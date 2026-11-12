"""Telegram link collector. Requires python-telegram-bot 22.x and python-docx."""
import logging
import re
import secrets
import sqlite3
from contextlib import contextmanager
from copy import deepcopy
from datetime import datetime, timedelta, timezone
from getpass import getpass
from io import BytesIO
from pathlib import Path
from urllib.parse import urlsplit

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from telegram import InlineKeyboardButton as Button, InlineKeyboardMarkup as Keyboard
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, MessageHandler, filters

BASE = Path(__file__).resolve().parent
DB = BASE / 'links.sqlite3'
TEMPLATE = BASE / 'template.docx'
TZ = timezone(timedelta(hours=5))
SECTIONS = ['Видео', 'Инфографика', 'Экспертное мнение']
JOIN_CODE = ADMIN_CODE = ''
TOPICS = []


@contextmanager
def connect():
    con = sqlite3.connect(DB)
    try:
        with con:
            yield con
    finally:
        con.close()


def initialize():
    global TOPICS
    doc = Document(TEMPLATE)
    if len(doc.tables) != 3 or any(len(t.rows[0].cells) != 10 for t in doc.tables):
        raise ValueError('В шаблоне нужны три таблицы с колонкой № и девятью темами.')
    TOPICS = [c.text.strip() for c in doc.tables[0].rows[0].cells[1:]]
    for table in doc.tables[1:]:
        if [c.text.strip() for c in table.rows[0].cells[1:]] != TOPICS:
            raise ValueError('Темы во всех трёх таблицах должны совпадать.')
    with connect() as con:
        con.execute('CREATE TABLE IF NOT EXISTS members (uid INTEGER PRIMARY KEY, name TEXT, admin INTEGER DEFAULT 0)')
        con.execute('''CREATE TABLE IF NOT EXISTS links (
            id INTEGER PRIMARY KEY, uid INTEGER, day TEXT, section INTEGER,
            topic INTEGER, url TEXT, UNIQUE(day, section, topic, url))''')


def role(uid):
    with connect() as con:
        row = con.execute('SELECT admin FROM members WHERE uid=?', (uid,)).fetchone()
    return None if row is None else row[0]


def today():
    return datetime.now(TZ).strftime('%d.%m.%Y')


def parse_day(text):
    return datetime.strptime(text, '%d.%m.%Y').strftime('%d.%m.%Y')


def menu(uid):
    rows = [[Button(name, callback_data=f's:{i}')] for i, name in enumerate(SECTIONS)]
    rows.append([Button('Дата публикаций', callback_data='date')])
    rows.append([Button('Мои ссылки', callback_data='mine')])
    if role(uid) == 1:
        rows.append([Button('Скачать Word', callback_data='report')])
    return Keyboard(rows)


async def start(update, context):
    if update.effective_chat.type != 'private':
        return
    uid = update.effective_user.id
    for key in ('section', 'topic', 'waiting'):
        context.user_data.pop(key, None)
    day = context.user_data.setdefault('day', today())
    if role(uid) is None:
        await update.message.reply_text('Для доступа отправьте /join и код отдела.\nАдминистратор: /admin и свой код.\nПример: /join ваш_код')
    else:
        await update.message.reply_text(f'Выберите раздел. Дата: {day}', reply_markup=menu(uid))


async def enroll(update, context):
    if update.effective_chat.type != 'private':
        return
    admin = update.message.text.split()[0].split('@')[0] == '/admin'
    expected = ADMIN_CODE if admin else JOIN_CODE
    supplied = ' '.join(context.args)
    if not secrets.compare_digest(supplied.encode(), expected.encode()):
        await update.message.reply_text('Неверный код доступа.')
        return
    user = update.effective_user
    with connect() as con:
        con.execute('''INSERT INTO members(uid,name,admin) VALUES(?,?,?)
            ON CONFLICT(uid) DO UPDATE SET name=excluded.name, admin=MAX(members.admin,excluded.admin)''',
            (user.id, user.full_name, int(admin)))
    await start(update, context)


def add_link(cell, url):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = 0
    link = OxmlElement('w:hyperlink')
    link.set(qn('r:id'), p.part.relate_to(url, RT.HYPERLINK, is_external=True))
    run = OxmlElement('w:r')
    props = OxmlElement('w:rPr')
    for tag, value in [('w:color', '0563C1'), ('w:sz', '16'), ('w:u', 'single')]:
        el = OxmlElement(tag)
        el.set(qn('w:val'), value)
        props.append(el)
    run.append(props)
    text = OxmlElement('w:t')
    text.text = url
    run.append(text)
    link.append(run)
    p._p.append(link)


def make_report(day):
    doc = Document(TEMPLATE)
    with connect() as con:
        records = con.execute('SELECT section,topic,url FROM links WHERE day=? ORDER BY id', (day,)).fetchall()
    for p in doc.paragraphs:
        if re.search(r'\d{2}\.\d{2}\.\d{4}', p.text):
            # Keep paragraph styling and replace only its date text.
            text = re.sub(r'\d{2}\.\d{2}\.\d{4}', day, p.text)
            if p.runs:
                p.runs[0].text = text
                for run in p.runs[1:]:
                    run.text = ''
            else:
                p.add_run(text)
    for section, table in enumerate(doc.tables):
        buckets = [[url for s,t,url in records if s == section and t == topic] for topic in range(9)]
        prototype = deepcopy(table.rows[1]._tr)
        for row in list(table.rows)[1:]:
            table._tbl.remove(row._tr)
        repeat = OxmlElement('w:tblHeader')
        table.rows[0]._tr.get_or_add_trPr().append(repeat)
        for index in range(max(1, max(map(len, buckets)))):
            table._tbl.append(deepcopy(prototype))
            row = table.rows[-1]
            for height in list(row._tr.xpath('./w:trPr/w:trHeight')):
                height.getparent().remove(height)
            for cell in row.cells:
                cell.text = ''
            row.cells[0].text = str(index + 1)
            for topic, urls in enumerate(buckets):
                if index < len(urls):
                    add_link(row.cells[topic + 1], urls[index])
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output, len(records)


async def button(update, context):
    query = update.callback_query
    await query.answer()
    uid = update.effective_user.id
    if update.effective_chat.type != 'private' or role(uid) is None:
        return
    data = query.data
    state = context.user_data
    state.pop('waiting', None)
    day = state.setdefault('day', today())
    if data == 'home':
        state.pop('topic', None)
        await query.edit_message_text(f'Выберите раздел. Дата: {day}', reply_markup=menu(uid))
    elif data.startswith('s:'):
        section = int(data.split(':')[1])
        state['section'] = section
        state.pop('topic', None)
        rows = [[Button(f'{i+1}. {title[:48]}', callback_data=f't:{section}:{i}')] for i,title in enumerate(TOPICS)]
        rows.append([Button('Главное меню', callback_data='home')])
        await query.edit_message_text(f'{SECTIONS[section]}: выберите тему.', reply_markup=Keyboard(rows))
    elif data.startswith('t:'):
        _, section, topic = data.split(':')
        state.update(section=int(section), topic=int(topic))
        await query.edit_message_text(f'{SECTIONS[int(section)]}\n{TOPICS[int(topic)]}\nДата: {day}\n\nОтправьте ссылки на эту тему, каждую с новой строки. Все ссылки сообщения попадут в эту одну тему.', reply_markup=Keyboard([[Button('Главное меню', callback_data='home')]]))
    elif data == 'date':
        state['waiting'] = 'date'
        await query.message.reply_text('Введите дату публикаций: ДД.ММ.ГГГГ. Например: ' + today())
    elif data == 'mine':
        with connect() as con:
            rows = con.execute('SELECT id,section,topic,url FROM links WHERE uid=? AND day=? ORDER BY id DESC LIMIT 10', (uid,day)).fetchall()
        lines = [f'№{i} | {SECTIONS[s]} | тема {t+1}\n{url}' for i,s,t,url in rows]
        await query.message.reply_text(('Последние 10 ваших ссылок за ' + day + ':\n\n' + '\n\n'.join(lines) + '\n\nУдалить свою ошибочную запись: /delete номер')[:4000] if rows else 'Ваших ссылок за эту дату пока нет.')
    elif data == 'report' and role(uid) == 1:
        state['waiting'] = 'report'
        await query.message.reply_text('За какую дату сформировать Word? Введите ДД.ММ.ГГГГ.')


async def remove(update, context):
    if update.effective_chat.type != 'private' or role(update.effective_user.id) is None:
        return
    if len(context.args) != 1 or not context.args[0].isdigit():
        await update.message.reply_text('Пример: /delete 12. Номер можно посмотреть в «Мои ссылки».')
        return
    with connect() as con:
        result = con.execute('DELETE FROM links WHERE id=? AND uid=?', (int(context.args[0]), update.effective_user.id))
    await update.message.reply_text('Запись удалена.' if result.rowcount else 'Такой вашей записи нет.')


async def message(update, context):
    uid = update.effective_user.id
    if role(uid) is None:
        await update.message.reply_text('Сначала получите доступ через /join и код отдела.')
        return
    state = context.user_data
    text = update.message.text.strip()
    waiting = state.get('waiting')
    if waiting:
        try:
            day = parse_day(text)
        except ValueError:
            await update.message.reply_text('Нужна дата в формате ДД.ММ.ГГГГ.')
            return
        state.pop('waiting', None)
        if waiting == 'date':
            state['day'] = day
            state.pop('topic', None)
            await update.message.reply_text(f'Дата установлена: {day}. Выберите раздел.', reply_markup=menu(uid))
        elif role(uid) == 1:
            output, count = make_report(day)
            if count:
                await update.message.reply_document(output, filename=f'Atyrau_{day}.docx', caption=f'Отчёт за {day}. Ссылок: {count}')
            else:
                await update.message.reply_text('За эту дату ссылок пока нет.')
        return
    if 'topic' not in state:
        await update.message.reply_text('Сначала выберите раздел и тему: /start')
        return
    urls = re.findall(r'https?://[^\s<>]+', text)
    urls = list(dict.fromkeys(u.rstrip('.,;') for u in urls))
    if not urls or any(not urlsplit(u).hostname for u in urls):
        await update.message.reply_text('Отправьте ссылки, начинающиеся с https:// или http://.')
        return
    day = state.setdefault('day', today())
    added = 0
    with connect() as con:
        for url in urls:
            result = con.execute('INSERT OR IGNORE INTO links(uid,day,section,topic,url) VALUES(?,?,?,?,?)', (uid,day,state['section'],state['topic'],url))
            added += result.rowcount
    await update.message.reply_text(f'Сохранено: {added}. Повторов: {len(urls)-added}.\n{SECTIONS[state["section"]]}, тема {state["topic"]+1}. Дата: {day}\nДругая тема: /start')


async def error(update, context):
    print('Ошибка обработки:', type(context.error).__name__)
    if update and update.effective_message:
        await update.effective_message.reply_text('Не удалось выполнить действие. Проверьте терминал; секретные ключи не отправляйте.')


def main():
    global JOIN_CODE, ADMIN_CODE
    logging.disable(logging.CRITICAL)  # Prevent third-party logs exposing a token.
    try:
        initialize()
        token = getpass('Токен BotFather (ввод скрыт): ').strip()
        if not re.fullmatch(r'[0-9]+:[A-Za-z0-9_-]+', token):
            print('Токен содержит лишние символы. Скопируйте только ключ из BotFather.')
            return
        JOIN_CODE = getpass('Придумайте код доступа для коллег (минимум 8 символов): ').strip()
        ADMIN_CODE = getpass('Придумайте отдельный код администратора (минимум 8 символов): ').strip()
        if min(len(JOIN_CODE), len(ADMIN_CODE)) < 8 or JOIN_CODE == ADMIN_CODE:
            print('Нужны два разных кода, каждый минимум 8 символов.')
            return
        app = Application.builder().token(token).build()
        private = filters.ChatType.PRIVATE
        for command, handler in [('start',start), ('join',enroll), ('admin',enroll), ('delete',remove)]:
            app.add_handler(CommandHandler(command, handler, filters=private))
        app.add_handler(CallbackQueryHandler(button))
        app.add_handler(MessageHandler(private & filters.TEXT & ~filters.COMMAND, message))
        app.add_error_handler(error)
        print('Бот запускается. В Telegram: /admin ваш_код. Остановка: Ctrl+C.')
        app.run_polling()
    except KeyboardInterrupt:
        print('Бот остановлен.')
    except Exception as exc:
        # Print the type only: Telegram exception messages may contain the token.
        print('Запуск не выполнен:', type(exc).__name__)
        print('Проверьте template.docx, токен и подключение к интернету.')


if __name__ == '__main__':
    main()
