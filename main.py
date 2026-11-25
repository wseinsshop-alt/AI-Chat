"""AI Chat — four-file Telegram AI client. Python 3.10+.

Runtime credentials, SQLite and encrypted backups live outside the source tree.
No model-produced Python, SQL, HTML or shell code is ever executed.
"""
from __future__ import annotations

import argparse
import base64
import copy
import csv
import getpass
import hashlib
import hmac
import html
import io
import ipaddress
import json
import logging
import math
import os
from pathlib import Path
import re
import secrets
import socket
import sqlite3
import threading
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime, timezone
from urllib.parse import urlsplit, urlunsplit

import requests
from requests.adapters import HTTPAdapter
import telebot
from telebot import types
from cryptography.fernet import Fernet
from PIL import Image, ImageOps
from pypdf import PdfReader
from docx import Document
from openpyxl import Workbook, load_workbook
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

ROOT = Path(__file__).resolve().parent
FORMATS = ('pdf', 'docx', 'xlsx', 'txt', 'md', 'csv')
DEFAULTS = {'invite_only': True, 'session_hours': 24, 'backup_hours': 24,
            'backup_keep': 7, 'daily_tokens': 100000, 'daily_usd': 0,
            'max_output_tokens': 4096, 'history_messages': 20, 'max_file_mb': 10,
            'workers': 4, 'prices': {}, 'transcription_model': 'whisper-1',
            'speech_model': 'tts-1', 'speech_voice': 'alloy',
            'transcription_usd_per_minute': None, 'speech_usd_per_million_characters': None,
            'system_prompt': 'Ты полезный помощник в Telegram. Не выдумывай выполненные действия.'}


def now() -> float:
    return time.time()


def day() -> str:
    return datetime.now(timezone.utc).date().isoformat()


def password_hash(text: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac('sha256', text.encode(), bytes.fromhex(salt), 300000).hex()


def safe_cell(value):
    if isinstance(value, str) and value.lstrip().startswith(('=', '+', '-', '@', '\t', '\r')):
        return "'" + value
    return value


def normalize_url(value: str) -> str:
    p = urlsplit(value.strip())
    if p.scheme != 'https' or not p.hostname or p.username or p.password or p.query or p.fragment:
        raise ValueError('Нужен публичный HTTPS base URL без пароля и параметров.')
    path = p.path.rstrip('/')
    return urlunsplit((p.scheme, p.netloc, '/v1' if path in ('', '/chat') else path, '', ''))


def public_host(url: str):
    p = urlsplit(normalize_url(url))
    answers = socket.getaddrinfo(p.hostname, p.port or 443, type=socket.SOCK_STREAM)
    if not answers or any(not ipaddress.ip_address(a[4][0]).is_global for a in answers):
        raise ValueError('Приватные, локальные и служебные IP-адреса не поддерживаются.')
    return answers[0][4][0]


class PinnedHTTPSAdapter(HTTPAdapter):
    """Use the validated IP while preserving TLS hostname verification and SNI."""
    def __init__(self, hostname, address, port):
        self.hostname, self.address, self.port = hostname, address, port
        super().__init__(max_retries=0)

    def get_connection_with_tls_context(self, request, verify, proxies=None, cert=None):
        return self.poolmanager.connection_from_host(self.address, self.port, scheme='https',
            pool_kwargs={'assert_hostname': self.hostname, 'server_hostname': self.hostname})


def configuration(setup=False):
    path = Path(os.environ.get('AI_CHAT_CONFIG', str(ROOT / 'config.json')))
    cfg = DEFAULTS | json.loads(path.read_text(encoding='utf-8'))
    default_dir = (Path(os.environ['LOCALAPPDATA']) / 'AIChat' if os.name == 'nt'
                   else Path(os.environ.get('XDG_DATA_HOME', str(Path.home() / '.local/share'))) / 'ai-chat')
    data = Path(os.environ.get('AI_CHAT_DATA_DIR') or cfg.get('data_dir') or default_dir).expanduser().resolve()
    data.mkdir(parents=True, exist_ok=True)
    os.chmod(data, 0o700)
    saved = data / 'settings.json'
    settings = json.loads(saved.read_text(encoding='utf-8')) if saved.exists() else {}
    cfg['telegram_token'] = os.environ.get('AI_CHAT_TOKEN') or cfg.get('telegram_token') or settings.get('telegram_token', '')
    admin_env = os.environ.get('AI_CHAT_ADMIN_IDS')
    cfg['admin_ids'] = ([int(x.strip()) for x in admin_env.split(',') if x.strip()] if admin_env
                        else cfg.get('admin_ids') or settings.get('admin_ids', []))
    if setup or not cfg['telegram_token'] or not cfg['admin_ids']:
        if not os.isatty(0):
            raise ValueError('Сначала выполни python main.py --setup или задай AI_CHAT_TOKEN и AI_CHAT_ADMIN_IDS.')
        token = getpass.getpass('Токен Telegram (Enter — сохранить текущий): ').strip()
        if token:
            cfg['telegram_token'] = token
        ids = input('Telegram ID администраторов через запятую: ').strip()
        if ids:
            cfg['admin_ids'] = [int(x.strip()) for x in ids.split(',')]
        if not cfg['telegram_token'] or not cfg['admin_ids']:
            raise ValueError('Нужны токен и хотя бы один ID администратора.')
        saved.write_text(json.dumps({k: cfg[k] for k in ('telegram_token', 'admin_ids')}), encoding='utf-8')
        os.chmod(saved, 0o600)
    cfg['data_path'] = data
    return cfg


SCHEMA = '''
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY,salt TEXT NOT NULL,hash TEXT NOT NULL,
 verified REAL DEFAULT 0,failures INTEGER DEFAULT 0,locked REAL DEFAULT 0,
 active INTEGER,current_chat INTEGER,daily_tokens INTEGER,daily_usd REAL);
CREATE TABLE IF NOT EXISTS panels(id INTEGER PRIMARY KEY AUTOINCREMENT,owner INTEGER NOT NULL,
 name TEXT NOT NULL,url TEXT NOT NULL,token TEXT NOT NULL,model TEXT DEFAULT '',stream INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS chats(id INTEGER PRIMARY KEY AUTOINCREMENT,owner INTEGER NOT NULL,
 panel INTEGER NOT NULL REFERENCES panels(id) ON DELETE CASCADE,title TEXT NOT NULL,
 model TEXT DEFAULT '',updated REAL NOT NULL);
CREATE TABLE IF NOT EXISTS messages(id INTEGER PRIMARY KEY AUTOINCREMENT,
 chat INTEGER NOT NULL REFERENCES chats(id) ON DELETE CASCADE,role TEXT NOT NULL,content TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS favorites(owner INTEGER,panel INTEGER REFERENCES panels(id) ON DELETE CASCADE,
 model TEXT,PRIMARY KEY(owner,panel,model));
CREATE TABLE IF NOT EXISTS invites(code TEXT PRIMARY KEY,uses INTEGER NOT NULL,expires REAL NOT NULL);
CREATE TABLE IF NOT EXISTS admitted(id INTEGER PRIMARY KEY);
CREATE TABLE IF NOT EXISTS usage(id INTEGER PRIMARY KEY AUTOINCREMENT,owner INTEGER,day TEXT,kind TEXT,
 model TEXT,input_tokens INTEGER,output_tokens INTEGER,cost REAL,estimated INTEGER,created REAL);
CREATE TABLE IF NOT EXISTS reservations(id TEXT PRIMARY KEY,owner INTEGER,day TEXT,tokens INTEGER,cost REAL);
'''


class Store:
    TABLES = ('users', 'panels', 'chats', 'messages', 'favorites', 'invites', 'admitted', 'usage')

    def __init__(self, data: Path):
        self.data = data
        self.path = data / 'bot.db'
        key = data / 'secret.key'
        if not key.exists():
            if self.path.exists():
                raise ValueError('Нет secret.key. Восстанови ключ, прежде чем открывать существующую базу.')
            key.write_bytes(Fernet.generate_key())
            os.chmod(key, 0o600)
        self.cipher = Fernet(key.read_bytes().strip())
        with self.db() as c:
            c.executescript(SCHEMA)
            # Upgrade the previous tg-ai-cabinet schema without replacing its tables.
            for table, field_name, definition in (
                ('users', 'current_chat', 'INTEGER'), ('users', 'daily_tokens', 'INTEGER'),
                ('users', 'daily_usd', 'REAL'), ('panels', 'stream', 'INTEGER DEFAULT 1')):
                if field_name not in [r['name'] for r in c.execute('PRAGMA table_info(' + table + ')')]:
                    c.execute(f'ALTER TABLE {table} ADD COLUMN {field_name} {definition}')
            if c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='history'").fetchone():
                for row in c.execute('SELECT * FROM history').fetchall():
                    panel = c.execute('SELECT * FROM panels WHERE id=? AND owner=?', (row['panel'], row['owner'])).fetchone()
                    if not panel:
                        continue
                    cid = c.execute('INSERT INTO chats(owner,panel,title,model,updated) VALUES(?,?,?,?,?)',
                                    (row['owner'], row['panel'], 'Перенесённый чат', panel['model'], now())).lastrowid
                    for msg in json.loads(row['messages']):
                        if msg.get('role') in ('user', 'assistant'):
                            c.execute('INSERT INTO messages(chat,role,content) VALUES(?,?,?)',
                                      (cid, msg['role'], json.dumps(msg['content'], ensure_ascii=False)))
                c.execute('DROP TABLE history')
        os.chmod(self.path, 0o600)

    @contextmanager
    def db(self):
        con = sqlite3.connect(self.path, timeout=30)
        con.row_factory = sqlite3.Row
        con.execute('PRAGMA foreign_keys=ON')
        try:
            with con:
                yield con
        finally:
            con.close()

    def rows(self, sql, args=()):
        with self.db() as c:
            return [dict(r) for r in c.execute(sql, args)]

    def one(self, sql, args=()):
        found = self.rows(sql, args)
        return found[0] if found else None

    def write(self, sql, args=()):
        with self.db() as c:
            return c.execute(sql, args).lastrowid

    def user(self, uid):
        return self.one('SELECT * FROM users WHERE id=?', (uid,))

    def panel(self, uid, pid=None):
        if pid is None:
            u = self.user(uid)
            pid = u['active'] if u else None
        return self.one('SELECT * FROM panels WHERE owner=? AND id=?', (uid, pid))

    def chat(self, uid):
        panel = self.panel(uid)
        if not panel:
            raise ValueError('Сначала добавь и выбери панель.')
        u = self.user(uid)
        chat = self.one('SELECT * FROM chats WHERE owner=? AND id=? AND panel=?',
                        (uid, u['current_chat'], panel['id']))
        if not chat:
            chat = self.one('SELECT * FROM chats WHERE owner=? AND panel=? ORDER BY updated DESC LIMIT 1', (uid, panel['id']))
        if not chat:
            return self.new_chat(uid)
        self.write('UPDATE users SET current_chat=? WHERE id=?', (chat['id'], uid))
        return chat

    def new_chat(self, uid):
        p = self.panel(uid)
        if not p:
            raise ValueError('Сначала выбери панель.')
        cid = self.write('INSERT INTO chats(owner,panel,title,model,updated) VALUES(?,?,?,?,?)',
                         (uid, p['id'], 'Новый чат', p['model'], now()))
        self.write('UPDATE users SET current_chat=? WHERE id=?', (cid, uid))
        return self.one('SELECT * FROM chats WHERE id=?', (cid,))

    def history(self, chat_id, limit=20):
        rows = self.rows('SELECT * FROM (SELECT * FROM messages WHERE chat=? ORDER BY id DESC LIMIT ?) ORDER BY id', (chat_id, limit))
        return [{'role': r['role'], 'content': json.loads(r['content'])} for r in rows]

    def append(self, cid, content, answer):
        with self.db() as c:
            for role, text in (('user', content), ('assistant', answer)):
                c.execute('INSERT INTO messages(chat,role,content) VALUES(?,?,?)', (cid, role, json.dumps(text, ensure_ascii=False)))
            c.execute('UPDATE chats SET updated=? WHERE id=?', (now(), cid))
            # Keep a bounded, persistent conversation; trim complete oldest pairs.
            c.execute('DELETE FROM messages WHERE chat=? AND id NOT IN (SELECT id FROM messages WHERE chat=? ORDER BY id DESC LIMIT 100)', (cid, cid))

    def backup(self):
        with self.db() as c:
            c.execute('BEGIN')
            payload = {'version': 2, 'created': now()}
            for table in self.TABLES:
                payload[table] = [dict(r) for r in c.execute('SELECT * FROM ' + table)]
        return self.cipher.encrypt(json.dumps(payload, ensure_ascii=False).encode())

    def decode_backup(self, blob):
        payload = json.loads(self.cipher.decrypt(blob))
        if payload.get('version') != 2 or not all(isinstance(payload.get(t), list) for t in self.TABLES):
            raise ValueError('Нужна резервная копия AI Chat версии 2 и исходный secret.key.')
        for panel in payload['panels']:
            normalize_url(panel['url'])
            self.cipher.decrypt(panel['token'].encode())
        return payload

    def restore(self, blob):
        payload = self.decode_backup(blob)
        # Validate data against an independent fresh database before changing the live one.
        check = sqlite3.connect(':memory:')
        try:
            check.executescript(SCHEMA)
            check.execute('PRAGMA foreign_keys=ON')
            for table in self.TABLES:
                fields = [r[1] for r in check.execute('PRAGMA table_info(' + table + ')')]
                for row in payload[table]:
                    check.execute(f"INSERT INTO {table} ({','.join(fields)}) VALUES ({','.join('?' for _ in fields)})", [row.get(k) for k in fields])
            if check.execute('PRAGMA foreign_key_check').fetchone():
                raise ValueError('Повреждённые связи в базе.')
        finally:
            check.close()
        self.automatic_backup(force=True)
        with self.db() as c:
            c.execute('DELETE FROM reservations')
            for table in reversed(self.TABLES):
                c.execute('DELETE FROM ' + table)
            for table in self.TABLES:
                fields = [r['name'] for r in c.execute('PRAGMA table_info(' + table + ')')]
                c.executemany(f"INSERT INTO {table} ({','.join(fields)}) VALUES ({','.join('?' for _ in fields)})", [[row.get(k) for k in fields] for row in payload[table]])
            c.execute('UPDATE users SET verified=0,failures=0,locked=0')

    def clear_database(self):
        """Make a forced encrypted backup, then remove user data but retain the DB schema."""
        self.automatic_backup(force=True)
        with self.db() as c:
            c.execute('DELETE FROM reservations')
            for table in reversed(self.TABLES):
                c.execute('DELETE FROM ' + table)

    def automatic_backup(self, hours=24, keep=7, force=False):
        folder = self.data / 'backups'
        folder.mkdir(exist_ok=True)
        files = sorted(folder.glob('*.enc'))
        if not force and files and now() - files[-1].stat().st_mtime < hours * 3600:
            return
        target = folder / (datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S') + '-' + secrets.token_hex(2) + '.enc')
        temp = target.with_suffix('.tmp')
        temp.write_bytes(self.backup())
        os.chmod(temp, 0o600)
        os.replace(temp, target)
        for old in sorted(folder.glob('*.enc'))[:-max(1, keep)]:
            old.unlink()


class Provider:
    def __init__(self, panel, cipher):
        self.url = panel['url']
        self.key = cipher.decrypt(panel['token'].encode()).decode()

    def request(self, path, payload=None, files=None, stream=False):
        address = public_host(self.url)
        parsed = urlsplit(self.url)
        session = requests.Session()
        session.trust_env = False
        session.mount('https://', PinnedHTTPSAdapter(parsed.hostname, address, parsed.port or 443))
        try:
            response = session.request('GET' if payload is None and files is None else 'POST', self.url + path,
                                       headers={'Authorization': 'Bearer ' + self.key, 'Host': parsed.netloc},
                                       json=payload if files is None else None, data=payload if files else None,
                                       files=files, stream=True, timeout=(15, 60), allow_redirects=False)
            if response.status_code != 200:
                response.close()
                raise ValueError(f'API: HTTP {response.status_code}. Проверь base URL, ключ, модель, поддержку функции и баланс.')
            # Keep the session alive until the response is consumed.
            response._ai_session = session
            return response
        except Exception:
            session.close()
            raise

    def json(self, path, payload=None):
        r = self.request(path, payload)
        try:
            raw = bounded_response(r, 10 * 1024 * 1024)
            return json.loads(raw)
        finally:
            r.close()
            r._ai_session.close()

    def models(self):
        found = self.json('/models').get('data', [])
        result = [str(m['id']) for m in found if isinstance(m, dict) and 'id' in m]
        if not result:
            raise ValueError('Панель вернула пустой список моделей.')
        return result


def bounded_response(response, limit):
    result = bytearray()
    for chunk in response.iter_content(65536):
        result.extend(chunk)
        if len(result) > limit:
            raise ValueError('Ответ сервера превышает лимит размера.')
    return bytes(result)


def image_content(raw, caption):
    with Image.open(io.BytesIO(raw)) as source:
        if source.width * source.height > 24000000:
            raise ValueError('Уменьши изображение: лимит 24 мегапикселя.')
        image = ImageOps.exif_transpose(source).convert('RGB')
        image.thumbnail((1600, 1600))
        out = io.BytesIO()
        image.save(out, 'JPEG', quality=82)
    return [{'type': 'text', 'text': caption or 'Что на изображении? Помоги разобраться.'},
            {'type': 'image_url', 'image_url': {'url': 'data:image/jpeg;base64,' + base64.b64encode(out.getvalue()).decode()}}]


def check_zip(raw):
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        if len(z.infolist()) > 2000 or sum(i.file_size for i in z.infolist()) > 40 * 1024 * 1024:
            raise ValueError('Слишком большой распакованный документ.')


def read_document(raw, name):
    ext = Path(name).suffix.lower()
    parts = []
    if ext == '.pdf':
        pdf = PdfReader(io.BytesIO(raw))
        if pdf.is_encrypted:
            raise ValueError('PDF защищён паролем.')
        if len(pdf.pages) > 60:
            raise ValueError('Лимит — 60 страниц PDF.')
        for page in pdf.pages:
            parts.append(page.extract_text() or '')
            if sum(map(len, parts)) > 60000:
                break
    elif ext == '.docx':
        check_zip(raw)
        doc = Document(io.BytesIO(raw))
        parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            parts.extend(' | '.join(cell.text for cell in row.cells) for row in table.rows)
    elif ext == '.xlsx':
        check_zip(raw)
        book = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        try:
            for sheet in book.worksheets[:5]:
                parts.append('Лист: ' + sheet.title)
                for row in sheet.iter_rows(max_row=1000, max_col=30, values_only=True):
                    parts.append(' | '.join('' if x is None else str(x) for x in row))
                    if sum(map(len, parts)) > 60000:
                        break
        finally:
            book.close()
    elif ext in ('.txt', '.md', '.csv', '.json'):
        try:
            parts = [raw.decode('utf-8-sig')]
        except UnicodeDecodeError:
            parts = [raw.decode('cp1251')]
    else:
        raise ValueError('Поддерживаются PDF, DOCX, XLSX, TXT, MD, CSV, JSON и изображения.')
    text = '\n'.join(parts)
    if not text.strip():
        raise ValueError('В документе нет текста. Скан PDF пришли фотографией.')
    return text[:60000], len(text) > 60000


def register_pdf_font():
    if 'AIChat' in pdfmetrics.getRegisteredFontNames():
        return
    font = next((p for p in (Path('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'), Path('C:/Windows/Fonts/arial.ttf')) if p.exists()), None)
    if font is None:
        raise ValueError('Для PDF установи fonts-dejavu-core на сервере.')
    bold = font.with_name('DejaVuSans-Bold.ttf' if font.name.startswith('Deja') else 'arialbd.ttf')
    pdfmetrics.registerFont(TTFont('AIChat', str(font)))
    pdfmetrics.registerFont(TTFont('AIChatBold', str(bold if bold.exists() else font)))
    pdfmetrics.registerFontFamily('AIChat', normal='AIChat', bold='AIChatBold', italic='AIChat', boldItalic='AIChatBold')


def create_document(text, fmt):
    if fmt not in FORMATS or len(text) > 120000:
        raise ValueError('Неподдерживаемый формат или слишком большой документ.')
    out = io.BytesIO()
    if fmt in ('txt', 'md'):
        out.write(text.encode('utf-8'))
    elif fmt == 'csv':
        cleaned = re.sub(r'^```(?:csv)?\s*|\s*```$', '', text.strip())
        writer_buffer = io.StringIO(newline='')
        writer = csv.writer(writer_buffer)
        for i, row in enumerate(csv.reader(io.StringIO(cleaned), strict=True)):
            if i >= 10000:
                raise ValueError('Слишком много строк CSV.')
            writer.writerow([safe_cell(x) for x in row])
        out.write(writer_buffer.getvalue().encode('utf-8-sig'))
    elif fmt == 'xlsx':
        cleaned = re.sub(r'^```(?:json)?\s*|\s*```$', '', text.strip())
        table = json.loads(cleaned)
        if not isinstance(table, list) or len(table) > 10000 or any(not isinstance(r, list) or len(r) > 100 for r in table):
            raise ValueError('Для XLSX нужен JSON-массив строк. Попроси ИИ исправить формат.')
        book = Workbook()
        sheet = book.active
        sheet.title = 'AI Chat'
        for row in table:
            sheet.append([safe_cell(x if isinstance(x, (str, int, float, bool)) or x is None else str(x)) for x in row])
        sheet.freeze_panes = 'A2'
        if sheet.max_row:
            from openpyxl.styles import Font, PatternFill
            for cell in sheet[1]:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='2563EB')
            sheet.auto_filter.ref = sheet.dimensions
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = min(50, max(12, max(len(str(c.value or '')) for c in col) + 2))
        book.save(out)
        book.close()
    elif fmt == 'docx':
        doc = Document()
        for line in text.splitlines():
            if line.startswith('#'):
                level = min(3, len(line) - len(line.lstrip('#')))
                doc.add_heading(line.lstrip('#').strip(), level=level)
            elif line.startswith(('- ', '* ')):
                doc.add_paragraph(line[2:], style='List Bullet')
            else:
                doc.add_paragraph(line)
        doc.save(out)
    else:
        register_pdf_font()
        doc = SimpleDocTemplate(out, pagesize=(595.28, 841.89), leftMargin=46, rightMargin=46, topMargin=58, bottomMargin=48)
        body = ParagraphStyle('Body', fontName='AIChat', fontSize=10.5, leading=16, spaceAfter=8, textColor=colors.HexColor('#243247'))
        heading = ParagraphStyle('Heading', parent=body, fontName='AIChatBold', fontSize=15, leading=21, spaceBefore=12, textColor=colors.HexColor('#1746A2'))
        story = []
        for line in text.splitlines():
            if not line.strip():
                story.append(Spacer(1, 6))
            elif not line.startswith('```'):
                clean = html.escape(line.lstrip('#').strip() if line.startswith('#') else line)
                clean = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', clean)
                story.append(Paragraph(clean, heading if line.startswith('#') else body))
        def page(canvas, document):
            canvas.saveState()
            canvas.setFont('AIChat', 8)
            canvas.setFillColor(colors.HexColor('#64748B'))
            canvas.drawString(46, 814, 'AI CHAT · DOCUMENT')
            canvas.drawRightString(549, 28, str(document.page))
            canvas.restoreState()
        doc.build(story or [Paragraph('Пустой документ', body)], onFirstPage=page, onLaterPages=page)
    out.seek(0)
    out.name = 'AI-Chat.' + fmt
    return out


@dataclass
class Job:
    cancel: threading.Event = field(default_factory=threading.Event)
    response: object = None


class Cancelled(Exception):
    pass


class SecretFilter(logging.Filter):
    def __init__(self, token):
        super().__init__()
        self.token = token

    def filter(self, record):
        text = record.getMessage().replace(self.token, '[TELEGRAM_TOKEN]')
        text = re.sub(r'(?i)Bearer\s+[^\s\"\']+', 'Bearer [REDACTED]', text)
        record.msg, record.args, record.exc_info, record.exc_text = text, (), None, None
        return True


def btn(text, data, style='primary'):
    # Telegram supports semantic button styles, not arbitrary RGB colors.
    if style not in ('primary', 'success', 'danger'):
        style = 'primary'
    return types.InlineKeyboardButton(text, callback_data=data, style=style)


def keyboard(*rows):
    kb = types.InlineKeyboardMarkup()
    for row in rows:
        kb.row(*row)
    return kb


class AIChat:
    def __init__(self, config, store=None, bot=None):
        self.cfg = config
        self.store = store or Store(config['data_path'])
        telebot.logger.addFilter(SecretFilter(config['telegram_token']))
        self.bot = bot or telebot.TeleBot(config['telegram_token'], threaded=True, num_threads=8)
        self.admins = set(map(int, config['admin_ids']))
        self.states = {}
        self.cache = {}
        self.jobs = {}
        self.lock = threading.RLock()
        self.user_locks = {}
        self.pool = ThreadPoolExecutor(max_workers=config['workers'])
        self.closed = threading.Event()
        self.bot.register_message_handler(self.command, commands=['start', 'menu', 'help', 'cancel', 'stop', 'new', 'backup', 'invite'])
        self.bot.register_callback_query_handler(self.callback, func=lambda call: True)
        self.bot.register_message_handler(self.message, content_types=['text', 'photo', 'document', 'voice', 'audio'])

    @staticmethod
    def model_label(model):
        """Readable model label; keep the exact provider model ID for API calls."""
        raw = (model or '').strip()
        low = raw.casefold()
        aliases = (
            (('gpt-4o-mini',), 'GPT-4o mini'), (('gpt-4o',), 'GPT-4o'),
            (('gpt-4.1-mini',), 'GPT-4.1 mini'), (('gpt-4.1-nano',), 'GPT-4.1 nano'),
            (('gpt-4.1',), 'GPT-4.1'), (('gpt-3.5-turbo',), 'GPT-3.5 Turbo'),
            (('claude-3-5-sonnet', 'claude-3.5-sonnet'), 'Claude 3.5 Sonnet'),
            (('claude-3-7-sonnet', 'claude-3.7-sonnet'), 'Claude 3.7 Sonnet'),
            (('claude-3-opus',), 'Claude 3 Opus'), (('claude-3-haiku',), 'Claude 3 Haiku'),
            (('gemini-2.5-pro',), 'Gemini 2.5 Pro'), (('gemini-2.5-flash',), 'Gemini 2.5 Flash'),
            (('gemini-2.0-flash',), 'Gemini 2.0 Flash'), (('deepseek-r1',), 'DeepSeek R1'),
            (('deepseek-v3',), 'DeepSeek V3'), (('llama-3.3',), 'Llama 3.3'),
            (('qwen2.5', 'qwen-2.5'), 'Qwen 2.5'),
        )
        for fragments, label in aliases:
            if any(part in low for part in fragments):
                return f'{label} · {raw}'
        return raw

    def model_label(model):
        """Readable display name; keep the exact provider model ID for API calls."""
        raw = (model or '').strip()
        low = raw.casefold()
        aliases = (
            (('gpt-4o-mini',), 'GPT-4o mini'), (('gpt-4o',), 'GPT-4o'),
            (('gpt-4.1-mini',), 'GPT-4.1 mini'), (('gpt-4.1-nano',), 'GPT-4.1 nano'),
            (('gpt-4.1',), 'GPT-4.1'), (('gpt-3.5-turbo',), 'GPT-3.5 Turbo'),
            (('claude-3-5-sonnet', 'claude-3.5-sonnet'), 'Claude 3.5 Sonnet'),
            (('claude-3-7-sonnet', 'claude-3.7-sonnet'), 'Claude 3.7 Sonnet'),
            (('claude-3-opus',), 'Claude 3 Opus'), (('claude-3-haiku',), 'Claude 3 Haiku'),
            (('gemini-2.5-pro',), 'Gemini 2.5 Pro'), (('gemini-2.5-flash',), 'Gemini 2.5 Flash'),
            (('gemini-2.0-flash',), 'Gemini 2.0 Flash'), (('deepseek-r1',), 'DeepSeek R1'),
            (('deepseek-v3',), 'DeepSeek V3'), (('llama-3.3',), 'Llama 3.3'),
            (('qwen2.5', 'qwen-2.5'), 'Qwen 2.5'),
        )
        for fragments, label in aliases:
            if any(part in low for part in fragments):
                return f'{label} · {raw}'
        return raw

    @staticmethod
    def model_label(model):
        """Readable model label; API requests still use the provider ID."""
        raw = (model or '').strip()
        low = raw.casefold()
        aliases = (
            (('gpt-4o-mini',), 'GPT-4o mini'), (('gpt-4o',), 'GPT-4o'),
            (('gpt-4.1-mini',), 'GPT-4.1 mini'), (('gpt-4.1-nano',), 'GPT-4.1 nano'),
            (('gpt-4.1',), 'GPT-4.1'), (('gpt-3.5-turbo',), 'GPT-3.5 Turbo'),
            (('claude-3-5-sonnet', 'claude-3.5-sonnet'), 'Claude 3.5 Sonnet'),
            (('claude-3-7-sonnet', 'claude-3.7-sonnet'), 'Claude 3.7 Sonnet'),
            (('claude-3-opus',), 'Claude 3 Opus'), (('claude-3-haiku',), 'Claude 3 Haiku'),
            (('gemini-2.5-pro',), 'Gemini 2.5 Pro'), (('gemini-2.5-flash',), 'Gemini 2.5 Flash'),
            (('gemini-2.0-flash',), 'Gemini 2.0 Flash'), (('deepseek-r1',), 'DeepSeek R1'),
            (('deepseek-v3',), 'DeepSeek V3'), (('llama-3.3',), 'Llama 3.3'),
            (('qwen2.5', 'qwen-2.5'), 'Qwen 2.5'),
        )
        for fragments, label in aliases:
            if any(part in low for part in fragments):
                return f'{label} · {raw}'
        return raw

    def tell(self, uid, text, keys=None, title=None):
        if title:
            text = '<b>' + html.escape(title) + '</b>\n\n' + html.escape(text)
        for pos in range(0, len(text), 1800):
            # Format each independent chunk to avoid breaking Telegram entities.
            chunk = text[pos:pos + 1800]
            if title and len(text) <= 1800:
                self.bot.send_message(uid, chunk, parse_mode='HTML', reply_markup=keys)
            else:
                self.bot.send_message(uid, chunk, reply_markup=keys if pos == 0 else None)

    def card(self, uid, title, body, keys=None):
        self.bot.send_message(uid, '<b>' + html.escape(title) + '</b>\n\n' + html.escape(body[:3500]),
                              parse_mode='HTML', reply_markup=keys)

    def erase(self, msg):
        try:
            self.bot.delete_message(msg.chat.id, msg.message_id)
        except Exception:
            pass

    def authenticated(self, uid):
        u = self.store.user(uid)
        return bool(u and u['verified'] + self.cfg['session_hours'] * 3600 > now())

    def gate(self, uid):
        if not self.store.user(uid):
            allowed = uid in self.admins or not self.cfg['invite_only'] or self.store.one('SELECT id FROM admitted WHERE id=?', (uid,))
            if not allowed:
                self.tell(uid, '🔒 Вход по приглашению. Попроси администратора прислать ссылку.')
                return False
            state = self.states.get(uid, {})
            if state.get('step') not in ('register', 'register_confirm'):
                self.states[uid] = {'step': 'register'}
                self.tell(uid, '🔐 Создай личный пароль: 8–128 символов. Не используй пароль Telegram. Я постараюсь удалить сообщение с паролем.')
            return False
        if not self.authenticated(uid):
            self.states[uid] = {'step': 'login'}
            self.tell(uid, '🔐 Введи личный пароль. Вход действует ' + str(self.cfg['session_hours']) + ' ч.')
            return False
        return True

    def verify_password(self, uid, text):
        u = self.store.user(uid)
        if not u or u['locked'] > now():
            raise ValueError('Слишком много попыток. Подожди 5 минут.')
        if not hmac.compare_digest(password_hash(text, u['salt']), u['hash']):
            count = u['failures'] + 1
            self.store.write('UPDATE users SET failures=?,locked=? WHERE id=?', (count, now() + 300 if count >= 5 else 0, uid))
            raise ValueError('Неверный пароль.')
        self.store.write('UPDATE users SET failures=0,locked=0 WHERE id=?', (uid,))

    def home(self, uid):
        self.states.pop(uid, None)
        p = self.store.panel(uid)
        body = 'Сначала добавь API-панель: «➕ Панель» → название → API URL → API-ключ.'
        if p:
            chat = self.store.chat(uid)
            body = (f"🟢 Панель: {p['name']}\n"
                    f"🧠 Модель: {self.model_label(chat['model'] or p['model'])}\n"
                    f"💬 Чат: {chat['title']}")
        body += '\n\nОтправляй сообщение, фото с подписью или документ. Для нового диалога нажми «Новый чат».'
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
        for row in (
            [('💬 Чаты', 'primary'), ('🧠 Модели', 'primary')],
            [('📚 Панели', 'primary'), ('➕ Добавить панель', 'success')],
            [('📎 Файл', 'success'), ('🆕 Новый чат', 'primary')],
            [('💰 Расходы', 'primary'), ('⚙️ Настройки', 'primary')],
        ):
            kb.row(*(types.KeyboardButton(t) for t, s in row))
        if uid in self.admins:
            kb.row(types.KeyboardButton('🛠 Админка · доступы'))
        self.bot.send_message(uid, '✨ AI CHAT  ·  личный кабинет', reply_markup=kb)
        self.card(uid, 'Привет! 👋', body, keyboard(
            [btn('🧠 Выбрать модель', 'models', 'success'), btn('📚 Мои панели', 'panels')],
            [btn('⭐ Избранные модели', 'models:favorites'), btn('📎 Создать файл', 'formats')]))

    def panels(self, uid):
        panels = self.store.rows('SELECT * FROM panels WHERE owner=? ORDER BY id', (uid,))
        current = self.store.panel(uid)
        buttons = []
        for p in panels:
            selected = current and current['id'] == p['id']
            buttons.append([btn(('✓ ' if selected else '') + p['name'][:38], f'panel:{p["id"]}', 'success' if selected else 'primary'),
                            btn('⚙️', f'panelcfg:{p["id"]}')])
        buttons += [[btn('➕ Добавить', 'addpanel', 'success')], [btn('🏠 Кабинет', 'home')]]
        self.card(uid, '📚 Твои панели', 'Выбери панель или открой её настройки.', keyboard(*buttons))

    def panel_settings(self, uid, pid):
        p = self.store.panel(uid, pid)
        if not p:
            raise ValueError('Панель не найдена.')
        self.card(uid, '⚙️ ' + p['name'], p['url'] + '\nПотоковый ответ: ' + ('включён' if p['stream'] else 'выключен'), keyboard(
            [btn('✏️ Название', f'renamepanel:{pid}'), btn('🔑 API-ключ', f'rekey:{pid}')],
            [btn('🔎 Проверить API', f'check:{pid}'), btn('↔️ Режим ответа', f'stream:{pid}')],
            [btn('🗑 Удалить', f'delpanel:{pid}', 'danger'), btn('‹ Панели', 'panels')]))

    def provider(self, panel):
        return Provider(panel, self.store.cipher)

    @staticmethod
    def model_label(model):
        """Readable display name; retain the provider's exact model ID for API calls."""
        raw = (model or '').strip()
        low = raw.casefold()
        aliases = (
            (('gpt-4o-mini',), 'GPT-4o mini'),
            (('gpt-4o',), 'GPT-4o'),
            (('gpt-4.1-mini',), 'GPT-4.1 mini'),
            (('gpt-4.1-nano',), 'GPT-4.1 nano'),
            (('gpt-4.1',), 'GPT-4.1'),
            (('gpt-3.5-turbo',), 'GPT-3.5 Turbo'),
            (('claude-3-5-sonnet', 'claude-3.5-sonnet'), 'Claude 3.5 Sonnet'),
            (('claude-3-7-sonnet', 'claude-3.7-sonnet'), 'Claude 3.7 Sonnet'),
            (('claude-3-opus', 'claude-3.0-opus'), 'Claude 3 Opus'),
            (('claude-3-haiku', 'claude-3.0-haiku'), 'Claude 3 Haiku'),
            (('gemini-2.5-pro',), 'Gemini 2.5 Pro'),
            (('gemini-2.5-flash',), 'Gemini 2.5 Flash'),
            (('gemini-2.0-flash',), 'Gemini 2.0 Flash'),
            (('deepseek-r1',), 'DeepSeek R1'),
            (('deepseek-v3',), 'DeepSeek V3'),
            (('llama-3.3',), 'Llama 3.3'),
            (('qwen2.5', 'qwen-2.5'), 'Qwen 2.5'),
        )
        for fragments, label in aliases:
            if any(part in low for part in fragments):
                return f'{label} · {raw}'
        return raw

    def models(self, uid, page=0, query='', favorites=False, refresh=False):
        p = self.store.panel(uid)
        if not p:
            raise ValueError('Сначала добавь панель кнопкой «➕ Добавить панель».')
        cached = self.cache.get(uid)
        if refresh or not cached or cached['panel'] != p['id'] or cached['expires'] < now():
            cached = {'panel': p['id'], 'all': self.provider(p).models(), 'expires': now() + 1800}
            self.cache[uid] = cached
        favs = {r['model'] for r in self.store.rows('SELECT model FROM favorites WHERE owner=? AND panel=?', (uid, p['id']))}
        available = [m for m in cached['all'] if query.casefold() in m.casefold() and (not favorites or m in favs)]
        available.sort(key=lambda m: (m not in favs, m.casefold()))
        cached.update(visible=available, query=query, favorites=favorites)
        cached['revision'] = secrets.token_hex(3)
        current = self.store.chat(uid)['model'] or p['model']
        total = max(1, math.ceil(len(available) / 7))
        page = max(0, min(page, total - 1))
        buttons = []
        for index in range(page * 7, min(len(available), (page + 1) * 7)):
            model = available[index]
            label = self.model_label(model)
            buttons.append([btn(('✅ ' if model == current else '🤖 ') + label[:55],
                                f'model:{cached["revision"]}:{index}',
                                'success' if model == current else 'primary'),
                            btn('⭐' if model in favs else '☆', f'favorite:{cached["revision"]}:{index}',
                                'success' if model in favs else 'primary')])
        nav = []
        if page:
            nav.append(btn('⬅️ Назад', f'modelpage:{page - 1}'))
        nav.append(btn(f'Страница {page + 1}/{total}', 'noop'))
        if page + 1 < total:
            nav.append(btn('Вперёд ➡️', f'modelpage:{page + 1}'))
        buttons += [nav,
                    [btn('🔎 Найти', 'searchmodel'), btn('📋 Все модели', 'models')],
                    [btn('⭐ Избранное', 'models:favorites', 'success'), btn('🔄 Обновить', 'refreshmodels')],
                    [btn('🏠 В кабинет', 'home')]]
        if not available:
            description = 'По этому запросу моделей нет. Нажми «📋 Все» или обнови список.'
        else:
            description = f'Найдено: {len(available)}\nНажми кнопку модели, чтобы выбрать её. ⭐ сохраняет модель в избранное.\nТекущий выбор: {self.model_label(current) if current else "не выбран"}.'
        self.card(uid, '🧠 Выбор модели · ' + p['name'], description, keyboard(*buttons))

    def chats(self, uid, page=0):
        all_chats = self.store.rows('SELECT c.*,p.name AS panel_name FROM chats c JOIN panels p ON p.id=c.panel WHERE c.owner=? ORDER BY c.updated DESC', (uid,))
        total = max(1, math.ceil(len(all_chats) / 8))
        page = max(0, min(page, total - 1))
        buttons = [[btn(c['title'][:35] + ' · ' + c['panel_name'][:15], f'chat:{c["id"]}'), btn('⚙️', f'chatcfg:{c["id"]}')]
                   for c in all_chats[page * 8:page * 8 + 8]]
        nav = [btn(f'{page + 1}/{total}', 'noop')]
        if page:
            nav.insert(0, btn('‹', f'chatpage:{page - 1}'))
        if page + 1 < total:
            nav.append(btn('›', f'chatpage:{page + 1}'))
        buttons += [nav, [btn('➕ Новый чат', 'newchat', 'success'), btn('🏠 Кабинет', 'home')]]
        self.card(uid, '💬 Твои разговоры', 'История сохраняется отдельно для каждого чата.', keyboard(*buttons))

    def owned_chat(self, uid, cid):
        chat = self.store.one('SELECT * FROM chats WHERE id=? AND owner=?', (cid, uid))
        if not chat:
            raise ValueError('Чат не найден.')
        return chat

    def settings(self, uid):
        self.card(uid, '⚙️ Настройки', 'Пароль, лимиты и управление сессией.', keyboard(
            [btn('🔐 Сменить пароль', 'password'), btn('💰 Лимиты', 'limits')],
            [btn('🚪 Выйти', 'logout', 'danger'), btn('🏠 Кабинет', 'home')]))

    def usage(self, uid):
        stats = self.store.one('SELECT COUNT(*) AS requests,COALESCE(SUM(input_tokens+output_tokens),0) AS tokens,COALESCE(SUM(cost),0) AS cost,COALESCE(SUM(estimated),0) AS estimates FROM usage WHERE owner=? AND day=?', (uid, day()))
        unknown = self.store.one('SELECT COUNT(*) AS n FROM usage WHERE owner=? AND day=? AND cost IS NULL', (uid, day()))['n']
        u = self.store.user(uid)
        token_limit = u['daily_tokens'] if u['daily_tokens'] is not None else self.cfg['daily_tokens']
        usd_limit = u['daily_usd'] if u['daily_usd'] is not None else self.cfg['daily_usd']
        self.card(uid, '💰 Сегодня · UTC',
                  f"Вызовов API: {stats['requests']}\nТокенов: {stats['tokens']} / {token_limit or 'без лимита'}\n"
                  f"Расчётная стоимость: ${stats['cost']:.6f}\nUSD-лимит: {usd_limit or 'не задан'}\n"
                  f"Вызовов с неизвестной ценой: {unknown}\nОценённых записей: {stats['estimates']}\n\n"
                  'Это локальный учёт по usage и заданным тарифам. Счёт провайдера может отличаться. При отсутствии usage используется консервативная оценка.',
                  keyboard([btn('Настроить лимиты', 'limits'), btn('🏠 Кабинет', 'home')]))

    def reserve(self, uid, model, tokens, estimated_cost):
        u = self.store.user(uid)
        token_limit = u['daily_tokens'] if u['daily_tokens'] is not None else self.cfg['daily_tokens']
        usd_limit = u['daily_usd'] if u['daily_usd'] is not None else self.cfg['daily_usd']
        if usd_limit and estimated_cost is None:
            raise ValueError('Задан USD-лимит, но тариф этой операции неизвестен. Настрой prices/тарифы голоса в config.json или используй только токеновый лимит.')
        rid = secrets.token_hex(12)
        with self.store.db() as c:
            c.execute('BEGIN IMMEDIATE')
            spent = c.execute('SELECT COALESCE(SUM(input_tokens+output_tokens),0),COALESCE(SUM(cost),0) FROM usage WHERE owner=? AND day=?', (uid, day())).fetchone()
            reserved = c.execute('SELECT COALESCE(SUM(tokens),0),COALESCE(SUM(cost),0) FROM reservations WHERE owner=? AND day=?', (uid, day())).fetchone()
            if token_limit and spent[0] + reserved[0] + tokens > token_limit:
                raise ValueError('Достигнут дневной лимит токенов. Уменьши историю, ответ или измени лимит.')
            if usd_limit and spent[1] + reserved[1] + (estimated_cost or 0) > usd_limit:
                raise ValueError('Недостаточно дневного USD-лимита для этого запроса.')
            c.execute('INSERT INTO reservations VALUES(?,?,?,?,?)', (rid, uid, day(), tokens, estimated_cost))
        return rid

    def charge(self, uid, rid, kind, model, incoming, outgoing, cost, estimated):
        with self.store.db() as c:
            reservation = c.execute('SELECT day FROM reservations WHERE id=? AND owner=?', (rid, uid)).fetchone()
            c.execute('INSERT INTO usage(owner,day,kind,model,input_tokens,output_tokens,cost,estimated,created) VALUES(?,?,?,?,?,?,?,?,?)',
                      (uid, reservation['day'] if reservation else day(), kind, model, incoming, outgoing, cost, int(estimated), now()))
            c.execute('DELETE FROM reservations WHERE id=?', (rid,))

    def price(self, panel, model, incoming, outgoing):
        price = self.cfg['prices'].get(panel['url'] + '|' + model, self.cfg['prices'].get(model))
        if not price:
            return None
        return (incoming * float(price['input']) + outgoing * float(price['output'])) / 1000000

    def launch(self, uid, function):
        with self.lock:
            if uid in self.jobs:
                raise ValueError('Уже выполняется запрос. Нажми «Стоп» или дождись ответа.')
            if len(self.jobs) >= self.cfg['workers'] * 2:
                raise ValueError('Сейчас много запросов. Попробуй чуть позже.')
            job = Job()
            self.jobs[uid] = job
        def worker():
            try:
                function(job)
            except Cancelled:
                self.tell(uid, '⏹ Запрос остановлен. Провайдер всё ещё может учесть уже выполненную работу.')
            except ValueError as error:
                self.tell(uid, str(error))
            except Exception:
                self.tell(uid, 'Не удалось обработать запрос. Проверь API и поддержку функции. Платные запросы автоматически не повторяются.')
            finally:
                with self.lock:
                    self.jobs.pop(uid, None)
        self.pool.submit(worker)

    def stop(self, uid):
        with self.lock:
            job = self.jobs.get(uid)
            if job:
                job.cancel.set()
        self.tell(uid, '⏹ Остановка запрошена. Подожди завершения сетевого чтения.' if job else 'Нет активного запроса.')

    def last_answer(self, uid, cid=None):
        chat = self.owned_chat(uid, cid) if cid else self.store.chat(uid)
        row = self.store.one("SELECT content FROM messages WHERE chat=? AND role='assistant' ORDER BY id DESC LIMIT 1", (chat['id'],))
        if not row:
            raise ValueError('В этом чате ещё нет ответа ИИ.')
        return json.loads(row['content'])

    def send_answer(self, uid, text, cid):
        for pos in range(0, len(text), 1800):
            chunk = html.escape(text[pos:pos + 1800])
            chunk = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', chunk)
            self.bot.send_message(uid, chunk, parse_mode='HTML')
        self.card(uid, '✨ Готово', 'Что сделать с ответом?', keyboard(
            [btn('Короче', f'rewrite:short:{cid}'), btn('Подробнее', f'rewrite:long:{cid}'), btn('Проще', f'rewrite:simple:{cid}')],
            [btn('↻ Повторить', f'retry:{cid}'), btn('🔊 Озвучить', f'speak:{cid}')],
            [btn('📄 Скачать', f'saveformats:{cid}', 'success'), btn('💬 Чаты', 'chats')]))

    def generation(self, uid, job, content, cid=None, fmt=None, retry=False):
        chat = self.owned_chat(uid, cid) if cid else self.store.chat(uid)
        panel = self.store.panel(uid, chat['panel'])
        model = chat['model'] or panel['model']
        if not model:
            raise ValueError('Сначала выбери модель.')
        history = self.store.history(chat['id'], self.cfg['history_messages'])
        if retry:
            # Replay the latest user input without duplicating its previous answer.
            if len(history) < 2 or history[-1]['role'] != 'assistant':
                raise ValueError('Нет ответа для повторной генерации.')
            content = history[-2]['content']
            history = history[:-2]
        if isinstance(content, list):
            for item in history:
                if isinstance(item['content'], list):
                    item['content'] = '\n'.join(x.get('text', '') for x in item['content']) + '\n[Ранее было изображение]'
        def weight(msg):
            return len(msg['content']) if isinstance(msg['content'], str) else sum(len(x.get('text', '')) for x in msg['content']) + 20000
        while history and sum(weight(x) for x in history) > 60000:
            history = history[2:]
        instruction = self.cfg['system_prompt']
        if fmt:
            instruction += f' Подготовь содержимое документа {fmt}. Файл создаст приложение; не обещай ссылку и не добавляй вводный комментарий.'
            if fmt == 'xlsx':
                instruction += ' Верни только JSON-массив строк таблицы: [["Заголовок"],["Значение"]], без ограждений.'
            elif fmt == 'csv':
                instruction += ' Верни только CSV с заголовком.'
            else:
                instruction += ' Используй простые заголовки и абзацы. Не используй широкие Markdown-таблицы.'
        messages = [{'role': 'system', 'content': instruction}] + history + [{'role': 'user', 'content': content}]
        # UTF-8 bytes are a conservative text-token reservation, not an exact tokenizer.
        incoming_est = sum(len(json.dumps(x, ensure_ascii=False).encode()) if isinstance(x['content'], str)
                           else sum(len(y.get('text', '').encode()) for y in x['content']) + 20000 for x in messages)
        outgoing_max = int(self.cfg['max_output_tokens'])
        estimated_cost = self.price(panel, model, incoming_est, outgoing_max)
        rid = self.reserve(uid, model, incoming_est + outgoing_max, estimated_cost)
        answer = ''
        usage = None
        response = None
        request_started = False
        finish = None
        status = None
        try:
            status = self.bot.send_message(uid, '✦ Думаю…', reply_markup=keyboard([btn('⏹ Стоп', 'stop', 'danger')]))
            if job.cancel.is_set():
                raise Cancelled()
            payload = {'model': model, 'messages': messages, 'max_tokens': outgoing_max, 'stream': bool(panel['stream'])}
            if panel['stream']:
                payload['stream_options'] = {'include_usage': True}
            request_started = True
            response = self.provider(panel).request('/chat/completions', payload, stream=bool(panel['stream']))
            job.response = response
            if panel['stream']:
                last_edit = now()
                for raw in response.iter_lines(chunk_size=256):
                    if job.cancel.is_set():
                        raise Cancelled()
                    if not raw or not raw.startswith(b'data:'):
                        continue
                    data = raw[5:].strip()
                    if data == b'[DONE]':
                        break
                    event = json.loads(data)
                    if event.get('error'):
                        raise ValueError('Провайдер вернул ошибку генерации.')
                    if event.get('usage'):
                        usage = event['usage']
                    for choice in event.get('choices', []):
                        delta = choice.get('delta', {}).get('content')
                        if isinstance(delta, str):
                            answer += delta
                        finish = choice.get('finish_reason') or finish
                    if len(answer) > 120000:
                        raise ValueError('Ответ слишком большой; уменьши лимит генерации.')
                    if now() - last_edit > 2 and answer:
                        try:
                            self.bot.edit_message_text('✦ ' + answer[-2800:], uid, status.message_id,
                                                       reply_markup=keyboard([btn('⏹ Стоп', 'stop', 'danger')]))
                        except Exception:
                            pass
                        last_edit = now()
            else:
                result = json.loads(bounded_response(response, 4 * 1024 * 1024))
                choice = result['choices'][0]
                answer = choice['message'].get('content', '')
                finish = choice.get('finish_reason')
                usage = result.get('usage')
            if job.cancel.is_set():
                raise Cancelled()
            if not isinstance(answer, str) or not answer.strip():
                raise ValueError('Модель не вернула текст. Проверь модель и режим потока в настройках панели.')
            self.store.append(chat['id'], content, answer)
            if chat['title'] == 'Новый чат':
                title = content if isinstance(content, str) else content[0].get('text', 'Изображение')
                self.store.write('UPDATE chats SET title=? WHERE id=? AND owner=?', (title[:50], chat['id'], uid))
            if fmt:
                self.bot.send_document(uid, create_document(answer, fmt), caption='📎 AI Chat · ' + fmt.upper())
            else:
                self.send_answer(uid, answer, chat['id'])
            if finish == 'length':
                self.tell(uid, '⚠️ Достигнут лимит ответа модели. Ответ или файл может быть неполным.')
        finally:
            if response is not None:
                response.close()
                response._ai_session.close()
            if usage:
                incoming = int(usage.get('prompt_tokens', usage.get('input_tokens', incoming_est)))
                outgoing = int(usage.get('completion_tokens', usage.get('output_tokens', 0)))
                cost = self.price(panel, model, incoming, outgoing)
                self.charge(uid, rid, 'chat', model, incoming, outgoing, cost, False)
            elif request_started:
                # Unknown/cancelled calls keep their full reservation in the local ledger.
                self.charge(uid, rid, 'chat', model, incoming_est, outgoing_max, estimated_cost, True)
            else:
                self.store.write('DELETE FROM reservations WHERE id=?', (rid,))
            try:
                if status is not None:
                    self.bot.delete_message(uid, status.message_id)
            except Exception:
                pass

    def formats(self, uid, cid=None):
        prefix = 'save' if cid else 'format'
        buttons = []
        for i in range(0, len(FORMATS), 2):
            buttons.append([btn(fmt.upper(), f'{prefix}:{fmt}' + (f':{cid}' if cid else ''), 'success') for fmt in FORMATS[i:i + 2]])
        buttons.append([btn('🏠 Кабинет', 'home')])
        self.card(uid, '📄 Выбери формат', 'Скачать последний ответ этого чата.' if cid else 'После выбора напиши, какой документ подготовить.', keyboard(*buttons))

    def download(self, obj):
        limit = min(19, int(self.cfg['max_file_mb'])) * 1024 * 1024
        if not obj.file_size or obj.file_size > limit:
            raise ValueError(f'Лимит вложения: {limit // (1024 * 1024)} МБ.')
        info = self.bot.get_file(obj.file_id)
        # Telegram's getFile download URL contains the bot token; never send it to a panel.
        raw = self.bot.download_file(info.file_path)
        if len(raw) > limit:
            raise ValueError('Слишком большой файл.')
        return raw

    def transcribe(self, uid, job, message):
        panel = self.store.panel(uid)
        if not panel:
            raise ValueError('Сначала добавь панель.')
        obj = message.voice or message.audio
        raw = self.download(obj)
        seconds = getattr(obj, 'duration', 0) or 60
        rate = self.cfg['transcription_usd_per_minute']
        cost = seconds / 60 * float(rate) if rate is not None else None
        model = self.cfg['transcription_model']
        rid = self.reserve(uid, model, 0, cost)
        started = False
        response = None
        try:
            if job.cancel.is_set():
                raise Cancelled()
            started = True
            name = getattr(obj, 'file_name', None) or 'voice.ogg'
            response = self.provider(panel).request('/audio/transcriptions', {'model': model},
                                                   files={'file': (name, raw, getattr(obj, 'mime_type', None) or 'audio/ogg')})
            job.response = response
            result = json.loads(bounded_response(response, 1024 * 1024))
            text = result.get('text', '').strip()
            if not text:
                raise ValueError('Не удалось распознать речь.')
            if job.cancel.is_set():
                raise Cancelled()
            self.tell(uid, '🎙 Распознано:\n' + text)
        finally:
            if response is not None:
                response.close()
                response._ai_session.close()
            if started:
                self.charge(uid, rid, 'transcription', model, 0, 0, cost, True)
            else:
                self.store.write('DELETE FROM reservations WHERE id=?', (rid,))
        # Speech and chat are two separately billable operations and budget checks.
        self.generation(uid, job, text)

    def speak(self, uid, job, cid):
        chat = self.owned_chat(uid, cid)
        panel = self.store.panel(uid, chat['panel'])
        text = self.last_answer(uid, cid)
        if len(text) > 4000:
            raise ValueError('Для озвучки лимит 4000 символов. Сначала нажми «Короче».')
        rate = self.cfg['speech_usd_per_million_characters']
        cost = len(text) / 1000000 * float(rate) if rate is not None else None
        model = self.cfg['speech_model']
        rid = self.reserve(uid, model, 0, cost)
        response = None
        started = False
        try:
            if job.cancel.is_set():
                raise Cancelled()
            started = True
            response = self.provider(panel).request('/audio/speech',
                {'model': model, 'voice': self.cfg['speech_voice'], 'input': text, 'response_format': 'mp3'})
            job.response = response
            audio = bounded_response(response, 15 * 1024 * 1024)
            if job.cancel.is_set():
                raise Cancelled()
            buf = io.BytesIO(audio)
            buf.name = 'AI-Chat.mp3'
            self.bot.send_audio(uid, buf, caption='🔊 Голос сгенерирован ИИ')
        finally:
            if response is not None:
                response.close()
                response._ai_session.close()
            if started:
                self.charge(uid, rid, 'speech', model, 0, 0, cost, True)
            else:
                self.store.write('DELETE FROM reservations WHERE id=?', (rid,))

    def admin_menu(self, uid):
        if uid not in self.admins:
            raise ValueError('Этот раздел доступен администратору.')
        users = self.store.one('SELECT COUNT(*) AS n FROM users')['n']
        self.card(uid, '🛠 Админка · доступы и база',
                  f'Пользователей: {users}\nРежим: ' + ('только по приглашению' if self.cfg['invite_only'] else 'регистрация открыта') +
                  '\n\nЧтобы добавить человека: нажми «🔗 Создать приглашение» и перешли ему ссылку. Он открывает её в Telegram и задаёт свой пароль. Ссылка одноразовая и действует 24 часа.',
                  keyboard(
                      [btn('🔗 Создать приглашение', 'invite', 'success')],
                      [btn('📤 Экспорт БД', 'backup'), btn('📥 Импорт БД', 'restore')],
                      [btn('🗑 Очистить базу', 'cleardb', 'danger')],
                      [btn('🏠 В кабинет', 'home')]))

    def invite(self, uid):
        if uid not in self.admins:
            raise ValueError('Нужны права администратора.')
        code = secrets.token_urlsafe(18)
        self.store.write('INSERT INTO invites VALUES(?,?,?)', (code, 1, now() + 86400))
        name = self.bot.get_me().username
        self.tell(uid, f'🔗 Ссылка готова — перешли её человеку, которого хочешь добавить. Она работает один раз и истекает через 24 часа.\n\nhttps://t.me/{name}?start=invite_{code}')

    def admit(self, uid, code):
        if uid in self.admins or self.store.user(uid):
            return
        hashed = code  # High-entropy bearer code; never included in logs.
        with self.store.db() as c:
            c.execute('BEGIN IMMEDIATE')
            invite = c.execute('SELECT * FROM invites WHERE code=?', (hashed,)).fetchone()
            if not invite or invite['uses'] <= 0 or invite['expires'] < now():
                raise ValueError('Приглашение недействительно или истекло.')
            c.execute('UPDATE invites SET uses=uses-1 WHERE code=?', (hashed,))
            c.execute('INSERT OR IGNORE INTO admitted VALUES(?)', (uid,))

    def export(self, uid):
        if uid not in self.admins:
            raise ValueError('Нужны права администратора.')
        blob = io.BytesIO(self.store.backup())
        blob.name = 'AI-Chat-backup.enc'
        self.bot.send_document(uid, blob, caption='🔐 Зашифрованная копия. Для восстановления нужен исходный secret.key — храни его отдельно.')

    def command(self, message):
        if message.chat.type != 'private':
            return
        uid = message.from_user.id
        with self.lock:
            lock = self.user_locks.setdefault(uid, threading.RLock())
        with lock:
            try:
                command, *args = message.text.split(maxsplit=1)
                command = command.split('@')[0]
                if command == '/start' and args and args[0].startswith('invite_'):
                    self.admit(uid, args[0][7:])
                if command == '/stop':
                    if self.authenticated(uid):
                        self.stop(uid)
                    return
                self.states.pop(uid, None)
                if not self.gate(uid):
                    return
                if command == '/new':
                    if uid in self.jobs:
                        raise ValueError('Сначала останови текущий запрос.')
                    self.store.new_chat(uid)
                elif command == '/backup':
                    self.export(uid)
                    return
                elif command == '/invite':
                    self.invite(uid)
                    return
                elif command == '/help':
                    self.tell(uid, 'AI Chat\n/start — кабинет\n/new — новый чат\n/cancel — отменить ввод\n/stop — остановить запрос\n\nФото: добавь вопрос в подпись. Файлы: PDF/DOCX/XLSX/TXT/MD/CSV/JSON. Голос работает при поддержке /audio/transcriptions у панели. Озвучка — /audio/speech. Это отдельные платные запросы API.')
                self.home(uid)
            except ValueError as error:
                self.tell(uid, str(error))
            except Exception:
                self.tell(uid, 'Не удалось выполнить команду. Проверь настройки.')

    def callback(self, call):
        try:
            self.bot.answer_callback_query(call.id)
        except Exception:
            pass
        if not call.message or call.message.chat.type != 'private':
            return
        uid = call.from_user.id
        with self.lock:
            lock = self.user_locks.setdefault(uid, threading.RLock())
        with lock:
            try:
                if not self.gate(uid):
                    return
                data = call.data
                parts = data.split(':')
                action = parts[0]
                if action == 'stop':
                    self.stop(uid)
                    return
                if action == 'noop':
                    return
                if uid in self.jobs and action not in ('home', 'usage'):
                    raise ValueError('Дождись ответа или нажми «Стоп», прежде чем менять настройки.')
                if action == 'home':
                    self.home(uid)
                elif action == 'panels':
                    self.states.pop(uid, None)
                    self.panels(uid)
                elif action == 'addpanel':
                    self.states[uid] = {'step': 'panel_name'}
                    self.tell(uid, 'Название панели? /cancel — отмена.')
                elif action == 'panel':
                    pid = int(parts[1])
                    if not self.store.panel(uid, pid):
                        raise ValueError('Панель не найдена.')
                    self.store.write('UPDATE users SET active=?,current_chat=NULL WHERE id=?', (pid, uid))
                    self.cache.pop(uid, None)
                    self.home(uid)
                elif action == 'panelcfg':
                    self.panel_settings(uid, int(parts[1]))
                elif action in ('renamepanel', 'rekey', 'stream', 'check', 'delpanel', 'yespanel'):
                    pid = int(parts[1])
                    panel = self.store.panel(uid, pid)
                    if not panel:
                        raise ValueError('Панель не найдена.')
                    if action in ('renamepanel', 'rekey'):
                        self.states[uid] = {'step': action, 'panel': pid}
                        self.tell(uid, 'Введи новое название.' if action == 'renamepanel' else 'Пришли новый API-ключ. Старый ключ сохранится, если проверка не пройдёт.')
                    elif action == 'stream':
                        self.store.write('UPDATE panels SET stream=1-stream WHERE owner=? AND id=?', (uid, pid))
                        self.panel_settings(uid, pid)
                    elif action == 'check':
                        start = now()
                        models = self.provider(panel).models()
                        self.card(uid, '✅ API доступен', f'Моделей: {len(models)}\nВремя: {now() - start:.2f} с\nПроверен /models. Поддержку фото и голоса нужно проверять выбранной моделью.')
                    elif action == 'delpanel':
                        self.card(uid, '🗑 Удалить панель?', panel['name'] + '\nБудут удалены ключ, все чаты панели и избранное.', keyboard(
                            [btn('Да, удалить', f'yespanel:{pid}', 'danger'), btn('Отмена', f'panelcfg:{pid}')]))
                    else:
                        with self.store.db() as c:
                            c.execute('DELETE FROM panels WHERE id=? AND owner=?', (pid, uid))
                            replacement = c.execute('SELECT id FROM panels WHERE owner=? ORDER BY id LIMIT 1', (uid,)).fetchone()
                            c.execute('UPDATE users SET active=?,current_chat=NULL WHERE id=? AND active=?', (replacement['id'] if replacement else None, uid, pid))
                        self.cache.pop(uid, None)
                        self.panels(uid)
                elif action in ('models', 'refreshmodels', 'modelpage'):
                    self.states.pop(uid, None)
                    cache = self.cache.get(uid, {})
                    self.models(uid, page=int(parts[1]) if action == 'modelpage' else 0,
                                query=cache.get('query', '') if action == 'modelpage' else '',
                                favorites=(len(parts) > 1 and parts[1] == 'favorites') if action == 'models' else cache.get('favorites', False),
                                refresh=action == 'refreshmodels')
                elif action == 'searchmodel':
                    self.states[uid] = {'step': 'searchmodel'}
                    self.tell(uid, 'Напиши часть названия модели. /cancel — отмена.')
                elif action in ('model', 'favorite'):
                    p = self.store.panel(uid)
                    cache = self.cache.get(uid, {})
                    if not p or cache.get('panel') != p['id'] or cache.get('expires', 0) < now():
                        raise ValueError('Список устарел. Открой модели заново.')
                    if len(parts) != 3 or parts[1] != cache.get('revision'):
                        raise ValueError('Это старое меню. Открой модели заново.')
                    model = cache['visible'][int(parts[2])]
                    if action == 'favorite':
                        exists = self.store.one('SELECT model FROM favorites WHERE owner=? AND panel=? AND model=?', (uid, p['id'], model))
                        if exists:
                            self.store.write('DELETE FROM favorites WHERE owner=? AND panel=? AND model=?', (uid, p['id'], model))
                        else:
                            self.store.write('INSERT INTO favorites VALUES(?,?,?)', (uid, p['id'], model))
                        self.models(uid, query=cache.get('query', ''), favorites=cache.get('favorites', False))
                    else:
                        self.store.write('UPDATE panels SET model=? WHERE owner=? AND id=?', (model, uid, p['id']))
                        self.store.new_chat(uid)
                        self.home(uid)
                elif action in ('chats', 'chatpage'):
                    self.chats(uid, int(parts[1]) if action == 'chatpage' else 0)
                elif action in ('chat', 'chatcfg', 'renamechat', 'delchat', 'yeschat', 'history'):
                    cid = int(parts[1])
                    chat = self.owned_chat(uid, cid)
                    if action == 'chat':
                        self.store.write('UPDATE users SET active=?,current_chat=? WHERE id=?', (chat['panel'], cid, uid))
                        self.home(uid)
                        recent = self.store.history(cid, 4)
                        if recent:
                            self.tell(uid, 'Последние сообщения:\n\n' + '\n\n'.join(('Ты: ' if m['role'] == 'user' else 'ИИ: ') + (m['content'][:700] if isinstance(m['content'], str) else '[изображение]') for m in recent))
                    elif action == 'chatcfg':
                        self.card(uid, '💬 ' + chat['title'], chat['model'] or 'Модель не выбрана', keyboard(
                            [btn('✏️ Название', f'renamechat:{cid}'), btn('📜 История TXT', f'history:{cid}')],
                            [btn('🗑 Удалить', f'delchat:{cid}', 'danger'), btn('‹ Чаты', 'chats')]))
                    elif action == 'renamechat':
                        self.states[uid] = {'step': 'renamechat', 'chat': cid}
                        self.tell(uid, 'Новое название чата?')
                    elif action == 'delchat':
                        self.card(uid, 'Удалить чат?', chat['title'], keyboard([btn('Удалить', f'yeschat:{cid}', 'danger'), btn('Отмена', f'chatcfg:{cid}')]))
                    elif action == 'yeschat':
                        self.store.write('DELETE FROM chats WHERE owner=? AND id=?', (uid, cid))
                        self.store.write('UPDATE users SET current_chat=NULL WHERE id=? AND current_chat=?', (uid, cid))
                        self.chats(uid)
                    else:
                        history = self.store.history(cid, 100)
                        text = '\n\n'.join(m['role'].upper() + '\n' + (m['content'] if isinstance(m['content'], str) else '[Изображение]\n' + m['content'][0].get('text', '')) for m in history)
                        self.bot.send_document(uid, create_document(text or 'Пустой чат', 'txt'))
                elif action == 'newchat':
                    self.store.new_chat(uid)
                    self.home(uid)
                elif action in ('formats', 'saveformats'):
                    self.formats(uid, int(parts[1]) if action == 'saveformats' else None)
                elif action == 'format':
                    if parts[1] not in FORMATS:
                        return
                    self.states[uid] = {'step': 'file_prompt', 'format': parts[1]}
                    self.tell(uid, 'Что подготовить в ' + parts[1].upper() + '? Опиши содержимое. /cancel — отмена.')
                elif action == 'save':
                    cid = int(parts[2])
                    self.bot.send_document(uid, create_document(self.last_answer(uid, cid), parts[1]))
                elif action in ('rewrite', 'retry', 'speak'):
                    cid = int(parts[-1])
                    self.owned_chat(uid, cid)
                    if action == 'speak':
                        self.launch(uid, lambda job: self.speak(uid, job, cid))
                    elif action == 'retry':
                        self.launch(uid, lambda job: self.generation(uid, job, '', cid=cid, retry=True))
                    else:
                        prompts = {'short': 'Сократи предыдущий ответ, сохрани суть.', 'long': 'Раскрой предыдущий ответ подробнее с примерами.', 'simple': 'Объясни предыдущий ответ проще, понятными словами.'}
                        prompt = prompts[parts[1]]
                        self.launch(uid, lambda job: self.generation(uid, job, prompt, cid=cid))
                elif action == 'settings':
                    self.settings(uid)
                elif action == 'usage':
                    self.usage(uid)
                elif action == 'limits':
                    self.states[uid] = {'step': 'limits'}
                    self.tell(uid, 'Отправь дневной лимит токенов и USD через пробел, например: 100000 2.\n0 — без лимита. Для USD нужны тарифы в config.json. Сутки по UTC.')
                elif action == 'password':
                    self.states[uid] = {'step': 'old_password'}
                    self.tell(uid, 'Введи текущий пароль.')
                elif action == 'logout':
                    self.store.write('UPDATE users SET verified=0 WHERE id=?', (uid,))
                    self.states.pop(uid, None)
                    self.tell(uid, 'Выход выполнен. /start — войти.')
                elif action == 'admin':
                    self.admin_menu(uid)
                elif action == 'invite':
                    self.invite(uid)
                elif action == 'backup':
                    self.export(uid)
                elif action == 'restore':
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    if self.jobs:
                        raise ValueError('Перед импортом дождись завершения всех запросов.')
                    self.states[uid] = {'step': 'import_password'}
                    self.tell(uid, 'Подтверди свой пароль для импорта всей базы.')
                elif action == 'cleardb':
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    if self.jobs:
                        raise ValueError('Перед очисткой дождись завершения запросов.')
                    self.states[uid] = {'step': 'clear_password', 'expires': now() + 300}
                    self.tell(uid, '⚠️ Очистить пользователей, панели, чаты, приглашения и статистику? Перед очисткой сохранится зашифрованная копия. Подтверди личным паролем в течение 5 минут.')
                elif action == 'confirmcleardb':
                    state = self.states.get(uid, {})
                    if uid not in self.admins or state.get('step') != 'clear_confirm' or state.get('expires', 0) < now():
                        raise ValueError('Подтверждение очистки истекло. База не изменена.')
                    with self.lock:
                        if self.jobs:
                            raise ValueError('Сначала дождись завершения всех запросов.')
                        self.store.clear_database()
                        self.states.clear()
                        self.cache.clear()
                    self.tell(uid, 'База очищена. Создана зашифрованная копия в каталоге backups. /start — начать настройку заново.')
                elif action == 'confirmrestore':
                    state = self.states.get(uid, {})
                    if uid not in self.admins or state.get('step') != 'import_confirm' or state.get('expires', 0) < now():
                        raise ValueError('Подтверждение импорта истекло.')
                    with self.lock:
                        if self.jobs:
                            raise ValueError('Сначала дождись всех запросов.')
                        self.store.restore(state['blob'])
                        self.states.clear()
                        self.cache.clear()
                    self.tell(uid, 'База восстановлена. Сессии сброшены. /start — войти.')
            except ValueError as error:
                self.tell(uid, str(error))
            except Exception:
                self.tell(uid, 'Операция не выполнена. Открой /start и попробуй снова. При импорте проверь файл и исходный ключ шифрования.')

    def message(self, message):
        if message.chat.type != 'private':
            return
        uid = message.from_user.id
        with self.lock:
            lock = self.user_locks.setdefault(uid, threading.RLock())
        with lock:
            try:
                self._message(message)
            except ValueError as error:
                self.tell(uid, str(error))
            except Exception:
                self.tell(uid, 'Не удалось обработать сообщение. Проверь файл или настройки API. Автоматического повтора платных запросов нет.')

    def _message(self, message):
        uid = message.from_user.id
        text = message.text or ''
        state = self.states.get(uid, {})
        step = state.get('step')
        u = self.store.user(uid)
        auth_steps = ('register', 'register_confirm', 'login', 'old_password', 'new_password', 'new_password_confirm', 'import_password', 'clear_password')
        if step in auth_steps and message.content_type == 'text':
            self.erase(message)
            if step not in ('register', 'register_confirm', 'login') and not self.authenticated(uid):
                self.gate(uid)
                return
            if step in ('login', 'old_password', 'import_password', 'clear_password'):
                self.verify_password(uid, text)
                if step == 'old_password':
                    self.states[uid] = {'step': 'new_password'}
                    self.tell(uid, 'Новый пароль: 8–128 символов.')
                elif step == 'import_password':
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    self.states[uid] = {'step': 'import_file', 'expires': now() + 300}
                    self.tell(uid, 'Пришли .enc-копию в течение 5 минут. Нужен тот же secret.key.')
                elif step == 'clear_password':
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    self.states[uid] = {'step': 'clear_confirm', 'expires': now() + 300}
                    self.card(uid, 'Последнее подтверждение',
                              'Это удалит данные всех пользователей, панели, историю, приглашения и статистику. Перед этим будет сделана зашифрованная копия. Продолжить?',
                              keyboard([btn('🗑 Да, очистить всё', 'confirmcleardb', 'danger'),
                                        btn('↩️ Отмена', 'admin')]))
                    return
                else:
                    self.store.write('UPDATE users SET verified=? WHERE id=?', (now(), uid))
                    self.home(uid)
                return
            if not 8 <= len(text) <= 128:
                raise ValueError('Пароль должен содержать 8–128 символов.')
            if step in ('register', 'new_password'):
                salt = secrets.token_hex(16)
                self.states[uid] = {'step': step + '_confirm', 'salt': salt, 'hash': password_hash(text, salt)}
                self.tell(uid, 'Повтори пароль.')
                return
            if not hmac.compare_digest(password_hash(text, state['salt']), state['hash']):
                raise ValueError('Пароли не совпали. Повтори или /cancel.')
            if step == 'register_confirm' and u:
                raise ValueError('Аккаунт уже существует. Начни /start.')
            self.store.write('INSERT INTO users(id,salt,hash,verified) VALUES(?,?,?,?) ON CONFLICT(id) DO UPDATE SET salt=excluded.salt,hash=excluded.hash,verified=excluded.verified,failures=0,locked=0',
                             (uid, state['salt'], state['hash'], now()))
            self.home(uid)
            return
        if not self.gate(uid):
            return
        if uid in self.jobs:
            raise ValueError('Сейчас выполняется запрос. /stop — остановить.')
        if message.content_type == 'document' and step == 'import_file':
            if uid not in self.admins or state.get('expires', 0) < now():
                raise ValueError('Импорт недоступен или время истекло.')
            blob = self.download(message.document)
            payload = self.store.decode_backup(blob)
            self.states[uid] = {'step': 'import_confirm', 'blob': blob, 'expires': now() + 300}
            self.card(uid, 'Заменить всю базу?', f"Пользователей: {len(payload['users'])}\nПанелей: {len(payload['panels'])}\nПеред заменой сохранится локальная копия.",
                      keyboard([btn('Заменить БД', 'confirmrestore', 'danger'), btn('Отмена', 'home')]))
            return
        if message.content_type != 'text':
            if step:
                raise ValueError('Закончи текущую настройку или /cancel, затем пришли вложение ещё раз.')
            self.launch(uid, lambda job: self.attachment(uid, job, message))
            return
        nav = {
            '💬 Чаты': 'chats',
            '🧠 Модели': 'models',
            '📚 Панели': 'panels',
            '➕ Добавить панель': 'addpanel',
            '➕ Панель': 'addpanel',  # Older reply keyboard label.
            '📎 Файл': 'formats',
            '📄 Документ': 'formats',  # Older reply keyboard label.
            '🆕 Новый чат': 'newchat',
            '💰 Расходы': 'usage',
            '⚙️ Настройки': 'settings',
            '🛠 Админка · доступы': 'admin',
            '🛠 Админка': 'admin',  # Older reply keyboard label.
        }
        if text in nav:
            self.states.pop(uid, None)
            # Use the same authorization-aware routing as inline buttons.
            call = type('LocalCall', (), {'id': '', 'from_user': message.from_user, 'message': message, 'data': nav[text]})()
            self.callback(call)
            return
        if step == 'panel_name':
            if not text.strip() or len(text) > 60:
                raise ValueError('Название: 1–60 символов.')
            self.states[uid] = {'step': 'panel_url', 'name': text.strip()}
            self.tell(uid, 'Пришли точный API base URL из панели. Ссылки вида /chat будут преобразованы в /v1.')
        elif step == 'panel_url':
            url = normalize_url(text)
            public_host(url)
            state.update(step='panel_token', url=url)
            self.tell(uid, 'API: ' + url + '\nПришли API-ключ этой панели. Я постараюсь удалить сообщение.')
        elif step == 'panel_token':
            self.erase(message)
            encrypted = self.store.cipher.encrypt(text.strip().encode()).decode()
            panel = {'url': state['url'], 'token': encrypted}
            available = self.provider(panel).models()
            pid = self.store.write('INSERT INTO panels(owner,name,url,token) VALUES(?,?,?,?)', (uid, state['name'], state['url'], encrypted))
            self.store.write('UPDATE users SET active=?,current_chat=NULL WHERE id=?', (pid, uid))
            self.states.pop(uid, None)
            self.cache[uid] = {'panel': pid, 'all': available, 'expires': now() + 1800}
            self.models(uid)
        elif step in ('renamepanel', 'rekey'):
            panel = self.store.panel(uid, state['panel'])
            if not panel:
                raise ValueError('Панель не найдена.')
            if step == 'renamepanel':
                if not text.strip() or len(text) > 60:
                    raise ValueError('Название: 1–60 символов.')
                self.store.write('UPDATE panels SET name=? WHERE id=? AND owner=?', (text.strip(), panel['id'], uid))
            else:
                self.erase(message)
                panel['token'] = self.store.cipher.encrypt(text.strip().encode()).decode()
                self.provider(panel).models()
                self.store.write('UPDATE panels SET token=? WHERE id=? AND owner=?', (panel['token'], panel['id'], uid))
            self.states.pop(uid, None)
            self.cache.pop(uid, None)
            self.panel_settings(uid, panel['id'])
        elif step == 'searchmodel':
            self.models(uid, query=text[:100])
            self.states.pop(uid, None)
        elif step == 'renamechat':
            chat = self.owned_chat(uid, state['chat'])
            if not text.strip() or len(text) > 80:
                raise ValueError('Название чата: 1–80 символов.')
            self.store.write('UPDATE chats SET title=? WHERE id=? AND owner=?', (text.strip(), chat['id'], uid))
            self.states.pop(uid, None)
            self.chats(uid)
        elif step == 'limits':
            tokens, usd = text.replace(',', '.').split()
            tokens, usd = int(tokens), float(usd)
            if tokens < 0 or usd < 0 or not math.isfinite(usd):
                raise ValueError('Лимиты должны быть конечными неотрицательными числами.')
            self.store.write('UPDATE users SET daily_tokens=?,daily_usd=? WHERE id=?', (tokens, usd, uid))
            self.states.pop(uid, None)
            self.usage(uid)
        elif step == 'file_prompt':
            if len(text) > 20000:
                raise ValueError('Сократи запрос до 20 000 символов.')
            fmt = state['format']
            self.launch(uid, lambda job: self.generation(uid, job, text, fmt=fmt))
            self.states.pop(uid, None)
        elif step:
            self.tell(uid, 'Ожидаю текущий шаг настройки. /cancel — отмена.')
        else:
            if len(text) > 20000:
                raise ValueError('Сократи запрос до 20 000 символов.')
            requested = re.search(r'(?i)(?:сделай|создай|подготовь|пришли|сохрани|оформи|отправь).{0,100}\b(pdf|docx|xlsx|txt|csv|markdown)\b', text)
            fmt = requested.group(1).lower() if requested else None
            fmt = 'md' if fmt == 'markdown' else fmt
            self.launch(uid, lambda job: self.generation(uid, job, text, fmt=fmt))

    def attachment(self, uid, job, message):
        self.card(uid, '📎 Обрабатываю вложение', 'Фото и содержимое файла будут отправлены выбранной панели.', keyboard([btn('⏹ Стоп', 'stop', 'danger')]))
        if message.content_type in ('voice', 'audio'):
            self.transcribe(uid, job, message)
            return
        obj = message.photo[-1] if message.content_type == 'photo' else message.document
        raw = self.download(obj)
        if job.cancel.is_set():
            raise Cancelled()
        name = getattr(obj, 'file_name', None) or 'photo.jpg'
        caption = message.caption or ''
        if message.content_type == 'photo' or Path(name).suffix.lower() in ('.jpg', '.jpeg', '.png', '.webp'):
            content = image_content(raw, caption)
        else:
            extracted, clipped = read_document(raw, name)
            if clipped:
                self.tell(uid, '⚠️ Для анализа взяты первые 60 000 символов документа.')
            content = (caption or 'Проанализируй документ и выдели главное.') + '\n\nФайл: ' + name + '\n--- Начало данных ---\n' + extracted + '\n--- Конец данных ---'
        self.generation(uid, job, content)

    def maintenance(self):
        while not self.closed.is_set():
            try:
                with self.lock:
                    self.store.automatic_backup(self.cfg['backup_hours'], self.cfg['backup_keep'])
            except Exception:
                print('Не удалось создать автоматическую копию; проверь права и свободное место.', flush=True)
            self.closed.wait(300)

    def run(self):
        # A reservation surviving a process crash becomes a conservative usage record.
        for r in self.store.rows('SELECT * FROM reservations'):
            self.charge(r['owner'], r['id'], 'interrupted', 'unknown', r['tokens'], 0, r['cost'], True)
        threading.Thread(target=self.maintenance, daemon=True).start()
        try:
            self.bot.set_my_commands([types.BotCommand('start', 'Кабинет'), types.BotCommand('new', 'Новый чат'),
                                      types.BotCommand('stop', 'Остановить'), types.BotCommand('cancel', 'Отменить ввод'), types.BotCommand('help', 'Помощь')])
            print('AI Chat запущен. Секретные ключи и сообщения не записываются в журнал.', flush=True)
            self.bot.infinity_polling(skip_pending=True, timeout=30, long_polling_timeout=30, allowed_updates=['message', 'callback_query'])
        finally:
            self.closed.set()
            for job in self.jobs.values():
                job.cancel.set()
            self.pool.shutdown(wait=False, cancel_futures=True)


def main():
    parser = argparse.ArgumentParser(description='AI Chat — Telegram AI client')
    parser.add_argument('--setup', action='store_true', help='Configure credentials outside the repository')
    args = parser.parse_args()
    cfg = configuration(args.setup)
    if args.setup:
        print('Настройки сохранены вне репозитория: ' + str(cfg['data_path']))
        return
    AIChat(cfg).run()


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print('AI Chat остановлен.')
    except Exception:
        print('Не удалось запустить AI Chat. Проверь config.json, зависимости и локальные настройки (--setup). Секреты не выводятся.')
        raise SystemExit(1)
