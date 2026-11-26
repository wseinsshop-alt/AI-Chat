"""AI Chat — Telegram AI client for private conversations and shared panels. Python 3.10+.

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
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from docx import Document
from openpyxl import Workbook, load_workbook
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

ROOT = Path(__file__).resolve().parent
FORMATS = ('pdf', 'docx', 'xlsx', 'txt', 'md', 'csv', 'pptx', 'json', 'html', 'zip')
DEFAULTS = {'invite_only': True, 'session_hours': 24, 'backup_hours': 24,
            'backup_keep': 7, 'daily_tokens': 100000, 'daily_usd': 0,
            'max_output_tokens': 4096, 'history_messages': 20, 'max_file_mb': 10,
            'workers': 4, 'prices': {}, 'image_prices': {}, 'image_model': '', 'transcription_model': 'whisper-1',
            'transcription_usd_per_minute': None,
            'system_prompt': 'Ты полезный помощник в Telegram. Не выдумывай выполненные действия.'}


def now() -> float:
    return time.time()


def day() -> str:
    return datetime.now(timezone.utc).date().isoformat()


def telegram_command(text):
    value = re.sub(r'^[\s\u200b\u200e\u200f\ufeff]+', '', str(text or ''))
    if not value.startswith('/'):
        return '', ''
    parts = value.split(maxsplit=1)
    return parts[0].split('@',1)[0].casefold(), parts[1] if len(parts)>1 else ''


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
        ids = input('Telegram ID администраторов через запятую (Enter — сохранить текущие): ').strip()
        if ids:
            cfg['admin_ids'] = [int(x.strip()) for x in ids.split(',')]
        if not cfg['telegram_token'] or not cfg['admin_ids']:
            raise ValueError('Нужны токен и хотя бы один ID администратора.')
        if not re.fullmatch(r'[0-9]+:[A-Za-z0-9_-]{20,}', cfg['telegram_token']):
            raise ValueError('Неверный формат токена. Скопируй полный токен из BotFather.')
        if any(not isinstance(uid, int) or uid <= 0 for uid in cfg['admin_ids']):
            raise ValueError('Telegram ID администратора должен быть положительным числом.')
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
CREATE TABLE IF NOT EXISTS invites(code TEXT PRIMARY KEY,uses INTEGER NOT NULL,expires REAL NOT NULL,days INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS panel_shares(panel INTEGER NOT NULL REFERENCES panels(id) ON DELETE CASCADE,
 friend INTEGER NOT NULL,PRIMARY KEY(panel,friend));
CREATE TABLE IF NOT EXISTS preferences(owner INTEGER PRIMARY KEY,style TEXT NOT NULL DEFAULT 'balanced');
CREATE TABLE IF NOT EXISTS panel_choices(owner INTEGER NOT NULL,panel INTEGER NOT NULL REFERENCES panels(id) ON DELETE CASCADE,
 model TEXT NOT NULL,PRIMARY KEY(owner,panel));
CREATE TABLE IF NOT EXISTS ui_sessions(owner INTEGER PRIMARY KEY,payload TEXT NOT NULL,expires REAL NOT NULL);
CREATE TABLE IF NOT EXISTS admitted(id INTEGER PRIMARY KEY);
CREATE TABLE IF NOT EXISTS subscriptions(id INTEGER PRIMARY KEY,expires REAL NOT NULL DEFAULT 0,
 blocked INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS usage(id INTEGER PRIMARY KEY AUTOINCREMENT,owner INTEGER,day TEXT,kind TEXT,
 model TEXT,input_tokens INTEGER,output_tokens INTEGER,cost REAL,estimated INTEGER,created REAL);
CREATE TABLE IF NOT EXISTS reservations(id TEXT PRIMARY KEY,owner INTEGER,day TEXT,tokens INTEGER,cost REAL);
'''


class Store:
    TABLES = ('users', 'panels', 'chats', 'messages', 'favorites', 'invites', 'admitted', 'subscriptions', 'usage', 'panel_shares', 'preferences', 'panel_choices')

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
            migrate_subscriptions = not c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='subscriptions'").fetchone()
            c.executescript(SCHEMA)
            if migrate_subscriptions:
                c.execute('INSERT OR IGNORE INTO subscriptions(id) SELECT id FROM users')
                c.execute('INSERT OR IGNORE INTO subscriptions(id) SELECT id FROM admitted')
            if 'days' not in [r['name'] for r in c.execute('PRAGMA table_info(invites)')]:
                c.execute('ALTER TABLE invites ADD COLUMN days INTEGER NOT NULL DEFAULT 0')
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

    def subscription(self, uid):
        return self.one('SELECT * FROM subscriptions WHERE id=?', (uid,))

    def has_subscription(self, uid):
        row = self.subscription(uid)
        return bool(row and not row['blocked'] and (row['expires'] == 0 or row['expires'] > now()))

    def grant_subscription(self, uid, days):
        if uid <= 0 or days not in (0, 7, 30, 90):
            raise ValueError('Нужны положительный Telegram ID и срок 7, 30, 90 дней или без срока.')
        with self.db() as c:
            c.execute('BEGIN IMMEDIATE')
            old = c.execute('SELECT * FROM subscriptions WHERE id=?', (uid,)).fetchone()
            base = max(now(), old['expires']) if old and not old['blocked'] else now()
            expires = 0 if days == 0 else base + days * 86400
            # A finite extension must not shorten an existing unlimited grant.
            if old and not old['blocked'] and old['expires'] == 0:
                expires = 0
            c.execute('INSERT OR IGNORE INTO admitted(id) VALUES(?)', (uid,))
            c.execute('INSERT INTO subscriptions(id,expires,blocked) VALUES(?,?,0) '
                      'ON CONFLICT(id) DO UPDATE SET expires=excluded.expires,blocked=0', (uid, expires))
        return expires

    def panel(self, uid, pid=None):
        if pid is None:
            u = self.user(uid)
            pid = u['active'] if u else None
        p = self.one('SELECT * FROM panels WHERE id=?', (pid,))
        if p and hasattr(self, 'admin_ids') and p['owner'] not in self.admin_ids:
            return None
        if p and (p['owner'] == uid or (self.has_subscription(uid) and self.one(
                'SELECT 1 FROM panel_shares WHERE panel=? AND friend=?', (pid, uid)))):
            return p
        return None

    def available_panels(self, uid):
        return [p for p in self.rows('SELECT * FROM panels ORDER BY id') if self.panel(uid, p['id'])]

    def preferred_model(self, uid, panel):
        choice = self.one('SELECT model FROM panel_choices WHERE owner=? AND panel=?', (uid, panel['id']))
        return choice['model'] if choice else panel['model']

    def answer_style(self, uid):
        row = self.one('SELECT style FROM preferences WHERE owner=?', (uid,))
        return row['style'] if row and row['style'] in ('balanced', 'short', 'detailed') else 'balanced'

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
                         (uid, p['id'], 'Новый чат', self.preferred_model(uid, p), now()))
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
            payload = {'version': 5, 'created': now()}
            for table in self.TABLES:
                payload[table] = [dict(r) for r in c.execute('SELECT * FROM ' + table)]
        return self.cipher.encrypt(json.dumps(payload, ensure_ascii=False).encode())

    def decode_backup(self, blob):
        payload = json.loads(self.cipher.decrypt(blob))
        if payload.get('version') == 2:
            existing = {int(r['id']) for r in payload.get('users', []) + payload.get('admitted', [])}
            payload['subscriptions'] = [{'id': uid, 'expires': 0, 'blocked': 0} for uid in existing]
            for invite in payload.get('invites', []):
                invite.setdefault('days', 0)
            payload['version'] = 3
        if payload.get('version') == 3:
            payload['panel_shares'] = []
            payload['version'] = 4
        if payload.get('version') == 4:
            payload['preferences'] = []
            payload['panel_choices'] = []
            payload['version'] = 5
        if payload.get('version') != 5 or not all(isinstance(payload.get(t), list) for t in self.TABLES):
            raise ValueError('Нужна резервная копия AI Chat версии 2/3/4/5 и исходный secret.key.')
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
            c.execute('DELETE FROM ui_sessions')
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
            c.execute('DELETE FROM ui_sessions')
            for table in reversed(self.TABLES):
                c.execute('DELETE FROM ' + table)
            if c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='history'").fetchone():
                c.execute('DELETE FROM history')

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


class ProviderQuotaError(ValueError):
    pass


class Provider:
    def __init__(self, panel, cipher):
        self.url = panel['url']
        self.key = cipher.decrypt(panel['token'].encode()).decode()

    def request(self, path, payload=None, files=None, stream=False, read_timeout=60):
        address = public_host(self.url)
        parsed = urlsplit(self.url)
        session = requests.Session()
        session.trust_env = False
        session.mount('https://', PinnedHTTPSAdapter(parsed.hostname, address, parsed.port or 443))
        try:
            response = session.request('GET' if payload is None and files is None else 'POST', self.url + path,
                                       headers={'Authorization': 'Bearer ' + self.key, 'Host': parsed.netloc},
                                       json=payload if files is None else None, data=payload if files else None,
                                       files=files, stream=True, timeout=(15, read_timeout), allow_redirects=False)
            if response.status_code != 200:
                status = response.status_code
                exhausted = status == 402
                if status in (400,403,429):
                    try:
                        info = json.loads(bounded_response(response,65536))
                        error = info.get('error',info)
                        code = str(error.get('code','')) if isinstance(error,dict) else ''
                        exhausted = code in ('insufficient_quota','quota_exceeded','insufficient_balance',
                            'insufficient_funds','credit_balance_too_low','balance_exhausted','tokens_exhausted')
                    except Exception:
                        pass
                response.close()
                if exhausted:
                    raise ProviderQuotaError('Баланс или квота API-панели исчерпаны.')
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
            try:
                return json.loads(raw)
            except (ValueError,UnicodeError):
                raise ValueError('Панель вернула не JSON. Проверь именно API base URL, а не адрес страницы чата.') from None
        finally:
            r.close()
            r._ai_session.close()

    def model_records(self):
        result = self.json('/models')
        found = result.get('data') if isinstance(result,dict) else None
        if not isinstance(found,list):
            raise ValueError('API не вернул список моделей data. Проверь совместимость панели и API-адрес.')
        unique = {}
        for m in found:
            if isinstance(m,dict) and isinstance(m.get('id'),str) and m['id'].strip():
                unique.setdefault(m['id'],m)
        return list(unique.values())

    def models(self):
        found = self.model_records()
        result = [str(m['id']) for m in found if isinstance(m, dict) and 'id' in m]
        if not result:
            raise ValueError('Панель вернула пустой список моделей.')
        return result




def readable_name(value):
    return re.sub(r'[\s\u200b-\u200f\ufeff]+', ' ', str(value or '')).strip()


def universal_model(records):
    """Prefer declared image-input support; model-name hints are only a fallback."""
    ranked = []
    for item in records:
        model = item.get('id','')
        lower = model.casefold()
        if re.search(r'(embed|rerank|whisper|transcri|tts|gpt-image|dall-e|flux|imagen|sdxl|moderation|sora)',lower):
            continue
        architecture = item.get('architecture') or {}
        if not isinstance(architecture,dict):
            architecture = {}
        inputs = item.get('input_modalities', architecture.get('input_modalities'))
        outputs = item.get('output_modalities',architecture.get('output_modalities'))
        if outputs and 'text' not in outputs:
            continue
        capabilities = item.get('capabilities') or {}
        if isinstance(capabilities,list):
            capabilities = {'vision':'vision' in capabilities}
        if not isinstance(capabilities,dict):
            capabilities = {}
        declared = bool(inputs is not None or 'vision' in capabilities or 'supports_vision' in item)
        vision = ('image' in inputs if isinstance(inputs,list) else False) or capabilities.get('vision') is True or item.get('supports_vision') is True
        hint = bool(re.search(r'(vision|[-/]vl\b|gpt-4o|gpt-4\.1|gpt-5|gemini|claude-3|claude-(sonnet|opus|haiku)-4)',lower))
        score = 10000 if vision else (1000 if hint and not declared else 0)
        score += 100 if any(name in lower for name in ('gpt','claude','gemini')) else 0
        score += 30 if any(name in lower for name in ('pro','sonnet','gpt-5')) else 0
        score -= 10 if any(name in lower for name in ('mini','nano','lite')) else 0
        score -= 30 if any(name in lower for name in ('preview','experimental')) else 0
        ranked.append((score,model))
    return sorted(ranked,key=lambda pair:(-pair[0],pair[1]))[0][1] if ranked else ''


def bounded_response(response, limit, cancel=None):
    result = bytearray()
    for chunk in response.iter_content(65536):
        if cancel is not None and cancel.is_set():
            raise Cancelled()
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



def output_parts(message):
    """Normalize text and image fields used by OpenAI-compatible chat panels."""
    text, images = [], []
    def image(value):
        if isinstance(value, dict):
            if value.get('b64_json'):
                images.append('data:image/png;base64,' + value['b64_json'])
                return
            value = value.get('image_url', value.get('url', ''))
        if isinstance(value, dict):
            value = value.get('url', '')
        if isinstance(value, str) and value:
            images.append(value)
    content = message.get('content', '')
    if isinstance(content, str):
        text.append(content)
    elif isinstance(content, list):
        for item in content:
            if not isinstance(item, dict):
                continue
            if item.get('type') in ('text', 'output_text'):
                value = item.get('text', '')
                if isinstance(value, str):
                    text.append(value)
            elif item.get('type') in ('image', 'image_url', 'output_image'):
                image(item)
    for item in message.get('images', []) or []:
        image(item)
    return ''.join(text), images


def picture_bytes(value, cancel):
    """Download a bounded image from a public HTTPS URL, never with panel credentials."""
    limit = 15 * 1024 * 1024
    if cancel.is_set():
        raise Cancelled()
    if value.startswith('data:image/'):
        if len(value) > limit * 4 // 3 + 256 or ';base64,' not in value:
            raise ValueError('Изображение слишком большое или имеет неверный формат.')
        try:
            return base64.b64decode(value.split(',', 1)[1], validate=True)
        except Exception:
            raise ValueError('Модель вернула некорректное изображение.') from None
    parsed = urlsplit(value)
    if parsed.scheme != 'https' or not parsed.hostname or parsed.username or parsed.password or parsed.fragment:
        raise ValueError('Изображение должно быть доступно по публичному HTTPS.')
    # Validate DNS without rejecting signed CDN query parameters.
    base = urlunsplit((parsed.scheme, parsed.netloc, parsed.path, '', ''))
    address = public_host(base)
    session = requests.Session()
    session.trust_env = False
    session.mount('https://', PinnedHTTPSAdapter(parsed.hostname, address, parsed.port or 443))
    response = None
    try:
        response = session.get(value, headers={'Host': parsed.netloc}, stream=True,
                               timeout=(15, 60), allow_redirects=False)
        if response.status_code != 200:
            raise ValueError('Не удалось скачать изображение от модели.')
        return bounded_response(response, limit, cancel)
    finally:
        if response is not None:
            response.close()
        session.close()


def photo_buffer(raw):
    """Send a Telegram-friendly preview; decode only bounded raster images."""
    with Image.open(io.BytesIO(raw)) as source:
        if source.width * source.height > 24000000 or source.width < 1 or source.height < 1:
            raise ValueError('Лимит изображения — 24 мегапикселя.')
        source.seek(0)
        image = ImageOps.exif_transpose(source).convert('RGBA')
        image.thumbnail((2048, 2048))
        canvas = Image.new('RGB', image.size, 'white')
        canvas.paste(image, mask=image.getchannel('A'))
        # Telegram rejects extremely narrow photographs.
        if max(canvas.size) > min(canvas.size) * 20:
            square = Image.new('RGB', (max(canvas.size), max(canvas.size)), 'white')
            square.paste(canvas, ((square.width-canvas.width)//2, (square.height-canvas.height)//2))
            canvas = square
        out = io.BytesIO()
        canvas.save(out, 'JPEG', quality=90)
    out.seek(0)
    out.name = 'AI-Chat.jpg'
    return out


def requested_output(text):
    """Explicit format wins over media mentioned inside the document request."""
    intent = re.search(r'(?i)\b(?:сделай|создай|сформируй|сгенерируй|подготовь|'
                       r'пришли|пришлите|скинь|скиньте|сбрось|сохрани|сохраните|'
                       r'оформи|отправь|отправьте|выгрузи|экспортируй|'
                       r'create|send|save|export|make)\b', text)
    document = re.search(r'(?i)\b(pdf|пдф(?:ку|ке|чик|чиком)?|docx|doc|word|ворд|'
                         r'xlsx|xls|excel|эксель|txt|csv|markdown|md|pptx|ppt|powerpoint|'
                         r'паверпоинт|поверпоинт|json|html|zip|зип)\b', text)
    if document and intent:
        mentions = list(document.re.finditer(text))
        if any(re.fullmatch(r'\s*(?:,|и|и ещё|а также)\s*(?:в\s+)?', text[left.end():right.start()], re.I)
               for left, right in zip(mentions, mentions[1:])):
            return None  # Let the structured file protocol return all requested formats.
        kind = document.group(1).lower()
        if kind.startswith('пдф'):
            return 'pdf'
        return {'doc': 'docx', 'word': 'docx', 'ворд': 'docx', 'xls': 'xlsx',
                'excel': 'xlsx', 'эксель': 'xlsx', 'markdown': 'md',
                'ppt': 'pptx', 'powerpoint': 'pptx', 'паверпоинт': 'pptx',
                'поверпоинт': 'pptx', 'зип': 'zip'}.get(kind, kind)
    if re.match(r'(?is)^\s*(?:нарисуй(?:те)?|изобрази(?:те)?|draw\b|'
                r'(?:создай|сгенерируй|сделай|пришли|отправь|скинь)\s+(?:мне\s+)?'
                r'(?:картинку|изображение|фото(?:графию|графии|чку)?|фотку|иллюстрацию|рисунок)|'
                r'(?:create|generate)\s+(?:an?\s+)?image\b)', text):
        return 'image'
    if intent:
        if re.search(r'(?i)\bпрезентаци[яюи]\b', text):
            return 'pptx'
        if re.search(r'(?i)\b(?:таблицу|табличку)\b', text):
            return 'xlsx'
        if re.search(r'(?i)\b(?:архив|архивом)\b', text):
            return 'zip'
        if re.search(r'(?i)\bдокумент(?:ом)?\b', text):
            return 'pdf'
        if re.search(r'(?i)\b(?:файлом|файл)\b', text):
            return 'txt'
    return None


def clean_json(text):
    return re.sub(r'^' + r'\x60{3}(?:json)?\s*|\s*\x60{3}$', '', text.strip())


def artifact_name(name, fmt):
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', '_', str(name or 'AI-Chat'))
    name = name.strip(' .')[:100] or 'AI-Chat'
    stem = name.rsplit('.', 1)[0] if '.' in name else name
    return stem + '.' + fmt


def response_artifacts(text):
    """Decode file declarations as data only; never run generated code."""
    files = []
    def extract(match):
        payload = json.loads(match.group(1))
        items = payload.get('files') if isinstance(payload, dict) else None
        if not isinstance(items, list) or not 1 <= len(items) <= 4:
            raise ValueError('Модель должна вернуть от 1 до 4 файлов.')
        for item in items:
            if isinstance(item, dict) and isinstance(item.get('format'), str):
                item['format'] = item['format'].lower()
            if not isinstance(item, dict) or item.get('format') not in FORMATS:
                raise ValueError('Модель запросила неподдерживаемый формат файла.')
            content = item.get('content')
            if isinstance(content, (dict, list)):
                content = json.dumps(content, ensure_ascii=False)
            if not isinstance(content, str) or len(content) > 120000:
                raise ValueError('Модель вернула слишком большое или неверное содержимое файла.')
            files.append((artifact_name(item.get('name'), item['format']), item['format'], content))
        if len(files) > 4 or sum(len(f[2]) for f in files) > 120000:
            raise ValueError('Слишком много файлов или текста в одном ответе.')
        return ''
    try:
        visible = re.sub(r'(?s)\x60{3}ai-files\s*\n(.*?)\n\x60{3}', extract, text)
    except (json.JSONDecodeError, TypeError):
        raise ValueError('Ответ с файлами обрезан или имеет неверный JSON. Попроси модель выдать файл ещё раз.') from None
    if re.search(r'\x60{3}ai-files\b', visible):
        raise ValueError('Модель не закончила описание файлов. Попроси уменьшить объём или увеличить лимит ответа.')
    return visible.strip(), files


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
    elif ext == '.pptx':
        check_zip(raw)
        deck = Presentation(io.BytesIO(raw))
        if len(deck.slides) > 60:
            raise ValueError('Лимит — 60 слайдов PPTX.')
        for index, slide in enumerate(deck.slides, 1):
            parts.append('Слайд ' + str(index))
            for shape in slide.shapes:
                if shape.has_text_frame:
                    parts.append(shape.text)
                elif shape.has_table:
                    parts.extend(' | '.join(cell.text for cell in row.cells) for row in shape.table.rows)
            if sum(map(len, parts)) > 60000:
                break
    elif ext in ('.txt', '.md', '.csv', '.json', '.html'):
        try:
            parts = [raw.decode('utf-8-sig')]
        except UnicodeDecodeError:
            parts = [raw.decode('cp1251')]
    else:
        raise ValueError('Поддерживаются PDF, DOCX, XLSX, PPTX, TXT, MD, CSV, JSON, HTML и изображения.')
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
    elif fmt == 'json':
        value = json.loads(clean_json(text))
        out.write(json.dumps(value, ensure_ascii=False, indent=2, allow_nan=False).encode())
    elif fmt == 'html':
        paragraphs = '\n'.join('<p>' + html.escape(line) + '</p>' for line in text.splitlines())
        out.write(('<!doctype html><html lang="ru"><meta charset="utf-8"><title>AI Chat</title>'
                   '<style>body{max-width:900px;margin:40px auto;padding:24px;font:18px/1.6 sans-serif}'
                   'p{white-space:pre-wrap;overflow-wrap:anywhere}</style><body>' +
                   paragraphs + '</body></html>').encode())
    elif fmt == 'zip':
        bundle = json.loads(clean_json(text))
        items = bundle.get('files') if isinstance(bundle, dict) else None
        if not isinstance(items, list) or not 1 <= len(items) <= 20:
            raise ValueError('Для ZIP нужны от 1 до 20 текстовых файлов в JSON.')
        seen, size = set(), 0
        with zipfile.ZipFile(out, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
            for item in items:
                if not isinstance(item, dict):
                    raise ValueError('Некорректная запись файла в архиве.')
                name, content = item.get('name'), item.get('content')
                if (not isinstance(name, str) or not name.strip(' .') or name in ('.', '..')
                        or re.search(r'[\\/:*?"<>|\x00-\x1f]', name) or len(name) > 100
                        or name.casefold() in seen or not isinstance(content, str)):
                    raise ValueError('В архиве нужны уникальные простые имена и текстовое содержимое.')
                size += len(content.encode())
                if size > 1024 * 1024:
                    raise ValueError('Слишком большой архив.')
                seen.add(name.casefold())
                archive.writestr(name, content)
    elif fmt == 'pptx':
        try:
            payload = json.loads(clean_json(text))
            slides = payload.get('slides') if isinstance(payload, dict) else None
        except json.JSONDecodeError:
            chunks = re.split(r'\n(?=#{1,2}\s)', text)
            slides = [{'title': chunk.splitlines()[0].lstrip('# ').strip()[:80],
                       'points': [line.lstrip('-* ').strip() for line in chunk.splitlines()[1:] if line.strip()][:6]}
                      for chunk in chunks if chunk.strip()]
        if not isinstance(slides, list) or not 1 <= len(slides) <= 20:
            raise ValueError('Для PPTX нужны от 1 до 20 слайдов.')
        deck = Presentation()
        deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
        for index, item in enumerate(slides, 1):
            if not isinstance(item, dict):
                raise ValueError('Некорректное содержимое слайда.')
            title, points = item.get('title', ''), item.get('points', [])
            if not isinstance(title, str) or not isinstance(points, list) or len(points) > 6:
                raise ValueError('На слайде нужен заголовок и до 6 пунктов.')
            if any(not isinstance(p, str) or len(p) > 180 for p in points) or len(title) > 100:
                raise ValueError('Сократи текст презентации: заголовок до 100, пункт до 180 символов.')
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor.from_string('F4F7FC')
            head = slide.shapes.add_textbox(Inches(.7), Inches(.6), Inches(11.9), Inches(1.2)).text_frame
            head.word_wrap = True
            head.text = title
            head.paragraphs[0].font.size = Pt(30)
            head.paragraphs[0].font.bold = True
            head.paragraphs[0].font.color.rgb = RGBColor.from_string('1746A2')
            body = slide.shapes.add_textbox(Inches(.8), Inches(2), Inches(11.6), Inches(4.6)).text_frame
            body.word_wrap = True
            for i, point in enumerate(points):
                p = body.paragraphs[0] if i == 0 else body.add_paragraph()
                p.text = '• ' + point
                p.font.size = Pt(21)
                p.space_after = Pt(18)
                p.font.color.rgb = RGBColor.from_string('243247')
            footer = slide.shapes.add_textbox(Inches(.8), Inches(7), Inches(11), Inches(.3)).text_frame
            footer.text = f'AI CHAT  ·  {index:02d}'
            footer.paragraphs[0].font.size = Pt(10)
        deck.save(out)
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
    token: str = field(default_factory=lambda: secrets.token_hex(8))


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
        self.store.admin_ids = self.admins
        self.reply_routes = {}
        self.reply_labels = {}
        self.catalog_cache = {}
        self.states = {}
        self.cache = {}
        self.jobs = {}
        self.lock = threading.RLock()
        self.user_locks = {}
        self.menu_ids = {}
        self.menu_lock = threading.RLock()
        self.pool = ThreadPoolExecutor(max_workers=config['workers'])
        self.closed = threading.Event()
        self.bot.register_message_handler(self.command, func=lambda message: telegram_command(message.text)[0]=='/start', content_types=['text'])
        self.bot.register_callback_query_handler(self.callback, func=lambda call: True)
        self.bot.register_message_handler(self.message, content_types=['text', 'photo', 'document', 'voice', 'audio'])

    @staticmethod
    def model_label(model):
        raw = readable_name(model)
        name = raw.rsplit('/',1)[-1]
        families = {'gpt':'GPT','chatgpt':'ChatGPT','claude':'Claude','gemini':'Gemini',
                    'deepseek':'DeepSeek','qwen':'Qwen','llama':'Llama','mistral':'Mistral'}
        known = False
        for prefix, label in families.items():
            if re.match(r'^'+prefix+r'(?=[\d_.-]|$)',name,re.I):
                name = label + name[len(prefix):]
                known = True
                break
        if not known:
            return raw or 'Не выбрана'
        # Only Claude's documented short version notation; never rewrite date suffixes.
        if name.startswith('Claude'):
            name = re.sub(r'(?<=-)(\d)-(\d)(?=-(?:sonnet|haiku|opus)|$)',r'\1.\2',name,flags=re.I)
        for word in ('mini','nano','flash','pro','sonnet','haiku','opus','vision','vl','preview'):
            name = re.sub(r'(?i)(?<!\w)'+word+r'(?!\w)',word.upper() if word=='vl' else word.capitalize(),name)
        return name

    @staticmethod
    def compact_label(value, limit=90):
        value = readable_name(value)
        if len(value)<=limit:
            return value
        return value[:limit-26]+'…'+value[-25:]

    @staticmethod
    def panel_labels(panels):
        result = {}
        groups = {}
        for p in panels:
            groups.setdefault(readable_name(p['name']) or 'Панель',[]).append(p)
        for name, group in groups.items():
            group.sort(key=lambda p:p['id'])
            hosts = [urlsplit(p['url']).hostname or '' for p in group]
            for i,p in enumerate(group):
                label = name
                if len(group)>1:
                    if len(set(hosts))==len(hosts):
                        label += ' · '+hosts[i]
                    elif i==0:
                        label += ' · основная'
                    else:
                        label += ' · резерв '+chr(ord('А')+i-1)
                result[p['id']] = label
        return result

    def load_routes(self, uid):
        if uid not in self.reply_routes:
            row = self.store.one('SELECT payload FROM ui_sessions WHERE owner=?',(uid,))
            try:
                routes = json.loads(row['payload']) if row else {}
                if not isinstance(routes,dict) or any(not isinstance(k,str) or not isinstance(v,str) for k,v in routes.items()):
                    routes = {}
            except (ValueError,TypeError):
                routes = {}
            self.reply_routes[uid] = routes
            self.reply_labels.setdefault(uid,set()).update(routes)
        return self.reply_routes[uid]

    def select_model(self, uid, pid, model):
        if not self.store.panel(uid,pid):
            raise ValueError('Доступ к этой панели отозван.')
        if not model or len(model)>1000:
            raise ValueError('Неверное название модели.')
        u = self.store.user(uid)
        old = self.store.one('SELECT * FROM chats WHERE id=? AND owner=?',(u['current_chat'],uid))
        if old and not self.store.panel(uid,old['panel']):
            old = None
        with self.store.db() as c:
            c.execute('UPDATE users SET active=? WHERE id=?',(pid,uid))
            c.execute('INSERT INTO panel_choices VALUES(?,?,?) ON CONFLICT(owner,panel) DO UPDATE SET model=excluded.model',(uid,pid,model))
            if old:
                c.execute('UPDATE chats SET panel=?,model=?,updated=? WHERE id=? AND owner=?',(pid,model,now(),old['id'],uid))
        if not old:
            chat = self.store.new_chat(uid)
            self.store.write('UPDATE chats SET model=? WHERE id=?',(model,chat['id']))
        self.resume(uid)

    def tell(self, uid, text, keys=None, title=None):
        step = self.states.get(uid, {}).get('step')
        if keys is None and step and step not in (
                'register', 'register_confirm', 'login', 'old_password', 'new_password',
                'new_password_confirm', 'import_password', 'clear_password'):
            keys = keyboard([btn('↩ Отмена', 'home'), btn('❔ Помощь', 'help')])
        if keys is not None:
            keys = self.bottom(uid, keys)
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
                              parse_mode='HTML', reply_markup=self.bottom(uid, keys))

    def main_keys(self, uid):
        rows = [[btn('🧠 Модели','models'),btn('💬 Чаты','chats')],
                [btn('✚ Новый чат','newchat'),btn('⭐ Избранное','models:favorites')],
                [btn('⚙️ Настройки','settings'),btn('🎟 Доступ','subscription')]]
        if uid in self.admins:
            rows.append([btn('🛠 Админка','admin'),btn('📚 Панели','panels')])
        return keyboard(*rows)

    def bottom(self, uid, keys=None):
        with self.lock:
            lock = self.user_locks.setdefault(uid,threading.RLock())
        with lock:
            return self._bottom(uid,keys)

    def _bottom(self, uid, keys=None):
        keys = keys or self.main_keys(uid)
        items = [item for row in keys.to_dict().get('inline_keyboard',[]) for item in row
                 if item.get('callback_data') not in ('home','stop')]
        items += [{'text':'🏠 Меню','callback_data':'home','style':'primary'}]
        routes, buttons = {}, []
        for item in items:
            label = readable_name(item['text'])
            if label and label[0].isalnum():
                label = '› '+label
            action = item['callback_data']
            if label in routes:
                if routes[label]==action:
                    continue
                raise ValueError('Повтор названия кнопки. Открой /start и обнови список.')
            routes[label] = action
            buttons.append(types.KeyboardButton(label,style=item.get('style','primary')))
        if len(buttons)%2:
            label = '❔ Как пользоваться'
            routes[label] = 'help'
            buttons.append(types.KeyboardButton(label,style='primary'))
        board = types.ReplyKeyboardMarkup(resize_keyboard=True,is_persistent=True,row_width=2,
                                         input_field_placeholder='Сообщение, фото или голосовое…')
        for i in range(0,len(buttons),2):
            board.row(*buttons[i:i+2])
        self.store.write('INSERT INTO ui_sessions VALUES(?,?,?) ON CONFLICT(owner) DO UPDATE SET payload=excluded.payload,expires=excluded.expires',
                         (uid,json.dumps(routes,ensure_ascii=False),now()+7*86400))
        self.reply_routes[uid] = routes
        self.reply_labels.setdefault(uid,set()).update(routes)
        return board

    @staticmethod
    def page_nav(page, pages, prefix):
        if pages <= 1:
            return []
        return [[btn('← Назад', f'{prefix}:{page-1}' if page else 'noop'),
                 btn('Далее →', f'{prefix}:{page+1}' if page+1 < pages else 'noop')]]

    def prepare_panel(self, uid):
        p = self.store.panel(uid)
        if not p:
            available = self.store.available_panels(uid)
            if available:
                p = available[0]
                self.store.write('UPDATE users SET active=?,current_chat=NULL WHERE id=?', (p['id'], uid))
        return p

    def erase(self, msg):
        try:
            self.bot.delete_message(msg.chat.id, msg.message_id)
        except Exception:
            pass

    def authenticated(self, uid):
        return bool(self.store.user(uid))

    def gate(self, uid):
        if not self.store.user(uid):
            allowed = uid in self.admins or not self.cfg['invite_only'] or self.store.one('SELECT id FROM admitted WHERE id=?',(uid,))
            if not allowed:
                self.tell(uid,f'🔒 Доступ по приглашению. Твой Telegram ID: {uid}. Попроси администратора добавить тебя.')
                return False
            self.store.write('INSERT INTO users(id,salt,hash,verified) VALUES(?,?,?,?)',(uid,'','',0))
        return True

    def subscription_label(self, uid):
        if uid in self.admins:
            return 'Администратор · без срока'
        row = self.store.subscription(uid)
        if not row:
            return 'Ожидает выдачи доступа администратором'
        if row['blocked']:
            return 'Доступ отключён администратором'
        if not row['expires']:
            return 'Активна · без срока'
        stamp = datetime.fromtimestamp(row['expires'], timezone.utc).strftime('%d.%m.%Y %H:%M UTC')
        return ('Активна до ' if row['expires'] > now() else 'Истекла ') + stamp

    def require_subscription(self, uid):
        if uid not in self.admins and not self.store.has_subscription(uid):
            raise ValueError('Доступ к ИИ не активен. Открой «🎟 Мой доступ» и попроси администратора выдать или продлить подписку.')

    def subscription_info(self, uid):
        self.card(uid, '🎟 Мой доступ', self.subscription_label(uid) +
                  f'\nТвой Telegram ID: {uid}\n\nДля выданной панели нужны и подписка, и разрешение администратора. Чаты личные, расход идёт с баланса владельца панели.',
                  keyboard([btn('📚 Панели', 'panels'), btn('🏠 Меню', 'home')]))

    def verify_password(self, uid, text):
        raise ValueError('Вход по паролю отключён. Используй /start.')

    def home(self, uid):
        self.states.pop(uid, None)
        p = self.prepare_panel(uid)
        body = 'Выбери модель и просто общайся: текст, фото и голосовые — в этом же чате. Файлы и изображения приходят сюда же.'
        if p:
            chat = self.store.chat(uid)
            model = chat['model'] or self.store.preferred_model(uid, p)
            body = f"Модель: {self.model_label(model) if model else 'не выбрана'}\nЧат: {chat['title']}\n\n" + body
        else:
            body += '\n\n' + ('Добавь панели через админку.' if uid in self.admins else 'Попроси администратора выдать доступ к моделям.')
        self.card(uid, 'AI CHAT', body, self.main_keys(uid))

    def resume(self, uid):
        if uid not in self.admins and not self.store.has_subscription(uid):
            self.subscription_info(uid)
            return
        p = self.prepare_panel(uid)
        if not p:
            self.home(uid)
            return
        chat = self.store.chat(uid)
        model = chat['model'] or self.store.preferred_model(uid, p)
        if not model:
            self.models(uid)
            return
        self.card(uid, '💬 Продолжай разговор',
                  f"Модель: {self.model_label(model)}\nЧат: {chat['title']}\n\nПиши, присылай фото или голосовые. Картинки и файлы появятся в этой переписке.",
                  self.main_keys(uid))

    def help_menu(self, uid):
        self.card(uid, '❔ Как пользоваться',
                  'Выбери модель снизу и пиши как в обычном чате. Можно присылать фото, документы и голосовые. Проси картинки, PDF, таблицы или презентации прямо в сообщении.\n\nМодель можно сменить, сохранив текущий разговор. «Новый чат» начинает отдельную личную историю. Панелями управляет только администратор.\n\nЕдинственная команда — /start. Меню доступно снизу; «Стоп» появляется только у текущего сообщения ожидания. Поддержка изображений и распознавания речи зависит от API-панели.',
                  self.main_keys(uid))

    def panels(self, uid, page=0):
        if uid not in self.admins:
            self.models(uid)
            return
        panels = self.store.rows('SELECT * FROM panels WHERE owner=? ORDER BY id',(uid,))
        labels = self.panel_labels(panels)
        pages = max(1,math.ceil(len(panels)/6))
        page = max(0,min(page,pages-1))
        choices = [btn('📚 '+self.compact_label(labels[p['id']]),f'panelcfg:{p["id"]}') for p in panels[page*6:page*6+6]]
        rows = [choices[i:i+2] for i in range(0,len(choices),2)]
        rows += self.page_nav(page,pages,'panelpage')
        rows += [[btn('➕ Добавить панель','addpanel'),btn('‹ Админка','admin')]]
        self.card(uid,f'📚 Панели · {page+1}/{pages}',
                  'Нажми название для настройки. В разговоре выбирается только модель.' if panels else
                  'Добавь первую панель: название, API-адрес и ключ.',keyboard(*rows))

    def panel_settings(self, uid, pid):
        if uid not in self.admins:
            raise ValueError('Управление панелями доступно только администратору.')
        p = self.store.panel(uid, pid)
        if not p:
            raise ValueError('Панель не найдена.')
        if p['owner'] != uid:
            self.card(uid, '🤝 ' + p['name'], 'Общая панель администратора. Чаты личные. API-ключ скрыт; запросы расходуют баланс владельца.',
                      keyboard([btn('✓ Выбрать', f'panel:{pid}'), btn('‹ Панели', 'panels')]))
            return
        self.card(uid, '⚙️ ' + p['name'], p['url'] + '\nПотоковый ответ: ' + ('включён' if p['stream'] else 'выключен'), keyboard(
            [btn('✏️ Название', f'renamepanel:{pid}'), btn('🔑 API-ключ', f'rekey:{pid}')],
            [btn('🔎 Проверить API', f'check:{pid}'), btn('↔️ Режим ответа', f'stream:{pid}')],
            [btn('🗑 Удалить', f'delpanel:{pid}', 'danger'), btn('‹ Панели', 'panels')]))

    def provider(self, panel, uid=None):
        if uid is not None:
            panel = self.store.panel(uid, panel['id'])
            if not panel:
                raise ValueError('Доступ к панели отозван.')
        return Provider(panel, self.store.cipher)


    def models(self, uid, page=0, query='', favorites=False, refresh=False):
        panels = self.store.available_panels(uid)
        signature = [(p['id'], p['url'], p['token']) for p in panels]
        cached = self.cache.get(uid)
        if refresh or not cached or cached.get('signature') != signature or cached.get('expires', 0) < now():
            def fetch(p):
                try:
                    old = self.catalog_cache.get(p['id'], {})
                    key = (p['url'], p['token'])
                    if not refresh and old.get('key') == key and old.get('expires',0) > now():
                        models = old['models']
                    else:
                        models = self.provider(p, uid).models()
                        self.catalog_cache[p['id']] = {'key':key,'expires':now()+600,'models':models}
                    return [(p['id'], m, p['name']) for m in models], None
                except Exception:
                    return [], p['name']
            entries, failed = [], []
            if panels:
                with ThreadPoolExecutor(max_workers=min(4,len(panels))) as pool:
                    for found, error in pool.map(fetch, panels):
                        entries.extend(found)
                        if error:
                            failed.append(error)
            cached = {'signature':signature,'all':entries,'failed':failed,'expires':now()+600}
            self.cache[uid] = cached
        favs = {(r['panel'],r['model']) for r in self.store.rows('SELECT panel,model FROM favorites WHERE owner=?',(uid,))}
        visible = [e for e in cached['all'] if query.casefold() in (e[1]+' '+e[2]).casefold() and
                   (not favorites or (e[0],e[1]) in favs)]
        visible = sorted(set(visible),key=lambda e:((e[0],e[1]) not in favs,e[1].casefold(),e[0]))
        cached.update(visible=visible,query=query,favorites=favorites,revision=secrets.token_hex(3))
        pages = max(1,math.ceil(len(visible)/6))
        page = max(0,min(page,pages-1))
        aliases = self.panel_labels(panels)
        display = {}
        for pid,model,source in visible:
            label = self.compact_label(self.model_label(model)+' · '+aliases.get(pid,source))
            display.setdefault(label,[]).append((pid,model,source))
        choices = []
        active = self.store.panel(uid)
        current = self.store.chat(uid) if active else None
        for pid,model,source in visible[page*6:page*6+6]:
            label = self.compact_label(self.model_label(model)+' · '+aliases.get(pid,source))
            if len(display[label])>1:
                label = model+' · '+aliases.get(pid,source)
            selected = current and current['panel']==pid and current['model']==model
            choices.append(btn(('✅ ' if selected else '🧠 ')+label,f'choose:{pid}:{model}',
                               'success' if selected else 'primary'))
        rows = [choices[i:i+2] for i in range(0,len(choices),2)]
        rows += self.page_nav(page,pages,'modelpage')
        rows += [[btn('🔎 Найти модель','searchmodel'),btn('⭐ Избранное','models:favorites')],
                 [btn('📋 Все модели','models'),btn('🔄 Обновить','refreshmodels')]]
        if current and current['model']:
            marked = (current['panel'],current['model']) in favs
            rows.append([btn('★ Убрать текущую' if marked else '☆ Сохранить текущую','favorite_current'),
                         btn('💬 Продолжить','resume')])
        body = 'Выбери модель. Источник указан после точки, разговор сохранится.'
        if current:
            body = 'Сейчас: '+self.model_label(current['model'])+'\n\n'+body
        if not visible:
            body = 'Список пуст. Попробуй другой поиск или обнови модели.'
        if not panels:
            body = 'Нет разрешённых панелей. '+('Добавь панель в админке.' if uid in self.admins else 'Попроси администратора выдать доступ.')
        if cached['failed']:
            body += '\nВременно недоступны: '+', '.join(cached['failed'])[:300]
        self.card(uid,f'🧠 Модели · {page+1}/{pages}',body,keyboard(*rows))

    def chats(self, uid, page=0):
        chats = self.store.rows('SELECT * FROM chats WHERE owner=? ORDER BY updated DESC,id DESC',(uid,))
        chats = [c for c in chats if self.store.panel(uid,c['panel'])]
        pages = max(1,math.ceil(len(chats)/6))
        page = max(0,min(page,pages-1))
        titles = {}
        for c in chats:
            base = readable_name(c['title']) or 'Новый чат'
            titles.setdefault(base,[]).append(c)
        choices = []
        for c in chats[page*6:page*6+6]:
            title = readable_name(c['title']) or 'Новый чат'
            if len(titles[title])>1:
                title += ' · '+datetime.fromtimestamp(c['updated'],timezone.utc).strftime('%d.%m.%Y %H:%M:%S')
                same = sorted(titles[readable_name(c['title']) or 'Новый чат'],key=lambda x:x['id'])
                stamp = datetime.fromtimestamp(c['updated'],timezone.utc).strftime('%d.%m.%Y %H:%M:%S')
                same_second = [x for x in same if datetime.fromtimestamp(x['updated'],timezone.utc).strftime('%d.%m.%Y %H:%M:%S')==stamp]
                if len(same_second)>1:
                    title += ' · '+chr(ord('А')+next(i for i,x in enumerate(same_second) if x['id']==c['id']))
            choices.append(btn('💬 '+self.compact_label(title),f'chatcfg:{c["id"]}'))
        rows = [choices[i:i+2] for i in range(0,len(choices),2)]
        rows += self.page_nav(page,pages,'chatpage')
        rows += [[btn('✚ Новый чат','newchat'),btn('💬 Продолжить','resume')]]
        self.card(uid,f'💬 Мои чаты · {page+1}/{pages}',
                  'Нажми чат: открыть, переименовать, скачать историю или удалить. История видна только тебе.' if chats else
                  'Пока нет разговоров. Выбери модель и напиши первое сообщение.',keyboard(*rows))

    def owned_chat(self, uid, cid):
        chat = self.store.one('SELECT * FROM chats WHERE id=? AND owner=?', (cid, uid))
        if not chat:
            raise ValueError('Чат не найден.')
        if not self.store.panel(uid, chat['panel']):
            raise ValueError('Доступ к панели отозван.')
        return chat

    def settings(self, uid):
        labels = {'balanced':'обычный','short':'краткий','detailed':'подробный'}
        self.card(uid,'⚙️ Личные настройки','Стиль: '+labels[self.store.answer_style(uid)]+
                  '\nВход по Telegram ID, без пароля. Доступ к панелям выдаёт администратор.',
                  keyboard([btn('✍️ Стиль','styles'),btn('💰 Лимиты','limits')]))

    def styles(self, uid):
        selected = self.store.answer_style(uid)
        def choice(label, value):
            return btn(('✓ ' if selected == value else '') + label, 'style:' + value,
                       'success' if selected == value else 'primary')
        self.card(uid, '✍️ Стиль ответа',
                  'Обычный — длина по задаче.\nКраткий — сразу суть.\nПодробный — объяснение и примеры.\n\nВыбор личный и применяется ко всем твоим чатам.',
                  keyboard([choice('Обычный', 'balanced'), choice('Краткий', 'short')],
                           [choice('Подробный', 'detailed'), btn('‹ Настройки', 'settings')]))

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
        self.require_subscription(uid)
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
            self.require_subscription(uid)
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


    def route_input(self, uid, job, text):
        if not text.strip() or len(text) > 20000:
            raise ValueError('Отправь сообщение до 20 000 символов.')
        self.prepare_panel(uid)
        output = requested_output(text)
        if output == 'image':
            self.generate_image(uid, job, text)
        elif output:
            self.generation(uid, job, text, fmt=output)
        else:
            self.generation(uid, job, text)

    def image_prompt(self, uid):
        p = self.prepare_panel(uid)
        if not p:
            self.panels(uid)
            return
        chat = self.store.chat(uid)
        model = chat['model'] or self.store.preferred_model(uid, p)
        self.states[uid] = {'step': 'image_prompt'}
        self.card(uid, '🖼 Нарисовать изображение',
                  f"Панель: {p['name']}\nМодель: {self.model_label(model) if model else 'не выбрана'}\n\n"
                  'Напиши, что нарисовать. Можно просто отправить «Нарисуй…» в обычный чат.\n\n'
                  'Для рисования бот найдёт модель изображений в этой панели. Можно выбрать её вручную через «Модели». Поддержка рисования зависит от панели.',
                  keyboard([btn('🧠 Сменить модель', 'models'), btn('↩ Отмена', 'home')]))

    def send_pictures(self, uid, job, references, as_file=False):
        unique = list(dict.fromkeys(references))
        if len(unique) > 4:
            raise ValueError('За один ответ можно получить не более четырёх изображений.')
        for reference in unique:
            raw = picture_bytes(reference, job.cancel)
            photo = photo_buffer(raw)  # Validate format and dimensions even when sending a file.
            if job.cancel.is_set():
                raise Cancelled()
            if as_file:
                with Image.open(io.BytesIO(raw)) as source:
                    ext = {'JPEG': 'jpg', 'PNG': 'png', 'WEBP': 'webp', 'GIF': 'gif'}.get(source.format)
                if not ext:
                    raise ValueError('Для отправки файлом нужны PNG, JPEG, WEBP или GIF.')
                original = io.BytesIO(raw)
                original.name = 'AI-Chat.' + ext
                self.bot.send_document(uid, original,reply_markup=self.bottom(uid))
            else:
                self.bot.send_photo(uid, photo,reply_markup=self.bottom(uid))

    def generate_image(self, uid, job, prompt):
        if not prompt.strip() or len(prompt) > 20000:
            raise ValueError('Описание картинки должно содержать от 1 до 20 000 символов.')
        self.prepare_panel(uid)
        chat = self.store.chat(uid)
        panel = self.store.panel(uid, chat['panel'])
        if not panel:
            raise ValueError('Доступ к панели отозван.')
        model = chat['model'] or self.store.preferred_model(uid, panel)
        self.require_subscription(uid)
        configured = self.cfg.get('image_model', '')
        is_image = lambda name: bool(re.search(r'(?i)(gpt-image|dall-e|flux|imagen|ideogram|seedream|recraft|sdxl|stable.diffusion)', name))
        if configured:
            model = configured
        elif not model or not is_image(model):
            available = self.provider(panel, uid).models()
            candidates = sorted(name for name in available if is_image(name))
            if not candidates:
                raise ValueError('В панели не найдена модель для картинок. Выбери модель изображений в «Модели» или попроси владельца подключить её.')
            model = candidates[0]
        rates = self.cfg.get('image_prices', {})
        cost = rates.get(panel['url'] + '|' + model, rates.get(model)) if isinstance(rates, dict) else None
        if cost is not None:
            cost = float(cost)
            if not math.isfinite(cost) or cost < 0:
                raise ValueError('Некорректная цена изображения в image_prices.')
        estimate = len(prompt.encode()) + int(self.cfg['max_output_tokens'])
        rid = self.reserve(uid, model, estimate, cost)
        started, response, status, usage = False, None, None, None
        try:
            status = self.bot.send_message(uid, '✦ Рисую…', reply_markup=keyboard([btn('⏹ Стоп', 'stop:'+job.token, 'danger')]))
            if job.cancel.is_set():
                raise Cancelled()
            payload = {'model': model, 'prompt': prompt, 'n': 1}
            started = True
            response = self.provider(panel, uid).request('/images/generations', payload, read_timeout=180)
            job.response = response
            result = json.loads(bounded_response(response, 24 * 1024 * 1024, job.cancel))
            usage = result.get('usage')
            _, references = output_parts({'images': result.get('data', [])})
            if not references:
                raise ValueError('Панель не вернула картинку. Выбери модель изображений и проверь поддержку /images/generations.')
            if job.cancel.is_set():
                raise Cancelled()
            as_file = bool(re.search(r'(?i)\b(?:файлом|оригинал|as a file)\b', prompt))
            self.send_pictures(uid, job, references, as_file=as_file)
            self.store.append(chat['id'], prompt, '[Создано изображение]')
            if chat['title'] == 'Новый чат':
                self.store.write('UPDATE chats SET title=? WHERE id=? AND owner=?', (prompt[:50], chat['id'], uid))
            with self.menu_lock:
                self.menu_ids.pop(uid, None)
        except ProviderQuotaError:
            started = False
            raise
        finally:
            if response is not None:
                response.close()
                response._ai_session.close()
            if started:
                incoming = int(usage.get('input_tokens', 0)) if isinstance(usage, dict) else estimate
                outgoing = int(usage.get('output_tokens', 0)) if isinstance(usage, dict) else 0
                self.charge(uid, rid, 'image', model, incoming, outgoing, cost, not bool(usage))
            else:
                self.store.write('DELETE FROM reservations WHERE id=?', (rid,))
            if status is not None:
                self.erase(status)


    def send_answer(self, uid, text, cid):
        with self.menu_lock:
            self.menu_ids.pop(uid, None)
        for pos in range(0, len(text), 1800):
            chunk = html.escape(text[pos:pos + 1800])
            chunk = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', chunk)
            self.bot.send_message(uid, chunk, parse_mode='HTML', reply_markup=self.bottom(uid))

    def generation(self, uid, job, content, cid=None, fmt=None, retry=False):
        if job.cancel.is_set():
            raise Cancelled()
        self.require_subscription(uid)
        chat = self.owned_chat(uid,cid) if cid else self.store.chat(uid)
        original = (chat['panel'],chat['model'] or self.store.preferred_model(uid,self.store.panel(uid,chat['panel'])))
        entries = []
        for p in self.store.available_panels(uid):
            try:
                cached = self.catalog_cache.get(p['id'],{})
                key = (p['url'],p['token'])
                if cached.get('key')==key and cached.get('expires',0)>now():
                    models = cached['models']
                else:
                    models = self.provider(p,uid).models()
                    self.catalog_cache[p['id']] = {'key':key,'expires':now()+600,'models':models}
                entries.extend((p['id'],m) for m in models)
            except Exception:
                continue
        priorities = self.cfg.get('failover_priority',[])
        reserves = [(pid,m) for pid,m in entries if m==original[1] and pid!=original[0]]
        for model in priorities:
            reserves.extend((pid,m) for pid,m in entries if m==model and (pid,m)!=original)
        choices = list(dict.fromkeys([original]+reserves))
        failed_panels = set()
        for pid, model in choices:
            if job.cancel.is_set():
                raise Cancelled()
            if pid in failed_panels or not self.store.panel(uid,pid):
                continue
            self.require_subscription(uid)
            self.store.write('UPDATE chats SET panel=?,model=? WHERE id=? AND owner=?',(pid,model,chat['id'],uid))
            self.store.write('UPDATE users SET active=?,current_chat=? WHERE id=?',(pid,chat['id'],uid))
            try:
                return self._generation_once(uid,job,content,cid=chat['id'],fmt=fmt,retry=retry)
            except ProviderQuotaError:
                failed_panels.add(pid)
        raise ValueError('У доступных панелей исчерпан баланс. Администратор может пополнить их или настроить резервные модели.')

    def _generation_once(self, uid, job, content, cid=None, fmt=None, retry=False):
        chat = self.owned_chat(uid, cid) if cid else self.store.chat(uid)
        panel = self.store.panel(uid, chat['panel'])
        if not panel:
            raise ValueError('Доступ к панели отозван.')
        model = chat['model'] or self.store.preferred_model(uid, panel)
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
        style = self.store.answer_style(uid)
        if style == 'short':
            instruction += ' Отвечай кратко и по делу. Сохраняй существенные условия и важные детали.'
        elif style == 'detailed':
            instruction += ' Объясняй подробно, структурированно, с примерами, когда они помогают.'
        if fmt:
            instruction += f' Пользователь запросил настоящий файл {fmt}. Приложение само создаст и прикрепит его в Telegram из твоего ответа. Подготовь только содержимое документа. Не пиши, что не можешь прикрепить файл, не обещай ссылку и не давай инструкций самостоятельного сохранения.'
            if fmt == 'xlsx':
                instruction += ' Верни только JSON-массив строк таблицы: [["Заголовок"],["Значение"]], без ограждений.'
            elif fmt == 'pptx':
                instruction += ' Верни только JSON {"slides":[{"title":"Заголовок","points":["Пункт"]}]}. До 20 слайдов, до 6 пунктов по 180 символов на слайд, заголовок до 100 символов.'
            elif fmt == 'zip':
                instruction += ' Верни только JSON {"files":[{"name":"notes.txt","content":"Текст"}]}. До 20 файлов, простые имена без путей, только текстовое содержимое.'
            elif fmt == 'json':
                instruction += ' Верни только корректный JSON без ограждений и объяснений.'
            elif fmt == 'csv':
                instruction += ' Верни только CSV с заголовком.'
            else:
                instruction += ' Используй простые заголовки и абзацы. Не используй широкие Markdown-таблицы.'
        else:
            instruction += (
                ' Приложение умеет создавать настоящие PDF, DOCX, XLSX, TXT, MD, CSV, PPTX, JSON, HTML и ZIP.'
                ' Если пользователь просит файл, не утверждай, что не можешь прикреплять файлы: выдай содержимое в блоке '
                '```ai-files\n{"files":[{"name":"document.pdf","format":"pdf","content":"Текст документа"}]}\n```.'
                ' До четырёх файлов, только перечисленные форматы, content — строка или JSON.'
                ' Для XLSX content — массив строк, для PPTX объект slides с title и points (до 20 слайдов, 6 коротких пунктов),'
                ' для ZIP объект files с name и текстовым content (до 20 простых имён без путей).'
                ' Пиши обычный ответ вне блока; не выдавай этот блок, если пользователь не просил файл.'
                ' Если пользователь хочет картинку, можешь передать один блок '
                '```ai-image\n{"prompt":"Описание картинки"}\n```.'
                ' Приложение создаст её и пришлёт сюда же. Не используй этот блок без просьбы пользователя.'
                ' Не выдумывай ссылки на несуществующие вложения.')
        messages = [{'role': 'system', 'content': instruction}] + history + [{'role': 'user', 'content': content}]
        # UTF-8 bytes are a conservative text-token reservation, not an exact tokenizer.
        incoming_est = sum(len(json.dumps(x, ensure_ascii=False).encode()) if isinstance(x['content'], str)
                           else sum(len(y.get('text', '').encode()) for y in x['content']) + 20000 for x in messages)
        outgoing_max = int(self.cfg['max_output_tokens'])
        estimated_cost = self.price(panel, model, incoming_est, outgoing_max)
        rid = self.reserve(uid, model, incoming_est + outgoing_max, estimated_cost)
        answer = ''
        pictures = []
        usage = None
        response = None
        request_started = False
        finish = None
        status = None
        try:
            status = self.bot.send_message(uid, '✦ Думаю…', reply_markup=keyboard([btn('⏹ Стоп', 'stop:'+job.token, 'danger')]))
            if job.cancel.is_set():
                raise Cancelled()
            payload = {'model': model, 'messages': messages, 'max_tokens': outgoing_max, 'stream': bool(panel['stream'])}
            if panel['stream']:
                payload['stream_options'] = {'include_usage': True}
            request_started = True
            response = self.provider(panel, uid).request('/chat/completions', payload, stream=bool(panel['stream']))
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
                        chunk_text, chunk_pictures = output_parts(choice.get('delta', {}))
                        answer += chunk_text
                        pictures.extend(chunk_pictures)
                        if len(pictures) > 4 or sum(map(len, pictures)) > 24 * 1024 * 1024:
                            raise ValueError('Ответ содержит слишком много изображений.')
                        finish = choice.get('finish_reason') or finish
                    if len(answer) > 24 * 1024 * 1024:
                        raise ValueError('Ответ слишком большой; уменьши лимит генерации.')
                    if now() - last_edit > 2 and answer and len(answer) < 120000 and 'data:image' not in answer and 'ai-files' not in answer:
                        try:
                            self.bot.edit_message_text('✦ ' + answer[-2800:], uid, status.message_id,
                                                       reply_markup=keyboard([btn('⏹ Стоп','stop:'+job.token,'danger')]))
                        except Exception:
                            pass
                        last_edit = now()
            else:
                result = json.loads(bounded_response(response, 24 * 1024 * 1024, job.cancel))
                choice = result['choices'][0]
                answer, pictures = output_parts(choice['message'])
                finish = choice.get('finish_reason')
                usage = result.get('usage')
            if job.cancel.is_set():
                raise Cancelled()
            # Some panels return images as Markdown rather than a separate field.
            markdown_images = re.findall(r'!\[[^\]]*\]\((https://[^\s)]+|data:image/[^)\s]+)\)', answer)
            if markdown_images:
                pictures.extend(markdown_images)
                answer = re.sub(r'!\[[^\]]*\]\((https://[^\s)]+|data:image/[^)\s]+)\)', '', answer).strip()
            image_jobs = []
            def image_task(match):
                value = json.loads(match.group(1))
                prompt = value.get('prompt') if isinstance(value, dict) else None
                if not isinstance(prompt, str) or not prompt.strip() or len(prompt) > 4000 or image_jobs:
                    raise ValueError('Некорректный запрос изображения от модели.')
                image_jobs.append(prompt)
                return ''
            answer = re.sub(r'(?s)\x60{3}ai-image\s*\n(.*?)\n\x60{3}', image_task, answer)
            answer, artifacts = response_artifacts(answer)
            prepared_files = []
            for name, kind, data in artifacts:
                file = create_document(data, kind)
                file.name = name
                prepared_files.append(file)
            if not answer.strip() and not pictures and not prepared_files and not image_jobs:
                raise ValueError('Модель не вернула текст или изображение. Проверь модель и режим потока панели.')
            if len(answer) > 120000:
                raise ValueError('Ответ слишком большой; уменьши лимит генерации.')
            stored_answer = answer + ('\n[Изображение от ИИ]' if pictures else '')
            if artifacts:
                stored_answer += '\n[Файлы: ' + ', '.join(item[0] for item in artifacts) + ']\n'
                stored_answer += '\n'.join(item[2] for item in artifacts)
            self.store.append(chat['id'], content, stored_answer)
            if chat['title'] == 'Новый чат':
                title = content if isinstance(content, str) else content[0].get('text', 'Изображение')
                self.store.write('UPDATE chats SET title=? WHERE id=? AND owner=?', (title[:50], chat['id'], uid))
            for file in prepared_files:
                if job.cancel.is_set():
                    raise Cancelled()
                self.bot.send_document(uid, file,reply_markup=self.bottom(uid))
            if pictures:
                self.send_pictures(uid, job, pictures)
            if fmt and answer.strip() and not prepared_files:
                self.bot.send_document(uid, create_document(answer, fmt),reply_markup=self.bottom(uid))
            elif answer.strip():
                self.send_answer(uid, answer, chat['id'])
            for prompt in image_jobs:
                try:
                    self.generate_image(uid, job, prompt)
                except ProviderQuotaError:
                    raise ValueError('Баланс API изображений исчерпан. Текстовый ответ сохранён.') from None
            if finish == 'length':
                self.tell(uid, '⚠️ Достигнут лимит ответа модели. Ответ или файл может быть неполным.')
        except ProviderQuotaError:
            request_started = False
            raise
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
                # Unknownкнопка «🏠 Меню»led calls keep their full reservation in the local ledger.
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
            buttons.append([btn(fmt.upper(), f'{prefix}:{fmt}' + (f':{cid}' if cid else '')) for fmt in FORMATS[i:i + 2]])
        buttons.append([btn('❔ Помощь', 'help'), btn('🏠 Меню', 'home')])
        self.card(uid, '📄 Создать документ', 'Скачать последний ответ этого чата.' if cid else 'PDF и DOCX — документы\nXLSX и CSV — таблицы · PPTX — презентации\nTXT и MD — текст · JSON — данные\nHTML — статическая страница · ZIP — архив текстов\n\nВыбери формат, затем опиши содержимое.', keyboard(*buttons))

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
        panel = self.prepare_panel(uid)
        if not panel:
            raise ValueError('Сначала выбери доступную панель в меню «Панели».')
        obj = message.voice or message.audio
        raw = self.download(obj)
        seconds = getattr(obj, 'duration', 0) or 60
        rate = self.cfg['transcription_usd_per_minute']
        cost = seconds / 60 * float(rate) if rate is not None else None
        model = self.cfg['transcription_model']
        rid = self.reserve(uid, model, 0, cost)
        started = False
        response = None
        status = None
        try:
            if job.cancel.is_set():
                raise Cancelled()
            status = self.bot.send_message(uid, '✦ Распознаю голос…', reply_markup=keyboard([btn('⏹ Стоп', 'stop:'+job.token, 'danger')]))
            started = True
            name = getattr(obj, 'file_name', None) or 'voice.ogg'
            response = self.provider(panel, uid).request('/audio/transcriptions', {'model': model},
                                                   files={'file': (name, raw, getattr(obj, 'mime_type', None) or 'audio/ogg')})
            job.response = response
            result = json.loads(bounded_response(response, 1024 * 1024))
            text = result.get('text', '').strip()
            if not text:
                raise ValueError('Не удалось распознать речь.')
            if job.cancel.is_set():
                raise Cancelled()
        except ProviderQuotaError:
            started = False
            raise
        finally:
            if response is not None:
                response.close()
                response._ai_session.close()
            if started:
                self.charge(uid, rid, 'transcription', model, 0, 0, cost, True)
            else:
                self.store.write('DELETE FROM reservations WHERE id=?', (rid,))
            if status is not None:
                self.erase(status)
        # Transcription and the requested answer are separately billed API operations.
        self.route_input(uid, job, text)

    def routing_settings(self, uid):
        if uid not in self.admins:
            raise ValueError('Нужны права администратора.')
        order = self.cfg.get('failover_priority',[])
        self.card(uid,'🔁 Резервные модели',
                  'Сначала та же модель в другой разрешённой панели, затем модели по списку ниже.\n\n'+
                  ('\n'.join(f'{i+1}. {m}' for i,m in enumerate(order)) or 'Список пока пуст.')+
                  '\n\nПереключение — только при явной ошибке баланса/квоты, до начала ответа. Личные лимиты не обходятся.',
                  keyboard([btn('✏️ Порядок','routing_edit'),btn('‹ Админка','admin')]))

    def admin_menu(self, uid):
        if uid not in self.admins:
            raise ValueError('Этот раздел доступен администратору.')
        self.states.pop(uid, None)
        users = self.store.one('SELECT COUNT(*) AS n FROM users')['n']
        panels = self.store.one('SELECT COUNT(*) AS n FROM panels WHERE owner=?', (uid,))['n']
        shared = self.store.one('SELECT COUNT(*) AS n FROM panel_shares s JOIN panels p ON p.id=s.panel WHERE p.owner=?', (uid,))['n']
        self.card(uid, '🛠 Управление ботом',
                  f'Пользователей: {users}\nТвоих панелей: {panels} · Выданных доступов: {shared}\n\n1. Добавь друга по ID или пригласи ссылкой.\n2. Выдай срок доступа в его карточке.\n3. Открой «Панели друга» и разреши нужные.',
                  keyboard([btn('👥 Друзья', 'friends'), btn('🔗 Пригласить', 'invite')],
                           [btn('🔁 Резервные модели','routing'),btn('⚙️ Настройки','settings')],
                           [btn('✚ Друг по ID', 'addfriend'), btn('📚 Мои панели', 'panels')],
                           [btn('📤 Копия базы', 'backup'), btn('📥 Восстановить', 'restore')],
                           [btn('🗑 Очистить базу', 'cleardb', 'danger'), btn('🏠 Меню', 'home')]))

    def friends(self, uid, page=0):
        if uid not in self.admins:
            raise ValueError('Нужны права администратора.')
        rows = self.store.rows('SELECT id FROM users UNION SELECT id FROM admitted UNION SELECT id FROM subscriptions ORDER BY id')
        friends = [r['id'] for r in rows if r['id'] not in self.admins]
        total = max(1, math.ceil(len(friends) / 6))
        page = max(0, min(page, total-1))
        buttons = []
        for friend in friends[page*6:page*6+6]:
            active = self.store.has_subscription(friend)
            count = self.store.one('SELECT COUNT(*) AS n FROM panel_shares s JOIN panels p ON p.id=s.panel WHERE s.friend=? AND p.owner=?', (friend, uid))['n']
            buttons.append([btn(('🟢 ' if active else '⚪ ') + str(friend), f'friend:{friend}',
                                'success' if active else 'primary'),
                            btn(f'🤝 Панели {friend}', f'shares:{friend}')])
        buttons += self.page_nav(page, total, 'friendpage')
        buttons += [[btn('✚ Друг по ID', 'addfriend'), btn('🔗 Пригласить', 'invite')],
                    [btn('🔄 Обновить', f'friendpage:{page}'), btn('‹ Админка', 'admin')]]
        self.card(uid, f'👥 Друзья · {len(friends)} · {page+1}/{total}',
                  '🟢 Есть подписка · ⚪ нет доступа к ИИ.\nКарточка друга — срок доступа и отключение. «Панели» — выбор твоих API-панелей для него.' if friends else
                  'Добавь друга по Telegram ID или отправь ему приглашение. Затем выдай срок доступа и выбери панели.',
                  keyboard(*buttons))

    def friend_card(self, uid, friend):
        if uid not in self.admins:
            raise ValueError('Нужны права администратора.')
        if friend <= 0 or friend in self.admins:
            raise ValueError('Укажи ID друга. У администраторов доступ без срока.')
        count = self.store.one('SELECT COUNT(*) AS n FROM panel_shares s JOIN panels p ON p.id=s.panel WHERE p.owner=? AND s.friend=?', (uid, friend))['n']
        registered = bool(self.store.user(friend))
        self.card(uid, f'👤 Друг · {friend}', self.subscription_label(friend) +
                  f'\nРегистрация: {"готова" if registered else "нужно открыть /start"}\nТвоих выданных панелей: {count}\n\nСначала выбери срок, затем выдай панели. Продление прибавляет дни; бессрочный доступ не сокращается. История у друга личная.',
                  keyboard([btn('+7 дней', f'grant:{friend}:7'), btn('+30 дней', f'grant:{friend}:30')],
                           [btn('+90 дней', f'grant:{friend}:90'), btn('Без срока', f'grant:{friend}:0')],
                           [btn('🤝 Панели друга', f'shares:{friend}'), btn('👥 Все друзья', 'friends')],
                           [btn('⛔ Отключить', f'revokefriend:{friend}', 'danger'), btn('‹ Админка', 'admin')]))

    def share_menu(self, uid, friend, page=0):
        if uid not in self.admins or friend<=0 or friend in self.admins:
            raise ValueError('Нужны права администратора и ID друга.')
        panels = self.store.rows('SELECT * FROM panels WHERE owner=? ORDER BY id',(uid,))
        aliases = self.panel_labels(panels)
        pages = max(1,math.ceil(len(panels)/6))
        page = max(0,min(page,pages-1))
        choices = []
        for p in panels[page*6:page*6+6]:
            granted = bool(self.store.one('SELECT 1 FROM panel_shares WHERE panel=? AND friend=?',(p['id'],friend)))
            choices.append(btn(('✅ ' if granted else '➕ ')+self.compact_label(aliases[p['id']]),
                               f'sharetoggle:{friend}:{p["id"]}:{page}','success' if granted else 'primary'))
        rows = [choices[i:i+2] for i in range(0,len(choices),2)]
        rows += self.page_nav(page,pages,f'shares:{friend}')
        rows += [[btn('👤 Карточка друга',f'friend:{friend}'),btn('👥 Все друзья','friends')]]
        body = '✅ Разрешена · ➕ Выдать. Нажми панель, чтобы включить или отключить доступ. Ключ скрыт, расходы идут с твоего API-баланса.'
        if not self.store.has_subscription(friend):
            body += '\n\nСначала выдай срок доступа в карточке друга.'
        self.card(uid,f'🤝 Панели для {friend} · {page+1}/{pages}',body,keyboard(*rows))

    def invite(self, uid, days=None):
        if uid not in self.admins:
            raise ValueError('Нужны права администратора.')
        if days is None:
            self.card(uid, '🔗 Пригласить друга',
                      'Выбери срок доступа. Подписка начнётся, когда новый друг откроет ссылку. Ссылка одноразовая и действует 24 часа.',
                      keyboard([btn('7 дней', 'invite:7', 'success'), btn('30 дней', 'invite:30', 'success')],
                               [btn('90 дней', 'invite:90', 'success'), btn('Без срока', 'invite:0', 'success')],
                               [btn('👥 Друзья', 'friends'), btn('‹ Админка', 'admin')]))
            return
        if days not in (0, 7, 30, 90):
            raise ValueError('Неизвестный срок доступа.')
        code = secrets.token_urlsafe(18)
        self.store.write('INSERT INTO invites(code,uses,expires,days) VALUES(?,?,?,?)', (code, 1, now() + 86400, days))
        name = self.bot.get_me().username
        term = f'{days} дней' if days else 'без срока'
        self.tell(uid, f'🔗 Доступ: {term}. Перешли ссылку другу. Она действует один раз, в течение 24 часов.\n\nhttps://t.me/{name}?start=invite_{code}')

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
            expires = now() + invite['days'] * 86400 if invite['days'] else 0
            c.execute('INSERT OR IGNORE INTO subscriptions(id,expires,blocked) VALUES(?,?,0)', (uid, expires))

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
                command, argument = telegram_command(message.text)
                if command != '/start':
                    self.tell(uid, 'Используй кнопки снизу или /start.')
                    return
                if argument.startswith('invite_'):
                    self.admit(uid,argument[7:])
                self.states.pop(uid,None)
                if self.gate(uid):
                    self.home(uid)
            except ValueError as error:
                self.tell(uid,str(error))

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
                if uid not in self.admins and action in ('panels','panelpage','panel','panelcfg','addpanel','renamepanel','rekey','stream','check','delpanel','yespanel'):
                    raise ValueError('Панелями управляет только администратор.')
                if action == 'stop':
                    job = self.jobs.get(uid)
                    if len(parts)>1 and (not job or parts[1]!=job.token):
                        return
                    self.stop(uid)
                    return
                if action == 'noop':
                    return
                if uid in self.jobs and action not in ('home', 'usage'):
                    raise ValueError('Дождись ответа или нажми «Стоп», прежде чем менять настройки.')
                if action not in ('confirmrestore', 'confirmcleardb', 'confirmrevoke', 'yespanel', 'yeschat'):
                    self.states.pop(uid, None)
                if action == 'home':
                    self.home(uid)
                elif action in ('panels', 'panelpage'):
                    self.states.pop(uid, None)
                    self.panels(uid, int(parts[1]) if action == 'panelpage' else 0)
                elif action == 'resume':
                    self.resume(uid)
                elif action == 'help':
                    self.help_menu(uid)
                elif action == 'addpanel':
                    self.states[uid] = {'step': 'panel_name'}
                    self.tell(uid, 'Название панели? «🏠 Меню» — отмена.')
                elif action == 'panel':
                    pid = int(parts[1])
                    if not self.store.panel(uid, pid):
                        raise ValueError('Панель не найдена.')
                    self.store.write('UPDATE users SET active=?,current_chat=NULL WHERE id=?', (pid, uid))
                    self.cache.pop(uid, None)
                    self.models(uid)
                elif action == 'panelcfg':
                    self.panel_settings(uid, int(parts[1]))
                elif action in ('renamepanel', 'rekey', 'stream', 'check', 'delpanel', 'yespanel'):
                    pid = int(parts[1])
                    panel = self.store.panel(uid, pid)
                    if not panel:
                        raise ValueError('Панель не найдена.')
                    if panel['owner'] != uid:
                        raise ValueError('Изменять общую панель может только её владелец.')
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
                        self.states[uid] = {'step':'delete_panel','panel':pid,'expires':now()+300}
                        self.card(uid, '🗑 Удалить панель?', panel['name'] + '\nБудут удалены ключ, все чаты панели и избранное.', keyboard(
                            [btn('Да, удалить', f'yespanel:{pid}', 'danger'), btn('Отмена', f'panelcfg:{pid}')]))
                    else:
                        state = self.states.get(uid,{})
                        if state.get('step')!='delete_panel' or state.get('panel')!=pid or state.get('expires',0)<now():
                            raise ValueError('Подтверди удаление панели заново.')
                        self.states.pop(uid,None)
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
                    self.tell(uid, 'Напиши часть названия модели. «🏠 Меню» — отмена.')
                elif action == 'choose':
                    if len(parts)<3:
                        raise ValueError('Модель не найдена.')
                    self.select_model(uid,int(parts[1]),':'.join(parts[2:]))
                elif action == 'favorite_current':
                    chat = self.store.chat(uid)
                    self.owned_chat(uid,chat['id'])
                    if not chat['model']:
                        raise ValueError('Сначала выбери модель.')
                    existing = self.store.one('SELECT 1 FROM favorites WHERE owner=? AND panel=? AND model=?',(uid,chat['panel'],chat['model']))
                    if existing:
                        self.store.write('DELETE FROM favorites WHERE owner=? AND panel=? AND model=?',(uid,chat['panel'],chat['model']))
                    else:
                        self.store.write('INSERT INTO favorites VALUES(?,?,?)',(uid,chat['panel'],chat['model']))
                    self.models(uid)
                elif action in ('model', 'favorite'):
                    cache = self.cache.get(uid,{})
                    if len(parts)!=3 or parts[1]!=cache.get('revision') or cache.get('expires',0)<now():
                        raise ValueError('Список устарел. Открой «Модели» заново.')
                    index = int(parts[2])
                    if index < 0 or index >= len(cache.get('visible',[])):
                        raise ValueError('Модель не найдена.')
                    pid, model, source = cache['visible'][index]
                    p = self.store.panel(uid,pid)
                    if not p:
                        raise ValueError('Доступ к модели отозван.')
                    if action == 'favorite':
                        exists = self.store.one('SELECT 1 FROM favorites WHERE owner=? AND panel=? AND model=?',(uid,pid,model))
                        if exists:
                            self.store.write('DELETE FROM favorites WHERE owner=? AND panel=? AND model=?',(uid,pid,model))
                        else:
                            self.store.write('INSERT INTO favorites VALUES(?,?,?)',(uid,pid,model))
                        self.models(uid,query=cache.get('query',''),favorites=cache.get('favorites',False))
                    else:
                        self.select_model(uid,pid,model)
                elif action in ('chats', 'chatpage'):
                    self.chats(uid, int(parts[1]) if action == 'chatpage' else 0)
                elif action in ('chat', 'chatcfg', 'renamechat', 'delchat', 'yeschat', 'history'):
                    cid = int(parts[1])
                    chat = self.owned_chat(uid, cid)
                    if action == 'chat':
                        self.store.write('UPDATE users SET active=?,current_chat=? WHERE id=?', (chat['panel'], cid, uid))
                        self.resume(uid)
                    elif action == 'chatcfg':
                        self.card(uid, '💬 ' + chat['title'], chat['model'] or 'Модель не выбрана', keyboard(
                            [btn('💬 Открыть', f'chat:{cid}'),btn('📜 История', f'history:{cid}')],
                            [btn('✏️ Название',f'renamechat:{cid}'),btn('🧠 Модели','models')],
                            [btn('🗑 Удалить', f'delchat:{cid}', 'danger'), btn('‹ Чаты', 'chats')]))
                    elif action == 'renamechat':
                        self.states[uid] = {'step': 'renamechat', 'chat': cid}
                        self.tell(uid, 'Новое название чата?')
                    elif action == 'delchat':
                        self.states[uid] = {'step':'delete_chat','chat':cid,'expires':now()+300}
                        self.card(uid, 'Удалить чат?', chat['title'], keyboard([btn('Удалить', f'yeschat:{cid}', 'danger'), btn('Отмена', f'chatcfg:{cid}')]))
                    elif action == 'yeschat':
                        state = self.states.get(uid,{})
                        if state.get('step')!='delete_chat' or state.get('chat')!=cid or state.get('expires',0)<now():
                            raise ValueError('Подтверди удаление чата заново.')
                        self.states.pop(uid,None)
                        self.store.write('DELETE FROM chats WHERE owner=? AND id=?', (uid, cid))
                        self.store.write('UPDATE users SET current_chat=NULL WHERE id=? AND current_chat=?', (uid, cid))
                        self.chats(uid)
                    else:
                        history = self.store.history(cid, 100)
                        text = '\n\n'.join(m['role'].upper() + '\n' + (m['content'] if isinstance(m['content'], str) else '[Изображение]\n' + m['content'][0].get('text', '')) for m in history)
                        self.bot.send_document(uid, create_document(text or 'Пустой чат', 'txt'))
                elif action == 'newchat':
                    if not self.prepare_panel(uid):
                        self.panels(uid)
                    else:
                        self.store.new_chat(uid)
                        self.resume(uid)
                elif action in ('formats', 'saveformats'):
                    self.formats(uid, int(parts[1]) if action == 'saveformats' else None)
                elif action == 'format':
                    if parts[1] not in FORMATS:
                        return
                    self.states[uid] = {'step': 'file_prompt', 'format': parts[1]}
                    self.tell(uid, 'Что подготовить в ' + parts[1].upper() + '? Опиши содержимое. «🏠 Меню» — отмена.')
                elif action == 'save':
                    cid = int(parts[2])
                    self.bot.send_document(uid, create_document(self.last_answer(uid, cid), parts[1]))
                elif action == 'speak':
                    self.tell(uid, 'Озвучка отключена. Ответы приходят текстом, изображениями и документами.')
                elif action in ('rewrite', 'retry'):
                    cid = int(parts[-1])
                    self.owned_chat(uid, cid)
                    if action == 'retry':
                        self.launch(uid, lambda job: self.generation(uid, job, '', cid=cid, retry=True))
                    else:
                        prompts = {'short': 'Сократи предыдущий ответ, сохрани суть.', 'long': 'Раскрой предыдущий ответ подробнее с примерами.', 'simple': 'Объясни предыдущий ответ проще, понятными словами.'}
                        prompt = prompts[parts[1]]
                        self.launch(uid, lambda job: self.generation(uid, job, prompt, cid=cid))
                elif action == 'styles':
                    self.styles(uid)
                elif action == 'style':
                    if len(parts) != 2 or parts[1] not in ('balanced', 'short', 'detailed'):
                        raise ValueError('Неизвестный стиль ответа.')
                    self.store.write('INSERT INTO preferences(owner,style) VALUES(?,?) ON CONFLICT(owner) DO UPDATE SET style=excluded.style', (uid, parts[1]))
                    self.styles(uid)
                elif action == 'imageprompt':
                    self.image_prompt(uid)
                elif action == 'settings':
                    self.settings(uid)
                elif action == 'usage':
                    self.usage(uid)
                elif action == 'limits':
                    self.states[uid] = {'step': 'limits'}
                    self.tell(uid, 'Отправь дневной лимит токенов и USD через пробел, например: 100000 2.\n0 — без лимита. Для USD нужны тарифы в config.json. Сутки по UTC.')
                elif action in ('password','logout'):
                    self.states.pop(uid,None)
                    self.home(uid)
                elif action in ('routing','routing_edit'):
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    if action == 'routing':
                        self.routing_settings(uid)
                    else:
                        self.states[uid] = {'step':'routing_edit'}
                        self.tell(uid,'Введи API ID резервных моделей по одному в строке, в порядке приоритета. Например: gpt-4o-mini. До 30 моделей. Один дефис очищает список.')
                elif action == 'admin':
                    self.admin_menu(uid)
                elif action == 'subscription':
                    self.subscription_info(uid)
                elif action in ('friends', 'friendpage'):
                    self.states.pop(uid, None)
                    self.friends(uid, int(parts[1]) if action == 'friendpage' else 0)
                elif action == 'addfriend':
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    self.states[uid] = {'step': 'friend_id'}
                    self.tell(uid, 'Отправь числовой Telegram ID друга. Он увидит свой ID в «🎟 Мой доступ». «🏠 Меню» — отмена.')
                elif action == 'friend':
                    self.friend_card(uid, int(parts[1]))
                elif action in ('shares', 'sharetoggle'):
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    friend = int(parts[1])
                    if friend <= 0 or friend in self.admins:
                        raise ValueError('Укажи ID друга.')
                    page = int(parts[-1]) if len(parts) > 2 else 0
                    if action == 'sharetoggle':
                        pid = int(parts[2])
                        p = self.store.one('SELECT * FROM panels WHERE id=? AND owner=?', (pid, uid))
                        if not p:
                            raise ValueError('Выдавать можно только собственные панели.')
                        with self.store.db() as c:
                            exists = c.execute('SELECT 1 FROM panel_shares WHERE panel=? AND friend=?', (pid, friend)).fetchone()
                            if exists:
                                c.execute('DELETE FROM panel_shares WHERE panel=? AND friend=?', (pid, friend))
                                c.execute('UPDATE users SET active=NULL,current_chat=NULL WHERE id=? AND active=?', (friend, pid))
                            else:
                                c.execute('INSERT INTO panel_shares VALUES(?,?)', (pid, friend))
                                c.execute('UPDATE users SET active=?,current_chat=NULL WHERE id=? AND active IS NULL', (pid, friend))
                        if exists:
                            with self.lock:
                                job = self.jobs.get(friend)
                                if job:
                                    job.cancel.set()
                        self.cache.pop(friend, None)
                    self.share_menu(uid, friend, page)
                elif action == 'grant':
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    friend, duration = int(parts[1]), int(parts[2])
                    if friend in self.admins:
                        raise ValueError('Администратору подписка не требуется.')
                    self.store.grant_subscription(friend, duration)
                    self.friend_card(uid, friend)
                elif action == 'revokefriend':
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    friend = int(parts[1])
                    if friend <= 0 or friend in self.admins:
                        raise ValueError('Нельзя отключить доступ администратору.')
                    self.states[uid] = {'step': 'revoke_friend', 'friend': friend, 'expires': now() + 300}
                    self.card(uid, 'Отключить доступ?', f'Друг {friend} больше не сможет отправлять запросы к ИИ. Его панели и история сохранятся.',
                              keyboard([btn('⛔ Отключить', 'confirmrevoke', 'danger'), btn('Отмена', 'friends')]))
                elif action == 'confirmrevoke':
                    state = self.states.get(uid, {})
                    if uid not in self.admins or state.get('step') != 'revoke_friend' or state.get('expires', 0) < now():
                        raise ValueError('Подтверждение истекло.')
                    friend = state['friend']
                    with self.lock:
                        self.store.write('INSERT INTO subscriptions(id,expires,blocked) VALUES(?,0,1) '
                                         'ON CONFLICT(id) DO UPDATE SET blocked=1', (friend,))
                        if friend in self.jobs:
                            self.jobs[friend].cancel.set()
                    self.states.pop(uid, None)
                    self.friend_card(uid, friend)
                elif action == 'invite':
                    self.invite(uid, int(parts[1]) if len(parts) > 1 else None)
                elif action == 'backup':
                    self.export(uid)
                elif action == 'restore':
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    if self.jobs:
                        raise ValueError('Перед импортом дождись завершения всех запросов.')
                    self.states[uid] = {'step': 'import_file','expires':now()+300}
                    self.tell(uid, 'Пришли зашифрованную .enc-копию в течение 5 минут. Перед заменой потребуется подтверждение.')
                elif action == 'cleardb':
                    if uid not in self.admins:
                        raise ValueError('Нужны права администратора.')
                    if self.jobs:
                        raise ValueError('Перед очисткой дождись завершения запросов.')
                    self.states[uid] = {'step':'clear_confirm','expires':now()+300}
                    self.card(uid,'🗑 Очистить базу?', 'Будут удалены все пользователи, панели и чаты. Сначала сохранится зашифрованная копия. Подтверждение действует 5 минут.',keyboard([btn('Да, очистить','confirmcleardb','danger'),btn('Отмена','admin')]))
                elif action == 'confirmcleardb':
                    state = self.states.get(uid, {})
                    if uid not in self.admins or state.get('step') != 'clear_confirm' or state.get('expires', 0) < now():
                        if state.get('step')=='clear_confirm' and state.get('expires',0)<now():
                            self.states.pop(uid,None)
                        raise ValueError('Подтверждение очистки истекло. База не изменена.')
                    with self.lock:
                        if self.jobs:
                            raise ValueError('Сначала дождись завершения всех запросов.')
                        self.store.clear_database()
                        self.states.clear()
                        self.cache.clear()
                        self.reply_routes.clear()
                        self.reply_labels.clear()
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
                        self.reply_routes.clear()
                        self.reply_labels.clear()
                    self.tell(uid, 'База восстановлена. /start — открыть чат.')
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
        with self.menu_lock:
            self.menu_ids.pop(uid, None)
        text = message.text or ''
        command, _ = telegram_command(text)
        if command:
            if command == '/start':
                self.command(message)
            else:
                self.tell(uid,'Осталась только команда /start. Остальное — кнопками снизу.')
            return
        route = self.load_routes(uid).get(text) or {
            '🏠 Меню':'home','⏹ Стоп':'stop','🧠 Модели':'models','💬 Чаты':'chats',
            '✚ Новый чат':'newchat','⭐ Избранное':'models:favorites','⚙️ Настройки':'settings',
            '🎟 Доступ':'subscription','🛠 Админка':'admin','📚 Панели':'panels'}.get(text)
        if route:
            call = type('LocalCall',(),{'id':'','from_user':message.from_user,'message':message,'data':route})()
            self.callback(call)
            return
        if text in self.reply_labels.get(uid,set()) or re.match(r'^\d+\. 🧠 ',text) or re.match(r'^(?:🧠|✅) .+ · .+',text):
            if self.gate(uid):
                self.home(uid)
            return
        state = self.states.get(uid, {})
        step = state.get('step')
        u = self.store.user(uid)
        if not self.gate(uid):
            return
        if uid in self.jobs:
            raise ValueError('Сейчас выполняется запрос. кнопка «⏹ Стоп» — остановить.')
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
                raise ValueError('Закончи текущую настройку или кнопка «🏠 Меню», затем пришли вложение ещё раз.')
            self.launch(uid, lambda job: self.attachment(uid, job, message))
            return
        nav = {
            '💬 Продолжить': 'resume',
            '🗂 Мои чаты': 'chats',
            '✚ Новый чат': 'newchat',
            '📄 Документы': 'formats',
            '🖼 Нарисовать': 'imageprompt',
            '❔ Помощь': 'help',
            '🏠 Меню': 'home',
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
            '🎟 Мой доступ': 'subscription',
            '👥 Друзья и подписки': 'friends',
        }
        if text in nav:
            self.states.pop(uid, None)
            # Use the same authorization-aware routing as inline buttons.
            call = type('LocalCall', (), {'id': '', 'from_user': message.from_user, 'message': message, 'data': nav[text]})()
            self.callback(call)
            return
        if step == 'routing_edit':
            if uid not in self.admins:
                raise ValueError('Нужны права администратора.')
            values = [] if text.strip()=='-' else [line.strip() for line in text.splitlines() if line.strip()]
            if len(values)>30 or any(len(m)>200 or any(ord(c)<32 for c in m) for m in values):
                raise ValueError('До 30 моделей, по одному API ID в строке.')
            path = Path(os.environ.get('AI_CHAT_CONFIG',str(ROOT/'config.json')))
            data = json.loads(path.read_text(encoding='utf-8'))
            data['failover_priority'] = list(dict.fromkeys(values))
            temp = path.with_suffix('.json.tmp')
            temp.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
            os.chmod(temp,0o600)
            os.replace(temp,path)
            self.cfg['failover_priority'] = data['failover_priority']
            self.states.pop(uid,None)
            self.routing_settings(uid)
        elif step == 'friend_id':
            if uid not in self.admins:
                raise ValueError('Нужны права администратора.')
            if not text.strip().isdigit() or int(text.strip()) <= 0 or len(text.strip()) > 18:
                raise ValueError('Отправь положительный числовой Telegram ID.')
            friend = int(text.strip())
            self.states.pop(uid, None)
            self.friend_card(uid, friend)
        elif step in ('panel_name','panel_url','panel_token','renamepanel','rekey') and uid not in self.admins:
            self.states.pop(uid,None)
            raise ValueError('Панелями управляет только администратор.')
        elif step == 'panel_name':
            name = readable_name(text)
            if not name or len(name) > 60:
                raise ValueError('Название: 1–60 символов.')
            self.states[uid] = {'step': 'panel_url', 'name': name}
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
            records = self.provider(panel).model_records()
            available = [m['id'] for m in records]
            selected = universal_model(records)
            if not selected:
                raise ValueError('В панели не найдены разговорные модели. Добавление отменено.')
            pid = self.store.write('INSERT INTO panels(owner,name,url,token,model) VALUES(?,?,?,?,?)', (uid, state['name'], state['url'], encrypted,selected))
            self.store.write('INSERT INTO panel_choices VALUES(?,?,?)',(uid,pid,selected))
            self.store.write('UPDATE users SET active=?,current_chat=NULL WHERE id=?', (pid, uid))
            self.states.pop(uid, None)
            self.cache[uid] = {'panel': pid, 'all': available, 'expires': now() + 1800}
            self.cache.pop(uid,None)
            self.resume(uid)
        elif step in ('renamepanel', 'rekey'):
            panel = self.store.panel(uid, state['panel'])
            if not panel or panel['owner'] != uid:
                raise ValueError('Панель не найдена или нет прав на изменение.')
            if step == 'renamepanel':
                name = readable_name(text)
                if not name or len(name) > 60:
                    raise ValueError('Название: 1–60 символов.')
                self.store.write('UPDATE panels SET name=? WHERE id=? AND owner=?', (name, panel['id'], uid))
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
            name = readable_name(text)
            if not name or len(name) > 80:
                raise ValueError('Название чата: 1–80 символов.')
            self.store.write('UPDATE chats SET title=? WHERE id=? AND owner=?', (name, chat['id'], uid))
            self.states.pop(uid, None)
            self.chats(uid)
        elif step == 'limits':
            try:
                tokens, usd = text.replace(',', '.').split()
                tokens, usd = int(tokens), float(usd)
            except (ValueError,TypeError):
                raise ValueError('Введи два числа: лимит токенов и USD. Например: 100000 2. Ноль отключает соответствующий лимит.') from None
            if tokens > 9223372036854775807:
                raise ValueError('Лимит токенов слишком большой. Для отключения ограничения укажи 0.')
            if tokens < 0 or usd < 0 or not math.isfinite(usd):
                raise ValueError('Лимиты должны быть конечными неотрицательными числами.')
            self.store.write('UPDATE users SET daily_tokens=?,daily_usd=? WHERE id=?', (tokens, usd, uid))
            self.states.pop(uid, None)
            self.usage(uid)
        elif step == 'image_prompt':
            if len(text) > 20000:
                raise ValueError('Сократи описание до 20 000 символов.')
            self.launch(uid, lambda job: self.generate_image(uid, job, text))
            self.states.pop(uid, None)
        elif step == 'file_prompt':
            if len(text) > 20000:
                raise ValueError('Сократи запрос до 20 000 символов.')
            fmt = state['format']
            self.launch(uid, lambda job: self.generation(uid, job, text, fmt=fmt))
            self.states.pop(uid, None)
        elif step:
            self.tell(uid, 'Ожидаю текущий шаг настройки. «🏠 Меню» — отмена.')
        else:
            if len(text) > 20000:
                raise ValueError('Сократи запрос до 20 000 символов.')
            self.launch(uid, lambda job: self.route_input(uid, job, text))

    def attachment(self, uid, job, message):
        self.prepare_panel(uid)
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
        output = requested_output(caption)
        self.generation(uid, job, content, fmt=output if output in FORMATS else None)

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
            self.bot.set_my_commands([types.BotCommand('start', 'Открыть чат')])
            try:
                self.bot.set_chat_menu_button(menu_button=types.MenuButtonCommands())
            except Exception:
                pass
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
