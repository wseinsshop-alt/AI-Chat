#!/usr/bin/env bash
# AI Chat: install/update on Ubuntu 22.04+ or Debian 12+.
set -Eeuo pipefail
umask 077
if [[ ${EUID} -ne 0 ]]; then echo 'Run with sudo bash install.sh (or as root).' >&2; exit 1; fi
[[ -d /run/systemd/system ]] || { echo 'A server with systemd is required.' >&2; exit 1; }
exec 9>/var/lock/ai-chat-install.lock
flock -n 9 || { echo 'Another AI Chat installation is running.' >&2; exit 1; }
configure=0
revision=${AI_CHAT_REVISION:-main}
while [[ $# -gt 0 ]]; do
  case "$1" in
    --configure) configure=1; shift ;;
    --revision) revision=${2:?Missing revision}; shift 2 ;;
    *) echo "Unknown option: $1" >&2; exit 1 ;;
  esac
done
[[ $revision =~ ^[A-Za-z0-9._/-]+$ ]] || { echo 'Invalid revision' >&2; exit 1; }
export DEBIAN_FRONTEND=noninteractive NEEDRESTART_MODE=a
app=/opt/ai-chat
data=/var/lib/ai-chat
unit=/etc/systemd/system/ai-chat.service
stage=$(mktemp -d /tmp/ai-chat-install.XXXXXX)
mutated=0
was_active=0
snapshot=
rollback() {
  code=$?
  trap - ERR
  if (( mutated )); then
    systemctl stop ai-chat.service 2>/dev/null || true
    for file in main.py README.md config.json requirements.txt install.sh; do
      if [[ -f "$stage/previous/$file" ]]; then cp -p "$stage/previous/$file" "$app/$file"; else rm -f "$app/$file"; fi
    done
    if [[ -f "$stage/previous/service" ]]; then cp -p "$stage/previous/service" "$unit"; else rm -f "$unit"; fi
    # Restore the consistent pre-upgrade SQLite snapshot before starting old code.
    if [[ -n "$snapshot" && -f "$snapshot/bot.db" ]]; then
      rm -f "$data/bot.db-wal" "$data/bot.db-shm"
      cp -p "$snapshot/bot.db" "$data/bot.db"
    fi
    for file in secret.key settings.json; do
      [[ ! -f "$snapshot/$file" ]] || cp -p "$snapshot/$file" "$data/$file"
    done
    systemctl daemon-reload
    (( ! was_active )) || systemctl start ai-chat.service || true
    for service in tg-ai-cabinet telegram-ai-cabinet telegram-ai-bot; do
      [[ ! -f "$stage/$service.active" ]] || systemctl start "$service.service" || true
    done
  fi
  echo "Installation failed. Previous files restored when available. Snapshot: ${snapshot:-not needed}." >&2
  exit "$code"
}
trap rollback ERR
trap 'rm -rf "$stage"' EXIT
apt-get update -qq
apt-get install -y -qq python3 python3-venv curl ca-certificates fonts-dejavu-core
python3 -c 'import sys; assert sys.version_info >= (3,10), "Python 3.10+ required"'
curl -fLsS --retry 3 "https://codeload.github.com/wseinsshop-alt/AI-Chat/tar.gz/$revision" -o "$stage/source.tar.gz"
mkdir "$stage/source" "$stage/previous"
tar -xzf "$stage/source.tar.gz" --strip-components=1 -C "$stage/source"
python3 - "$stage/source" <<'PY'
import ast,json,sys
from pathlib import Path
p=Path(sys.argv[1]); ast.parse((p/'main.py').read_text(encoding='utf-8'))
json.loads((p/'config.json').read_text(encoding='utf-8'))
PY
mkdir -p "$app" "$data"
chmod 700 "$data"
for file in main.py README.md config.json requirements.txt install.sh; do
  [[ ! -f "$app/$file" ]] || cp -p "$app/$file" "$stage/previous/$file"
done
[[ ! -f "$unit" ]] || cp -p "$unit" "$stage/previous/service"
python3 -m venv "$app/.venv"
"$app/.venv/bin/python" -m pip install --quiet -r "$stage/source/requirements.txt"
"$app/.venv/bin/python" -c 'import telebot,requests,cryptography,PIL,reportlab,pypdf,docx,openpyxl'
# Reuse the existing external credentials, including the previous cabinet install.
python3 - "$data" <<'PY'
import json,sys
from pathlib import Path
data=Path(sys.argv[1]); target=data/'settings.json'
if not target.exists():
    for folder in ('/opt/ai-chat','/opt/tg-ai-cabinet','/opt/telegram-ai-bot-v2','/opt/telegram-ai-bot'):
        path=Path(folder)/'config.json'
        if not path.exists(): continue
        try:
            old=json.loads(path.read_text(encoding='utf-8'))
            token=old.get('telegram_token') or old.get('token') or old.get('bot_token')
            ids=old.get('admin_ids') or [old.get('admin') or old.get('owner_id')]
            if isinstance(ids,str): ids=ids.split(',')
            if token and ids and all(ids):
                target.write_text(json.dumps({'telegram_token':token,'admin_ids':[int(x) for x in ids]}),encoding='utf-8')
                target.chmod(0o600); break
        except (ValueError,TypeError): continue
PY
systemctl is-active --quiet ai-chat.service && was_active=1 || true
for service in tg-ai-cabinet telegram-ai-cabinet telegram-ai-bot; do
  systemctl is-active --quiet "$service.service" && touch "$stage/$service.active" || true
done
# Prepare credentials in staging so a cancelled setup leaves a running bot untouched.
if (( configure )) || [[ ! -f "$data/settings.json" ]]; then
  mkdir -p "$stage/setup"
  [[ ! -f "$data/settings.json" ]] || cp "$data/settings.json" "$stage/setup/settings.json"
  AI_CHAT_DATA_DIR="$stage/setup" "$app/.venv/bin/python" "$stage/source/main.py" --setup </dev/tty
fi
mutated=1
systemctl stop ai-chat.service 2>/dev/null || true
for service in tg-ai-cabinet telegram-ai-cabinet telegram-ai-bot; do
  [[ ! -f "$stage/$service.active" ]] || systemctl stop "$service.service"
done
snapshot="$data/restore-points/$(date -u +%Y%m%dT%H%M%S)-$$"
mkdir -p "$snapshot"
python3 - "$data" "$snapshot" <<'PY'
import shutil,sqlite3,sys
from contextlib import closing
from pathlib import Path
data,snapshot=map(Path,sys.argv[1:])
for name in ('secret.key','settings.json'):
    if (data/name).exists(): shutil.copy2(data/name,snapshot/name)
if (data/'bot.db').exists():
    with closing(sqlite3.connect(data/'bot.db')) as src, closing(sqlite3.connect(snapshot/'bot.db')) as dst: src.backup(dst)
else:
    old=Path('/opt/tg-ai-cabinet/data')
    if (old/'bot.db').exists() and (old/'secret.key').exists():
        with closing(sqlite3.connect(old/'bot.db')) as src, closing(sqlite3.connect(data/'bot.db')) as dst: src.backup(dst)
        shutil.copy2(old/'secret.key',data/'secret.key')
PY
if [[ -f "$stage/setup/settings.json" ]]; then cp "$stage/setup/settings.json" "$data/settings.json"; fi
for file in main.py README.md requirements.txt install.sh; do cp "$stage/source/$file" "$app/$file"; done
# Preserve local behavior/price settings on updates.
[[ -f "$app/config.json" ]] || cp "$stage/source/config.json" "$app/config.json"
AI_CHAT_DATA_DIR="$data" PYTHONDONTWRITEBYTECODE=1 "$app/.venv/bin/python" -c "import sys; sys.path.insert(0,'$app'); import main; c=main.configuration(); main.Store(c['data_path'])"
cat > "$unit" <<'SERVICE'
[Unit]
Description=AI Chat Telegram Bot
Wants=network-online.target
After=network-online.target

[Service]
WorkingDirectory=/opt/ai-chat
Environment=AI_CHAT_DATA_DIR=/var/lib/ai-chat
Environment=PYTHONDONTWRITEBYTECODE=1
ExecStart=/opt/ai-chat/.venv/bin/python -u /opt/ai-chat/main.py
Restart=always
RestartSec=5
UMask=0077
NoNewPrivileges=true

[Install]
WantedBy=multi-user.target
SERVICE
systemctl daemon-reload
systemctl enable ai-chat.service >/dev/null
systemctl restart ai-chat.service
sleep 5
systemctl is-active --quiet ai-chat.service
[[ $(systemctl show ai-chat.service -p NRestarts --value) == 0 ]] || { echo 'Bot restarted during startup.' >&2; false; }
for service in tg-ai-cabinet telegram-ai-cabinet telegram-ai-bot; do
  [[ ! -f "$stage/$service.active" ]] || systemctl disable "$service.service" 2>/dev/null || true
done
trap - ERR
echo 'AI Chat installed and running. Send /start in Telegram.'
echo "Rollback snapshot: $snapshot"
exit 0
