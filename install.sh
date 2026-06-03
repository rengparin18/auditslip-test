#!/usr/bin/env bash
# Auditslip one-shot installer.
# Run on Ubuntu/Debian VPS as root:
#   curl -fsSL https://raw.githubusercontent.com/aamsainz1-ui/auditslip/main/install.sh | sudo bash
# Or from a cloned repo:
#   sudo bash install.sh
set -Eeuo pipefail

REPO_URL="${AUDITSLIP_REPO_URL:-https://github.com/aamsainz1-ui/auditslip.git}"
INSTALL_DIR="${AUDITSLIP_INSTALL_DIR:-/root/projects/auditslip}"
ENV_DIR="${AUDITSLIP_ENV_DIR:-/etc/auditslip}"
ENV_FILE="${AUDITSLIP_ENV_FILE:-$ENV_DIR/auditslip.env}"
NON_INTERACTIVE=0
START_SERVICES=1

for arg in "$@"; do
  case "$arg" in
    --non-interactive) NON_INTERACTIVE=1 ;;
    --no-start) START_SERVICES=0 ;;
    --help|-h)
      cat <<'EOF'
Auditslip installer

Options:
  --non-interactive   install files/services only; do not prompt for secrets
  --no-start          install/update systemd units but do not start services

Environment overrides:
  AUDITSLIP_REPO_URL      Git repository URL
  AUDITSLIP_INSTALL_DIR   install path (default /root/projects/auditslip)
  AUDITSLIP_ENV_FILE      env file path (default /etc/auditslip/auditslip.env)
EOF
      exit 0
      ;;
    *) echo "Unknown option: $arg" >&2; exit 2 ;;
  esac
done

log() { printf '\n\033[1;32m==>\033[0m %s\n' "$*"; }
warn() { printf '\n\033[1;33mWARN:\033[0m %s\n' "$*" >&2; }
need_root() {
  if [[ "${EUID:-$(id -u)}" -ne 0 ]]; then
    echo "Please run as root, e.g. sudo bash install.sh" >&2
    exit 1
  fi
}
rand_secret() {
  if command -v openssl >/dev/null 2>&1; then
    openssl rand -hex 24
  else
    python3 - <<'PY'
import secrets
print(secrets.token_hex(24))
PY
  fi
}
set_env_value() {
  local key="$1" value="$2"
  ENV_FILE="$ENV_FILE" KEY="$key" VALUE="$value" python3 - <<'PY'
import os
from pathlib import Path
path = Path(os.environ['ENV_FILE'])
key = os.environ['KEY']
value = os.environ['VALUE']
line = f"{key}={value}\n"
text = path.read_text(encoding='utf-8') if path.exists() else ''
lines = text.splitlines(keepends=True)
updated = False
for i, old in enumerate(lines):
    if old.startswith(key + '='):
        lines[i] = line
        updated = True
        break
if not updated:
    if lines and not lines[-1].endswith('\n'):
        lines[-1] += '\n'
    lines.append(line)
path.write_text(''.join(lines), encoding='utf-8')
PY
}
prompt_value() {
  local key="$1" label="$2" secret="${3:-0}" current=""
  current="$(grep -E "^${key}=" "$ENV_FILE" 2>/dev/null | tail -1 | cut -d= -f2- || true)"
  if [[ "$current" != "" && "$current" != "CHANGE_ME" ]]; then
    printf "%s already set; press Enter to keep, or type a new value.\n" "$key"
  fi
  local value=""
  if [[ "$secret" == "1" ]]; then
    read -r -s -p "$label: " value </dev/tty || true
    printf '\n' >/dev/tty
  else
    read -r -p "$label: " value </dev/tty || true
  fi
  if [[ -n "$value" ]]; then
    set_env_value "$key" "$value"
  fi
}
copy_unit() {
  local name="$1" src="$INSTALL_DIR/systemd/$name" dst="/etc/systemd/system/$name"
  python3 - "$src" "$dst" "$INSTALL_DIR" <<'PY'
from pathlib import Path
import sys
src, dst, install_dir = map(Path, sys.argv[1:4])
text = src.read_text(encoding='utf-8')
text = text.replace('/root/projects/auditslip', str(install_dir))
dst.write_text(text, encoding='utf-8')
PY
}

need_root

log "1/8 Install OS packages"
apt-get update
DEBIAN_FRONTEND=noninteractive apt-get install -y git curl sqlite3 python3 python3-requests python3-openpyxl ca-certificates openssl

log "2/8 Clone/update repository"
mkdir -p "$(dirname "$INSTALL_DIR")"
if [[ -d "$INSTALL_DIR/.git" ]]; then
  git -C "$INSTALL_DIR" fetch origin
  git -C "$INSTALL_DIR" checkout main
  git -C "$INSTALL_DIR" pull --ff-only origin main
else
  git clone "$REPO_URL" "$INSTALL_DIR"
fi
cd "$INSTALL_DIR"

log "3/8 Create runtime directories"
mkdir -p "$ENV_DIR" \
  "$INSTALL_DIR/data/slip-images" \
  "$INSTALL_DIR/exports" \
  "$INSTALL_DIR/imports/backend" \
  "$INSTALL_DIR/backups/db"

log "4/8 Create/update env file"
if [[ ! -f "$ENV_FILE" ]]; then
  cp "$INSTALL_DIR/.env.example" "$ENV_FILE"
  chmod 600 "$ENV_FILE"
else
  chmod 600 "$ENV_FILE"
  warn "Env file exists; installer will not overwrite secrets: $ENV_FILE"
fi
set_env_value AUDITSLIP_HOME "$INSTALL_DIR"
set_env_value AUDITSLIP_DATA_DIR "$INSTALL_DIR/data"
set_env_value AUDITSLIP_EXPORT_DIR "$INSTALL_DIR/exports"
set_env_value AUDITSLIP_BACKEND_IMPORT_DIR "$INSTALL_DIR/imports/backend"
set_env_value AUDITSLIP_DB "$INSTALL_DIR/data/auditslip.db"
set_env_value OCR_PROVIDER_HEALTH_PATH "$INSTALL_DIR/data/ocr-provider-health.json"
set_env_value AUDITSLIP_BACKUP_DIR "$INSTALL_DIR/backups/db"

if grep -q '^AUDITSLIP_DASHBOARD_TOKEN=CHANGE_ME\|^AUDITSLIP_DASHBOARD_TOKEN=$' "$ENV_FILE"; then
  set_env_value AUDITSLIP_DASHBOARD_TOKEN "$(rand_secret)"
fi
if grep -q '^AUDITSLIP_DASHBOARD_OWNER_PASSWORD=CHANGE_ME\|^AUDITSLIP_DASHBOARD_OWNER_PASSWORD=$' "$ENV_FILE"; then
  set_env_value AUDITSLIP_DASHBOARD_OWNER_PASSWORD "$(rand_secret)"
fi

if [[ "$NON_INTERACTIVE" -eq 0 && -r /dev/tty ]]; then
  log "5/8 Interactive setup (press Enter to skip/keep existing values)"
  prompt_value BOT_TOKEN "Telegram BOT_TOKEN for single-bot mode" 1
  prompt_value GEMINI_API_KEY "Gemini API key (recommended OCR provider)" 1
  prompt_value OPENAI_API_KEY "OpenAI API key (optional fallback OCR provider)" 1
  prompt_value AUDITSLIP_DASHBOARD_OWNER_USER "Dashboard owner username" 0
  prompt_value AUDITSLIP_DASHBOARD_OWNER_PASSWORD "Dashboard owner password" 1
  prompt_value AUDITSLIP_FLOW_MAP "Flow map, e.g. bot1|-100111=deposit,bot1|-100222=withdraw" 0
  cat <<'EOF'

Multi-bot setup:
  Edit /etc/auditslip/auditslip.env after this script if you need several companies:
    BOT_TOKEN_1=<token>
    BOT_TOKEN_2=<token>
    AUDITSLIP_TELEGRAM_BOTS="bot1:BOT_TOKEN_1:บริษัท 1,bot2:BOT_TOKEN_2:บริษัท 2"
EOF
else
  log "5/8 Non-interactive setup skipped prompts"
  warn "Edit $ENV_FILE and fill BOT_TOKEN or AUDITSLIP_TELEGRAM_BOTS plus OCR keys before starting production."
fi

log "6/8 Install systemd units"
copy_unit auditslip-bot.service
copy_unit auditslip-dashboard.service
copy_unit auditslip-bot-watchdog.service
copy_unit auditslip-bot-watchdog.timer
copy_unit auditslip-backup.service
copy_unit auditslip-backup.timer
systemctl daemon-reload

log "7/8 Validate code and configuration"
python3 -m py_compile auditslip_bot.py auditslip_dashboard.py auditslip_watchdog.py auditslip_bank_ledger.py tools/verify_audit_chain.py
python3 tests/check_auditslip_product_contract.py

log "8/8 Enable services"
if [[ "$START_SERVICES" -eq 1 ]]; then
  bot_token_value="$(grep -E '^BOT_TOKEN=' "$ENV_FILE" 2>/dev/null | tail -1 | cut -d= -f2- || true)"
  multi_bot_value="$(grep -E '^AUDITSLIP_TELEGRAM_BOTS=' "$ENV_FILE" 2>/dev/null | tail -1 | cut -d= -f2- || true)"
  gemini_value="$(grep -E '^GEMINI_API_KEY=' "$ENV_FILE" 2>/dev/null | tail -1 | cut -d= -f2- || true)"
  google_value="$(grep -E '^GOOGLE_API_KEY=' "$ENV_FILE" 2>/dev/null | tail -1 | cut -d= -f2- || true)"
  openai_value="$(grep -E '^OPENAI_API_KEY=' "$ENV_FILE" 2>/dev/null | tail -1 | cut -d= -f2- || true)"
  if [[ -z "$bot_token_value" && ( -z "$multi_bot_value" || "$multi_bot_value" == *BOT_TOKEN_1* ) ]]; then
    warn "No Telegram bot token configured yet; services will be enabled but not started. Edit $ENV_FILE then run: systemctl restart auditslip-bot auditslip-dashboard"
    START_SERVICES=0
  fi
  if [[ -z "$gemini_value" && -z "$google_value" && -z "$openai_value" ]]; then
    warn "No OCR API key configured yet; services will be enabled but not started. Edit $ENV_FILE then run: systemctl restart auditslip-bot auditslip-dashboard"
    START_SERVICES=0
  fi
fi
if [[ "$START_SERVICES" -eq 1 ]]; then
  systemctl enable auditslip-bot.service auditslip-dashboard.service auditslip-bot-watchdog.timer auditslip-backup.timer
  systemctl restart auditslip-bot.service auditslip-dashboard.service
  systemctl start auditslip-bot-watchdog.timer auditslip-backup.timer
  sleep 2
  systemctl --no-pager --full status auditslip-bot.service auditslip-dashboard.service | sed -n '1,80p' || true
  curl -fsS 'http://127.0.0.1:8095/api/health?quick=1' || true
else
  systemctl enable auditslip-bot.service auditslip-dashboard.service auditslip-bot-watchdog.timer auditslip-backup.timer
  warn "Services enabled but not started because --no-start was used."
fi

cat <<EOF

Done.
Project: $INSTALL_DIR
Env:     $ENV_FILE
Dashboard local: http://127.0.0.1:8095/

Next steps:
  1. Verify $ENV_FILE has Telegram/OCR keys.
  2. Add your bot to Telegram groups.
  3. Run: systemctl is-active auditslip-bot.service auditslip-dashboard.service
  4. Open dashboard and login with AUDITSLIP_DASHBOARD_OWNER_USER/PASSWORD.
EOF
