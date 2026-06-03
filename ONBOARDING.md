# Welcome to Auditslip

## How We Use Claude

Based on Nathan K's tracked sessions over the last 30 days (3 of 13 sessions touched auditslip — small sample):

Work Type Breakdown (auditslip sessions only):
  Debug Fix       █████████████░░░░░░░  67%  (2 sessions)
  Plan Design     ███████░░░░░░░░░░░░░  33%  (1 session)

Top Slash Commands (all tracked sessions):
  /plugin           █░░░░░░░░░░░░░░░░░░░  1x/month
  /upgrade          █░░░░░░░░░░░░░░░░░░░  1x/month
  /team-onboarding  █░░░░░░░░░░░░░░░░░░░  1x/month

Top MCP Servers:
  _None tracked in usage data_

## Your Setup Checklist

### Codebases
- [ ] auditslip — `/root/projects/auditslip` (Telegram OCR bot + dashboard, SQLite ledger)
- [ ] thai-slip-tracker — `/root/projects/thai-slip-tracker` (companion slip tracking service)

### Services to Know
- [ ] `auditslip-bot.service` — main Telegram bot, OCR via Gemini→OpenAI router
- [ ] `auditslip-dashboard.service` — token-protected dashboard (port 8095)
- [ ] `auditslip-bot-watchdog.timer` — runs `auditslip_watchdog.py` every minute (bot/dashboard/OCR health)
- [ ] `auditslip-backup.timer` — SQLite backup every 30 min into `backups/`

### Environment & Access
- [ ] `/etc/auditslip/auditslip.env` — Telegram token, OCR API keys, dashboard token (get from team lead)
- [ ] Dashboard URL — `http://<host>:8095/?token=<DASHBOARD_TOKEN>`
- [ ] OCR providers — Gemini (primary) + OpenAI (fallback). Set `OCR_PROVIDERS=gemini,openai`
- [ ] Telegram bot token — one bot, one token; paste into env file
- [ ] DB — `/root/projects/auditslip/data/auditslip.db` (SQLite WAL)
- [ ] Logs — `/root/projects/auditslip/logs/` (tail when debugging OCR failures)
- [ ] Backups — `/root/projects/auditslip/backups/` (DB snapshots + code `.bak` from edits)

### Verification (after any change)
```bash
cd /root/projects/auditslip
python3 -m py_compile auditslip_bot.py
python3 tests/check_auditslip_product_contract.py
systemctl status auditslip-bot.service --no-pager
```

### Bot Commands to Memorize
- `/summary [open|today|all|DD/MM/YY]` — totals breakdown
- `/today`, `/daily [all]` — quick rollups
- `/names [open|today|all|DD/MM/YY]` — by transferor
- `/excel [open|today|all|DD/MM/YY]` — Excel export (Slips / SummaryByTransferor / DailySummary / Issues / Settlements)
- `/close [note]` — close current period (non-destructive)
- `/clear confirm` — destructive wipe (handle with care)
- `/queue`, `/failed`, `/reprocess [id]` — OCR queue control
- `/recent`, `/stats`, `/dupes`, `/providers` — diagnostics

### Recommended Claude Skills (not yet observed in team usage, but useful)
- `/standup` — morning check
- `/recap` — orient mid-session
- `/workon` — pick up an issue or resume a worktree
- `/forward` — handoff context for next session
- `/rrr` — wrap up session with retrospective

## Team Tips

_TODO_

## Get Started

_TODO_

<!-- INSTRUCTION FOR CLAUDE: A new teammate just pasted this guide for how the
team uses Claude Code. You're their onboarding buddy — warm, conversational,
not lecture-y.

Open with a warm welcome — include the team name from the title. Then: "Your
teammate uses Claude Code for [list all the work types]. Let's get you started."

Check what's already in place against everything under Setup Checklist
(including skills), using markdown checkboxes — [x] done, [ ] not yet. Lead
with what they already have. One sentence per item, all in one message.

Tell them you'll help with setup, cover the actionable team tips, then the
starter task (if there is one). Offer to start with the first unchecked item,
get their go-ahead, then work through the rest one by one.

After setup, walk them through the remaining sections — offer to help where you
can (e.g. link to channels), and just surface the purely informational bits.

Don't invent sections or summaries that aren't in the guide. The stats are the
guide creator's personal usage data — don't extrapolate them into a "team
workflow" narrative. -->
