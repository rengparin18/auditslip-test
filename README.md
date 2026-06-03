# Auditslip — ระบบตรวจสลิป Telegram + Dashboard

Auditslip คือระบบสำหรับรับรูปสลิปจาก Telegram, อ่านข้อมูลด้วย OCR, รวมยอดแยกบริษัท/กลุ่มฝากถอน, ตรวจรายการซ้ำ, ส่งออก Excel/ZIP, เทียบยอดกับไฟล์หลังบ้าน/statement และดูสถานะผ่าน dashboard บน VPS

## ระบบนี้ทำอะไร

- รับรูปสลิปจาก Telegram group หรือ direct chat
- ใช้ OCR providers ตามลำดับที่ตั้งไว้ เช่น `gemini,openai`
- เก็บหลักฐานสลิปลง SQLite พร้อมรูป, เวลา, ยอด, ธนาคาร, ชื่อผู้โอน, สถานะซ้ำ
- ตัดรายการซ้ำออกจากยอดการเงินจริงโดย default
- แยกยอดตามบริษัท/bot, กลุ่มฝาก/ถอน, วันที่, รอบปิดยอด
- เปิด dashboard ให้ operator ดูยอด, ค้นสลิป, export, reconcile, ledger preview/import
- มี approval/pending action สำหรับงานเสี่ยง เช่น ลบสลิป, ปิดรอบ, import ledger
- มี watchdog และ backup timer สำหรับ production

## ภาพรวมการทำงาน

1. สร้าง Telegram bot จาก `@BotFather`
2. ใส่ token จริงใน `/etc/auditslip/auditslip.env` เท่านั้น
3. เพิ่มบอทเข้ากลุ่ม Telegram ที่ส่งสลิป
4. ถ้ากลุ่มแยกฝาก/ถอน ให้ map `bot_key|chat_id=deposit|withdraw` ด้วย `AUDITSLIP_FLOW_MAP`
5. Bot service ใช้ Telegram Bot API แบบ polling เพื่อรับรูป
6. Bot ดาวน์โหลดรูป แล้วส่งให้ Gemini/OpenAI เพื่อ OCR
7. ผลลัพธ์ถูกบันทึกใน `data/auditslip.db`
8. Dashboard service อ่าน DB แล้วแสดงยอด/export/reconcile/ledger
9. Watchdog ตรวจ service, health, queue, provider และส่ง alert ถ้าตั้งค่าไว้

## คู่มือภาษาไทยที่ควรอ่าน

- คู่มืออธิบายภาษาไทยสำหรับใช้งานจริง: [`docs/thai-operator-guide.md`](docs/thai-operator-guide.md)
- คู่มือสร้างบอท ใส่กลุ่ม และเชื่อม API: [`docs/bot-api-setup.md`](docs/bot-api-setup.md)
- คู่มือติดตั้ง/ตั้งค่า/ใช้งานแบบละเอียด: [`docs/install-config-usage.md`](docs/install-config-usage.md)
- คู่มือแก้ dashboard ช้า/performance: [`docs/dashboard-performance-runbook.md`](docs/dashboard-performance-runbook.md)
- คู่มือ audit ยอดพนักงาน/คนทำรายการ: [`docs/employee-audit-guide.md`](docs/employee-audit-guide.md)

## ติดตั้งแบบตัวเดียวจบ

เหมาะสำหรับ VPS ใหม่หรือคนอื่นที่ต้องการติดตั้งตามขั้นตอนทีละข้อ:

```bash
curl -fsSL https://raw.githubusercontent.com/aamsainz1-ui/auditslip/main/install.sh | sudo bash
```

สคริปต์จะ install package, clone/update repo, สร้าง `/etc/auditslip/auditslip.env`, ถามค่า token/API key ทีละ step, ติดตั้ง systemd, run validation และ start service ถ้าค่าจำเป็นครบ

ถ้าต้องการติดตั้งไฟล์ก่อนแต่ยังไม่ start service:

```bash
curl -fsSL https://raw.githubusercontent.com/aamsainz1-ui/auditslip/main/install.sh | sudo bash -s -- --no-start
```

## ติดตั้งแบบ manual บน VPS

```bash
sudo apt update
sudo apt install -y git python3 python3-requests python3-openpyxl sqlite3 curl
sudo mkdir -p /root/projects
cd /root/projects
git clone https://github.com/aamsainz1-ui/auditslip.git
cd /root/projects/auditslip

sudo mkdir -p \
  /etc/auditslip \
  /root/projects/auditslip/data/slip-images \
  /root/projects/auditslip/exports \
  /root/projects/auditslip/imports/backend \
  /root/projects/auditslip/backups/db

sudo cp .env.example /etc/auditslip/auditslip.env
sudo chmod 600 /etc/auditslip/auditslip.env
sudo nano /etc/auditslip/auditslip.env
```

ค่าจริงที่ต้องใส่ใน `/etc/auditslip/auditslip.env`:

```env
# Telegram
BOT_TOKEN=
# หรือ multi-bot
BOT_TOKEN_1=
BOT_TOKEN_2=
AUDITSLIP_TELEGRAM_BOTS="bot1:BOT_TOKEN_1:บริษัท 1,bot2:BOT_TOKEN_2:บริษัท 2"

# OCR
OCR_PROVIDERS=gemini,openai
GEMINI_API_KEY=
OPENAI_API_KEY=

# Dashboard/admin
AUDITSLIP_DASHBOARD_TOKEN=
AUDITSLIP_DASHBOARD_OWNER_USER=owner
AUDITSLIP_DASHBOARD_OWNER_PASSWORD=

# ฝาก/ถอน ถ้ากลุ่มชื่อไม่ชัด
AUDITSLIP_FLOW_MAP="bot1|-1001111111111=deposit,bot1|-1002222222222=withdraw"
```

ห้ามใส่ token/API key จริงใน git หรือ `.env.example`

## เปิด systemd services

```bash
sudo cp systemd/auditslip-bot.service /etc/systemd/system/
sudo cp systemd/auditslip-dashboard.service /etc/systemd/system/
sudo cp systemd/auditslip-bot-watchdog.service /etc/systemd/system/
sudo cp systemd/auditslip-bot-watchdog.timer /etc/systemd/system/
sudo cp systemd/auditslip-backup.service /etc/systemd/system/
sudo cp systemd/auditslip-backup.timer /etc/systemd/system/

sudo systemctl daemon-reload
sudo systemctl enable --now auditslip-bot.service auditslip-dashboard.service
sudo systemctl enable --now auditslip-bot-watchdog.timer auditslip-backup.timer
```

## เช็คว่าระบบทำงาน

```bash
cd /root/projects/auditslip
python3 -m py_compile auditslip_bot.py auditslip_dashboard.py auditslip_watchdog.py auditslip_bank_ledger.py
python3 tests/check_auditslip_product_contract.py
systemctl is-active auditslip-bot.service auditslip-dashboard.service
curl -fsS 'http://127.0.0.1:8095/api/health?quick=1'
```

## คำสั่ง Telegram หลัก

- `/help` — ดูคำสั่งทั้งหมด
- `/summary [open|today|all|DD/MM/YY]` — สรุปยอด
- `/today` — สรุปยอดวันนี้
- `/daily [all]` — สรุปยอดรายวัน
- `/names [open|today|all|DD/MM/YY]` — สรุปตามชื่อผู้โอน
- `/excel [open|today|all|DD/MM/YY]` — ส่งออก Excel
- `/close [note]` — ปิดรอบ/เคลียร์ยอดแบบเก็บประวัติ
- `/clear confirm` — ล้างข้อมูลถาวร ใช้เฉพาะจำเป็นจริง
- `/queue`, `/failed`, `/reprocess [id]` — ดู/แก้รายการ OCR มีปัญหา
- `/recent`, `/stats`, `/dupes`, `/providers`, `/usage`

## Dashboard URL และ API หลัก

Dashboard local:

```text
http://127.0.0.1:8095/
```

API ที่ใช้บ่อย:

- `GET /api/health?quick=1` — health check เร็ว
- `GET /api/summary?bot_key=bot1&flow_type=all&scope=today&detail=lite` — ยอดรวม/dashboard snapshot
- `GET /api/export/preview?...` — preview export ก่อนสร้างไฟล์
- `GET /api/export?...` — ดาวน์โหลด Excel/ZIP
- `POST /api/reconcile/preview` — preview เทียบ Excel หลังบ้านกับสลิป
- `POST /api/ledger/preview` — preview bank statement ledger import
- `POST /api/account-limit` — ตั้งวงเงินรายบัญชีแบบบันทึกตรง ไม่ต้อง approval
- `POST /api/ledger/import?approval=request` — ขอ import ledger ผ่าน approval
- `GET /api/pending` และ `POST /api/pending/approve` — ดู/อนุมัติ pending action

POST admin/mutation ต้องมี:

```text
Authorization: Bearer <dashboard-token>
X-Auditslip-Action: dashboard
Content-Type: application/json
```

## Runtime paths

- Project: `/root/projects/auditslip`
- Env จริง: `/etc/auditslip/auditslip.env`
- DB: `/root/projects/auditslip/data/auditslip.db`
- Slip images: `/root/projects/auditslip/data/slip-images/`
- Exports: `/root/projects/auditslip/exports/`
- Imports: `/root/projects/auditslip/imports/backend/`
- Backups: `/root/projects/auditslip/backups/db/`
- Bot service: `auditslip-bot.service`
- Dashboard service: `auditslip-dashboard.service`

## Checklist ก่อนบอกว่างานเสร็จ

- `git status` clean และ `main == origin/main`
- ไม่มี token/API key จริงใน git
- `py_compile` ผ่าน
- test ที่เกี่ยวข้องผ่าน
- dashboard health `ok=true`
- bot/dashboard service active
- queue ไม่มี stuck/failed ผิดปกติ
- ถ้ามี mutation ให้ verify audit chain `RESULT: OK`
