#!/usr/bin/env python3
"""Production watchdog for Auditslip bot/dashboard.

Checks are intentionally cheap and evidence-focused: service state, dashboard health,
OCR queue health, today's counted total sanity, and optional Telegram pending probe.
Secrets are never printed; Telegram alerts are throttled by fingerprint.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sqlite3
import subprocess
import sys
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

APP_DIR = Path(__file__).resolve().parent
DEFAULT_ENV_FILE = Path("/etc/auditslip/auditslip.env")
DEFAULT_DB = APP_DIR / "data" / "auditslip.db"
DEFAULT_STATE = APP_DIR / "data" / "watchdog-state.json"
DEFAULT_RESTART_LOG = APP_DIR / "data" / "watchdog_restart_log.json"
RESTART_WINDOW_MIN = 60
RESTART_MAX_PER_WINDOW = 3
BKK = timezone(timedelta(hours=7))

SECRET_KEY_RE = re.compile(r"(token|secret|password|api[_-]?key)", re.I)


@dataclass
class Alert:
    code: str
    severity: str
    title: str
    detail: str
    data: dict[str, Any] = field(default_factory=dict)

    def public(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "severity": self.severity,
            "title": self.title,
            "detail": self.detail,
            "data": scrub(self.data),
        }


def scrub(value: Any) -> Any:
    """Remove raw secrets and verbose provider errors from JSON/log output."""
    if isinstance(value, dict):
        out: dict[str, Any] = {}
        for key, item in value.items():
            if SECRET_KEY_RE.search(str(key)):
                out[key] = "[REDACTED]"
            elif str(key).lower() in {"error", "raw_error", "exception", "traceback"}:
                out[key] = "[REDACTED]"
            else:
                out[key] = scrub(item)
        return out
    if isinstance(value, list):
        return [scrub(item) for item in value]
    if isinstance(value, str):
        if re.search(r"\d{8,}:[A-Za-z0-9_-]{20,}", value):
            return "[REDACTED]"
        return value[:240]
    return value


def parse_env_line(line: str) -> tuple[str, str] | None:
    line = line.strip()
    if not line or line.startswith("#") or "=" not in line:
        return None
    if line.startswith("export "):
        line = line[7:].strip()
    key, value = line.split("=", 1)
    key = key.strip()
    value = value.strip()
    if not key:
        return None
    if (value.startswith("\"") and value.endswith("\"")) or (value.startswith("'") and value.endswith("'")):
        value = value[1:-1]
    return key, value


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(errors="ignore").splitlines():
        parsed = parse_env_line(line)
        if not parsed:
            continue
        key, value = parsed
        os.environ.setdefault(key, value)


def env_bool(name: str, default: bool = False) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "y", "on"}


def env_int(name: str, default: int) -> int:
    try:
        return int(str(os.environ.get(name, default)).strip())
    except Exception:
        return default


def now_ms() -> int:
    return int(time.time() * 1000)


def bkk_today_iso() -> str:
    return datetime.now(BKK).date().isoformat()


def run_cmd(cmd: list[str], timeout: int = 8) -> subprocess.CompletedProcess[str]:
    return subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=timeout)


def table_exists(conn: sqlite3.Connection, table: str) -> bool:
    row = conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone()
    return bool(row)


def check_services(args: argparse.Namespace, alerts: list[Alert], report: dict[str, Any]) -> None:
    if args.no_systemctl:
        report["services"] = {"skipped": True}
        return
    services = [
        (os.environ.get("AUDITSLIP_BOT_SERVICE", "auditslip-bot.service"), "bot"),
        (os.environ.get("AUDITSLIP_DASHBOARD_SERVICE", "auditslip-dashboard.service"), "dashboard"),
    ]
    auto_restart = env_bool("AUDITSLIP_WATCHDOG_AUTO_RESTART", True)
    restart_log = Path(os.environ.get("AUDITSLIP_WATCHDOG_RESTART_LOG") or DEFAULT_RESTART_LOG)
    report["services"] = {}
    report["restarts"] = []
    for service, label in services:
        proc = run_cmd(["systemctl", "is-active", service])
        state = proc.stdout.strip() or proc.stderr.strip() or "unknown"
        report["services"][label] = {"unit": service, "state": state}
        if proc.returncode == 0 and state == "active":
            continue
        alerts.append(
            Alert(
                "service_down",
                "critical",
                f"{label} service is not active",
                f"{service} state={state}",
                {"service": service, "state": state},
            )
        )
        if not auto_restart:
            continue
        now = datetime.now(BKK)
        allowed, recent = check_restart_backoff(restart_log, service, now)
        attempt = len(recent) + 1
        if not allowed:
            alerts.append(
                Alert(
                    "restart_blocked_backoff",
                    "critical",
                    f"{label} restart backoff tripped",
                    f"{service} restarted >={RESTART_MAX_PER_WINDOW}x/hour, NOT restarting, investigate",
                    {"service": service, "recent_count": len(recent), "window_min": RESTART_WINDOW_MIN},
                )
            )
            report["restarts"].append({"service": service, "skipped": "backoff", "recent_count": len(recent)})
            continue
        alerts.append(
            Alert(
                "service_crash_restart",
                "warning",
                f"{label} crashed, restarting",
                f"{service} crashed, restarting (attempt {attempt}/{RESTART_MAX_PER_WINDOW})",
                {"service": service, "attempt": attempt, "max": RESTART_MAX_PER_WINDOW},
            )
        )
        if args.dry_run:
            report["restarts"].append({"service": service, "dry_run": True, "attempt": attempt})
            continue
        restart = run_cmd(["systemctl", "restart", service], timeout=20)
        record_restart(restart_log, service, now)
        report["restarts"].append({"service": service, "exit_code": restart.returncode, "attempt": attempt})


def check_dashboard(args: argparse.Namespace, alerts: list[Alert], report: dict[str, Any]) -> None:
    if args.skip_dashboard:
        report["dashboard_health"] = {"skipped": True}
        return
    url = os.environ.get("AUDITSLIP_WATCHDOG_HEALTH_URL") or f"http://127.0.0.1:{os.environ.get('AUDITSLIP_DASHBOARD_PORT', '8095')}/api/health?quick=1"
    status: dict[str, Any] = {"url": re.sub(r"token=[^&]+", "token=[REDACTED]", url)}
    try:
        with urllib.request.urlopen(url, timeout=5) as resp:
            body = resp.read(4096)
        data = json.loads(body.decode("utf-8"))
        ok = bool(data.get("ok"))
        status.update({"ok": ok})
    except Exception as exc:
        ok = False
        status.update({"ok": False, "error": "[REDACTED]", "error_type": type(exc).__name__})
    report["dashboard_health"] = status
    if not ok:
        alerts.append(Alert("dashboard_health_failed", "critical", "Dashboard health endpoint failed", "dashboard /api/health did not return ok", status))


def check_queue_and_totals(args: argparse.Namespace, alerts: list[Alert], report: dict[str, Any]) -> None:
    db_path = Path(args.db or os.environ.get("AUDITSLIP_DB") or DEFAULT_DB)
    report["db"] = str(db_path)
    if not db_path.exists():
        alerts.append(Alert("db_missing", "critical", "Auditslip DB missing", "database file not found", {"db": str(db_path)}))
        return
    stale_minutes = env_int("AUDITSLIP_WATCHDOG_STALE_MINUTES", 15)
    failed_threshold = env_int("AUDITSLIP_WATCHDOG_FAILED_THRESHOLD", 1)
    cutoff = now_ms() - stale_minutes * 60_000
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        if table_exists(conn, "ocr_jobs"):
            rows = conn.execute("SELECT status, COUNT(*) AS count FROM ocr_jobs GROUP BY status ORDER BY status").fetchall()
            report["ocr_jobs"] = {str(r["status"]): int(r["count"] or 0) for r in rows}
            stale = conn.execute(
                """
                SELECT job_id, COALESCE(NULLIF(bot_key,''),'default') AS bot_key, chat_id, message_id, status,
                       created_at, updated_at
                FROM ocr_jobs
                WHERE status IN ('queued','processing') AND COALESCE(NULLIF(updated_at,0), created_at) < ?
                ORDER BY COALESCE(NULLIF(updated_at,0), created_at) ASC
                LIMIT 10
                """,
                (cutoff,),
            ).fetchall()
            if stale:
                oldest_ms = int(stale[0]["updated_at"] or stale[0]["created_at"] or 0)
                alerts.append(
                    Alert(
                        "queue_stale",
                        "critical",
                        "OCR queue has stale queued/processing jobs",
                        f"{len(stale)} sampled jobs older than {stale_minutes} minutes",
                        {
                            "sample_count": len(stale),
                            "oldest_age_min": round((now_ms() - oldest_ms) / 60000, 1) if oldest_ms else None,
                            "samples": [
                                {"job_id": r["job_id"], "bot_key": r["bot_key"], "chat_id": r["chat_id"], "message_id": r["message_id"], "status": r["status"]}
                                for r in stale[:5]
                            ],
                        },
                    )
                )
            failed_count = int(report["ocr_jobs"].get("failed", 0))
            if failed_count >= failed_threshold:
                failed_rows = conn.execute(
                    """
                    SELECT job_id, COALESCE(NULLIF(bot_key,''),'default') AS bot_key, chat_id, message_id, attempts, max_attempts, updated_at
                    FROM ocr_jobs
                    WHERE status='failed'
                    ORDER BY updated_at DESC
                    LIMIT 5
                    """
                ).fetchall()
                alerts.append(
                    Alert(
                        "ocr_failed",
                        "warning",
                        "OCR jobs failed",
                        f"{failed_count} OCR jobs are failed",
                        {"count": failed_count, "samples": [dict(r) for r in failed_rows]},
                    )
                )
        if table_exists(conn, "slips"):
            today = os.environ.get("AUDITSLIP_WATCHDOG_TODAY") or bkk_today_iso()
            row = conn.execute(
                """
                SELECT COUNT(*) AS count, COALESCE(SUM(amount),0) AS amount
                FROM slips
                WHERE status='success' AND COALESCE(is_duplicate,0)=0 AND COALESCE(slip_date_iso,'')=?
                """,
                (today,),
            ).fetchone()
            today_count = int(row["count"] or 0)
            today_amount = float(row["amount"] or 0)
            report["today"] = {"date": today, "success_nonduplicate_count": today_count, "success_nonduplicate_amount": today_amount}
            min_count = env_int("AUDITSLIP_WATCHDOG_MIN_TODAY_COUNT", 0)
            min_amount = float(os.environ.get("AUDITSLIP_WATCHDOG_MIN_TODAY_AMOUNT", "0") or 0)
            if (min_count and today_count < min_count) or (min_amount and today_amount < min_amount):
                alerts.append(
                    Alert(
                        "today_total_anomaly",
                        "warning",
                        "Today's counted total is below configured floor",
                        "success/non-duplicate total is lower than watchdog threshold",
                        {"date": today, "count": today_count, "amount": today_amount, "min_count": min_count, "min_amount": min_amount},
                    )
                )


def telegram_alert_target() -> tuple[str, str]:
    token = os.environ.get("AUDITSLIP_WATCHDOG_BOT_TOKEN") or os.environ.get("BOT_TOKEN") or os.environ.get("TELEGRAM_BOT_TOKEN") or ""
    chat_id = os.environ.get("AUDITSLIP_WATCHDOG_ALERT_CHAT_ID") or next((x.strip() for x in os.environ.get("AUDITSLIP_ADMIN_IDS", "").split(",") if x.strip()), "")
    return token, chat_id


def alert_fingerprint(alert: Alert) -> str:
    data = alert.data or {}
    key_parts = [alert.code, str(data.get("service", "")), str(data.get("job_id", "")), str(data.get("date", ""))]
    return "|".join(key_parts)


def load_state(path: Path) -> dict[str, Any]:
    try:
        return json.loads(path.read_text()) if path.exists() else {}
    except Exception:
        return {}


def save_state(path: Path, state: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2))
    tmp.replace(path)


def check_restart_backoff(
    log_path: Path,
    service: str,
    now: datetime,
    window_min: int = RESTART_WINDOW_MIN,
    max_restarts: int = RESTART_MAX_PER_WINDOW,
) -> tuple[bool, list[str]]:
    """Return (allowed, recent_timestamps_in_window) for a candidate restart."""
    try:
        log = json.loads(log_path.read_text()) if log_path.exists() else {}
    except Exception:
        log = {}
    entries = log.get(service, []) if isinstance(log, dict) else []
    cutoff = now - timedelta(minutes=window_min)
    recent: list[str] = []
    for ts in entries:
        try:
            parsed = datetime.fromisoformat(ts)
        except Exception:
            continue
        if parsed >= cutoff:
            recent.append(ts)
    return (len(recent) < max_restarts, recent)


def record_restart(log_path: Path, service: str, ts: datetime, window_min: int = RESTART_WINDOW_MIN) -> None:
    import fcntl
    log_path.parent.mkdir(parents=True, exist_ok=True)
    lock_path = log_path.with_suffix(log_path.suffix + ".lock")
    with open(lock_path, "a+") as lock_fh:
        fcntl.flock(lock_fh.fileno(), fcntl.LOCK_EX)
        try:
            try:
                log = json.loads(log_path.read_text()) if log_path.exists() else {}
            except Exception:
                log = {}
            if not isinstance(log, dict):
                log = {}
            entries = list(log.get(service, []))
            entries.append(ts.isoformat())
            cutoff = ts - timedelta(minutes=window_min)
            pruned: list[str] = []
            for entry in entries:
                try:
                    parsed = datetime.fromisoformat(entry)
                except Exception:
                    continue
                if parsed >= cutoff:
                    pruned.append(entry)
            log[service] = pruned
            tmp = log_path.with_suffix(log_path.suffix + ".tmp")
            tmp.write_text(json.dumps(log, ensure_ascii=False, indent=2))
            tmp.replace(log_path)
        finally:
            fcntl.flock(lock_fh.fileno(), fcntl.LOCK_UN)


def send_alerts(args: argparse.Namespace, alerts: list[Alert], report: dict[str, Any]) -> None:
    report["alert_delivery"] = {"enabled": False, "sent": 0, "throttled": 0}
    if args.no_telegram or args.dry_run or not alerts:
        return
    token, chat_id = telegram_alert_target()
    if not token or not chat_id:
        report["alert_delivery"]["reason"] = "missing target"
        return
    state_path = Path(args.state_file or os.environ.get("AUDITSLIP_WATCHDOG_STATE") or DEFAULT_STATE)
    state = load_state(state_path)
    sent_at = state.setdefault("sent_at", {})
    throttle_sec = env_int("AUDITSLIP_WATCHDOG_ALERT_THROTTLE_SEC", 1800)
    now = int(time.time())
    deliverable: list[Alert] = []
    for alert in alerts:
        fp = alert_fingerprint(alert)
        last = int(sent_at.get(fp, 0) or 0)
        if now - last >= throttle_sec:
            deliverable.append(alert)
            sent_at[fp] = now
        else:
            report["alert_delivery"]["throttled"] += 1
    if not deliverable:
        save_state(state_path, state)
        return
    lines = ["🚨 Auditslip watchdog"]
    for alert in deliverable[:8]:
        lines.append(f"- {alert.severity}: {alert.title} ({alert.code})")
        if alert.detail:
            lines.append(f"  {alert.detail}")
    if len(deliverable) > 8:
        lines.append(f"- +{len(deliverable)-8} more alerts")
    text = "\n".join(lines)
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = urllib.parse.urlencode({"chat_id": chat_id, "text": text}).encode("utf-8")
    try:
        with urllib.request.urlopen(url, data=data, timeout=8) as resp:
            resp.read(2048)
        report["alert_delivery"] = {"enabled": True, "sent": len(deliverable), "throttled": report["alert_delivery"].get("throttled", 0)}
        save_state(state_path, state)
    except Exception as exc:
        report["alert_delivery"] = {"enabled": True, "sent": 0, "error": "[REDACTED]", "error_type": type(exc).__name__}


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Auditslip production watchdog")
    p.add_argument("--env-file", default=str(DEFAULT_ENV_FILE))
    p.add_argument("--db", default="")
    p.add_argument("--state-file", default=str(DEFAULT_STATE))
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--json", action="store_true", dest="json_output")
    p.add_argument("--no-systemctl", action="store_true")
    p.add_argument("--skip-dashboard", action="store_true")
    p.add_argument("--no-telegram", action="store_true")
    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    load_env_file(Path(args.env_file))
    alerts: list[Alert] = []
    report: dict[str, Any] = {"ok": True, "dry_run": bool(args.dry_run), "restarts": []}
    try:
        check_services(args, alerts, report)
        check_dashboard(args, alerts, report)
        check_queue_and_totals(args, alerts, report)
        report["alerts"] = [a.public() for a in alerts]
        report["ok"] = not any(a.severity == "critical" for a in alerts)
        send_alerts(args, alerts, report)
    except Exception as exc:
        report["ok"] = False
        report["alerts"] = [Alert("watchdog_error", "critical", "Watchdog crashed", "watchdog raised an unexpected exception", {"error_type": type(exc).__name__}).public()]
    output = scrub(report)
    if args.json_output:
        print(json.dumps(output, ensure_ascii=False, sort_keys=True))
    else:
        if output.get("alerts"):
            print(json.dumps(output, ensure_ascii=False, sort_keys=True))
        else:
            print("ok: auditslip watchdog clean")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
