#!/usr/bin/env python3
"""auditslip_audit_employee.py — Employee-level audit helpers.

Phase 1: Reconcile slips ↔ bank_ledger per account
Phase 2: Daily variance per employee (transferor_name × date)
Phase 3: Cross-bot duplicate fingerprint detection
"""
from __future__ import annotations

import hashlib
import datetime as dt
import re
import sqlite3
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Helpers (inline so module is self-contained)
# ---------------------------------------------------------------------------

def _clean(v: Any) -> str:
    return str(v or "").strip()


def _float(v: Any) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _connect(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def _table_exists(conn: sqlite3.Connection, table: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)
    ).fetchone()
    return row is not None


def _bkk_today() -> str:
    return (dt.datetime.utcnow() + dt.timedelta(hours=7)).strftime("%Y-%m-%d")


def _scope_filter(scope: str, date_col: str, settlement_col: str = "settlement_id") -> Tuple[str, List[Any]]:
    raw = _clean(scope) or "open"
    low = raw.lower()
    if low in {"open", "current"}:
        return f"({settlement_col} IS NULL OR {settlement_col}='')", []
    if low in {"all", "__all__", "ทั้งหมด"}:
        return "1=1", []
    if low in {"today", "วันนี้"}:
        return f"{date_col}=?", [_bkk_today()]
    text = raw[6:] if low.startswith("range:") else raw
    if ".." in text:
        start, end = [p.strip() for p in text.split("..", 1)]
        if start and end and start > end:
            start, end = end, start
        if start and end:
            return f"{date_col} BETWEEN ? AND ?", [start, end]
        if start:
            return f"{date_col}>=?", [start]
        if end:
            return f"{date_col}<=?", [end]
        return "1=1", []
    if re.match(r"^\d{4}-\d{2}-\d{2}$", raw):
        return f"{date_col}=?", [raw]
    return f"{date_col}=?", [raw]


def _slip_flow_filter(flow_type: str, alias: str = "") -> str:
    flow = _clean(flow_type).lower()
    prefix = f"{alias}." if alias else ""
    title = f"LOWER(COALESCE({prefix}chat_title,''))"
    if flow == "deposit":
        return f"({title} LIKE '%ฝาก%' OR {title} LIKE '%เติม%' OR {title} LIKE '%deposit%' OR {title} LIKE '%topup%')"
    if flow == "withdraw":
        return f"({title} LIKE '%ถอน%' OR {title} LIKE '%withdraw%')"
    return "1=1"


def _now_iso() -> str:
    return (dt.datetime.utcnow() + dt.timedelta(hours=7)).isoformat(timespec="seconds") + "+07:00"


def ensure_employee_tables(conn: sqlite3.Connection) -> None:
    """Create non-destructive employee master/mapping tables if missing."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS employee_master (
          employee_id TEXT PRIMARY KEY,
          display_name TEXT NOT NULL,
          role TEXT NOT NULL DEFAULT 'employee',
          active INTEGER NOT NULL DEFAULT 1,
          notes TEXT NOT NULL DEFAULT '',
          created_at_iso TEXT NOT NULL,
          updated_at_iso TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS employee_aliases (
          alias_key TEXT PRIMARY KEY,
          employee_id TEXT NOT NULL,
          alias_type TEXT NOT NULL,
          alias_value TEXT NOT NULL,
          source TEXT NOT NULL DEFAULT 'telegram_sender',
          active INTEGER NOT NULL DEFAULT 1,
          created_at_iso TEXT NOT NULL,
          updated_at_iso TEXT NOT NULL,
          FOREIGN KEY(employee_id) REFERENCES employee_master(employee_id)
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_employee_aliases_employee ON employee_aliases(employee_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_employee_aliases_active ON employee_aliases(active, alias_type)")
    conn.commit()


def _employee_alias_key(alias_type: str, alias_value: Any) -> str:
    value = _clean(alias_value)
    if alias_type in {"telegram_username", "telegram_sender"}:
        value = value.lower()
    return f"{alias_type}:{value}" if value else ""


def _employee_aliases_from_row(row: Any) -> List[Tuple[str, str, str]]:
    def get(key: str) -> Any:
        try:
            return row[key]
        except Exception:
            return getattr(row, key, "")

    candidates = [
        ("telegram_user", _clean(get("user_id"))),
        ("telegram_username", _clean(get("username")).lstrip("@")),
        ("telegram_sender", _clean(get("sender_name"))),
    ]
    out: List[Tuple[str, str, str]] = []
    seen = set()
    for alias_type, alias_value in candidates:
        key = _employee_alias_key(alias_type, alias_value)
        if key and key not in seen:
            seen.add(key)
            out.append((key, alias_type, alias_value))
    return out


def _employee_id_for_alias(alias_key: str) -> str:
    return "emp_" + hashlib.sha256(alias_key.encode("utf-8")).hexdigest()[:12]


def _employee_display_from_row(row: Any, fallback_id: str) -> str:
    def get(key: str) -> Any:
        try:
            return row[key]
        except Exception:
            return getattr(row, key, "")

    return _clean(get("sender_name")) or _clean(get("username")).lstrip("@") or _clean(get("transferor_name")) or fallback_id


def _employee_alias_case_sql(alias: str = "s") -> str:
    prefix = f"{alias}." if alias else ""
    return (
        f"CASE "
        f"WHEN NULLIF(TRIM(COALESCE({prefix}user_id,'')), '') IS NOT NULL THEN 'telegram_user:' || TRIM(COALESCE({prefix}user_id,'')) "
        f"WHEN NULLIF(TRIM(COALESCE({prefix}username,'')), '') IS NOT NULL THEN 'telegram_username:' || LOWER(TRIM(REPLACE(COALESCE({prefix}username,''),'@',''))) "
        f"WHEN NULLIF(TRIM(COALESCE({prefix}sender_name,'')), '') IS NOT NULL THEN 'telegram_sender:' || LOWER(TRIM(COALESCE({prefix}sender_name,''))) "
        f"ELSE '' END"
    )


def sync_employee_master_from_slips(
    db_path: Path,
    bot_key: str = "",
    chat_id: str = "",
    scope: str = "open",
    flow_type: str = "all",
    limit: int = 5000,
) -> Dict[str, Any]:
    """Upsert a non-destructive employee master from Telegram sender identity.

    This does not alter slip rows. It creates/updates only `employee_master` and
    `employee_aliases`, using Telegram `user_id` first, then username, then sender_name.
    """
    with _connect(db_path) as conn:
        ensure_employee_tables(conn)
        bot = _clean(bot_key)
        cid = _clean(chat_id)
        where_parts = ["status='success'", "COALESCE(is_duplicate,0)=0"]
        params: List[Any] = []
        if bot and bot not in {"all", "__all__"}:
            where_parts.append("COALESCE(bot_key,'default')=?")
            params.append(bot)
        if cid:
            where_parts.append("chat_id=?")
            params.append(cid)
        scope_clause, scope_params = _scope_filter(scope, "slip_date_iso")
        where_parts.append(scope_clause)
        params.extend(scope_params)
        where_parts.append(_slip_flow_filter(flow_type))
        where = " AND ".join(where_parts)
        rows = conn.execute(
            f"""
            SELECT user_id, username, sender_name, transferor_name,
                   COALESCE(bot_key,'default') AS bot_key, company_name,
                   COUNT(*) AS slip_count, SUM(amount) AS total_amount
            FROM slips
            WHERE {where}
              AND (NULLIF(TRIM(COALESCE(user_id,'')), '') IS NOT NULL
                   OR NULLIF(TRIM(COALESCE(username,'')), '') IS NOT NULL
                   OR NULLIF(TRIM(COALESCE(sender_name,'')), '') IS NOT NULL)
            GROUP BY user_id, username, sender_name, transferor_name, bot_key, company_name
            ORDER BY slip_count DESC, total_amount DESC
            LIMIT ?
            """,
            [*params, limit],
        ).fetchall()

        now = _now_iso()
        inserted_employees = 0
        inserted_aliases = 0
        employee_ids: List[str] = []
        seen_ids = set()
        for row in rows:
            aliases = _employee_aliases_from_row(row)
            if not aliases:
                continue
            primary_key = aliases[0][0]
            employee_id = _employee_id_for_alias(primary_key)
            display_name = _employee_display_from_row(row, employee_id)
            cur = conn.execute(
                """
                INSERT OR IGNORE INTO employee_master(employee_id, display_name, role, active, notes, created_at_iso, updated_at_iso)
                VALUES(?, ?, 'employee', 1, '', ?, ?)
                """,
                (employee_id, display_name, now, now),
            )
            inserted_employees += max(0, cur.rowcount)
            conn.execute("UPDATE employee_master SET updated_at_iso=? WHERE employee_id=?", (now, employee_id))
            if employee_id not in seen_ids:
                seen_ids.add(employee_id)
                employee_ids.append(employee_id)
            for alias_key, alias_type, alias_value in aliases:
                cur = conn.execute(
                    """
                    INSERT OR IGNORE INTO employee_aliases(alias_key, employee_id, alias_type, alias_value, source, active, created_at_iso, updated_at_iso)
                    VALUES(?, ?, ?, ?, 'telegram_sender', 1, ?, ?)
                    """,
                    (alias_key, employee_id, alias_type, alias_value, now, now),
                )
                inserted_aliases += max(0, cur.rowcount)
                conn.execute(
                    "UPDATE employee_aliases SET updated_at_iso=?, active=1 WHERE alias_key=?",
                    (now, alias_key),
                )
        conn.commit()

        employees: List[Dict[str, Any]] = []
        if employee_ids:
            placeholders = ",".join("?" for _ in employee_ids)
            employee_rows = conn.execute(
                f"""
                SELECT em.employee_id, em.display_name, em.role, em.active,
                       COUNT(ea.alias_key) AS alias_count
                FROM employee_master em
                LEFT JOIN employee_aliases ea ON ea.employee_id=em.employee_id AND COALESCE(ea.active,1)=1
                WHERE em.employee_id IN ({placeholders})
                GROUP BY em.employee_id, em.display_name, em.role, em.active
                ORDER BY em.display_name
                """,
                employee_ids,
            ).fetchall()
            employees = [dict(r) for r in employee_rows]

    return {
        "ok": True,
        "inserted_employees": inserted_employees,
        "inserted_aliases": inserted_aliases,
        "employee_count": len(employees),
        "employees": employees,
    }


# ---------------------------------------------------------------------------
# Phase 3 helper: fingerprint schema migration
# ---------------------------------------------------------------------------

def ensure_dup_fingerprint_column(conn: sqlite3.Connection) -> None:
    """Add dup_fingerprint column to slips if not present."""
    cols = [row[1] for row in conn.execute("PRAGMA table_info(slips)").fetchall()]
    if "dup_fingerprint" not in cols:
        conn.execute("ALTER TABLE slips ADD COLUMN dup_fingerprint TEXT")
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_slips_dup_fingerprint "
            "ON slips(dup_fingerprint) WHERE dup_fingerprint IS NOT NULL"
        )
        conn.commit()


def _make_fingerprint(amount: float, slip_date_iso: str, reference_no: str) -> Optional[str]:
    """Return 16-char hex fingerprint or None if not enough data."""
    ref = _clean(reference_no)
    date = _clean(slip_date_iso)
    if not (amount > 0 and (ref or date)):
        return None
    raw = f"{round(amount, 2)}|{date}|{ref}"
    return hashlib.sha256(raw.encode()).hexdigest()[:16]


def backfill_dup_fingerprints(db_path: Path, limit: int = 5000) -> int:
    """Compute and store dup_fingerprint for slips that don't have one yet.
    Returns number of rows updated."""
    with _connect(db_path) as conn:
        ensure_dup_fingerprint_column(conn)
        rows = conn.execute(
            """SELECT id, amount, slip_date_iso, reference_no
               FROM slips
               WHERE dup_fingerprint IS NULL AND status='success'
               LIMIT ?""",
            (limit,),
        ).fetchall()
        updated = 0
        for r in rows:
            fp = _make_fingerprint(_float(r["amount"]), _clean(r["slip_date_iso"]), _clean(r["reference_no"]))
            if fp:
                conn.execute("UPDATE slips SET dup_fingerprint=? WHERE id=?", (fp, r["id"]))
                updated += 1
        conn.commit()
    return updated


# ---------------------------------------------------------------------------
# Phase 3: Cross-bot duplicate detection
# ---------------------------------------------------------------------------

def cross_bot_duplicates(
    db_path: Path,
    bot_key: str = "",
    scope: str = "open",
    limit: int = 200,
) -> Dict[str, Any]:
    """Find slips that share the same dup_fingerprint across different bot_key/chat_id.

    Returns groups where the same transaction appears in more than one source.
    """
    with _connect(db_path) as conn:
        ensure_dup_fingerprint_column(conn)
        # backfill on the fly (small batch)
        backfill_dup_fingerprints(db_path, limit=2000)

        bot = _clean(bot_key)
        scope_clause, scope_params = _scope_filter(scope, "slip_date_iso")
        bot_having = (
            "AND SUM(CASE WHEN COALESCE(bot_key,'default')=? THEN 1 ELSE 0 END) > 0"
            if bot and bot not in {"all", "__all__"}
            else ""
        )
        bot_params: list = [bot] if bot_having else []

        # Find fingerprints that appear with different bot_key/chat_id combos
        rows = conn.execute(
            f"""
            SELECT dup_fingerprint,
                   COUNT(DISTINCT bot_key||':'||chat_id) AS source_count,
                   COUNT(*) AS slip_count,
                   SUM(amount) AS total_amount,
                   MIN(slip_date_iso) AS earliest_date,
                   MAX(slip_date_iso) AS latest_date
            FROM slips
            WHERE dup_fingerprint IS NOT NULL
              AND status='success'
              AND {scope_clause}
            GROUP BY dup_fingerprint
            HAVING source_count > 1
              {bot_having}
            ORDER BY total_amount DESC, source_count DESC
            LIMIT ?
            """,
            [*scope_params, *bot_params, limit],
        ).fetchall()

        groups: List[Dict[str, Any]] = []
        for row in rows:
            fp = row["dup_fingerprint"]
            detail_rows = conn.execute(
                f"""
                SELECT id, bot_key, company_name, chat_id, chat_title,
                       transferor_name, sender_name, amount,
                       slip_date_iso, slip_date_display, reference_no,
                       is_duplicate, created_at_iso
                FROM slips
                WHERE dup_fingerprint=? AND status='success' AND {scope_clause}
                ORDER BY created_at_iso
                """,
                [fp, *scope_params],
            ).fetchall()
            groups.append({
                "fingerprint": fp,
                "source_count": row["source_count"],
                "slip_count": row["slip_count"],
                "total_amount": _float(row["total_amount"]),
                "earliest_date": _clean(row["earliest_date"]),
                "latest_date": _clean(row["latest_date"]),
                "slips": [dict(r) for r in detail_rows],
            })

    return {
        "ok": True,
        "group_count": len(groups),
        "total_suspicious_amount": sum(_float(g["total_amount"]) - _float(g["slips"][0]["amount"] if g["slips"] else 0) for g in groups),
        "groups": groups,
    }


# ---------------------------------------------------------------------------
# Phase 2: Daily variance per employee
# ---------------------------------------------------------------------------

def employee_daily_variance(
    db_path: Path,
    bot_key: str = "",
    chat_id: str = "",
    scope: str = "open",
    flow_type: str = "all",
    threshold: float = 100.0,
    auto_sync: bool = True,
) -> Dict[str, Any]:
    """Group slips by employee master identity × slip date and compute totals.

    Employee identity prefers Telegram sender mapping in `employee_master` /
    `employee_aliases`, then falls back to raw sender/username/transferor text.
    If bank_ledger is available, compare slip total vs ledger total per day.
    Flags rows where |variance| > threshold.
    """
    sync_result: Dict[str, Any] | None = None
    if auto_sync:
        sync_result = sync_employee_master_from_slips(db_path, bot_key=bot_key, chat_id=chat_id, scope=scope, flow_type=flow_type)
    with _connect(db_path) as conn:
        ensure_employee_tables(conn)
        bot = _clean(bot_key)
        cid = _clean(chat_id)

        where_parts = ["s.status='success'", "COALESCE(s.is_duplicate,0)=0"]
        params: List[Any] = []

        if bot and bot not in {"all", "__all__"}:
            where_parts.append("COALESCE(s.bot_key,'default')=?")
            params.append(bot)
        if cid:
            where_parts.append("s.chat_id=?")
            params.append(cid)

        scope_clause, scope_params = _scope_filter(scope, "s.slip_date_iso")
        where_parts.append(scope_clause)
        params.extend(scope_params)
        where_parts.append(_slip_flow_filter(flow_type, alias="s"))

        where = " AND ".join(where_parts)
        alias_expr = _employee_alias_case_sql("s")

        rows = conn.execute(
            f"""
            SELECT
                COALESCE(NULLIF(em.display_name,''), NULLIF(s.sender_name,''), NULLIF(s.username,''), NULLIF(s.transferor_name,''), '(ไม่ทราบ)') AS employee,
                COALESCE(em.employee_id, '') AS employee_id,
                CASE WHEN em.employee_id IS NOT NULL THEN 'employee_master'
                     WHEN NULLIF(TRIM(COALESCE(s.sender_name,'')), '') IS NOT NULL OR NULLIF(TRIM(COALESCE(s.username,'')), '') IS NOT NULL THEN 'raw_sender'
                     ELSE 'raw_transferor' END AS identity_source,
                COALESCE(s.bot_key,'default') AS bot_key,
                s.company_name,
                s.chat_id,
                COALESCE(s.slip_date_iso, '') AS date_key,
                COUNT(*) AS slip_count,
                SUM(s.amount) AS slip_total
            FROM slips s
            LEFT JOIN employee_aliases ea ON ea.alias_key=({alias_expr}) AND COALESCE(ea.active,1)=1
            LEFT JOIN employee_master em ON em.employee_id=ea.employee_id AND COALESCE(em.active,1)=1
            WHERE {where}
            GROUP BY 1, 2, 3, 4, 5, 6, 7
            ORDER BY date_key DESC, slip_total DESC
            """,
            params,
        ).fetchall()

        # Check if bank_ledger_entries exists for comparison
        has_ledger = _table_exists(conn, "bank_ledger_entries")
        ledger_by_date: Dict[str, float] = {}
        if has_ledger:
            led_params: List[Any] = []
            led_where = "1=1"
            if bot and bot not in {"all", "__all__"}:
                led_where += " AND COALESCE(bot_key,'default')=?"
                led_params.append(bot)
            led_scope_clause, led_scope_params = _scope_filter(scope, "date_key", "''")
            led_where += f" AND {led_scope_clause}"
            led_params.extend(led_scope_params)
            if flow_type in {"deposit", "withdraw"}:
                led_where += " AND flow_type=?"
                led_params.append(flow_type)
            ledger_rows = conn.execute(
                f"""
                SELECT date_key, SUM(amount) AS ledger_total
                FROM bank_ledger_entries
                WHERE flow_type IN ('deposit','withdraw','all') AND {led_where}
                GROUP BY date_key
                """,
                led_params,
            ).fetchall()
            ledger_by_date = {r["date_key"]: _float(r["ledger_total"]) for r in ledger_rows}

        result_rows: List[Dict[str, Any]] = []
        flagged_count = 0

        # Aggregate further by employee across all their days
        emp_summary: Dict[str, Dict[str, Any]] = {}

        for r in rows:
            emp = _clean(r["employee"])
            date = _clean(r["date_key"])
            slip_total = _float(r["slip_total"])
            ledger_total = ledger_by_date.get(date, None)
            variance = (slip_total - ledger_total) if ledger_total is not None else None
            flagged = variance is not None and abs(variance) > threshold

            employee_id = _clean(r["employee_id"])
            identity_source = _clean(r["identity_source"])
            row_dict = {
                "employee": emp,
                "employee_id": employee_id,
                "identity_source": identity_source,
                "bot_key": _clean(r["bot_key"]),
                "company_name": _clean(r["company_name"]),
                "chat_id": _clean(r["chat_id"]),
                "date": date,
                "slip_count": int(r["slip_count"] or 0),
                "slip_total": slip_total,
                "ledger_total": ledger_total,
                "variance": variance,
                "flagged": flagged,
            }
            result_rows.append(row_dict)
            if flagged:
                flagged_count += 1

            # aggregate per employee
            emp_key = employee_id or f"raw:{emp}"
            if emp_key not in emp_summary:
                emp_summary[emp_key] = {"employee": emp, "employee_id": employee_id, "identity_source": identity_source, "total_slips": 0, "total_amount": 0.0, "flagged_days": 0, "days": []}
            emp_summary[emp_key]["total_slips"] += int(r["slip_count"] or 0)
            emp_summary[emp_key]["total_amount"] += slip_total
            if flagged:
                emp_summary[emp_key]["flagged_days"] += 1
            emp_summary[emp_key]["days"].append(row_dict)

        emp_list = sorted(emp_summary.values(), key=lambda x: x["flagged_days"], reverse=True)

    return {
        "ok": True,
        "has_ledger": has_ledger,
        "identity_mode": "employee_master",
        "employee_sync": sync_result,
        "threshold": threshold,
        "flagged_count": flagged_count,
        "employee_count": len(emp_list),
        "employees": emp_list,
        "rows": result_rows,
    }


# ---------------------------------------------------------------------------
# Phase 1: Reconcile slips ↔ bank_ledger per account
# ---------------------------------------------------------------------------

def _parse_hhmm(t: str) -> Optional[int]:
    """Return minutes-since-midnight from 'HH:MM' or 'HH:MM:SS'. None if invalid."""
    t = _clean(t)
    if not t:
        return None
    parts = t.split(":")
    if len(parts) < 2:
        return None
    try:
        h = int(parts[0])
        m = int(parts[1])
    except (TypeError, ValueError):
        return None
    if not (0 <= h < 24 and 0 <= m < 60):
        return None
    return h * 60 + m


def _score_match(slip: Dict[str, Any], entry: Dict[str, Any], time_tol_min: int = 5) -> int:
    """Simple matcher: amount + date must match exactly; time within tolerance.

    Returns score >= 1 if matched, -1 if not.
    Ignores names, references, descriptions completely (per user request).
    """
    # 1. Amount must match (within rounding)
    if abs(_float(slip.get("amount")) - _float(entry.get("amount"))) > 0.009:
        return -1

    # 2. Date must match exactly (both present)
    slip_date = _clean(slip.get("slip_date_iso") or slip.get("date_key"))
    entry_date = _clean(entry.get("date_key") or entry.get("date"))
    if slip_date and entry_date:
        if slip_date != entry_date:
            return -1
    else:
        # if either side missing date, allow but lower score
        return 1

    # 3. Time check (within tolerance)
    slip_min = _parse_hhmm(slip.get("slip_time") or "")
    entry_min = _parse_hhmm(entry.get("time") or "")
    if slip_min is not None and entry_min is not None:
        diff = abs(slip_min - entry_min)
        if diff > time_tol_min:
            return -1
        return 100 - diff  # closer time = higher score
    # if one side missing time, accept on amount+date only
    return 50


def reconcile_slips_ledger(
    db_path: Path,
    bot_key: str = "",
    chat_id: str = "",
    account_key: str = "",
    scope: str = "open",
    flow_type: str = "all",
    limit: int = 500,
) -> Dict[str, Any]:
    """Match slips against bank_ledger_entries for an account.

    Returns three buckets:
    - matched: slip ↔ ledger entry paired
    - slip_only: in slips, not found in ledger
    - ledger_only: in ledger, not found in slips
    """
    with _connect(db_path) as conn:
        ensure_dup_fingerprint_column(conn)
        has_ledger = _table_exists(conn, "bank_ledger_entries")

        bot = _clean(bot_key)
        cid = _clean(chat_id)
        acct = _clean(account_key)

        # --- load slips ---
        slip_where_parts = ["status='success'"]
        slip_params: List[Any] = []
        if bot and bot not in {"all", "__all__"}:
            slip_where_parts.append("COALESCE(bot_key,'default')=?")
            slip_params.append(bot)
        if cid:
            slip_where_parts.append("chat_id=?")
            slip_params.append(cid)
        if acct:
            slip_where_parts.append(
                "(COALESCE(to_account,'') LIKE ? OR COALESCE(from_account,'') LIKE ?)"
            )
            slip_params.extend([f"%{acct}%", f"%{acct}%"])
        scope_clause, scope_params = _scope_filter(scope, "slip_date_iso")
        slip_where_parts.append(scope_clause)
        slip_params.extend(scope_params)
        slip_where_parts.append(_slip_flow_filter(flow_type))
        slip_where_parts.append("COALESCE(is_duplicate,0)=0")

        slip_where = " AND ".join(slip_where_parts)
        slip_rows = conn.execute(
            f"""SELECT id, bot_key, company_name, chat_id, chat_title,
                       transferor_name, sender_name, recipient_name,
                       from_bank, from_account, to_bank, to_account, issuer_bank,
                       amount, reference_no, seq, aid,
                       slip_date_iso, slip_time, created_at_iso,
                       is_duplicate, dup_fingerprint
                FROM slips
                WHERE {slip_where}
                ORDER BY COALESCE(slip_date_iso,''), id
                LIMIT ?""",
            [*slip_params, limit],
        ).fetchall()
        slips = [dict(r) for r in slip_rows]

        # --- load bank_ledger_entries ---
        ledger_entries: List[Dict[str, Any]] = []
        if has_ledger:
            led_where_parts = ["1=1"]
            led_params: List[Any] = []
            if bot and bot not in {"all", "__all__"}:
                led_where_parts.append("COALESCE(bot_key,'default')=?")
                led_params.append(bot)
            if acct:
                led_where_parts.append("(account_no LIKE ? OR account_key LIKE ?)")
                led_params.extend([f"%{acct}%", f"%{acct}%"])
            led_scope_clause, led_scope_params = _scope_filter(scope, "date_key", "''")
            led_where_parts.append(led_scope_clause)
            led_params.extend(led_scope_params)
            if flow_type in {"deposit", "withdraw"}:
                led_where_parts.append("flow_type=?")
                led_params.append(flow_type)
            led_where = " AND ".join(led_where_parts)
            ledger_rows = conn.execute(
                f"""SELECT entry_id, bot_key, account_key, company_name,
                           bank, account_no, account_name,
                           date, date_key, time, flow_type,
                           amount, reference, sender, receiver, description
                    FROM bank_ledger_entries
                    WHERE {led_where}
                    ORDER BY date_key, time, row_no
                    LIMIT ?""",
                [*led_params, limit],
            ).fetchall()
            ledger_entries = [dict(r) for r in ledger_rows]

    # --- match ---
    used_slips: set = set()
    used_entries: set = set()
    matched: List[Dict[str, Any]] = []
    slip_only: List[Dict[str, Any]] = []
    ledger_only: List[Dict[str, Any]] = []

    # For each ledger entry, find best slip match
    for ei, entry in enumerate(ledger_entries):
        best_si = -1
        best_score = 0
        for si, slip in enumerate(slips):
            if si in used_slips:
                continue
            score = _score_match(slip, entry)
            if score > best_score:
                best_score = score
                best_si = si
        if best_si >= 0 and best_score > 0:
            used_slips.add(best_si)
            used_entries.add(ei)
            matched.append({
                "slip": slips[best_si],
                "entry": entry,
                "score": best_score,
                "amount": _float(entry.get("amount")),
            })
        else:
            ledger_only.append(entry)

    for si, slip in enumerate(slips):
        if si not in used_slips:
            slip_only.append(slip)

    slip_total = sum(_float(s.get("amount")) for s in slips)
    ledger_total = sum(_float(e.get("amount")) for e in ledger_entries)
    matched_amount = sum(m["amount"] for m in matched)

    return {
        "ok": True,
        "has_ledger": has_ledger,
        "scope": {"bot_key": bot or "__all__", "chat_id": cid, "account_key": acct, "flow_type": flow_type, "scope": scope},
        "summary": {
            "slip_count": len(slips),
            "slip_total": slip_total,
            "ledger_count": len(ledger_entries),
            "ledger_total": ledger_total,
            "matched_count": len(matched),
            "matched_amount": matched_amount,
            "slip_only_count": len(slip_only),
            "slip_only_amount": sum(_float(s.get("amount")) for s in slip_only),
            "ledger_only_count": len(ledger_only),
            "ledger_only_amount": sum(_float(e.get("amount")) for e in ledger_only),
            "diff_amount": slip_total - ledger_total,
        },
        "matched": matched[:200],
        "slip_only": slip_only[:200],
        "ledger_only": ledger_only[:200],
    }
