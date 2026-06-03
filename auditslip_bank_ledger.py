#!/usr/bin/env python3
"""Bank-ledger reconciliation component for Auditslip.

This module owns the new per-account statement-ledger behavior: schema,
idempotent import, account-scoped matching, and read-only snapshots.  The
legacy dashboard injects its existing normalizers/parsers through configure()
so the component stays separate without broad dashboard rewrites.
"""
from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import time
from pathlib import Path
from typing import Any, Callable, Dict, List

_HELPERS: Dict[str, Any] = {}


def configure(**helpers: Any) -> None:
    _HELPERS.update(helpers)


def _value(name: str) -> Any:
    if name not in _HELPERS:
        raise RuntimeError(f"bank ledger helper not configured: {name}")
    return _HELPERS[name]


def _call(name: str, *args: Any, **kwargs: Any) -> Any:
    fn = _value(name)
    return fn(*args, **kwargs)


def clean_display(value: Any) -> str:
    return _call("clean_display", value)


def clean_company_name(value: Any, bot_key: Any = "") -> str:
    return _call("clean_company_name", value, bot_key)


def display_bank(value: Any) -> str:
    return _call("display_bank", value)


def account_key_for(company_name: Any, bank: Any, account_no: Any, account_name: Any) -> str:
    return _call("account_key_for", company_name, bank, account_no, account_name)


def normalize_match_date(value: Any) -> str:
    return _call("normalize_match_date", value)


def normalize_match_text(value: Any) -> str:
    return _call("normalize_match_text", value)


def normalize_flow_type(value: Any) -> str:
    return _call("normalize_flow_type", value)


def flow_label(value: Any) -> str:
    return _call("flow_label", value)


def sqlite_table_exists(conn: sqlite3.Connection, table: str) -> bool:
    return _call("sqlite_table_exists", conn, table)


def rows_to_dicts(rows: Any) -> List[Dict[str, Any]]:
    return _call("rows_to_dicts", rows)


def scope_date_range(scope: Any) -> Any:
    return _call("scope_date_range", scope)


def scope_to_date(scope: Any) -> Any:
    return _call("scope_to_date", scope)


def connect(db_path: Path) -> sqlite3.Connection:
    return _call("connect", db_path)


def parse_statement_file(path: Path) -> List[Dict[str, Any]]:
    return _call("parse_statement_file", path)


def filter_statement_reconcile_rows(rows: List[Dict[str, Any]], scope: str, flow_type: str) -> Any:
    return _call("filter_statement_reconcile_rows", rows, scope, flow_type)


def slip_reconcile_rows(conn: sqlite3.Connection, chat_id: str = "", scope: str = "all", bot_key: str = "", flow_type: str = "all") -> List[Dict[str, Any]]:
    return _call("slip_reconcile_rows", conn, chat_id=chat_id, scope=scope, bot_key=bot_key, flow_type=flow_type)


def amount_time_date_match(left: Dict[str, Any], right: Dict[str, Any]) -> bool:
    return _call("amount_time_date_match", left, right)


def reconcile_daily_summary(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return _call("reconcile_daily_summary", rows)


BANK_LEDGER_TABLES = ("bank_ledger_accounts", "bank_ledger_entries")


def bank_ledger_account_identity(bot_key: Any = "", company_name: Any = "", bank: Any = "", account_no: Any = "", account_name: Any = "") -> Dict[str, Any]:
    bot = clean_display(bot_key) or "default"
    company = clean_company_name(company_name, bot) or bot
    bank_display = display_bank(bank)
    account = clean_display(account_no)
    name = clean_display(account_name)
    key = account_key_for(company, bank_display, account, name)
    return {"bot_key": bot, "company_name": company, "bank": bank_display, "account_no": account, "account_name": name, "account_key": key}


def ensure_bank_ledger_tables(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS bank_ledger_accounts (
          bot_key TEXT NOT NULL DEFAULT 'default',
          account_key TEXT NOT NULL,
          company_name TEXT,
          bank TEXT,
          account_no TEXT,
          account_name TEXT,
          active INTEGER DEFAULT 1,
          updated_at INTEGER NOT NULL,
          PRIMARY KEY(bot_key, account_key)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS bank_ledger_entries (
          entry_id TEXT PRIMARY KEY,
          bot_key TEXT NOT NULL DEFAULT 'default',
          account_key TEXT NOT NULL,
          company_name TEXT,
          bank TEXT,
          account_no TEXT,
          account_name TEXT,
          source_name TEXT,
          source_hash TEXT NOT NULL,
          row_no INTEGER,
          date TEXT,
          date_key TEXT,
          time TEXT,
          flow_type TEXT,
          flow_label TEXT,
          description TEXT,
          amount REAL DEFAULT 0,
          reference TEXT,
          sender TEXT,
          receiver TEXT,
          imported_at INTEGER NOT NULL,
          raw_json TEXT,
          UNIQUE(bot_key, account_key, source_hash)
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_bank_ledger_bot_account_date ON bank_ledger_entries(bot_key, account_key, date_key, flow_type)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_bank_ledger_bot_date ON bank_ledger_entries(bot_key, date_key, flow_type)")


def bank_ledger_tables_exist(conn: sqlite3.Connection) -> bool:
    return all(sqlite_table_exists(conn, table) for table in BANK_LEDGER_TABLES)


def bank_ledger_source_hash(account: Dict[str, Any], row: Dict[str, Any]) -> str:
    amount = f"{float(row.get('amount') or 0):.2f}"
    parts = [
        clean_display(account.get("bot_key")),
        clean_display(account.get("account_key")),
        clean_display(row.get("source")),
        clean_display(row.get("row")),
        normalize_match_date(row.get("date_key") or row.get("date")),
        clean_display(row.get("time")),
        normalize_flow_type(row.get("flow_type")),
        amount,
        normalize_match_text(row.get("reference")),
        normalize_match_text(row.get("description")),
        normalize_match_text(row.get("sender")),
        normalize_match_text(row.get("receiver")),
    ]
    return hashlib.sha256("|".join(parts).encode("utf-8")).hexdigest()


def bank_ledger_entry_id(account: Dict[str, Any], source_hash: str) -> str:
    raw = f"{clean_display(account.get('bot_key'))}|{clean_display(account.get('account_key'))}|{source_hash}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def bank_ledger_entries_from_statement(statement_rows: List[Dict[str, Any]], account: Dict[str, Any]) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for row in statement_rows:
        flow = normalize_flow_type(row.get("flow_type"))
        if flow not in {"deposit", "withdraw"}:
            flow = clean_display(row.get("flow_type")) or ""
        source_hash = bank_ledger_source_hash(account, row)
        entry = {
            "entry_id": bank_ledger_entry_id(account, source_hash),
            "bot_key": account["bot_key"],
            "account_key": account["account_key"],
            "company_name": account["company_name"],
            "bank": account["bank"],
            "account_no": account["account_no"],
            "account_name": account["account_name"],
            "source_name": clean_display(row.get("source")),
            "source_hash": source_hash,
            "row_no": int(row.get("row") or 0),
            "date": clean_display(row.get("date")),
            "date_key": normalize_match_date(row.get("date_key") or row.get("date")),
            "time": clean_display(row.get("time")),
            "flow_type": flow,
            "flow_label": flow_label(flow) if flow else "",
            "description": clean_display(row.get("description")),
            "amount": float(row.get("amount") or 0),
            "reference": clean_display(row.get("reference")),
            "sender": clean_display(row.get("sender")),
            "receiver": clean_display(row.get("receiver")),
            "raw_json": json.dumps(row, ensure_ascii=False, default=str),
        }
        out.append(entry)
    return out


def account_no_matches(left: Any, right: Any) -> bool:
    left_key = normalize_match_text(left)
    right_key = normalize_match_text(right)
    return bool(left_key and right_key and left_key == right_key)


def slip_matches_bank_ledger_account(slip: Dict[str, Any], account: Dict[str, Any], flow_type: str = "") -> bool:
    flow = normalize_flow_type(flow_type or slip.get("flow_type") or "all")
    account_no = clean_display(account.get("account_no"))
    account_bank = display_bank(account.get("bank"))
    if flow == "deposit":
        if account_no:
            return account_no_matches(slip.get("to_account"), account_no)
        return bool(account_bank and display_bank(slip.get("to_bank")) == account_bank)
    if flow == "withdraw":
        if account_no:
            return account_no_matches(slip.get("from_account"), account_no)
        return bool(account_bank and display_bank(slip.get("from_bank") or slip.get("issuer_bank")) == account_bank)
    if account_no:
        return account_no_matches(slip.get("to_account"), account_no) or account_no_matches(slip.get("from_account"), account_no)
    return True


def filter_slips_for_bank_ledger_account(slips: List[Dict[str, Any]], account: Dict[str, Any], flow_type: str) -> List[Dict[str, Any]]:
    flow = normalize_flow_type(flow_type)
    return [slip for slip in slips if slip_matches_bank_ledger_account(slip, account, flow)]


def match_ledger_entries_to_slips(entries: List[Dict[str, Any]], slips: List[Dict[str, Any]], account: Dict[str, Any]) -> Dict[str, Any]:
    used_slips: set[int] = set()
    matched: List[Dict[str, Any]] = []
    unmatched_ledger: List[Dict[str, Any]] = []
    eligible_slip_indexes = {idx for idx, slip in enumerate(slips) if slip_matches_bank_ledger_account(slip, account, clean_display(slip.get("flow_type")) or "all")}
    for entry in entries:
        entry_flow = normalize_flow_type(entry.get("flow_type"))
        best_idx = -1
        for idx, slip in enumerate(slips):
            if idx in used_slips:
                continue
            if not slip_matches_bank_ledger_account(slip, account, entry_flow):
                continue
            if amount_time_date_match(entry, slip):
                best_idx = idx
                break
        if best_idx >= 0:
            used_slips.add(best_idx)
            eligible_slip_indexes.add(best_idx)
            matched.append({"ledger": entry, "slip": slips[best_idx], "confidence": "exact_amount_time_account"})
        else:
            unmatched_ledger.append(entry)
    unmatched_slips = [slip for idx, slip in enumerate(slips) if idx in eligible_slip_indexes and idx not in used_slips]
    matched_amount = sum(float(m["ledger"].get("amount") or 0) for m in matched)
    unmatched_ledger_amount = sum(float(r.get("amount") or 0) for r in unmatched_ledger)
    unmatched_slip_amount = sum(float(r.get("amount") or 0) for r in unmatched_slips)
    return {
        "matched": {"count": len(matched), "amount": matched_amount, "rows": matched[:100]},
        "ledger_extra": {"count": len(unmatched_ledger), "amount": unmatched_ledger_amount, "rows": unmatched_ledger[:100]},
        "slip_extra": {"count": len(unmatched_slips), "amount": unmatched_slip_amount, "rows": unmatched_slips[:100]},
    }


def existing_bank_ledger_hashes(conn: sqlite3.Connection, bot_key: str, account_key: str) -> set[str]:
    if not sqlite_table_exists(conn, "bank_ledger_entries"):
        return set()
    rows = conn.execute(
        "SELECT source_hash FROM bank_ledger_entries WHERE COALESCE(bot_key,'default')=? AND account_key=?",
        (clean_display(bot_key) or "default", clean_display(account_key)),
    ).fetchall()
    return {clean_display(r["source_hash"]) for r in rows}


def bank_ledger_query_rows(conn: sqlite3.Connection, bot_key: str = "", account_key: str = "", account_no: str = "", scope: str = "all", flow_type: str = "all", limit: int = 500) -> List[Dict[str, Any]]:
    if not sqlite_table_exists(conn, "bank_ledger_entries"):
        return []
    where = "1=1"
    params: List[Any] = []
    bot = clean_display(bot_key)
    if bot and bot not in {"__all__", "all"}:
        where += " AND COALESCE(bot_key,'default')=?"
        params.append(bot)
    if clean_display(account_key):
        where += " AND account_key=?"
        params.append(clean_display(account_key))
    flow = normalize_flow_type(flow_type)
    if flow in {"deposit", "withdraw", "other"}:
        where += " AND flow_type=?"
        params.append(flow)
    range_start, range_end, _ = scope_date_range(scope)
    normalized, _ = scope_to_date(scope)
    normalized = clean_display(normalized)
    if range_start or range_end:
        if range_start and range_end:
            where += " AND date_key BETWEEN ? AND ?"
            params.extend([range_start, range_end])
        elif range_start:
            where += " AND date_key >= ?"
            params.append(range_start)
        elif range_end:
            where += " AND date_key <= ?"
            params.append(range_end)
    elif normalized in {"", "all", "open"}:
        pass
    elif re.match(r"^\d{4}-\d{2}-\d{2}$", normalized):
        where += " AND date_key=?"
        params.append(normalized)
    rows = conn.execute(
        f"""
        SELECT * FROM bank_ledger_entries
        WHERE {where}
        ORDER BY date_key DESC, time DESC, row_no DESC, imported_at DESC
        LIMIT ?
        """,
        [*params, int(limit or 500)],
    ).fetchall()
    out = rows_to_dicts(rows)
    account_no_clean = clean_display(account_no)
    if account_no_clean:
        out = [row for row in out if account_no_matches(row.get("account_no"), account_no_clean)]
    return out


def import_bank_ledger_statement(
    db_path: Path,
    statement_path: Path,
    bot_key: str = "",
    company_name: str = "",
    bank: str = "",
    account_no: str = "",
    account_name: str = "",
    scope: str = "all",
    flow_type: str = "all",
    dry_run: bool = True,
) -> Dict[str, Any]:
    account = bank_ledger_account_identity(bot_key, company_name, bank, account_no, account_name)
    all_statement_rows = parse_statement_file(statement_path)
    filtered_rows, filtered_out = filter_statement_reconcile_rows(all_statement_rows, scope, flow_type)
    entries = bank_ledger_entries_from_statement(filtered_rows, account)
    now = int(time.time())
    inserted_count = 0
    duplicate_rows: List[Dict[str, Any]] = []
    with connect(db_path) as conn:
        existing_hashes = existing_bank_ledger_hashes(conn, account["bot_key"], account["account_key"])
        for entry in entries:
            if clean_display(entry.get("source_hash")) in existing_hashes:
                duplicate_rows.append(entry)
        slip_rows = slip_reconcile_rows(conn, chat_id="", scope=scope, bot_key=account["bot_key"], flow_type=flow_type)
        slip_rows = filter_slips_for_bank_ledger_account(slip_rows, account, flow_type)
        if not dry_run:
            ensure_bank_ledger_tables(conn)
            conn.execute(
                """
                INSERT INTO bank_ledger_accounts(bot_key, account_key, company_name, bank, account_no, account_name, active, updated_at)
                VALUES (?,?,?,?,?,?,1,?)
                ON CONFLICT(bot_key, account_key) DO UPDATE SET
                  company_name=excluded.company_name,
                  bank=excluded.bank,
                  account_no=excluded.account_no,
                  account_name=excluded.account_name,
                  active=1,
                  updated_at=excluded.updated_at
                """,
                (account["bot_key"], account["account_key"], account["company_name"], account["bank"], account["account_no"], account["account_name"], now),
            )
            for entry in entries:
                try:
                    conn.execute(
                        """
                        INSERT INTO bank_ledger_entries(
                          entry_id, bot_key, account_key, company_name, bank, account_no, account_name,
                          source_name, source_hash, row_no, date, date_key, time, flow_type, flow_label,
                          description, amount, reference, sender, receiver, imported_at, raw_json
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            entry["entry_id"], entry["bot_key"], entry["account_key"], entry["company_name"], entry["bank"], entry["account_no"], entry["account_name"],
                            entry["source_name"], entry["source_hash"], entry["row_no"], entry["date"], entry["date_key"], entry["time"], entry["flow_type"], entry["flow_label"],
                            entry["description"], entry["amount"], entry["reference"], entry["sender"], entry["receiver"], now, entry["raw_json"],
                        ),
                    )
                    inserted_count += 1
                except sqlite3.IntegrityError:
                    if entry not in duplicate_rows:
                        duplicate_rows.append(entry)
            conn.commit()
    match_result = match_ledger_entries_to_slips(entries, slip_rows, account)
    incoming_amount = sum(float(r.get("amount") or 0) for r in entries)
    inserted_rows = [entry for entry in entries if entry not in duplicate_rows]
    result = {
        "ok": True,
        "mode": "bank_ledger_import",
        "dry_run": bool(dry_run),
        "statement_path": str(statement_path),
        "account": account,
        "scope": {"bot_key": account["bot_key"], "flow_type": normalize_flow_type(flow_type), "flow_label": flow_label(flow_type), "date_scope": clean_display(scope)},
        "incoming": {"count": len(entries), "amount": incoming_amount, "rows": entries[:100]},
        "filtered_out": {"count": len(filtered_out), "amount": sum(float(r.get("amount") or 0) for r in filtered_out)},
        "duplicates": {"count": len(duplicate_rows), "amount": sum(float(r.get("amount") or 0) for r in duplicate_rows), "rows": duplicate_rows[:100]},
        "inserted": {"count": inserted_count if not dry_run else 0, "amount": sum(float(r.get("amount") or 0) for r in inserted_rows) if not dry_run else 0.0},
        "slips": {"count": len(slip_rows), "amount": sum(float(r.get("amount") or 0) for r in slip_rows)},
        "daily": {"ledger": reconcile_daily_summary(entries), "slips": reconcile_daily_summary(slip_rows)},
        "approval_required": False,
    }
    result.update(match_result)
    return result


def preview_bank_ledger_import(db_path: Path, statement_path: Path, **kwargs: Any) -> Dict[str, Any]:
    kwargs["dry_run"] = True
    return import_bank_ledger_statement(db_path, statement_path, **kwargs)


def bank_ledger_snapshot(db_path: Path | None = None, bot_key: str = "", account_key: str = "", account_no: str = "", scope: str = "all", flow_type: str = "all") -> Dict[str, Any]:
    if db_path is None:
        db_path = Path(_value("DB_PATH"))
    account = bank_ledger_account_identity(bot_key or "__all__", "", "", account_no, "")
    with connect(db_path) as conn:
        entries = bank_ledger_query_rows(conn, bot_key=bot_key, account_key=account_key, account_no=account_no, scope=scope, flow_type=flow_type)
        if entries:
            first = entries[0]
            account = bank_ledger_account_identity(first.get("bot_key"), first.get("company_name"), first.get("bank"), first.get("account_no"), first.get("account_name"))
        if not entries and not clean_display(account_no) and not clean_display(account_key):
            slip_rows: List[Dict[str, Any]] = []
        else:
            slip_rows = slip_reconcile_rows(conn, chat_id="", scope=scope, bot_key=bot_key, flow_type=flow_type)
            if account_no or (entries and entries[0].get("account_no")):
                slip_rows = filter_slips_for_bank_ledger_account(slip_rows, account, flow_type)
    match_result = match_ledger_entries_to_slips(entries, slip_rows, account)
    return {
        "ok": True,
        "account": account if entries or account_no else {},
        "entries": {"count": len(entries), "amount": sum(float(r.get("amount") or 0) for r in entries), "rows": entries[:100]},
        "slips": {"count": len(slip_rows), "amount": sum(float(r.get("amount") or 0) for r in slip_rows)},
        "daily": {"ledger": reconcile_daily_summary(entries), "slips": reconcile_daily_summary(slip_rows)},
        "matched": match_result["matched"],
        "unmatched_ledger": match_result["ledger_extra"],
        "unmatched_slips": match_result["slip_extra"],
    }


