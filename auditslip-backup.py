#!/usr/bin/env python3
"""Create a consistent compressed SQLite backup for Auditslip."""
from __future__ import annotations

import gzip
import os
import shutil
import sqlite3
import time
from pathlib import Path

DB_PATH = Path(os.environ.get('AUDITSLIP_DB', '/root/projects/auditslip/data/auditslip.db'))
BACKUP_DIR = Path(os.environ.get('AUDITSLIP_BACKUP_DIR', '/root/projects/auditslip/backups/db'))
RETENTION_DAYS = int(os.environ.get('AUDITSLIP_BACKUP_RETENTION_DAYS', '14'))


def main() -> None:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime('%Y%m%d-%H%M%S')
    tmp_db = BACKUP_DIR / f'auditslip-{stamp}.db.tmp'
    final = BACKUP_DIR / f'auditslip-{stamp}.db.gz'
    with sqlite3.connect(DB_PATH) as src, sqlite3.connect(tmp_db) as dst:
        src.backup(dst)
    with open(tmp_db, 'rb') as src, gzip.open(str(final) + '.tmp', 'wb') as gz:
        shutil.copyfileobj(src, gz)
    Path(str(final) + '.tmp').rename(final)
    tmp_db.unlink(missing_ok=True)
    cutoff = time.time() - RETENTION_DAYS * 86400
    removed = 0
    for p in BACKUP_DIR.glob('auditslip-*.db.gz'):
        if p.stat().st_mtime < cutoff:
            p.unlink()
            removed += 1
    print(f'backup={final} removed_old={removed}')


if __name__ == '__main__':
    main()
