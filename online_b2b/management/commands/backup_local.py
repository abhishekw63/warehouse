"""
Management command: backup TiDB -> local MySQL.

Copies ALL data from the TiDB profile into the local MySQL profile (schema + rows
+ views), so local mirrors the server. Reverse of ``db_push_to_tidb.py``, and the
same operation as the Setup page's "Backup TiDB -> local" card — exposed here so a
scheduler (Windows Task Scheduler / cron) can run it daily.

    py manage.py backup_local

Destructive on LOCAL only (local is overwritten to match TiDB); TiDB is read-only.
Exits non-zero on failure so a scheduler can detect a bad run.
"""
from __future__ import annotations

import datetime

from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = "Backup: copy ALL data from TiDB into the local MySQL profile (daily mirror)."

    def handle(self, *args, **options):
        from online_b2b.services import db_target
        ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.stdout.write(f"[{ts}] TiDB -> local backup starting...")
        res = db_target.backup_tidb_to_local()
        end = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        if res.get('ok'):
            self.stdout.write(self.style.SUCCESS(
                f"[{end}] OK - {res.get('n_tables')} table(s), "
                f"{res.get('total_rows', 0):,} rows, {res.get('views')} view(s) "
                f"in {res.get('elapsed')}s. {res.get('source')} -> {res.get('target')}"))
        else:
            extra = ''
            if res.get('copied_tables') is not None:
                extra = (f" (stopped after {res.get('copied_tables')} table(s), "
                         f"{res.get('rows_so_far', 0):,} rows)")
            self.stderr.write(self.style.ERROR(
                f"[{end}] FAILED - {res.get('error')}{extra}"))
            raise SystemExit(1)
