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

    def add_arguments(self, parser):
        # --if-due makes the command idempotent-per-day: it only backs up when the
        # last successful backup is older than --min-hours. This is what lets a
        # logon/daily scheduler fire it every time the laptop is opened without
        # re-running a full backup more than once a day (and it auto-catches-up a
        # day missed while the laptop was off).
        parser.add_argument('--if-due', action='store_true',
                            help='Skip if a successful backup ran within --min-hours.')
        parser.add_argument('--min-hours', type=float, default=20.0,
                            help='Freshness window for --if-due (default 20h).')

    def handle(self, *args, **options):
        from online_b2b.services import db_target
        if options.get('if_due'):
            lb = db_target.last_backup()
            if lb and lb.get('at'):
                try:
                    last = datetime.datetime.strptime(str(lb['at'])[:19], '%Y-%m-%d %H:%M:%S')
                    age_h = (datetime.datetime.now() - last).total_seconds() / 3600.0
                    if age_h < options['min_hours']:
                        self.stdout.write(
                            f"Not due — last backup {age_h:.1f}h ago "
                            f"(< {options['min_hours']:.0f}h). Skipping.")
                        return
                except (ValueError, TypeError):
                    pass   # unparseable marker → treat as due, run the backup
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
