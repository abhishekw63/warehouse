"""
Management command: send the daily Issues email.

Builds the Issues email for ONE day's issue lines and sends it to the configured
stakeholders — so a scheduler (Windows Task Scheduler / Render Cron) can mail each
day's issues automatically, no operator click needed.

    py manage.py send_daily_issues                  # today's issues
    py manage.py send_daily_issues --yesterday      # yesterday (for a next-morning run)
    py manage.py send_daily_issues --date 2026-08-20
    py manage.py send_daily_issues --marketplace Flipkart
    py manage.py send_daily_issues --if-not-sent    # skip if that day already went out
    py manage.py send_daily_issues --dry-run        # build + log, DON'T send

Sends the DECIDED issue lines (Excluded + Included) for the chosen day — the same
content as the Issues page "Send" button, scoped to one date. Every send (or
skip) is recorded in ``daily_email_log`` for audit + idempotency. Exits non-zero
on a real send failure so the scheduler can detect a bad run; an empty day is a
clean no-op (exit 0).
"""
from __future__ import annotations

import datetime

from django.core.management.base import BaseCommand

_KIND = 'daily_issues'


class Command(BaseCommand):
    help = "Send the daily Issues email (one day's decided issue lines) to stakeholders."

    def add_arguments(self, parser):
        parser.add_argument('--date', default='',
                            help='Target day YYYY-MM-DD (default: today).')
        parser.add_argument('--yesterday', action='store_true',
                            help='Target yesterday (handy for a next-morning schedule).')
        parser.add_argument('--marketplace', default='',
                            help='Scope to one marketplace (default: all).')
        parser.add_argument('--if-not-sent', action='store_true',
                            help='Skip if a successful send for that day is already logged.')
        parser.add_argument('--dry-run', action='store_true',
                            help='Build + log the attempt but do NOT actually send.')

    def handle(self, *args, **options):
        from online_b2b.services import daily_send_log as log
        from online_b2b.services.issue_email import IssuesEmailReport

        # ── resolve the target day ──
        if options['date']:
            try:
                day = datetime.date.fromisoformat(options['date'])
            except ValueError:
                self.stderr.write(self.style.ERROR(
                    f"Bad --date {options['date']!r} - use YYYY-MM-DD."))
                raise SystemExit(2)
        else:
            day = datetime.date.today()
            if options['yesterday']:
                day -= datetime.timedelta(days=1)
        ds = day.isoformat()
        mp = (options['marketplace'] or '').strip()

        # ── idempotency guard ──
        if options['if_not_sent'] and log.already_sent(_KIND, ds):
            self.stdout.write(f"Already sent the {ds} issues email - skipping.")
            return

        # ── build the report for that single day ──
        rep = IssuesEmailReport({'date_from': ds, 'date_to': ds, 'marketplace': mp,
                                 'resolution': 'all'})
        n = len(rep.rows)
        if not n:
            self.stdout.write(f"No issue lines for {ds}{f' ({mp})' if mp else ''} "
                              f"- nothing to send.")
            log.record(_KIND, ds, 0, ok=True, error='no rows (nothing to send)')
            return

        rcpts = rep.recipients()
        if not rcpts.get('to'):
            msg = 'No recipient configured (set DEFAULT_RECIPIENT / EMAIL_TO).'
            self.stderr.write(self.style.ERROR(msg))
            log.record(_KIND, ds, n, ok=False, error=msg)
            raise SystemExit(1)
        rcpt_str = ', '.join(rcpts.get('to', [])) + (
            ' | cc ' + ', '.join(rcpts.get('cc', [])) if rcpts.get('cc') else '')

        if options['dry_run']:
            self.stdout.write(self.style.WARNING(
                f"[dry-run] {ds}: {n} issue line(s) -> would send to {rcpt_str}"))
            log.record(_KIND, ds, n, ok=True, error='dry-run (not sent)',
                       recipients=rcpt_str)
            return

        # ── send ──
        ok, reason = rep.send()
        log.record(_KIND, ds, n, ok=ok, error='' if ok else (reason or 'send failed'),
                   recipients=rcpt_str)
        if ok:
            self.stdout.write(self.style.SUCCESS(
                f"Sent {ds} issues email: {n} line(s) -> {rcpt_str}"))
        else:
            self.stderr.write(self.style.ERROR(
                f"FAILED to send {ds} issues email: {reason}"))
            raise SystemExit(1)
