"""Retry any failed/pending auto Issues emails (the self-healing sweep).

Run manually, or from a FREE scheduler (a GitHub-Actions `schedule:` hitting this,
or the /healthz keep-alive worker later) as a safety net on top of the per-run
send that fires at Lock & Record:

    python manage.py flush_issue_emails [--limit 25]
"""
from django.core.management.base import BaseCommand

from online_b2b.services import auto_issue_email


class Command(BaseCommand):
    help = "Retry failed/pending auto Issues emails (self-healing sweep)."

    def add_arguments(self, parser):
        parser.add_argument('--limit', type=int, default=25,
                            help='Max runs to retry this pass (default 25).')

    def handle(self, *args, **opts):
        res = auto_issue_email.flush_pending(limit=opts['limit'])
        self.stdout.write(self.style.SUCCESS(
            f"Issue-email sweep: tried {res['tried']}, sent {res['sent']}."))
