"""
online_b2b.services.mailer
==========================

Reusable **email skeleton** for the web app — the single place that talks SMTP.

Design (DRY / skeleton-first): every email feature, now and later (issue
reports, daily summaries, exception alerts, …), plugs in by subclassing
:class:`EmailReport` and implementing ``subject()`` + ``html()``. No feature
re-implements SMTP, recipients, or the preview/send contract.

Config is **reused verbatim from the frozen desktop app**
(``online_po_processor.config.email_config.get_email_config``) so web and
desktop send from the same Gmail account / recipient list, and a single
``Calculation Data/email_config.json`` override re-routes both.

Public surface
--------------
``get_config()``                      → merged email config dict
``send_html(subject, html, to, cc)``  → ``(ok, reason)``; never raises
``class EmailReport``                 → base for every email feature
"""
from __future__ import annotations

import logging
import smtplib
from email.message import EmailMessage

log = logging.getLogger(__name__)


def get_config() -> dict:
    """Effective email config (sender, app-password, SMTP, default TO + CC).

    Reuses the desktop app's two-layer config (code defaults + optional
    ``Calculation Data/email_config.json`` override). Returns a plain dict;
    falls back to an empty dict if the desktop module can't be imported."""
    try:
        from online_po_processor.config.email_config import get_email_config
        return get_email_config()
    except Exception:  # noqa: BLE001 — config import must never crash a page
        log.exception('Could not load email config')
        return {}


def _as_list(v) -> list:
    if not v:
        return []
    return [v] if isinstance(v, str) else list(v)


def send_html(subject: str, html: str, to=None, cc=None, config=None):
    """Send one HTML email. Returns ``(ok: bool, reason: str)`` — never raises.

    ``to`` / ``cc`` default to the config's ``DEFAULT_RECIPIENT`` /
    ``CC_RECIPIENTS``. Mirrors the desktop ``EmailSender`` (Gmail STARTTLS:587,
    plain-text fallback + HTML alternative, specific error messages)."""
    config = config or get_config()
    sender = config.get('EMAIL_SENDER')
    pwd = config.get('EMAIL_PASSWORD')
    to_list = _as_list(to) or _as_list(config.get('DEFAULT_RECIPIENT'))
    cc_list = _as_list(cc) if cc is not None else _as_list(config.get('CC_RECIPIENTS'))
    smtp_host = config.get('SMTP_SERVER')

    if not sender:
        return False, 'Email config missing EMAIL_SENDER.'
    if not pwd:
        return False, 'Email config missing EMAIL_PASSWORD.'
    if not to_list:
        return False, 'No recipient — set DEFAULT_RECIPIENT or pass `to`.'
    if not smtp_host:
        return False, 'Email config missing SMTP_SERVER.'

    msg = EmailMessage()
    msg['From'] = sender
    msg['To'] = ', '.join(to_list)
    if cc_list:
        msg['Cc'] = ', '.join(cc_list)
    msg['Subject'] = subject
    msg.set_content('This email contains an HTML report. '
                    'Please view it in an HTML-capable client.')
    msg.add_alternative(html, subtype='html')

    server = None
    try:
        server = smtplib.SMTP(smtp_host, int(config.get('SMTP_PORT', 587)),
                              timeout=30)
        server.starttls()
        server.login(sender, pwd)
        server.send_message(msg, to_addrs=to_list + cc_list)
        server.quit()
        log.info('Email sent: %r → %s (+%d cc)', subject, to_list, len(cc_list))
        return True, ''
    except smtplib.SMTPAuthenticationError as e:
        _close(server)
        return False, (f'Authentication failed — check EMAIL_PASSWORD (Gmail '
                       f'App Password, not the account password). ({e.smtp_code})')
    except smtplib.SMTPException as e:
        _close(server)
        return False, f'SMTP error: {e}'
    except OSError as e:
        _close(server)
        return False, f'Network error — check the connection / SMTP port. ({e})'


def _close(server) -> None:
    if server is None:
        return
    try:
        server.close()
    except (smtplib.SMTPException, OSError):
        pass


class EmailReport:
    """Skeleton base for every web email feature.

    Subclass and implement :meth:`subject` + :meth:`html`; optionally override
    :meth:`to` / :meth:`cc` (return ``None`` to use the config defaults). The
    view layer calls :meth:`preview` (render only — for the on-screen modal),
    then :meth:`send` on confirm."""

    def subject(self) -> str:
        raise NotImplementedError

    def html(self) -> str:
        raise NotImplementedError

    def to(self):           # noqa: D102 — None → config DEFAULT_RECIPIENT
        return None

    def cc(self):           # noqa: D102 — None → config CC_RECIPIENTS
        return None

    def recipients(self) -> dict:
        cfg = get_config()
        return {
            'to': _as_list(self.to()) or _as_list(cfg.get('DEFAULT_RECIPIENT')),
            'cc': (_as_list(self.cc()) if self.cc() is not None
                   else _as_list(cfg.get('CC_RECIPIENTS'))),
        }

    def preview(self) -> dict:
        """Everything the modal needs — no network call."""
        r = self.recipients()
        return {'subject': self.subject(), 'html': self.html(),
                'to': r['to'], 'cc': r['cc']}

    def send(self):
        """Deliver. Returns ``(ok, reason)``."""
        return send_html(self.subject(), self.html(), to=self.to(), cc=self.cc())
