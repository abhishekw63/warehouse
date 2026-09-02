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
    # `to=None` means "use the configured default". An EXPLICIT empty list means
    # "send to nobody" and must NOT fall back — the old `or` treated [] and None
    # alike, so a caller that had filtered its recipients down to none silently
    # mailed the default recipient instead. Mirrors how `cc` already behaves.
    to_list = (_as_list(config.get('DEFAULT_RECIPIENT')) if to is None
               else _as_list(to))
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

    # ── HTTPS first, when a provider is configured ────────────────────────
    # The deployed host's outbound SMTP is blocked ([Errno 101] Network is
    # unreachable), so on the server this is the ONLY path that can work. It is
    # opt-in: with EMAIL_HTTP_PROVIDER unset we go straight to SMTP and nothing
    # changes for local use. If the HTTP send fails we still try SMTP, so a
    # misconfigured key can't take email away from a machine where SMTP works.
    http_err = ''
    try:
        from . import mail_http
        if mail_http.configured():
            ok, reason = mail_http.send(subject, html, sender, to_list, cc_list)
            if ok:
                return True, ''
            http_err = reason
            log.warning('HTTP email failed (%s) — falling back to SMTP', reason)
    except Exception as e:  # noqa: BLE001 — transport choice must never raise
        http_err = f'{type(e).__name__}: {e}'

    def _fail(msg):
        # Surface BOTH failures, else the HTTP problem is invisible behind the
        # SMTP one and whoever debugs it chases the wrong transport.
        return False, (f'{msg} (HTTP transport also failed: {http_err})'
                       if http_err else msg)

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
        return _fail(f'Authentication failed — check EMAIL_PASSWORD (Gmail '
                     f'App Password, not the account password). ({e.smtp_code})')
    except smtplib.SMTPException as e:
        _close(server)
        return _fail(f'SMTP error: {e}')
    except OSError as e:
        _close(server)
        return _fail(f'Network error — check the connection / SMTP port. ({e})')


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
