"""
online_b2b.services.mail_http
=============================

Send email over **HTTPS** instead of SMTP.

Why this exists: the deployed host cannot open a TCP connection to
``smtp.gmail.com:587`` — outbound SMTP egress is blocked, so every send from the
server fails with ``[Errno 101] Network is unreachable``. Proven by correlating
``issue_email_log`` against ``audit_log.host``: every successful send in the
app's history came from a developer laptop, every server-side attempt failed.
Retrying cannot help; it is the same blocked path.

HTTPS egress plainly works (the app reaches its database and serves traffic), so
the fix is to change the TRANSPORT, not the message. Recipients, subject and HTML
are untouched — :mod:`online_b2b.services.mailer` just hands them here first.

Configuration (environment, e.g. the host's dashboard)::

    EMAIL_HTTP_PROVIDER = resend | brevo | sendgrid | mailgun
    EMAIL_HTTP_KEY      = <the provider's API key>
    EMAIL_HTTP_DOMAIN   = <mailgun only: the sending domain>

Unset ``EMAIL_HTTP_PROVIDER`` and nothing changes — SMTP stays the only path, so
this is inert until someone opts in.

Deliberately stdlib-only (``urllib``): adding a dependency for four small POSTs
would be a poor trade, and it keeps this importable in any environment.
"""
from __future__ import annotations

import json
import logging
import os
import urllib.error
import urllib.request

log = logging.getLogger(__name__)

_TIMEOUT = 20
_ENDPOINTS = {
    'resend':   'https://api.resend.com/emails',
    'brevo':    'https://api.brevo.com/v3/smtp/email',
    'sendgrid': 'https://api.sendgrid.com/v3/mail/send',
    # mailgun's URL carries the domain, so it is built in _payload()
}


def provider() -> str:
    """The configured provider name, lowercased. '' when HTTP send is off."""
    name = (os.environ.get('EMAIL_HTTP_PROVIDER') or '').strip().lower()
    return name if name in ('resend', 'brevo', 'sendgrid', 'mailgun') else ''


def configured() -> bool:
    """True when a provider AND a key are set — i.e. HTTP send is usable."""
    return bool(provider() and (os.environ.get('EMAIL_HTTP_KEY') or '').strip())


def _payload(prov, key, subject, html, sender, to, cc):
    """-> (url, body_bytes, headers) for the chosen provider."""
    if prov == 'resend':
        body = {'from': sender, 'to': to, 'subject': subject, 'html': html}
        if cc:
            body['cc'] = cc
        return (_ENDPOINTS[prov], json.dumps(body).encode(),
                {'Authorization': f'Bearer {key}', 'Content-Type': 'application/json'})

    if prov == 'brevo':
        body = {'sender': {'email': sender}, 'subject': subject,
                'htmlContent': html, 'to': [{'email': a} for a in to]}
        if cc:
            body['cc'] = [{'email': a} for a in cc]
        return (_ENDPOINTS[prov], json.dumps(body).encode(),
                {'api-key': key, 'Content-Type': 'application/json',
                 'accept': 'application/json'})

    if prov == 'sendgrid':
        person = {'to': [{'email': a} for a in to]}
        if cc:
            person['cc'] = [{'email': a} for a in cc]
        body = {'personalizations': [person], 'from': {'email': sender},
                'subject': subject,
                'content': [{'type': 'text/html', 'value': html}]}
        return (_ENDPOINTS[prov], json.dumps(body).encode(),
                {'Authorization': f'Bearer {key}', 'Content-Type': 'application/json'})

    # mailgun — form-encoded, HTTP basic auth ('api' : key)
    import base64
    import urllib.parse
    domain = (os.environ.get('EMAIL_HTTP_DOMAIN') or '').strip()
    fields = [('from', sender), ('subject', subject), ('html', html)]
    fields += [('to', a) for a in to] + [('cc', a) for a in cc]
    tok = base64.b64encode(f'api:{key}'.encode()).decode()
    return (f'https://api.mailgun.net/v3/{domain}/messages',
            urllib.parse.urlencode(fields).encode(),
            {'Authorization': f'Basic {tok}',
             'Content-Type': 'application/x-www-form-urlencoded'})


def send(subject: str, html: str, sender: str, to: list, cc: list = None):
    """POST the mail to the configured provider. -> ``(ok, reason)``; never raises.

    Mirrors ``mailer.send_html``'s contract exactly so it can stand in for it.
    """
    cc = list(cc or [])
    prov, key = provider(), (os.environ.get('EMAIL_HTTP_KEY') or '').strip()
    if not prov or not key:
        return False, 'HTTP email not configured (EMAIL_HTTP_PROVIDER / EMAIL_HTTP_KEY).'
    if prov == 'mailgun' and not (os.environ.get('EMAIL_HTTP_DOMAIN') or '').strip():
        return False, 'Mailgun needs EMAIL_HTTP_DOMAIN.'
    if not sender:
        return False, 'Email config missing EMAIL_SENDER.'
    if not to:
        return False, 'No recipient.'

    url, body, headers = _payload(prov, key, subject, html, sender, to, cc)
    req = urllib.request.Request(url, data=body, headers=headers, method='POST')
    try:
        with urllib.request.urlopen(req, timeout=_TIMEOUT) as r:
            if 200 <= r.status < 300:
                log.info('Email sent via %s: %r → %s (+%d cc)',
                         prov, subject, to, len(cc))
                return True, ''
            return False, f'{prov} returned HTTP {r.status}'
    except urllib.error.HTTPError as e:
        # The provider's own message is the useful part (bad key, unverified
        # sender domain, quota) — surface it instead of a bare status code.
        try:
            detail = e.read().decode('utf-8', 'replace')[:300]
        except Exception:  # noqa: BLE001
            detail = ''
        return False, f'{prov} HTTP {e.code}: {detail or e.reason}'
    except urllib.error.URLError as e:
        return False, f'{prov} unreachable: {e.reason}'
    except Exception as e:  # noqa: BLE001 — a mail transport must never raise
        return False, f'{prov} error: {type(e).__name__}: {e}'
