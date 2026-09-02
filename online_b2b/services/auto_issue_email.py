"""
online_b2b.services.auto_issue_email
====================================

Auto **Issues email** — sent per run on Lock&Record (excluded + included lines),
with a self-healing retry sweep for anything that missed.

This module owns ONLY the *plumbing*: when to send, idempotency, retries, and the
log. Composition/content is delegated to :class:`IssuesEmailReport`, so the email
can be redesigned later (or swapped for a webhook) without touching any of this.

It NEVER raises to its caller — a failed email must never break the money-path
lock. Toggle off with ``AUTO_ISSUE_EMAIL=0``. See [[issues-email-exclude-only]].
"""
from __future__ import annotations

import os
import threading

from . import issue_email_log as _log


def enabled() -> bool:
    """Kill-switch: ``AUTO_ISSUE_EMAIL=0`` turns per-run auto-mail off."""
    return os.environ.get('AUTO_ISSUE_EMAIL', '1') not in ('0', 'false', 'False', '')


def send_for_run(run_id, marketplace: str = '', *, force: bool = False,
                 retries: int = 1) -> dict:
    """Send the Issues email for ONE run (excluded + included lines). Idempotent —
    a run already ``sent`` is skipped unless ``force``. Returns a small result for
    the UI toast: ``{status, detail, to, n_excluded, n_included}``. Never raises."""
    try:
        run_id = int(run_id)
    except (TypeError, ValueError):
        return {'status': 'skipped', 'detail': 'no run id'}

    if not enabled():
        return {'status': 'skipped', 'detail': 'auto-email off'}

    if not force and _log.status_of(run_id) == 'sent':
        return {'status': 'sent', 'detail': 'already sent',
                'n_excluded': 0, 'n_included': 0}

    try:
        from .issue_email import IssuesEmailReport
        rep = IssuesEmailReport({'run_id': str(run_id)})
    except Exception as e:  # noqa: BLE001
        _log.record(run_id, 'failed', marketplace=marketplace,
                    error=f'compose error: {type(e).__name__}: {e}')
        return {'status': 'failed', 'detail': f'compose error: {e}'}

    # Never mail a false 'no issues' on a failed data read.
    if rep.fetch_error:
        _log.record(run_id, 'failed', marketplace=marketplace,
                    error=f'issue-data read failed: {rep.fetch_error}')
        return {'status': 'failed', 'detail': str(rep.fetch_error)}

    n_excl, n_incl = len(rep.excluded), len(rep.included)
    if n_excl == 0 and n_incl == 0:
        _log.record(run_id, 'skipped', marketplace=marketplace,
                    error='no excluded/included issue lines')
        return {'status': 'skipped', 'detail': 'no issue lines',
                'n_excluded': 0, 'n_included': 0}

    r = rep.recipients()
    to_list = r.get('to') or []
    recips = ', '.join(to_list + (r.get('cc') or []))
    if not to_list:
        _log.record(run_id, 'skipped', marketplace=marketplace,
                    n_excluded=n_excl, n_included=n_incl,
                    error='no recipient configured')
        return {'status': 'skipped', 'detail': 'no recipient configured',
                'n_excluded': n_excl, 'n_included': n_incl}

    ok, reason = False, ''
    for _ in range(max(1, retries + 1)):          # initial try + `retries` retries
        try:
            ok, reason = rep.send()
        except Exception as e:  # noqa: BLE001
            ok, reason = False, f'{type(e).__name__}: {e}'
        if ok:
            break

    if ok:
        _log.record(run_id, 'sent', marketplace=marketplace, n_excluded=n_excl,
                    n_included=n_incl, recipients=recips, subject=rep.subject())
        return {'status': 'sent', 'detail': 'sent', 'to': to_list,
                'n_excluded': n_excl, 'n_included': n_incl}

    _log.record(run_id, 'failed', marketplace=marketplace, n_excluded=n_excl,
                n_included=n_incl, recipients=recips, subject=rep.subject(),
                error=(reason or 'send failed')[:500])
    return {'status': 'failed', 'detail': (reason or 'send failed'),
            'to': to_list, 'n_excluded': n_excl, 'n_included': n_incl}


def flush_pending(limit: int = 25) -> dict:
    """Re-attempt every run with a failed/pending auto-email (the self-healing
    sweep). Returns ``{tried, sent}``. Best-effort; never raises."""
    tried = sent = 0
    for rid in _log.pending_run_ids(limit):
        # Re-check RIGHT NOW instead of trusting the list we just read. The list
        # is a snapshot; a send can land between reading it and getting here (it
        # did — a sweep re-sent a run that had just succeeded from another host,
        # and its 'failed' write clobbered the 'sent' row). force=False makes
        # send_for_run itself skip an already-sent run, so the mail can never go
        # out twice and the newer state is never overwritten.
        tried += 1
        if send_for_run(rid, force=False).get('status') == 'sent':
            sent += 1
    return {'tried': tried, 'sent': sent}


def flush_pending_async(limit: int = 25) -> None:
    """Fire the sweep in a daemon thread so it adds no latency to a lock."""
    if not enabled():
        return
    try:
        threading.Thread(target=flush_pending, kwargs={'limit': limit},
                         daemon=True).start()
    except Exception:  # noqa: BLE001
        pass
