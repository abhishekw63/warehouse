"""
online_b2b.services.daily_checklist
===================================

Daily Activity Checklist — the operator's per-day work tracker so nothing is
missed after interruptions. For each day, each channel (from the
:mod:`marketplaces` registry) has 5 steps that mirror the real workflow:

    Uploaded (web) → Workbook downloaded → Entered in sheet
        → Posted to D365 → Cross check

Behaviour:
  * **"Uploaded (web)" auto-ticks** from that day's ``order_headers`` (for
    web-integrated channels) — with the actual record time.
  * Every manual tick stores a **timestamp + user** (full audit trail).
  * Per-channel progress + overall day %; yesterday's incomplete is surfaced.

API-ready: :func:`get_day` returns a plain JSON-serializable dict.
"""
from __future__ import annotations

import datetime as _dt

from . import marketplaces as reg
from .order_db import _conn

_TABLE = 'daily_checklist'

# (step key, label) — order = column order.
STEPS: list[tuple[str, str]] = [
    ('web', 'Uploaded (web)'),
    ('workbook', 'Workbook downloaded'),
    ('sheet', 'Entered in sheet'),
    ('d365', 'Posted to D365'),
    ('crosscheck', 'Cross check'),
]
_STEP_KEYS = {k for k, _ in STEPS}
_AUTO_STEP = 'web'

_CREATE = f"""
CREATE TABLE IF NOT EXISTS {_TABLE} (
    day         DATE NOT NULL,
    channel     VARCHAR(40) NOT NULL,
    step        VARCHAR(20) NOT NULL,
    checked     TINYINT DEFAULT 0,
    checked_at  DATETIME NULL,
    checked_by  VARCHAR(80),
    PRIMARY KEY (day, channel, step)
)
"""


def ensure_table() -> None:
    with _conn() as (cur, d):
        cur.execute(_CREATE)
        cur.connection.commit()


def _today() -> _dt.date:
    return _dt.date.today()


def _parse_day(day) -> _dt.date:
    if isinstance(day, _dt.date):
        return day
    if not day:
        return _today()
    return _dt.datetime.strptime(str(day)[:10], '%Y-%m-%d').date()


def _hhmm(v) -> str:
    if not v:
        return ''
    if isinstance(v, str):
        v = v[11:16] if len(v) >= 16 else v
        return v
    return v.strftime('%H:%M')


def _recorded_web(day: _dt.date) -> dict:
    """``{channel.key: 'HH:MM'}`` for channels with POs recorded on ``day`` —
    the auto "Uploaded (web)" signal, timed at the earliest record that day."""
    dk = reg.db_key_to_channel()
    out: dict = {}
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(
            f"SELECT marketplace, MIN(run_ts) FROM order_headers "
            f"WHERE DATE(run_ts)={ph} GROUP BY marketplace", (day.isoformat(),))
        for mk, ts in cur.fetchall():
            ch = dk.get(str(mk))
            if ch:
                out[ch] = _hhmm(ts)
    return out


def _stored(day: _dt.date) -> dict:
    """``{(channel, step): {checked, at, by}}`` for a day."""
    out: dict = {}
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(
            f"SELECT channel, step, checked, checked_at, checked_by FROM {_TABLE} "
            f"WHERE day={ph}", (day.isoformat(),))
        for ch, st, ck, at, by in cur.fetchall():
            out[(ch, st)] = {'checked': bool(ck), 'at': _hhmm(at), 'by': by or ''}
    return out


def get_day(day=None) -> dict:
    """Full JSON-safe grid for ``day`` (default today)."""
    ensure_table()
    day = _parse_day(day)
    stored = _stored(day)
    auto = _recorded_web(day)

    seg_out = []
    done_ch = 0
    for seg in reg.grouped():
        chans = []
        for c in seg['channels']:
            key = c['key']
            steps = []
            done_steps = 0
            for sk, slabel in STEPS:
                cell = stored.get((key, sk), {'checked': False, 'at': '', 'by': ''})
                is_auto = False
                if sk == _AUTO_STEP and key in auto:
                    cell = {'checked': True, 'at': auto[key], 'by': 'system'}
                    is_auto = True
                if cell['checked']:
                    done_steps += 1
                steps.append({'key': sk, 'label': slabel, 'checked': cell['checked'],
                              'at': cell['at'], 'by': cell['by'], 'auto': is_auto})
            done = done_steps == len(STEPS)
            if done:
                done_ch += 1
            chans.append({'key': key, 'display': c['display'], 'live': c['live'],
                          'db_key': c['db_key'], 'steps': steps,
                          'done_steps': done_steps, 'total_steps': len(STEPS),
                          'pct': round(done_steps * 100 / len(STEPS)), 'done': done})
        seg_out.append({'segment': seg['segment'], 'channels': chans})

    total_ch = len(reg.channels())
    return {
        'day': day.isoformat(),
        'is_today': day == _today(),
        'segments': seg_out,
        'steps': [{'key': k, 'label': lbl} for k, lbl in STEPS],
        'done_channels': done_ch,
        'total_channels': total_ch,
        'overall_pct': round(done_ch * 100 / total_ch) if total_ch else 0,
        'yesterday': _yesterday_incomplete(day),
    }


def _yesterday_incomplete(day: _dt.date) -> dict:
    """Count of channels that had ANY activity the previous day but weren't fully
    done — surfaced so carried-over work isn't silently forgotten."""
    prev = day - _dt.timedelta(days=1)
    stored = _stored(prev)
    auto = _recorded_web(prev)
    touched, incomplete = set(), []
    for c in reg.channels():
        steps_done = 0
        touched_ch = False
        for sk, _ in STEPS:
            checked = stored.get((c.key, sk), {}).get('checked', False)
            if sk == _AUTO_STEP and c.key in auto:
                checked = True
            if checked:
                steps_done += 1
                touched_ch = True
        if touched_ch:
            touched.add(c.key)
            if steps_done < len(STEPS):
                incomplete.append(c.display)
    return {'date': prev.isoformat(), 'count': len(incomplete),
            'channels': incomplete[:12]}


def toggle(day, channel: str, step: str, checked: bool, user: str = '') -> dict:
    """Set one cell. Records ``checked_at`` (now) + ``checked_by`` on tick."""
    ensure_table()
    day = _parse_day(day)
    if channel not in {c.key for c in reg.channels()} or step not in _STEP_KEYS:
        return {'ok': False, 'error': 'Unknown channel/step.'}
    if step == _AUTO_STEP and channel in _recorded_web(day):
        return {'ok': False,
                'error': 'This channel is already uploaded on the web (auto-ticked).'}
    now = _dt.datetime.now() if checked else None
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(
            f"UPDATE {_TABLE} SET checked={ph}, checked_at={ph}, checked_by={ph} "
            f"WHERE day={ph} AND channel={ph} AND step={ph}",
            (1 if checked else 0, now, user, day.isoformat(), channel, step))
        if cur.rowcount == 0:
            cur.execute(
                f"INSERT INTO {_TABLE} (day, channel, step, checked, checked_at, "
                f"checked_by) VALUES ({ph},{ph},{ph},{ph},{ph},{ph})",
                (day.isoformat(), channel, step, 1 if checked else 0, now, user))
        cur.connection.commit()
    return {'ok': True, 'checked': checked,
            'at': _hhmm(now), 'by': user if checked else ''}
