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
_NOPO = 'nopo'   # per-channel "no PO received today" flag (NOT a work step)
_HOLD = 'hold'   # per-channel "on hold" flag (e.g. a CP issue) — parks the
                 # channel so it's not chased/counted as pending until un-held.

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


_CREATE_HOLD_LOG = f"""
CREATE TABLE IF NOT EXISTS {_TABLE}_hold_log (
    id       BIGINT AUTO_INCREMENT PRIMARY KEY,
    day      DATE NOT NULL,
    channel  VARCHAR(40) NOT NULL,
    action   VARCHAR(10) NOT NULL,      -- 'hold' | 'unhold'
    at       DATETIME NOT NULL,
    by_user  VARCHAR(80)
)
"""


def ensure_table() -> None:
    with _conn() as (cur, d):
        cur.execute(_CREATE)
        cur.execute(_CREATE_HOLD_LOG)
        cur.connection.commit()


def hold_history(day, channel: str) -> list[dict]:
    """Append-only hold/un-hold audit for a channel on a day — every hold and
    un-hold with its timestamp + user, so hold DURATION is always recoverable."""
    ensure_table()
    day = _parse_day(day)
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(
            f"SELECT action, at, by_user FROM {_TABLE}_hold_log "
            f"WHERE day={ph} AND channel={ph} ORDER BY at",
            (day.isoformat(), channel))
        return [{'action': a, 'at': (at.strftime('%Y-%m-%d %H:%M:%S') if at else ''),
                 'by': by or ''} for a, at, by in cur.fetchall()]


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


def _leaf_dict(c: dict, stored: dict, auto: dict) -> dict:
    """Build the JSON-safe grid dict for ONE trackable leaf channel (its own 5
    steps, DB rows, state). ``c`` is a registry channel dict."""
    key = c['key']
    chan_steps = c['steps'] or STEPS   # per-channel custom flow, else default
    nsteps = len(chan_steps)
    nopo = stored.get((key, _NOPO), {})
    no_po = bool(nopo.get('checked'))
    held = stored.get((key, _HOLD), {})
    on_hold = bool(held.get('checked'))
    steps = []
    done_steps = 0
    for sk, slabel in chan_steps:
        cell = stored.get((key, sk), {'checked': False, 'at': '', 'by': ''})
        is_auto = False
        if sk == _AUTO_STEP and key in auto:
            cell = {'checked': True, 'at': auto[key], 'by': 'system'}
            is_auto = True
        if cell['checked']:
            done_steps += 1
        steps.append({'key': sk, 'label': slabel, 'checked': cell['checked'],
                      'at': cell['at'], 'by': cell['by'], 'auto': is_auto})
    work_done = done_steps == nsteps
    if on_hold:                       # hold wins — parked, not chased
        state = 'hold'
    elif no_po:
        state = 'nopo'
    elif work_done:
        state = 'done'
    else:
        state = 'partial' if done_steps else 'todo'
    return {'key': key, 'display': c['display'], 'live': c['live'],
            'db_key': c['db_key'], 'steps': steps, 'is_parent': False,
            'children': [], 'custom': bool(c['steps']),
            'done_steps': done_steps, 'total_steps': nsteps,
            'pct': round(done_steps * 100 / nsteps) if nsteps else 0,
            'done': work_done, 'no_po': no_po,
            'no_po_at': nopo.get('at', ''),
            'on_hold': on_hold, 'hold_at': held.get('at', ''),
            'hold_by': held.get('by', ''), 'state': state}


def _parent_dict(c: dict, kids: list[dict]) -> dict:
    """Build the grid dict for a PARENT/container channel (e.g. MT Select). It has
    NO work steps of its own — its progress is a ROLLUP of its children: done when
    every child is done/handled; pct = mean of child pct; shows X/N done."""
    n = len(kids)
    done_kids = sum(1 for k in kids if k['state'] == 'done')
    handled_kids = sum(1 for k in kids if k['state'] in ('done', 'nopo', 'hold'))
    pct = round(sum(k['pct'] for k in kids) / n) if n else 0
    all_done = n > 0 and handled_kids == n
    return {'key': c['key'], 'display': c['display'], 'live': c['live'],
            'db_key': c['db_key'], 'steps': [], 'is_parent': True,
            'children': kids, 'custom': False,
            'done_steps': done_kids, 'total_steps': n,
            'child_done': done_kids, 'child_total': n,
            'child_handled': handled_kids,
            'pct': pct, 'done': all_done, 'no_po': False, 'no_po_at': '',
            'on_hold': False, 'hold_at': '', 'hold_by': '', 'state':
            'done' if all_done else ('partial' if handled_kids else 'todo')}


def get_day(day=None) -> dict:
    """Full JSON-safe grid for ``day`` (default today).

    Channels nest: a PARENT row (e.g. MT Select) is a container whose progress
    rolls up its children; each CHILD is a normal trackable leaf. Overall counts
    run over LEAVES only (children + standalone channels) — the parent container
    never counts as a separate unit."""
    ensure_table()
    day = _parse_day(day)
    stored = _stored(day)
    auto = _recorded_web(day)

    seg_out = []
    done_ch = nopo_ch = held_ch = 0
    for seg in reg.grouped():
        chans = []
        for c in seg['channels']:
            if c.get('is_parent'):
                kids = [_leaf_dict(kc, stored, auto) for kc in c['children']]
                chans.append(_parent_dict(c, kids))
                leaves = kids
            else:
                leaf = _leaf_dict(c, stored, auto)
                chans.append(leaf)
                leaves = [leaf]
            # counts run over the leaves (children / standalone), never the parent
            for lf in leaves:
                if lf['state'] == 'hold':
                    held_ch += 1
                elif lf['state'] == 'nopo':
                    nopo_ch += 1
                elif lf['state'] == 'done':
                    done_ch += 1
        seg_out.append({'segment': seg['segment'], 'channels': chans})

    total_ch = len(reg.leaves())
    handled = done_ch + nopo_ch + held_ch
    return {
        'day': day.isoformat(),
        'is_today': day == _today(),
        'segments': seg_out,
        'steps': [{'key': k, 'label': lbl} for k, lbl in STEPS],
        'done_channels': done_ch,
        'nopo_channels': nopo_ch,
        'held_channels': held_ch,
        'handled_channels': handled,
        'pending_channels': total_ch - handled,
        'total_channels': total_ch,
        'overall_pct': round(handled * 100 / total_ch) if total_ch else 0,
        'yesterday': _yesterday_incomplete(day),
    }


def _yesterday_incomplete(day: _dt.date) -> dict:
    """Count of channels that had ANY activity the previous day but weren't fully
    done — surfaced so carried-over work isn't silently forgotten."""
    prev = day - _dt.timedelta(days=1)
    stored = _stored(prev)
    auto = _recorded_web(prev)
    touched, incomplete = set(), []
    for c in reg.leaves():                    # leaves only — parents are rollups
        if stored.get((c.key, _NOPO), {}).get('checked'):
            continue  # marked "no PO today" = handled, not carried over
        chan_steps = c.steps or STEPS
        steps_done = 0
        touched_ch = False
        for sk, _ in chan_steps:
            checked = stored.get((c.key, sk), {}).get('checked', False)
            if sk == _AUTO_STEP and c.key in auto:
                checked = True
            if checked:
                steps_done += 1
                touched_ch = True
        if touched_ch:
            touched.add(c.key)
            if steps_done < len(chan_steps):
                incomplete.append(c.display)
    return {'date': prev.isoformat(), 'count': len(incomplete),
            'channels': incomplete[:12]}


def toggle(day, channel: str, step: str, checked: bool, user: str = '') -> dict:
    """Set one cell. Records ``checked_at`` (now) + ``checked_by`` on tick."""
    ensure_table()
    day = _parse_day(day)
    ch = reg.get(channel)
    if not ch:
        return {'ok': False, 'error': 'Unknown channel.'}
    valid = {sk for sk, _ in (ch.steps or STEPS)} | {_NOPO, _HOLD}
    if step not in valid:
        return {'ok': False, 'error': 'Unknown step for this channel.'}
    if step == _AUTO_STEP and channel in _recorded_web(day):
        return {'ok': False,
                'error': 'This channel is already uploaded on the web (auto-ticked).'}
    if step == _NOPO and checked and channel in _recorded_web(day):
        return {'ok': False,
                'error': "This channel has POs recorded today — can't mark 'No PO today'."}
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
        # Hold audit — log EVERY hold + un-hold with its OWN timestamp (the main
        # row's checked_at is cleared on un-hold, so duration would otherwise be
        # lost). Append-only.
        if step == _HOLD:
            cur.execute(
                f"INSERT INTO {_TABLE}_hold_log (day, channel, action, at, by_user) "
                f"VALUES ({ph},{ph},{ph},{ph},{ph})",
                (day.isoformat(), channel, 'hold' if checked else 'unhold',
                 _dt.datetime.now(), user))
        cur.connection.commit()
    return {'ok': True, 'checked': checked,
            'at': _hhmm(now), 'by': user if checked else ''}
