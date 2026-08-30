"""
online_b2b.services.tat_store
=============================

**Order-Management TAT (24h SLA).** A PO/SO must be uploaded within **1 working
day** of its PO date — working days **exclude weekends + holidays**. A breach is
**computed** (po_date → upload time ``run_ts``), never stored. The operator's
breach **reason** is stored per order in the web-owned ``order_tat`` table (1:1
with ``order_headers``, FK CASCADE). Online + Offline both covered.

Rules (confirmed with the operator):
  * Clock start  = ``order_headers.po_date`` (the marketplace PO date).
  * Clock end    = ``order_headers.run_ts`` (when we recorded/uploaded it).
  * TAT          = 1 working day. Same-day or next-working-day upload is OK;
                   anything later is a breach.
  * Reason       = a dropdown code + optional note, filled on the TAT page.
"""

from __future__ import annotations

import datetime as _dt

from .order_db import _conn

# Company/national holidays (YYYY-MM-DD) excluded from the working-day count.
# Fill this list as the operator provides the holiday calendar.
HOLIDAYS: set[str] = set()

# Allowed working days within TAT (24 working-hours ≈ next working day).
TAT_DAYS = 1

# Standard breach reasons for the dropdown (+ an optional free note).
REASONS = [
    'PO received late', 'Portal/system down', 'Price clarification',
    'Stock/inventory issue', 'Holiday/weekend', 'Bulk backlog', 'Other',
]

_MYSQL = """
CREATE TABLE IF NOT EXISTS order_tat (
    order_id     BIGINT PRIMARY KEY,
    reason_code  VARCHAR(40),
    note         VARCHAR(500),
    reason_by    VARCHAR(80),
    reason_at    DATETIME,
    CONSTRAINT fk_tat_order FOREIGN KEY (order_id)
        REFERENCES order_headers(order_id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
_SQLITE = """
CREATE TABLE IF NOT EXISTS order_tat (
    order_id INTEGER PRIMARY KEY, reason_code TEXT, note TEXT,
    reason_by TEXT, reason_at TEXT
)
"""


_READY = False        # process-local: the fixed DDL only needs to run ONCE


def ensure_table() -> None:
    global _READY
    if _READY:
        return
    with _conn() as (cur, d):
        cur.execute(_MYSQL if d['kind'] == 'mysql' else _SQLITE)
        cur.connection.commit()
    _READY = True


def _to_date(x):
    if x is None:
        return None
    if isinstance(x, _dt.datetime):
        return x.date()
    if isinstance(x, _dt.date):
        return x
    # ISO first (unambiguous), then Indian DAY-FIRST formats. Trying day-first
    # (never month-first) means an ambiguous '01-07-2026' resolves to 1 Jul,
    # not 7 Jan — and a parseable date is never silently dropped to None.
    s = str(x).strip()[:10]
    for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d.%m.%Y', '%d/%m/%Y'):
        try:
            return _dt.datetime.strptime(s, fmt).date()
        except (ValueError, TypeError):
            continue
    return None


def business_days(start, end) -> int:
    """Working days strictly AFTER ``start`` through ``end``. Only **Sundays** and
    ``HOLIDAYS`` are non-working (Saturday counts as a working day). Same day → 0;
    next working day → 1."""
    start, end = _to_date(start), _to_date(end)
    if not start or not end or end <= start:
        return 0
    n, cur = 0, start
    while cur < end:
        cur += _dt.timedelta(days=1)
        if cur.weekday() != 6 and cur.isoformat() not in HOLIDAYS:  # 6 = Sunday
            n += 1
    return n


def breaches(marketplace='', segment='', q='', status='pending',
             date_from='', date_to='', run='', limit=300) -> dict:
    """Orders uploaded later than TAT (working days). ``status`` =
    'pending' (no reason) / 'resolved' (has reason) / 'all'. ``run`` filters to a
    single upload run_id. Returns rows with ``wd_late`` (working days taken) +
    ``days_over`` (beyond TAT) + the reason, plus counts + ``pct`` resolved."""
    ensure_table()
    out = {'ok': False, 'rows': [], 'counts': {}, 'status': status,
           'date_from': date_from, 'date_to': date_to, 'run': run}
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            # Pre-filter in SQL: business days ≤ calendar days, so a >1-working-day
            # breach needs calendar lateness ≥ TAT_DAYS+1. Exact working-day check
            # is done in Python below.
            where = ["h.po_date IS NOT NULL", "h.run_ts IS NOT NULL",
                     f"DATEDIFF(DATE(h.run_ts), h.po_date) >= {TAT_DAYS + 1}"]
            params: list = []
            if marketplace:
                where.append(f"h.marketplace_label={ph}"); params.append(marketplace)
            if segment:
                where.append(f"h.segment={ph}"); params.append(segment)
            if q:
                where.append(f"(h.po LIKE {ph} OR h.location LIKE {ph})")
                params += [f"%{q}%", f"%{q}%"]
            if date_from:
                where.append(f"DATE(h.run_ts) >= {ph}"); params.append(date_from)
            if date_to:
                where.append(f"DATE(h.run_ts) <= {ph}"); params.append(date_to)
            if run:
                where.append(f"h.run_id={ph}"); params.append(run)
            wsql = " AND ".join(where)
            cols = ['order_id', 'run_ts', 'po_date', 'segment', 'marketplace',
                    'po', 'location', 'qty', 'order_value', 'reason_code', 'note',
                    'reason_by', 'reason_at']
            cur.execute(
                f"SELECT h.order_id, h.run_ts, h.po_date, h.segment, "
                f"h.marketplace_label, h.po, h.location, h.qty, h.order_value, "
                f"t.reason_code, t.note, t.reason_by, t.reason_at "
                f"FROM order_headers h LEFT JOIN order_tat t "
                f"ON t.order_id = h.order_id WHERE {wsql} "
                f"ORDER BY h.po_date ASC, h.order_id DESC", tuple(params))
            pend = res = 0
            kept = []
            for r in cur.fetchall():
                rec = dict(zip(cols, r))
                wd = business_days(rec['po_date'], rec['run_ts'])
                if wd <= TAT_DAYS:        # not a real breach after working-day calc
                    continue
                rec['wd_late'] = wd
                rec['days_over'] = wd - TAT_DAYS
                has = bool(rec.get('reason_code'))
                if has:
                    res += 1
                else:
                    pend += 1
                if status == 'pending' and has:
                    continue
                if status == 'resolved' and not has:
                    continue
                kept.append(rec)
            out['rows'] = kept[:int(limit)]
            out['shown'] = len(out['rows'])
            out['matched'] = len(kept)
            tot = pend + res
            out['counts'] = {'pending': pend, 'resolved': res, 'total': tot}
            # progress = reasons given / total breaches (0/N → 100% when all done)
            out['pct'] = round(res / tot * 100) if tot else 100
        out['ok'] = True
    except Exception as e:  # noqa: BLE001
        out['error'] = f"{type(e).__name__}: {e}"
    return out


def set_reason(order_id, reason_code, note='', by='') -> dict:
    """Set / clear the TAT breach reason for one order (upsert order_tat).
    Empty ``reason_code`` clears it."""
    ensure_table()
    try:
        oid = int(order_id)
    except (TypeError, ValueError):
        return {'ok': False, 'error': 'bad order_id'}
    code = (reason_code or '').strip()
    now = _dt.datetime.now()
    with _conn() as (cur, d):
        ph = d['ph']
        if code:
            row = (oid, code, (note or '')[:500], (by or '')[:80], now)
            if d['kind'] == 'mysql':
                cur.execute(
                    f"INSERT INTO order_tat (order_id, reason_code, note, "
                    f"reason_by, reason_at) VALUES ({ph},{ph},{ph},{ph},{ph}) "
                    f"ON DUPLICATE KEY UPDATE reason_code=VALUES(reason_code), "
                    f"note=VALUES(note), reason_by=VALUES(reason_by), "
                    f"reason_at=VALUES(reason_at)", row)
            else:
                cur.execute(
                    f"INSERT OR REPLACE INTO order_tat (order_id, reason_code, "
                    f"note, reason_by, reason_at) VALUES ({ph},{ph},{ph},{ph},{ph})",
                    row)
        else:
            cur.execute(f"DELETE FROM order_tat WHERE order_id={ph}", (oid,))
        cur.connection.commit()
    return {'ok': True}


def breach_count() -> int:
    """Total TAT breaches (reason given or not) — for the hub KPI. Never raises."""
    try:
        return breaches(status='all', limit=1).get('counts', {}).get('total', 0)
    except Exception:  # noqa: BLE001
        return 0
