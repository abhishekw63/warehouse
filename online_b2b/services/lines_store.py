"""
online_b2b.services.lines_store
===============================

Web-owned **full line-item audit** (``order_lines``).

The engine (``online_po_processor``) is the frozen backup and is NOT modified —
so this module, on the Django side, owns the ``order_lines`` table end to end:
it ensures the table exists, builds line rows from an engine ``ProcessingResult``
(reading SORow attributes only — never mutating the engine), and inserts them
into the same MySQL ``renee_orders`` via the shared raw connection.

``order_lines`` is the single line table (2-table model with ``order_headers``):
it carries the full vendor-vs-our comparison columns + ``status``, so the
"affected / mismatch" view is just ``status IN ('MISMATCH','NOT_IN_MASTER')``.
"""

from __future__ import annotations

import datetime as _dt

from .order_db import _conn, _conn_tx


def _utc_now() -> _dt.datetime:
    """Naive UTC 'now' — stamped on run_ts/created_at so the store is uniformly UTC
    on EVERY host (Render already runs UTC; local dev used to write IST, which made
    the same row read 5.5h apart). Display converts UTC→IST (see order_db._to_ist)."""
    return _dt.datetime.now(_dt.timezone.utc).replace(tzinfo=None)

# 2-table split (scalable model): order_lines = immutable order FACTS;
# order_line_validation = the computed validation + operator-decision layer
# (1:1 by line_id, only for validated lines). Reads go through the join VIEW
# ``order_lines_full``. ``received_ean`` (wrong EAN as received) lives in the
# validation table — order_lines never holds wrong data.
_FACT_COLS = [
    'run_id', 'run_ts', 'marketplace', 'po', 'location', 'item_no', 'ean',
    'description', 'qty', 'order_type', 'gst_code', 'unit_price', 'output_file',
]
_VAL_COLS = [
    'our_mrp', 'vendor_mrp', 'our_landing', 'vendor_landing', 'our_cp',
    'vendor_cp', 'diff', 'margin_pct', 'status', 'exception_label',
    'received_ean', 'action', 'override_cp', 'remark', 'decided_at',
]
# Back-compat: the full logical column set of a line (facts + validation).
COLS = _FACT_COLS + _VAL_COLS

_MYSQL_FACTS = """
CREATE TABLE IF NOT EXISTS order_lines (
    line_id      BIGINT AUTO_INCREMENT PRIMARY KEY,
    run_id       BIGINT,
    run_ts       DATETIME,
    marketplace  VARCHAR(50),
    po           VARCHAR(100),
    location     VARCHAR(500),
    item_no      VARCHAR(50),
    ean          VARCHAR(20),
    description  VARCHAR(255),
    qty          INT,
    order_type   VARCHAR(10),
    gst_code     VARCHAR(20),
    unit_price   DECIMAL(14,2),
    output_file  VARCHAR(500),
    created_at   DATETIME DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_lines_run (run_id),
    INDEX idx_lines_mp_po (marketplace, po),
    INDEX idx_lines_item (item_no),
    INDEX idx_lines_po_run (po, run_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
_SQLITE_FACTS = """
CREATE TABLE IF NOT EXISTS order_lines (
    line_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id       INTEGER, run_ts TEXT, marketplace TEXT, po TEXT, location TEXT,
    item_no      TEXT, ean TEXT, description TEXT, qty INTEGER, order_type TEXT,
    gst_code     TEXT, unit_price REAL, output_file TEXT,
    created_at   TEXT DEFAULT CURRENT_TIMESTAMP
)
"""
_MYSQL_VAL = """
CREATE TABLE IF NOT EXISTS order_line_validation (
    line_id         BIGINT PRIMARY KEY,
    our_mrp         DECIMAL(14,2), vendor_mrp     DECIMAL(14,2),
    our_landing     DECIMAL(14,2), vendor_landing DECIMAL(14,2),
    our_cp          DECIMAL(14,2), vendor_cp      DECIMAL(14,2),
    diff            DECIMAL(14,2), margin_pct     DECIMAL(6,2),
    status          VARCHAR(20),   exception_label VARCHAR(50),
    received_ean    VARCHAR(20),
    action          VARCHAR(20),   override_cp    DECIMAL(14,2),
    remark          VARCHAR(255),  decided_at     DATETIME,
    INDEX idx_val_status (status), INDEX idx_val_recv (received_ean),
    CONSTRAINT fk_val_line FOREIGN KEY (line_id)
        REFERENCES order_lines(line_id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
_SQLITE_VAL = """
CREATE TABLE IF NOT EXISTS order_line_validation (
    line_id         INTEGER PRIMARY KEY,
    our_mrp REAL, vendor_mrp REAL, our_landing REAL, vendor_landing REAL,
    our_cp REAL, vendor_cp REAL, diff REAL, margin_pct REAL,
    status TEXT, exception_label TEXT, received_ean TEXT,
    action TEXT, override_cp REAL, remark TEXT, decided_at TEXT,
    FOREIGN KEY (line_id) REFERENCES order_lines(line_id) ON DELETE CASCADE
)
"""
# The join view reads go through — exposes the original full column set, so
# read queries only swap the table name. COALESCE(status,'OK') so unvalidated
# (offline/D365) lines, which have no validation row, still read as OK.
_VIEW_SELECT = """
  l.line_id, l.run_id, l.run_ts, l.marketplace, l.po, l.location,
  l.item_no, l.ean, l.description, l.qty, l.order_type, l.gst_code,
  l.unit_price, l.output_file, l.created_at,
  v.our_mrp, v.vendor_mrp, v.our_landing, v.vendor_landing,
  v.our_cp, v.vendor_cp, v.diff, v.margin_pct,
  COALESCE(v.status,'OK') AS status, v.exception_label, v.received_ean,
  v.action, v.override_cp, v.remark, v.decided_at
"""


def _f(x) -> float | None:
    """2-dp float; None for blanks / NaN / non-numeric."""
    if x is None:
        return None
    try:
        v = float(x)
        return None if v != v else round(v, 2)
    except (TypeError, ValueError):
        return None


_READY = False        # process-local: the fixed DDL below only needs to run ONCE


def ensure_table() -> None:
    """Create the facts table + validation table + join view (idempotent).
    Web owns all three; the engine schema is untouched. Guarded by ``_READY`` so
    the ~5 DDL round-trips (2 CREATE TABLE + CREATE OR REPLACE VIEW + 2
    information_schema index probes) run once per process, not on every line
    write. Render restarts (deploy/spin-down) re-apply the DDL, so a changed view
    / index definition still lands on the next boot."""
    global _READY
    if _READY:
        return
    with _conn() as (cur, d):
        mysql = d['kind'] == 'mysql'
        cur.execute(_MYSQL_FACTS if mysql else _SQLITE_FACTS)
        cur.execute(_MYSQL_VAL if mysql else _SQLITE_VAL)
        view_sql = (f"SELECT {_VIEW_SELECT} FROM order_lines l "
                    "LEFT JOIN order_line_validation v ON v.line_id = l.line_id")
        if mysql:
            cur.execute(f"CREATE OR REPLACE VIEW order_lines_full AS {view_sql}")
            # Composite (po, run_id) index — makes the "latest run per PO" joins
            # (Fulfilment Risk analytics + Availability Checker) index-seek instead
            # of full-scan order_lines. Added post-hoc for pre-existing tables.
            cur.execute(
                "SELECT COUNT(*) FROM information_schema.statistics "
                "WHERE table_schema=DATABASE() AND table_name='order_lines' "
                "AND index_name='idx_lines_po_run'")
            if not cur.fetchone()[0]:
                cur.execute("ALTER TABLE order_lines ADD INDEX idx_lines_po_run (po, run_id)")
            # created_at index on order_headers — every analytics / dashboard /
            # facility query filters order_headers by created_at on each load;
            # without an index they full-scan. Additive + best-effort (skip if the
            # table isn't created yet in a brand-new DB, or we lack ALTER rights).
            try:
                cur.execute(
                    "SELECT COUNT(*) FROM information_schema.statistics "
                    "WHERE table_schema=DATABASE() AND table_name='order_headers' "
                    "AND index_name='idx_oh_created_at'")
                if not cur.fetchone()[0]:
                    cur.execute("ALTER TABLE order_headers "
                                "ADD INDEX idx_oh_created_at (created_at)")
            except Exception:  # noqa: BLE001 — order_headers absent / no perms
                pass
        else:
            cur.execute("DROP VIEW IF EXISTS order_lines_full")
            cur.execute(f"CREATE VIEW order_lines_full AS {view_sql}")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_lines_po_run ON order_lines (po, run_id)")
            try:
                cur.execute("CREATE INDEX IF NOT EXISTS idx_oh_created_at "
                            "ON order_headers (created_at)")
            except Exception:  # noqa: BLE001
                pass
        cur.connection.commit()
    _READY = True


def ensure_cascade_trigger() -> dict:
    """Wire the relation **order → line items**: when an ``order_headers`` row
    (one PO = one "order") is deleted, its ``order_lines`` are auto-deleted.

    Implemented as an ``AFTER DELETE`` trigger on ``order_headers`` so the cascade
    fires no matter HOW the header is removed (Django admin, raw SQL, a future
    cleanup). It only deletes from the **web-owned** ``order_lines`` table — the
    engine's insert behaviour is untouched. MySQL only (no-op elsewhere)."""
    with _conn() as (cur, d):
        if d['kind'] != 'mysql':
            return {'ok': True, 'skipped': 'not-mysql'}
        try:
            cur.execute("DROP TRIGGER IF EXISTS trg_order_headers_cascade_lines")
            cur.execute(
                """
                CREATE TRIGGER trg_order_headers_cascade_lines
                AFTER DELETE ON order_headers
                FOR EACH ROW
                DELETE FROM order_lines
                 WHERE run_id      <=> OLD.run_id
                   AND marketplace <=> OLD.marketplace
                   AND po          <=> OLD.po
                """)
            cur.connection.commit()
            return {'ok': True}
        except Exception as e:  # noqa: BLE001 — needs TRIGGER privilege
            return {'ok': False, 'error': f"{type(e).__name__}: {e}"}


def ensure_run_cascade_trigger() -> dict:
    """Wire the relation **run → orders → line items**: when a ``runs`` row is
    deleted, every ``order_headers`` row for that ``run_id`` is auto-deleted,
    which in turn fires :func:`ensure_cascade_trigger` to drop their
    ``order_lines``. So deleting a run removes its headers *and* lines in one go
    (no more orphaned headers/lines when a run is cleaned up).

    ``AFTER DELETE`` trigger on ``runs`` → fires however the run is removed (UI,
    raw SQL, admin). MySQL only (no-op elsewhere). Needs TRIGGER privilege."""
    with _conn() as (cur, d):
        if d['kind'] != 'mysql':
            return {'ok': True, 'skipped': 'not-mysql'}
        try:
            cur.execute("DROP TRIGGER IF EXISTS trg_runs_cascade_headers")
            cur.execute(
                """
                CREATE TRIGGER trg_runs_cascade_headers
                AFTER DELETE ON runs
                FOR EACH ROW
                DELETE FROM order_headers WHERE run_id <=> OLD.run_id
                """)
            cur.connection.commit()
            return {'ok': True}
        except Exception as e:  # noqa: BLE001 — needs TRIGGER privilege
            return {'ok': False, 'error': f"{type(e).__name__}: {e}"}


def ean_alias_map() -> dict:
    """Historical EAN corrections, derived from the validation layer:
    ``{received_ean (wrong) → ean (correct)}``. Lets a repeat wrong EAN auto-
    resolve (fed into the engine's alias step) — no separate alias table.
    Read-only; empty on any error."""
    try:
        with _conn() as (cur, d):
            cur.execute(
                "SELECT DISTINCT received_ean, ean FROM order_lines_full "
                "WHERE received_ean IS NOT NULL AND received_ean <> ''")
            return {str(w).strip(): str(c).strip()
                    for w, c in cur.fetchall() if w and c}
    except Exception:  # noqa: BLE001
        return {}


def set_order_value(run_id, value_by_po: dict) -> dict:
    """Stamp the inc-GST order value onto ``order_headers`` for a run, per PO.
    Used for amount-less Transfer Orders (e.g. Flipkart Branch) whose dump
    carries no amount — the value is computed from our master pricing
    (Landing × qty) in the engine bridge. No-op on empty input."""
    if run_id is None or not value_by_po:
        return {'updated': 0}
    # Batched: ONE UPDATE with a per-PO CASE instead of N round-trips (the enrich
    # phase was ~1 query per PO → 100+ round-trips on a big run). Same effect:
    # order_value is overwritten for each listed PO of this run.
    pairs = [(str(po), _f(val)) for po, val in value_by_po.items()]
    with _conn() as (cur, d):
        ph = d['ph']
        whens = ' '.join(f"WHEN {ph} THEN {ph}" for _ in pairs)
        case_args = [x for po, val in pairs for x in (po, val)]
        in_ph = ', '.join([ph] * len(pairs))
        pos = [po for po, _ in pairs]
        cur.execute(
            f"UPDATE order_headers SET order_value = CASE po {whens} "
            f"ELSE order_value END WHERE run_id={ph} AND po IN ({in_ph})",
            tuple(case_args) + (run_id,) + tuple(pos))
        updated = cur.rowcount or 0
        # keep the run's stored aggregate in sync (engine wrote 0).
        cur.execute(
            f"UPDATE runs SET total_value="
            f"(SELECT COALESCE(SUM(order_value),0) FROM order_headers "
            f"WHERE run_id={ph}) WHERE run_id={ph}", (run_id, run_id))
        cur.connection.commit()
    return {'updated': updated}


def set_po_dates(run_id, dates_by_po: dict, force=False) -> dict:
    """Backfill ``po_date`` / ``exp_date`` on ``order_headers`` for a run, per PO.
    Default (``force=False``) uses ``COALESCE`` so engine-provided dates are never
    overwritten — for PDF marketplaces whose parser carries the date in the header
    (not a row column), filling blanks only. ``force=True`` OVERWRITES the engine's
    value — needed when the engine parsed the date WRONGLY (e.g. Swiggy's
    day-first ``PoCreatedAt`` timestamp misread month-first for days 1–12). Powers
    the TAT tracker."""
    if run_id is None or not dates_by_po:
        return {'updated': 0}
    # Batched: ONE UPDATE with a per-column CASE, instead of one round-trip per PO.
    # Semantics preserved EXACTLY: a PO only touches a column it carries; non-force
    # uses COALESCE (fill blanks only), force overwrites; POs carrying neither date
    # are excluded from the WHERE so they're untouched. Each column's ELSE keeps a
    # listed PO that lacks THAT date unchanged.
    po_items = [(str(po), dd['po_date']) for po, dd in dates_by_po.items() if dd.get('po_date')]
    exp_items = [(str(po), dd['exp_date']) for po, dd in dates_by_po.items() if dd.get('exp_date')]
    all_pos = [str(po) for po, dd in dates_by_po.items()
               if dd.get('po_date') or dd.get('exp_date')]
    if not all_pos:
        return {'updated': 0}
    with _conn() as (cur, d):
        ph = d['ph']
        sets, args = [], []
        if po_items:
            case_sql = 'po_date = CASE po '
            for po, val in po_items:
                case_sql += (f"WHEN {ph} THEN {ph} " if force
                             else f"WHEN {ph} THEN COALESCE(po_date,{ph}) ")
                args += [po, val]
            sets.append(case_sql + 'ELSE po_date END')
        if exp_items:
            case_sql = 'exp_date = CASE po '
            for po, val in exp_items:
                case_sql += (f"WHEN {ph} THEN {ph} " if force
                             else f"WHEN {ph} THEN COALESCE(exp_date,{ph}) ")
                args += [po, val]
            sets.append(case_sql + 'ELSE exp_date END')
        in_ph = ', '.join([ph] * len(all_pos))
        cur.execute(
            f"UPDATE order_headers SET {', '.join(sets)} "
            f"WHERE run_id={ph} AND po IN ({in_ph})",
            tuple(args) + (run_id,) + tuple(all_pos))
        updated = cur.rowcount or 0
        cur.connection.commit()
    return {'updated': updated}


def set_location(run_id, loc_by_po: dict) -> dict:
    """Overwrite ``order_headers.location`` on a run, per PO, with a resolved
    SHORT name (e.g. 'Mumbai', 'West bengal'). Used where the engine needs the
    RAW ship-to address to do its own resolution (so we can't pre-shorten what it
    reads) but the tracker should show the friendly short name — Myntra. Unlike
    the dates backfill this OVERWRITES (the engine wrote the raw address). No-op
    on empty input or blank values."""
    if run_id is None or not loc_by_po:
        return {'updated': 0}
    # Batched: ONE UPDATE with a per-PO CASE instead of one round-trip per PO.
    # Blank locations are skipped (same as the old `if not loc: continue`).
    pairs = [(str(po), loc) for po, loc in loc_by_po.items() if loc]
    if not pairs:
        return {'updated': 0}
    with _conn() as (cur, d):
        ph = d['ph']
        whens = ' '.join(f"WHEN {ph} THEN {ph}" for _ in pairs)
        case_args = [x for po, loc in pairs for x in (po, loc)]
        in_ph = ', '.join([ph] * len(pairs))
        pos = [po for po, _ in pairs]
        cur.execute(
            f"UPDATE order_headers SET location = CASE po {whens} "
            f"ELSE location END WHERE run_id={ph} AND po IN ({in_ph})",
            tuple(case_args) + (run_id,) + tuple(pos))
        updated = cur.rowcount or 0
        cur.connection.commit()
    return {'updated': updated}


def web_dedup(result, marketplace) -> list:
    """Web-owned replica of the engine's ``apply_dedup`` — drop POs already in
    ``order_headers`` for this marketplace from ``result.rows`` (so a re-upload
    isn't recorded twice) and summarise them on ``result.skipped_orders``. Reuses
    the engine's PURE ``build_tracker_rows`` for the summary. DB is only READ here
    — no engine history store is opened (so no desktop tables get recreated)."""
    result.skipped_orders = []
    rows = getattr(result, 'rows', None) or []
    if not rows:
        return []
    try:
        from online_po_processor.auto.history_db import (
            DEDUP_SKIP_ENABLED,
            ORDER_SEGMENT,
            build_tracker_rows,
        )
    except Exception:  # noqa: BLE001
        return []
    if not DEDUP_SKIP_ENABLED:
        return []
    existing = set()
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(f"SELECT DISTINCT po FROM order_headers WHERE "
                        f"marketplace={ph}", (marketplace,))
            existing = {str(r[0]) for r in cur.fetchall()}
    except Exception:  # noqa: BLE001
        return []
    dup = {str(so.po_number) for so in rows if str(so.po_number) in existing}
    if not dup:
        return []
    trk = {str(t['po']): t for t in build_tracker_rows(result)}
    from online_po_processor.auto.history_db import _to_date as _hd_to_date

    def _dd(v):   # clean date → DD-MM-YYYY for the "already uploaded" tab (display only)
        dv = _hd_to_date(v)
        return dv.strftime('%d-%m-%Y') if dv else ''
    skipped = []
    for po in dup:
        t = trk.get(po, {})
        skipped.append({
            'segment': ORDER_SEGMENT, 'marketplace': marketplace,
            'marketplace_label': t.get('market_place', marketplace), 'po': po,
            'location': t.get('location', '') or '',
            'po_date': _dd(t.get('po_date')), 'exp_date': _dd(t.get('exp_date')),
            'qty': int(t.get('order_qty') or 0),
            'order_value': float(t.get('order_value') or 0.0),
        })
    result.rows = [so for so in rows if str(so.po_number) not in dup]
    result.skipped_orders = skipped
    return skipped


import logging as _logging
_dlog = _logging.getLogger(__name__)

# Quick-commerce channels whose delivery window is DAYS, not months — used to
# catch DD/MM↔MM/DD swapped source dates at ingest (e.g. a Zepto exp read as
# 08-Nov instead of 11-Aug). Long-window channels (Nykaa ~60d, Reliance) are
# deliberately NOT listed: their large gaps are legit, so we never "correct" them.
_SHORT_TAT_CHANNELS = {'Zepto', 'Blinkit', 'Swiggy', 'BlinkMP'}
_SHORT_TAT_MAX = 45


def _swap_daymonth(d):
    import datetime as _d
    try:
        return _d.date(d.year, d.day, d.month)
    except (ValueError, TypeError, AttributeError):
        return None


def sane_po_exp(marketplace, po_d, exp_d):
    """Guard against DD/MM↔MM/DD swapped source dates. Returns
    ``(po_d, exp_d, note)`` — ``note`` is '' when unchanged, else the reason.
    Corrects ONLY provably-wrong cases: (1) exp before po (impossible), or (2) a
    SHORT-TAT channel gap that a day/month swap resolves into a sane window.
    Never touches legit long-window channels (they're not in the short-TAT set),
    so Nykaa's ~60-day terms are preserved."""
    if not po_d or not exp_d:
        return po_d, exp_d, ''
    gap = (exp_d - po_d).days
    if gap < 0:                                   # impossible → restore order
        sp = _swap_daymonth(po_d)
        if sp and 0 <= (exp_d - sp).days <= 90:
            return sp, exp_d, f'po_date {po_d}→{sp} (exp<po; DD/MM swap)'
        se = _swap_daymonth(exp_d)
        if se and 0 <= (se - po_d).days <= 90:
            return po_d, se, f'exp_date {exp_d}→{se} (exp<po; DD/MM swap)'
    elif gap > _SHORT_TAT_MAX and marketplace in _SHORT_TAT_CHANNELS:
        se = _swap_daymonth(exp_d)
        if se and 0 <= (se - po_d).days <= _SHORT_TAT_MAX:
            return po_d, se, f'exp_date {exp_d}→{se} ({marketplace} gap {gap}d; DD/MM swap)'
        sp = _swap_daymonth(po_d)
        if sp and 0 <= (exp_d - sp).days <= _SHORT_TAT_MAX:
            return sp, exp_d, f'po_date {po_d}→{sp} ({marketplace} gap {gap}d; DD/MM swap)'
    return po_d, exp_d, ''


def record_run_headers(result, marketplace, warehouse, output_file='',
                       as_of=None) -> dict:
    """Web-owned replica of the engine's ``record_manual`` — write ``runs`` +
    ``order_headers`` DIRECTLY (no engine history store, so ``order_issue_lines``
    is never recreated). Reuses the engine's PURE ``order_rows_from_result`` for
    header derivation, so the rows are byte-identical to the old path. Returns
    ``{run_id, new_orders}``.

    ``as_of`` (a ``datetime``) BACK-DATES ``run_ts`` + ``created_at`` to that
    moment instead of now — used when finalizing a *Review-later* draft, so the
    whole record belongs to the day the draft was parked (see the confirm flow).
    ``None`` = stamp now (the normal path)."""
    import datetime as _dt2
    import os as _os

    from online_po_processor.auto.history_db import (
        ORDER_SEGMENT,
        _to_date,
        order_rows_from_result,
    )
    rows = order_rows_from_result(result, marketplace, warehouse or '', output_file)
    if not rows:
        return {'run_id': None, 'new_orders': 0}
    run_ts = as_of or _utc_now()
    source = (f"MANUAL: {_os.path.basename(output_file)}" if output_file
              else 'MANUAL')
    meta = {
        'marketplaces': 1,
        'total_pos': len({(o['marketplace'], o['po']) for o in rows}),
        # total_items = LINE-item count (matches the engine's record_manual),
        # not the header/PO count.
        'total_items': len(getattr(result, 'rows', []) or []),
        'total_qty': sum(int(o['qty'] or 0) for o in rows),
        'total_value': sum(float(o['order_value'] or 0) for o in rows),
    }
    with _conn() as (cur, d):
        run_id = _insert_run_and_headers(cur, d['ph'], run_ts, source, meta, rows)
        cur.connection.commit()
    return {'run_id': run_id, 'new_orders': len(rows)}


def _insert_run_and_headers(cur, ph, run_ts, source, meta, rows, recorded_by=None):
    """INSERT the ``runs`` row + all ``order_headers`` on an EXISTING cursor (NO
    commit — the caller owns the transaction). ``recorded_by`` = the user who
    clicked Lock & Record (audit). Returns the new run_id."""
    from online_po_processor.auto.history_db import ORDER_SEGMENT, _to_date
    cur.execute(
        f"INSERT INTO runs (run_ts, mode, source, marketplaces, total_pos, "
        f"total_items, total_qty, total_value, consolidated_path, "
        f"tracker_path, recorded_by, recorded_at) VALUES ({ph},'MANUAL',{ph},{ph},"
        f"{ph},{ph},{ph},{ph},'','',{ph},{ph})",
        (run_ts, source, meta['marketplaces'], meta['total_pos'],
         meta['total_items'], meta['total_qty'], meta['total_value'],
         (str(recorded_by)[:150] if recorded_by else None),
         _utc_now()))                        # recorded_at = ACTUAL record time (UTC)
    run_id = cur.lastrowid
    hcols = ('run_id, run_ts, created_at, mode, segment, marketplace, '
             'marketplace_label, po, location, warehouse, po_date, exp_date, '
             'order_type, items, qty, order_value, output_file')
    marks = ', '.join([ph] * 17)
    # BATCHED (executemany) — the old per-PO loop paid one network round-trip PER
    # HEADER, so a 129-PO run cost ~129 hops over remote TiDB. Build the payload
    # (with the per-row date-guard) first, then insert in a single round-trip.
    payload = []
    for o in rows:
        # date-guard: catch DD/MM-swapped source dates before they land
        pod, exd = _to_date(o['po_date']), _to_date(o['exp_date'])
        pod, exd, _note = sane_po_exp(o['marketplace_label'], pod, exd)
        if _note:
            _dlog.warning("date-guard: PO %s %s", o['po'], _note)
        payload.append(
            (run_id, run_ts, run_ts, 'MANUAL', o.get('segment', ORDER_SEGMENT),
             o['marketplace'], o['marketplace_label'], o['po'], o['location'],
             o['warehouse'], pod, exd,
             o['order_type'], o['items'], o['qty'], o['order_value'],
             o['output_file']))
    if payload:
        cur.executemany(f"INSERT INTO order_headers ({hcols}) VALUES ({marks})", payload)
    return run_id


def _insert_line_rows(cur, ph, run_id, rows):
    """INSERT ``order_lines`` + ``order_line_validation`` on an EXISTING cursor
    (NO commit — caller owns the transaction). Returns the count inserted.

    BATCHED for speed: the old code inserted facts one row at a time to read each
    ``lastrowid`` — that meant one network round-trip PER LINE, so a 3.4k-line run
    took ~6 min over remote TiDB (89 ms/hop). Now facts go in via ``executemany``
    (a few round-trips), then we read the new ``line_id``s back **in insertion
    order** (``WHERE run_id AND line_id > prev_max ORDER BY line_id``) and pair the
    1:1 validation rows POSITIONALLY. This never assumes auto-increment
    contiguity, so it's safe on TiDB (cached, gap-prone IDs) and under concurrency
    (the run_id filter isolates this run). A readback/row-count mismatch raises →
    the whole transaction rolls back (atomic; never a mis-paired or partial write).
    ~250x fewer round-trips → ~6 min becomes ~1-2 s."""
    if not rows:
        return 0
    ins_f = (f"INSERT INTO order_lines ({', '.join(_FACT_COLS)}) "
             f"VALUES ({', '.join([ph] * len(_FACT_COLS))})")
    # remember the run's current high-water line_id so we read back ONLY the rows
    # this call inserts (a backfill run may already have some).
    cur.execute(f"SELECT COALESCE(MAX(line_id), 0) FROM order_lines WHERE run_id={ph}",
                (run_id,))
    prev_max = cur.fetchone()[0] or 0
    cur.executemany(ins_f, [tuple(r.get(c) for c in _FACT_COLS) for r in rows])
    # read back the freshly-inserted ids in insertion order
    cur.execute(f"SELECT line_id FROM order_lines WHERE run_id={ph} AND line_id > {ph} "
                f"ORDER BY line_id", (run_id, prev_max))
    ids = [r[0] for r in cur.fetchall()]
    if len(ids) != len(rows):
        # never pair mismatched sets — bail so the transaction rolls back
        raise RuntimeError(
            f"order_lines id read-back mismatch: {len(ids)} new ids for "
            f"{len(rows)} rows (run {run_id}) — rolling back.")
    vcols = ['line_id'] + _VAL_COLS
    ins_v = (f"INSERT INTO order_line_validation ({', '.join(vcols)}) "
             f"VALUES ({', '.join([ph] * len(vcols))})")
    cur.executemany(ins_v, [(ids[i], *(rows[i].get(c) for c in _VAL_COLS))
                            for i in range(len(rows))])
    return len(rows)


def insert_lines_for_run(run_id, run_ts, line_rows) -> int:
    """Write order_lines (+ 1:1 validation) for a run whose HEADERS were recorded
    elsewhere — used by the offline EKA path, which records headers via history_db
    but not lines. Each row is a fact+validation dict; run_id/run_ts are stamped
    in. Atomic (own transaction). Returns the count written."""
    rows = [r for r in (line_rows or []) if r]
    if run_id is None or not rows:
        return 0
    for r in rows:
        r['run_id'] = run_id
        if not r.get('run_ts'):
            r['run_ts'] = run_ts
    with _conn_tx() as (cur, d):
        return _insert_line_rows(cur, d['ph'], run_id, rows)


def _ensure_run_recorded_by(cur) -> None:
    """Additive: make sure the ``runs`` table has the audit columns —
    ``recorded_by`` (who clicked Lock & Record) and ``recorded_at`` (the ACTUAL
    record time; ``run_ts`` is back-dated to the upload day, so it's not the real
    moment). Idempotent — each ALTER is a no-op/ignored if it already exists. DDL,
    so call OUTSIDE the atomic transaction (it auto-commits)."""
    for ddl in ("ALTER TABLE runs ADD COLUMN recorded_by VARCHAR(150)",
                "ALTER TABLE runs ADD COLUMN recorded_at DATETIME"):
        try:
            cur.execute(ddl)
        except Exception:  # noqa: BLE001 — already present (or race) → fine
            pass


def record_run_atomic(result, marketplace, warehouse, output_file, line_rows,
                      as_of=None, recorded_by=None) -> dict:
    """Lock & Record as ONE transaction: ``runs`` + ``order_headers`` +
    ``order_lines`` + ``order_line_validation`` are written together and
    COMMITTED only after ALL succeed. Any error / interruption / crash rolls the
    WHOLE thing back — the DB is 100%% written or completely untouched, never a
    partial run. Returns ``{run_id, new_orders, lines_recorded}``; raises on
    failure so the caller can report it (nothing was recorded)."""
    import datetime as _dt2
    import os as _os

    from online_po_processor.auto.history_db import order_rows_from_result
    rows = order_rows_from_result(result, marketplace, warehouse or '', output_file)
    if not rows:
        return {'run_id': None, 'new_orders': 0, 'lines_recorded': 0}
    ensure_table()   # create order_lines/validation tables BEFORE the tx (DDL auto-commits)
    with _conn() as (_cur, _d):          # ensure recorded_by col before the tx (DDL)
        _ensure_run_recorded_by(_cur)
    run_ts = as_of or _utc_now()
    source = (f"MANUAL: {_os.path.basename(output_file)}" if output_file else 'MANUAL')
    meta = {
        'marketplaces': 1,
        'total_pos': len({(o['marketplace'], o['po']) for o in rows}),
        'total_items': len(getattr(result, 'rows', []) or []),
        'total_qty': sum(int(o['qty'] or 0) for o in rows),
        'total_value': sum(float(o['order_value'] or 0) for o in rows),
    }
    with _conn_tx() as (cur, d):
        ph = d['ph']
        run_id = _insert_run_and_headers(cur, ph, run_ts, source, meta, rows,
                                         recorded_by=recorded_by)
        for r in (line_rows or []):      # stamp the just-created run_id onto each line
            r['run_id'] = run_id
        n = _insert_line_rows(cur, ph, run_id, line_rows or [])
        # _conn_tx COMMITS on clean exit; ANY exception above → full ROLLBACK.
    return {'run_id': run_id, 'new_orders': len(rows), 'lines_recorded': n}


def apply_issue_ean_fix(line_id, correct_ean) -> dict:
    """Post-lock EAN correction (Issues page) for a NOT_IN_MASTER line: resolve
    the correct item, recompute OUR pricing with the ENGINE's own helpers, and
    update the locked line — facts (ean / item_no / description) +
    order_line_validation (received_ean=the wrong EAN, our_*, diff, status,
    exception_label='EAN remap'). Engine untouched. Returns
    {ok, status, item_no, ean, our_cp, diff}."""
    try:
        from online_po_processor.config.marketplaces import MARKETPLACE_CONFIGS
        from online_po_processor.data.master_loader import MasterLoader

        from . import item_master_loader as iml
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"{type(e).__name__}: {e}"}
    hit = iml.resolve_in_master(correct_ean)
    if not hit:
        return {'ok': False, 'error': f"'{correct_ean}' is not in the item master."}
    try:
        lid = int(line_id)
    except (TypeError, ValueError):
        return {'ok': False, 'error': 'bad line_id'}

    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(
            f"SELECT marketplace, ean, vendor_cp, vendor_landing, margin_pct "
            f"FROM order_lines_full WHERE line_id={ph}", (lid,))
        row = cur.fetchone()
        if not row:
            return {'ok': False, 'error': 'line not found'}
        mp, old_ean, v_cp, v_land, margin_pct = row
        cfg = MARKETPLACE_CONFIGS.get(mp, {})
        margin = (float(margin_pct) / 100.0 if margin_pct
                  else cfg.get('default_margin', 70) / 100.0)
        mrp = float(hit['mrp']) if hit.get('mrp') is not None else None
        gst = hit.get('gst_code', '') or ''
        our_land = MasterLoader.calc_landing_price(mrp, margin)
        our_cp = MasterLoader.calc_cost_price(mrp, gst, margin)
        compare = cfg.get('compare_basis', 'cost')
        sbasis = cfg.get('status_basis') or compare

        def _pair(b):
            if b == 'landing':
                return our_land, (float(v_land) if v_land is not None else None)
            return our_cp, (float(v_cp) if v_cp is not None else None)
        o_disp, t_disp = _pair(compare)          # diff shown on compare basis
        disp_diff = (round(t_disp - o_disp, 2)
                     if (o_disp is not None and t_disp is not None) else None)
        o_st, t_st = _pair(sbasis)               # status decided on status basis
        st_diff = (t_st - o_st) if (o_st is not None and t_st is not None) else None
        status = 'OK' if (st_diff is not None and abs(st_diff) <= 0.5) else 'MISMATCH'

        cur.execute(
            f"UPDATE order_lines SET ean={ph}, item_no={ph}, description={ph} "
            f"WHERE line_id={ph}",
            (hit['ean'] or correct_ean, hit['item_no'],
             (hit['description'] or '')[:255], lid))
        cur.execute(
            f"UPDATE order_line_validation SET received_ean={ph}, our_mrp={ph}, "
            f"our_landing={ph}, our_cp={ph}, diff={ph}, margin_pct={ph}, "
            f"status={ph}, exception_label={ph} WHERE line_id={ph}",
            (old_ean, _f(mrp), _f(our_land), _f(our_cp), _f(disp_diff),
             round(margin * 100, 2), status, 'EAN remap', lid))
        cur.connection.commit()
    return {'ok': True, 'status': status, 'item_no': hit['item_no'],
            'ean': hit['ean'] or correct_ean, 'our_cp': _f(our_cp),
            'diff': _f(disp_diff)}


def build_lines(result, run_id=None, output_file: str = '', actions=None,
                ean_fixes=None, as_of=None) -> list[dict]:
    """One dict per SO line from an engine ``ProcessingResult`` (reads only).
    Carries the full comparison columns so affected = status filter. ``actions``
    is an optional ``{"po|item_no|ean": {"action":..,"remark":..}}`` map of the
    operator's per-line decision captured on the review screen.

    ``as_of`` (a ``datetime``) BACK-DATES each line's ``run_ts`` to that moment
    instead of now — keeps a finalized *Review-later* draft's lines on the day it
    was parked (matches the back-dated headers). ``None`` = now."""
    actions = actions or {}
    ean_fixes = ean_fixes or {}
    basis = getattr(result, 'compare_basis', None) or 'landing'
    run_ts = (as_of or _utc_now()).strftime('%Y-%m-%d %H:%M:%S')
    otype = 'TO' if getattr(result, 'output_type', 'so') == 'to' else 'SO'
    out: list[dict] = []
    for so in result.rows:
        v_landing = so.fob_price if basis == 'landing' else so.ref_fob_price
        v_cp = so.fob_price if basis == 'cost' else so.ref_fob_price
        row_margin = (so.applied_margin_pct
                      if getattr(so, 'applied_margin_pct', None) is not None
                      else result.margin_pct)
        o_landing = None
        if so.mrp is not None and _f(so.mrp) is not None and row_margin:
            o_landing = round(float(so.mrp) * float(row_margin), 2)
        up_base = (so.forced_unit_price
                   if getattr(so, 'forced_unit_price', None) is not None
                   else getattr(so, 'cost_price_ref', None))
        po = str(so.po_number)
        item = str(so.item_no or '')
        raw_ean = str(so.ean or '')
        # If this EAN was corrected, ship on the CORRECT one and remember the
        # WRONG one as received_ean (audit). order_lines never holds wrong data.
        recv_ean = None
        ean = raw_ean
        if raw_ean and raw_ean in ean_fixes:
            ean = str(ean_fixes[raw_ean])
            recv_ean = raw_ean
        act = (actions.get(f"{po}|{item}|{ean}")
               or actions.get(f"{po}|{item}|{raw_ean}") or {})
        # Recorded Unit Price reflects the operator's decision so the DB matches
        # the workbook + D365: INCLUDE → their (vendor) CP, OVERRIDE → our CP,
        # else the engine's own price. [[deal-sku-exception-behavior]]
        _ak = str(act.get('action') or '').upper()
        if _ak == 'INCLUDE':
            up = v_cp if v_cp is not None else up_base
        elif _ak == 'OVERRIDE' and str(act.get('override_cp') or '') != '':
            try:
                up = round(float(act.get('override_cp')), 2)
            except (TypeError, ValueError):
                up = up_base
        else:
            up = up_base
        out.append({
            'run_id': run_id, 'run_ts': run_ts,
            'marketplace': result.marketplace,
            'po': po,
            'location': str(getattr(so, 'source_location', '') or ''),
            'item_no': item,
            'ean': ean,
            'description': (str(so.description or ''))[:255],
            'qty': int(so.qty or 0),
            'order_type': otype,
            'gst_code': str(getattr(so, 'gst_code', '') or ''),
            'unit_price': _f(up),
            'vendor_mrp': _f(so.vendor_mrp),
            'our_mrp': _f(so.mrp),
            'vendor_landing': _f(v_landing),
            'our_landing': _f(o_landing),
            'vendor_cp': _f(v_cp),
            'our_cp': _f(so.cost_price_ref),
            'diff': _f(so.diffn),
            'margin_pct': round(float(row_margin) * 100, 2) if row_margin else None,
            'status': getattr(so, 'validation_status', '') or '',
            'exception_label': getattr(so, 'exception_label', '') or '',
            # wrong EAN as received (set by the review-page EAN fix); else None.
            'received_ean': (recv_ean or act.get('received_ean') or None),
            'output_file': output_file or '',
            'action': str(act.get('action', '') or ''),
            'remark': str(act.get('remark', '') or '')[:255],
            'override_cp': _f(act.get('override_cp')),
            # stamp WHEN the decision was taken (only when there IS one)
            'decided_at': run_ts if act.get('action') else None,
            # which rate the validation compared on (highlight on the UI).
            'basis': 'CP' if basis == 'cost' else 'Landing',
        })
    return out


def insert_lines(run_id, rows: list[dict]) -> dict:
    """Bulk-insert the line rows for a confirmed run. No-op if run_id is None
    (nothing new recorded) or rows empty."""
    if run_id is None or not rows:
        return {'recorded': 0}
    ensure_table()
    with _conn() as (cur, d):
        n = _insert_line_rows(cur, d['ph'], run_id, rows)
        cur.connection.commit()
    return {'recorded': n}


def update_action(line_id, action, remark) -> dict:
    """Set/update the operator's Action + Remark on an existing order_lines row
    (from the Issues page, after upload). Web-owned table — UPDATE one row."""
    try:
        lid = int(line_id)
    except (TypeError, ValueError):
        return {'ok': False, 'error': 'bad line_id'}
    ensure_table()
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(
            f"UPDATE order_line_validation SET action={ph}, remark={ph} "
            f"WHERE line_id={ph}",
            (str(action or '')[:20], str(remark or '')[:255], lid))
        cur.connection.commit()
    return {'ok': True}


def update_action_bulk(line_ids, action, remark) -> dict:
    """Set the same Action + Remark on many lines at once (Issues page bulk)."""
    ids = []
    for x in (line_ids or []):
        try:
            ids.append(int(x))
        except (TypeError, ValueError):
            pass
    if not ids:
        return {'ok': False, 'updated': 0}
    ensure_table()
    with _conn() as (cur, d):
        ph = d['ph']
        marks = ','.join([ph] * len(ids))
        cur.execute(
            f"UPDATE order_line_validation SET action={ph}, remark={ph} "
            f"WHERE line_id IN ({marks})",
            tuple([str(action or '')[:20], str(remark or '')[:255]] + ids))
        cur.connection.commit()
    return {'ok': True, 'updated': len(ids)}
