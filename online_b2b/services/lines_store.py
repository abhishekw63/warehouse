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

from .order_db import _conn

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
    INDEX idx_lines_item (item_no)
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


def ensure_table() -> None:
    """Create the facts table + validation table + join view (idempotent).
    Web owns all three; the engine schema is untouched."""
    with _conn() as (cur, d):
        mysql = d['kind'] == 'mysql'
        cur.execute(_MYSQL_FACTS if mysql else _SQLITE_FACTS)
        cur.execute(_MYSQL_VAL if mysql else _SQLITE_VAL)
        view_sql = (f"SELECT {_VIEW_SELECT} FROM order_lines l "
                    "LEFT JOIN order_line_validation v ON v.line_id = l.line_id")
        if mysql:
            cur.execute(f"CREATE OR REPLACE VIEW order_lines_full AS {view_sql}")
        else:
            cur.execute("DROP VIEW IF EXISTS order_lines_full")
            cur.execute(f"CREATE VIEW order_lines_full AS {view_sql}")
        cur.connection.commit()


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
    updated = 0
    with _conn() as (cur, d):
        ph = d['ph']
        for po, val in value_by_po.items():
            cur.execute(
                f"UPDATE order_headers SET order_value={ph} "
                f"WHERE run_id={ph} AND po={ph}", (_f(val), run_id, str(po)))
            updated += cur.rowcount or 0
        # keep the run's stored aggregate in sync (engine wrote 0).
        cur.execute(
            f"UPDATE runs SET total_value="
            f"(SELECT COALESCE(SUM(order_value),0) FROM order_headers "
            f"WHERE run_id={ph}) WHERE run_id={ph}", (run_id, run_id))
        cur.connection.commit()
    return {'updated': updated}


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
                ean_fixes=None) -> list[dict]:
    """One dict per SO line from an engine ``ProcessingResult`` (reads only).
    Carries the full comparison columns so affected = status filter. ``actions``
    is an optional ``{"po|item_no|ean": {"action":..,"remark":..}}`` map of the
    operator's per-line decision captured on the review screen."""
    actions = actions or {}
    ean_fixes = ean_fixes or {}
    basis = getattr(result, 'compare_basis', None) or 'landing'
    run_ts = _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
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
        up = (so.forced_unit_price
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
        ph = d['ph']
        ins_f = (f"INSERT INTO order_lines ({', '.join(_FACT_COLS)}) "
                 f"VALUES ({', '.join([ph] * len(_FACT_COLS))})")
        vcols = ['line_id'] + _VAL_COLS
        ins_v = (f"INSERT INTO order_line_validation ({', '.join(vcols)}) "
                 f"VALUES ({', '.join([ph] * len(vcols))})")
        # Insert facts row-by-row to capture each line_id, then bulk-insert the
        # matching validation rows keyed by that line_id (1:1).
        val_payload = []
        for r in rows:
            cur.execute(ins_f, tuple(r.get(c) for c in _FACT_COLS))
            lid = cur.lastrowid
            val_payload.append((lid, *(r.get(c) for c in _VAL_COLS)))
        if val_payload:
            cur.executemany(ins_v, val_payload)
        cur.connection.commit()
    return {'recorded': len(rows)}


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
