"""
auto.history_db
===============

AUTO mode (v2.4.0) — order-upload **history** in a local SQLite file.

Purpose (the operator's words): *"track which orders we are uploading."*
Every Auto run appends a ``runs`` row and one ``orders`` row per
``(marketplace, PO)``. Because the orders are keyed on
``(marketplace, PO)``, the tool can tell on the next run whether a PO has
already been uploaded — **dedup**. Per the operator's choice, duplicates
are *flagged, not skipped*: the run still produces them, but the summary
calls out "already uploaded on <date>" so it's a conscious re-upload, not
an accident.

Storage
-------
A single file ``<Dump>/Tracker/history.db`` (alongside the timestamped
tracker files). SQLite is append-only here, single-user; no server, no
setup. The DB is the queryable source of truth; :meth:`export_to_xlsx`
dumps a human-readable view on demand.

Nothing in here is seeded from the old manual tracker — history starts
fresh from the first Auto run forward (operator's choice).
"""

from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from online_po_processor.exporter.sheets.tracker_sheet import build_tracker_rows


_RUNS_DDL = """
CREATE TABLE IF NOT EXISTS runs (
    run_id            INTEGER PRIMARY KEY AUTOINCREMENT,
    run_ts            TEXT,
    online_root       TEXT,
    marketplaces      INTEGER,
    total_pos         INTEGER,
    total_items       INTEGER,
    total_qty         INTEGER,
    total_value       REAL,
    consolidated_path TEXT,
    tracker_path      TEXT
)
"""

_ORDERS_DDL = """
CREATE TABLE IF NOT EXISTS orders (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id            INTEGER,
    run_ts            TEXT,
    marketplace       TEXT,   -- our config key (e.g. 'Firstcry')
    marketplace_label TEXT,   -- tracker label (e.g. 'First Cry')
    po                TEXT,
    location          TEXT,
    warehouse         TEXT,
    po_date           TEXT,
    exp_date          TEXT,
    order_type        TEXT,   -- 'SO' | 'TO'
    items             INTEGER,
    qty               INTEGER,
    order_value       REAL,   -- GST-inclusive
    output_file       TEXT,
    is_duplicate      INTEGER,-- 1 if (marketplace, po) seen in an earlier run
    first_seen_ts     TEXT    -- when this (marketplace, po) was first recorded
)
"""

_ORDERS_INDEX = (
    "CREATE INDEX IF NOT EXISTS idx_orders_mp_po ON orders(marketplace, po)"
)


class HistoryDB:
    """Thin SQLite wrapper for the Auto run history."""

    def __init__(self, db_path) -> None:
        self.path = str(db_path)
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.path)
        self._init_schema()

    def _init_schema(self) -> None:
        cur = self.conn.cursor()
        cur.execute(_RUNS_DDL)
        cur.execute(_ORDERS_DDL)
        cur.execute(_ORDERS_INDEX)
        self.conn.commit()

    # ── dedup lookup ────────────────────────────────────────────────────
    def existing_first_seen(self) -> Dict[Tuple[str, str], str]:
        """``{(marketplace, po): first_seen_ts}`` for everything recorded
        so far (across all prior runs)."""
        cur = self.conn.cursor()
        cur.execute(
            "SELECT marketplace, po, MIN(run_ts) FROM orders "
            "GROUP BY marketplace, po"
        )
        return {(m, p): ts for m, p, ts in cur.fetchall()}

    # ── write ───────────────────────────────────────────────────────────
    def record(self, run_meta: dict, order_rows: List[dict]):
        """
        Insert one run + its orders in a single transaction. Marks each
        order ``is_duplicate`` if its ``(marketplace, po)`` already
        existed BEFORE this run. Returns ``(run_id, duplicates)`` where
        ``duplicates`` is a list of ``(label, po, first_seen_ts)``.
        """
        existing = self.existing_first_seen()
        cur = self.conn.cursor()
        cur.execute(
            "INSERT INTO runs (run_ts, online_root, marketplaces, "
            "total_pos, total_items, total_qty, total_value, "
            "consolidated_path, tracker_path) VALUES (?,?,?,?,?,?,?,?,?)",
            (run_meta['run_ts'], run_meta['online_root'],
             run_meta['marketplaces'], run_meta['total_pos'],
             run_meta['total_items'], run_meta['total_qty'],
             run_meta['total_value'], run_meta['consolidated_path'],
             run_meta['tracker_path']),
        )
        run_id = cur.lastrowid

        duplicates: List[Tuple[str, str, str]] = []
        for o in order_rows:
            key = (o['marketplace'], o['po'])
            first_seen = existing.get(key)
            is_dup = 1 if first_seen is not None else 0
            if is_dup:
                duplicates.append((o['marketplace_label'], o['po'], first_seen))
            cur.execute(
                "INSERT INTO orders (run_id, run_ts, marketplace, "
                "marketplace_label, po, location, warehouse, po_date, "
                "exp_date, order_type, items, qty, order_value, "
                "output_file, is_duplicate, first_seen_ts) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (run_id, run_meta['run_ts'], o['marketplace'],
                 o['marketplace_label'], o['po'], o['location'],
                 o['warehouse'], o['po_date'], o['exp_date'],
                 o['order_type'], o['items'], o['qty'], o['order_value'],
                 o['output_file'], is_dup,
                 first_seen if is_dup else run_meta['run_ts']),
            )
        self.conn.commit()
        return run_id, duplicates

    # ── read / export ───────────────────────────────────────────────────
    def export_to_xlsx(self, out_path) -> str:
        """Dump the full ``orders`` history to an .xlsx (newest first)."""
        from openpyxl import Workbook
        from online_po_processor.exporter._styles import (
            auto_width, data_cell, hdr_cell,
        )
        cols = ['Run #', 'Run Time', 'Market Place', 'PO', 'Location',
                'Warehouse', 'PO Date', 'Exp Date', 'Type', 'Items', 'Qty',
                'Order Value', 'Duplicate?', 'First Uploaded', 'Output File']
        cur = self.conn.cursor()
        cur.execute(
            "SELECT run_id, run_ts, marketplace_label, po, location, "
            "warehouse, po_date, exp_date, order_type, items, qty, "
            "order_value, is_duplicate, first_seen_ts, output_file "
            "FROM orders ORDER BY run_id DESC, id ASC"
        )
        rows = cur.fetchall()
        wb = Workbook()
        ws = wb.active
        ws.title = 'Order History'
        for c, h in enumerate(cols, 1):
            hdr_cell(ws, 1, c, h)
        for r, rec in enumerate(rows, start=2):
            for c, val in enumerate(rec, start=1):
                if c == 13:                      # Duplicate? → Yes/No
                    val = 'Yes' if val else ''
                data_cell(ws, r, c, val,
                          align='left' if c in (5, 15) else 'center')
        auto_width(ws)
        ws.freeze_panes = 'A2'
        wb.save(str(out_path))
        return str(out_path)

    def close(self) -> None:
        self.conn.close()


# ── helpers / top-level entry point ─────────────────────────────────────

def _order_rows_for_run(run) -> List[dict]:
    """Per-PO order rows for one MarketplaceRun (reuses the tracker
    row-builder for label / location / dates / inc-GST value)."""
    res = run.result
    trk = {str(row['po']): row for row in build_tracker_rows(res)}
    items_by_po: Dict[str, int] = {}
    for so in res.rows:
        po = str(so.po_number)
        items_by_po[po] = items_by_po.get(po, 0) + 1
    otype = 'TO' if getattr(res, 'output_type', 'so') == 'to' else 'SO'

    out: List[dict] = []
    for po, t in trk.items():
        out.append({
            'marketplace': run.marketplace,
            'marketplace_label': t['market_place'],
            'po': po,
            'location': t['location'] or '',
            'warehouse': run.warehouse or '',
            'po_date': str(t['po_date']) if t['po_date'] else '',
            'exp_date': str(t['exp_date']) if t['exp_date'] else '',
            'order_type': otype,
            'items': items_by_po.get(po, 0),
            'qty': int(t['order_qty'] or 0),
            'order_value': float(t['order_value'] or 0.0),
            'output_file': run.output_path or '',
        })
    return out


def history_db_path(online_root: str) -> Path:
    """``<Dump>/Tracker/history.db`` (online_root is ``<Dump>/Online``)."""
    return Path(online_root).parent / 'Tracker' / 'history.db'


def record_history(runs: List, online_root: str,
                   consolidated_path: str = '', tracker_path: str = '',
                   run_ts: Optional[str] = None) -> dict:
    """
    Append this Auto run to the history DB and return a small summary.

    Returns ``{db_path, run_id, total_orders, new_orders, duplicates}``
    where ``duplicates`` is a list of ``(label, po, first_seen_ts)``.
    """
    ok = [r for r in runs if r.status == 'ok' and r.result is not None]
    order_rows: List[dict] = []
    for run in ok:
        order_rows.extend(_order_rows_for_run(run))

    run_meta = {
        'run_ts': run_ts or datetime.now().isoformat(timespec='seconds'),
        'online_root': str(online_root),
        'marketplaces': len({r.marketplace for r in ok}),
        'total_pos': sum(r.pos for r in ok),
        'total_items': sum(r.rows for r in ok),
        'total_qty': sum(r.qty for r in ok),
        'total_value': sum(o['order_value'] for o in order_rows),
        'consolidated_path': str(consolidated_path or ''),
        'tracker_path': str(tracker_path or ''),
    }

    db = HistoryDB(history_db_path(online_root))
    try:
        run_id, dups = db.record(run_meta, order_rows)
    finally:
        db.close()

    return {
        'db_path': str(history_db_path(online_root)),
        'run_id': run_id,
        'total_orders': len(order_rows),
        'new_orders': len(order_rows) - len(dups),
        'duplicates': dups,
    }
