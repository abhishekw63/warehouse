"""
auto.history_db
===============

Order-upload **history** with a backend-agnostic storage layer (v2.4.0).

Purpose (operator's words): *"track which orders we are uploading."* Both
**Auto** mode and **Manual** mode record every order they generate into one
shared history, so deduplication ("was this PO already uploaded?") spans
both. The history is the source of truth for what's been pushed to D365.

Scalable storage
----------------
Nothing in the app talks to SQLite directly — everything goes through the
:class:`HistoryStore` interface. Today the only implementation is
:class:`SqliteHistoryStore` (a single ``history.db`` file, viewable with
any desktop SQL tool like DB Browser / DBeaver). To move to a server
database later (SQL Server / Postgres for a shared GUI dashboard), add a
new ``HistoryStore`` implementation and point :func:`get_history_store`
at it — **no call-site changes** in Auto, Manual, the runner or the GUI.

Storage location
----------------
One shared DB at ``<Dump>/Tracker/history.db`` (:func:`default_history_db_path`),
used by both modes. Auto derives the same path from its run folder; Manual
uses the default resolver (it has no run folder).

Schema
------
``runs``   — one row per processing run (Auto batch or a single Manual
             generate): timestamp, mode, totals, output paths.
``orders`` — one row per ``(marketplace, PO)``: marketplace, PO, location,
             warehouse, dates, type (SO/TO), items, qty, GST-inclusive
             value, output file, ``is_duplicate`` + ``first_seen_ts``,
             and the mode that produced it.
"""

from __future__ import annotations

import json
import os
import sqlite3
from abc import ABC, abstractmethod
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from online_po_processor.config.constants import (
    DEDUP_SKIP_ENABLED, ORDER_SEGMENT,
)
from online_po_processor.exporter.sheets.tracker_sheet import build_tracker_rows


# ── DB backend config (selects SQLite vs MySQL) ─────────────────────────
# Read from a LOCAL file kept OUT of the repo (it holds the DB password):
#   1. $ONLINE_PO_DB_CONFIG  (explicit path), else
#   2. %LOCALAPPDATA%\OnlinePOProcessor\db_config.json
# Shape: {"backend":"mysql","host":..,"port":..,"user":..,"password":..,
#         "database":"renee_orders"}. Absent / backend!='mysql' ⇒ SQLite.

def _db_config_path() -> Path:
    env = os.environ.get('ONLINE_PO_DB_CONFIG')
    if env:
        return Path(env)
    base = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
    return Path(base) / 'OnlinePOProcessor' / 'db_config.json'


def load_db_config() -> Optional[dict]:
    p = _db_config_path()
    if not p.exists():
        return None
    try:
        # utf-8-sig tolerates a UTF-8 BOM (Windows editors / PowerShell
        # Set-Content add one), which plain json.load would reject.
        with open(p, 'r', encoding='utf-8-sig') as f:
            return json.load(f)
    except Exception:        # noqa: BLE001 — bad/locked config ⇒ fall back to SQLite
        return None


def _to_date(val):
    """Parse a tracker date string (dd-mm-yyyy / dd.mm.yyyy / …) to a
    ``date`` for DATE columns; ``None`` when blank/unparseable."""
    if not val:
        return None
    s = str(val).strip()
    for fmt in ('%d-%m-%Y', '%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _to_dt(val):
    """ISO/string timestamp → ``datetime`` for DATETIME columns."""
    if not val:
        return None
    if isinstance(val, datetime):
        return val
    try:
        return datetime.fromisoformat(str(val))
    except ValueError:
        return str(val).replace('T', ' ')


# ── Shared DB location ──────────────────────────────────────────────────

# Known Dump roots, in preference order. The history DB lives at
# ``<Dump>/Tracker/history.db`` so both modes share one file.
_DUMP_ROOTS = [
    r"D:\OneDrive - RENEE COSMETICS PRIVATE LIMITED\Dump",
]


def default_dump_root() -> Path:
    for p in _DUMP_ROOTS:
        if os.path.isdir(p):
            return Path(p)
    return Path(_DUMP_ROOTS[0])    # created on first write if absent


def default_history_db_path() -> Path:
    """The single shared history DB used by Manual mode (and the default
    Auto would resolve to as well)."""
    return default_dump_root() / 'Tracker' / 'history.db'


def history_db_path(online_root: str) -> Path:
    """Auto's DB path derived from its run folder (``<Dump>/Online``) —
    resolves to the SAME ``<Dump>/Tracker/history.db`` as the default."""
    return Path(online_root).parent / 'Tracker' / 'history.db'


# ── Storage interface ───────────────────────────────────────────────────

class HistoryStore(ABC):
    """Backend-agnostic history store. Swap the implementation (SQLite →
    SQL Server / Postgres) without touching any caller."""

    @abstractmethod
    def record(self, run_meta: dict, order_rows: List[dict]):
        """Insert a run + its (new) orders; return the new ``run_id``.
        Dedup is handled BEFORE this (already-uploaded POs are removed by
        ``apply_dedup``), so every row here is new — the DB holds only new
        POs and never tracks duplicates."""

    @abstractmethod
    def existing_pos(self) -> set:
        """Set of ``(marketplace, po)`` already in the DB — used by
        ``apply_dedup`` to decide which POs are already uploaded."""

    @abstractmethod
    def export_to_xlsx(self, out_path) -> str:
        """Dump the full order history to a readable .xlsx."""

    @abstractmethod
    def fetch_orders(self, run_id=None) -> List[dict]:
        """
        Read back order rows (for building the tracker FROM the DB — the DB
        is the single source of truth). Each dict carries: segment,
        marketplace, marketplace_label, po, location, po_date, exp_date,
        order_type, qty, order_value.

        ``run_id`` limits to one run. (The DB holds only new POs, so every
        row is genuinely new — no duplicate filtering is needed.)
        """

    @abstractmethod
    def close(self) -> None: ...


# Columns returned by fetch_orders (same shape for every backend).
_FETCH_COLS = ['segment', 'marketplace', 'marketplace_label', 'po', 'location',
               'po_date', 'exp_date', 'order_type', 'qty', 'order_value']


def _fetch_orders_sql(cur, table: str, ph: str, run_id):
    """Shared SELECT for fetch_orders (``ph`` is the param placeholder:
    '?' for SQLite, '%s' for MySQL; ``table`` differs per backend)."""
    q = f"SELECT {', '.join(_FETCH_COLS)} FROM {table}"
    params = []
    if run_id is not None:
        q += f" WHERE run_id={ph}"
        params.append(run_id)
    q += " ORDER BY " + ('id' if table == 'orders' else 'order_id') + " ASC"
    cur.execute(q, params)
    return [dict(zip(_FETCH_COLS, row)) for row in cur.fetchall()]


# Order History export — shared by both backends. Column headers (in DB
# select order); index-free styling keyed on the header NAME so columns can
# be added without breaking alignment.
_HISTORY_COLS = ['Run #', 'Run Time', 'Mode', 'Segment', 'Market Place', 'PO',
                 'Location', 'Warehouse', 'PO Date', 'Exp Date', 'Type',
                 'Items', 'Qty', 'Order Value', 'Output File']
_HISTORY_SELECT = (
    "run_id, run_ts, mode, segment, marketplace_label, po, location, "
    "warehouse, po_date, exp_date, order_type, items, qty, order_value, "
    "output_file")


def _write_history_xlsx(rows, out_path) -> str:
    from openpyxl import Workbook
    from online_po_processor.exporter._styles import (
        auto_width, data_cell, hdr_cell,
    )
    left = {'Location', 'Output File'}
    wb = Workbook()
    ws = wb.active
    ws.title = 'Order History'
    for c, h in enumerate(_HISTORY_COLS, 1):
        hdr_cell(ws, 1, c, h)
    for r, rec in enumerate(rows, start=2):
        for c, val in enumerate(rec, start=1):
            name = _HISTORY_COLS[c - 1]
            if isinstance(val, datetime):
                val = val.isoformat(sep=' ')
            elif hasattr(val, 'isoformat'):
                val = val.isoformat()
            data_cell(ws, r, c, val, align='left' if name in left else 'center')
    auto_width(ws)
    ws.freeze_panes = 'A2'
    wb.save(str(out_path))
    return str(out_path)


# ── SQLite implementation ───────────────────────────────────────────────

_RUNS_DDL = """
CREATE TABLE IF NOT EXISTS runs (
    run_id            INTEGER PRIMARY KEY AUTOINCREMENT,
    run_ts            TEXT,
    mode              TEXT,   -- 'AUTO' | 'MANUAL'
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
    mode              TEXT,   -- 'AUTO' | 'MANUAL'
    segment           TEXT,   -- 'OnlineB2B' (vs future offline/GT)
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
    output_file       TEXT
)
"""

_ORDERS_INDEX = (
    "CREATE INDEX IF NOT EXISTS idx_orders_mp_po ON orders(marketplace, po)"
)


class SqliteHistoryStore(HistoryStore):
    """``history.db`` SQLite backend."""

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
        # Forward-compatible migration: add 'mode' to DBs created before it.
        self._ensure_column(cur, 'runs', 'mode', 'TEXT')
        self._ensure_column(cur, 'orders', 'mode', 'TEXT')
        self._ensure_column(cur, 'orders', 'segment', 'TEXT')
        self.conn.commit()

    @staticmethod
    def _ensure_column(cur, table: str, col: str, decl: str) -> None:
        existing = [r[1] for r in
                    cur.execute(f"PRAGMA table_info({table})").fetchall()]
        if col not in existing:
            cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {decl}")

    def existing_pos(self) -> set:
        cur = self.conn.cursor()
        cur.execute("SELECT marketplace, po FROM orders")
        return {(m, p) for m, p in cur.fetchall()}

    def record(self, run_meta: dict, order_rows: List[dict]):
        cur = self.conn.cursor()
        cur.execute(
            "INSERT INTO runs (run_ts, mode, online_root, marketplaces, "
            "total_pos, total_items, total_qty, total_value, "
            "consolidated_path, tracker_path) VALUES (?,?,?,?,?,?,?,?,?,?)",
            (run_meta['run_ts'], run_meta.get('mode', ''),
             run_meta['online_root'], run_meta['marketplaces'],
             run_meta['total_pos'], run_meta['total_items'],
             run_meta['total_qty'], run_meta['total_value'],
             run_meta['consolidated_path'], run_meta['tracker_path']),
        )
        run_id = cur.lastrowid
        for o in order_rows:        # all rows are new (dedup already applied)
            cur.execute(
                "INSERT INTO orders (run_id, run_ts, mode, segment, "
                "marketplace, marketplace_label, po, location, warehouse, "
                "po_date, exp_date, order_type, items, qty, order_value, "
                "output_file) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (run_id, run_meta['run_ts'], run_meta.get('mode', ''),
                 o.get('segment', ORDER_SEGMENT),
                 o['marketplace'], o['marketplace_label'], o['po'],
                 o['location'], o['warehouse'], o['po_date'], o['exp_date'],
                 o['order_type'], o['items'], o['qty'], o['order_value'],
                 o['output_file']),
            )
        self.conn.commit()
        return run_id

    def fetch_orders(self, run_id=None) -> List[dict]:
        return _fetch_orders_sql(self.conn.cursor(), 'orders', '?', run_id)

    def export_to_xlsx(self, out_path) -> str:
        cur = self.conn.cursor()
        cur.execute(f"SELECT {_HISTORY_SELECT} FROM orders "
                    "ORDER BY run_id DESC, id ASC")
        return _write_history_xlsx(cur.fetchall(), out_path)

    def close(self) -> None:
        self.conn.close()


# ── MySQL implementation ────────────────────────────────────────────────

_MYSQL_RUNS_DDL = """
CREATE TABLE IF NOT EXISTS runs (
    run_id            BIGINT AUTO_INCREMENT PRIMARY KEY,
    run_ts            DATETIME,
    mode              ENUM('AUTO','MANUAL'),
    source            VARCHAR(500),
    marketplaces      INT,
    total_pos         INT,
    total_items       INT,
    total_qty         INT,
    total_value       DECIMAL(14,2),
    consolidated_path VARCHAR(500),
    tracker_path      VARCHAR(500),
    created_at        DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""

_MYSQL_ORDERS_DDL = """
CREATE TABLE IF NOT EXISTS order_headers (
    order_id          BIGINT AUTO_INCREMENT PRIMARY KEY,
    run_id            BIGINT,
    run_ts            DATETIME,
    mode              ENUM('AUTO','MANUAL'),
    segment           VARCHAR(20),
    marketplace       VARCHAR(50),
    marketplace_label VARCHAR(50),
    po                VARCHAR(100),
    location          VARCHAR(500),
    warehouse         VARCHAR(20),
    po_date           DATE NULL,
    exp_date          DATE NULL,
    order_type        ENUM('SO','TO'),
    items             INT,
    qty               INT,
    order_value       DECIMAL(14,2),
    output_file       VARCHAR(500),
    created_at        DATETIME DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_mp_po (marketplace, po),
    INDEX idx_po (po),
    INDEX idx_run_ts (run_ts)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""


class MySqlHistoryStore(HistoryStore):
    """
    MySQL backend — the production store. Same interface as the SQLite
    one, so Auto/Manual/GUI are unaffected by the swap. Tables:
    ``runs`` + ``order_headers`` (the operator's chosen header table; ready
    for a future ``order_lines`` child keyed on ``order_id``).
    """

    def __init__(self, cfg: dict) -> None:
        import pymysql                      # local import → optional dep
        self.conn = pymysql.connect(
            host=cfg.get('host', '127.0.0.1'),
            port=int(cfg.get('port', 3306)),
            user=cfg.get('user', 'root'),
            password=cfg.get('password', ''),
            database=cfg.get('database', 'renee_orders'),
            charset='utf8mb4',
            autocommit=False,
        )
        self._init_schema()

    def _init_schema(self) -> None:
        with self.conn.cursor() as cur:
            cur.execute(_MYSQL_RUNS_DDL)
            cur.execute(_MYSQL_ORDERS_DDL)
            # Forward-compatible migrations:
            self._ensure_column(cur, 'order_headers', 'segment',
                                "VARCHAR(20)", "AFTER mode")
            # Drop the old duplicate-tracking columns — the DB now holds
            # only new POs (no duplicate tracking). Rows are kept.
            self._drop_column(cur, 'order_headers', 'is_duplicate')
            self._drop_column(cur, 'order_headers', 'first_seen_ts')
        self.conn.commit()

    def _ensure_column(self, cur, table: str, col: str, decl: str,
                       after: str = '') -> None:
        if not self._has_column(cur, table, col):
            cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {decl} {after}")

    def _drop_column(self, cur, table: str, col: str) -> None:
        if self._has_column(cur, table, col):
            cur.execute(f"ALTER TABLE {table} DROP COLUMN {col}")

    @staticmethod
    def _has_column(cur, table: str, col: str) -> bool:
        cur.execute(
            "SELECT COUNT(*) FROM information_schema.columns WHERE "
            "table_schema=DATABASE() AND table_name=%s AND column_name=%s",
            (table, col))
        return cur.fetchone()[0] > 0

    def existing_pos(self) -> set:
        with self.conn.cursor() as cur:
            cur.execute("SELECT marketplace, po FROM order_headers")
            return {(m, p) for m, p in cur.fetchall()}

    def record(self, run_meta: dict, order_rows: List[dict]):
        with self.conn.cursor() as cur:
            cur.execute(
                "INSERT INTO runs (run_ts, mode, source, marketplaces, "
                "total_pos, total_items, total_qty, total_value, "
                "consolidated_path, tracker_path) "
                "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (_to_dt(run_meta['run_ts']), run_meta.get('mode'),
                 run_meta['online_root'], run_meta['marketplaces'],
                 run_meta['total_pos'], run_meta['total_items'],
                 run_meta['total_qty'], run_meta['total_value'],
                 run_meta['consolidated_path'], run_meta['tracker_path']))
            run_id = cur.lastrowid
            for o in order_rows:    # all rows are new (dedup already applied)
                cur.execute(
                    "INSERT INTO order_headers (run_id, run_ts, mode, segment, "
                    "marketplace, marketplace_label, po, location, warehouse, "
                    "po_date, exp_date, order_type, items, qty, order_value, "
                    "output_file) "
                    "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                    (run_id, _to_dt(run_meta['run_ts']), run_meta.get('mode'),
                     o.get('segment', ORDER_SEGMENT),
                     o['marketplace'], o['marketplace_label'], o['po'],
                     o['location'], o['warehouse'], _to_date(o['po_date']),
                     _to_date(o['exp_date']), o['order_type'], o['items'],
                     o['qty'], o['order_value'], o['output_file']))
        self.conn.commit()
        return run_id

    def fetch_orders(self, run_id=None) -> List[dict]:
        return _fetch_orders_sql(self.conn.cursor(), 'order_headers', '%s',
                                 run_id)

    def export_to_xlsx(self, out_path) -> str:
        with self.conn.cursor() as cur:
            cur.execute(f"SELECT {_HISTORY_SELECT} FROM order_headers "
                        "ORDER BY run_id DESC, order_id ASC")
            rows = cur.fetchall()
        return _write_history_xlsx(rows, out_path)

    def close(self) -> None:
        self.conn.close()


# ── Factory (single swap point for the backend) ─────────────────────────

def get_history_store(db_path=None) -> HistoryStore:
    """
    Return the configured history store. Reads the local db_config: when
    it selects ``backend='mysql'``, returns :class:`MySqlHistoryStore`;
    otherwise the SQLite store. This is the ONE place the backend is
    chosen — Auto, Manual, the runner and the GUI never change.
    """
    cfg = load_db_config()
    if cfg and str(cfg.get('backend', '')).lower() == 'mysql':
        try:
            return MySqlHistoryStore(cfg)
        except Exception as e:   # noqa: BLE001 — MySQL down? fall back to SQLite
            import logging
            logging.warning("MySQL history unavailable (%s) — falling back "
                            "to SQLite", e)
    return SqliteHistoryStore(db_path or default_history_db_path())


# ── Row building (shared by Auto + Manual) ──────────────────────────────

def order_rows_from_result(result, marketplace_key: str, warehouse: str,
                           output_file: str) -> List[dict]:
    """Per-PO order rows for one ``ProcessingResult`` (reuses the tracker
    row-builder for label / location / dates / inc-GST value)."""
    trk = {str(row['po']): row for row in build_tracker_rows(result)}
    items_by_po: Dict[str, int] = {}
    for so in result.rows:
        po = str(so.po_number)
        items_by_po[po] = items_by_po.get(po, 0) + 1
    otype = 'TO' if getattr(result, 'output_type', 'so') == 'to' else 'SO'

    out: List[dict] = []
    for po, t in trk.items():
        out.append({
            'segment': ORDER_SEGMENT,
            'marketplace': marketplace_key,
            'marketplace_label': t['market_place'],
            'po': po,
            'location': t['location'] or '',
            'warehouse': warehouse or '',
            'po_date': str(t['po_date']) if t['po_date'] else '',
            'exp_date': str(t['exp_date']) if t['exp_date'] else '',
            'order_type': otype,
            'items': items_by_po.get(po, 0),
            'qty': int(t['order_qty'] or 0),
            'order_value': float(t['order_value'] or 0.0),
            'output_file': output_file or '',
        })
    return out


def _record(order_rows: List[dict], run_meta: dict, db_path,
            skipped: int = 0) -> dict:
    # Nothing new → don't create an empty run row (DB holds only new POs).
    if not order_rows:
        return {'db_path': str(db_path), 'run_id': None,
                'new_orders': 0, 'skipped': skipped}
    store = get_history_store(db_path)
    try:
        run_id = store.record(run_meta, order_rows)
    finally:
        store.close()
    return {'db_path': str(db_path), 'run_id': run_id,
            'new_orders': len(order_rows), 'skipped': skipped}


# ── Dedup: remove already-uploaded POs from a result ────────────────────

def apply_dedup(result) -> List[dict]:
    """
    If ``DEDUP_SKIP_ENABLED``: drop every PO already present in the history
    DB from ``result.rows`` (so it never reaches Headers/Lines) and stash a
    summary of each removed PO on ``result.skipped_orders``. Returns that
    skipped list. The DB is *not* touched here — duplicates are simply not
    output and not recorded; they appear only on the "Skipped" output sheet.
    """
    result.skipped_orders = []
    if not DEDUP_SKIP_ENABLED or not result.rows:
        return []
    mp = result.marketplace
    store = get_history_store()
    try:
        existing = store.existing_pos()
    finally:
        store.close()
    dup_pos = {str(so.po_number) for so in result.rows
               if (mp, str(so.po_number)) in existing}
    if not dup_pos:
        return []

    # Summarise the skipped POs (qty / value / label / location / dates)
    # from the tracker rows BEFORE filtering them out.
    trk = {str(t['po']): t for t in build_tracker_rows(result)}
    skipped: List[dict] = []
    for po in dup_pos:
        t = trk.get(po, {})
        skipped.append({
            'segment': ORDER_SEGMENT,
            'marketplace': mp,
            'marketplace_label': t.get('market_place', mp),
            'po': po,
            'location': t.get('location', '') or '',
            'po_date': t.get('po_date', ''),
            'exp_date': t.get('exp_date', ''),
            'qty': int(t.get('order_qty') or 0),
            'order_value': float(t.get('order_value') or 0.0),
        })
    result.rows = [so for so in result.rows
                   if str(so.po_number) not in dup_pos]
    result.skipped_orders = skipped
    return skipped


# ── Public entry points (Auto + Manual) ─────────────────────────────────

def record_history(runs: List, online_root: str,
                   consolidated_path: str = '', tracker_path: str = '',
                   run_ts: Optional[str] = None) -> dict:
    """AUTO mode: append a whole batch run (list of MarketplaceRun). Only
    NEW POs are recorded (dedup already removed duplicates from each
    result)."""
    ok = [r for r in runs if r.status in ('ok', 'no_rows') and r.result]
    order_rows: List[dict] = []
    skipped = 0
    for run in ok:
        skipped += len(getattr(run.result, 'skipped_orders', []) or [])
        if run.result.rows:
            order_rows.extend(order_rows_from_result(
                run.result, run.marketplace, run.warehouse,
                run.output_path or ''))

    run_meta = {
        'run_ts': run_ts or datetime.now().isoformat(timespec='seconds'),
        'mode': 'AUTO',
        'online_root': str(online_root),
        'marketplaces': len({r.marketplace for r in ok if r.result.rows}),
        'total_pos': len({(o['marketplace'], o['po']) for o in order_rows}),
        'total_items': len(order_rows),
        'total_qty': sum(o['qty'] for o in order_rows),
        'total_value': sum(o['order_value'] for o in order_rows),
        'consolidated_path': str(consolidated_path or ''),
        'tracker_path': str(tracker_path or ''),
    }
    return _record(order_rows, run_meta, history_db_path(online_root),
                   skipped=skipped)


def record_manual(result, output_file: str = '',
                  run_ts: Optional[str] = None) -> dict:
    """MANUAL mode: append a single generated result to the SAME shared
    history. Only NEW POs are recorded (dedup already applied)."""
    rows = order_rows_from_result(
        result, result.marketplace,
        getattr(result, 'warehouse_display', '') or '', output_file)
    run_meta = {
        'run_ts': run_ts or datetime.now().isoformat(timespec='seconds'),
        'mode': 'MANUAL',
        'online_root': f"MANUAL: {os.path.basename(output_file)}"
                       if output_file else 'MANUAL',
        'marketplaces': 1,
        'total_pos': len({str(r.po_number) for r in result.rows}),
        'total_items': len(result.rows),
        'total_qty': sum(r.qty for r in result.rows),
        'total_value': sum(o['order_value'] for o in rows),
        'consolidated_path': '',
        'tracker_path': '',
    }
    return _record(rows, run_meta, default_history_db_path(),
                   skipped=len(getattr(result, 'skipped_orders', []) or []))
