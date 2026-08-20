"""Availability **run recorder** — snapshot a fill-rate check so it can be
enquired into later. Inventory keeps changing, so a check run today is NOT
reproducible tomorrow; recording freezes the whole result (per-PO / per-SKU fill
+ the best-warehouse comparison) alongside the *inventory-as-of* it was computed
from. Self-contained + independently removable: ONE table, a handful of
functions, no edits to the availability engine. [[availability-checker]]
[[modular-removable-features]]

Stored in the business DB (Render's filesystem is ephemeral — a file wouldn't
survive a deploy). The full result is kept as a JSON payload so a past run can be
replayed exactly, read-only.
"""
from __future__ import annotations

import datetime as _dt
import json

from . import availability as av
from .order_db import _conn

_DDL_MYSQL = """
CREATE TABLE IF NOT EXISTS availability_run (
    run_id        BIGINT AUTO_INCREMENT PRIMARY KEY,
    run_ts        DATETIME,
    actor         VARCHAR(80),
    n_orders      INT DEFAULT 0,
    n_skus        INT DEFAULT 0,
    order_nos     TEXT,
    wh_override   VARCHAR(40),
    inv_as_of     VARCHAR(255),
    fill_pct      DECIMAL(6,2) DEFAULT 0,
    fill_val_pct  DECIMAL(6,2) DEFAULT 0,
    ord_qty       DECIMAL(16,2) DEFAULT 0,
    fillable_qty  DECIMAL(16,2) DEFAULT 0,
    best_wh       VARCHAR(20),
    note          VARCHAR(255),
    payload       MEDIUMTEXT,
    INDEX idx_availrun_ts (run_ts)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
_DDL_SQLITE = """
CREATE TABLE IF NOT EXISTS availability_run (
    run_id INTEGER PRIMARY KEY AUTOINCREMENT, run_ts TEXT, actor TEXT,
    n_orders INTEGER DEFAULT 0, n_skus INTEGER DEFAULT 0, order_nos TEXT,
    wh_override TEXT, inv_as_of TEXT, fill_pct REAL DEFAULT 0,
    fill_val_pct REAL DEFAULT 0, ord_qty REAL DEFAULT 0, fillable_qty REAL DEFAULT 0,
    best_wh TEXT, note TEXT, payload TEXT)
"""
_READY = False


def _ensure() -> None:
    global _READY
    if _READY:
        return
    with _conn() as (cur, d):
        cur.execute(_DDL_MYSQL if d['kind'] == 'mysql' else _DDL_SQLITE)
    _READY = True


def _inv_as_of(summary: dict) -> str:
    """One human string of the inventory snapshot timestamp(s) the check used —
    so a recorded run always says which stock it was measured against."""
    m = summary.get('wh_stock_as_of') or {}
    return ' | '.join(f"{k}: {v}" for k, v in m.items()) if m else str(summary.get('stock_as_of') or '')


def record(order_nos, warehouse_override='', actor='', note='') -> dict:
    """Run the check + best-WH comparison and FREEZE the whole result as a row.
    ``order_nos`` is a parsed list. Returns ``{ok, run_id, recorded_at, ...}``."""
    if not order_nos:
        return {'ok': False, 'error': 'Paste at least one order number to record.'}
    data = av.check_orders(order_nos, warehouse_override)
    if not data.get('ok') or not data.get('orders'):
        return {'ok': False, 'error': 'Nothing to record — no recognised orders.'}
    sc = av.wh_scenarios(order_nos)
    s = data.get('summary', {})
    now = _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    payload = json.dumps({
        'check': data, 'scenarios': sc if sc.get('ok') else None,
        'order_nos': order_nos, 'warehouse_override': warehouse_override,
        'recorded_at': now, 'actor': actor, 'note': note,
    }, default=str)
    _ensure()
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(
            f"INSERT INTO availability_run (run_ts, actor, n_orders, n_skus, "
            f"order_nos, wh_override, inv_as_of, fill_pct, fill_val_pct, ord_qty, "
            f"fillable_qty, best_wh, note, payload) VALUES "
            f"({ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph})",
            (now, actor, s.get('orders', 0), s.get('skus', 0),
             ', '.join(order_nos), warehouse_override, _inv_as_of(s),
             s.get('fill_pct', 0), s.get('fill_val_pct', 0), s.get('ord_qty', 0),
             s.get('fillable_qty', 0), (sc.get('best_wh') if sc.get('ok') else ''),
             note[:255], payload))
        # id of the row we just inserted (portable across MySQL/SQLite)
        cur.execute("SELECT MAX(run_id) FROM availability_run")
        rid = cur.fetchone()[0]
    return {'ok': True, 'run_id': rid, 'recorded_at': now,
            'n_orders': s.get('orders', 0), 'fill_pct': s.get('fill_pct', 0)}


def list_runs(limit: int = 100) -> list[dict]:
    """Recent recorded runs (newest first) — list-view metadata, no payload."""
    _ensure()
    out: list[dict] = []
    with _conn() as (cur, d):
        cur.execute(
            f"SELECT run_id, run_ts, actor, n_orders, n_skus, wh_override, "
            f"inv_as_of, fill_pct, fill_val_pct, ord_qty, fillable_qty, best_wh, note "
            f"FROM availability_run ORDER BY run_id DESC LIMIT {int(limit)}")
        for r in cur.fetchall():
            out.append({
                'run_id': r[0], 'run_ts': str(r[1] or ''), 'actor': str(r[2] or ''),
                'n_orders': r[3] or 0, 'n_skus': r[4] or 0,
                'wh_override': str(r[5] or ''), 'inv_as_of': str(r[6] or ''),
                'fill_pct': float(r[7] or 0), 'fill_val_pct': float(r[8] or 0),
                'ord_qty': float(r[9] or 0), 'fillable_qty': float(r[10] or 0),
                'best_wh': str(r[11] or ''), 'note': str(r[12] or '')})
    return out


def get_run(run_id) -> dict:
    """Full recorded snapshot (frozen check + scenarios) for a replay view."""
    _ensure()
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"SELECT run_ts, actor, note, payload FROM availability_run "
                    f"WHERE run_id={ph}", (run_id,))
        row = cur.fetchone()
    if not row:
        return {'ok': False, 'error': 'Recorded run not found.'}
    try:
        payload = json.loads(row[3]) if row[3] else {}
    except (ValueError, TypeError):
        payload = {}
    return {'ok': True, 'run_id': run_id, 'run_ts': str(row[0] or ''),
            'actor': str(row[1] or ''), 'note': str(row[2] or ''),
            'payload': payload}


def delete_run(run_id) -> dict:
    _ensure()
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"DELETE FROM availability_run WHERE run_id={ph}", (run_id,))
    return {'ok': True, 'run_id': run_id}
