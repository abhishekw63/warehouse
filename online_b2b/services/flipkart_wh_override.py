"""
online_b2b.services.flipkart_wh_override
========================================

ADDITIVE our-layer override for the Flipkart **Origin-Warehouse → sub-marketplace**
label (FK / FK Hyperlocal / FK Grocery).

The frozen engine (:data:`online_po_processor.engine.flipkart_tracker.LOCATION_MARKETPLACE`)
is the LOCKED base map; any warehouse missing there resolves to ``'FK (review)'``.
Rather than edit the frozen file for every new FC, operators **promote** a
warehouse here — a small DB-backed, durable map. An override **wins** over the
frozen result for its warehouse (so a ``'FK (review)'`` — or any label — can be
lifted to the confirmed one). Mirrors the DMart-FC / Reliance-ship-to pattern:
the frozen engine is never touched; our layer injects at runtime.
"""
from __future__ import annotations

import datetime as _dt

from .order_db import _conn

_TABLE = 'flipkart_wh_map'

_CREATE = f"""
CREATE TABLE IF NOT EXISTS {_TABLE} (
    origin_warehouse VARCHAR(120) NOT NULL PRIMARY KEY,
    market_place     VARCHAR(60)  NOT NULL,
    source           VARCHAR(20)  DEFAULT 'manual',
    updated_at       DATETIME NULL
)
"""


def ensure_table() -> None:
    with _conn() as (cur, d):
        cur.execute(_CREATE)
        cur.connection.commit()


def _norm(s) -> str:
    # The engine keys on the exact (stripped) Origin Warehouse string, which is
    # already lower_snake (e.g. 'hos_bag_wh_nl_01nl'); normalise defensively.
    return str(s or '').strip().lower()


def overrides() -> dict:
    """``{origin_warehouse(lower): market_place}`` — the promoted labels."""
    ensure_table()
    with _conn() as (cur, d):
        cur.execute(f"SELECT origin_warehouse, market_place FROM {_TABLE}")
        return {_norm(w): m for w, m in cur.fetchall()}


def set_override(origin_warehouse: str, market_place: str,
                 source: str = 'manual') -> dict:
    """Promote (upsert) one warehouse → market place. Idempotent."""
    wh = _norm(origin_warehouse)
    mkt = str(market_place or '').strip()[:60]
    if not wh or not mkt:
        return {'ok': False, 'error': 'Origin warehouse and market place are required.'}
    ensure_table()
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(
            f"UPDATE {_TABLE} SET market_place={ph}, source={ph}, updated_at={ph} "
            f"WHERE origin_warehouse={ph}", (mkt, source, _dt.datetime.now(), wh))
        if cur.rowcount == 0:
            cur.execute(
                f"INSERT INTO {_TABLE} (origin_warehouse, market_place, source, "
                f"updated_at) VALUES ({ph},{ph},{ph},{ph})",
                (wh, mkt, source, _dt.datetime.now()))
        cur.connection.commit()
    return {'ok': True, 'origin_warehouse': wh, 'market_place': mkt}


def apply(rows: list) -> int:
    """Re-label tracker rows **in place** from the override map (keyed by each
    row's ``'Location'`` = the origin-warehouse code). Override wins over the
    frozen label. Returns the number of rows changed. Never raises."""
    try:
        ov = overrides()
    except Exception:  # noqa: BLE001 — a label refinement must never break the run
        return 0
    if not ov:
        return 0
    n = 0
    for r in rows:
        m = ov.get(_norm(r.get('Location')))
        if m and r.get('Market Place') != m:
            r['Market Place'] = m
            n += 1
    return n
