"""
online_b2b.services.channel_map
===============================

Per-channel **SKU-code → item / EAN** lookup, for marketplaces that send only
THEIR own code (no EAN) — Swiggy today, Health & Glow and any future code-only
channel tomorrow. Generalises the old Swiggy-only ``item_swiggy_map`` into ONE
channel-scoped table::

    channel_sku_map(channel, sku_code, ean, item_no, source)

The engine resolves such a code through ``master.swiggy_sku`` (``{sku_code: ean}``);
``DBMasterLoader`` fills that dict from this table (channel='Swiggy'), so the
engine stays untouched. ``source='manual'`` rows survive a re-seed (durable, like
the old map's intent).
"""

from __future__ import annotations

import datetime as _dt

from online_po_processor.data.master_loader import MasterLoader

from .order_db import _conn

_clean = MasterLoader._clean_code
_TABLE = 'channel_sku_map'

_MYSQL = """
CREATE TABLE IF NOT EXISTS channel_sku_map (
    id         BIGINT AUTO_INCREMENT PRIMARY KEY,
    channel    VARCHAR(40),
    sku_code   VARCHAR(80),
    ean        VARCHAR(32),
    item_no    VARCHAR(50),
    source     VARCHAR(10) DEFAULT 'excel',
    updated_at DATETIME,
    INDEX idx_csm_chan (channel),
    INDEX idx_csm_code (channel, sku_code)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
_SQLITE = """
CREATE TABLE IF NOT EXISTS channel_sku_map (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    channel TEXT, sku_code TEXT, ean TEXT, item_no TEXT,
    source TEXT DEFAULT 'excel', updated_at TEXT
)
"""
_COLS = ['channel', 'sku_code', 'ean', 'item_no', 'source', 'updated_at']


def ensure_table() -> None:
    with _conn() as (cur, d):
        cur.execute(_MYSQL if d['kind'] == 'mysql' else _SQLITE)
        cur.connection.commit()


def channel_codes(channel: str) -> dict:
    """``{sku_code: ean}`` for a channel — the exact shape the engine's
    ``master.swiggy_sku`` expects. The EAN is resolved LIVE from item_master via
    item_no (so a rebuilt master's fresh EANs flow through, matching the old
    join-at-load behaviour), falling back to the stored ean. Empty on any error."""
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(
                f"SELECT c.sku_code, m.ean, c.ean FROM {_TABLE} c "
                f"LEFT JOIN item_master m ON m.item_no = c.item_no "
                f"WHERE c.channel={ph}", (channel,))
            out = {}
            for sku, live_ean, stored_ean in cur.fetchall():
                ean = live_ean if (live_ean and str(live_ean).strip()) else stored_ean
                if sku and ean and str(ean).strip() and str(ean).lower() != 'nan':
                    out[_clean(sku)] = _clean(ean)
            return out
    except Exception:  # noqa: BLE001
        return {}


def upsert_code(channel: str, sku_code, item_no=None, ean=None,
                source: str = 'manual') -> dict:
    """Add/replace ONE ``(channel, sku_code)`` row (durable by default). Used by
    the Item Master add/edit form so a typed Swiggy SKU lands here — the single
    source of truth for per-channel codes — instead of an item_master column."""
    sku = _clean(sku_code)
    if not channel or not sku:
        return {'ok': False, 'error': 'channel and sku_code are required.'}
    now = _dt.datetime.now()
    ensure_table()
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"DELETE FROM {_TABLE} WHERE channel={ph} AND sku_code={ph}",
                    (channel, sku))
        cur.execute(
            f"INSERT INTO {_TABLE} ({', '.join(_COLS)}) "
            f"VALUES ({', '.join([ph] * len(_COLS))})",
            (channel, sku, _clean(ean) or None, _clean(item_no) or None,
             source, now))
        cur.connection.commit()
    return {'ok': True, 'channel': channel, 'sku_code': sku}


def table_count() -> int:
    try:
        with _conn() as (cur, d):
            cur.execute(f"SELECT COUNT(*) FROM {_TABLE}")
            return int(cur.fetchone()[0] or 0)
    except Exception:  # noqa: BLE001
        return 0


def migrate_from_swiggy_map() -> dict:
    """One-time fold of the legacy ``item_swiggy_map`` (item_no → sku_code) +
    ``item_master`` (item_no → ean) into ``channel_sku_map`` as channel='Swiggy'.
    Idempotent — replaces the Excel-sourced Swiggy rows (manual rows kept)."""
    ensure_table()
    now = _dt.datetime.now()
    rows = []
    with _conn() as (cur, d):
        ph = d['ph']
        try:
            cur.execute(
                "SELECT s.item_no, s.swiggy_sku_code, m.ean "
                "FROM item_swiggy_map s "
                "LEFT JOIN item_master m ON m.item_no = s.item_no")
            src = cur.fetchall()
        except Exception:  # noqa: BLE001
            src = []
        for item_no, sku, ean in src:
            if sku and str(sku).strip() and str(sku).lower() != 'nan':
                rows.append(('Swiggy', _clean(sku), _clean(ean),
                             _clean(item_no), 'excel', now))
        cur.execute(f"DELETE FROM {_TABLE} WHERE channel='Swiggy' AND "
                    f"COALESCE(source,'excel') <> 'manual'")
        if rows:
            cur.executemany(
                f"INSERT INTO {_TABLE} ({', '.join(_COLS)}) "
                f"VALUES ({', '.join([ph] * len(_COLS))})", rows)
        cur.connection.commit()
    return {'ok': True, 'channel': 'Swiggy', 'rows': len(rows)}


def status() -> dict:
    """Per-channel counts + last update for the overview/admin. Never raises."""
    try:
        ensure_table()
        with _conn() as (cur, d):
            cur.execute(f"SELECT channel, COUNT(*), MAX(updated_at) FROM {_TABLE} "
                        f"GROUP BY channel ORDER BY channel")
            by_channel = [{'channel': c, 'count': int(n), 'last_updated': u}
                          for c, n, u in cur.fetchall()]
            cur.execute(f"SELECT COUNT(*) FROM {_TABLE}")
            total = int(cur.fetchone()[0] or 0)
        return {'ok': True, 'total': total, 'by_channel': by_channel}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"{type(e).__name__}: {e}"}


def list_codes(channel: str = '', q: str = '', limit: int = 300) -> dict:
    """Browsable list (filter by channel / search). Read-only."""
    cols = ['id', 'channel', 'sku_code', 'ean', 'item_no', 'source']
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(f"SELECT DISTINCT channel FROM {_TABLE} ORDER BY channel")
            channels = [r[0] for r in cur.fetchall()]
            where, args = [], []
            if channel:
                where.append(f"channel={ph}"); args.append(channel)
            if q:
                where.append(f"(sku_code LIKE {ph} OR ean LIKE {ph} OR "
                             f"item_no LIKE {ph})")
                args += [f"%{q}%", f"%{q}%", f"%{q}%"]
            wsql = ('WHERE ' + ' AND '.join(where)) if where else ''
            cur.execute(f"SELECT COUNT(*) FROM {_TABLE} {wsql}", args)
            total = int(cur.fetchone()[0] or 0)
            cur.execute(f"SELECT {', '.join(cols)} FROM {_TABLE} {wsql} "
                        f"ORDER BY channel, sku_code LIMIT {int(limit)}", args)
            rows = [dict(zip(cols, r)) for r in cur.fetchall()]
        return {'rows': rows, 'total': total, 'channels': channels,
                'channel': channel, 'q': q}
    except Exception:  # noqa: BLE001
        return {'rows': [], 'total': 0, 'channels': [], 'channel': channel, 'q': q}
