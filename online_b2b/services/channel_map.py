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
import logging

from online_po_processor.data.master_loader import MasterLoader

from .order_db import _conn

_log = logging.getLogger(__name__)

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


_READY = False        # process-local: the fixed DDL only needs to run ONCE


def ensure_table() -> None:
    global _READY
    if _READY:
        return
    with _conn() as (cur, d):
        cur.execute(_MYSQL if d['kind'] == 'mysql' else _SQLITE)
        cur.connection.commit()
    _READY = True


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
        _log.exception('channel_codes(%s) failed — returning empty map', channel)
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


def _norm_col(c) -> str:
    return ''.join(ch for ch in str(c).strip().lower() if ch.isalnum())


# Candidate column names (normalised) for the vendor SKU code + the EAN, so ONE
# uploader handles every channel's master layout (HG 'sku_code'/'ENN code',
# Swiggy 'SkuCode'/'EAN', …) without per-channel code.
_SKU_COLS = ('skucode', 'sku', 'vendorsku', 'vendorcode', 'code', 'itemcode')
_EAN_COLS = ('enncode', 'ean', 'eancode', 'gtin', 'barcode', 'eanupccode')


def load_master_file(path: str, channel: str) -> dict:
    """Bulk-upsert a channel's **SKU→EAN** master (any .xlsx/.csv, all sheets)
    into ``channel_sku_map``. Auto-detects the vendor-SKU-code + EAN columns by
    name, resolves ``item_no`` from item_master, keeps manual rows, and reports
    unmapped EANs. ``{ok, parsed, inserted, updated, unmapped, error}``."""
    import datetime as _dt

    import pandas as pd
    if not channel:
        return {'ok': False, 'error': 'Pick a channel first.'}
    ensure_table()
    try:
        pairs: dict = {}                      # {sku_code: ean}, first non-blank wins
        if str(path).lower().endswith('.csv'):
            frames = [pd.read_csv(path, dtype=str)]
        else:
            frames = []
            for sh in pd.ExcelFile(path).sheet_names:
                raw = pd.read_excel(path, sheet_name=sh, header=None, dtype=str)
                hdr = None
                for i in range(min(6, len(raw))):
                    norm = {_norm_col(c) for c in raw.iloc[i]
                            if c is not None and str(c).strip()}
                    if (norm & set(_SKU_COLS)) and (norm & set(_EAN_COLS)):
                        hdr = i
                        break
                if hdr is not None:
                    frames.append(pd.read_excel(path, sheet_name=sh, header=hdr, dtype=str))
        for df in frames:
            low = {_norm_col(c): c for c in df.columns}
            sc = next((low[k] for k in _SKU_COLS if k in low), None)
            en = next((low[k] for k in _EAN_COLS if k in low), None)
            if not sc or not en:
                continue
            for _, r in df.iterrows():
                sku = _clean(r[sc])
                ean = _clean(r[en])
                if sku and ean and ean.lower() != 'nan' and sku not in pairs:
                    pairs[sku] = ean
        if not pairs:
            return {'ok': False, 'error': 'No sku_code + EAN columns detected in the file.'}
        now = _dt.datetime.now()
        ins = upd = unmapped = 0
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute('SELECT ean, item_no FROM item_master')
            e2i = {str(a): str(b) for a, b in cur.fetchall()}
            for sku, ean in pairs.items():
                item = e2i.get(ean, '')
                if not item:
                    unmapped += 1
                cur.execute(f"SELECT id FROM {_TABLE} WHERE channel={ph} AND sku_code={ph}",
                            (channel, sku))
                row = cur.fetchone()
                if row:
                    cur.execute(f"UPDATE {_TABLE} SET ean={ph}, item_no={ph}, "
                                f"source='upload', updated_at={ph} WHERE id={ph}",
                                (ean, item or None, now, row[0]))
                    upd += 1
                else:
                    cur.execute(f"INSERT INTO {_TABLE} (channel,sku_code,ean,item_no,"
                                f"source,updated_at) VALUES ({ph},{ph},{ph},{ph},"
                                f"'upload',{ph})", (channel, sku, ean, item or None, now))
                    ins += 1
            cur.connection.commit()
        return {'ok': True, 'parsed': len(pairs), 'inserted': ins,
                'updated': upd, 'unmapped': unmapped}
    except Exception as e:  # noqa: BLE001
        _log.exception('load_master_file failed')
        return {'ok': False, 'error': f"{type(e).__name__}: {e}"}


def delete_code(row_id) -> dict:
    """Delete ONE ``channel_sku_map`` row by id."""
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(f"DELETE FROM {_TABLE} WHERE id={ph}", (int(row_id),))
            n = cur.rowcount
            cur.connection.commit()
        return {'ok': bool(n), 'deleted': n or 0}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': str(e)}


def table_count() -> int:
    try:
        with _conn() as (cur, d):
            cur.execute(f"SELECT COUNT(*) FROM {_TABLE}")
            return int(cur.fetchone()[0] or 0)
    except Exception:  # noqa: BLE001
        _log.exception('table_count failed — returning 0')
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
            _log.exception('migrate_from_swiggy_map: source read failed — treating as empty')
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
        _log.exception('list_codes failed — returning empty result')
        return {'rows': [], 'total': 0, 'channels': [], 'channel': channel, 'q': q}
