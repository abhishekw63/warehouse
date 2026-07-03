"""
online_b2b.services.overrides_store
===================================

ONE unified **exceptions** table — every per-code override the engine reads:

  * **remap / CP override / vendor-CP** (from ``Master Exceptions.xlsx``): the
    Firstcry EAN-remap, Blink EPISENSE deal, Myntra Goddess 'use vendor CP'.
  * **swiggy_deal** (from ``Items March.xlsx`` → 'Swiggy Deal SKUs'): per-SKU
    negotiated deal prices.

They live in one table (``item_exceptions``) distinguished by a ``kind`` column.
This retires the last two bundled-Excel dependencies — the web is a single source
of truth (DB), no Excel files needed.

Parity strategy (zero drift): rows carry each sheet's raw cell values; on load we
split by ``kind`` and regenerate the engine's TWO expected sheets ('Exceptions' +
'Swiggy Deal SKUs') into a tiny in-memory workbook fed to the engine's OWN parsers
(``MasterLoader.load_exceptions`` / ``_load_swiggy_sheets``). The engine code that
interprets them is untouched, so DB-sourced overrides are byte-identical.
"""

from __future__ import annotations

import datetime as _dt
import os as _os
import tempfile as _tmp

import pandas as pd

from .order_db import _conn

_TABLE = 'item_exceptions'
KIND_EXC = 'exception'      # Master Exceptions row (remap / price / vendor_cp)
KIND_DEAL = 'swiggy_deal'   # Swiggy Deal SKUs row
KIND_MYNTRA_DEAL = 'myntra_deal'   # Myntra negotiated per-SKU transfer prices

# Unified columns. source_code = 'Source Code' (exception) OR 'EAN' (swiggy deal);
# override_mrp = 'Override MRP' OR 'Correct MRP'; note = 'Note' OR 'Name'.
_COLS = ['kind', 'source_code', 'maps_to', 'override_mrp', 'override_margin',
         'use_vendor_cp', 'marketplace', 'note', 'item_id', 'correct_gst',
         'cost_with_gst', 'cost_after_gst', 'source', 'updated_at']

# Excel-header ↔ DB-column maps, per sheet, used by both seed (read) and
# build_overlay_workbook (regenerate) → one definition, no drift.
_EXC_MAP = [
    ('source_code', 'Source Code'), ('maps_to', 'Maps To'),
    ('override_mrp', 'Override MRP'), ('override_margin', 'Override Margin %'),
    ('marketplace', 'Marketplace'), ('note', 'Note'),
    ('use_vendor_cp', 'Use Vendor CP'),
]
_DEAL_MAP = [
    ('item_id', 'Iteam ID'), ('source_code', 'EAN'), ('note', 'Name'),
    ('override_mrp', 'Correct MRP'), ('correct_gst', 'Correct GST'),
    ('cost_with_gst', 'Cost With GST'), ('cost_after_gst', 'Cost after GST'),
]
# Myntra negotiated-price sheet ('Myntra Deal SKU'): each SKU carries an agreed
# per-unit 'Cost With GST (Transfer Price)' that becomes the expected CP for
# Myntra POs (÷(1+GST) → pre-GST CP). Marketplace is stamped by the reader (the
# sheet has no Marketplace column). Same unified columns as the Swiggy deal.
_MYNTRA_DEAL_MAP = [
    ('item_id', 'Style ID'), ('source_code', 'EAN'), ('note', 'SKU Name'),
    ('override_mrp', 'MRP'), ('cost_with_gst', 'Cost With GST (Transfer Price)'),
]

_MYSQL = """
CREATE TABLE IF NOT EXISTS item_exceptions (
    id             BIGINT AUTO_INCREMENT PRIMARY KEY,
    kind           VARCHAR(16) DEFAULT 'exception',
    source_code    VARCHAR(80), maps_to VARCHAR(80),
    override_mrp   VARCHAR(40), override_margin VARCHAR(40),
    use_vendor_cp  VARCHAR(10), marketplace VARCHAR(60), note VARCHAR(500),
    item_id        VARCHAR(40), correct_gst VARCHAR(40),
    cost_with_gst  VARCHAR(40), cost_after_gst VARCHAR(40),
    source         VARCHAR(10) DEFAULT 'excel',
    updated_at     DATETIME,
    INDEX idx_iexc_kind (kind), INDEX idx_iexc_src (source_code)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
_SQLITE = """
CREATE TABLE IF NOT EXISTS item_exceptions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT DEFAULT 'exception', source_code TEXT, maps_to TEXT,
    override_mrp TEXT, override_margin TEXT, use_vendor_cp TEXT, marketplace TEXT,
    note TEXT, item_id TEXT, correct_gst TEXT, cost_with_gst TEXT,
    cost_after_gst TEXT, source TEXT DEFAULT 'excel', updated_at TEXT
)
"""


def _s(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    s = str(x).strip()
    return '' if s.lower() == 'nan' else s


def ensure_tables() -> None:
    with _conn() as (cur, d):
        cur.execute(_MYSQL if d['kind'] == 'mysql' else _SQLITE)
        # Upgrade a pre-existing (separate-tables era) item_exceptions to the
        # unified schema — add the columns it didn't have. Idempotent.
        for col, ddl in (
                ('kind', "VARCHAR(16) DEFAULT 'exception'"),
                ('item_id', 'VARCHAR(40)'), ('correct_gst', 'VARCHAR(40)'),
                ('cost_with_gst', 'VARCHAR(40)'), ('cost_after_gst', 'VARCHAR(40)')):
            try:
                cur.execute(f"ALTER TABLE {_TABLE} ADD COLUMN {col} {ddl}")
            except Exception:  # noqa: BLE001 — column already exists
                pass
        # legacy: the old separate swiggy-deals table is no longer used
        try:
            cur.execute("DROP TABLE IF EXISTS item_swiggy_deals")
        except Exception:  # noqa: BLE001
            pass
        cur.connection.commit()


def _read_sheet(xlsx_path, sheet, colmap, kind):
    """Read a sheet's raw cells into unified row-dicts tagged with ``kind``."""
    try:
        df = pd.read_excel(xlsx_path, sheet_name=sheet, header=0, dtype=str)
    except Exception:  # noqa: BLE001
        return []
    norm = {''.join(str(c).split()).lower(): c for c in df.columns}
    rows = []
    for _, r in df.iterrows():
        row = {c: '' for c in _COLS if c not in ('source', 'updated_at')}
        row['kind'] = kind
        for dbcol, header in colmap:
            src = norm.get(''.join(header.split()).lower())
            if src is not None:
                row[dbcol] = _s(r.get(src))
        if any(v for k, v in row.items() if k != 'kind'):
            rows.append(row)
    return rows


def replace_all(rows) -> int:
    """Replace ALL Excel-sourced rows (both kinds); manual rows are kept."""
    ensure_tables()
    now = _dt.datetime.now()
    payload = [tuple(r.get(c, '') for c in _COLS[:-2]) + ('excel', now) for r in rows]
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"DELETE FROM {_TABLE} WHERE COALESCE(source,'excel') <> 'manual'")
        if payload:
            cur.executemany(
                f"INSERT INTO {_TABLE} ({', '.join(_COLS)}) "
                f"VALUES ({', '.join([ph] * len(_COLS))})", payload)
        cur.connection.commit()
    return len(payload)


def _compilation_path():
    """The Online B2B dump compilation workbook — the operator's live source for
    the marketplace deal sheets (e.g. 'Myntra Deal SKU'). Overridable via the
    ``B2B_DEAL_COMPILATION`` Django setting / env var; falls back to the known
    OneDrive location. Returns None if not found (deal seeding is then skipped)."""
    import os
    try:
        from django.conf import settings
        cand = getattr(settings, 'B2B_DEAL_COMPILATION', None)
    except Exception:  # noqa: BLE001
        cand = None
    cand = cand or os.environ.get('B2B_DEAL_COMPILATION')
    default = r'D:/OneDrive - RENEE COSMETICS PRIVATE LIMITED/Online_B2B_Dump_Compilation.xlsx'
    for p in (cand, default):
        if p and os.path.exists(p):
            return p
    return None


def seed_from_bundled() -> dict:
    """Seed the unified item_exceptions from ALL sources: Master Exceptions +
    Swiggy Deal SKUs (bundled master) + Myntra Deal SKU (dump compilation)."""
    try:
        from pathlib import Path

        from .engine_bridge import _engine_imports
        master = _engine_imports()['get_bundled_master_path']()
        exc_path = Path(master).parent / 'Master Exceptions.xlsx'
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"Bundled sources unavailable: {e}"}
    rows = []
    if exc_path.exists():
        rows += _read_sheet(str(exc_path), 0, _EXC_MAP, KIND_EXC)
    rows += _read_sheet(str(master), 'Swiggy Deal SKUs', _DEAL_MAP, KIND_DEAL)
    # Myntra negotiated deal SKUs — read from the dump compilation the operator
    # maintains (the sheet isn't in the bundled master). Stamp marketplace so the
    # override applies to Myntra ONLY (never leaks to other channels).
    myn = 0
    comp = _compilation_path()
    if comp:
        mrows = _read_sheet(comp, 'Myntra Deal SKU', _MYNTRA_DEAL_MAP, KIND_MYNTRA_DEAL)
        for r in mrows:
            r['marketplace'] = 'Myntra'
        rows += mrows
        myn = len(mrows)
    n = replace_all(rows)
    exc = sum(1 for r in rows if r['kind'] == KIND_EXC)
    deals = sum(1 for r in rows if r['kind'] == KIND_DEAL)
    return {'ok': True, 'rows': n, 'exceptions': exc, 'swiggy_deals': deals,
            'myntra_deals': myn}


def _fetch(where=''):
    try:
        with _conn() as (cur, d):
            cur.execute(f"SELECT {', '.join(_COLS[:-2])}, source FROM {_TABLE} {where}")
            cols = _COLS[:-2] + ['source']
            return [dict(zip(cols, r)) for r in cur.fetchall()]
    except Exception:  # noqa: BLE001
        return []


def table_count() -> int:
    try:
        with _conn() as (cur, d):
            cur.execute(f"SELECT COUNT(*) FROM {_TABLE}")
            return int(cur.fetchone()[0] or 0)
    except Exception:  # noqa: BLE001
        return 0


def table_counts() -> dict:
    rows = _fetch()
    return {'exceptions': sum(1 for r in rows if r.get('kind') == KIND_EXC),
            'swiggy_deals': sum(1 for r in rows if r.get('kind') == KIND_DEAL),
            'myntra_deals': sum(1 for r in rows if r.get('kind') == KIND_MYNTRA_DEAL),
            'total': len(rows)}


def myntra_deal_map() -> dict:
    """``{clean EAN: agreed Cost With GST (Transfer Price)}`` for the Myntra deal
    SKUs — the per-SKU negotiated price used as the expected CP on Myntra POs."""
    out: dict = {}
    for r in _fetch(f"WHERE kind='{KIND_MYNTRA_DEAL}'"):
        ean = _s(r.get('source_code'))
        if ean.endswith('.0'):
            ean = ean[:-2]
        try:
            v = float(_s(r.get('cost_with_gst')))
        except (TypeError, ValueError):
            continue
        if ean and v > 0:
            out[ean] = v
    return out


def build_overlay_workbook() -> str | None:
    """Regenerate the engine's two overlay sheets from the ONE table (split by
    ``kind``) so its own parsers interpret them identically. Returns a temp path
    or None when empty; caller deletes the file."""
    rows = _fetch()
    if not rows:
        return None
    exc = [r for r in rows if r.get('kind') == KIND_EXC]
    deals = [r for r in rows if r.get('kind') == KIND_DEAL]
    exc_df = pd.DataFrame(
        [{h: r.get(c, '') for c, h in _EXC_MAP} for r in exc],
        columns=[h for _, h in _EXC_MAP])
    deal_df = pd.DataFrame(
        [{h: r.get(c, '') for c, h in _DEAL_MAP} for r in deals],
        columns=[h for _, h in _DEAL_MAP])
    fd, path = _tmp.mkstemp(suffix='.xlsx', prefix='ovl_')
    _os.close(fd)
    try:
        with pd.ExcelWriter(path, engine='openpyxl') as xw:
            exc_df.to_excel(xw, sheet_name='Exceptions', index=False)   # sheet 0
            deal_df.to_excel(xw, sheet_name='Swiggy Deal SKUs', index=False)
    except Exception:  # noqa: BLE001
        try:
            _os.remove(path)
        except OSError:
            pass
        return None
    return path


def status() -> dict:
    try:
        ensure_tables()
        c = table_counts()
        with _conn() as (cur, d):
            cur.execute(f"SELECT MAX(updated_at) FROM {_TABLE}")
            last = cur.fetchone()[0]
        return {'ok': True, 'last_updated': last, **c}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"{type(e).__name__}: {e}"}


def list_all() -> list:
    return _fetch("ORDER BY kind, source_code")
