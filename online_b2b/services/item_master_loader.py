"""
online_b2b.services.item_master_loader
=======================================

Move the **Item Master** off the bundled Excel and into the DB — a pure
*data-source shift*. The frozen engine (``online_po_processor``) is NOT touched:
it keeps calling ``master.lookup(...)`` exactly as before; we just feed it a
:class:`DBMasterLoader` that fills the same in-memory dict from MySQL instead of
``Items March.xlsx``.

The master is *calculated* from two ERP exports (no manual Excel curation):

  * **Items**  (``Items (6).xlsx``) — ``No.``, ``GTIN`` (EAN), ``Description``,
    ``GST Group Code``, ``HSN/SAC Code`` (+ UoM / Brand / Category).
  * **Item M.R.P.** (``Item M.R.P. (5).xlsx``) — ``Item No.`` → ``M.R.P.`` with
    ``Start Date`` / ``End Date`` validity windows.

Build rule (confirmed with the operator):
  * **Effective MRP = the period whose [Start, End] covers *today*** (the upload
    day). Older periods are discarded. → one MRP per item.
  * Join to Items by item No. for the rest of the attributes.

Per-channel SKU codes (Swiggy / Health & Glow / …) live in their OWN table now —
``channel_sku_map`` (see :mod:`channel_map`) — so item_master no longer carries a
``swiggy_sku_code`` column (that was duplicated data). Hand-added items are kept
durably IN ``item_master`` itself, flagged ``batch_id='manual'`` (no separate
overlay table); a full ERP rebuild preserves them and the source wins once the
ERP export starts carrying that item.

Never-skip-silently: items with no MRP window covering today, or an MRP row with
no matching Items row, are returned as **warnings** (not dropped quietly).
"""

from __future__ import annotations

import datetime as _dt

import pandas as pd

# Reuse the engine's identifier cleaner so EAN / Item-No keys match the master
# byte-for-byte (no spurious trailing '.0', scientific notation, etc.).
from online_po_processor.data.master_loader import MasterLoader

from .order_db import _conn

_clean = MasterLoader._clean_code


def _s(x) -> str:
    """NaN/None-safe stringify (a float NaN is truthy, so ``x or ''`` is unsafe).
    Also drops the literal 'nan' that ``astype(str)`` leaves on blank cells."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    s = str(x).strip()
    return '' if s.lower() == 'nan' else s

_MASTER_TABLE = 'item_master'

# Insert column order for item_master. Per-channel SKU codes (Swiggy/HG/…) are
# NOT stored here — they live in channel_sku_map (the single source of truth).
# Hand-added rows are flagged batch_id='manual' (durable; survive a full ERP
# rebuild) — no separate overlay table.
_COLS = ['item_no', 'ean', 'description', 'gst_code', 'hsn', 'mrp',
         'mrp_start', 'mrp_end', 'base_uom', 'brand',
         'category', 'batch_id', 'updated_at']

_MYSQL_MASTER = """
CREATE TABLE IF NOT EXISTS item_master (
    item_no          VARCHAR(50) PRIMARY KEY,
    ean              VARCHAR(32),
    description      VARCHAR(512),
    gst_code         VARCHAR(20),
    hsn              VARCHAR(20),
    mrp              DECIMAL(14,2),
    mrp_start        DATE,
    mrp_end          DATE,
    base_uom         VARCHAR(20),
    brand            VARCHAR(60),
    category         VARCHAR(100),
    batch_id         VARCHAR(40),
    updated_at       DATETIME,
    INDEX idx_im_ean (ean)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
_SQLITE_MASTER = """
CREATE TABLE IF NOT EXISTS item_master (
    item_no          TEXT PRIMARY KEY,
    ean              TEXT,
    description      TEXT,
    gst_code         TEXT,
    hsn              TEXT,
    mrp              REAL,
    mrp_start        TEXT,
    mrp_end          TEXT,
    base_uom         TEXT,
    brand            TEXT,
    category         TEXT,
    batch_id         TEXT,
    updated_at       TEXT
)
"""


def ensure_tables() -> None:
    """Create item_master if absent (idempotent) + the channel SKU map
    (Swiggy/HG/…). Hand-added items live IN item_master (batch_id='manual') — no
    separate overlay table; the old item_swiggy_map / item_master_manual are
    retired. Web owns these tables; the engine's schema is untouched."""
    with _conn() as (cur, d):
        mysql = d['kind'] == 'mysql'
        cur.execute(_MYSQL_MASTER if mysql else _SQLITE_MASTER)
        cur.connection.commit()
    from . import channel_map
    channel_map.ensure_table()


def _num(x):
    """Float (2-dp) or None for blank/NaN/non-numeric."""
    try:
        v = float(str(x).replace(',', '').strip())
        return None if v != v else round(v, 2)
    except (TypeError, ValueError):
        return None


def _date(x):
    """Parse a date-ish value to ``datetime.date`` or None."""
    if x is None or not str(x).strip():
        return None
    dt = pd.to_datetime(x, errors='coerce')
    return None if pd.isna(dt) else dt.date()


def _upsert(cur, d, table, cols, values) -> None:
    """INSERT … ON DUPLICATE KEY UPDATE (MySQL) / INSERT OR REPLACE (SQLite),
    treating ``cols[0]`` as the primary key."""
    ph = d['ph']
    marks = ', '.join([ph] * len(cols))
    collist = ', '.join(cols)
    if d['kind'] == 'mysql':
        sets = ', '.join(f"{c}=VALUES({c})" for c in cols[1:])
        cur.execute(f"INSERT INTO {table} ({collist}) VALUES ({marks}) "
                    f"ON DUPLICATE KEY UPDATE {sets}", values)
    else:
        cur.execute(f"INSERT OR REPLACE INTO {table} ({collist}) "
                    f"VALUES ({marks})", values)


def upsert_manual_item(data: dict) -> dict:
    """Add/overwrite ONE item by hand — writes the live ``item_master`` row
    flagged ``batch_id='manual'`` so it survives a full ERP rebuild (until the
    ERP export carries it, then the source row wins). A typed Swiggy SKU is
    routed to ``channel_sku_map`` (the per-channel source of truth), not an
    item_master column. ``item_no`` is required."""
    ensure_tables()
    item_no = _clean(data.get('item_no'))
    if not item_no:
        return {'ok': False, 'error': 'Item No is required.'}
    rec = {
        'item_no': item_no,
        'ean': _clean(data.get('ean')) or None,
        'description': _s(data.get('description'))[:512],
        'gst_code': _s(data.get('gst_code'))[:20],
        'hsn': _s(data.get('hsn'))[:20],
        'mrp': _num(data.get('mrp')),
        'mrp_start': _date(data.get('mrp_start')),
        'mrp_end': _date(data.get('mrp_end')),
    }
    swiggy_code = _s(data.get('swiggy_sku_code'))[:80]
    now = _dt.datetime.now()
    with _conn() as (cur, d):
        _upsert(cur, d, _MASTER_TABLE, _COLS,
                (rec['item_no'], rec['ean'], rec['description'], rec['gst_code'],
                 rec['hsn'], rec['mrp'], rec['mrp_start'], rec['mrp_end'],
                 '', '', '', 'manual', now))
        cur.connection.commit()
    # Swiggy SKU code lives in channel_sku_map now (per-channel source of truth).
    if swiggy_code:
        from . import channel_map
        channel_map.upsert_code('Swiggy', swiggy_code, item_no=item_no,
                                ean=rec['ean'], source='manual')
    return {'ok': True, 'item_no': item_no}


# ── Source-file readers ─────────────────────────────────────────────────────

def _read_items(path: str) -> pd.DataFrame:
    """Items export → frame keyed by clean Item No. Keeps the master attribute
    columns; tolerant of column-name drift (matched case/space-insensitively)."""
    df = pd.read_excel(path, sheet_name=0)
    cols = {''.join(str(c).split()).lower(): c for c in df.columns}

    def col(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return None
    out = pd.DataFrame()
    out['item_no'] = df[col('no.', 'no', 'itemno')].map(_clean)
    out['ean'] = df[col('gtin', 'ean', 'barcode')].map(_clean)
    out['description'] = df[col('description', 'itemdescription')].astype(str)
    out['gst_code'] = df[col('gstgroupcode', 'gstcode')].astype(str)
    hsn = col('hsn/saccode', 'hsn/sac', 'hsncode', 'hsn')
    out['hsn'] = df[hsn].map(_clean) if hsn else ''
    uom = col('baseunitofmeasure', 'baseuom', 'uom')
    out['base_uom'] = df[uom].astype(str) if uom else ''
    brand = col('brandcode', 'brand')
    out['brand'] = df[brand].astype(str) if brand else ''
    cat = col('catagory', 'category')
    out['category'] = df[cat].astype(str) if cat else ''
    return out.drop_duplicates('item_no', keep='first').set_index('item_no')


def _read_effective_mrp(path: str, as_of: _dt.date):
    """Item M.R.P. export → ``{item_no: {mrp, start, end}}`` keeping ONLY the
    period whose [Start, End] covers ``as_of`` (older/newer periods discarded).
    Returns ``(effective, warnings)``; warns for items with no covering window
    (falls back to the latest already-started period so the item is never lost).
    """
    df = pd.read_excel(path, sheet_name=0)
    cols = {''.join(str(c).split()).lower(): c for c in df.columns}

    def col(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return None
    c_item = col('itemno.', 'itemno', 'no.')
    c_mrp = col('m.r.p.', 'mrp')
    c_start = col('startdate', 'start')
    c_end = col('enddate', 'end')

    df = df[[c_item, c_mrp, c_start, c_end]].copy()
    df.columns = ['item_no', 'mrp', 'start', 'end']
    df['item_no'] = df['item_no'].map(_clean)
    df['start'] = pd.to_datetime(df['start'], errors='coerce').dt.date
    df['end'] = pd.to_datetime(df['end'], errors='coerce').dt.date
    df['mrp'] = pd.to_numeric(df['mrp'], errors='coerce')

    # Blank Start/End cells parse to NaT, which is TRUTHY — a bare ``if
    # r['start']`` lets NaT through and ``NaT <= today`` raises a TypeError.
    # Guard with pd.notna() wherever a date is tested or compared.
    def _has(d):
        return pd.notna(d)
    effective: dict = {}
    no_cover: list = []
    no_end: list = []   # effective period missing its End Date (open-ended)
    for item_no, grp in df.groupby('item_no'):
        rows = [r for _, r in grp.iterrows()]
        covering = [r for r in rows
                    if _has(r['start']) and _has(r['end'])
                    and r['start'] <= as_of <= r['end']]
        if covering:
            r = max(covering, key=lambda x: x['start'])
        else:
            started = [r for r in rows if _has(r['start']) and r['start'] <= as_of]
            if started:
                r = max(started, key=lambda x: x['start'])
                no_cover.append(item_no)
            else:
                r = min(rows, key=lambda x: x['start'] if _has(x['start'])
                        else _dt.date.max)
                no_cover.append(item_no)
        if not _has(r['end']):
            no_end.append(item_no)
        effective[item_no] = {'mrp': r['mrp'],
                              'start': r['start'] if _has(r['start']) else None,
                              'end': r['end'] if _has(r['end']) else None}

    warnings = []
    if no_cover:
        warnings.append(
            f"{len(no_cover)} item(s) had no MRP period covering {as_of} — used "
            f"the latest already-started price instead (e.g. {', '.join(no_cover[:8])}"
            f"{'…' if len(no_cover) > 8 else ''}).")
    if no_end:
        warnings.append(
            f"{len(no_end)} item(s) have an MRP period with a BLANK End Date — "
            f"treated as open-ended (never expires); the ERP normally fills a "
            f"far-future end (e.g. 31-03-2030), so verify these aren't a data slip "
            f"(e.g. {', '.join(no_end[:8])}{'…' if len(no_end) > 8 else ''}).")
    return effective, warnings


# ── Build (join the two sources) ────────────────────────────────────────────

def build_rows(items_path: str, mrp_path: str, as_of: _dt.date | None = None):
    """Compute the item-master rows from the two source files. The master is
    MRP-driven (one row per item that has an effective MRP), joined to Items for
    attributes. Returns ``(rows, stats, warnings)`` and writes NOTHING."""
    as_of = as_of or _dt.date.today()
    items = _read_items(items_path)
    effective, warnings = _read_effective_mrp(mrp_path, as_of)
    # {item_no: sku} from channel_sku_map (channel='Swiggy') — for the count stat
    # only; the codes themselves live in channel_sku_map, not an item_master col.
    swiggy = load_swiggy_map()

    rows = []
    missing_item = []
    for item_no, mi in effective.items():
        if item_no not in items.index:
            missing_item.append(item_no)
            continue
        it = items.loc[item_no]
        ean = _s(it['ean'])
        rows.append({
            'item_no': item_no,
            'ean': ean or None,
            'description': _s(it['description'])[:512],
            'gst_code': _s(it['gst_code']),
            'hsn': _s(it['hsn']),
            'mrp': None if pd.isna(mi['mrp']) else round(float(mi['mrp']), 2),
            'mrp_start': mi['start'],
            'mrp_end': mi['end'],
            'base_uom': _s(it['base_uom'])[:20],
            'brand': _s(it['brand'])[:60],
            'category': _s(it['category'])[:100],
        })
    if missing_item:
        warnings.append(
            f"{len(missing_item)} MRP item(s) not found in the Items file — "
            f"skipped (e.g. {', '.join(missing_item[:8])}"
            f"{'…' if len(missing_item) > 8 else ''}).")

    gst_spread: dict = {}
    for r in rows:
        gst_spread[r['gst_code']] = gst_spread.get(r['gst_code'], 0) + 1
    stats = {
        'as_of': str(as_of),
        'items': len(rows),
        'with_ean': sum(1 for r in rows if r['ean']),
        'swiggy_mapped': sum(1 for r in rows if swiggy.get(r['item_no'])),
        'no_mrp_window': sum('no MRP period' in w for w in warnings),
        'gst_spread': gst_spread,
    }
    return rows, stats, warnings


# ── Swiggy map (durable; seeded once from the legacy master) ─────────────────

def load_swiggy_map() -> dict:
    """``{item_no: swiggy_sku_code}`` for the Swiggy channel. The durable source
    is now ``channel_sku_map`` (channel='Swiggy') — the old ``item_swiggy_map`` is
    retired. Empty if not seeded yet."""
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute("SELECT item_no, sku_code FROM channel_sku_map "
                        f"WHERE channel={ph}", ('Swiggy',))
            return {_clean(i): (s or '').strip()
                    for i, s in cur.fetchall() if s and str(s).strip()}
    except Exception:  # noqa: BLE001 — table may not exist yet
        return {}


def seed_swiggy_from_excel(master_xlsx: str) -> dict:
    """One-time seed of the Swiggy channel in ``channel_sku_map`` from the legacy
    master's curated 'Item Master' sheet 'Swiggy Code' column (``No.`` → ``Swiggy
    Code``). Idempotent (manual rows kept); safe to re-run. Returns ``{'seeded': n}``."""
    ensure_tables()
    try:
        xl = pd.ExcelFile(master_xlsx)
        sheet = next((s for s in xl.sheet_names
                      if str(s).strip().lower() == 'item master'), 0)
        df = pd.read_excel(xl, sheet_name=sheet)
    except Exception as e:  # noqa: BLE001
        return {'seeded': 0, 'error': f"{type(e).__name__}: {e}"}
    cols = {''.join(str(c).split()).lower(): c for c in df.columns}
    c_no = cols.get('no.') or cols.get('no') or cols.get('itemno')
    c_sw = cols.get('swiggycode') or cols.get('swiggysku') or cols.get('swiggyskucode')
    if not c_no or not c_sw:
        return {'seeded': 0, 'error': 'No Swiggy Code column on the Item Master sheet.'}
    pairs = []
    for _, r in df.iterrows():
        ino = _clean(r.get(c_no))
        sku = _clean(r.get(c_sw))
        if ino and sku and sku.lower() != 'nan':
            pairs.append((ino, sku))
    # Seed the Swiggy channel into channel_sku_map (the durable code→item source).
    now = _dt.datetime.now()
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute("DELETE FROM channel_sku_map WHERE channel='Swiggy' AND "
                    "COALESCE(source,'excel') <> 'manual'")
        cur.executemany(
            "INSERT INTO channel_sku_map (channel, sku_code, item_no, source, "
            f"updated_at) VALUES ('Swiggy',{ph},{ph},'excel',{ph})",
            [(s, i, now) for i, s in pairs])
        cur.connection.commit()
    return {'seeded': len(pairs)}


# ── Write (full replace, transactional) ─────────────────────────────────────

def replace_item_master(rows: list) -> dict:
    """Rebuild ``item_master`` from ``rows`` (one ERP upload) in a transaction,
    PRESERVING hand-added rows (``batch_id='manual'``): only the previous ERP rows
    are cleared — manual rows stay in place. If the ERP export now carries a
    manual item_no, the upsert overwrites it (source wins once it appears).
    Per-channel SKU codes are NOT here — they live in channel_sku_map."""
    ensure_tables()
    batch = _dt.datetime.now().strftime('%Y%m%d%H%M%S')
    now = _dt.datetime.now()
    payload = [(
        r['item_no'], r['ean'], r['description'], r['gst_code'], r['hsn'],
        r['mrp'], r['mrp_start'], r['mrp_end'],
        r['base_uom'], r['brand'], r['category'], batch, now,
    ) for r in rows]
    cols = ', '.join(_COLS)
    with _conn() as (cur, d):
        ph = d['ph']
        marks = ', '.join([ph] * len(_COLS))
        # Clear only previous ERP rows; manual rows (batch_id='manual') survive.
        cur.execute(
            f"DELETE FROM {_MASTER_TABLE} WHERE COALESCE(batch_id,'') <> 'manual'")
        # Upsert so an ERP row overwrites a surviving manual row of the same
        # item_no (source wins once the export carries it).
        if d['kind'] == 'mysql':
            sets = ', '.join(f"{c}=VALUES({c})" for c in _COLS[1:])
            sql = (f"INSERT INTO {_MASTER_TABLE} ({cols}) VALUES ({marks}) "
                   f"ON DUPLICATE KEY UPDATE {sets}")
        else:
            sql = (f"INSERT OR REPLACE INTO {_MASTER_TABLE} ({cols}) "
                   f"VALUES ({marks})")
        cur.executemany(sql, payload)
        cur.execute(f"SELECT COUNT(*) FROM {_MASTER_TABLE}")
        total = int(cur.fetchone()[0] or 0)
        cur.execute(f"SELECT COUNT(*) FROM {_MASTER_TABLE} WHERE "
                    f"COALESCE(batch_id,'')='manual'")
        manual = int(cur.fetchone()[0] or 0)
        cur.connection.commit()
    return {'ok': True, 'rows': total, 'batch_id': batch,
            'manual_overlaid': manual}


def status() -> dict:
    """Current item_master snapshot for the status page: row count, last update,
    Swiggy-mapped count. Never raises."""
    try:
        ensure_tables()
        with _conn() as (cur, d):
            cur.execute(f"SELECT COUNT(*), MAX(updated_at) FROM {_MASTER_TABLE}")
            n, last = cur.fetchone()
            # Swiggy-mapped count comes from the per-channel map now (the
            # item_master.swiggy_sku_code column is gone).
            cur.execute("SELECT COUNT(*) FROM channel_sku_map "
                        "WHERE channel='Swiggy'")
            smap = int(cur.fetchone()[0] or 0)
        return {'ok': True, 'count': int(n or 0), 'last_updated': last,
                'swiggy_mapped': smap, 'swiggy_map_rows': smap}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"{type(e).__name__}: {e}",
                'count': 0, 'last_updated': None, 'swiggy_mapped': 0}


def list_items(q: str = '', limit: int = 100) -> dict:
    """Browsable overview of item_master: optional search across item_no / EAN /
    description. Returns ``{rows, total, shown, q}``. Read-only."""
    q = (q or '').strip()
    cols = ['item_no', 'ean', 'description', 'mrp', 'gst_code', 'hsn',
            'mrp_start', 'mrp_end']
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            where, args = '', []
            if q:
                like = f"%{q}%"
                where = (f"WHERE item_no LIKE {ph} OR ean LIKE {ph} OR "
                         f"description LIKE {ph}")
                args = [like, like, like]
            cur.execute(f"SELECT COUNT(*) FROM {_MASTER_TABLE} {where}", args)
            total = int(cur.fetchone()[0] or 0)
            cur.execute(
                f"SELECT {', '.join(cols)} FROM {_MASTER_TABLE} {where} "
                f"ORDER BY item_no LIMIT {int(limit)}", args)
            rows = [dict(zip(cols, r)) for r in cur.fetchall()]
        return {'rows': rows, 'total': total, 'shown': len(rows), 'q': q}
    except Exception:  # noqa: BLE001
        return {'rows': [], 'total': 0, 'shown': 0, 'q': q}


def export_rows(q: str = '') -> list:
    """ALL item_master rows (optionally filtered by the same search as
    :func:`list_items`) for a full export — no row limit. Read-only."""
    q = (q or '').strip()
    cols = ['item_no', 'ean', 'description', 'mrp', 'gst_code', 'hsn',
            'mrp_start', 'mrp_end']
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            where, args = '', []
            if q:
                like = f"%{q}%"
                where = (f"WHERE item_no LIKE {ph} OR ean LIKE {ph} OR "
                         f"description LIKE {ph}")
                args = [like, like, like]
            cur.execute(
                f"SELECT {', '.join(cols)} FROM {_MASTER_TABLE} {where} "
                f"ORDER BY item_no", args)
            return [dict(zip(cols, r)) for r in cur.fetchall()]
    except Exception:  # noqa: BLE001
        return []


def resolve_in_master(key) -> dict | None:
    """Look up ``key`` (an EAN or an Item No) in item_master. Returns
    ``{item_no, ean, description, mrp}`` on hit, else None. Used to validate an
    operator's 'correct EAN' before accepting it as an EAN fix."""
    k = _clean(key)
    if not k:
        return None
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(
                f"SELECT item_no, ean, description, mrp, gst_code FROM "
                f"{_MASTER_TABLE} WHERE ean={ph} OR item_no={ph} LIMIT 1", (k, k))
            r = cur.fetchone()
            if r:
                return {'item_no': r[0], 'ean': r[1], 'description': r[2],
                        'mrp': r[3], 'gst_code': r[4]}
    except Exception:  # noqa: BLE001
        pass
    return None


def table_count() -> int:
    """Row count of item_master (0 if the table is missing/empty). Used by the
    bridge to decide DB-master vs Excel-fallback."""
    try:
        with _conn() as (cur, d):
            cur.execute(f"SELECT COUNT(*) FROM {_MASTER_TABLE}")
            return int(cur.fetchone()[0] or 0)
    except Exception:  # noqa: BLE001
        return 0


# ── DB-backed master (engine reads this instead of the Excel) ────────────────

class DBMasterLoader(MasterLoader):
    """Drop-in for the engine's ``MasterLoader`` that fills the SAME in-memory
    structures from ``item_master`` instead of ``Items March.xlsx``. Inherits
    every lookup / pricing / exception method unchanged — the engine can't tell
    the difference. Overlay files (Master Exceptions / Swiggy deal sheets) are
    still loaded from the bundled workbook so nothing regresses."""

    def load_from_db(self, overlay_master_path: str | None = None):
        self.master = {}
        self.swiggy_sku = {}
        with _conn() as (cur, d):
            cur.execute(
                f"SELECT item_no, ean, description, gst_code, hsn, mrp "
                f"FROM {_MASTER_TABLE}")
            for item_no, ean, desc, gst, hsn, mrp in cur.fetchall():
                ino = _clean(item_no)
                entry = {
                    'item_no': ino,
                    'mrp': float(mrp) if mrp is not None else None,
                    'gst_code': gst or '',
                    'description': desc or '',
                    'hsn': hsn or '',
                }
                if ean:
                    self.master[_clean(ean)] = entry
                if ino and ino not in self.master:
                    self.master[ino] = entry

        # Per-channel SkuCode→EAN (Swiggy today, HG/others tomorrow) comes WHOLLY
        # from channel_sku_map now — the item_master swiggy_sku_code column is
        # gone. The EAN is resolved live from item_master via item_no inside
        # channel_codes(), so a rebuilt master's fresh EANs flow through.
        try:
            from . import channel_map
            self.swiggy_sku = channel_map.channel_codes('Swiggy')
        except Exception:  # noqa: BLE001 — leave swiggy_sku empty on any error
            pass

        # Pricing-override overlays (Master Exceptions + Swiggy deal SKUs) now
        # come from the DB too (overrides_store) — NO bundled Excel. We regenerate
        # a tiny workbook from the DB tables and feed it to the engine's OWN
        # parsers, so the result is byte-identical to the old Excel path (parity-
        # verified). ``overlay_master_path`` is accepted but unused. The curated DB
        # Swiggy SkuCode map is re-asserted on top so it wins.
        db_sku = dict(self.swiggy_sku)
        try:
            import os as _os

            from . import overrides_store
            wb = overrides_store.build_overlay_workbook()
            if wb:
                try:
                    self.load_exceptions(wb)            # 'Exceptions' sheet (0)
                    xl = pd.ExcelFile(wb)
                    try:
                        self._load_swiggy_sheets(xl)    # 'Swiggy Deal SKUs' sheet
                    finally:
                        xl.close()
                finally:
                    try:
                        _os.remove(wb)
                    except OSError:
                        pass
        except Exception as _e:  # noqa: BLE001 — overlays must never break the load,
            # but never SILENTLY: a failed deal-overlay load means deal SKUs can
            # price at flat margin instead of the negotiated price — log it loudly.
            import logging as _lg
            _lg.getLogger(__name__).warning(
                "Deal-SKU overlays (exceptions + Swiggy deal sheets) failed to load: "
                "%s — deal SKUs may fall back to flat margin. Check overrides_store.", _e)
        self.swiggy_sku.update(db_sku)

        # Historical EAN corrections (received_ean → correct EAN) feed the
        # engine's alias step, so a repeat wrong EAN auto-resolves. Derived from
        # the validation layer — no separate alias table.
        try:
            from . import lines_store
            for wrong, correct in lines_store.ean_alias_map().items():
                self.exceptions[_clean(wrong)] = _clean(correct)
        except Exception:  # noqa: BLE001
            pass
        return self

    def add_session_aliases(self, ean_fixes: dict) -> None:
        """Add the current upload's pending EAN fixes ({wrong → correct}) on top
        of the loaded master so this PO re-resolves before it's locked."""
        for wrong, correct in (ean_fixes or {}).items():
            w, c = _clean(wrong), _clean(correct)
            if w and c:
                self.exceptions[w] = c

    def count(self) -> int:
        return len(self.master)


# ── update diff + staleness (for the upload preview + Hub reminder) ───────
def diff_against_current(rows: list) -> dict:
    """Compare incoming item rows vs the LIVE item_master → what will actually
    change: **new** items (item_no absent), **mrp_changed** (MRP differs, old →
    new), **removed** (a current non-manual item not in the new file). Read-only;
    never raises. So the operator sees exactly what an update touches — and a
    clear 'nothing to update' when nothing differs."""
    cur_map: dict = {}
    try:
        ensure_tables()
        with _conn() as (cur, _d):
            cur.execute(f"SELECT item_no, mrp, description, "
                        f"COALESCE(batch_id,'') FROM {_MASTER_TABLE}")
            for item_no, mrp, desc, batch in cur.fetchall():
                cur_map[str(item_no)] = {
                    'mrp': None if mrp is None else round(float(mrp), 2),
                    'description': desc or '', 'manual': batch == 'manual'}
    except Exception:  # noqa: BLE001
        return {'ok': False, 'new': [], 'mrp_changed': [], 'removed': [],
                'counts': {'new': 0, 'mrp_changed': 0, 'removed': 0,
                           'unchanged': 0}, 'any': False}
    new, changed, seen = [], [], set()
    for r in rows:
        ino = str(r.get('item_no') or '')
        if not ino:
            continue
        seen.add(ino)
        nm = None if r.get('mrp') is None else round(float(r['mrp']), 2)
        cur_row = cur_map.get(ino)
        if cur_row is None:
            new.append({'item_no': ino, 'new_mrp': nm,
                        'description': (r.get('description') or '')[:80]})
        elif cur_row['mrp'] != nm:
            changed.append({'item_no': ino, 'old_mrp': cur_row['mrp'],
                            'new_mrp': nm,
                            'description': (r.get('description') or '')[:80]})
    removed = [{'item_no': k, 'old_mrp': v['mrp'],
                'description': v['description'][:80]}
               for k, v in cur_map.items() if k not in seen and not v['manual']]
    unchanged = max(0, len(seen) - len(new) - len(changed))
    counts = {'new': len(new), 'mrp_changed': len(changed),
              'removed': len(removed), 'unchanged': unchanged}
    return {'ok': True, 'new': new[:500], 'mrp_changed': changed[:500],
            'removed': removed[:500], 'counts': counts,
            'any': bool(new or changed or removed)}


def last_updated() -> dict:
    """``{'when', 'days', 'due'}`` from ``MAX(updated_at)`` — powers the Hub
    '15-day refresh' reminder. ``due`` is True at ≥ 15 days. Never raises."""
    import datetime as _d
    try:
        ensure_tables()
        with _conn() as (cur, _d2):
            cur.execute(f"SELECT MAX(updated_at) FROM {_MASTER_TABLE}")
            last = cur.fetchone()[0]
    except Exception:  # noqa: BLE001
        return {'when': None, 'days': None, 'due': False}
    if not last:
        return {'when': None, 'days': None, 'due': True}
    if isinstance(last, str):
        dt = None
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d'):
            try:
                dt = _d.datetime.strptime(last[:19], fmt)
                break
            except ValueError:
                continue
        if dt is None:
            return {'when': str(last)[:10], 'days': None, 'due': False}
    else:
        dt = last
    days = (_d.datetime.now() - dt).days
    return {'when': dt.strftime('%d %b %Y'), 'days': days, 'due': days >= 15}
