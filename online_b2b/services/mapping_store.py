"""
online_b2b.services.mapping_store
=================================

Move the **Ship-To B2B mapping** (delivery location → ERP Cust No + Ship-to code)
off the bundled ``Calculation Data/Ship to B2B.xlsx`` and into the DB — a pure
*data-source shift*, mirroring the item-master shift (see ``item_master_loader``).

The frozen engine (``online_po_processor``) is NOT touched: it keeps calling
``mapping.load(...)`` / ``mapping.lookup(...)`` exactly as before. We feed it a
:class:`DBMappingLoader` that fills the SAME in-memory structures
(``self.mappings`` / ``self.by_shipto``) from MySQL instead of the Excel, so every
lookup tier (exact → normalized → aggressive → substring → by-code → Flipkart
address-overlap) is inherited unchanged.

Table ``ship_to_mapping`` holds one row per Excel row (all 25 parties — online
marketplaces AND offline channels), with the four resolution columns plus the
reference address columns. A full upload **replaces** the table (latest wins).
"""

from __future__ import annotations

import datetime as _dt
import json as _json
import re as _re

import pandas as pd
from online_po_processor.data.mapping_loader import MappingLoader

from .order_db import _conn

_MAP_TABLE = 'ship_to_mapping'
_FIELD_TABLE = 'ship_to_field'   # user-defined custom-field definitions

# Insert column order. ``source`` = 'excel' (from a bulk upload, wiped+rebuilt on
# every re-upload) or 'manual' (added/edited from the UI — DURABLE: survives an
# Excel re-upload, like item_master_manual).
_COLS = ['party', 'del_location', 'cust_no', 'ship_to', 'name', 'address',
         'address2', 'postcode', 'city', 'source', 'batch_id', 'updated_at']

_MYSQL_MAP = """
CREATE TABLE IF NOT EXISTS ship_to_mapping (
    id            BIGINT AUTO_INCREMENT PRIMARY KEY,
    party         VARCHAR(60),
    del_location  VARCHAR(500),
    cust_no       VARCHAR(40),
    ship_to       VARCHAR(60),
    name          VARCHAR(255),
    address       VARCHAR(500),
    address2      VARCHAR(500),
    postcode      VARCHAR(20),
    city          VARCHAR(120),
    state         VARCHAR(60) DEFAULT '',
    gst_reg       VARCHAR(60) DEFAULT '',
    country       VARCHAR(60) DEFAULT '',
    source        VARCHAR(10) DEFAULT 'excel',
    batch_id      VARCHAR(40),
    updated_at    DATETIME,
    INDEX idx_stm_party (party),
    INDEX idx_stm_shipto (ship_to)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
_SQLITE_MAP = """
CREATE TABLE IF NOT EXISTS ship_to_mapping (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    party         TEXT, del_location TEXT, cust_no TEXT, ship_to TEXT,
    name          TEXT, address TEXT, address2 TEXT, postcode TEXT, city TEXT,
    state         TEXT DEFAULT '', gst_reg TEXT DEFAULT '', country TEXT DEFAULT '',
    source        TEXT DEFAULT 'excel', batch_id TEXT, updated_at TEXT
)
"""


def _s(x) -> str:
    """NaN/None-safe stringify (a float NaN is truthy, so ``x or ''`` is unsafe)."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    s = str(x).strip()
    return '' if s.lower() == 'nan' else s


def _cust(x) -> str:
    """Cust No as the ERP stores it — pandas reads integer codes as floats when a
    cell in the column is blank ('20011.0'); strip the trailing '.0'."""
    c = _s(x)
    return c[:-2] if c.endswith('.0') else c


def ensure_table() -> None:
    """Create ``ship_to_mapping`` if absent (idempotent), and add the ``source``
    column to a pre-existing table. Web owns this table; engine schema untouched."""
    with _conn() as (cur, d):
        cur.execute(_MYSQL_MAP if d['kind'] == 'mysql' else _SQLITE_MAP)
        try:                              # backfill source col on an older table
            cur.execute(f"ALTER TABLE {_MAP_TABLE} ADD COLUMN source "
                        f"VARCHAR(10) DEFAULT 'excel'")
        except Exception:  # noqa: BLE001 — column already exists → fine
            pass
        # D365 address-enrichment columns (state / GST reg / country) — additive,
        # idempotent; the full authoritative address is loaded from the D365 dump.
        for col in ('state', 'gst_reg', 'country'):
            try:
                cur.execute(f"ALTER TABLE {_MAP_TABLE} ADD COLUMN {col} "
                            f"VARCHAR(60) DEFAULT ''")
            except Exception:  # noqa: BLE001 — already exists → fine
                pass
        # Personalization: per-row values for user-defined custom fields live in
        # this JSON blob ({field_name: value}); the field *definitions* live in
        # ship_to_field. Both are additive — nothing existing is touched.
        try:
            cur.execute(f"ALTER TABLE {_MAP_TABLE} ADD COLUMN extra TEXT")
        except Exception:  # noqa: BLE001 — already exists → fine
            pass
        if d['kind'] == 'mysql':
            cur.execute(f"""
                CREATE TABLE IF NOT EXISTS {_FIELD_TABLE} (
                    id         BIGINT AUTO_INCREMENT PRIMARY KEY,
                    name       VARCHAR(60) UNIQUE,
                    label      VARCHAR(120),
                    created_at DATETIME
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
        else:
            cur.execute(f"""
                CREATE TABLE IF NOT EXISTS {_FIELD_TABLE} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE, label TEXT, created_at TEXT
                )""")
        cur.connection.commit()


# ── Personalization: user-defined custom fields ─────────────────────────────
#   The operator can add extra columns (e.g. "Contact", "GST hold", "Region")
#   without a code change. Definitions live in ship_to_field; per-row values in
#   ship_to_mapping.extra (a JSON map). Reserved names can't collide with the
#   built-in columns.

_RESERVED_FIELDS = {
    'id', 'party', 'del_location', 'cust_no', 'ship_to', 'name', 'address',
    'address2', 'postcode', 'city', 'state', 'gst_reg', 'country', 'source',
    'batch_id', 'updated_at', 'extra',
}


def _slug(label) -> str:
    """A safe machine key from a human label ("GST Hold" → "gst_hold")."""
    s = _re.sub(r'[^a-z0-9]+', '_', str(label).strip().lower()).strip('_')
    return s[:60]


def _parse_extra(v) -> dict:
    """JSON blob → dict (never raises)."""
    if not v:
        return {}
    try:
        x = _json.loads(v)
        return x if isinstance(x, dict) else {}
    except Exception:  # noqa: BLE001
        return {}


def list_custom_fields() -> list:
    """Ordered list of custom fields ``[{'name','label'}, ...]``. Never raises."""
    try:
        ensure_table()
        with _conn() as (cur, d):
            cur.execute(f"SELECT name, label FROM {_FIELD_TABLE} ORDER BY id")
            return [{'name': n, 'label': l or n} for n, l in cur.fetchall()]
    except Exception:  # noqa: BLE001
        return []


def add_custom_field(label: str) -> dict:
    """Define a new custom field from a human label. Returns {ok, name, label}."""
    ensure_table()
    label = _s(label)[:120]
    name = _slug(label)
    if not name:
        return {'ok': False, 'error': 'Enter a field name.'}
    if name in _RESERVED_FIELDS:
        return {'ok': False, 'error': f'"{label}" is a built-in column already.'}
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"SELECT 1 FROM {_FIELD_TABLE} WHERE name={ph}", (name,))
        if cur.fetchone():
            return {'ok': False, 'error': 'That field already exists.'}
        cur.execute(f"INSERT INTO {_FIELD_TABLE} (name, label, created_at) "
                    f"VALUES ({ph}, {ph}, {ph})", (name, label, _dt.datetime.now()))
        cur.connection.commit()
    return {'ok': True, 'name': name, 'label': label}


def delete_custom_field(name: str) -> dict:
    """Remove a custom-field definition (per-row values in ``extra`` are left as
    orphaned JSON keys — harmless, and restored if the field is re-added)."""
    ensure_table()
    name = _slug(name)
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"DELETE FROM {_FIELD_TABLE} WHERE name={ph}", (name,))
        n = cur.rowcount
        cur.connection.commit()
    return {'ok': bool(n), 'deleted': n or 0}


def _extract_extra(fields: dict, known: set | None = None) -> dict:
    """Pull ``cf_<name>`` keys from a submitted form dict into an ``extra`` map,
    keeping only currently-defined fields with a non-blank value."""
    if known is None:
        known = {c['name'] for c in list_custom_fields()}
    out = {}
    for k, v in (fields or {}).items():
        if not k.startswith('cf_'):
            continue
        name = _slug(k[3:])
        if name in known:
            val = _s(v)[:500]
            if val:
                out[name] = val
    return out


# ── Parse (Ship-To B2B sheet → rows) ────────────────────────────────────────

def build_rows(xlsx_path: str):
    """Parse the ``Ship-To B2B`` sheet into mapping rows. Lenient on header
    naming (same synonyms the engine's loader accepts). Returns
    ``(rows, stats, warnings)`` and writes NOTHING."""
    warnings: list[str] = []
    try:
        try:
            df = pd.read_excel(xlsx_path, sheet_name='Ship-To B2B', header=0)
        except ValueError:
            df = pd.read_excel(xlsx_path, header=0)
            warnings.append("Sheet 'Ship-To B2B' not found — used the first sheet.")
    except Exception as e:  # noqa: BLE001
        return [], {}, [f"Cannot read mapping file: {type(e).__name__}: {e}"]

    cmap: dict = {}
    for col in df.columns:
        cl = str(col).strip().lower()
        if cl == 'party':
            cmap['party'] = col
        elif cl in ('del location', 'delivery location', 'location'):
            cmap['del_location'] = col
        elif cl in ('cust no', 'cust no.', 'customer no', 'sell-to'):
            cmap['cust_no'] = col
        elif cl in ('ship to', 'ship-to', 'ship to code'):
            cmap['ship_to'] = col
        elif cl == 'name':
            cmap['name'] = col
        elif cl in ('address', 'address 1', 'address1'):
            cmap['address'] = col
        elif cl in ('address 2', 'address2'):
            cmap['address2'] = col
        elif cl in ('postcode', 'post code', 'pincode', 'pin code', 'zip'):
            cmap['postcode'] = col
        elif cl == 'city':
            cmap['city'] = col

    missing = [k for k in ('party', 'del_location', 'cust_no', 'ship_to')
               if k not in cmap]
    if missing:
        return [], {}, [f"Mapping file missing columns: {', '.join(missing)}. "
                        f"Available: {list(df.columns)}"]

    def g(row, key):
        return _s(row[cmap[key]]) if key in cmap else ''

    rows, dropped = [], 0
    for _, r in df.iterrows():
        party = g(r, 'party')
        loc = g(r, 'del_location')
        if not party or not loc:        # rows with no party/location are noise
            dropped += 1
            continue
        rows.append({
            'party': party[:60], 'del_location': loc[:500],
            'cust_no': _cust(r[cmap['cust_no']])[:40],
            'ship_to': g(r, 'ship_to')[:60],
            'name': g(r, 'name')[:255], 'address': g(r, 'address')[:500],
            'address2': g(r, 'address2')[:500], 'postcode': g(r, 'postcode')[:20],
            'city': g(r, 'city')[:120],
        })

    by_party: dict = {}
    for r in rows:
        by_party[r['party']] = by_party.get(r['party'], 0) + 1
    stats = {'rows': len(rows), 'parties': len(by_party), 'dropped': dropped,
             'by_party': dict(sorted(by_party.items()))}
    if dropped:
        warnings.append(f"{dropped} row(s) skipped (blank Party or Del Location).")
    return rows, stats, warnings


# ── Write (full replace, transactional) ─────────────────────────────────────

def replace_mapping(rows: list) -> dict:
    """Wipe and rebuild ``ship_to_mapping`` from ``rows`` in one transaction —
    the table always mirrors the latest upload exactly."""
    ensure_table()
    batch = _dt.datetime.now().strftime('%Y%m%d%H%M%S')
    now = _dt.datetime.now()
    payload = [(
        r['party'], r['del_location'], r['cust_no'], r['ship_to'], r['name'],
        r['address'], r['address2'], r['postcode'], r['city'], 'excel', batch, now,
    ) for r in rows]
    with _conn() as (cur, d):
        ph = d['ph']
        cols = ', '.join(_COLS)
        marks = ', '.join([ph] * len(_COLS))
        # Wipe only the Excel-sourced rows — manual (UI-added) rows are durable.
        cur.execute(f"DELETE FROM {_MAP_TABLE} WHERE "
                    f"COALESCE(source,'excel') <> 'manual'")
        cur.executemany(
            f"INSERT INTO {_MAP_TABLE} ({cols}) VALUES ({marks})", payload)
        cur.connection.commit()
    return {'ok': True, 'rows': len(payload), 'batch_id': batch}


# ── CRUD (single rows; UI-added rows are source='manual', survive re-uploads) ─

_CRUD_FIELDS = ['party', 'del_location', 'cust_no', 'ship_to', 'name', 'address',
                'address2', 'postcode', 'city']


def _clean_fields(fields: dict) -> dict:
    out = {k: _s(fields.get(k))[:500] for k in _CRUD_FIELDS}
    out['cust_no'] = _cust(fields.get('cust_no'))[:40]
    out['party'] = out['party'][:60]
    out['ship_to'] = out['ship_to'][:60]
    out['postcode'] = out['postcode'][:20]
    out['city'] = out['city'][:120]
    out['name'] = out['name'][:255]
    return out


def add_mapping(fields: dict) -> dict:
    """Insert a single mapping row (source='manual'). Requires party +
    del_location. Returns {ok, id} or {ok:False, error}."""
    ensure_table()
    f = _clean_fields(fields)
    if not f['party'] or not f['del_location']:
        return {'ok': False, 'error': 'Party and Del Location are required.'}
    extra = _extract_extra(fields)
    cols = _CRUD_FIELDS + ['source', 'updated_at', 'extra']
    vals = [f[k] for k in _CRUD_FIELDS] + ['manual', _dt.datetime.now(),
                                           _json.dumps(extra)]
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(
            f"INSERT INTO {_MAP_TABLE} ({', '.join(cols)}) "
            f"VALUES ({', '.join([ph] * len(cols))})", vals)
        new_id = cur.lastrowid
        cur.connection.commit()
    return {'ok': True, 'id': new_id}


def update_mapping(row_id, fields: dict) -> dict:
    """Edit ONE mapping row by id. **Partial**: only the standard fields actually
    present in ``fields`` are written — so editing e.g. City never blanks the
    enriched Address — plus any ``cf_*`` custom values merged into ``extra``.
    Marks the row source='manual' so the edit survives an Excel re-upload."""
    ensure_table()
    fields = fields or {}
    present = [k for k in _CRUD_FIELDS if k in fields]
    clean = _clean_fields(fields)
    for k in ('party', 'del_location'):
        if k in fields and not clean[k]:
            return {'ok': False, 'error': 'Party and Del Location cannot be blank.'}
    submitted_extra = _extract_extra(fields)
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"SELECT extra FROM {_MAP_TABLE} WHERE id={ph}", (int(row_id),))
        row = cur.fetchone()
        merged = _parse_extra(row[0]) if row else {}
        merged.update(submitted_extra)          # incoming custom values win
        sets = present + ['source', 'updated_at', 'extra']
        vals = [clean[k] for k in present] + ['manual', _dt.datetime.now(),
                                              _json.dumps(merged)]
        cur.execute(
            f"UPDATE {_MAP_TABLE} SET {', '.join(f'{c}={ph}' for c in sets)} "
            f"WHERE id={ph}", vals + [int(row_id)])
        n = cur.rowcount
        cur.connection.commit()
    return {'ok': bool(n), 'updated': n or 0}


def delete_mapping(row_id) -> dict:
    """Delete a single mapping row by id."""
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"DELETE FROM {_MAP_TABLE} WHERE id={ph}", (int(row_id),))
        n = cur.rowcount
        cur.connection.commit()
    return {'ok': bool(n), 'deleted': n or 0}


def get_mapping(row_id) -> dict | None:
    """Fetch one mapping row by id (for the edit form)."""
    cols = ['id'] + _CRUD_FIELDS + ['source']
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"SELECT {', '.join(cols)} FROM {_MAP_TABLE} WHERE id={ph}",
                    (int(row_id),))
        r = cur.fetchone()
    return dict(zip(cols, r)) if r else None


def seed_from_bundled() -> dict:
    """One-time (re)seed of ``ship_to_mapping`` from the engine's bundled
    ``Ship to B2B.xlsx`` so the DB is live immediately. Returns the replace
    result (+ stats), or an error dict."""
    try:
        from .engine_bridge import _engine_imports
        path = _engine_imports()['get_bundled_mapping_path']()
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"Bundled mapping unavailable: {e}"}
    if not path:
        return {'ok': False, 'error': "Bundled Ship-To B2B mapping not found."}
    rows, stats, warnings = build_rows(str(path))
    if not rows:
        return {'ok': False, 'error': '; '.join(warnings) or 'No rows parsed.'}
    res = replace_mapping(rows)
    res.update(stats=stats, warnings=warnings, source=str(path))
    return res


# ── Read (status / overview) ────────────────────────────────────────────────

def table_count() -> int:
    """Number of rows in ship_to_mapping (0 if table absent)."""
    try:
        with _conn() as (cur, d):
            cur.execute(f"SELECT COUNT(*) FROM {_MAP_TABLE}")
            return int(cur.fetchone()[0] or 0)
    except Exception:  # noqa: BLE001
        return 0


def status() -> dict:
    """Snapshot for the status page: total rows, distinct parties, last update,
    per-party counts. Never raises."""
    try:
        ensure_table()
        with _conn() as (cur, d):
            cur.execute(f"SELECT COUNT(*), COUNT(DISTINCT party), MAX(updated_at) "
                        f"FROM {_MAP_TABLE}")
            n, parties, last = cur.fetchone()
            cur.execute(f"SELECT party, COUNT(*) FROM {_MAP_TABLE} "
                        f"GROUP BY party ORDER BY party")
            by_party = [{'party': p, 'count': int(c)} for p, c in cur.fetchall()]
        return {'ok': True, 'count': int(n or 0), 'parties': int(parties or 0),
                'last_updated': last, 'by_party': by_party}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"{type(e).__name__}: {e}",
                'count': 0, 'parties': 0, 'last_updated': None, 'by_party': []}


def list_mappings(party: str = '', q: str = '', limit: int = 200) -> dict:
    """Browsable overview: filter by party and/or search location/code/city.
    Returns ``{rows, total, shown, party, q, parties}``. Read-only."""
    party = (party or '').strip()
    q = (q or '').strip()
    cfields = list_custom_fields()   # personalization columns (fetched first)
    cols = ['id', 'party', 'del_location', 'cust_no', 'ship_to', 'name',
            'address', 'city', 'state', 'postcode', 'source', 'extra']
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(f"SELECT DISTINCT party FROM {_MAP_TABLE} ORDER BY party")
            parties = [r[0] for r in cur.fetchall()]
            where, args = [], []
            plist = [p.strip() for p in party.split(',') if p.strip()]  # multi-select
            if plist:
                marks = ', '.join([ph] * len(plist))
                where.append(f"party IN ({marks})"); args += plist
            if q:
                like = f"%{q}%"
                where.append(f"(del_location LIKE {ph} OR ship_to LIKE {ph} OR "
                             f"cust_no LIKE {ph} OR city LIKE {ph})")
                args += [like, like, like, like]
            wsql = ('WHERE ' + ' AND '.join(where)) if where else ''
            cur.execute(f"SELECT COUNT(*) FROM {_MAP_TABLE} {wsql}", args)
            total = int(cur.fetchone()[0] or 0)
            cur.execute(f"SELECT {', '.join(cols)} FROM {_MAP_TABLE} {wsql} "
                        f"ORDER BY party, del_location LIMIT {int(limit)}", args)
            rows = [dict(zip(cols, r)) for r in cur.fetchall()]
        # flatten each row's custom-field values into template-friendly cells
        for r in rows:
            ex = _parse_extra(r.pop('extra', ''))
            r['custom_cells'] = [{'name': c['name'], 'label': c['label'],
                                  'value': ex.get(c['name'], '')} for c in cfields]
        return {'rows': rows, 'total': total, 'shown': len(rows),
                'party': party, 'selected_parties': plist, 'q': q,
                'parties': parties, 'custom_fields': cfields}
    except Exception:  # noqa: BLE001
        return {'rows': [], 'total': 0, 'shown': 0, 'party': party,
                'selected_parties': [], 'q': q, 'parties': [],
                'custom_fields': cfields}


def export_rows(party: str = '', q: str = '') -> tuple[list[str], list[dict]]:
    """ALL mapping rows with FULL columns for the Excel export, honoring the same
    party/search filter as the on-screen list (no limit). Read-only. Returns
    ``(columns, rows)``."""
    party = (party or '').strip()
    q = (q or '').strip()
    cols = ['party', 'del_location', 'cust_no', 'ship_to', 'name', 'address',
            'address2', 'postcode', 'city', 'source']
    with _conn() as (cur, d):
        ph = d['ph']
        where, args = [], []
        plist = [p.strip() for p in party.split(',') if p.strip()]  # multi-select
        if plist:
            marks = ', '.join([ph] * len(plist))
            where.append(f"party IN ({marks})"); args += plist
        if q:
            like = f"%{q}%"
            where.append(f"(del_location LIKE {ph} OR ship_to LIKE {ph} OR "
                         f"cust_no LIKE {ph} OR city LIKE {ph} OR name LIKE {ph})")
            args += [like, like, like, like, like]
        wsql = ('WHERE ' + ' AND '.join(where)) if where else ''
        cur.execute(f"SELECT {', '.join(cols)} FROM {_MAP_TABLE} {wsql} "
                    f"ORDER BY party, del_location", args)
        rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    return cols, rows


# ── DB-backed loader (drop-in for the engine's MappingLoader) ────────────────

class DBMappingLoader(MappingLoader):
    """Fills the SAME structures the Excel ``MappingLoader`` builds
    (``self.mappings`` / ``self.by_shipto``) from ``ship_to_mapping`` instead of
    the workbook. Every lookup tier is inherited unchanged — the engine can't
    tell the difference."""

    def load(self, filepath, party_name, logs) -> int:  # noqa: ARG002 — filepath unused
        self.party_name = party_name
        self.mappings = {}
        self.by_shipto = {}

        def _norm_party(p: str) -> str:
            return ''.join(str(p).split()).lower()
        want = _norm_party(party_name)

        with _conn() as (cur, d):
            cur.execute(
                f"SELECT party, del_location, cust_no, ship_to FROM {_MAP_TABLE}")
            fetched = cur.fetchall()

        for party, location, cust_no, ship_to in fetched:
            if _norm_party(party) != want:
                continue
            location = (location or '').strip()
            cust_no = (cust_no or '').strip()
            ship_to = (ship_to or '').strip()
            if cust_no.endswith('.0'):
                cust_no = cust_no[:-2]
            if location and location.lower() != 'nan':
                entry = {'cust_no': cust_no, 'ship_to': ship_to}
                self.mappings[location] = entry
                if ship_to and ship_to.lower() != 'nan':
                    self.by_shipto.setdefault(
                        ship_to.upper(), {**entry, 'matched_key': ship_to})

        self.total_loaded = len(self.mappings)
        return self.total_loaded
