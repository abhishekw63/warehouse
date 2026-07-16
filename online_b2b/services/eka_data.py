"""
online_b2b.services.eka_data
============================

**EKA location registry — shifted Excel → DB** (mirrors the item-master and
ship-to-mapping moves). The desktop EKA constructor reads ``EKA_DATA.xlsx`` (one
row per airport / EBO / kiosk store: Bill-to, Ship-to, Location code, prefixes,
SO-number examples, margin, active flag). We relocate that whole sheet into the
``eka_data`` table so the web app owns the data; the desktop stays untouched for
now and we migrate the reader to the DB later, one step at a time.

Web-owned table in ``renee_orders`` — never touches the business order tables.
"""
from __future__ import annotations

from .order_db import _conn

_TABLE = 'eka_data'

# Excel header (as it appears in EKA_DATA.xlsx) → DB column. Order-independent;
# matched case/space-insensitively so a stray trailing space ('Bill to ') is fine.
_COL_MAP = {
    'desc': 'descr',
    'billto': 'bill_to',
    'shipto': 'ship_to',
    'location': 'location_code',
    'genbizpostinggroup': 'posting_group',
    'shortname': 'short_name',
    'prefix': 'prefix',
    'shortcode': 'short_code',
    'transfercode': 'transfer_code',
    'type': 'kind',
    'exampleregular': 'example_regular',
    'exampletester': 'example_tester',
    'status': 'status',
    'mrgnpct': 'margin_pct',
}

_COLS = ['descr', 'bill_to', 'ship_to', 'location_code', 'posting_group',
         'short_name', 'prefix', 'short_code', 'transfer_code', 'kind',
         'example_regular', 'example_tester', 'status', 'margin_pct']

_CREATE = f"""
CREATE TABLE IF NOT EXISTS {_TABLE} (
    id             INT AUTO_INCREMENT PRIMARY KEY,
    descr          VARCHAR(255),
    bill_to        VARCHAR(20),
    ship_to        VARCHAR(20),
    location_code  VARCHAR(60),
    posting_group  VARCHAR(40),
    short_name     VARCHAR(80),
    prefix         VARCHAR(10),
    short_code     VARCHAR(30),
    transfer_code  VARCHAR(60),
    kind           VARCHAR(20),
    example_regular VARCHAR(60),
    example_tester  VARCHAR(60),
    status         VARCHAR(12),
    margin_pct     DECIMAL(6,3),
    updated_at     DATETIME,
    INDEX idx_eka_short (short_name),
    INDEX idx_eka_loc (location_code)
)"""


def ensure_table() -> None:
    with _conn() as (cur, d):
        cur.execute(_CREATE)
        cur.connection.commit()


def _norm(h) -> str:
    return ''.join(ch for ch in str(h or '').lower() if ch.isalnum())


def load_from_excel(path: str, replace: bool = True) -> dict:
    """Read ``EKA_DATA.xlsx`` and (re)load every store row into ``eka_data``.

    ``replace=True`` truncates first so the table mirrors the sheet exactly
    (the sheet is the authoritative desktop registry today). Returns counts.
    """
    import datetime as _dt

    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    # position of each DB column in the sheet (via the normalised header map)
    pos = {}
    for i, h in enumerate(header):
        key = _COL_MAP.get(_norm(h))
        if key:
            pos[key] = i

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        # skip fully-blank rows; a store row must have a Short Name or Desc
        sn = r[pos['short_name']] if 'short_name' in pos and pos['short_name'] < len(r) else None
        ds = r[pos['descr']] if 'descr' in pos and pos['descr'] < len(r) else None
        if not sn and not ds:
            continue
        vals = []
        for c in _COLS:
            idx = pos.get(c)
            v = r[idx] if idx is not None and idx < len(r) else None
            if isinstance(v, str):
                v = v.strip()
            vals.append(v)
        rows.append(vals)

    ensure_table()
    now = _dt.datetime.now()
    with _conn() as (cur, d):
        ph = d['ph']
        if replace:
            cur.execute(f"DELETE FROM {_TABLE}")
        marks = ', '.join([ph] * (len(_COLS) + 1))
        cur.executemany(
            f"INSERT INTO {_TABLE} ({', '.join(_COLS)}, updated_at) VALUES ({marks})",
            [tuple(v) + (now,) for v in rows])
        cur.connection.commit()
    return {'loaded': len(rows), 'replaced': replace}


def all_rows() -> list[dict]:
    """Every EKA store as a JSON-safe dict (for the web page / future engine)."""
    ensure_table()
    with _conn() as (cur, d):
        cur.execute(f"SELECT {', '.join(_COLS)} FROM {_TABLE} ORDER BY short_name")
        return [dict(zip(_COLS, r)) for r in cur.fetchall()]


def active_rows() -> list[dict]:
    return [r for r in all_rows() if str(r.get('status', '')).lower() != 'inactive']
