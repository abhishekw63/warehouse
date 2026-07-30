"""
online_b2b.services.tracker_store
=================================

**Manual rows for the Consolidated Tracker** — for POs that can't be uploaded
through the web app but still need tracking. Stored in ONE new, isolated,
web-owned table (``tracker_manual``); it touches nothing existing (no business/
core table or logic). The tracker view merges these with the auto rows so the
page stays a single source of truth.

Self-contained and removable: drop this module + its table and the auto tracker
is unaffected.
"""

from __future__ import annotations

import datetime as _dt

from .order_db import _conn

_MYSQL = """
CREATE TABLE IF NOT EXISTS tracker_manual (
    id           INT AUTO_INCREMENT PRIMARY KEY,
    dept         VARCHAR(20),
    warehouse    VARCHAR(60),
    marketplace  VARCHAR(80),
    po           VARCHAR(120),
    external_doc VARCHAR(120),
    location     VARCHAR(255),
    pincode      VARCHAR(12),
    zone         VARCHAR(20),
    po_date      DATE NULL,
    exp_date     DATE NULL,
    order_value  DECIMAL(16,2) DEFAULT 0,
    qty          INT DEFAULT 0,
    omt          VARCHAR(255),
    created_by   VARCHAR(80),
    created_at   DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
_SQLITE = """
CREATE TABLE IF NOT EXISTS tracker_manual (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dept TEXT, warehouse TEXT, marketplace TEXT, po TEXT, external_doc TEXT,
    location TEXT, pincode TEXT, zone TEXT, po_date TEXT, exp_date TEXT,
    order_value REAL DEFAULT 0, qty INTEGER DEFAULT 0, omt TEXT,
    created_by TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP
)
"""

_FIELDS = ('dept', 'warehouse', 'marketplace', 'po', 'external_doc', 'location',
           'pincode', 'zone', 'po_date', 'exp_date', 'order_value', 'qty', 'omt')


def ensure_table() -> None:
    with _conn() as (cur, d):
        cur.execute(_MYSQL if d['kind'] == 'mysql' else _SQLITE)
        cur.connection.commit()


def _date(v):
    v = (str(v or '')).strip()
    if not v:
        return None
    for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y'):
        try:
            return _dt.datetime.strptime(v, fmt).date()
        except ValueError:
            continue
    return None


def add(data: dict, user: str = '') -> dict:
    """Insert one manual tracker row. ``po`` is required; other fields optional."""
    ensure_table()
    po = str(data.get('po') or '').strip()
    if not po:
        return {'ok': False, 'error': 'PO is required.'}
    try:
        vals = {
            'dept': str(data.get('dept') or '').strip()[:20],
            'warehouse': str(data.get('warehouse') or '').strip()[:60],
            'marketplace': str(data.get('marketplace') or '').strip()[:80],
            'po': po[:120],
            'external_doc': str(data.get('external_doc') or '').strip()[:120],
            'location': str(data.get('location') or '').strip()[:255],
            'pincode': str(data.get('pincode') or '').strip()[:12],
            'zone': str(data.get('zone') or '').strip().upper()[:20],
            'po_date': _date(data.get('po_date')),
            'exp_date': _date(data.get('exp_date')),
            'order_value': float(data.get('order_value') or 0),
            'qty': int(float(data.get('qty') or 0)),
            'omt': str(data.get('omt') or '').strip()[:255],
        }
        with _conn() as (cur, d):
            ph = d['ph']
            cols = list(_FIELDS) + ['created_by', 'created_at']
            marks = ', '.join([ph] * len(cols))
            args = [vals[f] for f in _FIELDS] + [str(user or '')[:80],
                    _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
            cur.execute(f"INSERT INTO tracker_manual ({', '.join(cols)}) VALUES ({marks})",
                        tuple(args))
            cur.connection.commit()
        return {'ok': True}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f'{type(e).__name__}: {e}'}


def delete(row_id) -> dict:
    ensure_table()
    try:
        with _conn() as (cur, d):
            cur.execute(f"DELETE FROM tracker_manual WHERE id={d['ph']}", (int(row_id),))
            cur.connection.commit()
        return {'ok': True}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f'{type(e).__name__}: {e}'}


def list_manual() -> list[dict]:
    """Manual rows shaped like the auto tracker rows (+ ``source='manual'`` and
    ``id`` so they can be deleted). Never raises."""
    ensure_table()
    out = []
    try:
        with _conn() as (cur, d):
            cur.execute("SELECT id, dept, warehouse, marketplace, po, external_doc, "
                        "location, pincode, zone, po_date, exp_date, order_value, "
                        "qty, omt, created_at FROM tracker_manual ORDER BY id DESC")
            for r in cur.fetchall():
                out.append({
                    'id': r[0], 'dept': r[1] or '', 'wh': r[2] or '',
                    'marketplace': r[3] or '', 'po': r[4] or '',
                    'external_doc': r[5] or '', 'location': r[6] or '',
                    'pincode': r[7] or '', 'zone': r[8] or '', 'po_date': r[9],
                    'exp_date': r[10], 'order_value': float(r[11] or 0),
                    'qty': int(r[12] or 0), 'omt': r[13] or '',
                    'uploaded': r[14], 'file_source': '', 'source': 'manual',
                })
    except Exception:  # noqa: BLE001
        pass
    return out
