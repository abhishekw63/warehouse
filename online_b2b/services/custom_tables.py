"""
online_b2b.services.custom_tables
=================================

A generic, no-code **master-tables** store for the dashboard "Tables" tab.

Design goal (per the user): build ONE flexible store so *any* new master table
can be created from the UI with **zero DB/schema changes** in future. Two tables:

  * ``custom_tables``      — one row per table (name, slug, column defs, colour
                            rules) — a JSON column schema, so columns are data.
  * ``custom_table_rows``  — one row per data row (``data`` = JSON keyed by the
                            table's column keys).

Thin + JSON-returning (API-ready): every function returns plain dicts/lists so a
future React/DRF layer can consume it unchanged. Self-contained & removable —
touches nothing else in the app.
"""
from __future__ import annotations

import json
import re

from .order_db import _conn


def _slug(name: str) -> str:
    return re.sub(r'[^a-z0-9]+', '-', str(name).lower()).strip('-') or 'table'


def _loads(v):
    """MySQL JSON comes back as str (older) or already-decoded (newer)."""
    if v is None:
        return None
    if isinstance(v, (list, dict)):
        return v
    try:
        return json.loads(v)
    except Exception:  # noqa: BLE001
        return None


# ── schema ────────────────────────────────────────────────────────────────
def ensure_schema() -> None:
    with _conn() as (cur, _d):
        cur.execute("""
            CREATE TABLE IF NOT EXISTS custom_tables (
              id INT AUTO_INCREMENT PRIMARY KEY,
              name VARCHAR(255) NOT NULL,
              slug VARCHAR(255) NOT NULL UNIQUE,
              columns JSON NOT NULL,
              color_rules JSON NULL,
              sort INT DEFAULT 0,
              created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            ) CHARACTER SET utf8mb4""")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS custom_table_rows (
              id INT AUTO_INCREMENT PRIMARY KEY,
              table_id INT NOT NULL,
              data JSON NOT NULL,
              sort INT DEFAULT 0,
              updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
              INDEX idx_ctr_table (table_id)
            ) CHARACTER SET utf8mb4""")


# ── tables CRUD ─────────────────────────────────────────────────────────────
def list_tables() -> list[dict]:
    with _conn() as (cur, _d):
        cur.execute("SELECT id,name,slug,columns,color_rules,sort FROM custom_tables ORDER BY sort,id")
        rows = cur.fetchall()
        counts = {}
        cur.execute("SELECT table_id,COUNT(*) FROM custom_table_rows GROUP BY table_id")
        for tid, n in cur.fetchall():
            counts[tid] = int(n)
        return [{'id': tid, 'name': name, 'slug': slug,
                 'columns': _loads(cols) or [], 'color_rules': _loads(cr) or {},
                 'rows': counts.get(tid, 0), 'sort': sort}
                for tid, name, slug, cols, cr, sort in rows]


def get_table(ident) -> dict | None:
    with _conn() as (cur, _d):
        if str(ident).isdigit():
            cur.execute("SELECT id,name,slug,columns,color_rules FROM custom_tables WHERE id=%s", (int(ident),))
        else:
            cur.execute("SELECT id,name,slug,columns,color_rules FROM custom_tables WHERE slug=%s", (str(ident),))
        r = cur.fetchone()
        if not r:
            return None
        return {'id': r[0], 'name': r[1], 'slug': r[2],
                'columns': _loads(r[3]) or [], 'color_rules': _loads(r[4]) or {}}


def create_table(name: str, columns: list[dict], color_rules: dict | None = None) -> int:
    import uuid
    slug = _slug(name)
    with _conn() as (cur, _d):
        cur.execute("SELECT COUNT(*) FROM custom_tables WHERE slug=%s OR slug LIKE %s", (slug, slug + '-%'))
        if cur.fetchone()[0]:
            slug = f"{slug}-{uuid.uuid4().hex[:4]}"
        cur.execute("SELECT COALESCE(MAX(sort),0)+1 FROM custom_tables")
        sort = cur.fetchone()[0]
        cur.execute("INSERT INTO custom_tables (name,slug,columns,color_rules,sort) VALUES (%s,%s,%s,%s,%s)",
                    (name, slug, json.dumps(columns), json.dumps(color_rules or {}), sort))
        cur.execute("SELECT LAST_INSERT_ID()")
        return int(cur.fetchone()[0])


def update_table(table_id: int, name=None, columns=None, color_rules=None) -> None:
    sets, vals = [], []
    if name is not None:
        sets.append("name=%s"); vals.append(name)
    if columns is not None:
        sets.append("columns=%s"); vals.append(json.dumps(columns))
    if color_rules is not None:
        sets.append("color_rules=%s"); vals.append(json.dumps(color_rules))
    if not sets:
        return
    vals.append(table_id)
    with _conn() as (cur, _d):
        cur.execute(f"UPDATE custom_tables SET {','.join(sets)} WHERE id=%s", vals)


def delete_table(table_id: int) -> None:
    with _conn() as (cur, _d):
        cur.execute("DELETE FROM custom_table_rows WHERE table_id=%s", (table_id,))
        cur.execute("DELETE FROM custom_tables WHERE id=%s", (table_id,))


# ── rows CRUD ───────────────────────────────────────────────────────────────
def list_rows(table_id: int) -> list[dict]:
    with _conn() as (cur, _d):
        cur.execute("SELECT id,data,sort FROM custom_table_rows WHERE table_id=%s ORDER BY sort,id", (table_id,))
        return [{'id': r[0], 'data': _loads(r[1]) or {}, 'sort': r[2]} for r in cur.fetchall()]


def add_row(table_id: int, data: dict) -> int:
    with _conn() as (cur, _d):
        cur.execute("SELECT COALESCE(MAX(sort),0)+1 FROM custom_table_rows WHERE table_id=%s", (table_id,))
        sort = cur.fetchone()[0]
        cur.execute("INSERT INTO custom_table_rows (table_id,data,sort) VALUES (%s,%s,%s)",
                    (table_id, json.dumps(data), sort))
        cur.execute("SELECT LAST_INSERT_ID()")
        return int(cur.fetchone()[0])


def update_row(row_id: int, data: dict) -> None:
    with _conn() as (cur, _d):
        cur.execute("UPDATE custom_table_rows SET data=%s WHERE id=%s", (json.dumps(data), row_id))


def delete_row(row_id: int) -> None:
    with _conn() as (cur, _d):
        cur.execute("DELETE FROM custom_table_rows WHERE id=%s", (row_id,))
