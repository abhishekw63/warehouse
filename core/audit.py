"""
core.audit
==========

Lightweight **write-audit trail** — who did which mutating action, and when.

Written automatically by :class:`core.access.WriteGuardMiddleware` for every
Editor write (POST/PUT/PATCH/DELETE that isn't a read-only export/search), so the
"who did each step" record is captured app-wide with **no per-view code**. Stored
in the business DB (``renee_orders``) next to the data. Fully best-effort — a
logging failure NEVER breaks the underlying request.

This complements the on-entity attribution (``runs.recorded_by`` = who recorded a
run): the audit_log is the chronological "who touched what" stream (line
decisions, mapping edits, inventory, deletes, role changes, …).
"""
from __future__ import annotations

import datetime as _dt

_READY = False


def _conn():
    from online_b2b.services.order_db import _conn as c
    return c()


def ensure_table() -> None:
    global _READY
    if _READY:
        return
    with _conn() as (cur, d):
        if d.get('kind') == 'mysql':
            cur.execute("""
                CREATE TABLE IF NOT EXISTS audit_log (
                    id       BIGINT AUTO_INCREMENT PRIMARY KEY,
                    ts       DATETIME,
                    username VARCHAR(150),
                    method   VARCHAR(10),
                    url_name VARCHAR(120),
                    path     VARCHAR(300),
                    target   VARCHAR(300),
                    detail   VARCHAR(500),
                    ms       INT,
                    q        INT,
                    INDEX idx_audit_ts (ts),
                    INDEX idx_audit_user (username)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
        else:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS audit_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, username TEXT,
                    method TEXT, url_name TEXT, path TEXT, target TEXT, detail TEXT,
                    ms INTEGER, q INTEGER)""")
        # Self-migrate an already-existing table: add the timing columns (ms = wall
        # time, q = SQL query count) if missing. Best-effort — a duplicate-column
        # error just means they're already there.
        for _col in ('ms', 'q'):
            try:
                cur.execute(f"ALTER TABLE audit_log ADD COLUMN {_col} INT")
            except Exception:  # noqa: BLE001
                pass
        cur.connection.commit()
    _READY = True


def log(username, method, url_name, path, target='', detail='', ms=None, q=None):
    """Append one audit row; returns its id (or None). Best-effort — swallows all
    errors. ``ms``/``q`` (wall time + SQL count) are usually stamped LATER via
    :func:`set_timing` since the duration isn't known until the request finishes."""
    try:
        ensure_table()
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(
                f"INSERT INTO audit_log (ts, username, method, url_name, path, "
                f"target, detail, ms, q) "
                f"VALUES ({ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph})",
                (_dt.datetime.now(), str(username or '')[:150], str(method or '')[:10],
                 str(url_name or '')[:120], str(path or '')[:300],
                 str(target or '')[:300], str(detail or '')[:500],
                 int(ms) if ms is not None else None,
                 int(q) if q is not None else None))
            cur.connection.commit()
            return cur.lastrowid
    except Exception:  # noqa: BLE001 — audit must never break a request
        return None


def set_timing(row_id, ms, q=None) -> None:
    """Stamp elapsed ms (+ SQL query count) onto an audit row after the request has
    finished. The row is written in ``process_view`` (before the view runs), so the
    duration is only known at the very end — PerfMiddleware calls this. Best-effort."""
    if not row_id:
        return
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(f"UPDATE audit_log SET ms={ph}, q={ph} WHERE id={ph}",
                        (int(ms), int(q) if q is not None else None, row_id))
            cur.connection.commit()
    except Exception:  # noqa: BLE001
        pass


def recent(limit: int = 300, user: str = '', q: str = '') -> list[dict]:
    """Most-recent audit rows (newest first) for the staff Audit Log page."""
    try:
        ensure_table()
        cols = ['id', 'ts', 'username', 'method', 'url_name', 'path', 'target',
                'detail', 'ms', 'q']
        where, params = [], []
        with _conn() as (cur, d):
            ph = d['ph']
            if user:
                where.append(f"username={ph}"); params.append(user)
            if q:
                where.append(f"(path LIKE {ph} OR target LIKE {ph} OR url_name LIKE {ph})")
                params += [f"%{q}%", f"%{q}%", f"%{q}%"]
            wsql = (' WHERE ' + ' AND '.join(where)) if where else ''
            cur.execute(
                f"SELECT {', '.join(cols)} FROM audit_log{wsql} "
                f"ORDER BY id DESC LIMIT {int(limit)}", tuple(params))
            return [dict(zip(cols, r)) for r in cur.fetchall()]
    except Exception:  # noqa: BLE001
        return []
