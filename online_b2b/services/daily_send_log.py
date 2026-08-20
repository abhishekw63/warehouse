"""Daily automated-email send log — a tiny audit + idempotency guard so a
scheduled send (Windows Task Scheduler / Render Cron) never mails the same day
twice and you can see what went out. Self-contained + removable: one table, three
functions. [[modular-removable-features]] [[issues-email-exclude-only]]
"""
from __future__ import annotations

import datetime as _dt

from .order_db import _conn

_DDL_MYSQL = """
CREATE TABLE IF NOT EXISTS daily_email_log (
    id         BIGINT AUTO_INCREMENT PRIMARY KEY,
    kind       VARCHAR(40),
    for_date   VARCHAR(10),
    sent_at    DATETIME,
    n_rows     INT DEFAULT 0,
    ok         TINYINT DEFAULT 0,
    error      VARCHAR(255),
    recipients VARCHAR(500),
    INDEX idx_dailymail (kind, for_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
_DDL_SQLITE = """
CREATE TABLE IF NOT EXISTS daily_email_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT, kind TEXT, for_date TEXT, sent_at TEXT,
    n_rows INTEGER DEFAULT 0, ok INTEGER DEFAULT 0, error TEXT, recipients TEXT)
"""
_READY = False


def _ensure() -> None:
    global _READY
    if _READY:
        return
    with _conn() as (cur, d):
        cur.execute(_DDL_MYSQL if d['kind'] == 'mysql' else _DDL_SQLITE)
    _READY = True


def already_sent(kind: str, for_date: str) -> bool:
    """True if a SUCCESSFUL send of ``kind`` for ``for_date`` (YYYY-MM-DD) is
    already logged — used by ``--if-not-sent`` so a logon-triggered scheduler
    won't re-send the same day."""
    _ensure()
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"SELECT COUNT(*) FROM daily_email_log WHERE kind={ph} "
                    f"AND for_date={ph} AND ok=1", (kind, for_date))
        return (cur.fetchone()[0] or 0) > 0


def record(kind: str, for_date: str, n_rows: int, ok: bool,
           error: str = '', recipients: str = '') -> None:
    """Append one send attempt (success or failure) to the audit log."""
    _ensure()
    now = _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(
            f"INSERT INTO daily_email_log (kind, for_date, sent_at, n_rows, ok, "
            f"error, recipients) VALUES ({ph},{ph},{ph},{ph},{ph},{ph},{ph})",
            (kind, for_date, now, int(n_rows), 1 if ok else 0,
             (error or '')[:255], (recipients or '')[:500]))
