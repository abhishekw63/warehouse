"""
online_b2b.services.issue_email_log
===================================

One row per run's **auto Issues email** — the log + the recovery source of truth.

Keyed by ``run_id`` so a send is idempotent (never twice) and a miss is
re-drivable (a ``failed``/``pending`` row is retried on the next Lock&Record or
via ``manage.py flush_issue_emails``). Statuses:

  * ``sent``    — SMTP accepted.
  * ``failed``  — SMTP/exception error (``error`` + ``attempts`` recorded); retried.
  * ``skipped`` — nothing to send (no issue lines / no recipient / auto-mail off);
                  logged with the reason so "no email" is never a silent mystery.

Small, single table — justified: it's what prevents duplicates, drives retries,
and gives an audit trail (a column couldn't do all three). See [[minimize-tables-columns]].
"""
from __future__ import annotations

import datetime as _dt

from .order_db import _conn

_TABLE = 'issue_email_log'
_READY = False        # process-local: the fixed DDL only needs to run ONCE

_MYSQL = f"""
CREATE TABLE IF NOT EXISTS {_TABLE} (
  run_id BIGINT PRIMARY KEY,
  marketplace VARCHAR(64),
  status VARCHAR(16),
  n_excluded INT DEFAULT 0,
  n_included INT DEFAULT 0,
  recipients VARCHAR(500),
  subject VARCHAR(300),
  error VARCHAR(500),
  attempts INT DEFAULT 0,
  created_at DATETIME,
  sent_at DATETIME
)"""
_SQLITE = f"""
CREATE TABLE IF NOT EXISTS {_TABLE} (
  run_id INTEGER PRIMARY KEY, marketplace TEXT, status TEXT,
  n_excluded INTEGER DEFAULT 0, n_included INTEGER DEFAULT 0,
  recipients TEXT, subject TEXT, error TEXT, attempts INTEGER DEFAULT 0,
  created_at TEXT, sent_at TEXT
)"""

_COLS = ['run_id', 'marketplace', 'status', 'n_excluded', 'n_included',
         'recipients', 'subject', 'error', 'attempts', 'created_at', 'sent_at']


def ensure_table() -> None:
    global _READY
    if _READY:
        return
    with _conn() as (cur, d):
        cur.execute(_MYSQL if d['kind'] == 'mysql' else _SQLITE)
        cur.connection.commit()
    _READY = True


def get(run_id) -> dict | None:
    """The log row for a run, or None."""
    try:
        ensure_table()
        with _conn() as (cur, d):
            cur.execute(f"SELECT {', '.join(_COLS)} FROM {_TABLE} WHERE run_id={d['ph']}",
                        (int(run_id),))
            row = cur.fetchone()
            return dict(zip(_COLS, row)) if row else None
    except Exception:  # noqa: BLE001 — a log read must never break a page/lock
        return None


def status_of(run_id) -> str | None:
    r = get(run_id)
    return r.get('status') if r else None


def record(run_id, status: str, *, marketplace='', n_excluded=0, n_included=0,
           recipients='', subject='', error='') -> None:
    """Upsert the run's row. ``attempts`` increments on every call; ``sent_at`` is
    stamped only when status == 'sent'. Best-effort — never raises."""
    try:
        ensure_table()
        now = _dt.datetime.utcnow()
        sent_at = now if status == 'sent' else None
        with _conn() as (cur, d):
            ph = d['ph']
            if d['kind'] == 'mysql':
                cur.execute(
                    f"INSERT INTO {_TABLE} (run_id,marketplace,status,n_excluded,"
                    f"n_included,recipients,subject,error,attempts,created_at,sent_at) "
                    f"VALUES ({ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},1,{ph},{ph}) "
                    f"ON DUPLICATE KEY UPDATE status=VALUES(status),"
                    f"marketplace=VALUES(marketplace),n_excluded=VALUES(n_excluded),"
                    f"n_included=VALUES(n_included),recipients=VALUES(recipients),"
                    f"subject=VALUES(subject),error=VALUES(error),"
                    f"attempts={_TABLE}.attempts+1,sent_at=VALUES(sent_at)",
                    (int(run_id), marketplace[:64], status, int(n_excluded),
                     int(n_included), (recipients or '')[:500], (subject or '')[:300],
                     (error or '')[:500], now, sent_at))
            else:
                prev = get(run_id)
                attempts = (prev.get('attempts', 0) if prev else 0) + 1
                cur.execute(
                    f"INSERT OR REPLACE INTO {_TABLE} (run_id,marketplace,status,"
                    f"n_excluded,n_included,recipients,subject,error,attempts,"
                    f"created_at,sent_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                    (int(run_id), marketplace[:64], status, int(n_excluded),
                     int(n_included), (recipients or '')[:500], (subject or '')[:300],
                     (error or '')[:500], attempts,
                     (prev.get('created_at') if prev else now.isoformat()),
                     sent_at.isoformat() if sent_at else None))
            cur.connection.commit()
    except Exception:  # noqa: BLE001 — logging must never break the caller
        pass


def pending_run_ids(limit: int = 25) -> list:
    """Run ids whose auto-email still needs sending (failed/pending), oldest first
    — the self-healing sweep's work list."""
    try:
        ensure_table()
        with _conn() as (cur, d):
            cur.execute(
                f"SELECT run_id FROM {_TABLE} WHERE status IN ('failed','pending') "
                f"ORDER BY created_at ASC LIMIT {int(limit)}")
            return [r[0] for r in cur.fetchall()]
    except Exception:  # noqa: BLE001
        return []


def unsent_count() -> int:
    """How many runs have a failed/pending auto-email (Issues-page banner)."""
    try:
        ensure_table()
        with _conn() as (cur, d):
            cur.execute(f"SELECT COUNT(*) FROM {_TABLE} "
                        f"WHERE status IN ('failed','pending')")
            return int(cur.fetchone()[0] or 0)
    except Exception:  # noqa: BLE001
        return 0
