"""_conn_tx() atomicity — a partial/interrupted Lock & Record must leave NOTHING.

Uses a throwaway sqlite file (monkeypatched _backend) so it never touches the
real TiDB/MySQL. Proves: clean block → committed; any exception mid-block →
full rollback (no partial rows).
"""

import sqlite3

import pytest

from online_b2b.services import order_db


@pytest.fixture
def sqlite_backend(tmp_path, monkeypatch):
    dbf = tmp_path / 'tx.sqlite3'
    con = sqlite3.connect(str(dbf))
    con.execute('CREATE TABLE t (id INTEGER PRIMARY KEY, v TEXT)')
    con.commit()
    con.close()
    monkeypatch.setattr(order_db, '_backend', lambda: ('sqlite', str(dbf)))
    return str(dbf)


def _count(dbf):
    con = sqlite3.connect(dbf)
    try:
        return con.execute('SELECT COUNT(*) FROM t').fetchone()[0]
    finally:
        con.close()


def test_conn_tx_commits_on_clean_exit(sqlite_backend):
    with order_db._conn_tx() as (cur, d):
        cur.execute('INSERT INTO t (v) VALUES (?)', ('a',))
        cur.execute('INSERT INTO t (v) VALUES (?)', ('b',))
    assert _count(sqlite_backend) == 2          # both committed


def test_conn_tx_rolls_back_on_exception(sqlite_backend):
    with pytest.raises(RuntimeError):
        with order_db._conn_tx() as (cur, d):
            cur.execute('INSERT INTO t (v) VALUES (?)', ('a',))   # would be a partial write
            raise RuntimeError('simulated interruption mid-record')
    assert _count(sqlite_backend) == 0          # NOTHING left — full rollback


def test_conn_tx_isolated_failures_leave_prior_commits(sqlite_backend):
    # A committed run stays; a later failed run rolls back completely.
    with order_db._conn_tx() as (cur, d):
        cur.execute('INSERT INTO t (v) VALUES (?)', ('run1',))
    assert _count(sqlite_backend) == 1
    with pytest.raises(ValueError):
        with order_db._conn_tx() as (cur, d):
            cur.execute('INSERT INTO t (v) VALUES (?)', ('run2-partial',))
            raise ValueError('boom')
    assert _count(sqlite_backend) == 1          # run1 intact, run2 fully gone
