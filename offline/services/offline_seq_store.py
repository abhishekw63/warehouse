"""DB-backed SO-number sequence counter for the offline MT-Select family (and
Reliance Trends, which reuses the same engine).

The desktop engine persists the per-channel ``{date, next_counter}`` to a local
``mt_select_seq.json``. On an ephemeral server that file is wiped on every restart /
redeploy, so ``load_seq_state()`` returns ``{}`` ("first run") and the counter
RESTARTS from its base → **duplicate SO numbers pushed to D365**. This keeps the
counter in the DB instead, so it survives restarts and is shared across workers.

The frozen engine is NOT touched: :func:`patch_engine` swaps the module-level
``load_seq_state`` / ``save_seq_state`` for DB-backed versions, so BOTH the bridge's
direct calls AND the engine's own ``assign_so_numbers`` (which calls them as module
globals) read/write the DB.
"""
from __future__ import annotations

import datetime as _dt

from online_b2b.services.order_db import _conn, _conn_tx

_TABLE = 'offline_seq_state'


_READY = False        # process-local: the fixed DDL only needs to run ONCE


def _ensure(cur):
    global _READY
    if _READY:
        return
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS {_TABLE} ("
        f"channel VARCHAR(32) PRIMARY KEY, seq_date VARCHAR(16), "
        f"next_counter BIGINT, updated_at DATETIME)")
    _READY = True


def db_load_state() -> dict:
    """Full per-channel seq state from the DB: ``{channel: {date, next_counter}}``.
    Shape-compatible with the engine's file-based ``load_seq_state``."""
    out: dict = {}
    try:
        with _conn() as (cur, _d):
            _ensure(cur)
            cur.execute(f"SELECT channel, seq_date, next_counter FROM {_TABLE}")
            for ch, sd, nc in cur.fetchall():
                if sd == '__scalar__':          # a top-level scalar (e.g. 'last_sequence')
                    out[ch] = int(nc or 0)
                else:
                    out[ch] = {'date': sd or '', 'next_counter': int(nc or 0)}
    except Exception:  # noqa: BLE001
        pass
    return out


def db_save_state(state: dict) -> None:
    """Persist the whole seq-state dict (upsert per channel). Written VERBATIM — no
    monotonic guard, because the preview path deliberately restores an earlier
    snapshot to un-burn the counter, which must be allowed to lower it."""
    try:
        now = _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with _conn_tx() as (cur, d):
            _ensure(cur)
            ph = d['ph']
            for ch, s in (state or {}).items():
                if not ch:
                    continue
                if isinstance(s, dict):
                    sd = str(s.get('date', ''))
                    nc = int(s.get('next_counter', 0) or 0)
                else:                            # top-level scalar (e.g. 'last_sequence')
                    try:
                        nc = int(s)
                    except (TypeError, ValueError):
                        continue
                    sd = '__scalar__'
                cur.execute(
                    f"REPLACE INTO {_TABLE} "
                    f"(channel, seq_date, next_counter, updated_at) "
                    f"VALUES ({ph},{ph},{ph},{ph})",
                    (str(ch), sd, nc, now))
    except Exception:  # noqa: BLE001
        pass


def patch_engine(eng) -> None:
    """Idempotently swap the engine's file-based seq load/save for the DB versions.
    Covers ``assign_so_numbers`` (module global) + every bridge call site."""
    if getattr(eng, '_db_seq_patched', False):
        return
    eng.load_seq_state = db_load_state
    eng.save_seq_state = db_save_state
    eng._db_seq_patched = True


def seed_from_file(path, only_if_empty: bool = True) -> dict:
    """One-time: copy a local ``mt_select_seq.json`` into the DB so the server
    continues the sequence from the real current counter (never restarts). By
    default only seeds when the DB is still empty, so it can never clobber a counter
    the server has already advanced."""
    import json
    from pathlib import Path
    p = Path(path)
    if not p.exists():
        return {'ok': False, 'error': f'no seq file at {p}'}
    try:
        data = json.loads(p.read_text(encoding='utf-8'))
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': str(e)}
    if not isinstance(data, dict):
        return {'ok': False, 'error': 'seq file is not a dict'}
    if only_if_empty and db_load_state():
        return {'ok': True, 'skipped': 'DB already has seq state'}
    db_save_state(data)
    return {'ok': True, 'seeded': sorted(data.keys())}
