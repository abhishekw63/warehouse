"""DB-backed offline master workbooks — so the offline engines (MT-Select, EKA…)
need NO local Excel file on the server.

The authoritative master workbook lives in the DB; on each run the bridge
MATERIALISES it to a temp file and hands that path to the frozen engine's loader,
which is completely unchanged. Upload / replace the workbook through the app; the
server no longer depends on the filesystem (works on Render's ephemeral disk — the
temp file is regenerated from the DB per worker, cached by upload time so the blob
isn't re-fetched every run).

TiDB caps a single row/entry at ~6 MB, and these workbooks are larger, so the bytes
are CHUNKED across ``offline_master_chunk`` rows (≤4 MB each) and reassembled on read.
Meta lives in ``offline_master_file``. Kept in ``offline/services`` (never touches
``offline_po_management`` — the frozen engine). Uses the shared ``order_db`` helpers.
"""
from __future__ import annotations

import datetime as _dt
import tempfile
from pathlib import Path

from online_b2b.services.order_db import _conn, _conn_tx

_META = 'offline_master_file'
_CHUNK = 'offline_master_chunk'
_CHUNK_BYTES = 4_000_000                     # < TiDB's ~6 MB per-entry limit
_cache: dict[str, tuple[str, Path]] = {}     # channel -> (uploaded_at, temp Path)


def _ensure(cur):
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS {_META} ("
        f"channel VARCHAR(32) PRIMARY KEY, filename VARCHAR(255), "
        f"size_bytes BIGINT, n_chunks INT, "
        f"uploaded_at DATETIME, uploaded_by VARCHAR(64))")
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS {_CHUNK} ("
        f"channel VARCHAR(32), seq INT, content LONGBLOB, "
        f"PRIMARY KEY (channel, seq))")


def put_master(channel: str, file_path, uploaded_by: str = '') -> dict:
    """Store / replace a channel's master workbook in the DB from a local file,
    chunked so no single row exceeds TiDB's entry-size limit."""
    p = Path(file_path)
    data = p.read_bytes()
    chunks = [data[i:i + _CHUNK_BYTES] for i in range(0, len(data), _CHUNK_BYTES)] or [b'']
    now = _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with _conn_tx() as (cur, d):
        _ensure(cur)
        ph = d['ph']
        cur.execute(f"DELETE FROM {_CHUNK} WHERE channel={ph}", (channel,))
        for seq, blob in enumerate(chunks):
            cur.execute(
                f"INSERT INTO {_CHUNK} (channel, seq, content) VALUES ({ph},{ph},{ph})",
                (channel, seq, blob))
        cur.execute(
            f"REPLACE INTO {_META} "
            f"(channel, filename, size_bytes, n_chunks, uploaded_at, uploaded_by) "
            f"VALUES ({ph},{ph},{ph},{ph},{ph},{ph})",
            (channel, p.name, len(data), len(chunks), now, uploaded_by or ''))
    _cache.pop(channel, None)
    return {'ok': True, 'channel': channel, 'filename': p.name,
            'bytes': len(data), 'chunks': len(chunks), 'uploaded_at': now}


def master_info(channel: str) -> dict | None:
    """Metadata (no blob) for a channel's stored master, or None if absent."""
    try:
        with _conn() as (cur, d):
            _ensure(cur)
            ph = d['ph']
            cur.execute(
                f"SELECT filename, size_bytes, n_chunks, uploaded_at, uploaded_by "
                f"FROM {_META} WHERE channel={ph}", (channel,))
            r = cur.fetchone()
    except Exception:  # noqa: BLE001
        return None
    if not r:
        return None
    return {'filename': r[0] or '', 'size_bytes': int(r[1] or 0),
            'n_chunks': int(r[2] or 0), 'uploaded_at': str(r[3] or ''),
            'uploaded_by': r[4] or ''}


def materialize(channel: str) -> Path | None:
    """Reassemble the channel's DB master workbook to a temp file and return its
    path. Cached per (channel, uploaded_at) so chunks aren't re-fetched every run.
    Returns None if the DB holds no master for this channel (caller falls back)."""
    try:
        with _conn() as (cur, d):
            _ensure(cur)
            ph = d['ph']
            cur.execute(
                f"SELECT filename, uploaded_at FROM {_META} WHERE channel={ph}", (channel,))
            r = cur.fetchone()
            if not r:
                return None
            fname, ver = (r[0] or f'{channel}_masters.xlsx'), str(r[1] or '')
            cached = _cache.get(channel)
            if cached and cached[0] == ver and cached[1].exists():
                return cached[1]
            cur.execute(
                f"SELECT content FROM {_CHUNK} WHERE channel={ph} ORDER BY seq", (channel,))
            parts = [row[0] for row in cur.fetchall() if row[0] is not None]
    except Exception:  # noqa: BLE001
        return None
    if not parts:
        return None
    outdir = Path(tempfile.gettempdir()) / 'offline_masters'
    outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / f'{channel}__{fname}'
    out.write_bytes(b''.join(parts))
    _cache[channel] = (ver, out)
    return out
