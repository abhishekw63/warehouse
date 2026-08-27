"""DB-backed 'Review Later' drafts — so parked runs persist across restarts and are
visible on BOTH local and production (they lived only in the per-token upload folder
on the local disk, which Render neither shares nor keeps across redeploys).

A draft is a parked upload token = a small folder (``meta.json`` + ``preview.json`` +
the raw PO file(s)). We SNAPSHOT that whole folder into the DB on park, and MATERIALISE
it back to the upload folder on demand (``_load_meta`` restores it when the local copy
is missing), so the existing file-based review → re-validate → finalise flow works
unchanged on any server. Files are chunked (≤4 MB) to fit TiDB's ~6 MB entry limit,
mirroring ``offline_master_store``.
"""
from __future__ import annotations

import datetime as _dt
import json as _json
from pathlib import Path

from online_b2b.services.order_db import _conn, _conn_tx

_META = 'parked_draft'
_FILE = 'parked_draft_file'
_CHUNK = 4_000_000


def _ensure(cur):
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS {_META} ("
        f"token VARCHAR(64) PRIMARY KEY, marketplace VARCHAR(64), "
        f"draft_at VARCHAR(24), draft_note VARCHAR(320), pos INT, undecided INT, "
        f"files INT, meta_json MEDIUMTEXT, updated_at DATETIME)")
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS {_FILE} ("
        f"token VARCHAR(64), filename VARCHAR(255), seq INT, content LONGBLOB, "
        f"PRIMARY KEY (token, filename, seq))")


def _counts(dir_path: Path, meta: dict) -> tuple[int, int]:
    """(pos, undecided) from the cached preview — mirrors views._collect_drafts."""
    npos = undecided = 0
    cache = dir_path / 'preview.json'
    if cache.exists():
        try:
            res = (_json.loads(cache.read_text(encoding='utf-8')).get('res') or {})
            npos = len(res.get('headers') or [])
            dec = meta.get('decisions') or {}
            for ln in (res.get('affected') or []):
                k = f"{ln.get('po', '')}|{ln.get('item_no', '')}|{ln.get('ean', '')}"
                if not (dec.get(k) or {}).get('action'):
                    undecided += 1
        except Exception:  # noqa: BLE001
            pass
    return npos, undecided


def snapshot(token: str, dir_path, meta: dict | None = None) -> dict:
    """Snapshot a parked token folder (all files) into the DB. Best-effort."""
    try:
        d = Path(dir_path)
        if not d.is_dir():
            return {'ok': False, 'error': 'no dir'}
        if meta is None:
            mp = d / 'meta.json'
            meta = _json.loads(mp.read_text(encoding='utf-8')) if mp.exists() else {}
        files = [f for f in d.iterdir() if f.is_file()]
        pos, undecided = _counts(d, meta)
        now = _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with _conn_tx() as (cur, dd):
            _ensure(cur)
            ph = dd['ph']
            cur.execute(f"DELETE FROM {_FILE} WHERE token={ph}", (token,))
            for f in files:
                data = f.read_bytes()
                chunks = [data[i:i + _CHUNK] for i in range(0, len(data), _CHUNK)] or [b'']
                for seq, blob in enumerate(chunks):
                    cur.execute(
                        f"INSERT INTO {_FILE} (token, filename, seq, content) "
                        f"VALUES ({ph},{ph},{ph},{ph})", (token, f.name, seq, blob))
            cur.execute(
                f"REPLACE INTO {_META} (token, marketplace, draft_at, draft_note, "
                f"pos, undecided, files, meta_json, updated_at) "
                f"VALUES ({ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph})",
                (token, str(meta.get('marketplace', '')), str(meta.get('draft_at', '')),
                 str(meta.get('draft_note', ''))[:320], pos, undecided, len(files),
                 _json.dumps(meta), now))
        return {'ok': True, 'token': token, 'files': len(files)}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f'{type(e).__name__}: {e}'}


def list_drafts() -> list[dict]:
    """All DB-parked drafts as the same dicts views._collect_drafts yields."""
    out: list[dict] = []
    try:
        with _conn() as (cur, d):
            _ensure(cur)
            cur.execute(
                f"SELECT token, marketplace, draft_at, draft_note, pos, undecided, files "
                f"FROM {_META} ORDER BY draft_at DESC")
            for r in cur.fetchall():
                out.append({'token': r[0], 'marketplace': r[1] or '',
                            'draft_at': r[2] or '', 'note': r[3] or '',
                            'pos': int(r[4] or 0), 'undecided': int(r[5] or 0),
                            'files': int(r[6] or 0)})
    except Exception:  # noqa: BLE001
        pass
    return out


def has(token: str) -> bool:
    try:
        with _conn() as (cur, d):
            _ensure(cur)
            ph = d['ph']
            cur.execute(f"SELECT 1 FROM {_META} WHERE token={ph}", (token,))
            return cur.fetchone() is not None
    except Exception:  # noqa: BLE001
        return False


def materialize(token: str, dir_path) -> bool:
    """Reassemble a DB-parked token's folder onto the local disk (so the file-based
    review flow can read it). Returns True if it wrote anything."""
    try:
        with _conn() as (cur, d):
            _ensure(cur)
            ph = d['ph']
            cur.execute(f"SELECT 1 FROM {_META} WHERE token={ph}", (token,))
            if cur.fetchone() is None:
                return False
            cur.execute(
                f"SELECT filename, seq, content FROM {_FILE} WHERE token={ph} "
                f"ORDER BY filename, seq", (token,))
            parts: dict = {}
            for fn, seq, content in cur.fetchall():
                parts.setdefault(fn, []).append(content or b'')
    except Exception:  # noqa: BLE001
        return False
    if not parts:
        return False
    dd = Path(dir_path)
    dd.mkdir(parents=True, exist_ok=True)
    for fn, blobs in parts.items():
        (dd / fn).write_bytes(b''.join(blobs))
    return True


def delete(token: str) -> None:
    """Remove a draft from the DB (finalised or discarded)."""
    try:
        with _conn_tx() as (cur, d):
            _ensure(cur)
            ph = d['ph']
            cur.execute(f"DELETE FROM {_FILE} WHERE token={ph}", (token,))
            cur.execute(f"DELETE FROM {_META} WHERE token={ph}", (token,))
    except Exception:  # noqa: BLE001
        pass
