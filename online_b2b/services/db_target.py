"""
online_b2b.services.db_target
=============================

Active-database target management for the **Setup** page (and the db_switch.py
CLI). Switching = copy ``db_profiles/<name>.json`` onto the active db_config the
whole app reads (``history_db.load_db_config``). Because every business
connection re-reads that config per call (raw pymysql in ``order_db`` +
``history_db``), a switch takes effect on the next request — no code change, one
file swap.

Profiles (gitignored, hold creds): ``db_profiles/local.json`` (your current
working local MySQL) and ``db_profiles/tidb.json`` (the server). Templates:
``db_profiles/*.json.example``.
"""
from __future__ import annotations

import json
import shutil
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[2]
PROFILES = _ROOT / 'db_profiles'


def _active_path() -> Path:
    from online_po_processor.auto.history_db import _db_config_path
    return _db_config_path()


def profiles() -> list[str]:
    return sorted(p.stem for p in PROFILES.glob('*.json')) if PROFILES.exists() else []


def _read(p: Path) -> dict:
    try:
        return json.loads(p.read_text(encoding='utf-8-sig'))
    except Exception:  # noqa: BLE001
        return {}


def _match_profile(cfg: dict) -> str | None:
    """Which saved profile does the active config correspond to (by host:port)."""
    for name in profiles():
        pc = _read(PROFILES / f'{name}.json')
        if pc.get('host') == cfg.get('host') and str(pc.get('port')) == str(cfg.get('port')):
            return name
    return None


def status() -> dict:
    """Current active target + the profiles available to switch to."""
    from online_po_processor.auto.history_db import load_db_config
    cfg = load_db_config() or {}
    return {
        'backend': cfg.get('backend'),
        'host': cfg.get('host'),
        'port': cfg.get('port'),
        'database': cfg.get('database'),
        'tls': bool(cfg.get('ssl') or cfg.get('ssl_ca')),
        'active': _match_profile(cfg),
        'active_path': str(_active_path()),
        'profiles': profiles(),
    }


def save_current(name: str) -> dict:
    """Snapshot the CURRENT active connection as profile ``name`` (e.g. 'local')
    — so you can always switch back to exactly what works now."""
    from online_po_processor.auto.history_db import load_db_config
    cfg = load_db_config()
    if not cfg:
        return {'ok': False, 'error': 'No active DB config to save.'}
    try:
        PROFILES.mkdir(exist_ok=True)
        (PROFILES / f'{name}.json').write_text(
            json.dumps(cfg, indent=2), encoding='utf-8')
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f'Could not save: {e}'}
    return {'ok': True, **status()}


def save_tidb(fields) -> dict:
    """Write the TiDB profile from the Setup form (host/port/user/password/db).
    TLS is forced on (TiDB Serverless requires it)."""
    host = (fields.get('host') or '').strip()
    user = (fields.get('user') or '').strip()
    if not host or not user:
        return {'ok': False, 'error': 'TiDB host and user are required.'}
    prof = {
        'backend': 'mysql', 'host': host,
        'port': int(fields.get('port') or 4000),
        'user': user, 'password': fields.get('password') or '',
        'database': (fields.get('database') or 'renee_orders').strip(),
        'ssl': True,
    }
    try:
        PROFILES.mkdir(exist_ok=True)
        (PROFILES / 'tidb.json').write_text(
            json.dumps(prof, indent=2), encoding='utf-8')
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f'Could not save: {e}'}
    return {'ok': True, **status()}


def switch(name: str) -> dict:
    """Activate a profile by copying it onto the active db_config path."""
    prof = PROFILES / f'{name}.json'
    if not prof.exists():
        return {'ok': False, 'error': f'No DB profile "{name}". '
                f'Create db_profiles/{name}.json (see the .example).'}
    dest = _active_path()
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(prof, dest)
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f'Could not write {dest}: {e}'}
    return {'ok': True, 'active': name, **status()}
