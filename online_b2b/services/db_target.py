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
