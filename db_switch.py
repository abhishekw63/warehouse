#!/usr/bin/env python
"""One-click DB switch — point the WHOLE app (desktop engine + web + Django ORM)
at the LOCAL MySQL or the TiDB server, by swapping the active db_config.json.

Nothing else changes: every connection reads history_db.load_db_config(), so a
single file swap re-targets the app. Your current local setup is untouched until
you switch, and switching back is one command.

Profiles live in  db_profiles/<name>.json  (gitignored — they hold creds).
Templates: db_profiles/local.json.example, db_profiles/tidb.json.example.

Usage:
  python db_switch.py status          # show the active target
  python db_switch.py save local      # snapshot the CURRENT active config -> profile "local"
  python db_switch.py local           # switch to local MySQL
  python db_switch.py tidb            # switch to the TiDB server
"""
import json
import shutil
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(_ROOT))
sys.path.insert(0, str(_ROOT / 'online_po_management'))
from online_po_processor.auto.history_db import _db_config_path  # noqa: E402

PROFILES = _ROOT / 'db_profiles'


def _active():
    p = _db_config_path()
    if not p.exists():
        return None, p
    try:
        return json.loads(p.read_text(encoding='utf-8-sig')), p
    except Exception:  # noqa: BLE001
        return {'_error': 'unreadable'}, p


def _describe(cfg):
    if not cfg:
        return '(none — SQLite fallback)'
    return f"{cfg.get('host', '?')}:{cfg.get('port', '?')} / {cfg.get('database', '?')}" \
           + ('  [TLS]' if (cfg.get('ssl') or cfg.get('ssl_ca')) else '')


def main():
    arg = (sys.argv[1] if len(sys.argv) > 1 else 'status').lower()
    dest = _db_config_path()

    if arg == 'status':
        cfg, p = _active()
        print(f'Active db_config: {p}')
        print(f'  -> {_describe(cfg)}')
        if PROFILES.exists():
            names = sorted(x.stem for x in PROFILES.glob('*.json'))
            print(f'  profiles available: {", ".join(names) or "(none)"}')
        return

    if arg == 'save':
        name = (sys.argv[2] if len(sys.argv) > 2 else '').lower()
        if not name:
            print('Usage: python db_switch.py save <name>'); sys.exit(1)
        cfg, p = _active()
        if not cfg:
            print('No active db_config to save.'); sys.exit(1)
        PROFILES.mkdir(exist_ok=True)
        out = PROFILES / f'{name}.json'
        out.write_text(json.dumps(cfg, indent=2), encoding='utf-8')
        print(f'Saved current active config -> {out}\n  -> {_describe(cfg)}')
        return

    # otherwise: activate a profile by name
    prof = PROFILES / f'{arg}.json'
    if not prof.exists():
        print(f'No profile "{arg}" at {prof}\n'
              f'Create it (copy db_profiles/{arg}.json.example) or run '
              f'"python db_switch.py save {arg}" first.')
        sys.exit(1)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(prof, dest)
    cfg, _ = _active()
    print(f'Switched to "{arg}"  ->  {dest}\n  -> {_describe(cfg)}')


if __name__ == '__main__':
    main()
