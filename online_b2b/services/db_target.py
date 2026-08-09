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
_LAST_BACKUP = _ROOT / 'logs' / 'tidb_last_backup.json'   # last successful TiDB->local backup


def last_backup() -> dict | None:
    """Details of the last successful TiDB -> local backup ({at, tables, rows,
    views, elapsed, source, target}), or None if none has run yet."""
    try:
        return json.loads(_LAST_BACKUP.read_text(encoding='utf-8'))
    except Exception:  # noqa: BLE001
        return None


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
        'last_backup': last_backup(),
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


def test(name: str) -> dict:
    """Live-test a profile: connect (with TLS) + run a sample query, report the
    server version, table count and order_headers rows. Never switches anything —
    a safe pre-flight check. Short timeout so a bad host can't hang the request."""
    prof = PROFILES / f'{name}.json'
    if not prof.exists():
        return {'ok': False, 'error': f'No DB profile "{name}" to test.'}
    cfg = _read(prof)
    try:
        import pymysql
        from online_po_processor.auto.history_db import mysql_ssl
        conn = pymysql.connect(
            host=cfg.get('host', '127.0.0.1'), port=int(cfg.get('port', 3306)),
            user=cfg.get('user', 'root'), password=cfg.get('password', ''),
            database=cfg.get('database', 'renee_orders'),
            charset='utf8mb4', connect_timeout=8, read_timeout=8,
            **mysql_ssl(cfg))
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': str(e), 'host': cfg.get('host'),
                'database': cfg.get('database')}
    try:
        cur = conn.cursor()
        cur.execute("SELECT VERSION()")
        version = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM information_schema.tables "
                    "WHERE table_schema = DATABASE()")
        n_tables = cur.fetchone()[0]
        oh = None
        try:
            cur.execute("SELECT COUNT(*) FROM order_headers")
            oh = cur.fetchone()[0]
        except Exception:  # noqa: BLE001 — table not there yet (fresh DB)
            oh = None
        return {'ok': True, 'version': version, 'tables': n_tables,
                'order_headers': oh, 'host': cfg.get('host'),
                'database': cfg.get('database')}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': str(e)}
    finally:
        conn.close()


def backup_tidb_to_local() -> dict:
    """ONE-CLICK BACKUP: copy ALL data from TiDB (source) into the local MySQL
    profile (destination) so local becomes an exact mirror of the server. Per base
    table: SHOW CREATE TABLE on TiDB -> DROP+CREATE on local -> batch-copy every
    row; then recreate views. **Destructive on LOCAL only** (local is overwritten
    to match TiDB); TiDB is read-only here. Reverse of ``db_push_to_tidb.py``.
    Returns a per-table summary or {ok:False,error}. Does NOT change the active
    target — after a backup the app keeps pointing wherever it already was."""
    import re
    import time
    src_p, dst_p = PROFILES / 'tidb.json', PROFILES / 'local.json'
    if not src_p.exists():
        return {'ok': False, 'error': 'No db_profiles/tidb.json (the source).'}
    if not dst_p.exists():
        return {'ok': False, 'error': 'No db_profiles/local.json (the backup target). '
                'Save your local MySQL as the "local" profile first.'}
    src, dst = _read(src_p), _read(dst_p)
    try:
        import pymysql
        from online_po_processor.auto.history_db import mysql_ssl
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f'pymysql unavailable: {e}'}

    def _c(cfg, autocommit):
        return pymysql.connect(
            host=cfg.get('host', '127.0.0.1'), port=int(cfg.get('port', 3306)),
            user=cfg.get('user', 'root'), password=cfg.get('password', ''),
            database=cfg.get('database', 'renee_orders'),
            charset='utf8mb4', autocommit=autocommit,
            connect_timeout=15, read_timeout=900, write_timeout=900,
            **mysql_ssl(cfg))

    BATCH = 1000
    t0 = time.time()
    try:
        sconn = _c(src, True)                     # TiDB source (TLS via mysql_ssl)
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f'Cannot connect to TiDB (source): {e}'}
    try:
        dconn = _c(dst, False)                    # local MySQL destination
    except Exception as e:  # noqa: BLE001
        sconn.close()
        return {'ok': False, 'error': f'Cannot connect to local MySQL (target): {e}'}

    out = {'ok': True, 'tables': [], 'total_rows': 0, 'views': 0,
           'source': f"{src.get('host')}:{src.get('port')}/{src.get('database')}",
           'target': f"{dst.get('host')}:{dst.get('port')}/{dst.get('database')}"}
    try:
        scur, dcur = sconn.cursor(), dconn.cursor()
        scur.execute("SELECT table_name, table_type FROM information_schema.tables "
                     "WHERE table_schema = DATABASE() ORDER BY table_name")
        base, views = [], []
        for name, ttype in scur.fetchall():
            (views if str(ttype).upper() == 'VIEW' else base).append(name)
        dcur.execute('SET FOREIGN_KEY_CHECKS=0')
        for t in base:
            scur.execute(f"SHOW CREATE TABLE `{t}`")
            ddl = scur.fetchone()[1]
            dcur.execute(f"DROP TABLE IF EXISTS `{t}`")
            dcur.execute(ddl)
            scur.execute(f"SELECT * FROM `{t}`")
            cols = [d[0] for d in scur.description]
            collist = ', '.join(f'`{c}`' for c in cols)
            ph = ', '.join(['%s'] * len(cols))
            ins = f"INSERT INTO `{t}` ({collist}) VALUES ({ph})"
            n = 0
            while True:
                rows = scur.fetchmany(BATCH)
                if not rows:
                    break
                dcur.executemany(ins, rows)
                n += len(rows)
            dconn.commit()
            out['tables'].append({'table': t, 'rows': n})
            out['total_rows'] += n
        for v in views:
            scur.execute(f"SHOW CREATE VIEW `{v}`")
            vddl = scur.fetchone()[1]
            vddl = re.sub(r'DEFINER=`[^`]*`@`[^`]*`\s*', '', vddl)
            vddl = re.sub(r'SQL SECURITY DEFINER', 'SQL SECURITY INVOKER', vddl)
            dcur.execute(f"DROP VIEW IF EXISTS `{v}`")
            dcur.execute(vddl)
            out['views'] += 1
        dcur.execute('SET FOREIGN_KEY_CHECKS=1')
        dconn.commit()
    except Exception as e:  # noqa: BLE001
        try:
            dconn.rollback()
        except Exception:  # noqa: BLE001
            pass
        return {'ok': False, 'error': f'{type(e).__name__}: {e}',
                'copied_tables': len(out['tables']), 'rows_so_far': out['total_rows']}
    finally:
        for c in (dconn, sconn):
            try:
                c.close()
            except Exception:  # noqa: BLE001
                pass
    out['n_tables'] = len(out['tables'])
    out['elapsed'] = round(time.time() - t0, 1)
    # Record this successful backup so the Setup card can show "last backup at ...".
    try:
        _LAST_BACKUP.parent.mkdir(parents=True, exist_ok=True)
        _LAST_BACKUP.write_text(json.dumps({
            'at': time.strftime('%Y-%m-%d %H:%M:%S'),
            'tables': out['n_tables'], 'rows': out['total_rows'],
            'views': out['views'], 'elapsed': out['elapsed'],
            'source': out['source'], 'target': out['target'],
        }), encoding='utf-8')
    except Exception:  # noqa: BLE001 — never fail the backup over the marker
        pass
    return out


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
