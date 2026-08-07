#!/usr/bin/env python
"""One-time migration — copy the LOCAL MySQL database into the TiDB server
(schema + data), so the server starts as an exact copy of your current working DB.

Pure PyMySQL (no mysqldump / MySQL client tools needed). Reads two profiles:
  db_profiles/local.json  (source — your current local MySQL)
  db_profiles/tidb.json   (target — the TiDB server, with ssl)

Copies every BASE TABLE (DROP + CREATE via SHOW CREATE TABLE, then all rows) and
re-creates VIEWS last (e.g. order_lines_full). Idempotent — safe to re-run
(it replaces the target tables each time).

Usage:
  python db_push_to_tidb.py            # DRY RUN — list tables + row counts, no writes
  python db_push_to_tidb.py --push     # actually copy local -> TiDB
"""
import json
import re
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(_ROOT / 'online_po_management'))
from online_po_processor.auto.history_db import mysql_ssl  # noqa: E402

PROFILES = _ROOT / 'db_profiles'
BATCH = 1000


def _load(name):
    p = PROFILES / f'{name}.json'
    if not p.exists():
        sys.exit(f'Missing profile {p} — create it from db_profiles/{name}.json.example')
    return json.loads(p.read_text(encoding='utf-8-sig'))


def _connect(cfg, autocommit):
    import pymysql
    return pymysql.connect(
        host=cfg.get('host', '127.0.0.1'), port=int(cfg.get('port', 3306)),
        user=cfg.get('user', 'root'), password=cfg.get('password', ''),
        database=cfg.get('database', 'renee_orders'),
        charset='utf8mb4', autocommit=autocommit, **mysql_ssl(cfg))


def _tables(cur):
    cur.execute("SELECT table_name, table_type FROM information_schema.tables "
                "WHERE table_schema = DATABASE() ORDER BY table_name")
    base, views = [], []
    for name, ttype in cur.fetchall():
        (views if str(ttype).upper() == 'VIEW' else base).append(name)
    return base, views


def _rowcount(cur, t):
    cur.execute(f"SELECT COUNT(*) FROM `{t}`")
    return cur.fetchone()[0]


def main():
    push = '--push' in sys.argv
    src, dst = _load('local'), _load('tidb')
    print(f"SOURCE (local): {src.get('host')}:{src.get('port')}/{src.get('database')}")
    print(f"TARGET (tidb) : {dst.get('host')}:{dst.get('port')}/{dst.get('database')}  [TLS]")
    print('MODE:', 'PUSH (writes to TiDB)' if push else 'DRY RUN (no writes)')
    print('-' * 64)

    sconn = _connect(src, autocommit=True)
    scur = sconn.cursor()
    base, views = _tables(scur)
    print(f'{len(base)} base tables, {len(views)} views')

    if not push:
        for t in base:
            print(f'  {t:<28} {_rowcount(scur, t):>10,} rows')
        for v in views:
            print(f'  {v:<28} {"(view)":>10}')
        print('\nDRY RUN only. Re-run with --push to copy into TiDB.')
        sconn.close()
        return

    dconn = _connect(dst, autocommit=False)
    dcur = dconn.cursor()
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
        print(f'  [ok] {t:<28} {n:>10,} rows')

    for v in views:
        scur.execute(f"SHOW CREATE VIEW `{v}`")
        vddl = scur.fetchone()[1]
        vddl = re.sub(r'DEFINER=`[^`]*`@`[^`]*`\s*', '', vddl)   # strip DEFINER
        vddl = re.sub(r'SQL SECURITY DEFINER', 'SQL SECURITY INVOKER', vddl)
        dcur.execute(f"DROP VIEW IF EXISTS `{v}`")
        dcur.execute(vddl)
        print(f'  [ok] {v:<28} {"(view)":>10}')

    dcur.execute('SET FOREIGN_KEY_CHECKS=1')
    dconn.commit()
    dconn.close()
    sconn.close()
    print('\nDone — TiDB now mirrors local. Point the app at TiDB with '
          '"python db_switch.py tidb".')


if __name__ == '__main__':
    main()
