"""
online_b2b.services.app_settings
================================

Tiny **key/value settings store** for operator-tunable app behaviour — the data
layer behind the staff **Admin** page. One row per setting in ``app_settings``
(web-owned; never touches the business tables). Values are stored as strings;
helpers coerce to bool. API-ready: :func:`all_settings` returns a plain dict.

Add a new toggle by appending to :data:`DEFAULTS` (key → default, label, help) —
the Admin page renders every entry automatically, no template edits needed.
"""
from __future__ import annotations

from .order_db import _conn

_TABLE = 'app_settings'

# Every known setting: key → (default_bool, label, help-text). Adding a row here
# is all it takes to surface a new toggle on the Admin page.
DEFAULTS: dict[str, tuple] = {
    'auto_download_completed': (
        False,
        'Auto-download Completed workbook',
        'After you Lock & Record on the review page, the Completed SO Workbook '
        'download starts automatically. The manual download button stays as-is.'),
}

_CREATE = (
    f"CREATE TABLE IF NOT EXISTS {_TABLE} ("
    " skey VARCHAR(60) PRIMARY KEY,"
    " sval VARCHAR(255),"
    " updated_at DATETIME,"
    " updated_by VARCHAR(150))")


def ensure_table() -> None:
    with _conn() as (cur, d):
        cur.execute(_CREATE)
        cur.connection.commit()


def _truthy(v) -> bool:
    return str(v).strip().lower() in ('1', 'true', 'on', 'yes')


def get_bool(key: str, default: bool | None = None) -> bool:
    """Value of a boolean setting; falls back to the DEFAULTS entry (or ``default``)."""
    if default is None:
        default = bool(DEFAULTS.get(key, (False,))[0])
    ensure_table()
    with _conn() as (cur, d):
        ph = d['ph']
        cur.execute(f"SELECT sval FROM {_TABLE} WHERE skey={ph}", (key,))
        row = cur.fetchone()
    return _truthy(row[0]) if row else default


def set_value(key: str, value, user: str = '') -> dict:
    """Upsert one setting. Booleans are normalised to '1' / '0'."""
    import datetime as _dt
    if isinstance(value, bool):
        value = '1' if value else '0'
    ensure_table()
    with _conn() as (cur, d):
        ph = d['ph']
        now = _dt.datetime.now()
        # Portable upsert: try UPDATE, INSERT if nothing matched.
        cur.execute(
            f"UPDATE {_TABLE} SET sval={ph}, updated_at={ph}, updated_by={ph} "
            f"WHERE skey={ph}", (str(value), now, user, key))
        if cur.rowcount == 0:
            cur.execute(
                f"INSERT INTO {_TABLE} (skey, sval, updated_at, updated_by) "
                f"VALUES ({ph},{ph},{ph},{ph})", (key, str(value), now, user))
        cur.connection.commit()
    return {'ok': True, 'key': key, 'value': str(value)}


def all_settings() -> list[dict]:
    """Every known setting merged with its stored value — for the Admin page.
    ``[{key, label, help, value(bool), updated_at, updated_by}, ...]``."""
    ensure_table()
    stored: dict = {}
    with _conn() as (cur, d):
        cur.execute(f"SELECT skey, sval, updated_at, updated_by FROM {_TABLE}")
        for k, v, at, by in cur.fetchall():
            stored[k] = (v, at, by)
    out = []
    for key, (dflt, label, helptext) in DEFAULTS.items():
        v, at, by = stored.get(key, (None, None, None))
        out.append({
            'key': key, 'label': label, 'help': helptext,
            'value': _truthy(v) if v is not None else bool(dflt),
            'updated_at': at.strftime('%Y-%m-%d %H:%M') if at else '',
            'updated_by': by or ''})
    return out
