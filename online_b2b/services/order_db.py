"""
online_b2b.services.order_db
============================

Read-only access to the **order history DB** for the web dashboard.

The web app is now the OWNER of the order store: MySQL ``renee_orders``
(tables ``runs`` / ``order_headers`` + the 2-table line split
``order_lines`` / ``order_line_validation`` read via the ``order_lines_full``
view). Writes go through ``services.lines_store`` (``record_run_headers`` +
``insert_lines``) — web-owned replicas of the engine's ``record_manual`` /
``apply_dedup``, so the engine's history store is never invoked and the legacy
desktop ``order_issue_lines`` table is gone.

We open our own short-lived connection here and only ever ``SELECT``. The
backend choice and credentials come from the engine's own ``load_db_config``
(a local file kept out of the repo), so nothing is duplicated or hard-coded.
"""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager


def _backend():
    """Return ('mysql', cfg) or ('sqlite', path) based on the engine config."""
    from online_po_processor.auto.history_db import (
        default_history_db_path,
        load_db_config,
    )
    cfg = load_db_config()
    if cfg and str(cfg.get('backend', '')).lower() == 'mysql':
        return 'mysql', cfg
    return 'sqlite', default_history_db_path()


@contextmanager
def _conn():
    """Yield (cursor, dialect) where dialect carries the placeholder + the
    backend-specific order table name. Read-only; always closed."""
    kind, target = _backend()
    if kind == 'mysql':
        import pymysql
        c = pymysql.connect(
            host=target.get('host', '127.0.0.1'),
            port=int(target.get('port', 3306)),
            user=target.get('user', 'root'),
            password=target.get('password', ''),
            database=target.get('database', 'renee_orders'),
            charset='utf8mb4', autocommit=True,
        )
        dialect = {'ph': '%s', 'orders': 'order_headers', 'kind': 'mysql'}
        try:
            yield c.cursor(), dialect
        finally:
            c.close()
    else:
        c = sqlite3.connect(str(target))
        dialect = {'ph': '?', 'orders': 'orders', 'kind': 'sqlite'}
        try:
            yield c.cursor(), dialect
        finally:
            c.close()


def _rows(cur, cols: list[str]) -> list[dict]:
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def _tag_basis(rows: list[dict]) -> list[dict]:
    """Derive each line's comparison basis ('CP' / 'Landing') from which vendor
    value is present (cost-based marketplaces carry vendor_cp; landing-based —
    e.g. Flipkart — carry vendor_landing). Used to highlight the basis pair."""
    for r in rows:
        r['basis'] = ('CP' if r.get('vendor_cp') is not None
                      else 'Landing' if r.get('vendor_landing') is not None
                      else '')
    return rows


def db_available() -> bool:
    try:
        with _conn() as (cur, _):
            cur.execute("SELECT 1")
            cur.fetchone()
        return True
    except Exception:
        return False


def backend_label() -> str:
    try:
        kind, target = _backend()
        if kind == 'mysql':
            return f"MySQL · {target.get('database', 'renee_orders')}@{target.get('host')}"
        return f"SQLite · {target.name}"
    except Exception:
        return "unavailable"


# Unified order-management dashboard: ALL segments (Online B2B + Offline channels
# like Shoppers Stop / GT Mass / EKA) show together — they're all order-mgmt tools.
# '' = no segment filter. The Marketplace column distinguishes channels.
SEGMENT = ''
RECENT_DAYS = 2          # window for the "Updated (last 2d)" KPI


EXPIRY_SOON_DAYS = 7     # window for the "Expiring soon" KPI / row pill
PAGE_SIZE = 50           # orders per page (Load more appends another page)

# Sort key → SQL column (allow-list; never interpolate user input directly).
_SORT_COLS = {
    'date': 'run_ts', 'po': 'po', 'marketplace': 'marketplace_label',
    'qty': 'qty', 'value': 'order_value', 'items': 'items', 'expiry': 'exp_date',
}


def _cutoff(kind: str, days: int):
    """Return a run_ts lower-bound param for the given backend."""
    import datetime as _dt
    dt = _dt.datetime.now() - _dt.timedelta(days=days)
    return dt if kind == 'mysql' else dt.isoformat(sep=' ', timespec='seconds')


# Known segments for the dashboard switch.
SEGMENTS = ['OnlineB2B', 'Offline']


def _seg(ph: str, segment):
    """Segment WHERE fragment + params. '' / 'all' → no segment filter."""
    if segment and segment != 'all':
        return f"segment={ph}", [segment]
    return "1=1", []


_SEG_LABEL = {'OnlineB2B': 'Online B2B', 'Offline': 'Offline'}


def daily_intake(days: int = 30) -> dict:
    """Per-day order arrivals (by ``created_at``) split by segment — for the
    management daily stacked chart. Returns chart-ready arrays (gaps filled with
    0). Read-only; never raises."""
    import datetime as _dt
    out = {'labels': [], 'segments': [], 'value': {}, 'pos': {}, 'items': {}, 'qty': {}}
    try:
        with _conn() as (cur, d):
            ot = d['orders']
            cur.execute(
                f"SELECT DATE(created_at), segment, COUNT(DISTINCT po), "
                f"COALESCE(SUM(order_value),0), COALESCE(SUM(items),0), "
                f"COALESCE(SUM(qty),0) "
                f"FROM {ot} WHERE created_at >= (CURDATE() - INTERVAL {int(days) - 1} DAY) "
                f"GROUP BY DATE(created_at), segment")
            cells, segs = {}, []
            for dd, seg, pos, val, items, qty in cur.fetchall():
                seg = _SEG_LABEL.get(seg, seg or 'Other')
                if seg not in segs:
                    segs.append(seg)
                cells[(str(dd), seg)] = (pos or 0, float(val or 0), int(items or 0), int(qty or 0))
            today = _dt.date.today()
            day_list = [today - _dt.timedelta(days=i) for i in range(int(days) - 1, -1, -1)]
            value = {s: [] for s in segs}
            pos = {s: [] for s in segs}
            items = {s: [] for s in segs}
            qty = {s: [] for s in segs}
            for dd in day_list:
                ds = dd.isoformat()
                for s in segs:
                    p, v, it, q = cells.get((ds, s), (0, 0.0, 0, 0))
                    value[s].append(round(v, 2))
                    pos[s].append(p)
                    items[s].append(it)
                    qty[s].append(q)
            labels = [dd.strftime('%d %b') for dd in day_list]
            # Trim leading all-zero days so the bars fill the chart from the left
            # instead of clustering at the right edge (orders may start mid-window).
            start = 0
            for i in range(len(labels)):
                if any((value[s][i] or pos[s][i]) for s in segs):
                    start = i
                    break
            else:
                start = max(0, len(labels) - 1)
            if start > 0:
                labels = labels[start:]
                for s in segs:
                    value[s] = value[s][start:]
                    pos[s] = pos[s][start:]
                    items[s] = items[s][start:]
                    qty[s] = qty[s][start:]
            out = {'labels': labels, 'segments': segs,
                   'value': value, 'pos': pos, 'items': items, 'qty': qty}
    except Exception:  # noqa: BLE001
        pass
    return out


def intake_hierarchy(days: int = 30, date: str = '') -> dict:
    """segment → parent marketplace → child breakdown (pos/value/items) for the
    management tree. ``date`` (YYYY-MM-DD) scopes to a single day; otherwise the
    last N days. Read-only; never raises."""
    out = {'segments': [], 'total': {'pos': 0, 'value': 0.0, 'items': 0}}
    try:
        with _conn() as (cur, d):
            ot, ph = d['orders'], d['ph']
            if date:
                wsql, params = f"DATE(created_at) = {ph}", (date,)
            else:
                wsql, params = (f"created_at >= (CURDATE() - INTERVAL "
                                f"{int(days) - 1} DAY)", ())
            cur.execute(
                f"SELECT segment, marketplace, marketplace_label, COUNT(DISTINCT po), "
                f"COALESCE(SUM(order_value),0), COALESCE(SUM(items),0), "
                f"COALESCE(SUM(qty),0) "
                f"FROM {ot} WHERE {wsql} "
                f"GROUP BY segment, marketplace, marketplace_label", params)
            tree: dict = {}
            tp = tv = ti = tq = 0
            for seg, mkt, label, pos, val, items, qty in cur.fetchall():
                seg = _SEG_LABEL.get(seg, seg or 'Other')
                mkt = mkt or 'Other'
                label = label or mkt
                pos = pos or 0
                val = float(val or 0)
                items = int(items or 0)
                qty = int(qty or 0)
                s = tree.setdefault(seg, {'pos': 0, 'value': 0.0, 'items': 0, 'qty': 0, 'mkts': {}})
                m = s['mkts'].setdefault(mkt, {'pos': 0, 'value': 0.0, 'items': 0, 'qty': 0, 'children': []})
                m['children'].append({'label': label, 'pos': pos, 'value': round(val, 2),
                                      'items': items, 'qty': qty})
                m['pos'] += pos; m['value'] += val; m['items'] += items; m['qty'] += qty
                s['pos'] += pos; s['value'] += val; s['items'] += items; s['qty'] += qty
                tp += pos; tv += val; ti += items; tq += qty

            def _bar(part, whole):
                return round(part / whole * 100, 1) if whole else 0

            segments = []
            for seg, s in tree.items():
                sval = s['value'] or 1
                mkts = []
                for k, v in sorted(s['mkts'].items(), key=lambda kv: -kv[1]['value']):
                    mval = v['value'] or 1
                    kids = sorted(v['children'], key=lambda c: -c['value'])
                    for c in kids:
                        c['bar'] = _bar(c['value'], mval)
                    mkts.append({'marketplace': k, 'pos': v['pos'], 'value': round(v['value'], 2),
                                 'items': v['items'], 'qty': v['qty'],
                                 'multi': len(v['children']) > 1, 'children': kids,
                                 'bar': _bar(v['value'], sval)})
                segments.append({'segment': seg, 'pos': s['pos'], 'value': round(s['value'], 2),
                                 'items': s['items'], 'qty': s['qty'], 'marketplaces': mkts,
                                 'bar': _bar(s['value'], tv)})
            segments.sort(key=lambda x: -x['value'])
            out = {'segments': segments,
                   'total': {'pos': tp, 'value': round(tv, 2), 'items': ti, 'qty': tq}}
    except Exception:  # noqa: BLE001
        pass
    return out


def hub_extra_kpis() -> dict:
    """Extra hub KPIs in one round-trip: avg PO value, total line items, POs in the
    last 7 days, and resolved (actioned) issue lines. Read-only; never raises."""
    out = {'avg_po_value': 0.0, 'order_lines': 0, 'week_pos': 0, 'resolved': 0}
    try:
        with _conn() as (cur, d):
            ot = d['orders']
            cur.execute(f"SELECT COUNT(DISTINCT po), COALESCE(SUM(order_value),0) FROM {ot}")
            pos, val = cur.fetchone()
            out['avg_po_value'] = round(float(val or 0) / pos, 2) if pos else 0.0
            cur.execute(f"SELECT COUNT(DISTINCT po) FROM {ot} "
                        f"WHERE created_at >= (CURDATE() - INTERVAL 6 DAY)")
            out['week_pos'] = cur.fetchone()[0] or 0
            try:
                cur.execute("SELECT COUNT(*) FROM order_lines_full")
                out['order_lines'] = cur.fetchone()[0] or 0
                cur.execute("SELECT COUNT(*) FROM order_lines_full "
                            "WHERE action IS NOT NULL AND action <> ''")
                out['resolved'] = cur.fetchone()[0] or 0
            except Exception:  # noqa: BLE001 — order_lines may be empty/new
                pass
    except Exception:  # noqa: BLE001
        pass
    return out


def recent_orders(limit: int = 8) -> list:
    """Latest N orders across ALL channels (by created_at) for the hub's recent
    activity feed. Read-only; never raises."""
    out: list = []
    try:
        with _conn() as (cur, d):
            ot = d['orders']
            cur.execute(
                f"SELECT marketplace_label, segment, po, qty, "
                f"COALESCE(order_value,0), run_id, created_at "
                f"FROM {ot} ORDER BY created_at DESC, order_id DESC "
                f"LIMIT {int(limit)}")
            for r in cur.fetchall():
                out.append({
                    'marketplace': r[0] or r[1] or '',
                    'segment': r[1] or '',
                    'po': r[2] or '',
                    'qty': int(r[3] or 0),
                    'value': float(r[4] or 0),
                    'run_id': r[5],
                    'when': r[6],
                })
    except Exception:  # noqa: BLE001
        pass
    return out


def recent_runs(limit: int = 8) -> list:
    """Latest N upload RUNS (batches) across all channels — when it was uploaded,
    which marketplace(s), and what was uploaded amount-wise (#POs, qty, value).
    One row per run_id. Read-only; never raises."""
    out: list = []
    try:
        with _conn() as (cur, d):
            ot = d['orders']
            cur.execute(
                f"SELECT run_id, MAX(run_ts), GROUP_CONCAT(DISTINCT marketplace_label), "
                f"MAX(segment), COUNT(DISTINCT po), COALESCE(SUM(qty),0), "
                f"COALESCE(SUM(order_value),0) FROM {ot} "
                f"GROUP BY run_id ORDER BY run_id DESC LIMIT {int(limit)}")
            for r in cur.fetchall():
                mps = [m for m in (r[2] or '').split(',') if m]
                label = (mps[0] if len(mps) <= 1 else f"{mps[0]} +{len(mps) - 1}")
                out.append({
                    'run_id': r[0], 'when': r[1],
                    'marketplace': label or (r[3] or ''),
                    'marketplaces': r[2] or '',
                    'segment': r[3] or '',
                    'pos': int(r[4] or 0),
                    'qty': int(r[5] or 0),
                    'value': float(r[6] or 0),
                })
    except Exception:  # noqa: BLE001
        pass
    return out


def today_intake() -> dict:
    """Orders RECEIVED today — filtered by ``created_at`` (when recorded, not PO
    date). Returns total PO count + value plus a per-segment split for the hub
    'Received Today' card and its hover distribution. Read-only; never raises."""
    _label = {'OnlineB2B': 'Online B2B', 'Offline': 'Offline'}
    out = {'pos': 0, 'value': 0.0, 'by_segment': []}
    try:
        with _conn() as (cur, d):
            ot = d['orders']
            cur.execute(
                f"SELECT segment, COUNT(DISTINCT po), COALESCE(SUM(order_value),0) "
                f"FROM {ot} WHERE DATE(created_at)=CURDATE() "
                f"GROUP BY segment ORDER BY 3 DESC")
            segs, tot_pos, tot_val = [], 0, 0.0
            for s, p, v in cur.fetchall():
                p = p or 0
                v = float(v or 0)
                segs.append({'segment': _label.get(s, s or 'Other'),
                             'pos': p, 'value': v})
                tot_pos += p
                tot_val += v
            out = {'pos': tot_pos, 'value': round(tot_val, 2), 'by_segment': segs}
    except Exception:  # noqa: BLE001
        pass
    return out


def segment_kpis(segment: str) -> dict:
    """Lightweight POs / qty / value for one segment — for the hub group cards.
    Read-only; never raises (returns zeros on any error)."""
    out = {'pos': 0, 'qty': 0, 'value': 0.0}
    try:
        with _conn() as (cur, d):
            ot, ph = d['orders'], d['ph']
            seg, params = _seg(ph, segment)
            cur.execute(
                f"SELECT COUNT(DISTINCT po), COALESCE(SUM(qty),0), "
                f"COALESCE(SUM(order_value),0) FROM {ot} WHERE {seg}", tuple(params))
            r = cur.fetchone()
            out = {'pos': r[0] or 0, 'qty': int(r[1] or 0),
                   'value': float(r[2] or 0)}
    except Exception:  # noqa: BLE001
        pass
    return out


def _parse_date(v):
    import datetime as _dt
    if v is None or v == '':
        return None
    if isinstance(v, _dt.date):
        return v
    s = str(v)[:10]
    for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y'):
        try:
            return _dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _days_to_expiry(exp):
    import datetime as _dt
    d = _parse_date(exp)
    return (d - _dt.date.today()).days if d else None


def _norm_filters(f: dict) -> dict:
    return {
        'segment': (f.get('segment') or SEGMENT).strip(),
        'marketplace': (f.get('marketplace') or '').strip(),
        'days': int(f.get('days') or 0),
        'q': (f.get('q') or '').strip(),
        'warehouse': (f.get('warehouse') or '').strip(),
        'order_type': (f.get('order_type') or '').strip(),
        'date_from': (f.get('date_from') or '').strip(),
        'date_to': (f.get('date_to') or '').strip(),
    }


def _where(d: dict, f: dict):
    """Build the shared WHERE clause + params from a filters dict (segment-aware)."""
    ph, kind = d['ph'], d['kind']
    seg_sql, params = _seg(ph, f.get('segment') or SEGMENT)
    where = [seg_sql]
    if f['marketplace']:
        where.append(f"marketplace_label={ph}"); params.append(f['marketplace'])
    if f['days'] > 0:
        where.append(f"run_ts >= {ph}"); params.append(_cutoff(kind, f['days']))
    if f['q']:
        where.append(f"po LIKE {ph}"); params.append(f"%{f['q']}%")
    if f['warehouse']:
        where.append(f"warehouse={ph}"); params.append(f['warehouse'])
    if f['order_type']:
        where.append(f"order_type={ph}"); params.append(f['order_type'])
    if f['date_from']:
        where.append(f"po_date >= {ph}"); params.append(f['date_from'])
    if f['date_to']:
        where.append(f"po_date <= {ph}"); params.append(f['date_to'])
    return " AND ".join(where), params


def _order_cols(kind: str) -> list[str]:
    return ['run_id', 'run_ts', 'marketplace', 'po', 'location', 'warehouse',
            'po_date', 'exp_date', 'order_type', 'items', 'qty', 'value']


def _fetch_orders(cur, d, f, sort='date', direction='desc',
                  limit=PAGE_SIZE, offset=0):
    ot, kind = d['orders'], d['kind']
    wsql, params = _where(d, f)
    oid = 'order_id' if kind == 'mysql' else 'id'
    col = _SORT_COLS.get(sort, 'run_ts')
    dirn = 'ASC' if str(direction).lower() == 'asc' else 'DESC'
    cur.execute(
        f"SELECT run_id, run_ts, marketplace_label, po, location, warehouse, "
        f"po_date, exp_date, order_type, items, qty, order_value FROM {ot} "
        f"WHERE {wsql} ORDER BY {col} {dirn}, {oid} DESC "
        f"LIMIT {int(limit)} OFFSET {int(offset)}", tuple(params))
    rows = _rows(cur, _order_cols(kind))
    for r in rows:
        r['days_to_expiry'] = _days_to_expiry(r.get('exp_date'))
    return rows


def _count_orders(cur, d, f) -> int:
    wsql, params = _where(d, f)
    cur.execute(f"SELECT COUNT(*) FROM {d['orders']} WHERE {wsql}", tuple(params))
    return cur.fetchone()[0]


# ── Dashboard ───────────────────────────────────────────────────────────

def dashboard(segment='', marketplace='', days=0, q='', warehouse='', order_type='',
              date_from='', date_to='', sort='date', direction='desc',
              offset=0, limit=PAGE_SIZE) -> dict:
    """Online-B2B order dashboard read straight from the order DB. The filter
    args refine the Orders table + per-marketplace rollup; the headline KPIs +
    trend chart are always the full Online-B2B picture."""
    f = _norm_filters(locals())
    out: dict = {
        'ok': False, 'backend': backend_label(), 'kpis': {},
        'by_marketplace': [], 'orders': [], 'issue_count': 0,
        'marketplace_options': [], 'warehouse_options': [], 'type_options': [],
        'segments': SEGMENTS, 'trends': {}, 'total': 0, 'offset': offset,
        'limit': limit, 'sort': sort, 'direction': direction,
        'filters': dict(f, sort=sort, direction=direction),
    }
    try:
        with _conn() as (cur, d):
            ot, ph, kind = d['orders'], d['ph'], d['kind']
            seg, sp = _seg(ph, f['segment'])

            out['kpis'] = _kpis(cur, d, f['segment'])
            out['issue_count'] = out['kpis'].get('issue_lines', 0)
            out['trends'] = _trends(cur, d, 30, f['segment'])

            # Dropdown options (scoped to the selected segment)
            for key, colname in (('marketplace_options', 'marketplace_label'),
                                 ('warehouse_options', 'warehouse'),
                                 ('type_options', 'order_type')):
                cur.execute(f"SELECT DISTINCT {colname} FROM {ot} WHERE {seg} "
                            f"AND {colname} IS NOT NULL AND {colname} <> '' "
                            f"ORDER BY {colname}", tuple(sp))
                out[key] = [r[0] for r in cur.fetchall()]

            # Per-marketplace rollup (respects the filters)
            wsql, params = _where(d, f)
            cur.execute(
                f"SELECT marketplace_label, COUNT(DISTINCT po), "
                f"COALESCE(SUM(qty),0), COALESCE(SUM(order_value),0) "
                f"FROM {ot} WHERE {wsql} GROUP BY marketplace_label "
                f"ORDER BY 4 DESC", tuple(params))
            bm = _rows(cur, ['marketplace', 'pos', 'qty', 'value'])
            mx = max([m['value'] for m in bm], default=0) or 1
            sparks = _mp_sparklines(cur, d, 30, f['segment'])
            for m in bm:
                m['bar'] = round(float(m['value']) / float(mx) * 100, 1)
                m['spark'] = sparks.get(m['marketplace'], '')
            out['by_marketplace'] = bm

            out['total'] = _count_orders(cur, d, f)
            out['orders'] = _fetch_orders(cur, d, f, sort, direction,
                                          limit, offset)

            # Chart payloads as plain JSON-safe numbers (for ApexCharts).
            series = out['trends'].get('series', [])
            out['charts'] = {
                'trend': {
                    'labels': [s['label'] for s in series],
                    'value': [round(float(s['value']), 2) for s in series],
                    'orders': [int(s['orders']) for s in series],
                },
                'mix': {
                    'labels': [m['marketplace'] for m in bm[:8]],
                    'value': [round(float(m['value']), 2) for m in bm[:8]],
                },
            }
        out['ok'] = True
    except Exception as e:  # noqa: BLE001 — dashboard must never 500
        out['error'] = f"{type(e).__name__}: {e}"
    return out


def overview(segment=SEGMENT) -> dict:
    """Lean landing-page data: KPIs + 30-day trend + marketplace mix/summary
    (no order rows — the full list lives on the Orders page). ``segment`` scopes
    everything ('OnlineB2B' default, 'Offline', or 'all')."""
    out: dict = {
        'ok': False, 'backend': backend_label(), 'kpis': {},
        'by_marketplace': [], 'charts': {}, 'issue_count': 0, 'trends': {},
        'segments': SEGMENTS, 'segment': segment,
    }
    try:
        with _conn() as (cur, d):
            ot, ph = d['orders'], d['ph']
            seg, sp = _seg(ph, segment)
            out['kpis'] = _kpis(cur, d, segment)
            out['issue_count'] = out['kpis'].get('issue_lines', 0)
            out['trends'] = _trends(cur, d, 30, segment)

            cur.execute(
                f"SELECT marketplace_label, COUNT(DISTINCT po), "
                f"COALESCE(SUM(qty),0), COALESCE(SUM(order_value),0) "
                f"FROM {ot} WHERE {seg} GROUP BY marketplace_label "
                f"ORDER BY 4 DESC", tuple(sp))
            bm = _rows(cur, ['marketplace', 'pos', 'qty', 'value'])
            mx = max([m['value'] for m in bm], default=0) or 1
            sparks = _mp_sparklines(cur, d, 30, segment)
            for m in bm:
                m['bar'] = round(float(m['value']) / float(mx) * 100, 1)
                m['spark'] = sparks.get(m['marketplace'], '')
            out['by_marketplace'] = bm

            series = out['trends'].get('series', [])
            out['charts'] = {
                'trend': {
                    'labels': [s['label'] for s in series],
                    'value': [round(float(s['value']), 2) for s in series],
                    'orders': [int(s['orders']) for s in series],
                },
                'mix': {
                    'labels': [m['marketplace'] for m in bm[:8]],
                    'value': [round(float(m['value']), 2) for m in bm[:8]],
                },
            }
        out['ok'] = True
    except Exception as e:  # noqa: BLE001
        out['error'] = f"{type(e).__name__}: {e}"
    return out


def orders_page(sort='date', direction='desc', offset=0, limit=PAGE_SIZE,
                **filters) -> dict:
    """Just the next page of order rows (for 'Load more') + whether more exist."""
    f = _norm_filters(filters)
    out = {'ok': False, 'orders': [], 'has_more': False, 'next_offset': offset}
    try:
        with _conn() as (cur, d):
            total = _count_orders(cur, d, f)
            rows = _fetch_orders(cur, d, f, sort, direction, limit, offset)
            out['orders'] = rows
            out['next_offset'] = offset + len(rows)
            out['has_more'] = out['next_offset'] < total
        out['ok'] = True
    except Exception as e:  # noqa: BLE001
        out['error'] = f"{type(e).__name__}: {e}"
    return out


def _kpis(cur, d, segment=SEGMENT) -> dict:
    ot, ph, kind = d['orders'], d['ph'], d['kind']
    seg, sp = _seg(ph, segment)

    cur.execute(
        f"SELECT COUNT(DISTINCT po), COUNT(*), COALESCE(SUM(qty),0), "
        f"COALESCE(SUM(order_value),0), COUNT(DISTINCT marketplace_label) "
        f"FROM {ot} WHERE {seg}", tuple(sp))
    n_pos, n_lines, tot_qty, tot_val, n_mp = cur.fetchone()

    cur.execute(f"SELECT COUNT(DISTINCT po) FROM {ot} WHERE {seg} "
                f"AND run_ts >= {ph}", tuple(sp + [_cutoff(kind, RECENT_DAYS)]))
    updated_2d = cur.fetchone()[0]

    cur.execute(f"SELECT MAX(run_ts) FROM {ot} WHERE {seg}", tuple(sp))
    last_updated = cur.fetchone()[0]

    # Expiring soon (mysql DATE math; best-effort 0 on sqlite)
    expiring = 0
    if kind == 'mysql':
        try:
            cur.execute(
                f"SELECT COUNT(DISTINCT po) FROM {ot} WHERE {seg} "
                f"AND exp_date IS NOT NULL AND exp_date >= CURDATE() "
                f"AND exp_date <= DATE_ADD(CURDATE(), INTERVAL {EXPIRY_SOON_DAYS} DAY)",
                tuple(sp))
            expiring = cur.fetchone()[0]
        except Exception:
            expiring = 0

    # Affected lines + POs needing attention — derived from the SINGLE lines
    # table (status != OK), no separate issue-lines table.
    # Only UNRESOLVED (no action set) lines count as needs-attention / issues.
    n_issue = needs_attention = 0
    try:
        cur.execute("SELECT COUNT(*), COUNT(DISTINCT po) FROM order_lines_full "
                    "WHERE status IN ('MISMATCH','NOT_IN_MASTER') "
                    "AND (action IS NULL OR action = '')")
        n_issue, needs_attention = cur.fetchone()
    except Exception:
        pass

    k = {
        'pos': n_pos, 'lines': n_lines, 'qty': int(tot_qty or 0),
        'value': float(tot_val or 0.0), 'marketplaces': n_mp,
        'updated_2d': updated_2d, 'last_updated': last_updated,
        'expiring_soon': expiring, 'needs_attention': needs_attention,
        'issue_lines': n_issue,
    }
    k.update(_deltas(cur, d, segment))
    return k


def _deltas(cur, d, segment=SEGMENT) -> dict:
    """POs/value for the last 7 days vs the prior 7 days → % change."""
    ot, ph, kind = d['orders'], d['ph'], d['kind']
    seg, sp = _seg(ph, segment)

    def window(lo_days, hi_days):
        cur.execute(
            f"SELECT COUNT(DISTINCT po), COALESCE(SUM(order_value),0) FROM {ot} "
            f"WHERE {seg} AND run_ts >= {ph} AND run_ts < {ph}",
            tuple(sp + [_cutoff(kind, lo_days), _cutoff(kind, hi_days)]))
        return cur.fetchone()

    cur_pos, cur_val = window(7, 0)
    prev_pos, prev_val = window(14, 7)

    def pct(c, p):
        c, p = float(c or 0), float(p or 0)
        if p == 0:
            return None
        return round((c - p) / p * 100, 1)

    return {'pos_delta': pct(cur_pos, prev_pos),
            'value_delta': pct(cur_val, prev_val)}


def _trends(cur, d, days=30, segment=SEGMENT) -> dict:
    """Daily order count + value for the last ``days`` days (gaps filled), with
    pre-computed bar heights so the template can draw an SVG with no maths."""
    import datetime as _dt
    ot, ph, kind = d['orders'], d['ph'], d['kind']
    seg, sp = _seg(ph, segment)
    datefn = 'DATE(run_ts)'
    cur.execute(
        f"SELECT {datefn} AS dt, COUNT(DISTINCT po), COALESCE(SUM(order_value),0) "
        f"FROM {ot} WHERE {seg} AND run_ts >= {ph} GROUP BY dt ORDER BY dt",
        tuple(sp + [_cutoff(kind, days)]))
    raw = {}
    for dt, c, v in cur.fetchall():
        key = _parse_date(dt)
        if key:
            raw[key] = (int(c or 0), float(v or 0.0))

    today = _dt.date.today()
    series = []
    for i in range(days - 1, -1, -1):
        day = today - _dt.timedelta(days=i)
        c, v = raw.get(day, (0, 0.0))
        series.append({'date': day.isoformat(), 'label': day.strftime('%d %b'),
                       'orders': c, 'value': v})
    max_o = max([s['orders'] for s in series], default=0) or 1
    max_v = max([s['value'] for s in series], default=0) or 1
    for s in series:
        s['o_h'] = round(s['orders'] / max_o * 100, 1)
        s['v_h'] = round(s['value'] / max_v * 100, 1)

    # Pre-computed SVG area/line paths (viewBox 0..W × 0..H) for an elegant
    # sparkline — one set for Value, one for PO count (chart can toggle) — plus
    # transparent hover columns carrying a native tooltip.
    W, H, PAD = 1000.0, 100.0, 6.0
    n = len(series)

    def _paths(hkey):
        pts = []
        for i, s in enumerate(series):
            x = (i / (n - 1) * W) if n > 1 else 0.0
            y = PAD + (1 - s[hkey] / 100.0) * (H - PAD)
            pts.append((round(x, 1), round(y, 1)))
        ln = ('M ' + ' L '.join(f"{x} {y}" for x, y in pts)) if pts else ''
        ar = (ln + f" L {W} {H} L 0 {H} Z") if pts else ''
        return ln, ar

    line_v, area_v = _paths('v_h')
    line_o, area_o = _paths('o_h')
    colw = W / n if n else W
    cols = [{'x': round(i * colw, 1), 'w': round(colw, 1),
             'title': f"{s['label']} · {s['orders']} POs · ₹{s['value']:,.0f}"}
            for i, s in enumerate(series)]

    return {'series': series, 'max_orders': max_o, 'max_value': max_v,
            'total_orders': sum(s['orders'] for s in series),
            'total_value': sum(s['value'] for s in series),
            'svg': {'w': W, 'h': H, 'line': line_v, 'area': area_v,
                    'line_v': line_v, 'area_v': area_v,
                    'line_o': line_o, 'area_o': area_o, 'cols': cols}}


def _mp_sparklines(cur, d, days=30, segment=SEGMENT) -> dict[str, str]:
    """Tiny per-marketplace value sparklines (last ``days`` days) as SVG line
    paths in a 100×24 viewBox, keyed by marketplace_label."""
    import datetime as _dt
    ot, ph, kind = d['orders'], d['ph'], d['kind']
    seg, sp = _seg(ph, segment)
    cur.execute(
        f"SELECT marketplace_label, DATE(run_ts), COALESCE(SUM(order_value),0) "
        f"FROM {ot} WHERE {seg} AND run_ts >= {ph} GROUP BY 1, 2",
        tuple(sp + [_cutoff(kind, days)]))
    by: dict[str, dict] = {}
    for lbl, dt, v in cur.fetchall():
        key = _parse_date(dt)
        if key is not None:
            by.setdefault(lbl, {})[key] = float(v or 0.0)

    today = _dt.date.today()
    W, H, PAD = 100.0, 24.0, 2.0
    out: dict[str, str] = {}
    for lbl, dmap in by.items():
        vals = [dmap.get(today - _dt.timedelta(days=i), 0.0)
                for i in range(days - 1, -1, -1)]
        mx = max(vals) or 1.0
        nn = len(vals)
        pts = [(round(i / (nn - 1) * W, 1),
                round(PAD + (1 - v / mx) * (H - 2 * PAD), 1))
               for i, v in enumerate(vals)]
        out[lbl] = 'M ' + ' L '.join(f"{x} {y}" for x, y in pts)
    return out


# action set (any of KEEP/OVERRIDE/EXCLUDE) ⇒ the line is RESOLVED (handled).
# A line is RESOLVED if an action was set (mismatch decision) OR its EAN was
# corrected (received_ean set → a fixed NOT_IN_MASTER line). PENDING = neither.
_RESOLVED_SQL = ("((action IS NOT NULL AND action <> '') "
                 "OR (received_ean IS NOT NULL AND received_ean <> ''))")
_PENDING_SQL = ("(action IS NULL OR action = '') "
                "AND (received_ean IS NULL OR received_ean = '')")
# Belongs on the Issues page if currently affected OR an affected line that was
# EAN-corrected (received_ean ⇒ it *was* NOT_IN_MASTER).
_AFFECTED_SQL = ("(status IN ('MISMATCH','NOT_IN_MASTER') "
                 "OR (received_ean IS NOT NULL AND received_ean <> ''))")
_FIXED_EAN_SQL = "(received_ean IS NOT NULL AND received_ean <> '')"


def issues(marketplace='', q='', status='', resolution='pending',
           date_from='', date_to='', limit=300) -> dict:
    """Affected lines for the Issues page — price MISMATCH, NOT_IN_MASTER, AND
    EAN-corrected lines (now OK but were NOT_IN_MASTER, with the fix as the
    resolution). ``resolution`` = 'pending' / 'resolved' / 'all'.
    ``date_from`` / ``date_to`` (``YYYY-MM-DD``) filter by upload date (run_ts).
    ``limit`` caps rows (pass a large value / 0 for export = no cap)."""
    res_sql = {'pending': _PENDING_SQL, 'resolved': _RESOLVED_SQL}.get(
        resolution, '1=1')
    out = {'ok': False, 'rows': [], 'counts': {}, 'resolution': resolution,
           'date_from': date_from, 'date_to': date_to}
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            # 'Not in master' filter ⇒ currently-NOT_IN_MASTER *or* EAN-corrected
            # (those were not-in-master). MISMATCH is exact. Blank = all affected.
            if status == 'NOT_IN_MASTER':
                status_clause = f"(status='NOT_IN_MASTER' OR {_FIXED_EAN_SQL})"
                params: list = []
            elif status:
                status_clause = f"status={ph}"; params = [status]
            else:
                status_clause = _AFFECTED_SQL; params = []
            where = [status_clause, res_sql]
            if marketplace:
                where.append(f"marketplace={ph}"); params.append(marketplace)
            if q:
                where.append(f"(po LIKE {ph} OR item_no LIKE {ph} "
                             f"OR description LIKE {ph})")
                params += [f"%{q}%", f"%{q}%", f"%{q}%"]
            # Upload-date window (on run_ts). Same fragment reused for the counts.
            date_sql = ''
            date_params: list = []
            if date_from:
                date_sql += f" AND DATE(run_ts) >= {ph}"; date_params.append(date_from)
            if date_to:
                date_sql += f" AND DATE(run_ts) <= {ph}"; date_params.append(date_to)
            wsql = " AND ".join(where) + date_sql
            cols = ['line_id', 'run_ts', 'marketplace', 'po', 'item_no', 'ean',
                    'received_ean', 'exception_label', 'description', 'qty',
                    'vendor_mrp', 'our_mrp', 'vendor_cp', 'our_cp',
                    'vendor_landing', 'our_landing', 'diff', 'status', 'action',
                    'remark']
            lim_sql = f" LIMIT {int(limit)}" if limit and int(limit) > 0 else ''
            cur.execute(
                f"SELECT {', '.join(cols)} FROM order_lines_full WHERE {wsql} "
                f"ORDER BY line_id DESC{lim_sql}", tuple(params) + tuple(date_params))
            out['rows'] = _tag_basis(_rows(cur, cols))
            # headline counts within the chosen resolution scope (+ same date window).
            cur.execute(
                f"SELECT SUM(CASE WHEN status='MISMATCH' THEN 1 ELSE 0 END), "
                f"SUM(CASE WHEN status='NOT_IN_MASTER' OR {_FIXED_EAN_SQL} "
                f"THEN 1 ELSE 0 END) FROM order_lines_full "
                f"WHERE {_AFFECTED_SQL} AND {res_sql}{date_sql}", tuple(date_params))
            mm, nim = cur.fetchone()
            out['counts'] = {'MISMATCH': int(mm or 0), 'NOT_IN_MASTER': int(nim or 0)}
            # Totals across ALL resolutions (same date window) — so the cards can
            # show "X in this view · of N total" and never read as a misleading 0.
            cur.execute(
                f"SELECT SUM(CASE WHEN status='MISMATCH' THEN 1 ELSE 0 END), "
                f"SUM(CASE WHEN status='NOT_IN_MASTER' OR {_FIXED_EAN_SQL} "
                f"THEN 1 ELSE 0 END) FROM order_lines_full "
                f"WHERE {_AFFECTED_SQL}{date_sql}", tuple(date_params))
            tmm, tnim = cur.fetchone()
            out['counts_total'] = {'MISMATCH': int(tmm or 0),
                                   'NOT_IN_MASTER': int(tnim or 0)}
            cur.execute(f"SELECT COUNT(*) FROM order_lines_full "
                        f"WHERE {_AFFECTED_SQL} AND {_RESOLVED_SQL}{date_sql}",
                        tuple(date_params))
            out['resolved_total'] = cur.fetchone()[0]
            # Unresolved (pending) total for the simplified 2-card view.
            cur.execute(f"SELECT COUNT(*) FROM order_lines_full "
                        f"WHERE {_AFFECTED_SQL} AND {_PENDING_SQL}{date_sql}",
                        tuple(date_params))
            out['pending_total'] = cur.fetchone()[0]
        out['ok'] = True
    except Exception as e:  # noqa: BLE001
        out['error'] = f"{type(e).__name__}: {e}"
    return out


def mp_lot_qty(marketplace='', date_from='', date_to='') -> dict:
    """Total UPLOADED lot qty per marketplace over the given upload-date window
    (``run_ts``) — ALL lines, not just affected ones. Used by the Issues email
    to compute the uploaded-% (lot − excluded) ÷ lot. Returns ``{mp: qty}``."""
    out: dict = {}
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            where = ['1=1']
            params: list = []
            if marketplace:
                where.append(f"marketplace={ph}"); params.append(marketplace)
            if date_from:
                where.append(f"DATE(run_ts) >= {ph}"); params.append(date_from)
            if date_to:
                where.append(f"DATE(run_ts) <= {ph}"); params.append(date_to)
            cur.execute(
                f"SELECT marketplace, SUM(qty) FROM order_lines_full "
                f"WHERE {' AND '.join(where)} GROUP BY marketplace", tuple(params))
            for mp, q in cur.fetchall():
                out[str(mp or '—')] = int(q or 0)
    except Exception:  # noqa: BLE001 — best-effort; email falls back to flagged qty
        pass
    return out


def ean_corrections(limit=200) -> dict:
    """Audit of WRONG EANs received & corrected, grouped — the "vendor sent EAN
    X wrong N times" report for escalation. A wrong EAN auto-resolves going
    forward (temporary fix); the permanent fix is the vendor sending the right
    barcode. Derived from the validation layer (received_ean). Never raises."""
    out = {'ok': False, 'rows': [], 'total_lines': 0, 'distinct': 0}
    try:
        with _conn() as (cur, d):
            cur.execute(
                "SELECT v.received_ean, l.ean, l.item_no, MAX(l.description), "
                "COUNT(*), GROUP_CONCAT(DISTINCT l.marketplace), MAX(l.run_ts) "
                "FROM order_line_validation v "
                "JOIN order_lines l ON l.line_id = v.line_id "
                "WHERE v.received_ean IS NOT NULL AND v.received_ean <> '' "
                "GROUP BY v.received_ean, l.ean, l.item_no "
                f"ORDER BY COUNT(*) DESC LIMIT {int(limit)}")
            out['rows'] = _rows(cur, [
                'received_ean', 'correct_ean', 'item_no', 'description',
                'occurrences', 'marketplaces', 'last_seen'])
            out['total_lines'] = sum(r['occurrences'] for r in out['rows'])
            out['distinct'] = len(out['rows'])
            out['ok'] = True
    except Exception as e:  # noqa: BLE001
        out['error'] = f"{type(e).__name__}: {e}"
    return out


_SKU_AGG = """
  SELECT item_no, ean, MAX(description) AS description,
    MAX(our_mrp) AS our_mrp, MAX(vendor_mrp) AS vmrp_max, MIN(vendor_mrp) AS vmrp_min,
    SUM(qty) AS tot_qty,
    SUM(CASE WHEN status='OK' THEN qty ELSE 0 END)            AS ok_qty,
    SUM(CASE WHEN status='MISMATCH' THEN qty ELSE 0 END)      AS mis_qty,
    SUM(CASE WHEN status='NOT_IN_MASTER' THEN qty ELSE 0 END) AS nim_qty,
    SUM(CASE WHEN status='OK' THEN 1 ELSE 0 END)              AS ok_n,
    SUM(CASE WHEN status='MISMATCH' THEN 1 ELSE 0 END)        AS mis_n,
    SUM(CASE WHEN status='NOT_IN_MASTER' THEN 1 ELSE 0 END)   AS nim_n,
    COUNT(DISTINCT po) AS pos, MIN(diff) AS min_diff, MAX(diff) AS max_diff,
    GROUP_CONCAT(DISTINCT marketplace) AS marketplaces,
    COUNT(DISTINCT marketplace) AS mp_count
  FROM order_lines_full {wsql}
  GROUP BY item_no, ean
"""
_SKU_COLS = ['item_no', 'ean', 'description', 'our_mrp', 'vmrp_max', 'vmrp_min',
             'tot_qty', 'ok_qty', 'mis_qty', 'nim_qty', 'ok_n', 'mis_n', 'nim_n',
             'pos', 'min_diff', 'max_diff', 'marketplaces', 'mp_count']


def sku_summary(marketplace='', q='', date_from='', date_to='',
                issues_only=False, limit=5000) -> dict:
    """SKU-wise rollup of every recorded line, grouped by (item_no, EAN): qty +
    line-counts per status (OK / MISMATCH / NOT_IN_MASTER), MRP comparison with a
    'vendor MRP varies' flag, distinct POs, worst diff, marketplaces. Nothing is
    hidden — all SKUs and all three statuses. Never raises."""
    out = {'ok': False, 'rows': [], 'total_skus': 0, 'shown': 0,
           'marketplaces': [], 'totals': {}, 'capped': False}
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute("SELECT DISTINCT marketplace FROM order_lines_full "
                        "ORDER BY marketplace")
            out['marketplaces'] = [r[0] for r in cur.fetchall() if r[0]]
            where, args = [], []
            if marketplace:
                where.append(f"marketplace={ph}"); args.append(marketplace)
            if q:
                where.append(f"(item_no LIKE {ph} OR ean LIKE {ph} OR "
                             f"description LIKE {ph})")
                args += [f"%{q}%", f"%{q}%", f"%{q}%"]
            if date_from:
                where.append(f"run_ts >= {ph}"); args.append(date_from)
            if date_to:
                where.append(f"run_ts <= {ph}"); args.append(f"{date_to} 23:59:59")
            wsql = ('WHERE ' + ' AND '.join(where)) if where else ''
            having = "HAVING (mis_qty > 0 OR nim_qty > 0)" if issues_only else ""
            agg = _SKU_AGG.format(wsql=wsql)
            # total distinct SKUs (post-filter, post-having) for the "of N" note
            cur.execute(f"SELECT COUNT(*) FROM ({agg} {having}) t", args)
            out['total_skus'] = int(cur.fetchone()[0] or 0)
            cur.execute(f"{agg} {having} "
                        f"ORDER BY mis_qty DESC, nim_qty DESC, tot_qty DESC "
                        f"LIMIT {int(limit)}", args)
            rows = _rows(cur, _SKU_COLS)
            for r in rows:
                lo, hi = r.pop('vmrp_min'), r.pop('vmrp_max')
                r['vendor_mrp'] = hi
                r['vmrp_varies'] = (lo is not None and hi is not None and lo != hi)
                ds = [x for x in (r.pop('min_diff'), r.pop('max_diff'))
                      if x is not None]
                r['diff'] = min(ds) if ds else None        # worst (most negative)
            out['rows'] = rows
            out['shown'] = len(rows)
            out['capped'] = out['total_skus'] > len(rows)
            cur.execute(
                f"SELECT SUM(qty), "
                f"SUM(CASE WHEN status='OK' THEN qty ELSE 0 END), "
                f"SUM(CASE WHEN status='MISMATCH' THEN qty ELSE 0 END), "
                f"SUM(CASE WHEN status='NOT_IN_MASTER' THEN qty ELSE 0 END) "
                f"FROM order_lines_full {wsql}", args)
            tq, tok, tmis, tnim = cur.fetchone()
            out['totals'] = {'qty': int(tq or 0), 'ok': int(tok or 0),
                             'mismatch': int(tmis or 0), 'nim': int(tnim or 0)}
            out['ok'] = True
    except Exception as e:  # noqa: BLE001
        out['error'] = f"{type(e).__name__}: {e}"
    return out


def sku_analytics(date_from='', date_to='', marketplace='', top: int = 10,
                  full: bool = False) -> dict:
    """SKU-wise rollup of uploaded POs, filtered by **upload date** (``run_ts``)
    and **marketplace** — overall demanded qty + value, distinct SKUs / POs, and
    the top-N SKUs by **qty** and by **value** (Σ qty × unit_price). The caller
    sets the defaults (the Analytics view defaults to *today's* uploads). Also
    returns the marketplace list for the filter dropdown. With ``full=True`` it
    also returns ``rows`` = every SKU (value-desc) for the full-view page. Never
    raises."""
    out = {'ok': False, 'date_from': date_from, 'date_to': date_to,
           'marketplace': marketplace, 'marketplaces': [],
           'overall': {'skus': 0, 'qty': 0, 'value': 0, 'pos': 0, 'lines': 0},
           'top_qty': [], 'top_value': [], 'rows': []}
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute("SELECT DISTINCT marketplace FROM order_lines_full "
                        "WHERE marketplace IS NOT NULL AND marketplace <> '' "
                        "ORDER BY marketplace")
            out['marketplaces'] = [r[0] for r in cur.fetchall()]
            where, args = [], []
            if date_from:
                where.append(f"DATE(run_ts) >= {ph}"); args.append(date_from)
            if date_to:
                where.append(f"DATE(run_ts) <= {ph}"); args.append(date_to)
            if marketplace:
                where.append(f"marketplace={ph}"); args.append(marketplace)
            wsql = ' AND '.join(where) if where else '1=1'
            cur.execute(
                "SELECT item_no, MAX(description) AS description, SUM(qty) AS qty, "
                "SUM(qty * COALESCE(unit_price, 0)) AS value, "
                "COUNT(DISTINCT po) AS pos, COUNT(*) AS nlines, "
                "COUNT(DISTINCT marketplace) AS mps "
                f"FROM order_lines_full WHERE {wsql} GROUP BY item_no", tuple(args))
            rows = _rows(cur, ['item_no', 'description', 'qty', 'value',
                               'pos', 'lines', 'mps'])
            for r in rows:                       # normalise numerics
                r['qty'] = int(r['qty'] or 0)
                r['value'] = round(float(r['value'] or 0), 2)
            out['overall'] = {
                'skus': len(rows),
                'qty': sum(r['qty'] for r in rows),
                'value': round(sum(r['value'] for r in rows), 2),
                'pos': 0, 'lines': sum(r['lines'] for r in rows),
            }
            cur.execute(f"SELECT COUNT(DISTINCT po) FROM order_lines_full "
                        f"WHERE {wsql}", tuple(args))
            out['overall']['pos'] = int(cur.fetchone()[0] or 0)
            out['top_qty'] = sorted(rows, key=lambda r: -r['qty'])[:top]
            out['top_value'] = sorted(rows, key=lambda r: -r['value'])[:top]
            if full:
                out['rows'] = sorted(rows, key=lambda r: -r['value'])
            out['ok'] = True
    except Exception as e:  # noqa: BLE001
        out['error'] = f"{type(e).__name__}: {e}"
    return out


def sku_lines(item_no='', ean='', limit=300) -> dict:
    """Drill-down: the individual PO-lines behind one (item_no, EAN) SKU."""
    cols = ['po', 'marketplace', 'location', 'qty', 'our_mrp', 'vendor_mrp',
            'our_cp', 'vendor_cp', 'diff', 'status', 'received_ean', 'run_ts']
    out = {'rows': []}
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(
                f"SELECT {', '.join(cols)} FROM order_lines_full "
                f"WHERE item_no={ph} AND ean={ph} ORDER BY run_ts DESC, po "
                f"LIMIT {int(limit)}", (item_no, ean))
            out['rows'] = _rows(cur, cols)
    except Exception:  # noqa: BLE001
        pass
    return out


def ean_correction_counts() -> dict:
    """``{received_ean: occurrences}`` — how many times each wrong EAN was
    received. Used to annotate the review page ('received N times')."""
    try:
        with _conn() as (cur, d):
            cur.execute(
                "SELECT received_ean, COUNT(*) FROM order_line_validation "
                "WHERE received_ean IS NOT NULL AND received_ean <> '' "
                "GROUP BY received_ean")
            return {r[0]: r[1] for r in cur.fetchall()}
    except Exception:  # noqa: BLE001
        return {}


def orders_for_export(sort='date', direction='desc', **filters) -> list[dict]:
    """All filtered order rows (no page limit) for the Excel export."""
    f = _norm_filters(filters)
    with _conn() as (cur, d):
        return _fetch_orders(cur, d, f, sort, direction,
                             limit=1000000, offset=0)


# ── Line items explorer (order_lines, all lines) ────────────────────────

_LINE_VIEW_COLS = [
    'line_id', 'run_id', 'run_ts', 'marketplace', 'po', 'item_no', 'ean',
    'received_ean', 'exception_label', 'description', 'qty', 'unit_price',
    'our_mrp', 'vendor_mrp', 'vendor_cp', 'our_cp', 'vendor_landing',
    'our_landing', 'diff', 'status', 'action', 'remark',
]


def _lines_where(ph, marketplace, status, po, q):
    where, params = ['1=1'], []
    if marketplace:
        where.append(f"marketplace={ph}"); params.append(marketplace)
    if status:
        where.append(f"status={ph}"); params.append(status)
    if po:
        where.append(f"po={ph}"); params.append(po)
    if q:
        where.append(f"(po LIKE {ph} OR item_no LIKE {ph} OR ean LIKE {ph} "
                     f"OR description LIKE {ph})")
        params += [f"%{q}%"] * 4
    return " AND ".join(where), params


def line_items(marketplace='', status='', po='', q='', offset=0,
               limit=PAGE_SIZE) -> dict:
    """Browsable per-line view of ``order_lines`` (the full line audit) with
    filters + pagination. (Only web-confirmed runs have line items.)"""
    out = {
        'ok': False, 'backend': backend_label(), 'rows': [], 'total': 0,
        'offset': offset, 'limit': limit, 'marketplace_options': [],
        'status_options': [], 'kpis': {},
        'filters': {'marketplace': marketplace, 'status': status, 'po': po, 'q': q},
    }
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            wsql, params = _lines_where(ph, marketplace, status, po, q)

            cur.execute(f"SELECT COUNT(*), COALESCE(SUM(qty),0), "
                        f"COUNT(DISTINCT po) FROM order_lines_full WHERE {wsql}",
                        tuple(params))
            n, qty, pos = cur.fetchone()
            out['total'] = n
            out['kpis'] = {'lines': n, 'qty': int(qty or 0), 'pos': pos}

            cur.execute(
                f"SELECT {', '.join(_LINE_VIEW_COLS)} FROM order_lines_full "
                f"WHERE {wsql} ORDER BY line_id DESC "
                f"LIMIT {int(limit)} OFFSET {int(offset)}", tuple(params))
            out['rows'] = _tag_basis(_rows(cur, _LINE_VIEW_COLS))

            cur.execute("SELECT DISTINCT marketplace FROM order_lines_full "
                        "WHERE marketplace <> '' ORDER BY marketplace")
            out['marketplace_options'] = [r[0] for r in cur.fetchall()]
            cur.execute("SELECT DISTINCT status FROM order_lines_full "
                        "WHERE status <> '' ORDER BY status")
            out['status_options'] = [r[0] for r in cur.fetchall()]
        out['ok'] = True
    except Exception as e:  # noqa: BLE001
        out['error'] = f"{type(e).__name__}: {e}"
    return out


def line_items_page(marketplace='', status='', po='', q='', offset=0,
                    limit=PAGE_SIZE) -> dict:
    """Next page of line rows (Load more)."""
    out = {'ok': False, 'rows': [], 'has_more': False, 'next_offset': offset}
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            wsql, params = _lines_where(ph, marketplace, status, po, q)
            cur.execute(f"SELECT COUNT(*) FROM order_lines_full WHERE {wsql}",
                        tuple(params))
            total = cur.fetchone()[0]
            cur.execute(
                f"SELECT {', '.join(_LINE_VIEW_COLS)} FROM order_lines_full "
                f"WHERE {wsql} ORDER BY line_id DESC "
                f"LIMIT {int(limit)} OFFSET {int(offset)}", tuple(params))
            out['rows'] = _tag_basis(_rows(cur, _LINE_VIEW_COLS))
            out['next_offset'] = offset + len(out['rows'])
            out['has_more'] = out['next_offset'] < total
        out['ok'] = True
    except Exception as e:  # noqa: BLE001
        out['error'] = f"{type(e).__name__}: {e}"
    return out


def run_detail(run_id: int) -> dict:
    """Run meta + order headers + full line items (+ affected subset) for one
    run_id. Line items come from ``order_lines`` (only confirmed web runs have
    them; Tkinter-only runs show headers but no lines)."""
    out: dict = {'ok': False, 'run': None, 'orders': [], 'lines': [],
                 'issues': []}
    try:
        with _conn() as (cur, d):
            ot, ph = d['orders'], d['ph']
            cur.execute(
                f"SELECT run_id, run_ts, mode, marketplaces, total_pos, "
                f"total_items, total_qty, total_value FROM runs "
                f"WHERE run_id={ph}", (run_id,))
            r = cur.fetchone()
            if r:
                out['run'] = dict(zip(
                    ['run_id', 'run_ts', 'mode', 'marketplaces', 'total_pos',
                     'total_items', 'total_qty', 'total_value'], r))

            cur.execute(
                f"SELECT marketplace_label, po, location, warehouse, po_date, "
                f"exp_date, order_type, items, qty, order_value, output_file "
                f"FROM {ot} WHERE run_id={ph} ORDER BY po", (run_id,))
            out['orders'] = _rows(cur, [
                'marketplace', 'po', 'location', 'warehouse', 'po_date',
                'exp_date', 'order_type', 'items', 'qty', 'value',
                'output_file'])

            # Full line items for this run (order_lines, keyed by run_id);
            # affected = the status != OK subset (no separate table).
            try:
                cur.execute(
                    f"SELECT po, item_no, ean, description, qty, unit_price, "
                    f"our_mrp, vendor_mrp, vendor_cp, our_cp, vendor_landing, "
                    f"our_landing, diff, status, exception_label, action, remark "
                    f"FROM order_lines_full WHERE run_id={ph} ORDER BY po, item_no",
                    (run_id,))
                out['lines'] = _tag_basis(_rows(cur, [
                    'po', 'item_no', 'ean', 'description', 'qty', 'unit_price',
                    'our_mrp', 'vendor_mrp', 'vendor_cp', 'our_cp',
                    'vendor_landing', 'our_landing', 'diff', 'status',
                    'exception_label', 'action', 'remark']))
            except Exception:
                out['lines'] = []
            out['issues'] = [l for l in out['lines']
                             if l['status'] in ('MISMATCH', 'NOT_IN_MASTER')]
        out['ok'] = True
    except Exception as e:  # noqa: BLE001
        out['error'] = f"{type(e).__name__}: {e}"
    return out
