"""
online_b2b.services.inventory_fill
==================================

**Fill-rate analytics** — joins the recorded demand (``order_lines_full``) against
the CURRENT inventory snapshot (:mod:`inventory_store`) to answer, for a chosen
scope: how much of the ordered qty can we actually ship, what's out of stock, and
what will it tentatively bill — **PO-wise and MP-wise**, plus a clean-vs-affected
PO flag.

Value basis is the SAME inc-GST line value the Triangular reconcile and the
Summary Email use (:func:`triangular_validation._line_val`), so "tentative billing"
here reconciles with "uploaded value" there — tentative billing = uploaded value
× fill ratio.

Allocation model (v1, deterministic, no priority ordering): per item, demand
``D`` = Σ ordered qty of the NON-dropped lines in scope; available ``A`` = current
sellable stock; fill ratio ``f = min(A, D) / D``. Each line of that item is filled
**proportionally** (``line_fillable = line_qty × f``) — a fair-share split that
needs no tie-break. A PO/line is **clean** when every item it needs is fully in
stock (``A ≥ D`` → ``f = 1``); **affected** if any item is short.

Read-only; never writes; never raises (returns ``ok=False``).
"""
from __future__ import annotations

import datetime as _dt

from .order_db import _conn
from . import inventory_store as store


def _pct(part, whole) -> float:
    return round(100.0 * part / whole, 1) if whole else 0.0


def fill_rate(date_from='', date_to='', marketplace='', warehouse='',
              segment='') -> dict:
    """Fill-rate rollup for orders in [date_from, date_to] (default: today) against
    the current stock of ``warehouse`` ('' = all warehouses combined).

    Demand is attributed to a warehouse via each order's stored WH, normalized to
    the inventory Location code (+ marketplace overrides, e.g. BlinkMP → DS_BL_OFF1)
    — so when ``warehouse`` is set, only the lines that ship from THAT warehouse
    are counted. ``segment`` = 'online' / 'offline' / '' (both)."""
    from .triangular_validation import _line_val, _is_dropped

    today = _dt.date.today().isoformat()
    date_from = date_from or today
    date_to = date_to or date_from
    seg = str(segment or '').strip().lower()
    out = {'ok': False, 'error': '', 'warehouse': warehouse,
           'warehouse_label': (store.wh_name(warehouse) if warehouse
                               else 'All warehouses'),
           'date_from': date_from, 'date_to': date_to, 'marketplace': marketplace,
           'segment': seg,
           'items': [], 'pos': [], 'mps': [], 'totals': {},
           'stock_as_of': None, 'has_stock': False}

    # current stock (item → available) for the chosen WH (or combined)
    stock = store.current_stock_map(warehouse)
    snaps = store.current_snapshots()
    out['has_stock'] = bool(stock)
    caps = [s['captured_at'] for c, s in snaps.items()
            if (not warehouse or c == warehouse)]
    out['stock_as_of'] = max((str(c) for c in caps if c), default=None)

    try:
        with _conn() as (cur, d):
            ph = d['ph']
            # header map: (marketplace, po) → (warehouse, segment) so each demand
            # line can be attributed to its fulfilment warehouse + segment.
            hwhere = [f"DATE(run_ts) >= {ph}", f"DATE(run_ts) <= {ph}"]
            hparams: list = [date_from, date_to]
            if marketplace:
                hwhere.append(f"marketplace={ph}"); hparams.append(marketplace)
            cur.execute(
                f"SELECT marketplace, marketplace_label, po, warehouse, segment "
                f"FROM order_headers WHERE {' AND '.join(hwhere)}", tuple(hparams))
            # Key by PO (unique) AND by (marketplace, po). MT child channels store
            # the header marketplace as 'MT' but their LINES carry the label
            # ('Health & Glow'), so an (mk, po) join misses and the line would fall
            # back to the default warehouse — losing the operator's WH choice. The
            # PO-only key rescues that; the exact key stays for any legacy caller.
            hmap = {}
            _ov = store.wh_override_map()   # per-PO manual WH shifts win
            for mk, lbl, po, wh, sg in cur.fetchall():
                rec = {'wh': store.effective_order_wh(po, wh, mk, lbl, _ov),
                       'seg': str(sg or '')}
                hmap[str(po or '')] = rec
                hmap[(str(mk or ''), str(po or ''))] = rec

            where = [f"DATE(run_ts) >= {ph}", f"DATE(run_ts) <= {ph}"]
            params: list = [date_from, date_to]
            if marketplace:
                where.append(f"marketplace={ph}"); params.append(marketplace)
            cur.execute(
                f"SELECT po, marketplace, item_no, ean, description, qty, "
                f"our_landing, unit_price, gst_code, status, action "
                f"FROM order_lines_full WHERE {' AND '.join(where)}", tuple(params))
            cols = ['po', 'marketplace', 'item_no', 'ean', 'description', 'qty',
                    'our_landing', 'unit_price', 'gst_code', 'status', 'action']
            rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    except Exception as e:  # noqa: BLE001
        out['error'] = f'{type(e).__name__}: {e}'
        return out

    def _seg_ok(sg: str) -> bool:
        if not seg:
            return True
        online = str(sg or '') == 'OnlineB2B'
        return online if seg == 'online' else (not online)

    # keep only lines that actually get fulfilled (non-dropped) + carry value,
    # attributed to their fulfilment warehouse (+ segment) — filter to scope.
    lines = []
    demand: dict = {}                       # item_no → total ordered qty
    for r in rows:
        if _is_dropped(r['status'], r['action']):
            continue
        q = int(r['qty'] or 0)
        if q <= 0:
            continue
        h = (hmap.get((str(r['marketplace'] or ''), str(r['po'] or '')))
             or hmap.get(str(r['po'] or '')) or {})
        line_wh = h.get('wh') or store.DEFAULT_WH
        if warehouse and line_wh != warehouse:
            continue
        if not _seg_ok(h.get('seg', '')):
            continue
        item = str(r['item_no'] or '').strip()
        val = _line_val(r['our_landing'], r['unit_price'], r['gst_code'], q)
        r['_item'] = item
        r['_qty'] = q
        r['_val'] = val
        r['_wh'] = line_wh
        lines.append(r)
        demand[item] = demand.get(item, 0) + q

    # per-item fill ratio. Clamp available at 0 — D365 can carry a NEGATIVE
    # on-hand (over-picked / correction rows); a negative must read as 0% fill,
    # never a negative ratio.
    fill: dict = {}
    for item, dqty in demand.items():
        avail = max(0.0, float(stock.get(item, 0) or 0))
        fill[item] = min(avail, dqty) / dqty if dqty else 0.0

    def _fillable(r):
        return r['_qty'] * fill.get(r['_item'], 0.0)

    # ── per-PO and per-MP rollups ──
    def _blank(key, label):
        return {'key': key, 'label': label, 'ordered_qty': 0, 'fillable_qty': 0.0,
                'oos_qty': 0.0, 'value': 0.0, 'billing': 0.0,
                'lines': 0, 'short_lines': 0, 'affected': False}

    pos: dict = {}
    mps: dict = {}
    for r in lines:
        f = fill.get(r['_item'], 0.0)
        fq = r['_qty'] * f
        short = f < 0.999
        for bucket, key, label in ((pos, str(r['po'] or '—'), str(r['po'] or '—')),
                                   (mps, str(r['marketplace'] or '—'),
                                    str(r['marketplace'] or '—'))):
            g = bucket.get(key)
            if g is None:
                g = bucket[key] = _blank(key, label)
            g['ordered_qty'] += r['_qty']
            g['fillable_qty'] += fq
            g['oos_qty'] += (r['_qty'] - fq)
            g['value'] += r['_val']
            g['billing'] += r['_val'] * f
            g['lines'] += 1
            if short:
                g['short_lines'] += 1
                g['affected'] = True

    def _finish(bucket):
        rows_out = []
        for g in bucket.values():
            g['fillable_qty'] = round(g['fillable_qty'], 1)
            g['oos_qty'] = round(g['oos_qty'], 1)
            g['value'] = round(g['value'], 2)
            g['billing'] = round(g['billing'], 2)
            g['fill_pct'] = _pct(g['fillable_qty'], g['ordered_qty'])
            g['bill_pct'] = _pct(g['billing'], g['value'])
            g['clean'] = not g['affected']
            rows_out.append(g)
        rows_out.sort(key=lambda x: (x['affected'], -x['oos_qty']), reverse=False)
        # affected first (most OOS on top), then clean
        rows_out.sort(key=lambda x: (0 if x['affected'] else 1, -x['oos_qty']))
        return rows_out

    out['pos'] = _finish(pos)
    out['mps'] = _finish(mps)

    # ── per-item OOS list (only items with a shortfall, worst first) ──
    items_out = []
    for item, dqty in demand.items():
        avail = max(0.0, float(stock.get(item, 0) or 0))
        oos = max(dqty - avail, 0)
        meta = next((r for r in lines if r['_item'] == item), {})
        items_out.append({
            'item_no': item, 'ean': str(meta.get('ean') or ''),
            'description': str(meta.get('description') or ''),
            'demand': dqty, 'available': round(avail, 1), 'oos': round(oos, 1),
            'fill_pct': _pct(min(avail, dqty), dqty)})
    items_out.sort(key=lambda x: (-x['oos'], -x['demand']))
    out['items'] = items_out

    # ── grand totals ──
    ordered = sum(r['_qty'] for r in lines)
    fillable = sum(_fillable(r) for r in lines)
    value = sum(r['_val'] for r in lines)
    billing = sum(r['_val'] * fill.get(r['_item'], 0.0) for r in lines)
    aff_pos = sum(1 for g in out['pos'] if g['affected'])
    out['totals'] = {
        'ordered_qty': ordered, 'fillable_qty': round(fillable, 1),
        'oos_qty': round(ordered - fillable, 1), 'fill_pct': _pct(fillable, ordered),
        'value': round(value, 2), 'billing': round(billing, 2),
        'bill_pct': _pct(billing, value),
        'po_count': len(out['pos']), 'affected_pos': aff_pos,
        'clean_pos': len(out['pos']) - aff_pos,
        'mp_count': len(out['mps']),
        'oos_items': sum(1 for i in items_out if i['oos'] > 0),
        'demand_items': len(demand),
    }
    out['ok'] = True
    return out


def item_fill_ratios(date_from='', date_to='', marketplace='', warehouse='',
                     segment='') -> dict:
    """Per-item fill ratio ``{item_no: ratio}`` for the SAME scope + rules as
    :func:`fill_rate` (shared-stock fair-share: demand is summed across ALL POs in
    scope, so one SKU has ONE ratio no matter how many POs ordered it). Used by the
    SKU drill-down so a per-SKU fill can be shown that reconciles EXACTLY with the
    PO/MP/segment fill already on the board. Read-only; ``{}`` on any error.

    NOTE: pass the SAME ``segment`` (and leave ``warehouse``/``marketplace`` blank)
    the board used, or the ratios will not reconcile — the board builds with
    ``fill_rate(date_from=day, date_to=day, segment=seg)``."""
    from .triangular_validation import _line_val, _is_dropped  # noqa: F401
    today = _dt.date.today().isoformat()
    date_from = date_from or today
    date_to = date_to or date_from
    seg = str(segment or '').strip().lower()
    stock = store.current_stock_map(warehouse)
    demand: dict = {}
    try:
        with _conn() as (cur, d):
            ph = d['ph']
            hwhere = [f"DATE(run_ts) >= {ph}", f"DATE(run_ts) <= {ph}"]
            hparams: list = [date_from, date_to]
            if marketplace:
                hwhere.append(f"marketplace={ph}"); hparams.append(marketplace)
            cur.execute(
                f"SELECT marketplace, marketplace_label, po, warehouse, segment "
                f"FROM order_headers WHERE {' AND '.join(hwhere)}", tuple(hparams))
            hmap = {}
            _ov = store.wh_override_map()   # per-PO manual WH shifts win
            for mk, lbl, po, wh, sg in cur.fetchall():
                rec = {'wh': store.effective_order_wh(po, wh, mk, lbl, _ov),
                       'seg': str(sg or '')}
                hmap[str(po or '')] = rec           # PO-unique key (MT child safe)
                hmap[(str(mk or ''), str(po or ''))] = rec
            where = [f"DATE(run_ts) >= {ph}", f"DATE(run_ts) <= {ph}"]
            params: list = [date_from, date_to]
            if marketplace:
                where.append(f"marketplace={ph}"); params.append(marketplace)
            cur.execute(
                f"SELECT po, marketplace, item_no, qty, status, action "
                f"FROM order_lines_full WHERE {' AND '.join(where)}", tuple(params))
            rows = cur.fetchall()
    except Exception:  # noqa: BLE001
        return {}

    def _seg_ok(sg: str) -> bool:
        if not seg:
            return True
        online = str(sg or '') == 'OnlineB2B'
        return online if seg == 'online' else (not online)

    for po, mk, item, qty, status, action in rows:
        if _is_dropped(status, action):
            continue
        q = int(qty or 0)
        if q <= 0:
            continue
        h = (hmap.get((str(mk or ''), str(po or '')))
             or hmap.get(str(po or '')) or {})
        line_wh = h.get('wh') or store.DEFAULT_WH
        if warehouse and line_wh != warehouse:
            continue
        if not _seg_ok(h.get('seg', '')):
            continue
        it = str(item or '').strip()
        demand[it] = demand.get(it, 0) + q

    ratios: dict = {}
    for it, dqty in demand.items():
        avail = max(0.0, float(stock.get(it, 0) or 0))
        ratios[it] = (min(avail, dqty) / dqty) if dqty else 0.0
    return ratios
