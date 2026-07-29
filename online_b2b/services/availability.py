"""
online_b2b.services.availability
================================

**Order Availability Checker** — paste order number(s) straight from the Excel
tracker → for each, pull its recorded line items from the DB → check every SKU
against the CURRENT inventory snapshot in the *mapped* warehouse (auto-resolved
from the order's warehouse/marketplace, with a manual override).

Read-only. Reuses the existing building blocks — NO duplication of stock or
order logic:
  * :func:`order_db._conn` + the ``order_lines_full`` view (recorded lines) and
    ``order_headers`` (warehouse + marketplace).
  * :mod:`inventory_store` — ``current_stock_map`` (available qty per item for a
    warehouse), ``resolve_order_wh`` (WH auto-map incl. MP overrides), warehouse
    metadata.
"""

from __future__ import annotations

import re

from . import inventory_store as inv
from .order_db import _conn

# Split pasted text into tokens: Excel copy is tab/space/newline separated; also
# tolerate commas and semicolons. Order numbers themselves keep their internal
# '/' and '-' (e.g. 'SO/RL/07/280728'), so we only break on whitespace + , ;.
_SPLIT = re.compile(r'[\s,;]+')


def parse_order_nos(text) -> list[str]:
    """Pasted blob → de-duplicated, order-preserving list of order numbers."""
    seen: set[str] = set()
    out: list[str] = []
    for tok in _SPLIT.split(str(text or '').strip()):
        t = tok.strip()
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def _q(x):
    """Qty display — whole numbers as int, else 1-dp float."""
    x = float(x or 0)
    return int(x) if x == int(x) else round(x, 1)


def _line_status(found: bool, ordered: float, available: float) -> str:
    if not found:
        return 'NO STOCK'          # item not present in the current snapshot
    if available <= 0:
        return 'OOS'               # present but zero available
    if available < ordered:
        return 'SHORT'             # partial cover
    return 'OK'                    # fully coverable


def check_orders(order_nos, wh_override: str = '') -> dict:
    """For each order number: resolve its warehouse (override wins, else the
    order's own mapped WH) and compare each recorded line's qty to the available
    stock there. Returns a render-ready dict::

        {ok, orders:[{po, marketplace, wh, wh_short, wh_auto, overridden,
                      lines:[{item_no, ean, description, ordered, available,
                              fillable, short, status}],
                      ord_qty, fillable_qty, short_qty, fill_pct, skus}],
         not_found:[po,...], override, wh_options, summary}
    """
    override_code = inv.wh_normalize(wh_override) if (wh_override or '').strip() else ''

    # One stock map per distinct warehouse actually used (cheap + avoids re-query).
    _stock: dict[str, dict] = {}

    def stock_for(wh: str) -> dict:
        if wh not in _stock:
            _stock[wh] = inv.current_stock_map(wh)
        return _stock[wh]

    # Snapshot timestamps → "inventory as of …" per warehouse.
    _snaps = inv.current_snapshots()

    def snap_ts(wh: str) -> str:
        s = _snaps.get(wh)
        return str(s['captured_at']) if s and s.get('captured_at') else ''

    orders: list[dict] = []
    not_found: list[str] = []
    # SKU-wise aggregate across ALL pasted orders, keyed by (warehouse, item) so
    # cumulative demand for one SKU (spanning several POs) is netted against the
    # single stock figure for its warehouse.
    sku_agg: dict = {}

    with _conn() as (cur, d):
        ph, ot = d['ph'], d['orders']
        for po in order_nos:
            # Most recent run for this PO — re-uploads supersede, never double-count.
            cur.execute(
                f"SELECT run_id, warehouse, marketplace_label FROM {ot} "
                f"WHERE po={ph} ORDER BY run_ts DESC LIMIT 1", (po,))
            hdr = cur.fetchone()
            if not hdr:
                not_found.append(po)
                continue
            run_id, wh_raw, mp_label = hdr
            wh_auto = inv.resolve_order_wh(wh_raw, mp_label, mp_label)
            wh = override_code or wh_auto
            sm = stock_for(wh)

            cur.execute(
                f"SELECT item_no, ean, description, qty, unit_price, our_landing "
                f"FROM order_lines_full WHERE po={ph} AND run_id={ph} ORDER BY item_no",
                (po, run_id))
            lrows: list[dict] = []
            ord_qty = fill_qty = short_qty = 0.0
            ord_val = fill_val = short_val = 0.0
            for item_no, ean, desc, qty, unit_price, our_landing in cur.fetchall():
                key = str(item_no or '').strip()
                q = float(qty or 0)
                # per-unit value: inc-GST landing preferred, else unit price (CP).
                uv = float(our_landing or 0) or float(unit_price or 0)
                found = key in sm
                avail = float(sm.get(key, 0) or 0)
                avail_eff = avail if avail > 0 else 0.0   # oversold (<0) → 0 fillable
                fillable = min(q, avail_eff)
                short = q - fillable                      # ≤ ordered, always
                lo_v, lf_v, ls_v = q * uv, fillable * uv, short * uv
                lrows.append({
                    'item_no': key, 'ean': str(ean or ''),
                    'description': str(desc or ''),
                    'ordered': _q(q),
                    'available': _q(avail), 'fillable': _q(fillable), 'short': _q(short),
                    'unit_value': round(uv, 2),
                    'ordered_value': round(lo_v, 2), 'fillable_value': round(lf_v, 2),
                    'short_value': round(ls_v, 2),
                    'status': _line_status(found, q, avail),
                })
                ord_qty += q; fill_qty += fillable; short_qty += short
                ord_val += lo_v; fill_val += lf_v; short_val += ls_v
                # accumulate SKU-wise (demand summed; availability captured once)
                a = sku_agg.get((wh, key))
                if a is None:
                    a = sku_agg[(wh, key)] = {
                        'item_no': key, 'ean': str(ean or ''),
                        'description': str(desc or ''), 'wh': wh,
                        'wh_short': inv.wh_short(wh), 'ordered': 0.0,
                        'ordered_value': 0.0, 'available': avail, 'found': found,
                        'pos': set()}
                a['ordered'] += q
                a['ordered_value'] += lo_v
                a['pos'].add(po)
                if not a['ean'] and ean:
                    a['ean'] = str(ean)
                if not a['description'] and desc:
                    a['description'] = str(desc)
            orders.append({
                'po': po, 'marketplace': str(mp_label or ''),
                'wh': wh, 'wh_short': inv.wh_short(wh), 'stock_as_of': snap_ts(wh),
                'wh_auto': wh_auto, 'wh_auto_short': inv.wh_short(wh_auto),
                'overridden': bool(override_code) and override_code != wh_auto,
                'lines': lrows, 'skus': len(lrows),
                'ord_qty': _q(ord_qty), 'fillable_qty': _q(fill_qty), 'short_qty': _q(short_qty),
                'fill_pct': round(fill_qty / ord_qty * 100, 1) if ord_qty else 0.0,
                'ord_value': round(ord_val, 2), 'fillable_value': round(fill_val, 2),
                'short_value': round(short_val, 2),
                'fill_val_pct': round(fill_val / ord_val * 100, 1) if ord_val else 0.0,
                'has_value': ord_val > 0,
                'fully': short_qty <= 0 and len(lrows) > 0,
            })

    # SKU-wise fill rate — demand netted against stock (worst-case truth when the
    # same SKU is pulled by multiple pasted POs from one warehouse).
    skus: list[dict] = []
    for a in sku_agg.values():
        o = a['ordered']
        av = a['available']
        ov = a['ordered_value']
        avail_eff = av if av > 0 else 0.0                 # oversold (<0) → 0 fillable
        fillable = min(o, avail_eff)
        short = o - fillable                              # ≤ ordered, always
        uv = (ov / o) if o else 0.0                       # avg per-unit value
        fv, sv = fillable * uv, short * uv
        skus.append({
            'item_no': a['item_no'], 'ean': a['ean'],
            'description': a['description'], 'wh': a['wh'], 'wh_short': a['wh_short'],
            'pos': len(a['pos']), 'ordered': _q(o), 'available': _q(av),
            'fillable': _q(fillable), 'short': _q(short),
            'ordered_value': round(ov, 2), 'fillable_value': round(fv, 2),
            'short_value': round(sv, 2),
            'fill_pct': round(fillable / o * 100, 1) if o else 0.0,
            'fill_val_pct': round(fv / ov * 100, 1) if ov else 0.0,
            'status': _line_status(a['found'], o, av),
        })
    skus.sort(key=lambda x: (-x['short'], -x['ordered']))

    tot_ord = sum(o['ord_qty'] for o in orders)
    tot_fill = sum(o['fillable_qty'] for o in orders)
    tot_ordv = sum(o['ord_value'] for o in orders)
    tot_fillv = sum(o['fillable_value'] for o in orders)
    # "inventory as of" — per warehouse actually used (short tag → captured_at).
    used_wh = {(o['wh_short'], snap_ts(o['wh'])) for o in orders if snap_ts(o['wh'])}
    wh_stock_as_of = {short: ts for short, ts in used_wh}
    summary = {
        'orders': len(orders), 'not_found': len(not_found),
        'skus': len(skus),
        'ord_qty': _q(tot_ord), 'fillable_qty': _q(tot_fill),
        'short_qty': _q(sum(o['short_qty'] for o in orders)),
        'fill_pct': round(tot_fill / tot_ord * 100, 1) if tot_ord else 0.0,
        'ord_value': round(tot_ordv, 2), 'fillable_value': round(tot_fillv, 2),
        'short_value': round(sum(o['short_value'] for o in orders), 2),
        'fill_val_pct': round(tot_fillv / tot_ordv * 100, 1) if tot_ordv else 0.0,
        'has_value': tot_ordv > 0,
        'fully': sum(1 for o in orders if o['fully']),
        'wh_stock_as_of': wh_stock_as_of,
        'stock_as_of': next(iter(wh_stock_as_of.values()), '') if len(wh_stock_as_of) == 1 else '',
    }
    # Bin classification for the warehouse(s) touched — which bins are INCLUDED
    # vs EXCLUDED (so the WH team can see WHY an item reads short: its stock may
    # sit in an excluded return/QC bin, an unclassified 'new' bin, or a negative
    # pick face). Aggregated per bin (not per item — that detail isn't stored).
    bins: dict = {}
    _DEC = {'include': 'INCLUDED', 'exclude': 'EXCLUDED', 'new': 'NEW (unclassified)'}
    for wh in {o['wh'] for o in orders}:
        snap = _snaps.get(wh)
        if not snap:
            continue
        try:
            rows = inv.bin_audit(snap['snapshot_id'])
        except Exception:  # noqa: BLE001
            rows = []
        bins[inv.wh_short(wh)] = [{
            'bin': r.get('bin_code', ''), 'zone': r.get('zone_code', ''),
            'decision': _DEC.get(r.get('decision', ''), str(r.get('decision', '')).upper()),
            'lines': r.get('lines', 0), 'qty': _q(r.get('qty', 0)),
        } for r in rows]

    return {'ok': True, 'orders': orders, 'skus': skus, 'not_found': not_found,
            'bins': bins, 'override': override_code, 'wh_options': inv.WAREHOUSES,
            'summary': summary}


# ── Styled Excel export (same look as our other workbook downloads) ──────────

def to_workbook(data: dict):
    """Render an availability result (from :func:`check_orders`) into a styled
    multi-sheet .xlsx — Summary · By Order (PO-SKU line items) · By SKU · Not
    Found — matching our standard workbook styling (navy header, frozen header,
    auto-filter). Returns a ``BytesIO`` positioned at 0."""
    import datetime as _dt
    import io

    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    NAVY = PatternFill('solid', fgColor='1A237E')
    HEADF = Font(bold=True, color='FFFFFF')
    CENTER = Alignment(horizontal='center', vertical='center')
    OK = PatternFill('solid', fgColor='DCFCE7')
    SHORT = PatternFill('solid', fgColor='FEF3C7')
    OOS = PatternFill('solid', fgColor='FEE2E2')
    NOST = PatternFill('solid', fgColor='EAECEF')
    STFILL = {'OK': OK, 'SHORT': SHORT, 'OOS': OOS, 'NO STOCK': NOST}
    CUR = '[$₹-4009]#,##,##0.00'         # Indian-grouped rupee currency format

    def _cur_cols(ws, cols):
        """Apply the rupee format to `cols` (1-indexed) for every data row."""
        for row in range(2, ws.max_row + 1):
            for col in cols:
                ws.cell(row=row, column=col).number_format = CUR

    def _sheet(ws, heads, widths):
        ws.append(heads)
        for c in ws[1]:
            c.font = HEADF; c.fill = NAVY; c.alignment = CENTER
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = 'A2'

    def _finish(ws, ncols):
        if ws.max_row > 1:
            ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{ws.max_row}"

    s = data.get('summary', {})
    wb = openpyxl.Workbook()

    # 1) Summary
    ws = wb.active; ws.title = 'Summary'
    ws['A1'] = 'AVAILABILITY CHECK'; ws['A1'].font = Font(bold=True, size=14, color='1A237E')
    asof = s.get('wh_stock_as_of') or {}
    pairs = [
        ('Generated', f"{_dt.datetime.now():%d-%b-%Y %H:%M}"),
        ('Orders checked', s.get('orders', 0)),
        ('Not found', s.get('not_found', 0)),
        ('Distinct SKUs', s.get('skus', 0)),
        ('Fully coverable orders', s.get('fully', 0)),
        ('— Quantity —', ''),
        ('Ordered qty', s.get('ord_qty', 0)),
        ('Fillable qty', s.get('fillable_qty', 0)),
        ('Short qty', s.get('short_qty', 0)),
        ('Fill rate % (qty)', s.get('fill_pct', 0)),
        ('— Value (₹) —', ''),
        ('Ordered value', s.get('ord_value', 0)),
        ('Fillable value', s.get('fillable_value', 0)),
        ('Short value', s.get('short_value', 0)),
        ('Fill rate % (value)', s.get('fill_val_pct', 0)),
        ('Inventory as of', ' | '.join(f"{k}: {v}" for k, v in asof.items()) or '—'),
    ]
    _cur_labels = {'Ordered value', 'Fillable value', 'Short value'}
    for i, (k, v) in enumerate(pairs, start=3):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        cell = ws.cell(row=i, column=2, value=v)
        if k in _cur_labels:
            cell.number_format = CUR
    ws.column_dimensions['A'].width = 24; ws.column_dimensions['B'].width = 40

    # 2) PO Summary — one row per order, fill rate qty AND value
    ws = wb.create_sheet('PO Summary')
    _sheet(ws, ['Order No', 'Marketplace', 'Warehouse', 'SKUs', 'Ordered Qty',
                'Fillable Qty', 'Short Qty', 'Fill % (Qty)', 'Ordered ₹',
                'Fillable ₹', 'Short ₹', 'Fill % (Val)', 'Fully'],
           [20, 18, 12, 7, 11, 11, 10, 11, 14, 14, 13, 11, 8])
    for o in data.get('orders', []):
        ws.append([o['po'], o['marketplace'], o['wh_short'], o['skus'],
                   o['ord_qty'], o['fillable_qty'], o['short_qty'], o['fill_pct'],
                   o['ord_value'], o['fillable_value'], o['short_value'],
                   o['fill_val_pct'], 'YES' if o['fully'] else 'NO'])
    _finish(ws, 13)
    _cur_cols(ws, [9, 10, 11])          # Ordered ₹ · Fillable ₹ · Short ₹

    # 3) By Order — PO-SKU line items (qty + value)
    ws = wb.create_sheet('By Order Lines')
    _sheet(ws, ['Order No', 'Marketplace', 'Warehouse', 'Item No', 'EAN',
                'Description', 'Ordered', 'Available', 'Fillable', 'Short',
                'Unit ₹', 'Ordered ₹', 'Fillable ₹', 'Short ₹', 'Status'],
           [20, 18, 12, 12, 16, 40, 9, 10, 9, 8, 10, 13, 13, 12, 11])
    for o in data.get('orders', []):
        for l in o['lines']:
            ws.append([o['po'], o['marketplace'], o['wh_short'], l['item_no'],
                       l['ean'], l['description'], l['ordered'], l['available'],
                       l['fillable'], l['short'], l['unit_value'],
                       l['ordered_value'], l['fillable_value'], l['short_value'],
                       l['status']])
            fill = STFILL.get(l['status'])
            if fill:
                ws.cell(row=ws.max_row, column=15).fill = fill
    _finish(ws, 15)
    _cur_cols(ws, [11, 12, 13, 14])     # Unit ₹ · Ordered ₹ · Fillable ₹ · Short ₹

    # 4) By SKU — aggregated across pasted orders (qty + value)
    ws = wb.create_sheet('By SKU')
    _sheet(ws, ['Item No', 'EAN', 'Description', 'Warehouse', 'POs', 'Ordered',
                'Available', 'Fillable', 'Short', 'Fill % (Qty)', 'Ordered ₹',
                'Fillable ₹', 'Short ₹', 'Fill % (Val)', 'Status'],
           [12, 16, 40, 12, 6, 10, 11, 10, 9, 11, 13, 13, 12, 11, 11])
    for k in data.get('skus', []):
        ws.append([k['item_no'], k['ean'], k['description'], k['wh_short'],
                   k['pos'], k['ordered'], k['available'], k['fillable'],
                   k['short'], k['fill_pct'], k['ordered_value'],
                   k['fillable_value'], k['short_value'], k['fill_val_pct'],
                   k['status']])
        fill = STFILL.get(k['status'])
        if fill:
            ws.cell(row=ws.max_row, column=15).fill = fill
    _finish(ws, 15)
    _cur_cols(ws, [11, 12, 13])         # Ordered ₹ · Fillable ₹ · Short ₹

    # 5) Bin Classification — which bins we INCLUDE vs EXCLUDE per warehouse, so
    #    the WH team can see why an item reads short (stock in an excluded
    #    return/QC bin, an unclassified 'new' bin, or a negative pick face).
    bins = data.get('bins') or {}
    if bins:
        ws = wb.create_sheet('Bin Classification')
        _sheet(ws, ['Warehouse', 'Bin', 'Zone', 'Decision', 'Lines', 'Qty'],
               [12, 32, 18, 20, 8, 12])
        INC = PatternFill('solid', fgColor='DCFCE7')
        EXC = PatternFill('solid', fgColor='FEE2E2')
        NEWF = PatternFill('solid', fgColor='FEF3C7')
        for wh_short, rows in bins.items():
            for b in rows:
                ws.append([wh_short, b['bin'], b['zone'], b['decision'],
                           b['lines'], b['qty']])
                dec = b['decision']
                f = (INC if dec.startswith('INCLUDED')
                     else EXC if dec.startswith('EXCLUDED') else NEWF)
                ws.cell(row=ws.max_row, column=4).fill = f
        _finish(ws, 6)

    # 6) SKU Bins — per-item bin breakdown (where each SKU's stock sits: INCLUDED
    #    pick faces vs EXCLUDED return/QC bins) — explains a short SKU bin-by-bin.
    sku_bins = data.get('sku_bins') or {}
    if sku_bins:
        ws = wb.create_sheet('SKU Bins')
        _sheet(ws, ['Item No', 'Description', 'Warehouse', 'Bin', 'Zone',
                    'Decision', 'Qty'], [12, 40, 12, 30, 16, 14, 12])
        INC = PatternFill('solid', fgColor='DCFCE7')
        EXC = PatternFill('solid', fgColor='FEE2E2')
        NEWF = PatternFill('solid', fgColor='FEF3C7')
        for k in data.get('skus', []):
            blist = (sku_bins.get(k['wh']) or {}).get(k['item_no'], [])
            for b in blist:
                ws.append([k['item_no'], k['description'], k['wh_short'],
                           b['bin'], b['zone'], b['decision'], b['qty']])
                dec = b['decision']
                f = (INC if dec.startswith('INCLUDED')
                     else EXC if dec.startswith('EXCLUDED') else NEWF)
                ws.cell(row=ws.max_row, column=6).fill = f
        _finish(ws, 7)

    # 7) Not Found (only if any)
    nf = data.get('not_found', [])
    if nf:
        ws = wb.create_sheet('Not Found')
        _sheet(ws, ['Order No (not in system)'], [30])
        for po in nf:
            ws.append([po])
        _finish(ws, 1)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
