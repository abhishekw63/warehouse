"""
online_b2b.services.summary_email
=================================

The **Consolidated Daily Summary** email — built on the same reusable
:mod:`mailer` skeleton as :mod:`issue_email` / :mod:`daily_email`, so it reuses
the one SMTP/send path and the preview/send contract (no new mail plumbing).

It is a *summary only* (the detailed issue LINES stay on the Issues tab):

1. **Received board — every Online-B2B marketplace**, each marked *received
   today* / *NOT received today*; a not-received channel also shows when it
   **last received** (audit trail).
2. **Per-marketplace details** — PO count · items · qty · value · **included**
   (items that pass) · **excluded** (today's issue lines) for the received ones;
   dashes for the rest. Plus a headline *N issue lines — see the Issues tab*.

All figures are REUSED from the data layer
(:func:`order_db.marketplace_daily_intake`, itself read-only) and the marketplace
registry — no business logic is recomputed here.
"""
from __future__ import annotations

import datetime as _dt
from html import escape

from . import marketplaces as reg
from . import order_db
from .issue_email import _clean_emails      # reuse the shared recipient cleaner
from .mailer import EmailReport

try:                                        # reuse the ONE Indian-grouping helper
    from ..templatetags.b2b_extras import indnum as _indnum
except Exception:                           # noqa: BLE001 — defensive fallback
    def _indnum(value, dp=0):
        try:
            return f'{float(value or 0):,.{int(dp)}f}'
        except (TypeError, ValueError):
            return str(value)


_SKU_EMAIL_CAP = 60          # SKU rows shown in the email (rest flagged, not dropped)

import os as _os
# Cockpit summary is ~1.9 s to assemble; memoise briefly so revisits are instant.
# Short window → today's data stays fresh; an upload confirm busts it immediately.
_SUMMARY_TTL = float(_os.environ.get('ORDERDB_SUMMARY_TTL', '25'))


def _inr(v) -> str:
    """₹ with Indian digit grouping, no paise — ₹97,76,736."""
    return '₹' + _indnum(v, 0)


def _grp(v) -> str:
    """Plain Indian-grouped integer (for qty/counts) — 34,523."""
    return _indnum(v, 0)


def _pctg(part, whole) -> float:
    """part / whole as a 1-dp percentage (0 if whole is 0)."""
    try:
        return round(100.0 * float(part) / float(whole), 1) if whole else 0.0
    except (TypeError, ValueError, ZeroDivisionError):
        return 0.0


def _finish_leg(d: dict) -> dict:
    """Round the value legs + add uploaded/excluded % (of raw) for qty & value.
    Shared by the PO and SKU drill-down levels so every level renders identically."""
    d = dict(d)
    d['raw_value'] = round(d['raw_value'], 2)
    d['uploaded_value'] = round(d['uploaded_value'], 2)
    d['excluded_value'] = round(d['excluded_value'], 2)
    d['uploaded_qty_pct'] = _pctg(d['uploaded_qty'], d['raw_qty'])
    d['excluded_qty_pct'] = _pctg(d['excluded_qty'], d['raw_qty'])
    d['uploaded_value_pct'] = _pctg(d['uploaded_value'], d['raw_value'])
    d['excluded_value_pct'] = _pctg(d['excluded_value'], d['raw_value'])
    d['billing_value'] = None
    return d


def _finish_pd(pd: dict, skip_skus: bool = False) -> dict:
    """A finished PO row. SKU rows are LAZY — the page passes ``skip_skus=True`` so
    the (potentially thousands of) SKU rows are NOT built/rendered upfront; they're
    fetched per-PO on click via :func:`po_skus`. ``sku_count`` is always kept so the
    PO row can show "N SKU" and know it's expandable."""
    fp = _finish_leg(pd)
    fp['sku_count'] = len(pd.get('skus', {}))
    fp.pop('skus', None)
    if skip_skus:
        fp['sku_detail'] = []
    else:
        fp['sku_detail'] = [_finish_leg(sd) for _it, sd in
                            sorted(pd.get('skus', {}).items(), key=lambda kv: -kv[1]['raw_qty'])]
    return fp


def po_skus(day, marketplace, po, segment='') -> list:
    """Finished SKU rows for ONE PO — the lazy 2nd drill level (fetched on click).
    One focused query for this PO's lines + ONE scoped per-item fill-ratio map
    (shared-stock, so it reconciles EXACTLY with the PO/MP/segment fill on the
    board). ``segment`` = 'online'/'offline' — MUST match the board's scope or the
    per-SKU fill won't tie out. Returns [] on any error."""
    try:
        pd = order_db.po_sku_detail(day=day, marketplace=marketplace, po=po)
    except Exception:  # noqa: BLE001
        return []
    skus = pd.get('skus', {}) if pd else {}
    # per-item fill ratio for the board's scope — a SKU that's OOS reads its true
    # fill here (e.g. 0%), so the PO's 68% is explainable line-by-line.
    ratios: dict = {}
    try:
        from . import inventory_fill
        ratios = inventory_fill.item_fill_ratios(date_from=day, date_to=day,
                                                 segment=segment)
    except Exception:  # noqa: BLE001
        ratios = {}
    out = []
    for _it, sd in sorted(skus.items(), key=lambda kv: -kv[1]['raw_qty']):
        leg = _finish_leg(sd)
        f = ratios.get(str(_it).strip())
        if f is not None:
            leg['billing_qty'] = round((sd.get('uploaded_qty') or 0) * f, 1)
            leg['billing_value'] = round((sd.get('uploaded_value') or 0) * f, 2)
            # fill % is the ratio itself (fillable ÷ uploaded) — same for qty & value.
            leg['fill_qty_pct'] = _pctg(leg['billing_qty'], sd.get('uploaded_qty') or 0)
            leg['fill_val_pct'] = _pctg(leg['billing_value'], sd.get('uploaded_value') or 0)
        out.append(leg)
    return out


def _num(v):
    """Decimal/None → plain int/float for arithmetic + display."""
    if v is None:
        return 0
    try:
        f = float(v)
        return int(f) if f == int(f) else round(f, 2)
    except (TypeError, ValueError):
        return 0


def _fmt_ts(ts) -> str:
    """A run_ts (datetime or 'YYYY-MM-DD HH:MM:SS' string) → '21-Jul 13:04'."""
    if not ts:
        return ''
    if isinstance(ts, _dt.datetime):
        return ts.strftime('%d-%b %H:%M')
    s = str(ts)
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M'):
        try:
            return _dt.datetime.strptime(s[:19], fmt).strftime('%d-%b %H:%M')
        except ValueError:
            continue
    return s[:16]


# DB segment value ↔ registry segment value ↔ display name.
_SEGMENTS = [
    {'key': 'online', 'db': 'OnlineB2B', 'reg': 'Online', 'name': 'Online B2B'},
    {'key': 'offline', 'db': 'Offline', 'reg': 'Offline', 'name': 'Offline'},
]


def _channels_for(reg_seg: str) -> list:
    """Leaf channels of one registry segment ('Online'/'Offline'), registry order."""
    parents = {c.parent for c in reg.channels() if c.parent}
    return [c for c in reg.channels()
            if c.segment == reg_seg and c.key not in parents]


def _resolve_key(marketplace: str, label: str, channels: list):
    """Map a DB (marketplace, label) group to a registry channel KEY, reusing the
    same resolution the Daily page uses (db_label → db_key), with a lenient
    key/display fallback for channels that carry no db_key (e.g. BlinkMP)."""
    m = str(marketplace or '').strip().lower()
    l = str(label or '').strip().lower()
    for c in channels:
        if c.db_label and c.db_label.strip().lower() == l:
            return c.key
    for c in channels:
        if c.db_key and c.db_key.strip().lower() == m:
            return c.key
    for c in channels:
        if c.key.strip().lower() == m or c.display.strip().lower() == m:
            return c.key
    return None


def _segment_board(day, seg_def: dict, skip_skus: bool = False) -> dict:
    """One segment's received board (Online OR Offline). Folds today's intake +
    value legs + issues + last-received into that segment's channels only —
    cross-segment marketplaces are skipped (create=False), so the two boards never
    bleed into each other."""
    data = order_db.marketplace_daily_intake(segment=seg_def['db'], day=day)
    channels = _channels_for(seg_def['reg'])

    def _blank(key, display):
        return {'channel': key, 'display': display, 'received': False,
                'pos': 0, 'items': 0, 'qty': 0, 'value': 0.0, 'excluded': 0,
                'raw_value': 0.0, 'uploaded_value': 0.0, 'excluded_value': 0.0,
                'raw_qty': 0, 'uploaded_qty': 0, 'excluded_qty': 0,
                'billing_value': 0.0, 'last_received': None, 'labels': [],
                'pos_detail': {}}       # po → per-PO legs (for the drill-down)

    buckets: dict = {c.key: _blank(c.key, c.display) for c in channels}
    order = [c.key for c in channels]

    def _bucket(mk, label, create=False):
        key = _resolve_key(mk, label, channels)
        if key is None:
            key = str(mk or label or '').strip()
        if key in buckets:
            return buckets[key]
        if create and key:
            buckets[key] = _blank(key, key)
            order.append(key)
            return buckets[key]
        return None                           # cross-segment MP → skip

    for r in data.get('today', []):
        b = _bucket(r['marketplace'], r['marketplace_label'], create=True)
        if b is None:
            continue
        b['received'] = True
        b['pos'] += int(_num(r['pos']))
        b['items'] += int(_num(r['items']))
        b['qty'] += int(_num(r['qty']))
        b['value'] += float(_num(r['value']))
        lbl = str(r.get('marketplace_label') or '').strip()
        if lbl and lbl not in b['labels']:
            b['labels'].append(lbl)

    for r in data.get('issues', []):
        b = _bucket(r['marketplace'], r['marketplace'])
        if b:
            b['excluded'] += int(_num(r['count']))

    for r in data.get('value_legs', []):
        b = _bucket(r['marketplace'], r['marketplace'])
        if b:
            b['raw_value'] += float(_num(r.get('raw_value')))
            b['uploaded_value'] += float(_num(r.get('uploaded_value')))
            b['excluded_value'] += float(_num(r.get('excluded_value')))
            b['raw_qty'] += int(_num(r.get('raw_qty')))
            b['uploaded_qty'] += int(_num(r.get('uploaded_qty')))
            b['excluded_qty'] += int(_num(r.get('excluded_qty')))

    # per-PO breakdown (for the click-to-expand drill-down) — fold into the row's
    # channel, keyed by PO.
    def _blank_pd(po):
        return {'po': po, 'raw_qty': 0, 'uploaded_qty': 0, 'excluded_qty': 0,
                'raw_value': 0.0, 'uploaded_value': 0.0, 'excluded_value': 0.0,
                'skus': {}}      # item_no → per-SKU legs (2nd drill level)

    def _add_legs(tgt, r):
        tgt['raw_qty'] += int(_num(r.get('raw_qty')))
        tgt['uploaded_qty'] += int(_num(r.get('uploaded_qty')))
        tgt['excluded_qty'] += int(_num(r.get('excluded_qty')))
        tgt['raw_value'] += float(_num(r.get('raw_value')))
        tgt['uploaded_value'] += float(_num(r.get('uploaded_value')))
        tgt['excluded_value'] += float(_num(r.get('excluded_value')))

    for r in data.get('po_legs', []):
        b = _bucket(r['marketplace'], r['marketplace'])
        if b:
            po = str(r.get('po') or '')
            _add_legs(b['pos_detail'].setdefault(po, _blank_pd(po)), r)

    # per-SKU breakdown (2nd drill level: click a PO → its SKUs)
    for r in data.get('sku_legs', []):
        b = _bucket(r['marketplace'], r['marketplace'])
        if b:
            po = str(r.get('po') or '')
            pd = b['pos_detail'].setdefault(po, _blank_pd(po))
            item = str(r.get('item_no') or '')
            sd = pd['skus'].setdefault(item, {
                'item_no': item, 'ean': str(r.get('ean') or ''),
                'description': str(r.get('description') or ''),
                'raw_qty': 0, 'uploaded_qty': 0, 'excluded_qty': 0,
                'raw_value': 0.0, 'uploaded_value': 0.0, 'excluded_value': 0.0})
            _add_legs(sd, r)

    for r in data.get('last_received', []):
        b = _bucket(r['marketplace'], r['marketplace_label'])
        if b:
            ts = r.get('last_received')
            if ts and (b['last_received'] is None
                       or str(ts) > str(b['last_received'])):
                b['last_received'] = ts

    rows = []
    for key in order:
        b = buckets[key]
        excl = b['excluded'] if b['received'] else 0
        incl = max(b['items'] - excl, 0) if b['received'] else 0
        rows.append({
            'channel': b['channel'], 'display': b['display'],
            'received': b['received'],
            'pos': b['pos'], 'items': b['items'], 'qty': b['qty'],
            'value': round(b['value'], 2),
            'raw_value': round(b['raw_value'], 2) if b['received'] else 0.0,
            'uploaded_value': round(b['uploaded_value'], 2) if b['received'] else 0.0,
            'excluded_value': round(b['excluded_value'], 2) if b['received'] else 0.0,
            'raw_qty': b['raw_qty'] if b['received'] else 0,
            'uploaded_qty': b['uploaded_qty'] if b['received'] else 0,
            'excluded_qty': b['excluded_qty'] if b['received'] else 0,
            # % of raw (shown in brackets) — uploaded/excluded for qty AND value
            'uploaded_qty_pct': _pctg(b['uploaded_qty'], b['raw_qty']) if b['received'] else 0,
            'excluded_qty_pct': _pctg(b['excluded_qty'], b['raw_qty']) if b['received'] else 0,
            'uploaded_value_pct': _pctg(b['uploaded_value'], b['raw_value']) if b['received'] else 0,
            'excluded_value_pct': _pctg(b['excluded_value'], b['raw_value']) if b['received'] else 0,
            'billing_value': None,            # per-MP; set by build_summary
            'pos_detail': [_finish_pd(pd, skip_skus=skip_skus)
                           for _po, pd in sorted(b['pos_detail'].items())
                           ] if b['received'] else [],
            'included': incl, 'excluded': excl,
            'labels': ', '.join(b['labels']),
            'last_received': _fmt_ts(b['last_received']),
        })

    recv = [r for r in rows if r['received']]
    totals = {
        'received_count': len(recv), 'total_count': len(rows),
        'pos': sum(r['pos'] for r in recv),
        'items': sum(r['items'] for r in recv),
        'qty': sum(r['qty'] for r in recv),
        'value': round(sum(r['value'] for r in recv), 2),
        'raw_value': round(sum(r['raw_value'] for r in recv), 2),
        'uploaded_value': round(sum(r['uploaded_value'] for r in recv), 2),
        'excluded_value': round(sum(r['excluded_value'] for r in recv), 2),
        'raw_qty': sum(r['raw_qty'] for r in recv),
        'uploaded_qty': sum(r['uploaded_qty'] for r in recv),
        'excluded_qty': sum(r['excluded_qty'] for r in recv),
        'included': sum(r['included'] for r in recv),
        'excluded': sum(r['excluded'] for r in recv),
        'billing_value': None,                # set by build_summary from fill-rate
    }
    totals['uploaded_qty_pct'] = _pctg(totals['uploaded_qty'], totals['raw_qty'])
    totals['excluded_qty_pct'] = _pctg(totals['excluded_qty'], totals['raw_qty'])
    totals['uploaded_value_pct'] = _pctg(totals['uploaded_value'], totals['raw_value'])
    totals['excluded_value_pct'] = _pctg(totals['excluded_value'], totals['raw_value'])
    return {'key': seg_def['key'], 'name': seg_def['name'],
            'rows': rows, 'received': recv, 'totals': totals,
            'issue_total': totals['excluded']}


def build_summary(day=None, seg_filter='', segment=None, skip_skus=False) -> dict:
    """Assemble the consolidated summary — segment boards + grand totals + embedded
    excluded lines & SKU summary. ``seg_filter`` = 'online' / 'offline' / '' (both)
    picks which segment board(s) to include (the master filter). Shared by the
    on-screen page AND the email body so they never drift. Read-only; never raises.

    The whole assembly is ~1.9 s (three day-line scans + the inventory fill-rate),
    so the result is memoised for a short window (env ``ORDERDB_SUMMARY_TTL``, default
    25 s) keyed by day+segment+skip_skus. This makes re-opening the Cockpit instant;
    an upload confirm busts the ``'summary:'`` prefix so a new PO shows at once."""
    iso = (day or order_db._ist_today().isoformat())      # 'today' = IST day
    sel0 = str(seg_filter or '').strip().lower()
    key = f"summary:{iso}:{sel0}:{int(bool(skip_skus))}"
    return order_db._stable(
        key, lambda: _build_summary(day=day, seg_filter=seg_filter,
                                    segment=segment, skip_skus=skip_skus),
        ttl=_SUMMARY_TTL)


def _build_summary(day=None, seg_filter='', segment=None, skip_skus=False) -> dict:
    iso = (day or order_db._ist_today().isoformat())
    try:
        nice = _dt.date.fromisoformat(iso).strftime('%A, %d %b %Y')
    except (ValueError, TypeError):
        nice = iso

    sel = str(seg_filter or '').strip().lower()
    seg_defs = [s for s in _SEGMENTS if s['key'] == sel] if sel in ('online', 'offline') \
        else _SEGMENTS

    boards = []
    for seg_def in seg_defs:
        board = _segment_board(iso, seg_def, skip_skus=skip_skus)
        # Tentative billing = fill-rate-adjusted (uploaded × stock availability),
        # now COMPUTABLE from the Inventory snapshot. None when no stock uploaded
        # for that segment → the UI still shows a "with Inventory" placeholder.
        try:
            from . import inventory_fill
            fr = inventory_fill.fill_rate(date_from=iso, date_to=iso,
                                          segment=seg_def['key'])
            if fr.get('ok') and fr.get('has_stock'):
                # Fold per-MP tentative billing onto each board row (resolve the
                # fill-rate marketplace to this segment's channel), so the column
                # SUMS to the segment total (owner requirement).
                channels = _channels_for(seg_def['reg'])
                ch_bill: dict = {}
                ch_qty: dict = {}
                for g in fr.get('mps', []):
                    key = (_resolve_key(g.get('label'), g.get('label'), channels)
                           or str(g.get('label') or ''))
                    ch_bill[key] = ch_bill.get(key, 0.0) + float(g.get('billing') or 0)
                    ch_qty[key] = ch_qty.get(key, 0.0) + float(g.get('fillable_qty') or 0)
                po_bill = {str(g.get('label')): float(g.get('billing') or 0)
                           for g in fr.get('pos', [])}
                po_qty = {str(g.get('label')): float(g.get('fillable_qty') or 0)
                          for g in fr.get('pos', [])}
                # Fill rate = fillable (in-stock) share of what we UPLOADED, both
                # qty-wise and value-wise (value-wise = tentative billing). % is
                # against the row's own Uploaded leg so it reconciles on-screen.
                for row in board['rows']:
                    row['billing_value'] = round(ch_bill.get(row['channel'], 0.0), 2)
                    row['billing_qty'] = round(ch_qty.get(row['channel'], 0.0), 1)
                    row['fill_qty_pct'] = _pctg(row['billing_qty'], row.get('uploaded_qty') or 0)
                    row['fill_val_pct'] = _pctg(row['billing_value'], row.get('uploaded_value') or 0)
                    for pd in row.get('pos_detail', []):
                        pd['billing_value'] = (round(po_bill.get(str(pd['po']), 0.0), 2)
                                               or None)
                        pd['billing_qty'] = (round(po_qty.get(str(pd['po']), 0.0), 1)
                                             or None)
                        pd['fill_qty_pct'] = _pctg(pd.get('billing_qty') or 0, pd.get('uploaded_qty') or 0)
                        pd['fill_val_pct'] = _pctg(pd.get('billing_value') or 0, pd.get('uploaded_value') or 0)
                board['totals']['billing_value'] = round(
                    sum(r['billing_value'] or 0 for r in board['received']), 2)
                board['totals']['billing_qty'] = round(
                    sum(r.get('billing_qty') or 0 for r in board['received']), 1)
                board['totals']['fill_qty_pct'] = _pctg(
                    board['totals']['billing_qty'], board['totals'].get('uploaded_qty') or 0)
                board['totals']['fill_val_pct'] = _pctg(
                    board['totals']['billing_value'], board['totals'].get('uploaded_value') or 0)
                board['fill'] = {
                    'fill_pct': fr['totals'].get('fill_pct'),
                    'oos_qty': fr['totals'].get('oos_qty'),
                    'stock_as_of': fr.get('stock_as_of'),
                    'affected_pos': fr['totals'].get('affected_pos'),
                    'clean_pos': fr['totals'].get('clean_pos')}
        except Exception:                     # noqa: BLE001 — best-effort
            pass
        boards.append(board)

    def _sum(field):
        return sum(b['totals'].get(field, 0) or 0 for b in boards)
    bill = [b['totals'].get('billing_value') for b in boards
            if b['totals'].get('billing_value') is not None]
    grand = {
        'received_count': _sum('received_count'), 'total_count': _sum('total_count'),
        'pos': _sum('pos'), 'items': _sum('items'), 'qty': _sum('qty'),
        'raw_value': round(_sum('raw_value'), 2),
        'uploaded_value': round(_sum('uploaded_value'), 2),
        'excluded_value': round(_sum('excluded_value'), 2),
        'raw_qty': _sum('raw_qty'), 'uploaded_qty': _sum('uploaded_qty'),
        'excluded_qty': _sum('excluded_qty'),
        'excluded': _sum('excluded'),
        'billing_value': round(sum(bill), 2) if bill else None,
    }

    # (Excluded-lines + SKU-summary sections were removed from the email — the
    # digest is now the single MP-wise received board per segment. Full detail
    # lives on the Issues + SKU Summary tabs.)
    return {'day': iso, 'day_nice': nice, 'segments': boards, 'grand': grand,
            'seg_filter': sel if sel in ('online', 'offline') else 'both',
            'issue_total': grand['excluded']}


class SummaryEmailReport(EmailReport):
    """Consolidated daily summary → a review-and-send email (reuses the shared
    SMTP send path via :class:`EmailReport`)."""

    def __init__(self, day=None, segment: str = 'OnlineB2B', note: str = '',
                 subject: str = '', to=None, cc=None, seg_filter=''):
        self.note = (note or '').strip()
        self._subject = (subject or '').strip()
        self._to = _clean_emails(to) if to is not None else None
        self._cc = _clean_emails(cc) if cc is not None else None
        # skip_skus=True: neither the page (SKUs lazy-load on click) nor the email
        # (MP-level board) renders SKU rows inline — so never build them upfront.
        self.data = build_summary(day=day, seg_filter=seg_filter, skip_skus=True)

    # ── recipients (typed in the page; None → config defaults) ──
    def to(self):
        return self._to

    def cc(self):
        return self._cc

    def subject(self) -> str:
        if self._subject:
            return self._subject
        t = self.data['grand']
        return (f"Daily Summary — {self.data['day_nice']}: "
                f"{t['received_count']}/{t['total_count']} received · "
                f"{_grp(t['pos'])} PO · {_inr(t['raw_value'])} ordered · "
                f"{_inr(t['uploaded_value'])} to D365")

    # ── body pieces ──
    def _note_block(self) -> str:
        if not self.note:
            return ''
        body = escape(self.note).replace('\n', '<br>')
        return (
            '<div style="margin:0 0 16px;padding:12px 14px;border-radius:10px;'
            'background:#eef2ff;border:1px solid #c7d2fe;">'
            '<div style="font-size:10.5px;font-weight:800;letter-spacing:.04em;'
            'text-transform:uppercase;color:#3730a3;margin-bottom:5px;">'
            'Note from sender</div>'
            f'<div style="font-size:13px;color:#1e293b;line-height:1.5;">{body}</div>'
            '</div>')

    def _card(self, label, value, colour, accent, sub='', wide=False):
        lbl = ('font-size:9.5px;font-weight:800;letter-spacing:.06em;'
               'text-transform:uppercase;color:#7c8698;')
        val = f'font-size:{"20px" if wide else "23px"};font-weight:800;line-height:1.05;margin-top:5px;'
        card = ('display:inline-block;box-sizing:border-box;'
                f'min-width:{"150px" if wide else "104px"};'
                'padding:12px 14px 13px;border-radius:14px;border:1px solid #eceff4;'
                'margin:0 9px 9px 0;vertical-align:top;background:#ffffff;'
                f'border-top:3px solid {accent};box-shadow:0 2px 6px rgba(20,30,60,.05);')
        subhtml = (f'<div style="font-size:10px;color:#9aa1b2;margin-top:4px;">{sub}</div>'
                   if sub else '')
        return (f'<div style="{card}"><div style="{lbl}">{label}</div>'
                f'<div style="{val}color:{colour};">{value}</div>{subhtml}</div>')

    def _kpi_block(self, board) -> str:
        t = board['totals']
        c = self._card
        # Row 1 — operational counts
        ops = (
            '<div style="margin:0 0 4px;">'
            + c('Received', f"{t['received_count']}/{t['total_count']}", '#0f172a', '#6366f1')
            + c('PO count', _grp(t['pos']), '#0f172a', '#10b981')
            + c('Items', _grp(t['items']), '#0f172a', '#3949AB')
            + c('Qty', _grp(t['qty']), '#0f172a', '#3949AB')
            + c('Excluded lines', _grp(t['excluded']),
                '#b45309' if t['excluded'] else '#64748b', '#f59e0b')
            + '</div>')
        # Row 2 — the VALUE breakdown (sum-wise): raw = uploaded + excluded
        has_bill = t['billing_value'] is not None
        bill = _inr(t['billing_value']) if has_bill else '—'
        bill_col = '#0f9d6b' if has_bill else '#94a3b8'
        bill_acc = '#00b894' if has_bill else '#cbd5e1'
        fill = board.get('fill') or {}
        bill_sub = (f"fill {fill.get('fill_pct')}% × uploaded" if has_bill
                    else 'available with Inventory')
        vals = (
            '<div style="font-size:10.5px;font-weight:800;letter-spacing:.05em;'
            'text-transform:uppercase;color:#334155;margin:14px 0 8px;">'
            'Value breakdown &middot; today</div>'
            '<div style="margin:0 0 18px;">'
            + c('Total PO value', _inr(t['raw_value']), '#0f172a', '#3949AB', 'raw &mdash; as ordered', True)
            + c('Uploaded value', _inr(t['uploaded_value']), '#0f9d6b', '#00b894', 'clean &rarr; D365', True)
            + c('Excluded value', _inr(t['excluded_value']),
                '#b45309' if t['excluded_value'] else '#64748b', '#f59e0b', 'dropped lines', True)
            + c('Tentative billing', bill, bill_col, bill_acc, bill_sub, True)
            + '</div>')
        return ops + vals

    def _board(self, board) -> str:
        # "Inventory as of …" — the stock snapshot the tentative fill was computed
        # against (same figure the on-page cockpit shows).
        _sa = (board.get('fill') or {}).get('stock_as_of')
        _asof = (f' &middot; <b>inventory as of {escape(str(_sa))}</b>' if _sa else '')
        th = ('padding:10px 12px;text-align:left;font-size:10px;color:#64748b;'
              'white-space:nowrap;text-transform:uppercase;letter-spacing:.05em;'
              'font-weight:800;border-bottom:1px solid #e5e9f0;')
        thr = th + 'text-align:right;'
        rows_html, i = [], 0
        for r in board['rows']:
            i += 1
            bg = '#ffffff' if i % 2 else '#fafbfd'
            td = (f'padding:10px 12px;font-size:12.5px;vertical-align:middle;'
                  f'background:{bg};border-bottom:1px solid #f0f2f6;')
            tdr = td + 'text-align:right;font-variant-numeric:tabular-nums;'
            if r['received']:
                pill = ('<span style="display:inline-block;padding:3px 10px;border-radius:20px;'
                        'font-weight:700;font-size:11px;color:#0f9d6b;background:#e7f6ef;'
                        'white-space:nowrap;">&#10003; Received</span>')
                bv = r.get('billing_value')
                pcts = 'color:#94a3b8;font-weight:400;font-size:10px;'
                subs = 'font-size:11px;color:#64748b;margin-top:2px;'   # value line

                def _leg(qty, val, qpct=None, vpct=None, colour='#0f172a'):
                    """One stacked cell: qty (top) + ₹value (bottom), each with %."""
                    qp = (f' <span style="{pcts}">({qpct}%)</span>'
                          if qpct is not None else '')
                    vp = (f' <span style="{pcts}">({vpct}%)</span>'
                          if vpct is not None else '')
                    return (f'<td style="{tdr}">'
                            f'<div style="color:{colour};font-weight:700;">{_grp(qty)}{qp}</div>'
                            f'<div style="{subs}">{_inr(val)}{vp}</div></td>')
                cells = (
                    f'<td style="{tdr}">{_grp(r["pos"])}</td>'
                    + _leg(r['raw_qty'], r['raw_value'])
                    + _leg(r['uploaded_qty'], r['uploaded_value'],
                           r['uploaded_qty_pct'], r['uploaded_value_pct'], '#0f9d6b')
                    + (_leg(r['excluded_qty'], r['excluded_value'],
                            r['excluded_qty_pct'], r['excluded_value_pct'], '#b45309')
                       if r['excluded_qty'] or r['excluded_value']
                       else f'<td style="{tdr}color:#94a3b8;">—</td>')
                    # Tentative billing — qty (fill-rate adj.) on top, ₹ below,
                    # each with its fill % (same shape as the on-page cockpit).
                    + (_leg(r.get('billing_qty') or 0, bv,
                            r.get('fill_qty_pct'), r.get('fill_val_pct'), '#0f9d6b')
                       if bv else f'<td style="{tdr}color:#94a3b8;">—</td>'))
            else:
                last = (f'<span style="color:#94a3b8;font-size:11.5px;">last received '
                        f'<b style="color:#64748b;">{escape(r["last_received"])}</b></span>'
                        if r['last_received'] else
                        '<span style="color:#cbd2dc;font-size:11.5px;">no prior record</span>')
                pill = ('<span style="display:inline-block;padding:3px 10px;border-radius:20px;'
                        'font-weight:700;font-size:11px;color:#94a3b8;background:#f1f5f9;'
                        'white-space:nowrap;">&#9675; Not today</span>')
                cells = (f'<td colspan="5" style="{td}text-align:right;">{last}</td>')
            name = (f'<span style="font-weight:700;color:#0f172a;font-size:13px;">'
                    f'{escape(r["display"])}</span>'
                    + (f'<div style="font-size:10px;color:#a0a7b4;">{escape(r["labels"])}</div>'
                       if r['labels'] and r['labels'].lower() != r['display'].lower() else ''))
            rows_html.append(
                f'<tr><td style="{td}">{name}</td>'
                f'<td style="{td}white-space:nowrap;">{pill}</td>{cells}</tr>')

        # ── TOTAL row (matches the KPI cards exactly) ──
        t = board['totals']
        ttd = ('padding:11px 12px;font-size:12.5px;background:#eef2ff;'
               'border-top:2px solid #c7d2fe;vertical-align:middle;')
        ttr = ttd + 'text-align:right;font-variant-numeric:tabular-nums;'
        tp = 'color:#7c8698;font-weight:400;font-size:10px;'
        tsub = 'font-size:11px;color:#475569;margin-top:2px;'

        def _tleg(qty, val, qpct=None, vpct=None, colour='#1A237E'):
            qp = f' <span style="{tp}">({qpct}%)</span>' if qpct is not None else ''
            vp = f' <span style="{tp}">({vpct}%)</span>' if vpct is not None else ''
            return (f'<td style="{ttr}"><div style="color:{colour};font-weight:800;">'
                    f'{_grp(qty)}{qp}</div>'
                    f'<div style="{tsub}font-weight:700;">{_inr(val)}{vp}</div></td>')
        bvt = t.get('billing_value')
        total_row = (
            f'<tr><td style="{ttd}font-weight:800;color:#1A237E;">TOTAL</td>'
            f'<td style="{ttd}"></td><td style="{ttr}font-weight:800;">{_grp(t["pos"])}</td>'
            + _tleg(t['raw_qty'], t['raw_value'])
            + _tleg(t['uploaded_qty'], t['uploaded_value'],
                    t.get('uploaded_qty_pct'), t.get('uploaded_value_pct'), '#0f9d6b')
            + _tleg(t['excluded_qty'], t['excluded_value'],
                    t.get('excluded_qty_pct'), t.get('excluded_value_pct'), '#b45309')
            + (_tleg(t.get('billing_qty') or 0, bvt,
                     t.get('fill_qty_pct'), t.get('fill_val_pct'), '#0f9d6b')
               if bvt else f'<td style="{ttr}color:#94a3b8;font-weight:800;">—</td>')
            + '</tr>')
        rows_html.append(total_row)
        return (
            '<div style="font-size:12px;font-weight:800;letter-spacing:.05em;'
            'text-transform:uppercase;color:#334155;margin:8px 0 8px;">'
            'Received board · all marketplaces</div>'
            '<table style="border-collapse:separate;border-spacing:0;width:100%;'
            'border:1px solid #e5e9f0;border-radius:12px;overflow:hidden;">'
            f'<thead><tr><th style="{th}">Marketplace</th><th style="{th}">Status</th>'
            f'<th style="{thr}">PO</th>'
            f'<th style="{thr}">Raw<br><span style="font-weight:600;color:#94a3b8;">qty / &#8377;</span></th>'
            f'<th style="{thr}">Uploaded<br><span style="font-weight:600;color:#94a3b8;">qty / &#8377;</span></th>'
            f'<th style="{thr}">Excluded<br><span style="font-weight:600;color:#94a3b8;">qty / &#8377;</span></th>'
            f'<th style="{thr}">Tentative bill<br><span style="font-weight:600;color:#94a3b8;">qty / &#8377;</span></th>'
            f'</tr></thead>'
            f'<tbody>{"".join(rows_html)}</tbody></table>'
            '<div style="font-size:10.5px;color:#a0a7b4;margin:8px 2px 0;">'
            'Each cell = qty (top) / value &#8377; (bottom), % of raw in brackets &middot; '
            'Raw = as ordered (inc-GST), Uploaded = clean &rarr; D365, Excluded = dropped '
            '&middot; <b>Raw = Uploaded + Excluded</b> &middot; Tentative bill = fill-rate &times; '
            'uploaded (qty &amp; &#8377;)' + _asof + '.</div>')

    def _issue_ref(self, board) -> str:
        n = board['issue_total']
        if not n:
            return ('<div style="margin:14px 0 0;font-size:12px;color:#0f9d6b;">'
                    '&#10003; No issue lines today — nothing excluded.</div>')
        return (
            '<div style="margin:14px 0 0;padding:11px 14px;border-radius:10px;'
            'background:#fff7ed;border:1px solid #fed7aa;font-size:12.5px;color:#9a3412;">'
            f'<b>{n} issue line(s)</b> today were excluded (price mismatch / not in '
            'master) — detailed below. Full ongoing detail lives on the '
            '<b>Issues tab</b> of the dashboard.'
            '</div>')

    def _sec_title(self, title, count, colour) -> str:
        return (f'<div style="margin:26px 0 8px;font-size:12px;font-weight:800;'
                f'letter-spacing:.05em;text-transform:uppercase;color:{colour};">'
                f'{title}'
                + (f' <span style="color:#94a3b8;font-weight:700;">({count})</span>'
                   if count is not None else '') + '</div>')

    def _excluded_section(self) -> str:
        """Today's EXCLUDED (dropped) lines — same columns as the Issues-tab email,
        embedded so the digest is self-contained."""
        rows = self.data.get('excluded_lines', [])
        if not rows:
            return (self._sec_title('Excluded lines', 0, '#b45309')
                    + '<div style="font-size:12px;color:#0f9d6b;">&#10003; No lines '
                      'were excluded today.</div>')
        th = ('padding:8px 10px;text-align:left;font-size:10px;color:#7a2e0a;'
              'background:#fff4ec;white-space:nowrap;text-transform:uppercase;'
              'letter-spacing:.04em;border-bottom:2px solid #fed7aa;')
        thr = th + 'text-align:right;'
        out = []
        for i, r in enumerate(rows):
            bg = '#ffffff' if i % 2 == 0 else '#fffaf5'
            td = (f'padding:7px 10px;font-size:12px;background:{bg};'
                  'border-bottom:1px solid #f4ede6;')
            tdr = td + 'text-align:right;font-variant-numeric:tabular-nums;'
            mono = 'font-family:Consolas,monospace;font-size:11px;'
            out.append(
                f'<tr><td style="{td}">{escape(str(r.get("marketplace") or "-"))}</td>'
                f'<td style="{td}{mono}">{escape(str(r.get("po") or "-"))}</td>'
                f'<td style="{td}{mono}">{escape(str(r.get("item_no") or "-"))}</td>'
                f'<td style="{td}{mono}">{escape(str(r.get("ean") or "-"))}</td>'
                f'<td style="{td}">{escape(str(r.get("description") or "-"))}</td>'
                f'<td style="{tdr}">{_grp(r.get("qty"))}</td>'
                f'<td style="{tdr}">{_inr(r.get("our_cp"))}</td>'
                f'<td style="{tdr}">{_inr(r.get("vendor_cp"))}</td>'
                f'<td style="{td}">{escape(str(r.get("status") or "-"))}</td>'
                f'<td style="{td}">{escape(str(r.get("remark") or ""))}</td></tr>')
        return (
            self._sec_title('Excluded lines', len(rows), '#b45309')
            + '<div style="font-size:11px;color:#9a6a4a;margin:-2px 0 8px;">'
              'Dropped from the confirmed PO — the real loss (same as the Issues tab).</div>'
            '<table style="border-collapse:separate;border-spacing:0;width:100%;'
            'border:1px solid #fde3cc;border-radius:10px;overflow:hidden;">'
            f'<thead><tr><th style="{th}">MP</th><th style="{th}">PO</th>'
            f'<th style="{th}">Item</th><th style="{th}">EAN</th>'
            f'<th style="{th}">Description</th><th style="{thr}">Qty</th>'
            f'<th style="{thr}">Our CP</th><th style="{thr}">Their CP</th>'
            f'<th style="{th}">Status</th><th style="{th}">Remark</th></tr></thead>'
            f'<tbody>{"".join(out)}</tbody></table>')

    def _sku_section(self) -> str:
        """Today's SKU-wise rollup — the SKU-Summary tab, day-scoped & compact."""
        sku = self.data.get('sku', {})
        rows = sku.get('rows', [])
        if not rows:
            return ''
        tt = sku.get('totals', {})
        th = ('padding:8px 10px;text-align:left;font-size:10px;color:#3730a3;'
              'background:#eef2ff;white-space:nowrap;text-transform:uppercase;'
              'letter-spacing:.04em;border-bottom:2px solid #c7d2fe;')
        thr = th + 'text-align:right;'
        out = []
        for i, r in enumerate(rows):
            bg = '#ffffff' if i % 2 == 0 else '#f8f9ff'
            td = (f'padding:7px 10px;font-size:12px;background:{bg};'
                  'border-bottom:1px solid #eef0f8;')
            tdr = td + 'text-align:right;font-variant-numeric:tabular-nums;'
            mis = int(r.get('mis_qty') or 0)
            nim = int(r.get('nim_qty') or 0)
            flag = ''
            if mis:
                flag += (f'<span style="color:#b45309;">M {_grp(mis)}</span> ')
            if nim:
                flag += (f'<span style="color:#b91c1c;">NIM {_grp(nim)}</span>')
            mps = str(r.get('marketplaces') or '')
            out.append(
                f'<tr><td style="{td}font-family:Consolas,monospace;font-size:11px;">'
                f'{escape(str(r.get("item_no") or "-"))}</td>'
                f'<td style="{td}">{escape(str(r.get("description") or "-"))}</td>'
                f'<td style="{tdr}font-weight:700;">{_grp(r.get("tot_qty"))}</td>'
                f'<td style="{tdr}color:#0f9d6b;">{_grp(r.get("ok_qty"))}</td>'
                f'<td style="{td}">{flag or "&mdash;"}</td>'
                f'<td style="{tdr}">{int(r.get("mp_count") or 0)}</td>'
                f'<td style="{td}font-size:10.5px;color:#64748b;">{escape(mps)}</td></tr>')
        note = ''
        if sku.get('capped'):
            more = int(sku.get('total_skus', 0)) - len(rows)
            note = ('<div style="font-size:11px;color:#7c8698;margin:8px 2px 0;">'
                    f'Showing top {len(rows)} of {_grp(sku.get("total_skus"))} SKUs '
                    f'&mdash; {_grp(more)} more on the <b>SKU Summary tab</b>.</div>')
        tot = (f'<div style="font-size:11px;color:#64748b;margin:-2px 0 8px;">'
               f'{_grp(sku.get("total_skus"))} SKUs today &middot; total qty '
               f'{_grp(tt.get("qty"))} (OK {_grp(tt.get("ok"))} &middot; '
               f'mismatch {_grp(tt.get("mismatch"))} &middot; '
               f'not-in-master {_grp(tt.get("nim"))}).</div>')
        return (
            self._sec_title('SKU summary', sku.get('total_skus'), '#3730a3')
            + tot
            + '<table style="border-collapse:separate;border-spacing:0;width:100%;'
            'border:1px solid #dbe0f5;border-radius:10px;overflow:hidden;">'
            f'<thead><tr><th style="{th}">Item</th><th style="{th}">Description</th>'
            f'<th style="{thr}">Total qty</th><th style="{thr}">OK</th>'
            f'<th style="{th}">Issues</th><th style="{thr}">MPs</th>'
            f'<th style="{th}">Marketplaces</th></tr></thead>'
            f'<tbody>{"".join(out)}</tbody></table>' + note)

    def _hero(self) -> str:
        d = self.data
        t = d['grand']
        return f"""
  <div style="background:#1A237E;background:linear-gradient(120deg,#1A237E 0%,#3949AB 55%,#5C6BC0 100%);
              border-radius:18px;padding:22px 24px;color:#ffffff;margin:0 0 18px;
              box-shadow:0 10px 26px rgba(26,35,126,.28);">
    <div style="font-size:11px;font-weight:800;letter-spacing:.18em;text-transform:uppercase;
                color:#c5cae9;">REN&Eacute;E &middot; Daily Summary &middot; Online + Offline</div>
    <div style="font-size:21px;font-weight:800;margin:4px 0 2px;">{escape(d['day_nice'])}</div>
    <div style="font-size:12.5px;color:#dfe3f7;">
      {t['received_count']} of {t['total_count']} marketplaces received today
      &middot; {_grp(t['pos'])} PO &middot; {_grp(t['qty'])} qty
      &middot; {_inr(t['raw_value'])} ordered &middot; {_inr(t['uploaded_value'])} to D365</div>
  </div>"""

    def _segment_section(self, board) -> str:
        """One segment (Online / Offline): heading + KPIs + received board."""
        t = board['totals']
        return (
            '<div style="margin:24px 0 10px;padding:9px 14px;border-radius:11px;'
            'background:#f1f4fb;border:1px solid #e2e8f4;display:flex;'
            'align-items:baseline;gap:10px;">'
            f'<span style="font-size:14px;font-weight:800;color:#1A237E;">'
            f'{escape(board["name"])}</span>'
            f'<span style="font-size:11.5px;color:#64748b;">'
            f'{t["received_count"]}/{t["total_count"]} received &middot; '
            f'{_grp(t["pos"])} PO &middot; {_inr(t["raw_value"])} ordered</span></div>'
            + self._kpi_block(board)
            + self._board(board))

    def html(self) -> str:
        segments = ''.join(self._segment_section(b) for b in self.data['segments'])
        return f"""\
<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#eef1f7;">
<div style="font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#0f172a;
            max-width:1040px;margin:0 auto;padding:22px;background:#eef1f7;">
  {self._hero()}
  <div style="background:#ffffff;border-radius:18px;padding:22px 24px;
              box-shadow:0 6px 20px rgba(20,30,60,.06);">
    {self._note_block()}
    {segments}
    <p style="margin:18px 0 0;font-size:11px;color:#a0a7b4;border-top:1px solid #eef0f4;padding-top:12px;">
      Consolidated summary — auto-generated from Claude AI
      &middot; {order_db._ist_now():%d-%b-%Y %H:%M} IST. Full issue-line + SKU detail: Issues &amp; SKU Summary tabs.
    </p>
  </div>
</div>
</body></html>"""
