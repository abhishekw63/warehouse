"""
online_b2b.services.issue_email
===============================

The **Issues** email report — built on the reusable :mod:`mailer` skeleton.

Emails the operator's DECIDED issue lines to management, in two honest buckets:

  * **Excluded** — lines dropped from the confirmed PO. This is the real *loss*.
  * **Included (intimation)** — lines that hit a CP issue but were **kept anyway**:
      - *Include (their CP)*  → action ``INCLUDE``  (kept at the vendor's CP)
      - *Include (our CP)*    → action ``OVERRIDE`` (kept, repriced to our CP)
    These are **NOT a loss** — they were processed into the order. We list them
    purely as an **intimation/record**: "we faced this issue and included it", so
    if anything comes up later there's a paper trail.

The view layer passes the same filter dict the Issues page / export use, so the
email always matches what the operator is looking at.
"""
from __future__ import annotations

import datetime as _dt
import re as _re
from decimal import Decimal, InvalidOperation
from html import escape

from . import order_db
from .mailer import EmailReport

# Operator action → human label (aligned with the review-page vocabulary).
_ACTION_LABEL = {
    'EXCLUDE': 'Excluded',
    'OVERRIDE': 'Included (our CP)',
    'INCLUDE': 'Included (their CP)',
    'KEEP': 'Kept (flagged)',
    '': '— (no action yet)',
}
_ACTION_COLOR = {
    'EXCLUDE': '#b91c1c',       # red — the actual drop / loss
    'OVERRIDE': '#0a7d5a',      # green — included at our CP
    'INCLUDE': '#1d4ed8',       # blue — included at their CP
    'KEEP': '#1d4ed8', '': '#6b7280',
}
_INCLUDE_ACTIONS = ('INCLUDE', 'OVERRIDE')


def _fmt(v) -> str:
    if v is None or v == '':
        return '—'
    return escape(str(v))


_EMAIL_RE = _re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')


def _clean_emails(v) -> list:
    """Normalise a recipient value (list / comma-or-newline-separated string)
    to a de-duplicated list of syntactically-valid addresses."""
    if not v:
        return []
    if isinstance(v, str):
        parts = _re.split(r'[,;\n]+', v)
    else:
        parts = list(v)
    out, seen = [], set()
    for p in parts:
        e = str(p).strip()
        low = e.lower()
        if e and _EMAIL_RE.match(e) and low not in seen:
            seen.add(low)
            out.append(e)
    return out


def _num(v) -> Decimal:
    """Coerce a cell (Decimal / int / str / None) to Decimal — blank → 0."""
    if v is None or v == '':
        return Decimal('0')
    if isinstance(v, Decimal):
        return v
    try:
        return Decimal(str(v))
    except (InvalidOperation, ValueError, TypeError):
        return Decimal('0')


def _ind(v, dp=0) -> str:
    """Indian digit grouping — 22,47,616.24 (last 3 digits, then groups of 2)."""
    try:
        n = float(v)
    except (TypeError, ValueError):
        n = 0.0
    neg = n < 0
    if dp:
        whole, frac = f'{abs(n):.{dp}f}'.split('.')
    else:
        whole, frac = f'{int(round(abs(n)))}', ''
    if len(whole) > 3:
        head, last3 = whole[:-3], whole[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        whole = ','.join(parts) + ',' + last3
    s = whole + (('.' + frac) if frac else '')
    return ('-' + s) if neg else s


def _rupee(v) -> str:
    """₹ with 2 dp, Indian-grouped (₹22,47,616.24)."""
    return f'₹{_ind(v, 2)}'


def _unit_rate(r: dict) -> Decimal:
    """Our expected per-unit rate on the line's comparison basis: our_cp for
    CP-based lines, our_landing for landing-based (Flipkart), falling back to
    our_mrp so a value is never silently 0 when a rate is present."""
    basis = (r.get('basis') or '').upper()
    if basis == 'LANDING' and r.get('our_landing') is not None:
        return _num(r.get('our_landing'))
    if r.get('our_cp') is not None:
        return _num(r.get('our_cp'))
    if r.get('our_landing') is not None:
        return _num(r.get('our_landing'))
    return _num(r.get('our_mrp'))


def _mp_of(r: dict) -> str:
    return str(r.get('marketplace') or r.get('marketplace_label') or r.get('mp') or '—')


class IssuesEmailReport(EmailReport):
    """Decided issue lines (current filter) → management email.

    Two buckets: **Excluded** (loss) + **Included/intimation** (kept despite the
    CP issue). Every non-action filter (marketplace / status / date / search)
    still applies."""

    def __init__(self, filters: dict | None = None, note: str = '',
                 to=None, cc=None):
        self.filters = dict(filters or {})
        self.note = (note or '').strip()
        self._to = _clean_emails(to) if to is not None else None
        self._cc = _clean_emails(cc) if cc is not None else None
        # Fetch across ALL resolution states first (the page's resolution filter
        # only scopes the on-screen view); we then keep the DECIDED lines.
        fetch = {**self.filters, 'resolution': 'all'}
        data = order_db.issues(limit=0, **fetch)
        all_rows = data.get('rows', []) if data.get('ok') else []
        # DECIDED lines only = Excluded + Included(their/our CP). Undecided /
        # legacy KEEP flagged lines stay on the Issues page, not in the mail.
        self.excluded = [r for r in all_rows
                         if (r.get('action') or '').upper() == 'EXCLUDE']
        self.included = [r for r in all_rows
                         if (r.get('action') or '').upper() in _INCLUDE_ACTIONS]
        self.rows = self.excluded + self.included    # everything the mail lists
        # EAN remaps in scope: the marketplace sent a wrong/variant EAN that we
        # remapped to the correct one (received_ean ≠ shipped ean). Listed in
        # their OWN table so the ecom team sees exactly which SKUs need the vendor
        # to fix the barcode. Drawn from ALL flagged rows, not just decided ones.
        self.remaps = [
            r for r in all_rows
            if str(r.get('received_ean') or '').strip()
            and str(r.get('received_ean')).strip() != str(r.get('ean') or '').strip()]
        self.total_flagged = len(all_rows)
        self.tally: dict = {}
        for r in self.rows:
            a = (r.get('action') or '').upper()
            self.tally[a] = self.tally.get(a, 0) + 1
        self.summary = self._compute_summary()

    # ── recipient overrides (from the modal) ────────────────────────────
    def to(self):
        return self._to

    def cc(self):
        return self._cc

    # ── summary metrics ─────────────────────────────────────────────────
    def _compute_summary(self) -> dict:
        """Excluded (loss) + Included (intimation, NOT a loss) totals, plus a
        per-marketplace breakdown of both. Loss = excluded value ONLY."""
        exc_qty = inc_qty = 0
        exc_val = Decimal('0')
        inc_val = Decimal('0')
        by_mp: dict = {}   # mp → {exc_qty, exc, inc_qty, inc}

        def _g(mp):
            return by_mp.setdefault(mp, {'exc_qty': 0, 'exc': Decimal('0'),
                                         'inc_qty': 0, 'inc': Decimal('0')})
        for r in self.excluded:
            qty = int(_num(r.get('qty')))
            lv = _unit_rate(r) * qty
            exc_qty += qty
            exc_val += lv
            g = _g(_mp_of(r))
            g['exc_qty'] += qty
            g['exc'] += lv
        for r in self.included:
            qty = int(_num(r.get('qty')))
            lv = _unit_rate(r) * qty
            inc_qty += qty
            inc_val += lv
            g = _g(_mp_of(r))
            g['inc_qty'] += qty
            g['inc'] += lv
        # Uploaded-% is about the drop: (lot − excluded qty) ÷ lot per MP.
        lot = order_db.mp_lot_qty(
            marketplace=self.filters.get('marketplace', '') or '',
            date_from=self.filters.get('date_from', '') or '',
            date_to=self.filters.get('date_to', '') or '')
        for mp, g in by_mp.items():
            lot_qty = lot.get(mp) or (g['exc_qty'] + g['inc_qty'])
            g['lot_qty'] = lot_qty
            up = lot_qty - g['exc_qty']
            g['uploaded_pct'] = (Decimal(up) / Decimal(lot_qty) * 100) if lot_qty else Decimal('0')
        return {
            'excluded_qty': exc_qty, 'excluded_value': exc_val,
            'included_qty': inc_qty, 'included_value': inc_val,
            'loss_excluded': exc_val,        # loss = excluded value only
            'by_mp': by_mp,
        }

    # ── header ──────────────────────────────────────────────────────────
    def subject(self) -> str:
        mp = self.filters.get('marketplace') or 'All MPs'
        d = _dt.date.today().strftime('%d-%b-%Y')
        return (f"Online B2B — Issue lines: {len(self.excluded)} excluded, "
                f"{len(self.included)} included (intimation) [{mp}] — {d}")

    # ── body ────────────────────────────────────────────────────────────
    def _summary_line(self) -> str:
        if not self.rows:
            return 'No decided issue lines in the selected filter.'
        parts = []
        if self.excluded:
            parts.append(f"<b>{len(self.excluded)} excluded</b> (dropped)")
        if self.included:
            parts.append(f"<b>{len(self.included)} included</b> despite the issue "
                         f"(intimation — not a loss)")
        tail = ''
        undecided = self.total_flagged - len(self.rows)
        if undecided > 0:
            tail = (f" &nbsp;({undecided} other flagged line(s) still awaiting a "
                    f"decision are not listed here)")
        return ' &nbsp;·&nbsp; '.join(parts) + '.' + tail

    def _scope_line(self) -> str:
        f = self.filters
        bits = ["Scope: <b>decided lines (excluded + included)</b>"]
        if f.get('marketplace'):
            bits.append(f"Marketplace: <b>{escape(f['marketplace'])}</b>")
        if f.get('status'):
            bits.append(f"Status: <b>{escape(f['status'])}</b>")
        if f.get('date_from') or f.get('date_to'):
            bits.append("Upload date: <b>"
                        f"{escape(f.get('date_from') or '…')} → "
                        f"{escape(f.get('date_to') or '…')}</b>")
        if f.get('q'):
            bits.append(f"Search: <b>{escape(f['q'])}</b>")
        return ' &nbsp;·&nbsp; '.join(bits)

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

    def _summary_block(self) -> str:
        """KPI cards: Excluded (loss, red) + Included (intimation, green)."""
        s = self.summary
        card = ('display:inline-block;box-sizing:border-box;min-width:145px;'
                'padding:10px 14px;border-radius:10px;border:1px solid #e5e7eb;'
                'margin:0 8px 8px 0;vertical-align:top;background:#f8fafc;')
        lbl = ('font-size:10px;font-weight:800;letter-spacing:.05em;'
               'text-transform:uppercase;color:#64748b;')
        val = 'font-size:26px;font-weight:800;color:#0f172a;margin-top:3px;line-height:1.15;'
        red = val.replace('#0f172a', '#b91c1c')
        grn = val.replace('#0f172a', '#0a7d5a')
        return (
            '<div style="margin:0 0 16px;">'
            # excluded (loss)
            f'<div style="{card}"><div style="{lbl}">Excluded Lines</div>'
            f'<div style="{val}">{_ind(len(self.excluded))}</div></div>'
            f'<div style="{card}"><div style="{lbl}">Excluded Qty</div>'
            f'<div style="{val}">{_ind(s["excluded_qty"])}</div></div>'
            f'<div style="{card}border-color:#fecaca;background:#fef2f2;">'
            f'<div style="{lbl}color:#b91c1c;">Excluded Value (loss)</div>'
            f'<div style="{red}">{_rupee(s["loss_excluded"])}</div></div>'
            # included (intimation)
            f'<div style="{card}border-color:#bbf7d0;background:#f0fdf4;">'
            f'<div style="{lbl}color:#0a7d5a;">Included Lines (intimation)</div>'
            f'<div style="{grn}">{_ind(len(self.included))}</div></div>'
            f'<div style="{card}border-color:#bbf7d0;background:#f0fdf4;">'
            f'<div style="{lbl}color:#0a7d5a;">Included Qty</div>'
            f'<div style="{grn}">{_ind(s["included_qty"])}</div></div>'
            '</div>'
            + self._by_mp_block())

    def _by_mp_block(self) -> str:
        """Per-marketplace: Lot Qty · Uploaded % · Excluded (qty/value) ·
        Included (qty/value)."""
        by_mp = self.summary.get('by_mp') or {}
        if not by_mp:
            return ''
        th = ('padding:7px 10px;text-align:left;font-size:10.5px;color:#334155;'
              'background:#eef2f7;white-space:nowrap;text-transform:uppercase;'
              'letter-spacing:.04em;border-bottom:2px solid #dbe3ec;')
        thr = th + 'text-align:right;'
        td = 'padding:6px 10px;font-size:12px;border-bottom:1px solid #eef1f5;'
        tdr = td + 'text-align:right;font-variant-numeric:tabular-nums;'
        rows = ''
        for mp, g in sorted(by_mp.items(), key=lambda x: -(x[1]['exc'] + x[1]['inc'])):
            pct = g.get('uploaded_pct', Decimal('0'))
            pct_col = '#0f9d6b' if pct >= Decimal('100') else '#b45309'
            rows += (
                f'<tr><td style="{td}"><b>{escape(mp)}</b></td>'
                f'<td style="{tdr}">{_ind(g.get("lot_qty", 0))}</td>'
                f'<td style="{tdr}color:{pct_col};font-weight:700;">{pct:.2f}%</td>'
                f'<td style="{tdr}color:#b91c1c;font-weight:700;">{_ind(g["exc_qty"])}</td>'
                f'<td style="{tdr}color:#b91c1c;">{_rupee(g["exc"])}</td>'
                f'<td style="{tdr}color:#0a7d5a;font-weight:700;">{_ind(g["inc_qty"])}</td>'
                f'<td style="{tdr}color:#0a7d5a;">{_rupee(g["inc"])}</td></tr>')
        return (
            '<div style="margin:0 0 16px;">'
            '<div style="font-size:11px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;'
            'color:#64748b;margin:0 0 6px;">By marketplace</div>'
            '<table style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb;">'
            f'<thead><tr><th style="{th}">Marketplace</th><th style="{thr}">Lot Qty</th>'
            f'<th style="{thr}">Uploaded %</th><th style="{thr}">Excl. Qty</th>'
            f'<th style="{thr}">Excl. Value</th><th style="{thr}">Incl. Qty</th>'
            f'<th style="{thr}">Incl. Value</th></tr></thead>'
            f'<tbody>{rows}</tbody></table></div>')

    def _by_sku_block(self) -> str:
        """SKU-wise rollup so the ecom team can see, per problem SKU, the TOTAL
        affected qty, on how many POs, and across how many marketplaces (named).
        Aggregated over every decided issue line in scope."""
        if not self.rows:
            return ''
        agg: dict = {}
        for r in self.rows:
            item = str(r.get('item_no') or '').strip()
            ean = str(r.get('ean') or '').strip()
            key = item or ean or str(r.get('description') or '')
            a = agg.get(key)
            if a is None:
                a = agg[key] = {'item': item, 'ean': ean,
                                'desc': str(r.get('description') or ''),
                                'qty': 0, 'pos': set(), 'mps': set()}
            a['qty'] += int(_num(r.get('qty')))
            if r.get('po'):
                a['pos'].add(str(r.get('po')))
            if r.get('marketplace'):
                a['mps'].add(str(r.get('marketplace')))
        th = ('padding:7px 10px;text-align:left;font-size:10.5px;color:#334155;'
              'background:#eef2f7;white-space:nowrap;text-transform:uppercase;'
              'letter-spacing:.04em;border-bottom:2px solid #dbe3ec;')
        thr = th + 'text-align:right;'
        td = 'padding:6px 10px;font-size:12px;border-bottom:1px solid #eef1f5;'
        tdr = td + 'text-align:right;font-variant-numeric:tabular-nums;'
        rows = ''
        for a in sorted(agg.values(), key=lambda x: -x['qty']):
            mps = sorted(a['mps'])
            code = ' · '.join([c for c in (a['item'], a['ean']) if c])
            sub = (f'<br><span style="font-family:monospace;font-size:10.5px;'
                   f'color:#94a3b8;">{escape(code)}</span>') if code else ''
            name = escape(a['desc'] or a['item'] or a['ean'] or '—')
            rows += (
                f'<tr><td style="{td}"><b>{name}</b>{sub}</td>'
                f'<td style="{tdr}color:#b91c1c;font-weight:700;">{_ind(a["qty"])}</td>'
                f'<td style="{tdr}">{len(a["pos"])}</td>'
                f'<td style="{td}">{len(mps)} '
                f'<span style="color:#64748b;">({escape(", ".join(mps))})</span></td></tr>')
        return (
            '<div style="margin:0 0 16px;">'
            '<div style="font-size:11px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;'
            'color:#64748b;margin:0 0 6px;">By SKU — issue summary</div>'
            '<table style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb;">'
            f'<thead><tr><th style="{th}">SKU</th><th style="{thr}">Affected Qty</th>'
            f'<th style="{thr}">PO Count</th><th style="{th}text-align:left;">Marketplaces (count · names)</th>'
            f'</tr></thead><tbody>{rows}</tbody></table></div>')

    def _remap_block(self) -> str:
        """EAN remaps shown separately: the marketplace's wrong/variant EAN →
        the correct one we shipped on, per SKU (qty · POs · marketplaces)."""
        if not self.remaps:
            return ''
        agg: dict = {}
        for r in self.remaps:
            recv = str(r.get('received_ean') or '').strip()
            good = str(r.get('ean') or '').strip()
            key = (recv, good)
            a = agg.get(key)
            if a is None:
                a = agg[key] = {'recv': recv, 'good': good,
                                'desc': str(r.get('description') or ''),
                                'item': str(r.get('item_no') or ''),
                                'qty': 0, 'pos': set(), 'mps': set()}
            a['qty'] += int(_num(r.get('qty')))
            if r.get('po'):
                a['pos'].add(str(r.get('po')))
            if r.get('marketplace'):
                a['mps'].add(str(r.get('marketplace')))
        th = ('padding:7px 10px;text-align:left;font-size:10.5px;color:#334155;'
              'background:#f3ecff;white-space:nowrap;text-transform:uppercase;'
              'letter-spacing:.04em;border-bottom:2px solid #e2d6ff;')
        thr = th + 'text-align:right;'
        td = 'padding:6px 10px;font-size:12px;border-bottom:1px solid #eef1f5;'
        tdr = td + 'text-align:right;font-variant-numeric:tabular-nums;'
        mono = 'font-family:monospace;font-size:11.5px;'
        rows = ''
        for a in sorted(agg.values(), key=lambda x: -x['qty']):
            mps = sorted(a['mps'])
            rows += (
                f'<tr><td style="{td}"><b>{escape(a["desc"] or a["item"] or "-")}</b></td>'
                f'<td style="{td}{mono}color:#b45309;">{escape(a["recv"])}</td>'
                f'<td style="{td}{mono}color:#7c3aed;font-weight:700;">&rarr; {escape(a["good"])}</td>'
                f'<td style="{tdr}font-weight:700;">{_ind(a["qty"])}</td>'
                f'<td style="{tdr}">{len(a["pos"])}</td>'
                f'<td style="{td}">{len(mps)} '
                f'<span style="color:#64748b;">({escape(", ".join(mps))})</span></td></tr>')
        return (
            '<div style="margin:0 0 16px;">'
            '<div style="font-size:11px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;'
            'color:#7c3aed;margin:0 0 6px;">EAN remaps &mdash; vendor sent the wrong barcode</div>'
            '<table style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb;">'
            f'<thead><tr><th style="{th}">SKU</th><th style="{th}">Received EAN</th>'
            f'<th style="{th}">Remapped to</th><th style="{thr}">Qty</th>'
            f'<th style="{thr}">PO Count</th>'
            f'<th style="{th}text-align:left;">Marketplaces (count &middot; names)</th>'
            f'</tr></thead><tbody>{rows}</tbody></table></div>')

    def _lines_table(self, rows) -> str:
        """One issue-lines table for a subset of rows (Excluded / Included)."""
        th = ('padding:8px 10px;text-align:left;font-size:11.5px;color:#334155;'
              'background:#eef2f7;white-space:nowrap;text-transform:uppercase;'
              'letter-spacing:.02em;border-bottom:2px solid #dbe3ec;')
        td = 'padding:7px 10px;font-size:12px;border-bottom:1px solid #eef0f4;'
        rows_html = []
        for i, r in enumerate(rows):
            act = (r.get('action') or '').upper()
            bg = '#ffffff' if i % 2 == 0 else '#f7f8fb'
            badge = (f'<span style="display:inline-block;padding:2px 8px;'
                     f'border-radius:10px;font-weight:700;font-size:11px;'
                     f'color:#fff;background:{_ACTION_COLOR.get(act, "#6b7280")};">'
                     f'{escape(_ACTION_LABEL.get(act, act or "—"))}</span>')
            rows_html.append(
                f'<tr style="background:{bg};">'
                f'<td style="{td}">{_fmt(r.get("marketplace"))}</td>'
                f'<td style="{td}font-family:monospace;">{_fmt(r.get("po"))}</td>'
                f'<td style="{td}font-family:monospace;">{_fmt(r.get("item_no"))}</td>'
                f'<td style="{td}font-family:monospace;">{_fmt(r.get("ean"))}</td>'
                f'<td style="{td}">{_fmt(r.get("description"))}</td>'
                f'<td style="{td}text-align:right;">{_fmt(r.get("qty"))}</td>'
                f'<td style="{td}text-align:right;">{_fmt(r.get("our_cp"))}</td>'
                f'<td style="{td}text-align:right;">{_fmt(r.get("vendor_cp"))}</td>'
                f'<td style="{td}text-align:right;">{_fmt(r.get("diff"))}</td>'
                f'<td style="{td}">{_fmt(r.get("status"))}</td>'
                f'<td style="{td}">{badge}</td>'
                f'<td style="{td}">{_fmt(r.get("remark"))}</td>'
                f'</tr>')
        body = ''.join(rows_html) or (
            f'<tr><td colspan="12" style="{td}text-align:center;color:#6b7280;">'
            f'No lines.</td></tr>')
        return (
            '<table style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb;">'
            '<thead><tr>'
            f'<th style="{th}">MP</th><th style="{th}">PO</th><th style="{th}">Item</th>'
            f'<th style="{th}">EAN</th><th style="{th}">Description</th>'
            f'<th style="{th}text-align:right;">Qty</th>'
            f'<th style="{th}text-align:right;">Our CP</th>'
            f'<th style="{th}text-align:right;">Their CP</th>'
            f'<th style="{th}text-align:right;">Diff</th>'
            f'<th style="{th}">Status</th><th style="{th}">Action</th><th style="{th}">Remark</th>'
            '</tr></thead>'
            f'<tbody>{body}</tbody></table>')

    def _section(self, title, colour, subtitle, rows) -> str:
        if not rows:
            return ''
        return (
            f'<div style="margin:20px 0 3px;font-size:13px;font-weight:800;'
            f'color:{colour};letter-spacing:.01em;">{title} ({len(rows)})</div>'
            f'<div style="margin:0 0 7px;font-size:11.5px;color:#64748b;">{subtitle}</div>'
            + self._lines_table(rows))

    def html(self) -> str:
        """Two sections: Excluded (the loss) then Included (intimation — kept
        despite the CP issue, listed for the record only)."""
        excl = self._section(
            'Excluded lines', '#b91c1c',
            'Dropped from the confirmed PO — the real loss.', self.excluded)
        incl = self._section(
            'Included despite the issue', '#0a7d5a',
            'Kept and processed into the order — NOT a loss. Listed as an '
            'intimation so the CP issue is on record.', self.included)
        empty = ('<p style="font-size:12px;color:#6b7280;">No decided issue lines '
                 'in the selected scope.</p>' if not self.rows else '')
        return f"""\
<div style="font-family:Segoe UI,Arial,sans-serif;color:#0f172a;max-width:1000px;">
  <h2 style="margin:0 0 4px;color:#1A237E;">RENÉE · Online B2B — Issue lines</h2>
  <p style="margin:0 0 2px;font-size:13px;">{self._summary_line()}</p>
  <p style="margin:0 0 16px;font-size:12px;color:#475569;">{self._scope_line()}</p>
  {self._note_block()}
  {self._summary_block()}
  {self._by_sku_block()}
  {self._remap_block()}
  {excl}
  {incl}
  {empty}
  <p style="margin:16px 0 0;font-size:11px;color:#94a3b8;">
    Auto-generated from the Order Management dashboard · {_dt.datetime.now():%d-%b-%Y %H:%M}.
  </p>
</div>"""


# ── Review-Later (CP issue) auto-email ──────────────────────────────────────
class ReviewLaterEmailReport(EmailReport):
    """Sent automatically when an operator **Saves a run for Review Later** —
    usually because of an unresolved **CP issue**. Lists the flagged lines (item,
    EAN, issue/status, affected qty, our-CP vs their-CP, diff) so the **ecom
    team** can see exactly what's blocked and resolve the CP, while the run sits
    parked. Recipients = the same stakeholders as the Issues email (config
    defaults, unless overridden). Built from the draft's parsed preview (the run
    isn't recorded yet), NOT the DB."""

    def __init__(self, marketplace: str, affected: list, note: str = '',
                 draft_at: str = '', to=None, cc=None, kpis: dict | None = None):
        self.marketplace = marketplace or 'All MPs'
        self.rows = [r for r in (affected or [])]
        self.note = (note or '').strip()
        self.draft_at = draft_at or ''
        self._to = _clean_emails(to) if to is not None else None
        self._cc = _clean_emails(cc) if cc is not None else None
        # Full run KPIs (same as the review page): pos/lines/qty/value/affected/
        # affected_qty/ok_qty_pct. Optional — falls back to affected-only cards.
        self.kpis = dict(kpis or {})
        self.mm = sum(1 for r in self.rows
                      if (r.get('status') or '').upper() == 'MISMATCH')
        self.nim = sum(1 for r in self.rows
                       if (r.get('status') or '').upper() == 'NOT_IN_MASTER')
        self.aff_qty = sum(int(_num(r.get('qty'))) for r in self.rows)

    def to(self):
        return self._to

    def cc(self):
        return self._cc

    def subject(self) -> str:
        d = _dt.date.today().strftime('%d-%b-%Y')
        return (f"⏸ Online B2B — CP issue parked (Review Later): "
                f"{len(self.rows)} flagged line(s) [{escape(self.marketplace)}] — {d}")

    def _card(self, label, value, tone='') -> str:
        base = ('display:inline-block;box-sizing:border-box;min-width:118px;'
                'padding:10px 13px;border-radius:10px;border:1px solid #e5e7eb;'
                'margin:0 7px 8px 0;vertical-align:top;background:#f8fafc;')
        lbl = ('font-size:10px;font-weight:800;letter-spacing:.04em;'
               'text-transform:uppercase;color:#64748b;')
        val = 'font-size:18px;font-weight:800;color:#0f172a;margin-top:2px;'
        if tone == 'red':
            base += 'border-color:#fecaca;background:#fef2f2;'
            lbl += 'color:#b91c1c;'
            val = val.replace('#0f172a', '#b91c1c')
        elif tone == 'green':
            base += 'border-color:#bbf7d0;background:#f0fdf4;'
            lbl += 'color:#0a7d5a;'
            val = val.replace('#0f172a', '#0a7d5a')
        return (f'<div style="{base}"><div style="{lbl}">{label}</div>'
                f'<div style="{val}">{value}</div></div>')

    def _cards(self) -> str:
        """The same KPI row as the review page (when the run summary is passed),
        so the recipient sees POs / Line items / Qty / Value / Affected lines /
        Affected qty / OK qty. Falls back to affected-only cards otherwise."""
        k = self.kpis
        if k:
            ok = _num(k.get('ok_qty_pct'))
            return ('<div style="margin:0 0 16px;">'
                    + self._card('POs', f"{int(_num(k.get('pos'))):,}")
                    + self._card('Line items', f"{int(_num(k.get('lines'))):,}")
                    + self._card('Qty', f"{int(_num(k.get('qty'))):,}")
                    + self._card('Value', _rupee(k.get('value')))
                    + self._card('Affected lines', f"{int(_num(k.get('affected'))):,}", 'red')
                    + self._card('Affected qty', f"{int(_num(k.get('affected_qty'))):,}", 'red')
                    + self._card('OK qty', f"{ok:.1f}%", 'green' if ok >= 98 else 'red')
                    + '</div>')
        # fallback: affected-only cards
        return ('<div style="margin:0 0 16px;">'
                + self._card('Flagged Lines', f"{len(self.rows):,}")
                + self._card('Affected Qty', f"{self.aff_qty:,}")
                + self._card('CP Mismatch', f"{self.mm:,}", 'red')
                + self._card('Not in Master', f"{self.nim:,}", 'red')
                + '</div>')

    def _table(self) -> str:
        th = ('padding:8px 10px;text-align:left;font-size:11.5px;color:#334155;'
              'background:#eef2f7;white-space:nowrap;text-transform:uppercase;'
              'letter-spacing:.02em;border-bottom:2px solid #dbe3ec;')
        td = 'padding:7px 10px;font-size:12px;border-bottom:1px solid #eef0f4;'
        body = []
        for i, r in enumerate(self.rows):
            bg = '#ffffff' if i % 2 == 0 else '#f7f8fb'
            st = (r.get('status') or '').upper()
            sc = '#b91c1c' if st == 'MISMATCH' else ('#b45309' if st == 'NOT_IN_MASTER' else '#334155')
            body.append(
                f'<tr style="background:{bg};">'
                f'<td style="{td}font-family:monospace;">{_fmt(r.get("po"))}</td>'
                f'<td style="{td}font-family:monospace;">{_fmt(r.get("item_no"))}</td>'
                f'<td style="{td}font-family:monospace;">{_fmt(r.get("ean"))}</td>'
                f'<td style="{td}">{_fmt(r.get("description"))}</td>'
                f'<td style="{td}text-align:right;">{_fmt(r.get("qty"))}</td>'
                f'<td style="{td}color:{sc};font-weight:700;">{_fmt(r.get("status"))}</td>'
                f'<td style="{td}text-align:right;">{_fmt(r.get("our_cp"))}</td>'
                f'<td style="{td}text-align:right;">{_fmt(r.get("vendor_cp"))}</td>'
                f'<td style="{td}text-align:right;">{_fmt(r.get("diff"))}</td>'
                f'<td style="{td}">{_fmt(r.get("remark"))}</td></tr>')
        rows = ''.join(body) or (
            f'<tr><td colspan="10" style="{td}text-align:center;color:#6b7280;">No lines.</td></tr>')
        return (
            '<table style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb;">'
            f'<thead><tr><th style="{th}">PO</th><th style="{th}">Item</th>'
            f'<th style="{th}">EAN</th><th style="{th}">Description</th>'
            f'<th style="{th}text-align:right;">Qty</th><th style="{th}">Issue</th>'
            f'<th style="{th}text-align:right;">Our CP</th><th style="{th}text-align:right;">Their CP</th>'
            f'<th style="{th}text-align:right;">Diff</th><th style="{th}">Remark</th></tr></thead>'
            f'<tbody>{rows}</tbody></table>')

    def html(self) -> str:
        note = ''
        if self.note:
            body = escape(self.note).replace('\n', '<br>')
            note = (
                '<div style="margin:0 0 16px;padding:12px 14px;border-radius:10px;'
                'background:#eef2ff;border:1px solid #c7d2fe;">'
                '<div style="font-size:10.5px;font-weight:800;letter-spacing:.04em;'
                'text-transform:uppercase;color:#3730a3;margin-bottom:5px;">Reason for parking</div>'
                f'<div style="font-size:13px;color:#1e293b;line-height:1.5;">{body}</div></div>')
        parked = f" · parked {escape(self.draft_at)}" if self.draft_at else ''
        return f"""\
<div style="font-family:Segoe UI,Arial,sans-serif;color:#0f172a;max-width:1000px;">
  <h2 style="margin:0 0 4px;color:#b45309;">⏸ RENÉE · Online B2B — CP issue parked for review</h2>
  <p style="margin:0 0 2px;font-size:13px;"><b>{escape(self.marketplace)}</b> — {len(self.rows)}
     flagged line(s) held pending a CP decision{parked}. Please review the price / SKU issue below.</p>
  <p style="margin:0 0 16px;font-size:12px;color:#475569;">
     This run is <b>saved for review later</b> — NOT recorded yet. It will be finalized once the CP is resolved.</p>
  {note}
  {self._cards()}
  {self._table()}
  <p style="margin:16px 0 0;font-size:11px;color:#94a3b8;">
    Auto-sent when the run was parked · {_dt.datetime.now():%d-%b-%Y %H:%M}.
  </p>
</div>"""
