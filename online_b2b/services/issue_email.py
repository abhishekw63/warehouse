"""
online_b2b.services.issue_email
===============================

The **Issues** email report — built on the reusable :mod:`mailer` skeleton.

Emails the currently-filtered issue lines to management: what was flagged and,
for each, what was **Excluded / Override / Kept**, respectively. The view layer
passes the same filter dict the Issues page / export use, so the email always
matches what the operator is looking at.
"""
from __future__ import annotations

import datetime as _dt
import re as _re
from decimal import Decimal, InvalidOperation
from html import escape

from . import order_db
from .mailer import EmailReport

# Operator action → human label shown in the email.
_ACTION_LABEL = {
    'EXCLUDE': 'Excluded',
    'OVERRIDE': 'Override',
    'KEEP': 'Kept (flagged)',
    '': '— (no action yet)',
}
_ACTION_COLOR = {
    'EXCLUDE': '#b91c1c', 'OVERRIDE': '#b45309',
    'KEEP': '#1d4ed8', '': '#6b7280',
}


def _fmt(v) -> str:
    if v is None or v == '':
        return '—'
    return escape(str(v))


_EMAIL_RE = _re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')


def _clean_emails(v) -> list:
    """Normalise a recipient value (list / comma-or-newline-separated string)
    to a de-duplicated list of syntactically-valid addresses. Invalid tokens
    are dropped silently here; the view validates + surfaces before sending."""
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


def _rupee(v) -> str:
    """₹ with Indian-style-friendly thousands separators, 2 dp."""
    d = _num(v)
    return f'₹{d:,.2f}'


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


class IssuesEmailReport(EmailReport):
    """Issue lines (current filter) → management email."""

    def __init__(self, filters: dict | None = None, note: str = '',
                 to=None, cc=None):
        self.filters = dict(filters or {})
        # Operator-authored intro note + recipient overrides (from the modal).
        # None / empty ⇒ fall back to the config defaults (see to()/cc()).
        self.note = (note or '').strip()
        self._to = _clean_emails(to) if to is not None else None
        self._cc = _clean_emails(cc) if cc is not None else None
        data = order_db.issues(limit=0, **self.filters)
        self.rows = data.get('rows', []) if data.get('ok') else []
        # Tally actions for the summary line.
        self.tally: dict = {}
        for r in self.rows:
            a = (r.get('action') or '').upper()
            self.tally[a] = self.tally.get(a, 0) + 1
        self.summary = self._compute_summary()

    # ── recipient overrides (from the modal) ────────────────────────────
    def to(self):
        return self._to        # None → config DEFAULT_RECIPIENT

    def cc(self):
        return self._cc        # None → config CC_RECIPIENTS

    # ── summary metrics (Total Qty / Total Value / Loss) ────────────────
    def _compute_summary(self) -> dict:
        """Totals over the SAME filtered lines the email lists.

        * total_qty   = Σ qty
        * total_value = Σ (qty × our expected per-unit rate on the line basis)
        * loss (value at risk) = value tied up in lines that are NOT clean,
          split into three honest buckets:
            - mismatch    : Σ qty × |diff| over MISMATCH lines (the actual per-
              unit price gap × qty — the rupee exposure of the price mismatch)
            - not_in_master: Σ (qty × rate) over NOT_IN_MASTER lines (whole line
              value is at risk — the SKU can't be verified against our master)
            - excluded    : Σ (qty × rate) over lines the operator EXCLUDEd
              (dropped from the confirmed PO)
          Buckets are disjoint by precedence excluded > not_in_master >
          mismatch so no line is counted twice."""
        tot_qty = 0
        tot_val = Decimal('0')
        loss_mm = Decimal('0')
        loss_nim = Decimal('0')
        loss_exc = Decimal('0')
        for r in self.rows:
            qty = int(_num(r.get('qty')))
            rate = _unit_rate(r)
            line_val = rate * qty
            tot_qty += qty
            tot_val += line_val
            status = (r.get('status') or '').upper()
            action = (r.get('action') or '').upper()
            if action == 'EXCLUDE':
                loss_exc += line_val
            elif status == 'NOT_IN_MASTER':
                loss_nim += line_val
            elif status == 'MISMATCH':
                loss_mm += abs(_num(r.get('diff'))) * qty
        loss_total = loss_mm + loss_nim + loss_exc
        return {
            'total_qty': tot_qty,
            'total_value': tot_val,
            'loss_total': loss_total,
            'loss_mismatch': loss_mm,
            'loss_not_in_master': loss_nim,
            'loss_excluded': loss_exc,
        }

    # ── header ──────────────────────────────────────────────────────────
    def subject(self) -> str:
        res = self.filters.get('resolution', 'pending') or 'pending'
        mp = self.filters.get('marketplace') or 'All MPs'
        d = _dt.date.today().strftime('%d-%b-%Y')
        return (f"Online B2B — Issue lines ({res}): {len(self.rows)} "
                f"[{mp}] — {d}")

    # ── body ────────────────────────────────────────────────────────────
    def _summary_line(self) -> str:
        if not self.rows:
            return 'No issue lines in the selected filter.'
        parts = []
        for code in ('EXCLUDE', 'OVERRIDE', 'KEEP', ''):
            n = self.tally.get(code, 0)
            if n:
                parts.append(f"{n} {_ACTION_LABEL[code].lower()}")
        return f"{len(self.rows)} flagged line(s) — " + ', '.join(parts) + '.'

    def _scope_line(self) -> str:
        f = self.filters
        bits = [f"Resolution: <b>{escape(f.get('resolution', 'pending') or 'pending')}</b>"]
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
        """Operator's intro note, rendered as a clearly-marked section above the
        summary. Empty note ⇒ nothing is rendered."""
        if not self.note:
            return ''
        # Preserve the operator's line breaks; escape everything (untrusted).
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
        """Compact Total Qty / Total Value / Loss card above the lines table."""
        s = self.summary
        card = ('display:inline-block;box-sizing:border-box;min-width:150px;'
                'padding:10px 14px;border-radius:10px;border:1px solid #e5e7eb;'
                'margin:0 8px 8px 0;vertical-align:top;background:#f8fafc;')
        lbl = ('font-size:10px;font-weight:800;letter-spacing:.05em;'
               'text-transform:uppercase;color:#64748b;')
        val = 'font-size:19px;font-weight:800;color:#0f172a;margin-top:2px;'
        loss_val = val.replace('#0f172a', '#b91c1c')
        breakdown = (
            f'mismatch {_rupee(s["loss_mismatch"])} &middot; '
            f'not-in-master {_rupee(s["loss_not_in_master"])} &middot; '
            f'excluded {_rupee(s["loss_excluded"])}')
        return (
            '<div style="margin:0 0 16px;">'
            f'<div style="{card}"><div style="{lbl}">Total Qty</div>'
            f'<div style="{val}">{s["total_qty"]:,}</div></div>'
            f'<div style="{card}"><div style="{lbl}">Total Value</div>'
            f'<div style="{val}">{_rupee(s["total_value"])}</div></div>'
            f'<div style="{card}border-color:#fecaca;background:#fef2f2;">'
            f'<div style="{lbl}color:#b91c1c;">Loss (value at risk)</div>'
            f'<div style="{loss_val}">{_rupee(s["loss_total"])}</div>'
            f'<div style="font-size:11px;color:#7f1d1d;margin-top:3px;">'
            f'{breakdown}</div></div>'
            '</div>')

    def html(self) -> str:
        th = ('padding:8px 10px;text-align:left;font-size:12px;color:#fff;'
              'background:#1A237E;white-space:nowrap;')
        td = 'padding:7px 10px;font-size:12px;border-bottom:1px solid #eef0f4;'
        rows_html = []
        for i, r in enumerate(self.rows):
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
            f'No issue lines.</td></tr>')

        return f"""\
<div style="font-family:Segoe UI,Arial,sans-serif;color:#0f172a;max-width:1000px;">
  <h2 style="margin:0 0 4px;color:#1A237E;">RENÉE · Online B2B — Issue lines</h2>
  <p style="margin:0 0 2px;font-size:13px;">{self._summary_line()}</p>
  <p style="margin:0 0 16px;font-size:12px;color:#475569;">{self._scope_line()}</p>
  {self._note_block()}
  {self._summary_block()}
  <table style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb;">
    <thead><tr>
      <th style="{th}">MP</th><th style="{th}">PO</th><th style="{th}">Item</th>
      <th style="{th}">EAN</th><th style="{th}">Description</th>
      <th style="{th}text-align:right;">Qty</th>
      <th style="{th}text-align:right;">Our CP</th>
      <th style="{th}text-align:right;">Their CP</th>
      <th style="{th}text-align:right;">Diff</th>
      <th style="{th}">Status</th><th style="{th}">Action</th><th style="{th}">Remark</th>
    </tr></thead>
    <tbody>{body}</tbody>
  </table>
  <p style="margin:16px 0 0;font-size:11px;color:#94a3b8;">
    Auto-generated from the Order Management dashboard · {_dt.datetime.now():%d-%b-%Y %H:%M}.
  </p>
</div>"""
