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
from html import escape

from . import order_db
from .mailer import EmailReport

# Operator action → human label shown in the email.
_ACTION_LABEL = {
    "EXCLUDE": "Excluded",
    "OVERRIDE": "Override",
    "KEEP": "Kept (flagged)",
    "": "— (no action yet)",
}
_ACTION_COLOR = {
    "EXCLUDE": "#b91c1c",
    "OVERRIDE": "#b45309",
    "KEEP": "#1d4ed8",
    "": "#6b7280",
}


def _fmt(v) -> str:
    if v is None or v == "":
        return "—"
    return escape(str(v))


class IssuesEmailReport(EmailReport):
    """Issue lines (current filter) → management email."""

    def __init__(self, filters: dict | None = None):
        self.filters = dict(filters or {})
        data = order_db.issues(limit=0, **self.filters)
        self.rows = data.get("rows", []) if data.get("ok") else []
        # Tally actions for the summary line.
        self.tally: dict = {}
        for r in self.rows:
            a = (r.get("action") or "").upper()
            self.tally[a] = self.tally.get(a, 0) + 1

    # ── header ──────────────────────────────────────────────────────────
    def subject(self) -> str:
        res = self.filters.get("resolution", "pending") or "pending"
        mp = self.filters.get("marketplace") or "All MPs"
        d = _dt.date.today().strftime("%d-%b-%Y")
        return f"Online B2B — Issue lines ({res}): {len(self.rows)} [{mp}] — {d}"

    # ── body ────────────────────────────────────────────────────────────
    def _summary_line(self) -> str:
        if not self.rows:
            return "No issue lines in the selected filter."
        parts = []
        for code in ("EXCLUDE", "OVERRIDE", "KEEP", ""):
            n = self.tally.get(code, 0)
            if n:
                parts.append(f"{n} {_ACTION_LABEL[code].lower()}")
        return f"{len(self.rows)} flagged line(s) — " + ", ".join(parts) + "."

    def _scope_line(self) -> str:
        f = self.filters
        bits = [f"Resolution: <b>{escape(f.get('resolution', 'pending') or 'pending')}</b>"]
        if f.get("marketplace"):
            bits.append(f"Marketplace: <b>{escape(f['marketplace'])}</b>")
        if f.get("status"):
            bits.append(f"Status: <b>{escape(f['status'])}</b>")
        if f.get("date_from") or f.get("date_to"):
            bits.append(
                "Upload date: <b>"
                f"{escape(f.get('date_from') or '…')} → "
                f"{escape(f.get('date_to') or '…')}</b>"
            )
        if f.get("q"):
            bits.append(f"Search: <b>{escape(f['q'])}</b>")
        return " &nbsp;·&nbsp; ".join(bits)

    def html(self) -> str:
        th = (
            "padding:8px 10px;text-align:left;font-size:12px;color:#fff;"
            "background:#1A237E;white-space:nowrap;"
        )
        td = "padding:7px 10px;font-size:12px;border-bottom:1px solid #eef0f4;"
        rows_html = []
        for i, r in enumerate(self.rows):
            act = (r.get("action") or "").upper()
            bg = "#ffffff" if i % 2 == 0 else "#f7f8fb"
            badge = (
                f'<span style="display:inline-block;padding:2px 8px;'
                f"border-radius:10px;font-weight:700;font-size:11px;"
                f'color:#fff;background:{_ACTION_COLOR.get(act, "#6b7280")};">'
                f"{escape(_ACTION_LABEL.get(act, act or '—'))}</span>"
            )
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
                f"</tr>"
            )
        body = "".join(rows_html) or (
            f'<tr><td colspan="12" style="{td}text-align:center;color:#6b7280;">'
            f"No issue lines.</td></tr>"
        )

        return f"""\
<div style="font-family:Segoe UI,Arial,sans-serif;color:#0f172a;max-width:1000px;">
  <h2 style="margin:0 0 4px;color:#1A237E;">RENÉE · Online B2B — Issue lines</h2>
  <p style="margin:0 0 2px;font-size:13px;">{self._summary_line()}</p>
  <p style="margin:0 0 16px;font-size:12px;color:#475569;">{self._scope_line()}</p>
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
