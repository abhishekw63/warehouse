"""
online_b2b.services.daily_email
===============================

The **Daily Activity** email report — built on the reusable :mod:`mailer`
skeleton, exactly like :mod:`issue_email`.

Purpose: let an operator send their senior a time-stamped digest of the day —
which channel steps were **worked** (and when / by whom), what is **on hold**
(since when), what has **no PO**, and what is still **pending**, plus the
personal **My Tasks** list (done today + still open). The aim is simply to make
the senior aware of where the day stands.
"""
from __future__ import annotations

import datetime as _dt
from html import escape

from .issue_email import _clean_emails   # reuse the shared recipient cleaner
from .mailer import EmailReport

# state → (label, colour) for the status pill
_STATE = {
    'done':    ('✓ Completed',  '#0f9d6b'),
    'partial': ('◐ In progress', '#2563eb'),
    'hold':    ('⏸ On hold',    '#b45309'),
    'nopo':    ('— No PO',      '#64748b'),
    'todo':    ('○ Pending',    '#94a3b8'),
}


def _fmt(v) -> str:
    return escape(str(v)) if v not in (None, '') else '—'


class DailyTasksEmailReport(EmailReport):
    """One day's activity grid + My Tasks → a senior-awareness email."""

    def __init__(self, day=None, note: str = '', to=None, cc=None):
        from . import daily_checklist as dc
        self.note = (note or '').strip()
        self._to = _clean_emails(to) if to is not None else None
        self._cc = _clean_emails(cc) if cc is not None else None
        self.data = dc.get_day(day)
        self.adhoc = dc.adhoc_list()
        # Flatten to leaves (parents' children + standalone channels), tagged
        # with their segment, so the email reads as one ordered list.
        self.leaves: list = []
        for seg in self.data.get('segments', []):
            for c in seg.get('channels', []):
                if c.get('is_parent'):
                    for kid in c.get('children', []):
                        self.leaves.append((seg['segment'], c['display'], kid))
                else:
                    self.leaves.append((seg['segment'], '', c))
        # Anything with real activity today (worked / hold / no-PO / done).
        self.active = [t for t in self.leaves
                       if t[2]['state'] in ('done', 'partial', 'hold', 'nopo')]

    # ── recipients (typed in the modal; None → config defaults) ─────────
    def to(self):
        return self._to

    def cc(self):
        return self._cc

    # ── header ──────────────────────────────────────────────────────────
    def subject(self) -> str:
        d = self.data
        try:
            nice = _dt.date.fromisoformat(d['day']).strftime('%d-%b-%Y')
        except (ValueError, KeyError):
            nice = d.get('day', '')
        return (f"Daily Activity — {nice}: {d.get('done_channels', 0)} done · "
                f"{d.get('held_channels', 0)} on hold · "
                f"{d.get('pending_channels', 0)} pending")

    # ── body pieces ─────────────────────────────────────────────────────
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

    def _kpi_block(self) -> str:
        d = self.data
        lbl = ('font-size:9.5px;font-weight:800;letter-spacing:.06em;'
               'text-transform:uppercase;color:#7c8698;')
        val = 'font-size:24px;font-weight:800;line-height:1;margin-top:5px;'

        def c(label, value, colour, accent, icon, delay):
            card = ('display:inline-block;box-sizing:border-box;min-width:112px;'
                    'padding:12px 14px 13px;border-radius:14px;border:1px solid #eceff4;'
                    'margin:0 9px 9px 0;vertical-align:top;background:#ffffff;'
                    f'border-top:3px solid {accent};'
                    'box-shadow:0 2px 6px rgba(20,30,60,.05);')
            return (f'<div class="kpi" style="{card}animation-delay:{delay}s;">'
                    f'<div style="{lbl}">{icon}&nbsp;{label}</div>'
                    f'<div style="{val}color:{colour};">{value}</div></div>')
        return (
            '<div style="margin:0 0 18px;">'
            + c('Handled', f"{d.get('handled_channels', 0)}/{d.get('total_channels', 0)}",
                '#0f172a', '#6366f1', '&#9673;', 0.02)
            + c('Completed', d.get('done_channels', 0), '#0f9d6b', '#10b981', '&#10003;', 0.08)
            + c('On hold', d.get('held_channels', 0), '#b45309', '#f59e0b', '&#9208;', 0.14)
            + c('No PO', d.get('nopo_channels', 0), '#64748b', '#94a3b8', '&#8212;', 0.20)
            + c('Pending', d.get('pending_channels', 0), '#b91c1c', '#ef4444', '&#9675;', 0.26)
            + '</div>')

    # soft (tinted) pill palette per state — bg + text, email-safe fixed hex
    _PILL = {
        'done':    ('#e7f6ef', '#0f9d6b'),
        'partial': ('#e8f0fe', '#2563eb'),
        'hold':    ('#fdf1e3', '#b45309'),
        'nopo':    ('#eef1f5', '#64748b'),
        'todo':    ('#f1f5f9', '#94a3b8'),
    }

    def _step_chips(self, lf) -> str:
        """Worked steps as compact wrapping chips (label · time) + a slim
        progress bar — replaces the old 5-line stack that read as congested."""
        done_steps = [s for s in lf['steps'] if s['checked']]
        total = lf['total_steps'] or 1
        if not done_steps:
            return '<span style="color:#b0b7c3;font-size:12px;">—</span>'
        chips = ''
        for s in done_steps:
            t = escape(s['at']) or '—'
            chips += (
                '<span style="display:inline-block;margin:0 5px 5px 0;padding:3px 9px;'
                'border-radius:20px;background:#f0f7f2;border:1px solid #cfe9db;'
                'font-size:11px;color:#0b7a54;white-space:nowrap;">'
                f'<span style="color:#0f9d6b;font-weight:700;">✓</span> '
                f'{escape(s["label"])} '
                f'<span style="color:#8aa398;">· {t}</span></span>')
        pct = round(len(done_steps) * 100 / total)
        # Table-based bar — an empty inline-block span collapses in Gmail, so the
        # fill is a real <td> with content (&nbsp; + line-height:0).
        bar = (
            '<table role="presentation" cellpadding="0" cellspacing="0" '
            'style="border-collapse:collapse;margin-top:5px;"><tr>'
            '<td style="padding:0;"><table role="presentation" cellpadding="0" '
            'cellspacing="0" style="border-collapse:collapse;width:96px;height:6px;'
            'background:#eef2f7;border-radius:6px;"><tr>'
            f'<td class="pfill" style="width:{pct}%;height:6px;background:#10b981;'
            f'border-radius:6px;font-size:0;line-height:0;">&nbsp;</td>'
            f'<td style="font-size:0;line-height:0;">&nbsp;</td></tr></table></td>'
            f'<td style="padding-left:8px;font-size:10.5px;color:#94a3b8;'
            f'white-space:nowrap;">{len(done_steps)}/{total} steps</td>'
            '</tr></table>')
        return chips + bar

    def _channels_table(self) -> str:
        th = ('padding:11px 14px;text-align:left;font-size:10.5px;color:#64748b;'
              'white-space:nowrap;text-transform:uppercase;letter-spacing:.05em;'
              'font-weight:800;border-bottom:1px solid #e5e9f0;')
        rows_html, cur_seg, i = [], None, 0
        for seg, parent, lf in self.active:
            if seg != cur_seg:
                cur_seg = seg
                rows_html.append(
                    f'<tr><td colspan="4" style="padding:10px 14px 6px;'
                    f'font-size:10.5px;font-weight:800;letter-spacing:.06em;'
                    f'text-transform:uppercase;color:#1A237E;background:#f7f9fc;'
                    f'border-bottom:1px solid #e5e9f0;">{escape(seg)}</td></tr>')
                i = 0
            i += 1
            bg = '#ffffff' if i % 2 else '#fafbfd'
            td = (f'padding:12px 14px;font-size:12.5px;vertical-align:top;'
                  f'background:{bg};border-bottom:1px solid #f0f2f6;')
            label, _ = _STATE.get(lf['state'], ('—', '#64748b'))
            pbg, ptx = self._PILL.get(lf['state'], ('#f1f5f9', '#64748b'))
            crumb = (f'<div style="font-size:10px;color:#a0a7b4;margin-bottom:1px;">'
                     f'{escape(parent)}</div>' if parent else '')
            name = (f'{crumb}<span style="font-weight:700;color:#0f172a;font-size:13px;">'
                    f'{escape(lf["display"])}</span>')
            pill = (f'<span style="display:inline-block;padding:3px 11px;'
                    f'border-radius:20px;font-weight:700;font-size:11px;'
                    f'color:{ptx};background:{pbg};white-space:nowrap;">{label}</span>')
            if lf['state'] == 'hold':
                note = (f'<span style="color:#b45309;">&#9208; on hold since '
                        f'<b>{escape(lf["hold_at"]) or "—"}</b>'
                        f'{(" · " + escape(lf["hold_by"])) if lf["hold_by"] else ""}</span>')
            elif lf['state'] == 'nopo':
                note = (f'<span style="color:#94a3b8;">no PO today'
                        f'{(" · " + escape(lf["no_po_at"])) if lf["no_po_at"] else ""}</span>')
            else:
                note = '<span style="color:#cbd2dc;">—</span>'
            rows_html.append(
                f'<tr><td style="{td}">{name}</td>'
                f'<td style="{td}white-space:nowrap;">{pill}</td>'
                f'<td style="{td}line-height:1.7;">{self._step_chips(lf)}</td>'
                f'<td style="{td}font-size:11.5px;">{note}</td></tr>')
        if not rows_html:
            rows_html.append(
                '<tr><td colspan="4" style="padding:16px;text-align:center;'
                'color:#94a3b8;font-size:12.5px;">No channel activity recorded yet '
                'for this day.</td></tr>')
        return (
            '<div style="font-size:12px;font-weight:800;letter-spacing:.05em;'
            'text-transform:uppercase;color:#334155;margin:8px 0 8px;">'
            'Channel activity</div>'
            '<table style="border-collapse:separate;border-spacing:0;width:100%;'
            'border:1px solid #e5e9f0;border-radius:12px;overflow:hidden;">'
            f'<thead><tr><th style="{th}">Channel</th><th style="{th}">Status</th>'
            f'<th style="{th}">Steps worked</th><th style="{th}">Note</th>'
            '</tr></thead>'
            f'<tbody>{"".join(rows_html)}</tbody></table>')

    def _pending_block(self) -> str:
        """Compact list of channels not started (state todo) so the picture is
        complete without cluttering the main table."""
        todo = [lf['display'] for _seg, _p, lf in self.leaves
                if lf['state'] == 'todo']
        if not todo:
            return ''
        return (
            '<div style="margin:12px 0 16px;font-size:12px;color:#64748b;">'
            f'<b>Not started yet ({len(todo)}):</b> {escape(", ".join(todo))}</div>')

    def _adhoc_block(self) -> str:
        a = self.adhoc
        done, open_ = a.get('done_today', []), a.get('open', [])
        if not done and not open_:
            return ''
        li = 'margin:3px 0;font-size:12.5px;'
        parts = ['<div style="font-size:12px;font-weight:800;letter-spacing:.05em;'
                 'text-transform:uppercase;color:#334155;margin:18px 0 6px;">'
                 'My Tasks (ad-hoc)</div>']
        if done:
            items = ''.join(
                f'<li style="{li}color:#0f172a;"><b>{escape(t["title"])}</b> '
                f'<span style="color:#64748b;">· done {escape(t["done_at"]) or "—"}'
                f'{(" · " + escape(t["done_by"])) if t["done_by"] else ""}</span></li>'
                for t in done)
            parts.append(f'<div style="font-size:11.5px;color:#0f9d6b;font-weight:700;'
                         f'margin:4px 0 2px;">✓ Completed today ({len(done)})</div>'
                         f'<ul style="margin:0 0 8px 18px;padding:0;">{items}</ul>')
        if open_:
            items = ''.join(
                f'<li style="{li}color:{"#b91c1c" if t.get("overdue") else "#0f172a"};">'
                f'<b>{escape(t["title"])}</b>'
                f'{(" <span style=\'color:#b91c1c;\'>· 📅 " + escape(t["due"]) + (" · overdue" if t.get("overdue") else "") + "</span>") if t.get("due") else ""}'
                f' <span style="color:#94a3b8;">· added {escape(t["added"])}'
                f'{(" · " + str(t["age"]) + "d ago") if t.get("age") else ""}</span></li>'
                for t in open_)
            parts.append(f'<div style="font-size:11.5px;color:#b45309;font-weight:700;'
                         f'margin:4px 0 2px;">◷ Still open ({len(open_)})</div>'
                         f'<ul style="margin:0 0 8px 18px;padding:0;">{items}</ul>')
        return ''.join(parts)

    def _hero(self) -> str:
        """Branded gradient header with the date + an animated overall-progress
        bar. Gradient/rounded corners degrade gracefully in older clients."""
        d = self.data
        try:
            nice = _dt.date.fromisoformat(d['day']).strftime('%A, %d %b %Y')
        except (ValueError, KeyError):
            nice = d.get('day', '')
        pct = d.get('overall_pct', 0)
        return f"""
  <div style="background:#1A237E;background:linear-gradient(120deg,#1A237E 0%,#3949AB 55%,#5C6BC0 100%);
              border-radius:18px;padding:22px 24px;color:#ffffff;margin:0 0 18px;
              box-shadow:0 10px 26px rgba(26,35,126,.28);">
    <div style="font-size:11px;font-weight:800;letter-spacing:.18em;text-transform:uppercase;
                color:#c5cae9;">RENÉE &middot; Daily Activity</div>
    <div style="font-size:21px;font-weight:800;margin:4px 0 2px;">{escape(nice)}</div>
    <div style="font-size:12.5px;color:#dfe3f7;margin-bottom:14px;">
      Every step is time-stamped — here's exactly what moved and what's parked.</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
      <tr>
        <td style="width:60%;padding-right:16px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
                 style="border-collapse:collapse;background:rgba(255,255,255,.22);border-radius:10px;">
            <tr><td style="padding:0;">
              <table role="presentation" width="{pct}%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <tr><td class="pfill" style="height:12px;background:#00e5a0;border-radius:10px;
                        font-size:0;line-height:0;">&nbsp;</td></tr>
              </table>
            </td></tr>
          </table>
        </td>
        <td style="white-space:nowrap;vertical-align:middle;">
          <span style="font-size:22px;font-weight:800;">{pct}%</span>
          <span style="font-size:12px;color:#dfe3f7;"> handled &middot;
            {d.get('handled_channels', 0)}/{d.get('total_channels', 0)} channels</span>
        </td>
      </tr>
    </table>
  </div>"""

    _STYLE = """
  <style>
    @keyframes daFade { from{opacity:0;transform:translateY(10px)} to{opacity:1;transform:none} }
    @keyframes daPop  { 0%{opacity:0;transform:scale(.88)} 100%{opacity:1;transform:scale(1)} }
    @keyframes daBar  { from{opacity:.2} to{opacity:1} }
    .da-sec { animation:daFade .55s cubic-bezier(.2,.8,.2,1) both; }
    .da-sec:nth-of-type(2){animation-delay:.06s} .da-sec:nth-of-type(3){animation-delay:.12s}
    .da-sec:nth-of-type(4){animation-delay:.18s} .da-sec:nth-of-type(5){animation-delay:.24s}
    .kpi   { animation:daPop .5s cubic-bezier(.2,.9,.25,1) both; }
    .pfill { animation:daBar .9s ease both; }
    .da-tbl tbody tr { transition:background .15s ease; }
    .da-tbl tbody tr:hover td { background:#f4f7ff !important; }
    @media (prefers-reduced-motion: reduce){ *{animation:none!important} }
  </style>"""

    def html(self) -> str:
        return f"""\
<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
{self._STYLE}
</head>
<body style="margin:0;padding:0;background:#eef1f7;">
<div style="font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#0f172a;
            max-width:1040px;margin:0 auto;padding:22px;background:#eef1f7;">
  {self._hero()}
  <div style="background:#ffffff;border-radius:18px;padding:22px 24px;
              box-shadow:0 6px 20px rgba(20,30,60,.06);">
    <div class="da-sec">{self._note_block()}</div>
    <div class="da-sec">{self._kpi_block()}</div>
    <div class="da-sec da-tbl">{self._channels_table()}</div>
    <div class="da-sec">{self._pending_block()}</div>
    <div class="da-sec">{self._adhoc_block()}</div>
    <p style="margin:18px 0 0;font-size:11px;color:#a0a7b4;border-top:1px solid #eef0f4;padding-top:12px;">
      Auto-generated from the Order Management dashboard · {_dt.datetime.now():%d-%b-%Y %H:%M}.
    </p>
  </div>
</div>
</body></html>"""
