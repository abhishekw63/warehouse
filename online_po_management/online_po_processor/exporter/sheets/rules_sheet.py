"""
exporter.sheets.rules_sheet
===========================

Writes the **Rules & Exceptions** sheet (v2.4.4) — a one-row-per-marketplace
overview of the pricing RULE each marketplace follows plus its EXCEPTIONS,
rendered on EVERY output so the operator has the whole pricing picture at a
glance:

    Blinkit   | 70% landing                  | EPISENSE …: deal MRP 899 @ 76%
    Flipkart… | 77% landing                  | —
    Swiggy    | 80% cost                     | 5 deal SKU(s)
    Nykaa     | per category: Perfume 69% …  | —
    Reliance  | GST-based: keep 1-31%×(1+GST)| —

The marketplace being processed is highlighted (green/bold). Rules come from
``MARKETPLACE_CONFIGS`` (``margin_rules`` / ``gst_margin_discount`` / straight
``default_margin`` × ``compare_basis``); exceptions from
``result.exception_registry`` (the full Master-Exceptions + Swiggy-deal
registry). This complements the detailed per-EAN **Exceptions** sheet with a
per-marketplace summary.
"""
from __future__ import annotations

from online_po_processor.config.marketplaces import MARKETPLACE_CONFIGS
from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    BOLD_DATA_FONT, INFO_ITALIC_FONT, OK_FILL,
    auto_width, data_cell, hdr_cell,
)
from online_po_processor.exporter.sheets.tracker_sheet import _MARKETPLACE_DISPLAY

_HEADERS = ['Marketplace', 'Pricing Rule', 'Exceptions']


def _norm(s) -> str:
    return ''.join(str(s or '').split()).lower()


def _rule_for_config(cfg: dict) -> str:
    """Human pricing rule for a marketplace config — mirrors
    ``summary_sheet.pricing_rule_str`` but driven by the config dict so it
    works for marketplaces other than the one being processed."""
    mr = cfg.get('margin_rules')
    if mr:
        parts = []
        for rule in mr.get('rules', []):
            kp = rule.get('keep_pct')
            lbl = rule.get('label', 'rule')
            parts.append(f"{lbl} {kp}%" if kp is not None else str(lbl))
        dk = mr.get('default_keep_pct')
        if dk is not None:
            parts.append(f"{mr.get('default_label', 'Default')} {dk}%")
        return 'per category: ' + ' / '.join(parts)

    gmd = cfg.get('gst_margin_discount')
    if gmd is not None:
        pct = round(gmd * 100, 2)
        slabs = ' / '.join(
            f"{round((1 - gmd * (1 + g)) * 100, 2)}%@{int(g * 100)}%GST"
            for g in (0.0, 0.05, 0.18))
        return f"GST-based: keep 1-{pct}%x(1+GST) = {slabs}"

    dm = cfg.get('default_margin')
    basis = cfg.get('compare_basis', 'cost')
    return f"{dm}% {basis}" if dm is not None else '—'


def _exceptions_for(registry: list, mp_key: str, mp_display: str) -> str:
    """Summarise the registry exceptions that apply to one marketplace —
    individual non-deal exceptions spelled out, Swiggy deal SKUs collapsed to
    a count. '' when none."""
    want = {_norm(mp_key), _norm(mp_display)}
    want.discard('')
    rows = [e for e in registry if _norm(e.get('marketplace')) in want]
    if not rows:
        return ''
    deals = [e for e in rows if 'swiggy_deal' in (e.get('kinds') or [])]
    others = [e for e in rows if 'swiggy_deal' not in (e.get('kinds') or [])]
    parts = []
    for e in others:
        # Prefer the operator's own Note (carries the product name + %, e.g.
        # 'EPISENSE … 24% discount'); fall back to source EAN + derived effect.
        note = e.get('note', '')
        parts.append(note if note
                     else f"{e.get('source_code', '')}: {e.get('effect', '')}"
                     .strip(' :'))
    if deals:
        parts.append(f"{len(deals)} deal SKU(s)")
    return '  ;  '.join(p for p in parts if p)


def write(wb, result: ProcessingResult) -> None:
    """Append the 'Rules & Exceptions' sheet — one row per marketplace, the
    current one highlighted."""
    registry = getattr(result, 'exception_registry', None) or []
    current = result.marketplace or ''

    ws = wb.create_sheet('Rules & Exceptions')
    for c, h in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, c, h)

    # Current marketplace first (the operator's focus), then the rest in
    # config order.
    keys = list(MARKETPLACE_CONFIGS.keys())
    keys.sort(key=lambda k: _norm(k) != _norm(current))

    own_n = 0
    r = 2
    for key in keys:
        cfg = MARKETPLACE_CONFIGS[key]
        disp = _MARKETPLACE_DISPLAY.get(key, key)
        is_own = _norm(key) == _norm(current)
        own_n += 1 if is_own else 0

        data_cell(ws, r, 1, disp, align='center')
        data_cell(ws, r, 2, _rule_for_config(cfg), align='left')
        data_cell(ws, r, 3, _exceptions_for(registry, key, disp) or '—',
                  align='left')

        if is_own:
            for c in range(1, len(_HEADERS) + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = OK_FILL
                cell.font = BOLD_DATA_FONT
        r += 1

    r += 1
    disp_cur = _MARKETPLACE_DISPLAY.get(current, current) or current
    ws.cell(
        row=r, column=1,
        value=(f"ℹ Pricing rule + exceptions for every marketplace; "
               f"highlighted (green) = {disp_cur} (this run). Rules from the "
               f"marketplace config; exceptions from 'Master Exceptions.xlsx' "
               f"+ the master's Swiggy deal sheet."),
    ).font = INFO_ITALIC_FONT
    auto_width(ws)
