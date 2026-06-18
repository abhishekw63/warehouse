"""
exporter.sheets.exceptions_sheet
================================

Writes the **Exceptions** sheet — the FULL cross-marketplace exception
registry from the central ``Master Exceptions.xlsx`` file, rendered on
EVERY marketplace's output (v2.4.3).

Why the whole list, every time
------------------------------
The operator wants a single, always-present view of *all* exceptions across
*all* marketplaces each time any output is generated — so nothing silently
drifts out of mind. The rows that belong to **the marketplace being
processed** are HIGHLIGHTED (light-green, bold) so they stand out at a
glance — e.g. Blinkit's EPISENSE price override is highlighted in Blinkit's
output, Myntra's Goddess vendor-CP in Myntra's output — while the other
marketplaces' exceptions are still listed (un-highlighted) for awareness.

A blank ``Marketplace`` in the file means the exception applies to EVERY
channel, so those rows are highlighted in every output too.

Columns mirror the source file (Marketplace / Source Code / Maps To /
Override MRP / Override Margin % / Use Vendor CP / Note) plus a derived
``Effect`` summary and an **Applied (this run)** column that cross-references
``result.exceptions_applied`` — so the operator sees both the registry AND
which overrides actually fired on this batch.

No-op only when there's no registry at all (the exceptions file is absent),
so a setup without the file behaves exactly as before.
"""

from __future__ import annotations

from collections import Counter

from online_po_processor.data.master_loader import MasterLoader
from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    BOLD_DATA_FONT, INFO_ITALIC_FONT, OK_FILL,
    auto_width, data_cell, hdr_cell,
)
from online_po_processor.exporter.sheets.tracker_sheet import _MARKETPLACE_DISPLAY

_HEADERS = [
    'Marketplace', 'Type', 'Source Code (EAN)', 'Maps To',
    'Override MRP', 'Override Margin %', 'Use Vendor CP',
    'Effect', 'Note', 'Applied (this run)',
]

_KIND_LABEL = {
    'item_alias': 'Item remap',
    'price_override': 'Price override',
    'vendor_cp': 'Vendor CP',
    'swiggy_deal': 'Swiggy deal',
}


def _norm(s: str) -> str:
    """Space-insensitive, lower-cased marketplace name for comparison."""
    return ''.join(str(s or '').split()).lower()


def write(wb, result: ProcessingResult) -> None:
    registry = getattr(result, 'exception_registry', None) or []
    if not registry:
        return

    # Names that count as "this marketplace's own": the config key
    # (result.marketplace) and its display name (so a file that spells the
    # marketplace either way still highlights). Blank row marketplace = ALL.
    mp_key = result.marketplace or ''
    own_names = {_norm(mp_key),
                 _norm(_MARKETPLACE_DISPLAY.get(mp_key, mp_key))}
    own_names.discard('')

    # Applied-this-run counts, keyed by cleaned source code.
    applied_counts: Counter = Counter()
    for e in (getattr(result, 'exceptions_applied', None) or []):
        applied_counts[MasterLoader._clean_code(e.get('ean', ''))] += 1

    ws = wb.create_sheet('Exceptions')
    for c, h in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, c, h)

    def _is_own(row_mp: str) -> bool:
        rn = _norm(row_mp)
        return rn == '' or rn in own_names      # blank = applies to all

    # Own-marketplace rows first (the operator's focus), then the rest —
    # each group keeps the file's order.
    ordered = sorted(registry, key=lambda e: (not _is_own(e.get('marketplace', '')),))

    own_n = 0
    r = 2
    for e in ordered:
        own = _is_own(e.get('marketplace', ''))
        own_n += 1 if own else 0

        types = ' + '.join(_KIND_LABEL.get(k, k) for k in e.get('kinds', [])) \
            or '—'
        mrp = e.get('override_mrp')
        margin = e.get('override_margin_pct')
        applied = applied_counts.get(MasterLoader._clean_code(e.get('source_code', '')), 0)

        data_cell(ws, r, 1, e.get('marketplace') or 'ALL', align='center')
        data_cell(ws, r, 2, types, align='center')
        data_cell(ws, r, 3, e.get('source_code', ''), align='center')
        data_cell(ws, r, 4, e.get('maps_to') or '', align='center')
        data_cell(ws, r, 5, '' if mrp is None else round(mrp, 2),
                  number_format='#,##0.00', align='right')
        data_cell(ws, r, 6, '' if margin is None else round(margin * 100, 2),
                  number_format='#,##0.00', align='right')
        data_cell(ws, r, 7, 'Y' if e.get('use_vendor_cp') else '—',
                  align='center')
        data_cell(ws, r, 8, e.get('effect', ''), align='left')
        data_cell(ws, r, 9, e.get('note', ''), align='left')
        data_cell(ws, r, 10, f'Yes ×{applied}' if applied else '—',
                  align='center')

        # Highlight this marketplace's own rows (incl. apply-to-all rows).
        if own:
            for c in range(1, len(_HEADERS) + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = OK_FILL
                cell.font = BOLD_DATA_FONT
        r += 1

    # Footer legend.
    r += 1
    disp = _MARKETPLACE_DISPLAY.get(mp_key, mp_key) or mp_key
    ws.cell(
        row=r, column=1,
        value=(f"ℹ Full exception registry from 'Master Exceptions.xlsx' "
               f"({len(registry)} total). Highlighted (green) = applies to "
               f"{disp} ({own_n} row(s), incl. ALL-marketplace rows); others "
               f"shown for awareness. 'Applied (this run)' = overrides that "
               f"fired on this batch. Edit that file to add/remove — no code "
               f"change."),
    ).font = INFO_ITALIC_FONT
    auto_width(ws)
