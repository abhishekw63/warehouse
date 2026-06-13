"""
exporter.sheets.skipped_sheet
=============================

Writes the **Skipped POs** sheet (v2.4.0) — the POs that were *removed*
from this output because they were **already uploaded** (present in the
history DB). Dedup-skip drops them from Headers/Lines so they're never
re-sent to D365; this sheet is the in-file record of what was removed and
why, so nothing disappears silently.

No-op when nothing was skipped (``result.skipped_orders`` empty), so a
normal run's output is unchanged.

Source: ``result.skipped_orders`` (set by
:func:`online_po_processor.auto.history_db.apply_dedup`). Each entry:
``{segment, marketplace, marketplace_label, po, location, po_date,
exp_date, qty, order_value}``.
"""

from __future__ import annotations

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    INFO_ITALIC_FONT, LOC_MISMATCH_FILL,
    auto_width, data_cell, hdr_cell,
)

_HEADERS = ['Segment', 'Market Place', 'PO', 'Location', 'PO Date',
            'Exp Date', 'Order Value', 'Order Qty', 'Status']

_INR = ('[>=10000000]"₹ "##\\,##\\,##\\,##0.00;'
        '[>=100000]"₹ "##\\,##\\,##0.00;'
        '"₹ "##,##0.00')


def write(wb, result: ProcessingResult) -> None:
    skipped = getattr(result, 'skipped_orders', None) or []
    if not skipped:
        return

    ws = wb.create_sheet('Skipped POs')
    for c, h in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, c, h)

    r = 2
    for s in skipped:
        data_cell(ws, r, 1, s.get('segment', ''), align='center')
        data_cell(ws, r, 2, s.get('marketplace_label', ''), align='center')
        data_cell(ws, r, 3, s.get('po', ''), align='center')
        data_cell(ws, r, 4, s.get('location', ''), align='left')
        data_cell(ws, r, 5, s.get('po_date', ''), align='center')
        data_cell(ws, r, 6, s.get('exp_date', ''), align='center')
        data_cell(ws, r, 7, s.get('order_value', 0),
                  number_format=_INR, align='right')
        data_cell(ws, r, 8, s.get('qty', 0), align='center')
        cell = data_cell(ws, r, 9, 'Already uploaded — removed', align='center')
        cell.fill = LOC_MISMATCH_FILL          # amber "attention" tint
        r += 1

    # Footer note.
    r += 1
    ws.cell(row=r, column=1,
            value=(f"⚠ {len(skipped)} PO(s) were already in the history DB and "
                   f"removed from Headers/Lines (not re-sent to D365).")
            ).font = INFO_ITALIC_FONT
    auto_width(ws)
