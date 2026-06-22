"""
exporter.sheets.flipkart_tracker_sheet
======================================

Writes the **Tracker** sheet for Flipkart from the rows built off the
optional uploaded header file (``result.flipkart_tracker_rows``; see
``engine.flipkart_tracker``). One row per PO with Market Place (resolved by
location), PO / Location / PO Date / Exp Date / PO Aging For Exp / Order Value
/ Order Qty.

No-op when ``flipkart_tracker_rows`` is empty (i.e. the operator didn't upload
a header file, or the marketplace isn't Flipkart) — so every other run is
unchanged. Unknown-location rows show Market Place ``'FK (review)'`` and are
amber-highlighted so they stand out for the operator to map.
"""
from __future__ import annotations

from online_po_processor.data.models import ProcessingResult
from online_po_processor.engine.flipkart_tracker import (
    TRACKER_COLUMNS, unknown_locations,
)
from online_po_processor.exporter._styles import (
    BOLD_DATA_FONT, INFO_ITALIC_FONT, LOC_MISMATCH_FILL,
    auto_width, data_cell, hdr_cell,
)

_REVIEW = 'FK (review)'


def write(wb, result: ProcessingResult) -> None:
    rows = getattr(result, 'flipkart_tracker_rows', None) or []
    if not rows:
        return

    ws = wb.create_sheet('Tracker')
    for c, h in enumerate(TRACKER_COLUMNS, start=1):
        hdr_cell(ws, 1, c, h)

    money_cols = {'Order Value'}
    int_cols = {'Order Qty', 'PO Aging For Exp'}
    r = 2
    review_n = 0
    for row in rows:
        is_review = row.get('Market Place') == _REVIEW
        review_n += 1 if is_review else 0
        for c, key in enumerate(TRACKER_COLUMNS, start=1):
            val = row.get(key)
            if val is None:
                val = ''
            if key in money_cols and val != '':
                # Indian grouping + ₹ (matches the master tracker: ₹4,27,962.90).
                cell = data_cell(ws, r, c, round(float(val), 2),
                                 number_format='"₹" #,##,##0.00', align='right')
            elif key in int_cols and val != '':
                cell = data_cell(ws, r, c, val, align='center')
            else:
                cell = data_cell(ws, r, c, val,
                                 align='center' if key != 'Location' else 'left')
            # Flag unknown-location rows (the Market Place cell) for review.
            if is_review and key == 'Market Place':
                cell.fill = LOC_MISMATCH_FILL
                cell.font = BOLD_DATA_FONT
        r += 1

    # Footer note.
    r += 1
    note = (f"ℹ Flipkart Tracker from the uploaded header file "
            f"({len(rows)} PO(s)). Market Place is assigned by Location "
            f"(locked mapping).")
    if review_n:
        unk = ', '.join(unknown_locations(rows))
        note += (f"  {review_n} PO(s) at UNKNOWN location(s) shown as "
                 f"'{_REVIEW}' — add to the mapping: {unk}.")
    ws.cell(row=r, column=1, value=note).font = INFO_ITALIC_FONT
    auto_width(ws)
