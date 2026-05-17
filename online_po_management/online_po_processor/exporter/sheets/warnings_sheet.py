"""
exporter.sheets.warnings_sheet
==============================

Writes the **Warnings** sheet — surfaces every non-fatal issue the
engine hit during processing. The sheet is **only created when there
is at least one warning** — a clean run produces a workbook without
this tab, which is itself a signal to the user.

Column layout (3 columns)::

    1. PO        — source PO (empty for global warnings)
    2. Location  — raw location string (empty for non-location warnings)
    3. Warning   — human-readable message (wrapped, multi-line OK)

Warnings come from:

* Unmapped locations (per PO+location, deduped).
* Price mismatches (per item, deduped).
* Unknown GST codes (per code, deduped).
* Missing optional columns (global).
* Rows skipped due to missing data (per PO where feasible).

All warnings use the orange :data:`~._styles.WARN_FILL` header to
distinguish them from ordinary data sheets.

Layout (v2.1.0)
---------------
The Warning column was previously rendered as a single non-wrapping
line capped at the column's natural width — long messages either
overflowed visibly into adjacent cells (when no value sat there) or
got clipped. Now it uses :func:`wrap_data_cell` which sets
``wrap_text=True``, paired with an ``auto_width`` cap of 80 chars on
column C. Excel auto-grows the row height to fit wrapped content as
long as we don't set an explicit row height — so nothing else is
needed for the row to look right at any message length.

PO and Location columns get center alignment (they're identifiers),
Warning gets left alignment (paragraph-style text).
"""

from __future__ import annotations

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    WARN_FILL, auto_width, data_cell, hdr_cell, wrap_data_cell,
)


_HEADERS = ['PO', 'Location', 'Warning']

# v2.1.0: max width for the Warning column. Long messages wrap to this
# width; row height grows to fit. Other columns auto-size as usual.
# 80 chars is roughly the comfortable English-prose line length and
# matches what most code-style guides recommend for readability.
_WARNING_COL_CAP = 80


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Warnings' sheet to ``wb``, but only if there are
    warnings to report. No-op on clean runs.
    """
    if not result.warnings:
        return

    ws = wb.create_sheet('Warnings')

    for col_idx, header in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header, fill=WARN_FILL)

    for r, (po, loc, msg) in enumerate(result.warnings, start=2):
        # PO + Location are identifier columns — center for visual
        # consistency with how identifiers render elsewhere in the
        # workbook (Summary's PO column, Validation's Item No, etc.).
        data_cell(ws, r, 1, po, align='center')
        data_cell(ws, r, 2, loc, align='center')
        # Warning is free-form prose — wrap_data_cell enables
        # wrap_text=True so long messages don't get clipped at the
        # column edge. Row height is left unset; Excel auto-grows it
        # to fit the wrapped content.
        wrap_data_cell(ws, r, 3, msg, align='left')

    # Cap the Warning column at 80 chars so wrap_text actually wraps.
    # Without the cap, auto_width would expand C to fit the longest
    # single line — defeating the whole purpose of wrap_text.
    auto_width(ws, caps={'C': _WARNING_COL_CAP})