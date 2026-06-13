"""
exporter.sheets.headers_sheet
=============================

Writes the **Headers (SO)** sheet — the SO header rows that the ERP
imports to create the document shells. One row per unique PO number.

Column layout (18 columns)::

    1.  Document Type             = 'Order'
    2.  No.                       = the PO number
    3.  Sell-to Customer No.      = from mapping
    4.  Ship-to Code              = from mapping
    5.  Posting Date              = today
    6.  Order Date                = today
    7.  Document Date             = today
    8.  Invoice From Date         = today
    9.  Invoice To Date           = today
    10. External Document No.     = the PO number (same as col 2)
    11. Location Code             = the selected dispatch warehouse's
                                    resolved D365 code (v2.3.1); 'PICK'
                                    when no warehouse was chosen
    12. Dimension Set ID          = ''
    13. Supply Type               = 'B2B'   (always)
    14. Voucher Narration         = ''
    15. Brand Code (Dimension)    = ''
    16. Channel Code (Dimension)  = ''
    17. Catagory (Dimension)      = ''      (sic — matches ERP's spelling)
    18. Geography Code (Dimension)= ''

The dimension columns (15-18) are blank by design — the ERP fills them
from defaults.

Transfer Order layout (v2.0.0)
------------------------------
When ``result.output_type == 'to'`` (currently Flipkart-TO), this
module renders a 'Headers (TO)' sheet instead with a Transfer Header
column layout. One row per unique PO, mirroring the D365 TO import
columns the ERP team consumes:

    1. No.                  = PO number (numeric in source — kept as-is)
    2. Transfer-from Code   = 'PICK'  (always — origin warehouse)
    3. Transfer-to Code     = ship_to from Ship-To B2B (e.g. FK_BHW_BTS)
    4. Posting Date         = today
    5. In-Transit Code      = 'IN TRANSIT' (always)
    6. Direct Transfer      = 'false' (always)

Dimension columns (Brand/Channel/Category/etc.) are not rendered in
TO mode — they don't apply to inter-warehouse transfers and weren't
in the manual D365 TO dump that defined this layout.
"""

from __future__ import annotations
from datetime import datetime

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    auto_width, data_cell, hdr_cell,
)


# Column headers in display order (1-based positions match docstring).
_HEADERS = [
    'Document Type', 'No.', 'Sell-to Customer No.', 'Ship-to Code',
    'Posting Date', 'Order Date', 'Document Date',
    'Invoice From Date', 'Invoice To Date',
    'External Document No.', 'Location Code', 'Dimension Set ID',
    'Supply Type', 'Voucher Narration',
    'Brand Code (Dimension)', 'Channel Code (Dimension)',
    'Catagory (Dimension)', 'Geography Code (Dimension)',
]

# v2.0.0: TO header columns. Same row-per-unique-PO logic, different
# column set. Constants kept here at module top so they're visible
# alongside the SO _HEADERS for direct comparison.
_TO_HEADERS = [
    'No.', 'Transfer-from Code', 'Transfer-to Code',
    'Posting Date', 'In-Transit Code', 'Direct Transfer',
]
_TO_TRANSFER_FROM = 'PICK'
_TO_IN_TRANSIT = 'IN TRANSIT'
_TO_DIRECT_TRANSFER = 'false'


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Headers (SO)' or 'Headers (TO)' sheet to ``wb``.

    Args:
        wb:     openpyxl Workbook to write into.
        result: ProcessingResult — only ``result.rows`` is consulted (to
                derive the unique PO list). The output_type field
                governs whether the SO or TO layout is rendered.
    """
    # v2.0.0: branch on output_type — TO marketplaces (Flipkart-TO)
    # render a Transfer Header layout; everything else uses the SO
    # layout. Defaulting to 'so' via getattr keeps backwards compat
    # with results constructed before output_type was added.
    if getattr(result, 'output_type', 'so') == 'to':
        _write_to(wb, result)
        return

    ws = wb.create_sheet('Headers (SO)')

    for col_idx, header in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header)

    today_str = datetime.now().strftime("%d-%m-%Y")

    # v2.3.1: Location Code now follows the dispatch warehouse the
    # operator picked in the GUI (AHD/BLR/...), resolved to its D365
    # code on ``result.warehouse_code`` — the SAME value the Summary
    # footer reports. Falls back to the legacy 'PICK' constant when no
    # warehouse was selected (empty/None) or for legacy runs predating
    # the AHD/BLR selector, preserving prior behaviour.
    location_code = getattr(result, 'warehouse_code', '') or 'PICK'

    # Collect unique POs preserving the order they were processed in.
    # We use a set for O(1) membership check and a parallel list for order.
    seen: set = set()
    unique_po_rows = []
    for so_row in result.rows:
        if so_row.po_number not in seen:
            seen.add(so_row.po_number)
            unique_po_rows.append(so_row)

    # One header row per unique PO. We pull cust_no / ship_to from the
    # FIRST SORow we saw for that PO (all rows of a single PO share the
    # same delivery location, so the values are identical).
    #
    # v2.1.0 alignment: every column on Headers (SO) is a short
    # fixed-width identifier (Order/Item literals, dates, ERP codes,
    # PO numbers) — center-align all of them so the row reads as a
    # neat ID strip.
    for r, so_row in enumerate(unique_po_rows, start=2):
        data_cell(ws, r, 1, 'Order', align='center')              # Document Type
        data_cell(ws, r, 2, so_row.po_number, align='center')     # No.
        data_cell(ws, r, 3, so_row.cust_no, align='center')       # Sell-to Customer No.
        data_cell(ws, r, 4, so_row.ship_to, align='center')       # Ship-to Code
        data_cell(ws, r, 5, today_str, align='center')            # Posting Date
        data_cell(ws, r, 6, today_str, align='center')            # Order Date
        data_cell(ws, r, 7, today_str, align='center')            # Document Date
        data_cell(ws, r, 8, today_str, align='center')            # Invoice From Date
        data_cell(ws, r, 9, today_str, align='center')            # Invoice To Date
        data_cell(ws, r, 10, so_row.po_number, align='center')    # External Document No.
        data_cell(ws, r, 11, location_code, align='center')       # Location Code (v2.3.1)
        data_cell(ws, r, 12, '', align='center')                  # Dimension Set ID
        data_cell(ws, r, 13, 'B2B', align='center')               # Supply Type
        # Columns 14–18 left blank (Voucher Narration + 4 dimension cols).

    auto_width(ws)


def _write_to(wb, result: ProcessingResult) -> None:
    """
    v2.0.0: TO-mode counterpart of :func:`write`.

    Renders one row per unique PO using the Transfer Header column
    layout. Same row-collection logic as the SO path — collect unique
    POs preserving processing order, write one header row each.

    The PO number is written as numeric when it parses cleanly to int
    (Flipkart's Po Number column is int64) — that matches the manual
    D365 TO dump which stores 204345116 as a number, not a string.
    Falls back to string for non-numeric PO codes (defensive only;
    Flipkart-TO is numeric today).
    """
    ws = wb.create_sheet('Headers (TO)')

    for col_idx, header in enumerate(_TO_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header)

    today_str = datetime.now().strftime("%d-%m-%Y")

    seen: set = set()
    unique_po_rows = []
    for so_row in result.rows:
        if so_row.po_number not in seen:
            seen.add(so_row.po_number)
            unique_po_rows.append(so_row)

    for r, so_row in enumerate(unique_po_rows, start=2):
        # v2.1.0: Headers (TO) is identical-shape to Headers (SO) — all
        # cells are short fixed-width identifiers, so center-align all.
        # Col 1: No. — numeric when possible
        try:
            data_cell(ws, r, 1, int(so_row.po_number), align='center')
        except (ValueError, TypeError):
            data_cell(ws, r, 1, so_row.po_number, align='center')
        # Col 2: Transfer-from Code (constant)
        data_cell(ws, r, 2, _TO_TRANSFER_FROM, align='center')
        # Col 3: Transfer-to Code (= ship_to from Ship-To B2B). Left
        # blank when the mapping row had no Ship to code (Howrah case).
        data_cell(ws, r, 3, so_row.ship_to or '', align='center')
        # Col 4: Posting Date
        data_cell(ws, r, 4, today_str, align='center')
        # Col 5: In-Transit Code (constant)
        data_cell(ws, r, 5, _TO_IN_TRANSIT, align='center')
        # Col 6: Direct Transfer (constant)
        data_cell(ws, r, 6, _TO_DIRECT_TRANSFER, align='center')

    auto_width(ws)