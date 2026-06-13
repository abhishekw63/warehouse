"""
exporter.sheets.lines_sheet
===========================

Writes the **Lines (SO)** sheet — the SO line rows imported into the
ERP. One row per ordered item.

Column layout (8 columns)::

    1. Document Type   = 'Order'
    2. Document No.    = the PO number (matches Headers (SO) col 2)
    3. Line No.        = 10000, 20000, 30000, ... (resets on new PO)
    4. Type            = 'Item'  (always — no service/charge lines here)
    5. No.             = the resolved Item No
    6. Location Code   = the selected dispatch warehouse's resolved D365
                         code (v2.3.1); 'PICK' when no warehouse chosen
    7. Quantity        = order qty
    8. Unit Price      = ''      (left blank — WMS computes downstream)

The 10000-step Line No. is an ERP convention so users can later insert
extra lines between existing ones.

Transfer Order layout (v2.0.0)
------------------------------
When ``result.output_type == 'to'`` (currently Flipkart-TO), this
module renders a 'Lines (TO)' sheet with the Transfer Line column
layout instead. One row per ``SORow``, mirroring the manual D365 TO
dump:

    1. Document No.        = PO number
    2. Line No.            = 10000, 20000, ... (resets on new PO,
                              STRING in source so we mirror as-is)
    3. Item No.            = item_no
    4. Quantity            = qty
    5. Unit of Measure     = ''  (blank; matches manual dump)
    6. Qty. to Ship        = ''
    7. Qty. to Receive     = ''
    8. Dimension Set ID    = ''
    9. Transfer Price      = calc_price (post-GST cost: MRP × m% ÷ GST)

The 9th column is the key reason for a TO-specific layout. Without
it the ERP would derive Transfer Price from the item master's
default vendor cost, which doesn't reflect the marketplace's actual
agreed cost. We carry our calculated value through so the ERP
records the correct intercompany transfer value.
"""

from __future__ import annotations

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    INFO_ITALIC_FONT, WARN_FILL,
    auto_width, data_cell, hdr_cell,
)


_HEADERS = [
    'Document Type', 'Document No.', 'Line No.', 'Type',
    'No.', 'Location Code', 'Quantity', 'Unit Price',
]

# v2.0.0: TO line columns. 9 cols vs SO's 8 — the extra is Transfer
# Price (col 9) which carries our calculated post-GST cost. Other
# differences: no 'Document Type'/'Type' constant cells, blank
# Unit of Measure / Qty.to Ship / Qty.to Receive / Dimension Set ID.
_TO_HEADERS = [
    'Document No.', 'Line No.', 'Item No.', 'Quantity',
    'Unit of Measure', 'Qty. to Ship', 'Qty. to Receive',
    'Dimension Set ID', 'Transfer Price',
]

# Step between consecutive Line No. values within a PO.
_LINE_NO_STEP = 10_000


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Lines (SO)' or 'Lines (TO)' sheet to ``wb``.

    Args:
        wb:     openpyxl Workbook to write into.
        result: ProcessingResult — emits one line row per ``result.rows``
                entry, in the engine's processing order. The
                output_type field decides SO vs TO layout.
    """
    if getattr(result, 'output_type', 'so') == 'to':
        _write_to(wb, result)
        return

    ws = wb.create_sheet('Lines (SO)')

    # v2.1.3: read the runtime override flag. When True (toggle on in
    # the GUI), we populate col 8 with the engine-computed Cost Price
    # and tint the header amber + add an info row at the bottom so the
    # user can tell at a glance that values were stamped (vs a normal
    # run where col 8 is blank for downstream WMS computation).
    override = bool(getattr(result, 'override_unit_price', False))

    # v2.3.1: Location Code now follows the dispatch warehouse the
    # operator picked in the GUI (AHD/BLR/...), resolved to its D365
    # code on ``result.warehouse_code`` — the SAME value the Summary
    # footer reports. Falls back to the legacy 'PICK' constant when no
    # warehouse was selected (empty/None) or for legacy runs predating
    # the AHD/BLR selector, preserving prior behaviour.
    location_code = getattr(result, 'warehouse_code', '') or 'PICK'

    for col_idx, header in enumerate(_HEADERS, start=1):
        cell = hdr_cell(ws, 1, col_idx, header)
        # v2.1.3: amber tint on the Unit Price header when override is
        # active. Reuses WARN_FILL (the orange used by Warnings sheet
        # headers) so the visual language stays consistent — orange
        # means "needs attention, atypical state". When override is
        # off, the column header keeps the standard deep-blue HEADER_FILL.
        if override and col_idx == 8:
            cell.fill = WARN_FILL

    # Track Line No. per PO. We don't pre-group — we rely on the engine
    # emitting a PO's rows contiguously, which is true today (rows are in
    # input-file order, and a PO's lines are always contiguous in punch
    # files).
    current_po = None
    line_no = 0

    # v2.1.0 alignment policy for Lines (SO):
    #   Cols 1-6 (Document Type, PO, Line No, Type, Item No, Location)
    #     → center (all are short identifiers / fixed literals)
    #   Col 7 (Quantity)   → center (operational value, single integer)
    #   Col 8 (Unit Price) → right (monetary; blank by default for WMS
    #     computation, populated when override toggle is on)
    last_row = 1
    for r, so_row in enumerate(result.rows, start=2):
        if so_row.po_number != current_po:
            current_po = so_row.po_number
            line_no = 0

        line_no += _LINE_NO_STEP

        data_cell(ws, r, 1, 'Order', align='center')          # Document Type
        data_cell(ws, r, 2, so_row.po_number, align='center') # Document No.
        data_cell(ws, r, 3, line_no, align='center')          # Line No.
        data_cell(ws, r, 4, 'Item', align='center')           # Type
        data_cell(ws, r, 5, so_row.item_no, align='center')   # No. (Item No)
        data_cell(ws, r, 6, location_code, align='center')    # Location Code (v2.3.1)
        data_cell(ws, r, 7, so_row.qty, align='center')       # Quantity

        # v2.1.5: Unit Price override must ALWAYS use post-GST Cost
        # Price, not the marketplace's active comparison value. For
        # landing-basis marketplaces (BlinkMP/Myntra/Flipkart),
        # ``calc_price`` is intentionally the pre-GST Landing Rate for
        # validation; ``cost_price_ref`` is the naked Cost Price we
        # want to stamp into Unit Price. Per-row defensive: if the
        # master lookup failed and no cost can be computed, leave col 8
        # blank for THIS row only, matching the D365 exporter's lenient
        # behaviour.
        if override and so_row.cost_price_ref is not None:
            data_cell(ws, r, 8, round(so_row.cost_price_ref, 2),
                       number_format='#,##0.00', align='right')
        else:
            data_cell(ws, r, 8, '', align='right')

        last_row = r

    auto_width(ws)

    # v2.1.3: info-row footer when override is active. Single merged
    # cell, italic, light grey — same pattern Summary/Validation use
    # for footer notes. Sits 2 rows below the data so it doesn't get
    # mistaken for an extra line. Only rendered when override is on,
    # so non-override runs look identical to pre-v2.1.3 output.
    if override:
        info_row = last_row + 2
        info_text = (
            "ⓘ Unit Price overridden — values written from the "
            "engine's computed Cost Price (MRP × margin% ÷ GST). "
            "The downstream D365 import will use these instead of the "
            "vendor master's default. Untick 'Override Unit Price' in "
            "the GUI to revert to blank Unit Price (WMS-computed)."
        )
        cell = ws.cell(row=info_row, column=1, value=info_text)
        cell.font = INFO_ITALIC_FONT
        # Merge across all 8 columns so the note reads as one paragraph
        # rather than cramped into col A.
        ws.merge_cells(start_row=info_row, start_column=1,
                        end_row=info_row, end_column=8)


def _write_to(wb, result: ProcessingResult) -> None:
    """
    v2.0.0: TO-mode counterpart of :func:`write`.

    Writes one Transfer Line per ``SORow``. Same line numbering
    convention as SO (10000-step within a PO, reset at PO boundary).

    Transfer Price (col 9) populated when the engine computed a
    cost_price; left blank if the row's MRP was unparseable (engine
    skips those rows before reaching here, so this is defensive only).
    """
    ws = wb.create_sheet('Lines (TO)')

    for col_idx, header in enumerate(_TO_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header)

    current_po = None
    line_no = 0

    # v2.1.0 alignment for Lines (TO):
    #   Cols 1-3 (PO, Line No, Item No)  → center (identifiers)
    #   Col 4 (Quantity)                 → center (operational value)
    #   Cols 5-8 (blank pass-throughs)   → center (matches surrounding visual)
    #   Col 9 (Transfer Price)           → right (monetary value)
    for r, so_row in enumerate(result.rows, start=2):
        if so_row.po_number != current_po:
            current_po = so_row.po_number
            line_no = 0
        line_no += _LINE_NO_STEP

        # Col 1: Document No. (PO)
        try:
            data_cell(ws, r, 1, int(so_row.po_number), align='center')
        except (ValueError, TypeError):
            data_cell(ws, r, 1, so_row.po_number, align='center')
        # Col 2: Line No. — STRING ('10000', '20000', ...) per manual
        # D365 TO dump
        data_cell(ws, r, 2, str(line_no), align='center')
        # Col 3: Item No.
        data_cell(ws, r, 3, so_row.item_no, align='center')
        # Col 4: Quantity
        data_cell(ws, r, 4, so_row.qty, align='center')
        # Cols 5-8: blank
        for c in range(5, 9):
            data_cell(ws, r, c, '', align='center')
        # Col 9: Transfer Price — post-GST cost from engine
        if so_row.calc_price is not None:
            data_cell(ws, r, 9, so_row.calc_price, '#,##0.0000',
                       align='right')
        else:
            data_cell(ws, r, 9, '', align='right')

    auto_width(ws)
