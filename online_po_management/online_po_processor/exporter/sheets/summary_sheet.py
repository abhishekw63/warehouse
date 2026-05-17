"""
exporter.sheets.summary_sheet
=============================

Writes the **Summary** sheet — a per-PO grouped view for human
verification before the SO is imported.

Column layout (9 columns, v1.5.3)::

    1. PO
    2. Location (Raw)      — what the marketplace sent us
    3. Location (Mapped)   — canonical key matched to from Ship-To registry
    4. Cust No
    5. Ship-to
    6. Items               — count of lines on this PO
    7. Total Qty           — sum of quantities across lines
    8. Total Amount        — sum of SORow.amount across lines (₹, Indian
                             format). Populated when the marketplace has
                             ``amount_col`` configured (Blink, RK);
                             displays ₹0 when not (Myntra today).
    9. Status              — 'OK' (green) or 'UNMAPPED' (red)

Visual aids
-----------
* **Pale yellow fill on both location cells** when the raw and mapped
  names differ (case-insensitive). Means we used a fuzzy match — worth
  a quick eyeball to confirm we matched to the right Ship-To.
* Status pill: green for OK, red for UNMAPPED.
* TOTAL row at the bottom for Items + Qty + Amount — v2.1.0 paints
  every cell across the row with a light grey fill + thick top border
  so it reads as one continuous summary strip rather than fragmented
  values.
* Info sub-row: marketplace, margin %, warehouse, filename, generation
  timestamp, **and run duration** (v2.1.0).
* Legend row appears **only when** there's at least one yellow
  highlight — keeps clean runs free of noise.

Alignment policy (v2.1.0)
-------------------------
``data_cell`` now defaults to right-align for numbers and left-align
for text. The Summary sheet adds explicit overrides:

    * PO, Cust No, Ship-to, Items, Qty, Status   → center
    * Location (Raw) / (Mapped)                  → left
    * Total Amount                               → right (default)
    * 'TOTAL' label                              → center
"""

from __future__ import annotations
from datetime import datetime
from typing import Dict, Optional

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    BOLD_DATA_FONT, INFO_ITALIC_FONT, LEGEND_ITALIC_FONT,
    LOC_MISMATCH_FILL, STATUS_BAD_FILL, STATUS_BAD_FONT,
    STATUS_OK_FILL, STATUS_OK_FONT,
    TOTAL_ROW_FILL, TOTAL_ROW_BORDER,
    auto_width, data_cell, hdr_cell,
)


_HEADERS = [
    'PO', 'Location (Raw)', 'Location (Mapped)',
    'Cust No', 'Ship-to', 'Items', 'Total Qty', 'Total Amount', 'Status',
]

# 1-based column indices for cells we style specially.
_COL_PO = 1
_COL_RAW_LOC = 2
_COL_MAPPED_LOC = 3
_COL_CUST_NO = 4
_COL_SHIP_TO = 5
_COL_ITEMS = 6
_COL_QTY = 7
_COL_AMOUNT = 8    # v1.5.3 — new column
_COL_STATUS = 9    # shifted right by 1 to make room for Amount

# Indian-format rupee number format. The backslash-escaped rupee symbol
# (₹) is literal-escaped so Excel treats it as a currency prefix rather
# than trying to parse it. ``##\\,##\\,##0`` gives the lakh/crore
# grouping (e.g. 14,29,265 instead of 1,429,265). No decimals — amount
# values on marketplace punch files are already at whole-rupee precision
# for the big-picture summary view.
_INR_INDIAN_FORMAT = '[>=10000000]"\u20B9"##\\,##\\,##\\,##0;' \
                      '[>=100000]"\u20B9"##\\,##\\,##0;' \
                      '"\u20B9"##,##0'


def _format_duration(seconds: Optional[float]) -> str:
    """
    Format an elapsed-seconds value for the footer.

    ``None`` → empty string (caller skips the duration segment so the
    footer stays compact for older code paths that don't pass timing).
    Otherwise returns a short human form: ``2.34s`` for sub-minute,
    ``1m 23s`` for sub-hour, ``1h 02m`` for longer runs.
    """
    if seconds is None:
        return ''
    if seconds < 60:
        return f"{seconds:.2f}s"
    if seconds < 3600:
        m = int(seconds // 60)
        s = int(seconds % 60)
        return f"{m}m {s:02d}s"
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return f"{h}h {m:02d}m"


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Summary' sheet to ``wb``.

    For Transfer Order results (output_type='to', currently
    Flipkart-TO), the column layout drops Cust No (TOs have no
    customer) and Total Amount (TOs aren't revenue) — see
    :func:`_write_to`. Otherwise the standard SO layout is rendered.
    """
    if getattr(result, 'output_type', 'so') == 'to':
        _write_to(wb, result)
        return

    ws = wb.create_sheet('Summary')

    # ── Header row ──────────────────────────────────────────────────────
    for col_idx, header in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header)

    # ── Group by PO ─────────────────────────────────────────────────────
    # Every row of a given PO shares location/cust_no/ship_to (guaranteed
    # by the engine — one PO = one delivery location). So we capture those
    # from the first SORow seen for each PO, then accumulate Items + Qty
    # + Amount.
    po_groups: Dict[str, dict] = {}
    for so_row in result.rows:
        if so_row.po_number not in po_groups:
            po_groups[so_row.po_number] = {
                'location': so_row.location,
                'mapped_location': so_row.mapped_location,
                'cust_no': so_row.cust_no,
                'ship_to': so_row.ship_to,
                'mapped': so_row.mapped,
                'items': 0,
                'qty': 0,
                'amount': 0.0,
            }
        po_groups[so_row.po_number]['items'] += 1
        po_groups[so_row.po_number]['qty'] += so_row.qty
        # v1.5.3: sum amount; None (Myntra) contributes 0 silently.
        po_groups[so_row.po_number]['amount'] += float(so_row.amount or 0.0)

    # ── Data rows ───────────────────────────────────────────────────────
    # v2.1.0 alignment overrides per column (everything not listed
    # follows data_cell's smart default — text=left, numbers=right):
    #   PO        → center  (ID-style, even when numeric like Blink's int64)
    #   Cust No   → center  (ERP code)
    #   Ship-to   → center  (ERP code)
    #   Items     → center  (small integer count)
    #   Qty       → center  (operational total — center reads better than right)
    #   Status    → center  (single-word badge)
    #   Locations → left    (long descriptive text)
    #   Amount    → right   (default — column-wise sum legibility)
    r = 2
    for po, info in po_groups.items():
        status = 'OK' if info['mapped'] else 'UNMAPPED'

        data_cell(ws, r, _COL_PO, po, align='center')
        data_cell(ws, r, _COL_RAW_LOC, info['location'], align='left')
        data_cell(ws, r, _COL_MAPPED_LOC, info['mapped_location'],
                   align='left')
        data_cell(ws, r, _COL_CUST_NO, info['cust_no'], align='center')
        data_cell(ws, r, _COL_SHIP_TO, info['ship_to'], align='center')
        data_cell(ws, r, _COL_ITEMS, info['items'], align='center')
        data_cell(ws, r, _COL_QTY, info['qty'], align='center')
        # v1.5.3: Total Amount in INR Indian format (lakh/crore grouping).
        # Stored as the raw float; Excel applies the Indian format so the
        # visible value reads like ₹14,29,265 while sums/filters still
        # work on the underlying number.
        data_cell(
            ws, r, _COL_AMOUNT, info['amount'],
            number_format=_INR_INDIAN_FORMAT,
            align='right',
        )
        data_cell(ws, r, _COL_STATUS, status, align='center')

        # Yellow highlight when raw ≠ mapped (case-insensitive).
        # Indicates a fuzzy match — worth a human glance.
        raw_norm = (info['location'] or '').strip().lower()
        mapped_norm = (info['mapped_location'] or '').strip().lower()
        if info['mapped'] and raw_norm and mapped_norm and raw_norm != mapped_norm:
            ws.cell(row=r, column=_COL_RAW_LOC).fill = LOC_MISMATCH_FILL
            ws.cell(row=r, column=_COL_MAPPED_LOC).fill = LOC_MISMATCH_FILL

        # Status pill
        status_cell = ws.cell(row=r, column=_COL_STATUS)
        if status == 'OK':
            status_cell.fill = STATUS_OK_FILL
            status_cell.font = STATUS_OK_FONT
        else:
            status_cell.fill = STATUS_BAD_FILL
            status_cell.font = STATUS_BAD_FONT

        r += 1

    # ── Totals row ──────────────────────────────────────────────────────
    # v2.1.0: paint the entire row (all 9 columns) with TOTAL_ROW_FILL +
    # bold font + TOTAL_ROW_BORDER (thick top side). Empty gap cells
    # (Loc Raw, Loc Mapped, Cust No, Ship-to, Status) get the same
    # treatment so the strip reads as one continuous summary band.
    total_items = sum(g['items'] for g in po_groups.values())
    total_qty = sum(g['qty'] for g in po_groups.values())
    total_amount = sum(g['amount'] for g in po_groups.values())

    # Per-cell values for the TOTAL row. None means "blank visual cell"
    # but we still write/style it so the row looks continuous.
    total_row_cells = [
        (_COL_PO,         'TOTAL',       'center'),
        (_COL_RAW_LOC,    None,          'left'),
        (_COL_MAPPED_LOC, None,          'left'),
        (_COL_CUST_NO,    None,          'center'),
        (_COL_SHIP_TO,    None,          'center'),
        (_COL_ITEMS,      total_items,   'center'),
        (_COL_QTY,        total_qty,     'center'),
        (_COL_AMOUNT,     total_amount,  'right'),
        (_COL_STATUS,     None,          'center'),
    ]
    for col, value, align in total_row_cells:
        # Use the standard data_cell so border + alignment apply, then
        # overlay the TOTAL-row treatment (fill + bold font + thick
        # top border) on top.
        nf = _INR_INDIAN_FORMAT if col == _COL_AMOUNT and value is not None else None
        cell = data_cell(ws, r, col, value if value is not None else '',
                          number_format=nf, align=align)
        cell.fill = TOTAL_ROW_FILL
        cell.font = BOLD_DATA_FONT
        cell.border = TOTAL_ROW_BORDER

    # ── Info sub-row ────────────────────────────────────────────────────
    r += 2
    margin_str = f"{int(result.margin_pct * 100)}%"

    # v1.9.0: surface the warehouse the D365 export used so operations
    # can reconcile which RENEE warehouse this batch ships from.
    # Backwards-compatible: older ProcessingResults (constructed
    # without going through v1.9.0 GUI) keep the pre-v1.9.0 footer
    # shape since warehouse_display defaults to 'AHD' silently.
    wh_display = getattr(result, 'warehouse_display', '') or 'AHD'
    wh_code = getattr(result, 'warehouse_code', '') or 'PICK'

    # v2.1.0: append run duration when available. The exporter computes
    # ``elapsed_seconds`` just before saving so this footer shows the
    # full pipeline time (engine + export). On code paths that don't
    # populate it, the segment is silently omitted.
    duration_str = _format_duration(result.elapsed_seconds)
    duration_segment = f"  |  Duration: {duration_str}" if duration_str else ''

    info_text = (f"Marketplace: {result.marketplace}  |  "
                 f"Margin: {margin_str}  |  "
                 f"Warehouse: {wh_display} ({wh_code})  |  "
                 f"File: {result.input_file}  |  "
                 f"Generated: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
                 f"{duration_segment}")
    ws.cell(row=r, column=1, value=info_text).font = INFO_ITALIC_FONT

    # ── Legend row (conditional) ────────────────────────────────────────
    # Only show when at least one yellow highlight exists — otherwise the
    # legend is noise in a clean run.
    any_loc_mismatch = any(
        (g['mapped']
         and (g['location'] or '').strip().lower()
             != (g['mapped_location'] or '').strip().lower()
         and g['location'] and g['mapped_location'])
        for g in po_groups.values()
    )
    if any_loc_mismatch:
        r += 1
        ws.cell(
            row=r, column=1,
            value=("🟨 Yellow = raw and mapped location differ "
                   "(fuzzy match) — please verify."),
        ).font = LEGEND_ITALIC_FONT

    auto_width(ws)


# ── v2.0.0: Transfer Order summary ─────────────────────────────────────

# TO summary columns. Drops 'Cust No' (TOs have no customer) and
# 'Total Amount' (TOs aren't revenue, only inter-warehouse stock
# movements). Renames 'Ship-to' → 'Transfer-to' since that's what
# the value semantically is in TO context.
_TO_HEADERS = [
    'PO', 'Location (Raw)', 'Location (Mapped)',
    'Transfer-to', 'Items', 'Total Qty', 'Status',
]
# 1-based column indices for TO mode
_TO_COL_PO = 1
_TO_COL_RAW_LOC = 2
_TO_COL_MAPPED_LOC = 3
_TO_COL_TRANSFER_TO = 4
_TO_COL_ITEMS = 5
_TO_COL_QTY = 6
_TO_COL_STATUS = 7


def _write_to(wb, result: ProcessingResult) -> None:
    """
    v2.0.0: TO-mode counterpart of :func:`write`.

    Same per-PO grouping and visual treatment as the SO summary,
    but with a slimmer 7-column layout that drops customer/amount
    info that doesn't apply to inter-warehouse transfers.

    v2.1.0: alignment overrides + TOTAL-row strip + duration footer
    applied identically to the SO writer.

    Status semantics:
      * 'OK'         — mapping resolved cleanly with non-empty
                       Transfer-to Code
      * 'NO_TO_CODE' — mapping row exists in Ship-To B2B but its
                       Ship to column is blank (Howrah-style row)
      * 'UNMAPPED'   — no mapping row at all (shouldn't happen in a
                       clean Flipkart run)
    """
    ws = wb.create_sheet('Summary')

    for col_idx, header in enumerate(_TO_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header)

    # Group by PO — every row of a PO shares location/ship_to.
    po_groups: Dict[str, dict] = {}
    for so_row in result.rows:
        if so_row.po_number not in po_groups:
            po_groups[so_row.po_number] = {
                'location': so_row.location,
                'mapped_location': so_row.mapped_location,
                'ship_to': so_row.ship_to,
                'mapped': so_row.mapped,
                'items': 0,
                'qty': 0,
            }
        po_groups[so_row.po_number]['items'] += 1
        po_groups[so_row.po_number]['qty'] += so_row.qty

    # Data rows
    r = 2
    for po, info in po_groups.items():
        if not info['mapped']:
            status = 'UNMAPPED'
        elif not info['ship_to']:
            status = 'NO_TO_CODE'   # Howrah-style — mapped but Ship to blank
        else:
            status = 'OK'

        data_cell(ws, r, _TO_COL_PO, po, align='center')
        data_cell(ws, r, _TO_COL_RAW_LOC, info['location'], align='left')
        data_cell(ws, r, _TO_COL_MAPPED_LOC, info['mapped_location'],
                   align='left')
        data_cell(ws, r, _TO_COL_TRANSFER_TO, info['ship_to'], align='center')
        data_cell(ws, r, _TO_COL_ITEMS, info['items'], align='center')
        data_cell(ws, r, _TO_COL_QTY, info['qty'], align='center')
        data_cell(ws, r, _TO_COL_STATUS, status, align='center')

        # Yellow highlight for fuzzy location matches (same as SO).
        raw_norm = (info['location'] or '').strip().lower()
        mapped_norm = (info['mapped_location'] or '').strip().lower()
        if (info['mapped'] and raw_norm and mapped_norm
                and raw_norm != mapped_norm):
            ws.cell(row=r, column=_TO_COL_RAW_LOC).fill = LOC_MISMATCH_FILL
            ws.cell(row=r, column=_TO_COL_MAPPED_LOC).fill = LOC_MISMATCH_FILL

        # Status pill — green for OK, red for everything else.
        status_cell = ws.cell(row=r, column=_TO_COL_STATUS)
        if status == 'OK':
            status_cell.fill = STATUS_OK_FILL
            status_cell.font = STATUS_OK_FONT
        else:
            status_cell.fill = STATUS_BAD_FILL
            status_cell.font = STATUS_BAD_FONT

        r += 1

    # Totals row (v2.1.0 continuous strip — all 7 columns styled).
    total_items = sum(g['items'] for g in po_groups.values())
    total_qty = sum(g['qty'] for g in po_groups.values())

    total_row_cells = [
        (_TO_COL_PO,          'TOTAL',     'center'),
        (_TO_COL_RAW_LOC,     None,        'left'),
        (_TO_COL_MAPPED_LOC,  None,        'left'),
        (_TO_COL_TRANSFER_TO, None,        'center'),
        (_TO_COL_ITEMS,       total_items, 'center'),
        (_TO_COL_QTY,         total_qty,   'center'),
        (_TO_COL_STATUS,      None,        'center'),
    ]
    for col, value, align in total_row_cells:
        cell = data_cell(ws, r, col, value if value is not None else '',
                          align=align)
        cell.fill = TOTAL_ROW_FILL
        cell.font = BOLD_DATA_FONT
        cell.border = TOTAL_ROW_BORDER

    # Info sub-row (omit "Margin: 60%" — present but less relevant
    # for TO, kept anyway so users can verify the value used).
    r += 2
    margin_str = f"{int(result.margin_pct * 100)}%"
    wh_display = getattr(result, 'warehouse_display', '') or 'AHD'
    wh_code = getattr(result, 'warehouse_code', '') or 'PICK'
    duration_str = _format_duration(result.elapsed_seconds)
    duration_segment = f"  |  Duration: {duration_str}" if duration_str else ''

    info_text = (f"Marketplace: {result.marketplace}  |  "
                 f"Margin: {margin_str}  |  "
                 f"Warehouse: {wh_display} ({wh_code})  |  "
                 f"File: {result.input_file}  |  "
                 f"Generated: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
                 f"{duration_segment}")
    ws.cell(row=r, column=1, value=info_text).font = INFO_ITALIC_FONT

    # Legend (conditional — same as SO)
    any_loc_mismatch = any(
        (g['mapped']
         and (g['location'] or '').strip().lower()
             != (g['mapped_location'] or '').strip().lower()
         and g['location'] and g['mapped_location'])
        for g in po_groups.values()
    )
    if any_loc_mismatch:
        r += 1
        ws.cell(
            row=r, column=1,
            value=("🟨 Yellow = raw and mapped location differ "
                   "(fuzzy match) — please verify."),
        ).font = LEGEND_ITALIC_FONT

    auto_width(ws)