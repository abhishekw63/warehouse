"""
exporter.sheets.validation_sheet
================================

Writes the **Validation** sheet — per-item price check with clear
PASS/FAIL status for each line.

Column layout (11 columns, or 14 when HSN check is enabled)::

    1.  PO
    2.  Item No                    — resolved from master
    3.  EAN
    4.  Description                — from Items_March (readable product name)
    5.  MRP                        ─┐
    6.  Landing (m%)                ├ GREEN headers — our calculated values
    7.  GST Code                    │
    8.  Our Cost Price             ─┘
    9.  Marketplace <Label>         — fob_col value from the punch
    10. Difference with <Label>     — fob_price − calc_price
    11. Status                      — OK / MISMATCH / NOT_IN_MASTER / NO_PRICE

    — HSN cross-check columns (v1.6.0, conditional) —
    12. HSN (Marketplace)           — hsn_col value from the punch
    13. HSN (Master)                — HSN/SAC Code from Items_March
    14. HSN Check                   — OK / MISMATCH / NOT_IN_MASTER

``<Label>`` is the marketplace's ``compare_label`` from config (e.g.
"Landing Rate" for Myntra, "Cost" for RK).

Column meaning depends on ``compare_basis``
-------------------------------------------
* ``basis='landing'`` (Myntra, Reliance): Marketplace value is compared
  against "Landing (m%)" (= MRP × m%, pre-GST). Diff is clean (no GST
  rounding).
* ``basis='cost'`` (RK, Blink): Marketplace value is compared against
  "Our Cost Price" (= MRP × m% ÷ GST, post-GST). Diff may have tiny
  rounding noise — we treat ≤ 1 rupee as OK.

HSN cross-check (v1.6.0)
------------------------
When the marketplace has ``hsn_col`` set in its config (currently
Reliance only), the engine compares the punch's HSN against the
master's HSN per row. The three trailing columns appear only when at
least one row has a non-empty ``hsn_check_status`` — otherwise this
sheet keeps its familiar 11-column layout.

Visual cues
-----------
* **Mismatch rows** get a pale-pink fill across the entire row, Status
  cell in bold red.
* **OK rows** get a green status pill only (the bulk of a clean batch,
  so we keep row fill neutral to reduce visual fatigue).
* **NOT_IN_MASTER rows** get a pale-orange fill so these are easy to
  spot and fix by adding the item to Items_March.
* **HSN mismatches** get a red status pill (same bold-red font as
  price mismatches) but don't repaint the full row — the price-diff
  styling stays the primary signal, HSN is a secondary audit flag.

The trailing info row records ``basis=... | Margin: m%`` so someone
reviewing the output three months later can tell at a glance what the
numbers mean.
"""

from __future__ import annotations

import pandas as pd

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    CALC_FILL, HEADER_FILL, INFO_ITALIC_FONT, MISMATCH_FILL,
    MISMATCH_TEXT_FONT, NO_MASTER_FILL, NOT_IN_MASTER_TEXT_FONT,
    STATUS_OK_FILL, STATUS_OK_FONT,
    auto_width, data_cell, hdr_cell,
)


# Calculated column indices (1-based). These get a green header instead
# of the default blue to visually separate "our math" from "their data".
_CALC_COL_INDICES = {5, 6, 7, 8}  # MRP, Landing, GST Code, Our Cost Price


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Validation' sheet to ``wb``.

    v2.1.4: TO marketplaces with master lookup go through the same
    validation pipeline as SO. Pre-v2.1.4, TO mode skipped this sheet
    entirely because Flipkart-TO read Item No / MRP / GST from the
    dump file directly — no master to validate against. With v2.1.4
    Flipkart-TO now uses ``item_resolution='from_ean'`` and master
    lookup, so we DO have data to validate, and a dedicated reference-
    only comparison mode (``compare_mode='reference_only'``) for
    surfacing diffs without flagging them as MISMATCH.

    The sheet is still skipped only when there's truly nothing to
    show — TO mode without a configured ``compare_basis`` (i.e. older
    TO marketplaces or future TO configs that opt out of validation).
    """
    if getattr(result, 'output_type', 'so') == 'to':
        # Only skip if the marketplace genuinely has nothing to
        # validate (no compare_basis configured). With v2.1.4
        # Flipkart-TO sets compare_basis='cost' so the sheet renders.
        if not getattr(result, 'compare_basis', None):
            return

    # v2.1.4: detect reference-only mode for the banner row.
    compare_mode = (result.resolved_config or {}).get('compare_mode')
    is_reference_only = (compare_mode == 'reference_only')

    ws = wb.create_sheet('Validation')

    label = result.compare_label or 'Price'
    margin_pct_int = int(result.margin_pct * 100)

    # v1.6.0: decide up front whether HSN columns belong on this run.
    # Only shown when the engine actually populated an hsn_check_status
    # on at least one row (which happens when the marketplace config
    # has ``hsn_col`` set). This keeps the sheet's width consistent
    # with previous versions for marketplaces that don't opt in.
    has_hsn_check = any(
        so_row.hsn_check_status for so_row in result.rows
    )

    headers = [
        'PO', 'Item No', 'EAN', 'Description', 'MRP',
        f'Landing ({margin_pct_int}%)', 'GST Code',
        'Our Cost Price',
        f'Marketplace {label}',
        f'Difference with {label}',
        'Status',
    ]
    if has_hsn_check:
        headers.extend([
            'HSN (Marketplace)', 'HSN (Master)', 'HSN Check',
        ])

    # ── Header row ──────────────────────────────────────────────────────
    # v2.1.4: when compare_mode='reference_only', prepend a banner row
    # above the headers explaining that the 'Difference' column is
    # informational only — Transfer Prices in the output use the
    # engine's calculated values regardless of any diff. Without this
    # banner, an operator looking at the Validation sheet would
    # reasonably expect the diffs to be acted on (since that's how
    # all other marketplaces' Validation sheets work).
    #
    # The actual cell merge happens AFTER auto_width() at the bottom
    # of this function — auto_width can't iterate columns when row 1
    # has merged cells (MergedCell objects don't expose column_letter).
    # We write the banner text now so it's in place; the merge that
    # spans it across the header columns happens last.
    header_row_idx = 1
    banner_n_cols_for_merge = 0  # tracks merge width; 0 = no banner
    if is_reference_only:
        banner_msg = (
            f"\u2139 Reference-only comparison: '{label}' values from "
            f"the punch file are shown for audit only. Transfer Prices "
            f"in the output use ENGINE-CALCULATED values "
            f"({margin_pct_int}% margin from Items_March master) "
            f"regardless of any diff. Diffs above \u20b90.01 are also "
            f"logged as warnings."
        )
        banner_cell = ws.cell(row=1, column=1, value=banner_msg)
        banner_cell.font = INFO_ITALIC_FONT
        banner_n_cols_for_merge = len(headers)
        header_row_idx = 2

    for col_idx, header in enumerate(headers, start=1):
        fill = CALC_FILL if col_idx in _CALC_COL_INDICES else HEADER_FILL
        hdr_cell(ws, header_row_idx, col_idx, header, fill=fill)

    n_cols = len(headers)
    status_col = 11   # Price-validation status column index (fixed)
    # HSN columns, when present, occupy 12/13/14.
    hsn_punch_col = 12
    hsn_master_col = 13
    hsn_status_col = 14

    # ── Data rows ───────────────────────────────────────────────────────
    # v2.1.0 alignment policy:
    #   PO, Item No, EAN, GST Code, Status, HSN cols → center (identifiers/badges)
    #   Description                                  → left   (long prose)
    #   MRP, Landing, Cost, Marketplace, Difference  → right  (monetary)
    # v2.1.4: data starts at row 3 when reference-only banner is present
    # (banner=row 1, headers=row 2, data=row 3); otherwise row 2 as before.
    r = header_row_idx + 1
    mismatches = 0
    hsn_mismatches = 0
    for so_row in result.rows:
        data_cell(ws, r, 1, so_row.po_number, align='center')
        data_cell(ws, r, 2, so_row.item_no, align='center')
        data_cell(ws, r, 3, so_row.ean, align='center')
        data_cell(ws, r, 4, so_row.description, align='left')
        data_cell(ws, r, 5, so_row.mrp,
                   '#,##0.00' if so_row.mrp else None,
                   align='right')

        # Landing cost (MRP × margin%) — computed fresh for display so the
        # sheet stays self-consistent even if calc_price was derived
        # differently (e.g. RK uses cost basis, but we still want to show
        # Landing here for reference).
        landing = (float(so_row.mrp) * result.margin_pct
                   if so_row.mrp and not pd.isna(so_row.mrp) else None)
        data_cell(ws, r, 6,
                   round(landing, 2) if landing else '',
                   '#,##0.00', align='right')

        data_cell(ws, r, 7, so_row.gst_code, align='center')

        # Our Cost Price (naked CP) — always shown regardless of basis.
        data_cell(ws, r, 8,
                   round(so_row.cost_price_ref, 2)
                   if so_row.cost_price_ref else '',
                   '#,##0.00', align='right')

        # Marketplace value (fob_col)
        data_cell(ws, r, 9,
                   round(so_row.fob_price, 2) if so_row.fob_price else '',
                   '#,##0.00', align='right')

        # Difference (rounded to 2dp — finer is floating-point dust)
        data_cell(ws, r, 10,
                   round(so_row.diffn, 2) if so_row.diffn is not None else '',
                   '#,##0.00', align='right')

        data_cell(ws, r, status_col, so_row.validation_status, align='center')

        # ── Per-status row styling ──────────────────────────────────────
        if so_row.validation_status == 'MISMATCH':
            mismatches += 1
            for c in range(1, n_cols + 1):
                ws.cell(row=r, column=c).fill = MISMATCH_FILL
            ws.cell(row=r, column=status_col).font = MISMATCH_TEXT_FONT

        elif so_row.validation_status == 'OK':
            ws.cell(row=r, column=status_col).fill = STATUS_OK_FILL
            ws.cell(row=r, column=status_col).font = STATUS_OK_FONT

        elif so_row.validation_status == 'NOT_IN_MASTER':
            for c in range(1, n_cols + 1):
                ws.cell(row=r, column=c).fill = NO_MASTER_FILL
            ws.cell(row=r, column=status_col).font = NOT_IN_MASTER_TEXT_FONT

        # ── HSN columns (v1.6.0, only when marketplace opts in) ─────────
        if has_hsn_check:
            data_cell(ws, r, hsn_punch_col, so_row.hsn_punch, align='center')
            data_cell(ws, r, hsn_master_col, so_row.hsn_master, align='center')
            data_cell(ws, r, hsn_status_col, so_row.hsn_check_status,
                       align='center')

            # Mirror the price-status pill styling on the HSN Check
            # cell so the user can scan the column for red at a glance.
            hsn_cell = ws.cell(row=r, column=hsn_status_col)
            if so_row.hsn_check_status == 'OK':
                hsn_cell.fill = STATUS_OK_FILL
                hsn_cell.font = STATUS_OK_FONT
            elif so_row.hsn_check_status == 'MISMATCH':
                hsn_mismatches += 1
                hsn_cell.font = MISMATCH_TEXT_FONT
            elif so_row.hsn_check_status == 'NOT_IN_MASTER':
                hsn_cell.font = NOT_IN_MASTER_TEXT_FONT

        r += 1

    # ── Footer summary ──────────────────────────────────────────────────
    r += 1
    total = len(result.rows)
    ok_count = sum(1 for so_row in result.rows
                    if so_row.validation_status == 'OK')
    basis_note = (f"basis={result.compare_basis} "
                  f"(compared against '{label}')")
    summary_parts = [
        f"Total: {total} items",
        f"OK: {ok_count}",
    ]
    # v2.1.4: in reference-only mode, "Mismatches: 0" is technically true
    # but misleading — the engine downgrades MISMATCH→OK on purpose. Show
    # the count of rows whose |diff| exceeds tolerance instead, which is
    # what an auditor actually cares about.
    if is_reference_only:
        ref_diff_count = sum(
            1 for so_row in result.rows
            if so_row.diffn is not None and abs(so_row.diffn) > 0.01
        )
        summary_parts.append(f"Diffs > \u20b90.01: {ref_diff_count}")
    else:
        summary_parts.append(f"Mismatches: {mismatches}")
    summary_parts.extend([
        f"Margin: {margin_pct_int}%",
        basis_note,
    ])
    if has_hsn_check:
        # Only mention HSN stats when the check ran; otherwise the
        # zero in "HSN mismatches: 0" reads as a claim when really
        # the feature wasn't active.
        summary_parts.insert(3, f"HSN mismatches: {hsn_mismatches}")

    ws.cell(
        row=r, column=1,
        value=" | ".join(summary_parts),
    ).font = INFO_ITALIC_FONT

    auto_width(ws)

    # v2.1.4: NOW it's safe to merge the banner across the header
    # columns. Doing it before auto_width breaks the column iterator
    # (the MergedCell objects in cols 2..N of row 1 don't have
    # column_letter attribute, which auto_width's col[0] indexing
    # depends on).
    if banner_n_cols_for_merge > 0:
        ws.merge_cells(
            start_row=1, start_column=1,
            end_row=1, end_column=banner_n_cols_for_merge,
        )

    # v2.1.4: freeze below headers — A3 when banner present (row1=banner,
    # row2=headers), A2 otherwise. Without this the banner row scrolls
    # away with the data and operators lose the context note.
    ws.freeze_panes = f'A{header_row_idx + 1}'