"""
exporter.sheets.tracker_sheet
=============================

Writes the **Tracker** sheet — a per-PO pivot shaped to drop straight
into the operations team's master PO tracker by copy-paste. Unlike the
Summary sheet (which is for in-workbook human verification), this sheet
exists purely to be selected and pasted into an external tracker, so the
column order, labels, and blank columns are fixed to match that tracker's
layout cell-for-cell.

Column layout (9 columns, paste-ready order)::

    1. Market Place        — result.marketplace (e.g. 'Zepto')
    2. PO                  — PO number
    3. Location            — raw delivery location from the punch
    4. State Name          — intentionally BLANK (filled manually)
    5. PO Date             — from the punch's 'PO Date' column
    6. Exp Date            — from the punch's 'PO Expiry Date' column
    7. PO Aging For Exp    — intentionally BLANK (filled manually)
    8. Order Value         — Σ row amount across the PO (₹, 2 dp)
    9. Order Qty           — Σ qty across the PO

State Name and PO Aging For Exp are left empty on purpose — the user
fills them downstream — but the columns are still emitted so the paste
lands in the right cells.

Scope (v1)
----------
Zepto only. Other marketplaces don't carry the PO Date / PO Expiry Date
columns this view depends on, and the tracker format hasn't been
confirmed for them yet. :func:`write` is a no-op for any non-Zepto
result, so adding it to the exporter's sheet list is safe for every
marketplace — the sheet simply won't appear unless the run is Zepto.
Broaden the ``_SUPPORTED`` guard once the other marketplaces' date
columns and tracker layout are confirmed.
"""

from __future__ import annotations
from typing import Dict, Optional

import pandas as pd

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    auto_width, data_cell, hdr_cell,
)


_HEADERS = [
    'Market Place', 'PO', 'Location', 'State Name',
    'PO Date', 'Exp Date', 'PO Aging For Exp',
    'Order Value', 'Order Qty',
]

# 1-based column indices.
_COL_MARKETPLACE = 1
_COL_PO = 2
_COL_LOCATION = 3
_COL_STATE = 4          # intentionally blank
_COL_PO_DATE = 5
_COL_EXP_DATE = 6
_COL_AGING = 7          # intentionally blank
_COL_ORDER_VALUE = 8
_COL_ORDER_QTY = 9

# Marketplaces this sheet is wired up for. Zepto only for now — see the
# module docstring's Scope note.
_SUPPORTED = {'Zepto'}

# Rupee, Indian (lakh/crore) digit grouping, 2 decimals. Matches the
# master tracker's "₹ 20,683.60" presentation. Mirrors the Summary
# sheet's Indian format but keeps the paise instead of rounding to
# whole rupees (the tracker reconciles against invoices to the paisa).
_INR_FORMAT = ('[>=10000000]"₹ "##\\,##\\,##\\,##0.00;'
               '[>=100000]"₹ "##\\,##\\,##0.00;'
               '"₹ "##,##0.00')

# Candidate raw-column names for the two date fields, in preference
# order. Matched case-insensitively + whitespace-tolerantly against the
# punch's actual headers (see :func:`_find_col`) so a casing/spacing
# drift in a future Zepto dump doesn't silently blank the column.
_PO_DATE_CANDIDATES = ['PO Date']
_EXP_DATE_CANDIDATES = ['PO Expiry Date', 'PO Expiry', 'Expiry Date']


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Tracker' sheet to ``wb`` for supported marketplaces.

    No-op (no sheet created) when:
      * the marketplace isn't in ``_SUPPORTED`` (currently Zepto only),
      * there are no rows to summarise, or
      * the raw DataFrame is missing (can't source the date columns).

    Skipping cleanly rather than writing an empty/partial sheet keeps
    non-Zepto output identical to before this sheet existed.
    """
    if result.marketplace not in _SUPPORTED:
        return
    if not result.rows:
        return

    # Per-PO date lookup, sourced from the raw punch. Empty dict when
    # raw_df is unavailable — date cells then render blank rather than
    # crashing.
    po_dates = _build_po_date_lookup(result)

    ws = wb.create_sheet('Tracker')

    # ── Header row ──────────────────────────────────────────────────────
    for col_idx, header in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header)

    # ── Group by PO (insertion order preserved → rows match punch order)
    # Every row of a PO shares location (one PO = one delivery location,
    # guaranteed by the engine), so we capture it from the first row seen
    # and accumulate qty + amount.
    po_groups: Dict[str, dict] = {}
    for so_row in result.rows:
        if so_row.po_number not in po_groups:
            po_groups[so_row.po_number] = {
                'location': so_row.location,
                'qty': 0,
                'amount': 0.0,
            }
        po_groups[so_row.po_number]['qty'] += so_row.qty
        po_groups[so_row.po_number]['amount'] += float(so_row.amount or 0.0)

    # ── Data rows ───────────────────────────────────────────────────────
    r = 2
    for po, info in po_groups.items():
        dates = po_dates.get(po, {})

        data_cell(ws, r, _COL_MARKETPLACE, result.marketplace, align='center')
        data_cell(ws, r, _COL_PO, po, align='center')
        data_cell(ws, r, _COL_LOCATION, info['location'], align='left')
        # State Name — blank by design.
        data_cell(ws, r, _COL_STATE, '', align='center')
        data_cell(ws, r, _COL_PO_DATE, dates.get('po_date', ''), align='center')
        data_cell(ws, r, _COL_EXP_DATE, dates.get('exp_date', ''), align='center')
        # PO Aging For Exp — blank by design.
        data_cell(ws, r, _COL_AGING, '', align='center')
        data_cell(ws, r, _COL_ORDER_VALUE, info['amount'],
                   number_format=_INR_FORMAT, align='right')
        data_cell(ws, r, _COL_ORDER_QTY, info['qty'], align='center')

        r += 1

    auto_width(ws)


# ── Helpers ────────────────────────────────────────────────────────────

def _build_po_date_lookup(result: ProcessingResult) -> Dict[str, dict]:
    """
    Build ``{po_number: {'po_date': str, 'exp_date': str}}`` from the raw
    punch DataFrame.

    The PO key is coerced with the same rules the engine uses for
    ``SORow.po_number`` (:meth:`MarketplaceEngine._coerce_po_to_str`) so
    the keys line up with ``result.rows``. Dates are formatted
    ``dd-mm-yyyy`` to match the master tracker. The first non-empty value
    seen for each PO wins (every row of a PO carries the same PO/expiry
    date on Zepto dumps).

    Returns an empty dict when the raw DataFrame or the PO column is
    unavailable — callers then render blank date cells.
    """
    df = result.raw_df
    if df is None or getattr(df, 'empty', True):
        return {}

    cfg = result.resolved_config or {}
    po_col = cfg.get('po_col', 'PO No.')
    if not isinstance(po_col, str) or po_col not in df.columns:
        return {}

    po_date_col = _find_col(df, _PO_DATE_CANDIDATES)
    exp_date_col = _find_col(df, _EXP_DATE_CANDIDATES)

    lookup: Dict[str, dict] = {}
    for _, raw_row in df.iterrows():
        po = _coerce_po(raw_row[po_col])
        if not po:
            continue
        entry = lookup.setdefault(po, {'po_date': '', 'exp_date': ''})
        if po_date_col and not entry['po_date']:
            entry['po_date'] = _fmt_date(raw_row[po_date_col])
        if exp_date_col and not entry['exp_date']:
            entry['exp_date'] = _fmt_date(raw_row[exp_date_col])

    return lookup


def _find_col(df: pd.DataFrame, candidates) -> Optional[str]:
    """
    Return the actual DataFrame column name matching the first candidate
    found, comparing case-insensitively with internal whitespace runs
    collapsed. ``None`` when no candidate is present.
    """
    norm = {' '.join(str(c).split()).lower(): c for c in df.columns}
    for cand in candidates:
        key = ' '.join(str(cand).split()).lower()
        if key in norm:
            return norm[key]
    return None


def _coerce_po(po_raw) -> str:
    """
    Mirror of ``MarketplaceEngine._coerce_po_to_str`` — clean PO string,
    int-coerced for whole numbers so keys match ``SORow.po_number``.
    """
    if pd.isna(po_raw):
        return ''
    if isinstance(po_raw, int) or (
        isinstance(po_raw, float) and po_raw.is_integer()
    ):
        return str(int(po_raw))
    return str(po_raw).strip()


def _fmt_date(val) -> str:
    """
    Format a raw date cell as ``dd-mm-yyyy``. Pandas Timestamps and
    datetime-likes are formatted; blank/NaN becomes ''; anything else is
    passed through as a stripped string (already-formatted text dates).
    """
    if pd.isna(val):
        return ''
    if isinstance(val, pd.Timestamp):
        return val.strftime('%d-%m-%Y')
    # datetime.datetime / date also expose strftime.
    if hasattr(val, 'strftime'):
        try:
            return val.strftime('%d-%m-%Y')
        except (ValueError, TypeError):
            pass
    return str(val).strip()
