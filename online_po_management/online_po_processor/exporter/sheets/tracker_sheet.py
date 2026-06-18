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

    1. Market Place        — display name (e.g. 'First Cry', 'Zepto')
    2. PO                  — PO number
    3. Location            — raw delivery location from the punch
    4. State Name          — delivery state (when the source carries it;
                             else blank)
    5. PO Date             — from the source's PO-date column
    6. Exp Date            — from the source's PO-expiry column
    7. PO Aging For Exp    — Exp Date − today, in days (negative = past
                             expiry); blank when Exp Date is unavailable
    8. Order Value         — Σ row amount across the PO (₹, 2 dp)
    9. Order Qty           — Σ qty across the PO

Scope
-----
Per-marketplace via ``_SUPPORTED``. A marketplace qualifies once its
source exposes the PO-date / PO-expiry (and optionally state) fields this
view needs:
  * **Zepto** — date columns come straight from the punch ('PO Date',
    'PO Expiry Date'); no state, so State stays blank.
  * **Firstcry** — the PDF parser injects ``__po_date__`` / ``__exp_date__``
    / ``__state__`` synthetic columns, so all three are populated.
:func:`write` is a no-op for any unsupported marketplace, so adding it to
the exporter's sheet list is safe everywhere — the sheet only appears for
a supported run.
"""

from __future__ import annotations
from datetime import date, datetime
from typing import Dict, Optional

import pandas as pd

from online_po_processor.config.constants import ORDER_SEGMENT
from online_po_processor.data.master_loader import MasterLoader
from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    auto_width, data_cell, hdr_cell,
)


# v2.4.0: based on the master "New PO format" tracker (Feb-June'26) — no
# 'State Name'; PO Aging blank (operator preference). A leading 'Segment'
# column ('OnlineB2B') groups these against future offline (GT) orders that
# will share the same tracker / history DB.
_HEADERS = [
    'Segment',
    'Market Place', 'PO', 'Location',
    'PO Date', 'Exp Date', 'PO Aging For Exp',
    'Order Value', 'Order Qty',
]

# 1-based column indices.
_COL_SEGMENT = 1
_COL_MARKETPLACE = 2
_COL_PO = 3
_COL_LOCATION = 4
_COL_PO_DATE = 5
_COL_EXP_DATE = 6
_COL_AGING = 7          # intentionally blank
_COL_ORDER_VALUE = 8
_COL_ORDER_QTY = 9

# Marketplaces the per-marketplace Tracker sheet is wired up for. (The
# Auto consolidated/standalone tracker is NOT gated by this — it tracks
# every processed marketplace via build_tracker_rows.)
_SUPPORTED = {'Zepto', 'Firstcry', 'Reliance', 'RK', 'Bigbasket', 'Blink',
              'Meesho-TO', 'Nykaa', 'Purplle', 'Swiggy', 'Myntra'}

# Our config key → the exact 'Market Place' label the master tracker uses
# (verified against New PO format.xlsx → Feb-June'26). Names not listed
# already match (RK, Zepto, Nykaa, Reliance, Myntra, BlinkMP).
_MARKETPLACE_DISPLAY = {
    'Firstcry':    'First Cry',
    'Blink':       'Blinkit',
    'Dmart':       'D Mart',
    'Flipkart':    'Flipkart Alpha',
    'Flipkart-TO': 'Flipkart Branch',
    'Meesho-TO':   'Meesho-SB',
    'Bigbasket':   'Big Basket',
}

# v2.3.1: we currently sync only the PO details, so State Name and PO
# Aging are emitted BLANK — but the wiring is kept (state parsed from the
# source, aging = Exp − today). Flip either flag to True to populate.
_FILL_STATE = False
_FILL_AGING = False

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
#   * '__po_date__' / '__exp_date__'  — PDF parsers (FirstCry, Reliance)
#   * 'PO Date' / 'PO Expiry Date'     — Zepto punch
#   * 'Order date' / 'Cancellation deadline' — RK punch (RK's expiry is
#     its cancellation deadline; PO date is the order date)
#   * 'order_date' / 'expiry_date'     — Blink (underscores; matched via _norm_col)
#   * 'Date'                           — Flipkart (LAST, so specific names win)
_PO_DATE_CANDIDATES = ['__po_date__', 'PO Date', 'Order date', 'Date']
_EXP_DATE_CANDIDATES = ['__exp_date__', 'PO Expiry Date', 'PO Expiry',
                        'Expiry Date', 'Cancellation deadline']
_STATE_CANDIDATES = ['__state__', 'State Name', 'State']


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

    ws = wb.create_sheet('Tracker')
    for col_idx, header in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header)

    for r, row in enumerate(build_tracker_rows(result), start=2):
        write_tracker_row(ws, r, row)

    auto_width(ws)


# ── Reusable row builder / writer (shared with the Auto consolidated
#    Tracker so both stay identical) ──────────────────────────────────────

def build_tracker_rows(result: ProcessingResult) -> list:
    """
    Build the per-PO tracker rows for a result — a list of dicts with the
    9 tracker fields. Reused by :func:`write` (per-marketplace sheet) and
    by the Auto-mode consolidated Tracker.

    NOT gated by ``_SUPPORTED`` — the consolidated tracker wants a row for
    every processed marketplace. PO Date / Exp Date / State are filled
    only when the source carries them (else blank, by the ``_FILL_*``
    flags / missing columns). Order Value is GST-inclusive for
    ``amount_is_pre_gst`` marketplaces (RK), matching every other
    marketplace's already-inclusive amount.

    One row per PO, in processing order (every row of a PO shares its
    delivery location, so it's captured from the first row seen).
    """
    po_dates = _build_po_date_lookup(result)
    market_place = _MARKETPLACE_DISPLAY.get(
        result.marketplace, result.marketplace)
    incl_gst = bool((result.resolved_config or {}).get('amount_is_pre_gst'))

    # v2.4.0: some sources carry a per-PO GRAND TOTAL on every line (e.g.
    # Nykaa's 'PO Amount' = the portal's exact PO value, GST-inclusive,
    # repeated on each line). When the config names that column via
    # ``po_total_col``, the tracker uses it VERBATIM as Order Value (taken
    # once per PO) instead of summing per-line amounts — so the tracker
    # matches the marketplace portal to the rupee.
    po_totals = _build_po_total_lookup(result)

    po_groups: Dict[str, dict] = {}
    for so_row in result.rows:
        if so_row.po_number not in po_groups:
            po_groups[so_row.po_number] = {
                'location': so_row.location, 'qty': 0, 'amount': 0.0}
        po_groups[so_row.po_number]['qty'] += so_row.qty
        amt = float(so_row.amount or 0.0)
        if incl_gst:
            amt *= MasterLoader.row_gst_divisor(so_row)
        po_groups[so_row.po_number]['amount'] += amt

    rows = []
    for po, info in po_groups.items():
        dates = po_dates.get(po, {})
        exp_str = dates.get('exp_date', '')
        # Per-PO total column wins when present (exact portal match).
        order_value = po_totals.get(po)
        if order_value is None:
            order_value = info['amount']
        rows.append({
            'segment': ORDER_SEGMENT,
            'market_place': market_place,
            'po': po,
            'location': info['location'],
            'state': dates.get('state', '') if _FILL_STATE else '',
            'po_date': dates.get('po_date', ''),
            'exp_date': exp_str,
            'aging': _aging_days(exp_str) if _FILL_AGING else '',
            'order_value': order_value,
            'order_qty': info['qty'],
        })
    return rows


def _build_po_total_lookup(result: ProcessingResult) -> Dict[str, float]:
    """Build ``{po_number: per-PO total}`` from the raw DataFrame when the
    config names a ``po_total_col`` (a column carrying the whole-PO value on
    every line, e.g. Nykaa's 'PO Amount'). First non-null value per PO wins
    (it's identical across a PO's lines). Empty dict when not configured or
    the column is missing — callers then fall back to the summed amount."""
    cfg = result.resolved_config or {}
    col = cfg.get('po_total_col')
    df = result.raw_df
    if not col or df is None or getattr(df, 'empty', True):
        return {}
    po_col = cfg.get('po_col', 'PO Number')
    if (not isinstance(po_col, str) or po_col not in df.columns
            or col not in df.columns):
        return {}
    out: Dict[str, float] = {}
    for _, raw_row in df.iterrows():
        po = _coerce_po(raw_row[po_col])
        if not po or po in out:
            continue
        val = raw_row[col]
        if pd.isna(val):
            continue
        try:
            out[po] = float(str(val).replace(',', '').strip())
        except (ValueError, TypeError):
            continue
    return out


# Excel number format for the date columns. Real DATE values are written
# (not text) so the cell is reformattable / sortable AND pastes into the
# master tracker as a date.
_DATE_FMT = 'DD-MM-YYYY'


def _coerce_date(v):
    """Return a ``date`` for date-like input (date / datetime / dd-mm-yyyy
    or ISO string), else ``None``. Lets the tracker write a real Excel
    date regardless of whether the source gave us an object or a string."""
    if v is None or v == '':
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ('%d-%m-%Y', '%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    try:
        ts = pd.to_datetime(s, errors='raise')
        if pd.notna(ts):
            return ts.date()
    except Exception:        # noqa: BLE001 — not a date
        pass
    return None


def _write_date_cell(ws, r: int, col: int, v) -> None:
    d = _coerce_date(v)
    if d is not None:
        data_cell(ws, r, col, d, number_format=_DATE_FMT, align='center')
    else:
        data_cell(ws, r, col, v or '', align='center')


def write_tracker_row(ws, r: int, row: dict) -> None:
    """Write one tracker row dict (from :func:`build_tracker_rows`) at
    sheet row ``r``. Shared so the per-marketplace and consolidated
    trackers render identically. PO Date / Exp Date are written as real
    Excel dates (``DD-MM-YYYY``)."""
    data_cell(ws, r, _COL_SEGMENT,
              row.get('segment', ORDER_SEGMENT), align='center')
    data_cell(ws, r, _COL_MARKETPLACE, row['market_place'], align='center')
    data_cell(ws, r, _COL_PO, row['po'], align='center')
    data_cell(ws, r, _COL_LOCATION, row['location'], align='left')
    _write_date_cell(ws, r, _COL_PO_DATE, row['po_date'])
    _write_date_cell(ws, r, _COL_EXP_DATE, row['exp_date'])
    data_cell(ws, r, _COL_AGING, row['aging'], align='center')
    data_cell(ws, r, _COL_ORDER_VALUE, row['order_value'],
              number_format=_INR_FORMAT, align='right')
    data_cell(ws, r, _COL_ORDER_QTY, row['order_qty'], align='center')


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

    # v2.4.0: prefer the config's explicit date columns (e.g. Nykaa's
    # 'PO Release Date' / 'Expiry Date') over the generic candidate lists,
    # so a source with non-standard date headers still fills the tracker.
    def _cfg_col(key):
        c = cfg.get(key)
        return c if (isinstance(c, str) and c in df.columns) else None
    po_date_col = _cfg_col('po_date_col') or _find_col(df, _PO_DATE_CANDIDATES)
    exp_date_col = _cfg_col('exp_date_col') or _find_col(df, _EXP_DATE_CANDIDATES)
    state_col = _find_col(df, _STATE_CANDIDATES)

    lookup: Dict[str, dict] = {}
    for _, raw_row in df.iterrows():
        po = _coerce_po(raw_row[po_col])
        if not po:
            continue
        entry = lookup.setdefault(
            po, {'po_date': '', 'exp_date': '', 'state': ''})
        if po_date_col and not entry['po_date']:
            entry['po_date'] = _fmt_date(raw_row[po_date_col])
        if exp_date_col and not entry['exp_date']:
            entry['exp_date'] = _fmt_date(raw_row[exp_date_col])
        if state_col and not entry['state']:
            v = raw_row[state_col]
            entry['state'] = '' if pd.isna(v) else str(v).strip()

    return lookup


def _aging_days(exp_str: str):
    """
    Days between the expiry date (``dd-mm-yyyy`` string) and today —
    Exp − today. Negative = past expiry. Returns '' when the string is
    blank or doesn't parse, so the cell renders empty rather than wrong.
    """
    if not exp_str:
        return ''
    try:
        exp = datetime.strptime(str(exp_str).strip(), '%d-%m-%Y').date()
    except (ValueError, TypeError):
        return ''
    return (exp - date.today()).days


def _norm_col(c) -> str:
    """Normalise a column name for matching: underscores → spaces, runs of
    whitespace collapsed, lower-cased. So ``order_date`` / ``Order Date`` /
    ``ORDER  DATE`` all compare equal — this is what lets Blink's
    ``order_date`` / ``expiry_date`` match the ``Order date`` / ``Expiry
    Date`` candidates."""
    return ' '.join(str(c).replace('_', ' ').split()).lower()


def _find_col(df: pd.DataFrame, candidates) -> Optional[str]:
    """
    Return the actual DataFrame column name matching the first candidate
    found, comparing case-insensitively with underscores treated as
    spaces and whitespace runs collapsed. ``None`` when no candidate is
    present.
    """
    norm = {_norm_col(c): c for c in df.columns}
    for cand in candidates:
        key = _norm_col(cand)
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
    Format a raw date cell as ``dd-mm-yyyy``.

    Handles every shape a marketplace throws at us:
      * pandas Timestamp / datetime  → strftime
      * day-first text ('09-06-2026', '09.06.2026') → kept day-first
      * ISO datetime strings with time/timezone ('2026-06-01 08:45:17+00:00',
        '2026-06-16T18:29:59Z' — Blink) → parsed and reduced to the date
    Blank / NaN / unparseable → '' (or the original string if it had
    content but no recognisable date).

    Returning a clean ``dd-mm-yyyy`` here matters twice: the tracker shows
    it, AND the history DB's DATE parser receives it (so Blink's dates land
    as real DATEs instead of NULL).
    """
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return ''
    if isinstance(val, pd.Timestamp):
        return val.strftime('%d-%m-%Y')
    if hasattr(val, 'strftime') and not isinstance(val, str):
        try:
            return val.strftime('%d-%m-%Y')
        except (ValueError, TypeError):
            pass

    s = str(val).strip()
    if not s:
        return ''
    # Day-first text formats FIRST (so '09-06-2026' isn't misread as
    # month-first by the generic parser below).
    for fmt in ('%d-%m-%Y', '%d.%m.%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%d-%m-%Y')
        except ValueError:
            continue
    # ISO-ish (year-first, optional time/zone) — let pandas parse it.
    try:
        ts = pd.to_datetime(s, errors='raise')
        if pd.notna(ts):
            return ts.strftime('%d-%m-%Y')
    except Exception:        # noqa: BLE001 — not a date we recognise
        pass
    return s
