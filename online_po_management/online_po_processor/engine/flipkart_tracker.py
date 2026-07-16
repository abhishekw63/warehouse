"""
engine.flipkart_tracker
=======================

Build the **Flipkart Tracker** rows from the portal's PO-list "header file"
(``purchase-orders-*.csv``) — the per-PO pivot the ops team pastes into the
master tracker. Flipkart's individual ``purchase_order_*.xlsx`` files (used for
the SO) carry the line items but NOT the per-PO Order Value / Order Qty / dates;
the header CSV does. So the operator optionally uploads it alongside the POs and
we generate the Tracker from it.

This mirrors the original Marketplace_Automation ``flipkart.py`` approach
(read CSV → map columns → assign Market Place by location → summarise), with the
current taxonomy/labels.

Market Place is assigned from the PO's ``Origin Warehouse`` via a LOCKED
location→marketplace mapping (below). The CSV has no marketplace field, so this
mapping is the source of truth. A location NOT in the map is never left blank —
it gets ``'FK (review)'`` + a warning so the operator can add it.

Column mapping (CSV → Tracker)
------------------------------
    Purchase Order ID      → PO
    Origin Warehouse       → Location
    Order Date             → PO Date
    Expiry Date            → Exp Date
    (computed)             → PO Aging For Exp  (= Exp − today, in days)
    Total Amount           → Order Value
    Total Ordered Quantity → Order Qty
    (locked map)           → Market Place
"""
from __future__ import annotations

import datetime as _dt
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd


# ── LOCKED location → Market Place mapping ────────────────────────────────
# Source of truth (operator-confirmed 2026-06-19). Keyed by the exact
# ``Origin Warehouse`` string. Anything not here → '_REVIEW' (see below).
_FK = 'FK'
_FK_HYPERLOCAL = 'FK Hyperlocal'
_FK_GROCERY = 'FK Grocery'
_REVIEW = 'FK (review)'      # unknown location — never blank; flagged for review

LOCATION_MARKETPLACE: Dict[str, str] = {
    'new_new_wh_nl_01nl': _FK_HYPERLOCAL,
    'ane_gsh_wh_nl_01nl': _FK_HYPERLOCAL,
    'bal_gsh_wh_nl_01nl': _FK_HYPERLOCAL,
    'ben_hos_wh_nl_02nl': _FK_HYPERLOCAL,
    'bhi_pad_wh_nl_04nl': _FK,                 # NB: _04nl = FK …
    'bhi_pad_wh_nl_05nl': _FK_HYPERLOCAL,      # … _05nl = FK Hyperlocal
     'bhu_men_wh_g_01':    _FK_GROCERY,         # operator-confirmed 2026-06-23: Grocery
    'bin_sh_wh_nl_01nl':  _FK_HYPERLOCAL,
    # added 2026-06-25 (were missing → fell through to 'FK (review)')
    'che_gsh_wh_nl_01nl': _FK_HYPERLOCAL,      # Chennai
    'coi_app_wh_g_01':    _FK_GROCERY,         # Coimbatore — _wh_g_ = Grocery
    'guw_gsh_wh_nl_01nl': _FK_HYPERLOCAL,      # Guwahati
    'jai_sh_wh_nl_01nl':  _FK_HYPERLOCAL,      # Jaipur
    'hyd_gsh_wh_nl_01nl': _FK_HYPERLOCAL,
    'kol_gsh_wh_nl_01nl': _FK_HYPERLOCAL,
    'luc_gsh_wh_nl_01nl': _FK_HYPERLOCAL,
    'lud_gsh_wh_nl_01nl': _FK_HYPERLOCAL,      # added 2026-06-23 (was missing)
    'pat_sh_wh_nl_01nl':  _FK_HYPERLOCAL,
    'pun_dhl_wh_nl_01nl': _FK_HYPERLOCAL,      # added 2026-07-14: Maval/Pune (Hyperlocal)
    'sai_gsh_wh_nl_01nl': _FK_HYPERLOCAL,
    'son_gsh_wh_nl_01nl': _FK_HYPERLOCAL,
    'ulu_sh_wh_nl_01nl':  _FK_HYPERLOCAL,      # operator: Hyperlocal (not Grocery)
    'ahm_sh_wh_nl_02nl':  _FK_HYPERLOCAL,
}

# Tracker column order (matches the master tracker paste layout).
TRACKER_COLUMNS = [
    'Market Place', 'PO', 'Location', 'PO Date', 'Exp Date',
    'PO Aging For Exp', 'Order Value', 'Order Qty',
]


def marketplace_for_location(location: str) -> str:
    """Locked location → Market Place. Unknown → 'FK (review)' (never blank)."""
    return LOCATION_MARKETPLACE.get(str(location or '').strip(), _REVIEW)


def _fmt_date(val) -> str:
    """ISO-8601 (with tz) → 'DD-MM-YYYY'; blank-safe."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    try:
        d = pd.to_datetime(val, format='ISO8601', errors='coerce')
        if pd.isna(d):
            d = pd.to_datetime(val, errors='coerce', dayfirst=True)
        return '' if pd.isna(d) else d.tz_localize(None).strftime('%d-%m-%Y') \
            if d.tzinfo else d.strftime('%d-%m-%Y')
    except Exception:  # noqa: BLE001
        return str(val)


def _aging_days(exp_str: str, today: _dt.date) -> Optional[int]:
    """Exp − today, in days (negative = already expired). None if unparseable."""
    if not exp_str:
        return None
    try:
        exp = _dt.datetime.strptime(exp_str, '%d-%m-%Y').date()
        return (exp - today).days
    except ValueError:
        return None


def build_flipkart_tracker(header_csv: str | Path,
                           today: Optional[_dt.date] = None) -> List[dict]:
    """
    Read the Flipkart header CSV and return one tracker row per PO.

    Each row is a dict keyed by :data:`TRACKER_COLUMNS`. ``Market Place`` is
    resolved from the locked location map (unknown → 'FK (review)'). Raises
    on a missing/unreadable file or absent required columns — surfaced rather
    than producing a silently-wrong tracker.
    """
    header_csv = Path(header_csv)
    df = pd.read_csv(header_csv, dtype=str)

    required = ['Purchase Order ID', 'Origin Warehouse', 'Order Date',
                'Expiry Date', 'Total Amount', 'Total Ordered Quantity']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"{header_csv.name}: not a Flipkart header file — missing columns "
            f"{missing}. Expected the portal 'purchase-orders-*.csv' export.")

    today = today or _dt.date.today()
    rows: List[dict] = []
    for _, r in df.iterrows():
        po = str(r['Purchase Order ID']).strip()
        if not po or po.lower() == 'nan':
            continue
        loc = str(r['Origin Warehouse']).strip()
        po_date = _fmt_date(r['Order Date'])
        exp_date = _fmt_date(r['Expiry Date'])
        try:
            order_value = float(str(r['Total Amount']).replace(',', ''))
        except (ValueError, TypeError):
            order_value = None
        try:
            order_qty = int(float(str(r['Total Ordered Quantity'])))
        except (ValueError, TypeError):
            order_qty = None
        rows.append({
            'Market Place':     marketplace_for_location(loc),
            'PO':               po,
            'Location':         loc,
            'PO Date':          po_date,
            'Exp Date':         exp_date,
            'PO Aging For Exp': _aging_days(exp_date, today),
            'Order Value':      order_value,
            'Order Qty':        order_qty,
        })
    # Sort by Location (mirrors the original tool), keeping the paste tidy.
    rows.sort(key=lambda x: x['Location'])
    return rows


def unknown_locations(rows: List[dict]) -> List[str]:
    """Distinct locations that resolved to 'FK (review)' — for warnings."""
    seen = []
    for r in rows:
        if r['Market Place'] == _REVIEW and r['Location'] not in seen:
            seen.append(r['Location'])
    return seen
