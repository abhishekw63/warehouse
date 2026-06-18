"""
Big Basket (Innovative Retail Concepts Pvt Ltd / Bigbasket) Excel PO parser.

v2.7 — Big Basket is the first marketplace whose Excel ISN'T a flat table:
each PO arrives as its OWN ``<PO>.xlsx`` (one sheet ``<PO>AUTO-PO``) with a
multi-row HEADER block (DC / warehouse / delivery addresses, supplier,
GSTINs, ``PO Number`` / ``PO Date`` / ``PO Expiry date``) ABOVE the
line-item table. So a plain ``pd.read_excel`` can't find the columns — this
parser locates the table header row, reads the lines, and pulls the PO
number, dates and warehouse out of the preamble.

Layout (reference ``39455894.xlsx``)
------------------------------------
* Preamble: 'Warehouse Address' label with the DC CODE on the next row
  (e.g. ``Ahmedabad-FV-FMCG-DC`` — this MATCHES the Bigbasket Ship-To B2B
  'Del Location' exactly), then 'Delivery Address', 'Supplier', GSTINs.
* A line carrying ``PO Number:IRA39455894``, ``PO Date:15/Jun/2026``,
  ``PO Expiry date:15/Jul/2026``.
* The line-item table header:
    S.No | HSN Code | SKU Code | Description | EAN/UPC Code |
    Case Quantity | Quantity | Basic Cost | SGST% | SGST | CGST% | CGST |
    IGST% | IGST | GST% | GST Amount

Key column decisions
--------------------
* **EAN = 'EAN/UPC Code'** (``item_resolution='from_ean'`` → Items master).
* **Qty = 'Quantity'**; **Basic Cost** is the per-unit PRE-GST cost (the
  ``fob_col`` validated against MRP × margin% ÷ GST).
* No MRP column — MRP/GST come from the master via the EAN.
* **Order value** = Basic Cost × Qty (pre-GST), grossed up by the PO's own
  ``GST%`` (``amount_is_pre_gst`` + ``gst_pct_col='GST%'``) → inc-GST.

The parser emits the engine's synthetic ``__po__`` / ``__loc__`` /
``__po_date__`` / ``__exp_date__`` columns (same bridge pattern as the
Reliance / Myntra parsers) so the standard SO pipeline runs unchanged.
"""
from __future__ import annotations
import re
from pathlib import Path
from typing import Optional

import pandas as pd


_PO_RE   = re.compile(r'PO\s*Number\s*:?\s*(\S+)', re.IGNORECASE)
_DATE_RE = re.compile(r'PO\s*Date\s*:?\s*([0-9]{1,2}[/\-][A-Za-z0-9]+[/\-][0-9]{2,4})',
                      re.IGNORECASE)
_EXP_RE  = re.compile(r'PO\s*Expiry\s*date\s*:?\s*([0-9]{1,2}[/\-][A-Za-z0-9]+[/\-][0-9]{2,4})',
                      re.IGNORECASE)


def _fmt_date(s: str) -> str:
    """'15/Jun/2026' → '2026-06-15' (blank-safe; passes odd strings through)."""
    s = (s or '').strip()
    if not s:
        return ''
    try:
        return str(pd.to_datetime(s, dayfirst=True).date())
    except Exception:
        return s


def parse_bigbasket_excel(filepath: str | Path) -> pd.DataFrame:
    """Parse a Big Basket ``<PO>.xlsx`` into the engine's flat DataFrame."""
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(filepath)

    raw = pd.read_excel(filepath, sheet_name=0, header=None)

    # ── Preamble: PO number / dates ──
    head_text = '\n'.join(
        ' '.join(str(v) for v in raw.iloc[i].tolist() if str(v) != 'nan')
        for i in range(min(20, len(raw))))
    m = _PO_RE.search(head_text)
    po_number = m.group(1).strip() if m else ''
    m = _DATE_RE.search(head_text)
    po_date = _fmt_date(m.group(1)) if m else ''
    m = _EXP_RE.search(head_text)
    exp_date = _fmt_date(m.group(1)) if m else ''

    # ── Warehouse code = the value on the row after the 'Warehouse Address'
    #    label. This string equals the Bigbasket Del Location exactly. ──
    warehouse = ''
    for i in range(len(raw)):
        if str(raw.iloc[i, 0]).strip().lower() == 'warehouse address':
            if i + 1 < len(raw):
                warehouse = str(raw.iloc[i + 1, 0]).strip()
            break

    # ── Line-item table: locate the header row (carries 'EAN/UPC Code') ──
    header_row: Optional[int] = None
    for i in range(len(raw)):
        if any('ean/upc' in str(v).lower() for v in raw.iloc[i].tolist()):
            header_row = i
            break
    if header_row is None:
        raise ValueError(
            f"{filepath.name}: Big Basket line-item table not found "
            f"(no 'EAN/UPC Code' header). Inspect the sheet layout.")

    df = pd.read_excel(filepath, sheet_name=0, header=header_row, dtype=str)
    df.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df.columns]
    df = df[df['EAN/UPC Code'].notna()].copy()
    # Drop any footer/total rows (EAN must be a 8-14 digit barcode).
    df = df[df['EAN/UPC Code'].astype(str).str.replace(r'\D', '', regex=True)
            .str.len().between(8, 14)].copy()

    if df.empty:
        raise ValueError(
            f"{filepath.name}: no Big Basket line items found.")

    # Synthetic columns for the SO pipeline (repeated per row).
    df['__po__'] = po_number
    df['__loc__'] = warehouse
    df['__po_date__'] = po_date
    df['__exp_date__'] = exp_date
    return df


def load_bigbasket_excel_as_dataframe(filepath: str | Path) -> pd.DataFrame:
    """One-shot: parse the Big Basket Excel → engine-ready DataFrame.

    Registered in ``marketplace_engine.PDF_PARSERS`` under 'bigbasket' and
    invoked when the config sets ``file_parser='bigbasket'``."""
    return parse_bigbasket_excel(filepath)
