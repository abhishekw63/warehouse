"""
Flipkart (Flipkart India Private Limited) PO Excel parser.

v2.7.x — Flipkart's new vendor portal emits ONE Excel per PO
(``purchase_order_<PO>.xlsx``), replacing the old TWO-STEP flow where a
separate desktop tool compiled many PO files into a single
``FL_DUMP_COMPILATION.xlsx`` that was then fed to this app. This parser
reads ONE such PO file directly; because the Flipkart config now sets a
``file_parser``, the engine's multi-file path (``process_multi`` +
``_supports_multi_file``) kicks in — so the operator drops ALL of the
day's ``purchase_order_*.xlsx`` files at once and they're compiled into
one combined SO batch in memory. No intermediate dump file, no xlwings.

Layout (reference ``purchase_order_FLS05608C31C.xlsx``)
-------------------------------------------------------
* Preamble: retailer header; a key/value row with ``Expiry Date`` /
  ``ORDER DATE`` (row ~4); supplier + retailer addresses with a
  ``Shipped To Address`` value (row ~9).
* A **TWO-ROW hierarchical** line-item header:

    top  (cols 0-7) : S. no. | HSN/SAC Code | Product ID | Quantity |
                      Approved Units | UOM | Pending Quantity | Product Details
    sub  (cols 7-29): Title | Ean | Brand | Color | … | Supplier Unit
                      Price | IGST Rate | … | Supplier MRP | Taxable
                      Value | Tax Amount | Total Amount

  (``Product Details`` in the top row is a GROUP label spanning the sub
  columns; that's why a plain ``pd.read_excel`` sees non-unique columns.)
* Data rows until a ``Total Quantity=`` footer.

Mapping → the dump columns the Flipkart config already consumes
---------------------------------------------------------------
The parser emits the SAME flat column names the historical dump used, so
the Flipkart config's ``po_col``/``loc_col``/``qty_col``/``ean_col``/
``fob_col``/``amount_col`` mapping is unchanged:

* **PO**          ← the FILENAME (``purchase_order_<PO>`` → ``<PO>``), NOT
                    a cell — the in-sheet "PURCHASE ORDER NO" is ignored
                    per the operator's rule (use the file name).
* **EAN**         ← ``Ean``  (``item_resolution='from_ean'`` → Items master).
* **Qty**         ← ``Quantity``.
* **COST PRICE**  ← ``Supplier Unit Price`` (= MRP × 77%; the landing the
                    config validates against, ``compare_basis='landing'``).
* **total_amount**← ``Total Amount`` (verbatim, the GST-inclusive row total).
* **MRP**         ← ``Supplier MRP`` (NEW — carried so the Validation sheet
                    shows the vendor MRP pair; the new portal exposes it).
* **FSN Code**    ← ``Product ID``;  **description** ← ``Title``.
* **Address**     ← ``Shipped To Address`` value (→ Ship-To B2B mapping).

It also injects synthetic ``__po_date__`` / ``__exp_date__`` (ORDER DATE /
Expiry Date) so the dates are available if Flipkart is later added to the
Tracker. Numeric cells carry a ``" (INR)"`` / ``" %"`` suffix which is
stripped.
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd


# Normalized header label → canonical dump field. Matched against BOTH
# header rows (the layout splits labels across two visual rows).
_LABELS = {
    'productid':         'fsn',
    'ean':               'ean',
    'quantity':          'qty',
    'supplierunitprice': 'cost',
    'totalamount':       'total_amount',
    'suppliermrp':       'mrp',
    'title':             'description',
}

# Fields we MUST resolve for a usable line item.
_REQUIRED = ('ean', 'qty', 'cost')


def _norm(s) -> str:
    """Whitespace-free, lower-cased label for matching."""
    return re.sub(r'\s+', '', str(s if s is not None else '')).lower()


def _clean_num(v) -> str:
    """Strip ' (INR)' / '%' / thousands separators from a numeric cell."""
    s = str(v if v is not None else '').strip()
    s = re.sub(r'\(?\s*INR\s*\)?', '', s, flags=re.IGNORECASE)
    return s.replace('%', '').replace(',', '').strip()


def _to_float(v) -> Optional[float]:
    s = _clean_num(v)
    try:
        return float(s) if s and s.lower() != 'nan' else None
    except ValueError:
        return None


def _to_int(v) -> Optional[int]:
    s = _clean_num(v)
    try:
        return int(float(s)) if s and s.lower() != 'nan' else None
    except ValueError:
        return None


def extract_po_number(filepath: str | Path) -> str:
    """``purchase_order_FLS05608C31C.xlsx`` → ``FLS05608C31C``.

    Strips the ``purchase_order_`` prefix (any case/sep) and a trailing
    ``' (1)'`` duplicate-download suffix; the in-sheet PO number is ignored
    on purpose (the operator keys on the file name)."""
    name = Path(filepath).stem
    m = re.match(r'(?i)^purchase[\s_-]*order[\s_-]*(.+)$', name)
    po = m.group(1) if m else name
    po = re.sub(r'\s*\(\d+\)\s*$', '', po)        # drop ' (1)' dup suffix
    return po.strip().upper()


def _fmt_date(s: str) -> str:
    """'17/06/2026' → '2026-06-17' (blank-safe; passes odd strings through)."""
    s = (s or '').strip()
    if not s:
        return ''
    try:
        return str(pd.to_datetime(s, dayfirst=True).date())
    except Exception:  # noqa: BLE001
        return s


def _value_after_label(raw: pd.DataFrame, needle: str) -> str:
    """First non-empty cell to the RIGHT of a cell containing ``needle``
    (case-insensitive substring). Used for the ship-to address and the
    in-row key/value date fields."""
    needle = needle.lower()
    for i in range(len(raw)):
        row = raw.iloc[i].tolist()
        for j, cell in enumerate(row):
            if cell is not None and needle in str(cell).lower():
                for k in range(j + 1, len(row)):
                    nx = row[k]
                    if nx is not None and str(nx).strip() and str(nx) != 'nan':
                        return str(nx).strip()
    return ''


def clean_address(address: str) -> str:
    """Minimal clean: keep from 'Flipkart India' if present, truncate to the
    LAST 6-digit pincode, collapse whitespace. Mirrors the legacy dump
    generator so the existing Address→ship-to matching is unchanged."""
    if not address:
        return ''
    addr = str(address)
    start = addr.lower().find('flipkart india')
    if start != -1:
        addr = addr[start:]
    pins = list(re.finditer(r'\b\d{6}\b', addr))
    if pins:
        addr = addr[:pins[-1].end()]
    return re.sub(r'\s+', ' ', addr).strip()


def _header_top_row(raw: pd.DataFrame) -> Optional[int]:
    """Index of the top header row — the one carrying 'Product ID'."""
    for i in range(len(raw)):
        if any(_norm(c) == 'productid' for c in raw.iloc[i].tolist()):
            return i
    return None


def _build_col_map(raw: pd.DataFrame, top: int) -> Dict[str, int]:
    """{canonical_field: column_index} from the two-row header (rows
    ``top`` and ``top+1``). First occurrence wins."""
    col_map: Dict[str, int] = {}
    for hr in (top, top + 1):
        if hr >= len(raw):
            break
        for j, cell in enumerate(raw.iloc[hr].tolist()):
            key = _LABELS.get(_norm(cell))
            if key and key not in col_map:
                col_map[key] = j
    return col_map


def parse_flipkart_po(filepath: str | Path) -> pd.DataFrame:
    """Parse ONE ``purchase_order_<PO>.xlsx`` into the engine's flat dump
    DataFrame (one row per line item)."""
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(filepath)

    raw = pd.read_excel(filepath, sheet_name=0, header=None, dtype=str)

    po_number = extract_po_number(filepath)
    po_date = _fmt_date(_value_after_label(raw, 'order date'))
    exp_date = _fmt_date(_value_after_label(raw, 'expiry date'))
    address = clean_address(_value_after_label(raw, 'shipped to address'))

    top = _header_top_row(raw)
    if top is None:
        raise ValueError(
            f"{filepath.name}: Flipkart line-item table not found "
            f"(no 'Product ID' header). Inspect the sheet layout.")

    col_map = _build_col_map(raw, top)
    missing = [k for k in _REQUIRED if k not in col_map]
    if missing:
        raise ValueError(
            f"{filepath.name}: Flipkart header missing required columns "
            f"{missing} (mapped: {sorted(col_map)}). The portal layout may "
            f"have changed — inspect rows {top}/{top + 1}.")

    today = datetime.today().strftime('%d-%m-%Y')
    rows: List[dict] = []
    for i in range(top + 2, len(raw)):
        cells = raw.iloc[i].tolist()
        c0 = _norm(cells[0]) if cells else ''
        # Footer markers cap the line-item table.
        if c0.startswith('totalquantity') or c0.startswith('total='):
            break

        def g(field: str):
            ci = col_map.get(field)
            return cells[ci] if (ci is not None and ci < len(cells)) else None

        ean = re.sub(r'\D', '', str(g('ean') or ''))
        if not (8 <= len(ean) <= 14):       # skip non-item / footer rows
            continue
        qty = _to_int(g('qty'))
        if not qty:
            continue

        rows.append({
            'Date':         today,
            'PO':           po_number,
            'FSN Code':     str(g('fsn') or '').strip(),
            'EAN':          ean,
            'Qty':          qty,
            'COST PRICE':   _to_float(g('cost')),
            'total_amount': _to_float(g('total_amount')),
            'MRP':          _to_float(g('mrp')),
            'description':  str(g('description') or '').strip(),
            'Address':      address,
            '__po_date__':  po_date,
            '__exp_date__': exp_date,
        })

    if not rows:
        raise ValueError(
            f"{filepath.name}: no Flipkart line items found below the header "
            f"(rows {top + 2}+). The PDF/Excel layout may differ from the "
            f"reference — inspect the sheet.")

    return pd.DataFrame(rows)


def load_flipkart_po_as_dataframe(filepath: str | Path) -> pd.DataFrame:
    """One-shot: parse the Flipkart PO Excel → engine-ready DataFrame.

    Registered in ``marketplace_engine.PDF_PARSERS`` under 'flipkart' and
    invoked when the Flipkart config sets ``file_parser='flipkart'``."""
    return parse_flipkart_po(filepath)
