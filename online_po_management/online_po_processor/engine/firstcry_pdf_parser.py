"""
FirstCry (Siddeshwari Trading / firstcry.com) PDF PO parser.

v2.3.1 — FirstCry is the SECOND PDF-source marketplace (after Avenue/
DMart) and produces a STANDARD Sales Order, not a Transfer Order.

Why a table parser (not a text parser like Avenue)
--------------------------------------------------
The Avenue parser token-parses free ``extract_text()`` output because
Avenue's PDF has no clean cell borders. FirstCry's PO is a fully
bordered table, so ``page.extract_tables()`` recovers the grid directly
and far more reliably. We map columns by HEADER NAME (not fixed index),
so the parser tolerates column reordering / width changes between
exports — only the header labels must stay recognisable.

PDF layout (from the reference PO ``pin212260609ity61c4``)
----------------------------------------------------------
* A header key/value block: ``PONO``, ``PO Date``, ``Vendor Name``,
  ``Delivered To`` (+ address + buyer ``GST Number``), ``Currency``.
* A line-item table with these columns (18):

    Sr No | FCID | HSN Code | Manufacturer | Style Code | Image |
    Product Name | Product Description | Brand Name | Color |
    MRP | Base Cost | Tax | Landing Rate |
    Billed Qty | Free Qty | Total Qty | Total Amount

* A footer ``Sub Total <billed> <free> <total> <amount>`` row.

Key column decisions (verified against Items_March master)
----------------------------------------------------------
* **EAN = the 'Manufacturer' column**, NOT 'Style Code'. Manufacturer
  holds the GTIN on every row and resolves in master; Style Code is
  unreliable (sometimes the FCID, sometimes blank, sometimes a
  different EAN).
* **Landing Rate = MRP × 0.60** to the paisa across the sample, so the
  engine's ``compare_basis='landing'`` formula (MRP × margin%) matches
  it exactly at ``default_margin=60``.
* **Qty = 'Total Qty'** (Billed + Free).

The parser emits the engine's synthetic-column convention
(``__po__`` / ``__loc__``) so the standard SO pipeline (column
resolution → EAN→master lookup → ship-to mapping → validation) runs
unchanged — identical bridge pattern to the Avenue parser.
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import pdfplumber


# ── Header regexes (run against extract_text() of the header block) ───
_PO_RE   = re.compile(r'PONO\s*:?\s*([A-Za-z0-9][A-Za-z0-9\-]*)', re.IGNORECASE)
_DATE_RE = re.compile(r'PO\s*Date\s*:?\s*(\d{2}[.\-/]\d{2}[.\-/]\d{4})',
                      re.IGNORECASE)
# 15-char Indian GSTIN.
_GSTIN_RE = re.compile(r'\b(\d{2}[0-9A-Z]{13})\b')
# 'Delivered To: <name>' up to the next label ('GST', 'Address', EOL).
_DELIVERED_TO_RE = re.compile(
    r'Delivered\s*To\s*:?\s*(.+?)\s*(?:GST|Address|Currency|$)',
    re.IGNORECASE)
# 'PO Expiry Date:- DD-MM-YYYY' (v2.3.1 — for the Tracker sheet).
_EXPIRY_RE = re.compile(
    r'PO\s*Expiry\s*Date[^0-9]*(\d{2}[.\-/]\d{2}[.\-/]\d{4})', re.IGNORECASE)
# Delivery state: the name immediately before ', India - <pincode>' in
# the Delivered-To address (e.g. '…, Karnataka, India - 562114').
_STATE_RE = re.compile(
    r'([A-Za-z][A-Za-z ]*?),\s*India\s*[-–]\s*\d{6}', re.IGNORECASE)


@dataclass
class FirstcryLineItem:
    """One line item from a FirstCry PO."""
    sr_no:        int
    ean:          str            # from the 'Manufacturer' column (GTIN)
    fcid:         str            # FirstCry internal id
    style_code:   str            # unreliable — kept for reference only
    hsn_code:     str
    description:  str
    mrp:          float
    base_cost:    float
    tax_pct:      float          # GST % (e.g. 18.0)
    landing_rate: float          # = MRP × margin% (post-GST landed cost)
    billed_qty:   int
    free_qty:     int
    total_qty:    int
    total_amount: float


@dataclass
class FirstcryPOHeader:
    """Header info from a FirstCry PO PDF."""
    po_number:      str = ''
    po_date:        str = ''
    po_expiry:      str = ''      # PO Expiry Date (Tracker sheet)
    delivered_to:   str = ''      # ship-to location key (→ Ship-To B2B)
    delivery_state: str = ''      # state in the Delivered-To address
    buyer_gst:      str = ''
    raw_header:     str = ''


@dataclass
class FirstcryPO:
    header: FirstcryPOHeader
    items:  List[FirstcryLineItem] = field(default_factory=list)
    footer_total_qty:    Optional[int]   = None
    footer_total_amount: Optional[float] = None


# ── number / text cleaning ────────────────────────────────────────────

def _clean_num(cell: Any) -> str:
    """
    Strip ALL whitespace from a table cell so values pdfplumber split
    across visual lines ('89061216\\n40779') rejoin ('8906121640779').
    """
    return re.sub(r'\s+', '', str(cell or ''))


def _clean_text(cell: Any) -> str:
    """Collapse internal whitespace/newlines to single spaces."""
    return re.sub(r'\s+', ' ', str(cell or '')).strip()


def _to_float(cell: Any) -> float:
    t = _clean_num(cell).replace(',', '')
    if not t or t == '-':
        return 0.0
    try:
        return float(t)
    except ValueError:
        return 0.0


def _to_int(cell: Any) -> int:
    t = _clean_num(cell).replace(',', '')
    if not t or t == '-':
        return 0
    try:
        return int(float(t))
    except ValueError:
        return 0


# ── Header ────────────────────────────────────────────────────────────

def _parse_header(text: str) -> FirstcryPOHeader:
    """Extract PO number / date / delivered-to / buyer GST from text."""
    header = FirstcryPOHeader(raw_header=text[:1500])
    flat = re.sub(r'[ \t]+', ' ', text)

    if m := _PO_RE.search(flat):
        header.po_number = m.group(1).strip()
    if m := _DATE_RE.search(flat):
        header.po_date = m.group(1)
    if m := _EXPIRY_RE.search(flat):
        header.po_expiry = m.group(1)

    # v2.3.1: delivery state for the Tracker sheet. Search AFTER the
    # 'Delivered To' marker so the vendor's own state (Gujarat, which
    # appears earlier) isn't picked up — the first '<State>, India -
    # <pin>' after it is the ship-to state. Blank when the address
    # doesn't follow that shape (left empty in the tracker, by design).
    _idx = text.find('Delivered To')
    if m := _STATE_RE.search(text[_idx:] if _idx >= 0 else text):
        header.delivery_state = m.group(1).strip()

    # 'Delivered To' line carries the ship-to location AND (usually) the
    # buyer GST on the same row. Capture the location name as the mapping
    # key; the operator's Ship-To B2B 'Del Location' for FirstCry should
    # match it (exact or via the mapping's fuzzy/substring tiers).
    for ln in text.splitlines():
        if re.search(r'Delivered\s*To', ln, re.IGNORECASE):
            if m := _DELIVERED_TO_RE.search(re.sub(r'[ \t]+', ' ', ln)):
                header.delivered_to = m.group(1).strip()
            # Buyer GST on the same line, if present.
            if g := _GSTIN_RE.search(ln):
                header.buyer_gst = g.group(1)
            break

    # Fallback: if 'Delivered To' wasn't isolatable, the document title
    # (first non-empty line) is the same trading-name in this PO format.
    if not header.delivered_to:
        for ln in text.splitlines():
            if ln.strip():
                header.delivered_to = ln.strip()
                break

    return header


# ── Line-item table ───────────────────────────────────────────────────

# Map a normalized header-cell label → our canonical field key. Labels
# are normalized by removing ALL whitespace first, because pdfplumber
# splits narrow headers across visual lines ('Manufact\nurer',
# 'Landin\ng Rate', 'Total\nQty', 'Base\nCost'). So needles here are
# whitespace-free and matched as substrings. Order matters: longer /
# more-specific needles first ('totalamount' before 'totalqty'; 'sr'
# last because it's the shortest).
_COL_PATTERNS: List[tuple] = [
    ('totalamount',  'total_amount'),
    ('totalqty',     'total_qty'),
    ('billed',       'billed_qty'),
    ('free',         'free_qty'),
    ('landing',      'landing_rate'),
    ('basecost',     'base_cost'),
    ('tax',          'tax_pct'),
    ('mrp',          'mrp'),
    ('manufacturer', 'ean'),         # ← the GTIN column
    ('style',        'style_code'),
    ('hsn',          'hsn_code'),
    ('fcid',         'fcid'),
    ('productname',  'description'),
    ('sr',           'sr_no'),
]


def _map_columns(header_row: List[Any]) -> Optional[Dict[str, int]]:
    """
    Build {canonical_field: column_index} from a table's header row.

    Returns None if the row doesn't look like the line-item header
    (must contain at least Manufacturer + MRP + Landing + Total Qty).
    """
    norm = [re.sub(r'\s+', '', str(c or '')).lower() for c in header_row]
    mapping: Dict[str, int] = {}
    for idx, label in enumerate(norm):
        if not label:
            continue
        for needle, field_key in _COL_PATTERNS:
            if field_key in mapping:
                continue
            if needle in label:
                mapping[field_key] = idx
                break
    # Require the columns we actually depend on.
    if all(k in mapping for k in ('ean', 'mrp', 'landing_rate', 'total_qty')):
        return mapping
    return None


def _parse_items(tables: List[List[List[Any]]]) -> List[FirstcryLineItem]:
    """
    Build FirstcryLineItems from EVERY line-item table.

    A multi-page FirstCry PO renders one table per page (the column
    header repeats on each), so we must process ALL tables whose header
    is recognised — not just the first — or page-2+ items are lost.
    Rows are de-duped by Sr No as a guard against pdfplumber occasionally
    returning overlapping table regions.
    """
    items: List[FirstcryLineItem] = []
    seen_sr: set = set()
    for table in tables:
        if not table or len(table) < 2:
            continue
        # Locate the header row within the first few rows of this table.
        col_map = None
        header_idx = 0
        for hi in range(min(3, len(table))):
            col_map = _map_columns(table[hi])
            if col_map:
                header_idx = hi
                break
        if not col_map:
            continue

        def cell(row, key, _cm=col_map):
            i = _cm.get(key)
            return row[i] if i is not None and i < len(row) else ''

        for row in table[header_idx + 1:]:
            sr_raw = _clean_num(cell(row, 'sr_no'))
            ean = _clean_num(cell(row, 'ean'))
            # Skip non-data rows (Sub Total footer, blank rows): a real
            # line has an integer Sr No AND a non-empty EAN.
            if not sr_raw.isdigit() or not ean:
                continue
            sr_no = int(sr_raw)
            if sr_no in seen_sr:
                continue
            seen_sr.add(sr_no)
            items.append(FirstcryLineItem(
                sr_no=sr_no,
                ean=ean,
                fcid=_clean_num(cell(row, 'fcid')),
                style_code=_clean_num(cell(row, 'style_code')),
                hsn_code=_clean_num(cell(row, 'hsn_code')),
                description=_clean_text(cell(row, 'description')),
                mrp=_to_float(cell(row, 'mrp')),
                base_cost=_to_float(cell(row, 'base_cost')),
                tax_pct=_to_float(cell(row, 'tax_pct')),
                landing_rate=_to_float(cell(row, 'landing_rate')),
                billed_qty=_to_int(cell(row, 'billed_qty')),
                free_qty=_to_int(cell(row, 'free_qty')),
                total_qty=_to_int(cell(row, 'total_qty')),
                total_amount=_to_float(cell(row, 'total_amount')),
            ))
    return items


def _parse_footer(text: str) -> tuple:
    """Find 'Sub Total <billed> <free> <total> <amount>' for cross-check."""
    for ln in text.splitlines():
        m = re.match(
            r'^\s*Sub\s*Total\s+(\d+)\s+(\d+)\s+(\d+)\s+'
            r'([\d,]+(?:\.\d+)?)\s*$', ln, re.IGNORECASE)
        if m:
            return int(m.group(3)), _to_float(m.group(4))
    return None, None


# ── Public entry points ───────────────────────────────────────────────

def parse_firstcry_pdf(filepath: str | Path) -> FirstcryPO:
    """Parse a FirstCry PO PDF into a FirstcryPO."""
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(filepath)

    with pdfplumber.open(filepath) as pdf:
        text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
        tables: List[List[List[Any]]] = []
        for page in pdf.pages:
            tables.extend(page.extract_tables() or [])

    header = _parse_header(text)
    items = _parse_items(tables)
    if not items:
        raise ValueError(
            f"{filepath.name}: no FirstCry line items found. The table "
            f"header (Manufacturer / MRP / Landing Rate / Total Qty) "
            f"wasn't recognised — the PDF layout may differ from the "
            f"reference. Inspect page.extract_tables() output to tune "
            f"_COL_PATTERNS."
        )
    f_qty, f_amt = _parse_footer(text)
    return FirstcryPO(
        header=header, items=items,
        footer_total_qty=f_qty, footer_total_amount=f_amt,
    )


def firstcry_po_to_dataframe(po: FirstcryPO):
    """
    Convert a FirstcryPO into the flat DataFrame the engine consumes.

    Uses the ``__po__`` / ``__loc__`` synthetic-column convention so the
    standard SO pipeline runs unchanged. The FirstCry config references
    these names via ``po_col`` / ``loc_col`` / ``qty_col`` / ``ean_col``
    (= 'Manufacturer') / ``fob_col`` (= 'Landing Rate') / ``amount_col``
    / ``hsn_col``.
    """
    import pandas as pd

    rows = []
    for it in po.items:
        rows.append({
            '__po__':        po.header.po_number,
            '__loc__':       po.header.delivered_to,
            # v2.3.1: header fields replicated per row for the Tracker sheet.
            '__po_date__':   po.header.po_date,
            '__exp_date__':  po.header.po_expiry,
            '__state__':     po.header.delivery_state,
            'Sr No':         it.sr_no,
            'FCID':          it.fcid,
            'Manufacturer':  it.ean,          # ← ean_col
            'Style Code':    it.style_code,   # reference only
            'HSN Code':      it.hsn_code,
            'Product Name':  it.description,
            'MRP':           it.mrp,
            'Base Cost':     it.base_cost,
            'Tax':           it.tax_pct,
            'GST Rate':      it.tax_pct,      # alias for any rate-aware step
            'Landing Rate':  it.landing_rate,
            'Billed Qty':    it.billed_qty,
            'Free Qty':      it.free_qty,
            'Total Qty':     it.total_qty,
            'Total Amount':  it.total_amount,
        })
    return pd.DataFrame(rows)


def load_firstcry_pdf_as_dataframe(filepath: str | Path):
    """
    One-shot: parse PDF → engine-ready DataFrame.

    Registered in ``marketplace_engine.PDF_PARSERS`` under 'firstcry';
    called when a config sets ``source_format='pdf'`` +
    ``pdf_parser='firstcry'``.
    """
    return firstcry_po_to_dataframe(parse_firstcry_pdf(filepath))
