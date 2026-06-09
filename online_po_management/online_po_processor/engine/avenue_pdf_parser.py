"""
Avenue/DMart Ready PDF PO parser.

Parses Avenue E-Commerce Ltd Purchase Order PDFs into a structured
header dict + list of line items.

PDF layout (verified against Renee_AE00.pdf):
    Page 1: Header block + table header + ~12 items
    Page 2: ~4 items + Total row + Amount in Words + Terms

Each line item occupies TWO consecutive text lines after pdfplumber's
extract_text():
    Primary:  Sr  EAN  HSN_p1  Description_p1  UOM  Qty  MRP  Basic
              CGST% SGST% CESS% IGST% UGST% Landed Total
    Secondary: Article  HSN_p2  Description_p2  Clot  CGSTV SGSTV
               CESSV IGSTV UGSTV

Header info (PO number, dates, vendor code, GST, ship-to) lives in the
first ~12 text lines.
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import pdfplumber


# ── Regex patterns ────────────────────────────────────────────────────
# EANs are 13-digit (occasionally 12 or 14, hence the 12-14 range to
# tolerate edge cases). Avenue article numbers are typically 9 digits.
_EAN_RE      = re.compile(r'^\d{12,14}$')
_ARTICLE_RE  = re.compile(r'^\d{6,12}$')
_NUMBER_RE   = re.compile(r'^[\d,]+(?:\.\d+)?$')
_PO_RE       = re.compile(r'PurchaseOrder\s+(\d+)', re.IGNORECASE)
_SLOC_RE     = re.compile(r'S\.?Loc\.?\s*-?\s*(\d+)', re.IGNORECASE)
_VENDOR_RE   = re.compile(r'Vendor\s*:\s*(\d+)', re.IGNORECASE)
_GST_RE      = re.compile(r'GST\s*#\s*([0-9A-Z]+)', re.IGNORECASE)
_DATE_RE     = re.compile(
    r'PurchaseOrderDate\s*:\s*(\d{2}[.\-/]\d{2}[.\-/]\d{4})', re.IGNORECASE)
_VALIDITY_RE = re.compile(
    r'POValidity\s*:\s*(\d{2}[.\-/]\d{2}[.\-/]\d{4})\s*to\s*'
    r'(\d{2}[.\-/]\d{2}[.\-/]\d{4})', re.IGNORECASE)


@dataclass
class AvenueLineItem:
    """One line item from an Avenue PO, fully merged from primary+secondary rows."""
    sr_no:        int
    ean:          str            # 13-digit GTIN
    article_no:   str            # Avenue's internal SKU code
    hsn_code:     str            # Combined HSN (8-digit, e.g. '33041000')
    description:  str            # Combined product description
    uom:          str            # Usually 'EA'
    po_qty:       int
    mrp:          float
    basic_price:  float          # Pre-tax unit price
    landed_price: float          # Post-tax unit price ( = MRP × 0.45 expected)
    total_value:  float          # Landed × Qty
    # GST rate fields — exactly one of (CGST+SGST) or IGST will be populated
    cgst_pct:     Optional[float] = None
    sgst_pct:     Optional[float] = None
    cess_pct:     Optional[float] = None
    igst_pct:     Optional[float] = None
    ugst_pct:     Optional[float] = None
    # Tax-value (₹) fields from the secondary row
    cgst_v:       Optional[float] = None
    sgst_v:       Optional[float] = None
    cess_v:       Optional[float] = None
    igst_v:       Optional[float] = None
    ugst_v:       Optional[float] = None


@dataclass
class AvenuePOHeader:
    """Header info extracted from the Avenue PO PDF."""
    po_number:        str = ''
    storage_location: str = ''
    po_date:          str = ''         # Original DD.MM.YYYY string
    validity_from:    str = ''
    validity_to:      str = ''
    vendor_code:      str = ''
    vendor_name:      str = ''
    buyer_gst:        str = ''
    vendor_gst:       str = ''
    ship_to_text:     str = ''         # Full multi-line ship-to block
    ship_to_pincode:  str = ''
    raw_header_lines: List[str] = field(default_factory=list)


@dataclass
class AvenuePO:
    """Top-level container: header + list of items."""
    header: AvenuePOHeader
    items:  List[AvenueLineItem] = field(default_factory=list)
    # Footer totals (for cross-validation against summed items)
    footer_total_qty:   Optional[float] = None
    footer_total_value: Optional[float] = None


def parse_avenue_pdf(filepath: str | Path) -> AvenuePO:
    """
    Parse an Avenue/DMart Ready PO PDF.

    Args:
        filepath: Path to the Avenue PO PDF file.

    Returns:
        AvenuePO with header and items.

    Raises:
        ValueError: If the file doesn't look like an Avenue PO.
    """
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(filepath)

    with pdfplumber.open(filepath) as pdf:
        text_pages = [page.extract_text() or '' for page in pdf.pages]

    full_text = '\n'.join(text_pages)
    lines = [ln.rstrip() for ln in full_text.split('\n')]

    # Sanity check — make sure this looks like an Avenue PO
    if not any('Avenue' in ln and 'E-Commerce' in ln for ln in lines[:6]):
        raise ValueError(
            f"{filepath.name} doesn't look like an Avenue PO "
            f"(no 'Avenue E-Commerce' in first 6 lines)"
        )

    header = _parse_header(lines)
    items  = _parse_line_items(lines)
    total_qty, total_value = _parse_footer_totals(lines)

    return AvenuePO(
        header=header,
        items=items,
        footer_total_qty=total_qty,
        footer_total_value=total_value,
    )


# ── Internal helpers ──────────────────────────────────────────────────

def _to_float(token: str) -> float:
    """Parse '1,889.99' → 1889.99. Returns 0.0 for blanks/dashes."""
    token = (token or '').strip()
    if not token or token == '-':
        return 0.0
    return float(token.replace(',', ''))


def _to_float_or_none(token: str) -> Optional[float]:
    """Same as _to_float but returns None for blanks/dashes (preserve absence)."""
    token = (token or '').strip()
    if not token or token == '-':
        return None
    try:
        return float(token.replace(',', ''))
    except ValueError:
        return None


def _parse_header(lines: List[str]) -> AvenuePOHeader:
    """Extract PO number, dates, vendor, GST, ship-to from header lines."""
    header = AvenuePOHeader()
    # Look in the first ~15 lines for header fields
    search_lines = lines[:15]
    header.raw_header_lines = search_lines.copy()
    blob = ' '.join(search_lines)

    if m := _PO_RE.search(blob):
        header.po_number = m.group(1)
    if m := _SLOC_RE.search(blob):
        header.storage_location = m.group(1)
    if m := _DATE_RE.search(blob):
        header.po_date = m.group(1)
    if m := _VALIDITY_RE.search(blob):
        header.validity_from = m.group(1)
        header.validity_to   = m.group(2)
    if m := _VENDOR_RE.search(blob):
        header.vendor_code = m.group(1)

    # Both GSTs appear in the header; the first one (vendor) is on a
    # line near 'GST#24...' and the second (buyer) is on the BillTo
    # block. We capture both.
    gst_matches = _GST_RE.findall(blob)
    if gst_matches:
        # In the Avenue layout, the FIRST GST# is the vendor's (Renee's),
        # and the second one (in the BillTo/ShipTo block) is the buyer's.
        header.vendor_gst = gst_matches[0]
        if len(gst_matches) >= 2:
            header.buyer_gst = gst_matches[1]

    # Vendor name — the line after 'Vendor:<code>'. pdfplumber emits
    # this with columns glued together (vendor info + BillTo + ShipTo
    # all on one line). Vendor's chunk ends at the next column's start,
    # which we detect by trimming at known landmark substrings.
    _COL_TERMINATORS = ('IFC-', 'IFC ', 'BillTo', 'ShipTo', 'ValidFr',
                          'ValidTo', 'FSSAI', 'Email')
    for i, ln in enumerate(search_lines):
        if 'Vendor:' in ln and i + 1 < len(search_lines):
            next_ln = search_lines[i + 1].strip()
            # Trim at the first column-boundary landmark
            for term in _COL_TERMINATORS:
                idx = next_ln.find(term)
                if idx > 0:
                    next_ln = next_ln[:idx].strip()
                    break
            if next_ln:
                header.vendor_name = next_ln
            break

    # Ship-to pincode — Avenue puts the ship-to pincode in the THIRD
    # column of the header table. After concatenating columns,
    # pdfplumber emits text like:
    #   '380009 Kharbav,Thane,,Maharashtra Kharbav,Thane,,Maharashtra'
    #   'EmailID:accounts@reneecosmetics.in 421302 421302'
    # The vendor's pincode (380009 — Gujarat) appears first, then
    # BillTo's, then ShipTo's (often the same as BillTo's). We capture
    # the LAST 6-digit number in the first ~12 header lines since the
    # ShipTo block is always rendered rightmost.
    pincode_matches = re.findall(r'\b(\d{6})\b', ' '.join(search_lines))
    if pincode_matches:
        header.ship_to_pincode = pincode_matches[-1]

    # Capture the ship-to text block. Avenue's ShipTo block contains
    # 'IFC- Kukse Bhiwandi' / 'IFC-KukseBhiwandi' / 'IFC- Kukse Bhiwandi'
    # (depending on how pdfplumber's text extraction joined the columns).
    #
    # We normalize to the CANONICAL form 'IFC- Kukse Bhiwandi' — with a
    # SPACE after the dash AND between Kukse and Bhiwandi — because that
    # is the EXACT string the operator's Ship-To B2B sheet uses as the
    # 'Del Location' key for this warehouse:
    #
    #     Party  Del Location           Cust No  Ship to
    #     -----  ---------------------  -------  --------
    #     Dmart  IFC- Kukse Bhiwandi    20001    20001_34
    #
    # Any deviation here (missing space, lowercased, etc.) and the
    # mapping lookup fails → engine emits cust_no / ship_to blanks.
    ship_label = ''
    for ln in search_lines:
        # Accept any of: 'IFC-KukseBhiwandi', 'IFC- KukseBhiwandi',
        # 'IFC-Kukse Bhiwandi', 'IFC- Kukse Bhiwandi', 'IFC Kukse Bhiwandi'.
        # The pattern is intentionally permissive — Avenue's PDF
        # renderer has been observed to emit any of these spacings
        # across exports, but every variant maps to the same warehouse.
        m = re.search(
            r'IFC[\s\-]+Kukse[\s]*Bhiwandi', ln, re.IGNORECASE)
        if m:
            # Canonicalize: 'IFC-' + ' ' + 'Kukse' + ' ' + 'Bhiwandi'
            # (matches Vishal's Ship-To B2B 'Del Location' exactly)
            ship_label = 'IFC- Kukse Bhiwandi'
            break
    header.ship_to_text = ship_label

    return header


def _parse_line_items(lines: List[str]) -> List[AvenueLineItem]:
    """
    Walk text lines, pairing up primary+secondary rows into items.

    A primary line starts with a small integer Sr No, has a 13-digit
    EAN as token 2. A secondary line starts with a 6-12 digit article
    number (NOT 13 digits, so it can't be confused with an EAN).
    """
    items: List[AvenueLineItem] = []
    i = 0
    last_sr = 0
    while i < len(lines):
        line = lines[i]
        primary = _try_parse_primary(line, expected_sr_min=last_sr + 1)
        if primary is None:
            i += 1
            continue

        # Look for the secondary row on the very next line
        if i + 1 < len(lines):
            secondary = _try_parse_secondary(lines[i + 1])
            if secondary is not None:
                item = _merge_primary_secondary(primary, secondary)
                items.append(item)
                last_sr = item.sr_no
                i += 2
                continue

        # Primary without secondary — still capture it (degraded mode)
        item = _merge_primary_secondary(primary, {})
        items.append(item)
        last_sr = item.sr_no
        i += 1

    return items


def _try_parse_primary(line: str, expected_sr_min: int = 1) -> Optional[Dict[str, Any]]:
    """
    Try to parse a line as a primary item row. Returns None if it
    doesn't match the pattern.
    """
    tokens = line.split()
    if len(tokens) < 13:
        return None

    # Token 0 must be small integer (Sr No, 1-99 covers all real POs)
    if not tokens[0].isdigit():
        return None
    sr_no = int(tokens[0])
    if not (1 <= sr_no <= 999):
        return None
    # Sr Nos should be monotonically increasing — sanity check that
    # we're not picking up some footer line that happens to start
    # with digits.
    if sr_no < expected_sr_min - 2:
        return None

    # Token 1 must be a 13-digit EAN
    if not _EAN_RE.match(tokens[1]):
        return None

    # Token 2 is HSN part 1 (4 digits typically)
    # Tokens 3..N include description words + UOM (EA) + Qty + ...
    # Find UOM ('EA' usually) — that's our column anchor
    uom_idx = None
    for j in range(3, min(8, len(tokens))):
        if tokens[j] == 'EA':
            uom_idx = j
            break
    if uom_idx is None:
        return None

    # After UOM: Qty MRP Basic [- - -] IGST% [-] Landed Total
    # Counting back from the end: the LAST 2 tokens are Landed + Total.
    # Before that: 5 GST percentages (CGST SGST CESS IGST UGST).
    # Before that: Basic, MRP, Qty.
    # That's 2 + 5 + 3 = 10 trailing tokens. After UOM.
    if len(tokens) - (uom_idx + 1) < 10:
        return None

    qty_idx     = uom_idx + 1
    mrp_idx     = uom_idx + 2
    basic_idx   = uom_idx + 3
    cgst_p_idx  = uom_idx + 4
    sgst_p_idx  = uom_idx + 5
    cess_p_idx  = uom_idx + 6
    igst_p_idx  = uom_idx + 7
    ugst_p_idx  = uom_idx + 8
    landed_idx  = uom_idx + 9
    total_idx   = uom_idx + 10

    try:
        return {
            'sr_no':         sr_no,
            'ean':           tokens[1],
            'hsn_p1':        tokens[2],
            'desc_p1':       ' '.join(tokens[3:uom_idx]),
            'uom':           tokens[uom_idx],
            'po_qty':        int(tokens[qty_idx]),
            'mrp':           _to_float(tokens[mrp_idx]),
            'basic_price':   _to_float(tokens[basic_idx]),
            'cgst_pct':      _to_float_or_none(tokens[cgst_p_idx]),
            'sgst_pct':      _to_float_or_none(tokens[sgst_p_idx]),
            'cess_pct':      _to_float_or_none(tokens[cess_p_idx]),
            'igst_pct':      _to_float_or_none(tokens[igst_p_idx]),
            'ugst_pct':      _to_float_or_none(tokens[ugst_p_idx]),
            'landed_price':  _to_float(tokens[landed_idx]),
            'total_value':   _to_float(tokens[total_idx]),
        }
    except (ValueError, IndexError):
        return None


def _try_parse_secondary(line: str) -> Optional[Dict[str, Any]]:
    """
    Try to parse a line as a secondary item row.

    Secondary format: Article HSN_p2 Desc_p2 [Clot=1.00] - - - IGSTV -
    """
    tokens = line.split()
    if len(tokens) < 5:
        return None

    # Token 0 is article number (6-12 digits, NOT 13 — else it'd be EAN)
    if not _ARTICLE_RE.match(tokens[0]):
        return None
    if len(tokens[0]) >= 13:
        # Could collide with another EAN; reject
        return None

    # Token 1 is HSN part 2 (4 digits)
    hsn_p2 = tokens[1] if len(tokens) > 1 and tokens[1].isdigit() else ''

    # Find the 'Clot' value (typically '1.00') — anchor to the right
    # of the description.
    clot_idx = None
    for j in range(2, len(tokens)):
        if tokens[j] == '1.00' or (tokens[j].endswith('.00')
                                    and len(tokens[j]) <= 5):
            clot_idx = j
            break

    if clot_idx is None:
        # Degraded: assume description is just token 2
        desc_p2 = tokens[2] if len(tokens) > 2 else ''
        gst_v_idx = len(tokens) - 2 if len(tokens) >= 5 else None
        return {
            'article_no': tokens[0],
            'hsn_p2':     hsn_p2,
            'desc_p2':    desc_p2,
            'igst_v':     _to_float_or_none(tokens[gst_v_idx])
                            if gst_v_idx is not None else None,
        }

    # Description is the concatenated tokens between hsn_p2 and clot
    desc_p2 = ' '.join(tokens[2:clot_idx])

    # After Clot: CGSTV SGSTV CESSV IGSTV UGSTV
    after_clot = tokens[clot_idx + 1:]
    cgst_v = _to_float_or_none(after_clot[0]) if len(after_clot) > 0 else None
    sgst_v = _to_float_or_none(after_clot[1]) if len(after_clot) > 1 else None
    cess_v = _to_float_or_none(after_clot[2]) if len(after_clot) > 2 else None
    igst_v = _to_float_or_none(after_clot[3]) if len(after_clot) > 3 else None
    ugst_v = _to_float_or_none(after_clot[4]) if len(after_clot) > 4 else None

    return {
        'article_no': tokens[0],
        'hsn_p2':     hsn_p2,
        'desc_p2':    desc_p2,
        'cgst_v':     cgst_v,
        'sgst_v':     sgst_v,
        'cess_v':     cess_v,
        'igst_v':     igst_v,
        'ugst_v':     ugst_v,
    }


def _merge_primary_secondary(primary: Dict[str, Any],
                                secondary: Dict[str, Any]) -> AvenueLineItem:
    """Combine the two-row item layout into a single AvenueLineItem."""
    # Combined HSN: 4-digit HSN + 4-digit suffix → 8-digit full code
    hsn_p1 = primary.get('hsn_p1', '')
    hsn_p2 = secondary.get('hsn_p2', '')
    hsn_combined = f"{hsn_p1}{hsn_p2}" if hsn_p2 else hsn_p1

    # Combined description: primary's chunk + secondary's chunk.
    # PDF text extraction tends to collapse spaces inside descriptions,
    # so we keep both halves joined.
    desc = f"{primary.get('desc_p1','').strip()} {secondary.get('desc_p2','').strip()}".strip()

    return AvenueLineItem(
        sr_no=primary['sr_no'],
        ean=primary['ean'],
        article_no=secondary.get('article_no', ''),
        hsn_code=hsn_combined,
        description=desc,
        uom=primary['uom'],
        po_qty=primary['po_qty'],
        mrp=primary['mrp'],
        basic_price=primary['basic_price'],
        landed_price=primary['landed_price'],
        total_value=primary['total_value'],
        cgst_pct=primary.get('cgst_pct'),
        sgst_pct=primary.get('sgst_pct'),
        cess_pct=primary.get('cess_pct'),
        igst_pct=primary.get('igst_pct'),
        ugst_pct=primary.get('ugst_pct'),
        cgst_v=secondary.get('cgst_v'),
        sgst_v=secondary.get('sgst_v'),
        cess_v=secondary.get('cess_v'),
        igst_v=secondary.get('igst_v'),
        ugst_v=secondary.get('ugst_v'),
    )


def _parse_footer_totals(lines: List[str]) -> tuple:
    """Find the 'Total <qty> <value>' line near the end of the PDF."""
    for ln in lines:
        # 'Total 568.00 143,846.98'
        m = re.match(
            r'^Total\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s*$', ln.strip())
        if m:
            return _to_float(m.group(1)), _to_float(m.group(2))
    return None, None


# ──────────────────────────────────────────────────────────────────────
# Engine-bridge: AvenuePO → pandas DataFrame matching the engine's
# expected column shape.
# ──────────────────────────────────────────────────────────────────────

def avenue_po_to_dataframe(po: AvenuePO):
    """
    Convert an AvenuePO into a flat DataFrame that the existing
    MarketplaceEngine can consume.

    The DataFrame uses the engine's synthetic-column convention
    (``__po__`` / ``__loc__``) — same pattern Reliance's pre_process
    hook uses to inject header-only fields into per-row columns. This
    keeps the engine's downstream logic (column resolution, ship-to
    mapping, validation) format-agnostic.

    Columns emitted:
        __po__              — replicated PO number on every row
        __loc__             — replicated ship-to label on every row
        Sr No, EAN, Article No, HSN Code, Description, UOM
        PO Qty, MRP, Basic Price, Landed Price, Total Value
        CGST %, SGST %, CESS %, IGST %, UGST %
        CGST V, SGST V, CESS V, IGST V, UGST V
        GST Rate            — coalesced effective rate (IGST OR CGST+SGST)

    The engine's Avenue config references these names via ``po_col``,
    ``loc_col``, ``qty_col``, ``ean_col``, ``fob_col``, ``amount_col``,
    ``hsn_col``.
    """
    import pandas as pd

    rows = []
    for it in po.items:
        # Effective GST rate: IGST for inter-state, sum(CGST,SGST)
        # for intra-state. Stored as a single number for any downstream
        # validation that cares.
        if it.igst_pct is not None:
            gst_rate = it.igst_pct
        else:
            gst_rate = (it.cgst_pct or 0) + (it.sgst_pct or 0)

        rows.append({
            '__po__':       po.header.po_number,
            '__loc__':      po.header.ship_to_text or 'Bhiwandi',
            'Sr No':        it.sr_no,
            'EAN':          it.ean,
            'Article No':   it.article_no,
            'HSN Code':     it.hsn_code,
            'Description':  it.description,
            'UOM':          it.uom,
            'PO Qty':       it.po_qty,
            'MRP':          it.mrp,
            'Basic Price':  it.basic_price,
            'Landed Price': it.landed_price,
            'Total Value':  it.total_value,
            'CGST %':       it.cgst_pct,
            'SGST %':       it.sgst_pct,
            'CESS %':       it.cess_pct,
            'IGST %':       it.igst_pct,
            'UGST %':       it.ugst_pct,
            'CGST V':       it.cgst_v,
            'SGST V':       it.sgst_v,
            'CESS V':       it.cess_v,
            'IGST V':       it.igst_v,
            'UGST V':       it.ugst_v,
            'GST Rate':     gst_rate,
        })

    return pd.DataFrame(rows)


def load_avenue_pdf_as_dataframe(filepath: str | Path):
    """
    One-shot: parse PDF → return engine-ready DataFrame.

    This is what the engine should call when ``source_format == 'pdf'``
    AND ``pdf_parser == 'avenue'``. Same return type as ``pd.read_excel``,
    so the engine's downstream code runs unchanged.
    """
    po = parse_avenue_pdf(filepath)
    return avenue_po_to_dataframe(po)