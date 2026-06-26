"""
Myntra (Myntra Jabong India Pvt Ltd / MJIPL) PDF PO parser.

v2.4.0 — Myntra is the FIRST **dual-format** marketplace: it keeps its
historical Excel "punch" path (``source_format='excel'``) AND gains a
PDF path so the operator can feed *either* the dashboard Excel export
*or* the PO PDF Myntra emails. The engine routes by the uploaded file's
extension (``.pdf`` → this parser; everything else → the Excel path),
so neither format's behaviour changes. Myntra POs are standard Sales
Orders (not Transfer Orders).

Why a geometry parser (not ``extract_tables`` / text tokens)
------------------------------------------------------------
The Myntra PO table has clean *vertical* column rules but **no
horizontal row borders**, and on page 2+ the vertical rules don't even
render — so ``page.extract_tables()`` recovers only the header row, and
free ``extract_text()`` interleaves the wrapped multi-line article
names unpredictably. Two facts make a robust reconstruction possible:

* The 17 column x-boundaries are identical on every page (same PO
  template), so we derive them **once** from whichever page exposes the
  full vertical-rule grid and reuse them on the borderless pages.
* Every line item has exactly one **SKU Code** (col 0, e.g.
  ``RENEBLSH130474319``). We anchor a row band on each SKU's y-position
  (band = this SKU's top → the next SKU's top), then bucket every word
  in the band into a column by its x-centre. This rejoins values
  pdfplumber splits across visual lines — notably the 13-digit EAN,
  which renders as two fragments (``89044731`` + ``05960`` →
  ``8904473105960``).

PDF layout (reference PO ``MYNJ-RNEE090626-1``)
-----------------------------------------------
* A header key/value block: ``PO #``, ``PO Approved Date``,
  ``Estimated Shipment Date``, and a BILL-TO / SHIP-TO address block
  (Myntra ships to its own FC, so bill == ship).
* A 17-column line-item table:

    SKU Code | HSN Code | Vendor Article Name | Vendor Article Number |
    Color | Size | Style ID | Qty | Bis Certificate Number | MRP |
    List Price | Landed Price | CGST % | CGST Amt | SGST % | SGST Amt |
    Total (plus Taxes)

* A footer ``Total Quantity: <n>  Grand Total: <amount>``.

Key column decisions (mapped to the Myntra config)
--------------------------------------------------
* **EAN = 'Vendor Article Number'** — on Myntra POs this column carries
  the 13-digit GTIN. We emit it under the column name **'GTIN'** because
  that is the Myntra config's ``ean_col`` (``item_resolution='from_ean'``
  → looked up in the Items master). Verified against the master: the
  joined EANs resolve.
* **Landing Price = 'Landed Price'** (config ``fob_col`` /
  ``amount_col`` factor) and **List price = 'List Price'** (config
  ``ref_fob_col``). Order value = Landed × Qty, matching the Excel path.
* **Qty = 'Qty'** (config ``qty_col`` = 'Quantity').
* Real **PO dates** are injected (``__po_date__`` = PO Approved Date,
  ``__exp_date__`` = Estimated Shipment Date) — the Excel punch has no
  PO-date column, so the PDF path actually yields richer Tracker data.

The parser emits the engine's synthetic-column convention (``__po__`` /
``__loc__`` / ``__po_date__`` / ``__exp_date__`` / ``__state__``)
*alongside* the real config column names so the standard SO pipeline
(column resolution → EAN→master lookup → ship-to mapping → validation)
runs unchanged — the same bridge pattern as the FirstCry/Reliance
parsers.
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import pdfplumber


# ── Header regexes (run against extract_text() of the header block) ───
_PO_RE = re.compile(r'PO\s*#\s*:?\s*([A-Za-z0-9][A-Za-z0-9\-/]*)', re.IGNORECASE)
# PO Approved Date: 2026-06-09  (ISO).
_PO_DATE_RE = re.compile(
    r'PO\s*Approved\s*Date\s*:?\s*(\d{4}-\d{2}-\d{2})', re.IGNORECASE)
# Estimated Shipment Date: 24/07/2026  (day-first).
_EXP_DATE_RE = re.compile(
    r'Estimated\s*Ship(?:ment)?\s*Date\s*:?\s*'
    r'(\d{2}[/\-.]\d{2}[/\-.]\d{4})', re.IGNORECASE)
# A 6-digit Indian pincode — terminates a ship-to address line.
_PINCODE_RE = re.compile(r'^\d{6}$')
# A 2-letter uppercase state code (Bangalore,KA,560067).
_STATE_CODE_RE = re.compile(r'^[A-Z]{2}$')

# 2-letter state code → name (Tracker 'State Name'; currently display-off
# but captured for completeness / future use).
_STATE_CODES = {
    'AP': 'Andhra Pradesh', 'AR': 'Arunachal Pradesh', 'AS': 'Assam',
    'BR': 'Bihar', 'CG': 'Chhattisgarh', 'GA': 'Goa', 'GJ': 'Gujarat',
    'HR': 'Haryana', 'HP': 'Himachal Pradesh', 'JH': 'Jharkhand',
    'KA': 'Karnataka', 'KL': 'Kerala', 'MP': 'Madhya Pradesh',
    'MH': 'Maharashtra', 'MN': 'Manipur', 'ML': 'Meghalaya',
    'MZ': 'Mizoram', 'NL': 'Nagaland', 'OD': 'Odisha', 'OR': 'Odisha',
    'PB': 'Punjab', 'RJ': 'Rajasthan', 'SK': 'Sikkim', 'TN': 'Tamil Nadu',
    'TS': 'Telangana', 'TG': 'Telangana', 'TR': 'Tripura',
    'UP': 'Uttar Pradesh', 'UK': 'Uttarakhand', 'UA': 'Uttarakhand',
    'WB': 'West Bengal', 'DL': 'Delhi', 'JK': 'Jammu and Kashmir',
}

# Fields we MUST resolve for a usable line item. We gate page selection
# on these being mapped (not on a fixed column count) because Myntra
# renders TWO table layouts that share these left-hand columns but differ
# on the right: INTRA-state POs carry CGST%+CGSTAmt+SGST%+SGSTAmt (17
# cols), INTER-state POs a single IGST%+IGSTAmt (15 cols). 'list_price'
# is intentionally optional — it's reference-only (ref_fob).
_REQUIRED_FIELDS = ('sku_code', 'ean', 'qty', 'mrp', 'landed_price')
# Minimum column boundaries for a page to be considered the item grid —
# enough rules to reach 'Landed Price' (col 11 → 13 boundaries). Rejects
# footer/address mini-tables (≤5 boundaries) on continuation pages.
_MIN_GRID_BOUNDS = 13

# A line item's SKU Code (col 0): letters then digits, e.g.
# 'RENEBLSH130474319', 'RENECMPT97268797'. Whitespace is stripped before
# matching (pdfplumber never splits the SKU, but be defensive).
_SKU_RE = re.compile(r'^[A-Z]{2,}[A-Z0-9]*\d{4,}$')
# Footer marker that caps the last row band ('Total Quantity:' /
# 'Grand Total:' / 'Terms and conditions'). NB: must NOT include bare
# 'Total' — the table HEADER cell 'Total (plus Taxes)' sits above the
# first row and would otherwise truncate the whole page.
_FOOTER_RE = re.compile(r'^(Quantity|Grand|Terms)', re.IGNORECASE)


@dataclass
class MyntraLineItem:
    """One line item from a Myntra PO."""
    sku_code:     str            # Myntra internal SKU (reference only)
    ean:          str            # 'Vendor Article Number' (the GTIN)
    hsn_code:     str
    article_name: str
    color:        str
    size:         str
    style_id:     str
    qty:          int
    mrp:          float
    list_price:   float          # 'List Price' (→ ref_fob)
    landed_price: float          # 'Landed Price' (→ fob / amount factor)
    total_amount: float          # 'Total (plus Taxes)'


@dataclass
class MyntraPOHeader:
    """Header info from a Myntra PO PDF."""
    po_number:      str = ''
    po_date:        str = ''      # PO Approved Date (ISO)
    exp_date:       str = ''      # Estimated Shipment Date (day-first)
    location:       str = ''      # ship-to City (→ Ship-To B2B 'Del Location')
    delivery_state: str = ''      # state name from the ship-to code
    raw_header:     str = ''


@dataclass
class MyntraPO:
    header: MyntraPOHeader
    items:  List[MyntraLineItem] = field(default_factory=list)
    footer_total_qty:    Optional[int]   = None
    footer_total_amount: Optional[float] = None


# ── number / text cleaning ────────────────────────────────────────────

def _clean_num(cell: Any) -> str:
    """Strip ALL whitespace so values split across lines rejoin
    ('89044731' + '05960' → '8904473105960')."""
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


# ── Geometry: column boundaries + field mapping ───────────────────────

def _col_bounds(page) -> List[float]:
    """
    Column x-boundaries from a page's vertical rules.

    The table draws double rules (two edges ~1pt apart) at each column
    border; we cluster edges within 4pt into a single boundary. N+1
    boundaries → N columns. Returns [] when the page has no rules (the
    borderless continuation pages — the caller reuses the gridded page's
    boundaries for those).
    """
    xs = sorted(e['x0'] for e in page.vertical_edges)
    if not xs:
        return []
    bounds: List[float] = []
    cur = [xs[0]]
    for x in xs[1:]:
        if x - cur[-1] <= 4:
            cur.append(x)
        else:
            bounds.append(sum(cur) / len(cur))
            cur = [x]
    bounds.append(sum(cur) / len(cur))
    return bounds


def _assign_col(x_centre: float, bounds: List[float]) -> Optional[int]:
    """Column index whose [bound, next bound] span contains x_centre."""
    for i in range(len(bounds) - 1):
        if bounds[i] - 1 <= x_centre <= bounds[i + 1] + 1:
            return i
    return None


# Map a normalized (whitespace-free, lower-cased) header label →
# canonical field key. Order matters: more-specific needles first so
# 'articlenumber' wins over the 'article' in 'articlename', and
# 'totalqty'/'plustaxes' don't collide with bare 'total'.
_COL_PATTERNS: List[tuple] = [
    ('articlenumber', 'ean'),          # ← the GTIN column
    ('articlename',   'article_name'),
    ('skucode',       'sku_code'),
    ('hsn',           'hsn_code'),
    ('color',         'color'),
    ('colour',        'color'),
    ('size',          'size'),
    ('styleid',       'style_id'),
    ('plustaxes',     'total_amount'),
    ('totalamount',   'total_amount'),
    ('listprice',     'list_price'),
    ('landedprice',   'landed_price'),
    ('landingprice',  'landed_price'),
    ('mrp',           'mrp'),
    ('qty',           'qty'),
]


def _map_columns(page, bounds: List[float],
                 first_anchor_top: float) -> Dict[str, int]:
    """
    Build {canonical_field: column_index} from the header row.

    Header cells (like data cells) wrap across visual lines
    ('Vendor\\nArticle\\nNumber'), so we collect every word that sits in
    the band just above the first line item, bucket it into a column by
    x-centre, join per column, then needle-match the joined label. This
    tolerates column reordering between exports — only the labels must
    stay recognisable.
    """
    labels: Dict[int, List[tuple]] = {}
    for w in page.extract_words():
        top = w['top']
        # Header band: the wrapped header cells sit just above the first
        # line item. 60pt (was 45) because when the FIRST item's name wraps
        # tall, its SKU code (col 0, the anchor) sits lower — pushing the
        # header row up to ~50pt above the anchor (seen on real POs where the
        # first item has a long name). Anything extra pulled in is harmless:
        # _map_columns only assigns a column when a header NEEDLE matches, and
        # emails / bare long-digit strings are dropped below.
        if not (first_anchor_top - 60 <= top < first_anchor_top - 3):
            continue
        text = w['text']
        # Drop non-label pollution that can still fall in the band: emails
        # and bare long-digit strings (EANs / Style IDs wrapping up from
        # the first data row). Real header labels are short alpha words.
        if '@' in text or re.fullmatch(r'\d{5,}', text):
            continue
        ci = _assign_col((w['x0'] + w['x1']) / 2, bounds)
        if ci is not None:
            labels.setdefault(ci, []).append((top, w['x0'], text))

    mapping: Dict[str, int] = {}
    for ci, words in labels.items():
        words.sort()
        label = re.sub(r'\s+', '', ''.join(t for _, _, t in words)).lower()
        for needle, field_key in _COL_PATTERNS:
            if field_key in mapping:
                continue
            if needle in label:
                mapping[field_key] = ci
                break
    return mapping


def _first_anchor_top(page, bounds: List[float]) -> Optional[float]:
    """Top y of the first SKU-Code word (col 0) — the table's first row."""
    lo, hi = bounds[0], bounds[1]
    tops = [w['top'] for w in page.extract_words()
            if lo - 1 <= (w['x0'] + w['x1']) / 2 <= hi + 1
            and _SKU_RE.match(_clean_num(w['text']))]
    return min(tops) if tops else None


def _page_bottom(page) -> float:
    """Y of the footer/Terms block — data rows end above this."""
    tops = [w['top'] for w in page.extract_words()
            if _FOOTER_RE.match(w['text'])]
    return min(tops) if tops else 1e9


def _rows_from_page(page, bounds: List[float],
                    col_map: Dict[str, int]) -> List[Dict[str, str]]:
    """
    Reconstruct line-item rows from one page using SKU-anchored bands.

    Each row spans [this SKU's top, next SKU's top), capped at the
    footer. Words in the band are bucketed by column; the EAN column is
    *concatenated* (to rejoin its split fragments), all others are
    space-joined in reading order.
    """
    words = page.extract_words()
    bottom = _page_bottom(page)
    lo_x, hi_x = bounds[0], bounds[1]
    anchors = sorted(
        (w for w in words
         if lo_x - 1 <= (w['x0'] + w['x1']) / 2 <= hi_x + 1
         and _SKU_RE.match(_clean_num(w['text']))
         and w['top'] < bottom),
        key=lambda w: w['top'])

    ean_col = col_map.get('ean')
    rows: List[Dict[str, str]] = []
    for i, a in enumerate(anchors):
        band_lo = a['top'] - 4
        band_hi = min(anchors[i + 1]['top'] - 4 if i + 1 < len(anchors)
                      else 1e9, bottom)
        cells: Dict[int, List[tuple]] = {}
        for w in words:
            if band_lo <= w['top'] < band_hi:
                ci = _assign_col((w['x0'] + w['x1']) / 2, bounds)
                if ci is not None:
                    cells.setdefault(ci, []).append(
                        (w['top'], w['x0'], w['text']))
        row: Dict[int, str] = {}
        for ci, lst in cells.items():
            lst.sort()
            joiner = ''.join if ci == ean_col else ' '.join
            row[ci] = joiner(t for _, _, t in lst)
        # Re-key by canonical field for the caller.
        rows.append({fk: row.get(ci, '') for fk, ci in col_map.items()})
    return rows


def _rows_all_pages(pages, bounds: List[float],
                    col_map: Dict[str, int]) -> List[Dict[str, str]]:
    """
    Reconstruct rows across the WHOLE document on one global y-axis.

    v2.4.2 (cross-page wrap fix): a line item's cells can wrap onto the
    TOP of the next page — most damagingly the EAN, whose last-5 fragment
    renders as an "orphan" above the next page's first SKU anchor (e.g.
    PO MYNJ-RNEE160626-3's body-mist EAN ``89061216`` + ``48782`` →
    ``8906121648782``, the tail sitting at the top of page 2). Banding each
    page in isolation (the old :func:`_rows_from_page`) dropped that tail,
    truncating the last item on every page.

    Stitching the pages onto a single axis — ``global_top = top +
    Σ heights of earlier pages`` — makes the last item's band run from its
    anchor straight through the page break to the next anchor, so the orphan
    fragment falls inside the band and rejoins. Column x-boundaries are
    identical on every page (same template) so a single ``bounds`` applies.

    Header blocks sit ABOVE the first SKU anchor, so they're never inside
    any band (bands only span [anchor, next-anchor]); the global footer
    marker caps the final band.
    """
    # Stitch pages onto one axis. +1pt guard so a word at the very top of a
    # page never ties the previous page's bottom edge.
    gwords: List[dict] = []
    offset = 0.0
    for p in pages:
        for w in p.extract_words():
            gwords.append({'top': w['top'] + offset, 'x0': w['x0'],
                           'x1': w['x1'], 'text': w['text']})
        offset += p.height + 1.0

    lo_x, hi_x = bounds[0], bounds[1]

    # Global footer bottom (first 'Total Quantity:' / 'Grand Total:' /
    # 'Terms' anywhere). _FOOTER_RE deliberately excludes bare 'Total' (the
    # 'Total (plus Taxes)' header cell) so this doesn't truncate page 1.
    foot = [w['top'] for w in gwords if _FOOTER_RE.match(w['text'])]
    bottom = min(foot) if foot else 1e18

    anchors = sorted(
        (w for w in gwords
         if lo_x - 1 <= (w['x0'] + w['x1']) / 2 <= hi_x + 1
         and _SKU_RE.match(_clean_num(w['text']))
         and w['top'] < bottom),
        key=lambda w: w['top'])

    ean_col = col_map.get('ean')
    rows: List[Dict[str, str]] = []
    for i, a in enumerate(anchors):
        band_lo = a['top'] - 4
        band_hi = min(anchors[i + 1]['top'] - 4 if i + 1 < len(anchors)
                      else 1e18, bottom)
        cells: Dict[int, List[tuple]] = {}
        for w in gwords:
            if band_lo <= w['top'] < band_hi:
                ci = _assign_col((w['x0'] + w['x1']) / 2, bounds)
                if ci is not None:
                    cells.setdefault(ci, []).append(
                        (w['top'], w['x0'], w['text']))
        row: Dict[int, str] = {}
        for ci, lst in cells.items():
            lst.sort()
            joiner = ''.join if ci == ean_col else ' '.join
            row[ci] = joiner(t for _, _, t in lst)
        rows.append({fk: row.get(ci, '') for fk, ci in col_map.items()})
    return rows


# ── Header ────────────────────────────────────────────────────────────

def _parse_header(text: str) -> MyntraPOHeader:
    """Extract PO number / dates / ship-to city + state from page text."""
    header = MyntraPOHeader(raw_header=text[:1500])
    flat = re.sub(r'[ \t]+', ' ', text)

    if m := _PO_RE.search(flat):
        header.po_number = m.group(1).strip()
    if m := _PO_DATE_RE.search(flat):
        header.po_date = m.group(1)
    if m := _EXP_DATE_RE.search(flat):
        header.exp_date = m.group(1)

    header.location, header.delivery_state = _extract_location(text)
    return header


def _extract_location(text: str) -> tuple:
    """
    Recover the ship-to ``(location, state)`` from the address block.

    Myntra renders the address tail two ways, both terminating in a
    6-digit pincode:

        '…, <City>, <ST>, <pincode>'         (2-letter state code)
        '…, <Area>, <City>, <StateName>, <pincode>'   (full state name)

    The token right before the pincode is the state; if it's a 2-letter
    code the City is the token before that, otherwise the state name
    doubles as the best Ship-To B2B key (Myntra's 'Del Location' list
    carries both city names like 'Bangalore'/'Mumbai' and state names
    like 'Haryana'/'West bengal'). We scan from the BILL-TO / SHIP-TO
    marker and take the FIRST pincode-terminated line — Myntra bills and
    ships to its own FC (identical addresses) and the line comes before
    the vendor's own address, so this never picks up Renee's Ahmedabad
    address that renders lower down.
    """
    lines = text.splitlines()
    start = 0
    for i, ln in enumerate(lines):
        if re.search(r'(BILL|SHIP)\s*TO', ln, re.IGNORECASE):
            start = i + 1
            break

    for ln in lines[start:]:
        parts = [p.strip(' .\t') for p in ln.split(',')]
        for idx, p in enumerate(parts):
            if _PINCODE_RE.match(p) and idx >= 1:
                prev = parts[idx - 1]
                if _STATE_CODE_RE.match(prev) and idx >= 2:
                    city = parts[idx - 2]
                    return city, _STATE_CODES.get(prev, prev)
                # Full state name (or a single city,pincode) before pin.
                return prev, _STATE_CODES.get(prev.upper(), prev)
    return '', ''


# ── Public entry points ───────────────────────────────────────────────

def parse_myntra_pdf(filepath: str | Path) -> MyntraPO:
    """Parse a Myntra PO PDF into a MyntraPO."""
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(filepath)

    with pdfplumber.open(filepath) as pdf:
        pages = list(pdf.pages)
        text = '\n'.join(p.extract_text() or '' for p in pages)

        # Column geometry + field mapping come from the page that exposes
        # the line-item grid (page 1); reused on borderless continuation
        # pages. We accept the first page whose mapping resolves every
        # required field — NOT a fixed column count — so both the
        # intra-state (17-col CGST+SGST) and inter-state (15-col IGST)
        # layouts parse through the same code.
        bounds: List[float] = []
        col_map: Dict[str, int] = {}
        for p in pages:
            b = _col_bounds(p)
            if len(b) < _MIN_GRID_BOUNDS:
                continue
            fa = _first_anchor_top(p, b)
            if fa is None:
                continue
            cm = _map_columns(p, b, fa)
            if all(k in cm for k in _REQUIRED_FIELDS):
                bounds, col_map = b, cm
                break

        if not bounds or not col_map:
            raise ValueError(
                f"{filepath.name}: Myntra PO line-item grid not recognised "
                f"(need column rules + SKU Code / Vendor Article Number / "
                f"Qty / MRP / Landed Price headers). Mapped columns: "
                f"{sorted(col_map)}. Inspect page.vertical_edges / "
                f"extract_words() to retune.")

        # v2.4.2: reconstruct across ALL pages on one global y-axis so an
        # item whose cells wrap onto the next page's top (notably the EAN
        # tail) rejoins instead of being truncated. Replaces the per-page
        # loop (_rows_from_page), which dropped each page's last-item wrap.
        raw_rows = _rows_all_pages(pages, bounds, col_map)

    items: List[MyntraLineItem] = []
    for r in raw_rows:
        ean = _clean_num(r.get('ean', ''))
        sku = _clean_num(r.get('sku_code', ''))
        if not ean or not sku:
            continue
        items.append(MyntraLineItem(
            sku_code=sku,
            ean=ean,
            hsn_code=_clean_num(r.get('hsn_code', '')),
            article_name=_clean_text(r.get('article_name', '')),
            color=_clean_text(r.get('color', '')),
            size=_clean_text(r.get('size', '')),
            style_id=_clean_num(r.get('style_id', '')),
            qty=_to_int(r.get('qty', '')),
            mrp=_to_float(r.get('mrp', '')),
            list_price=_to_float(r.get('list_price', '')),
            landed_price=_to_float(r.get('landed_price', '')),
            total_amount=_to_float(r.get('total_amount', '')),
        ))

    if not items:
        raise ValueError(
            f"{filepath.name}: no Myntra line items found. The SKU-anchored "
            f"row reconstruction returned nothing — the PDF layout may "
            f"differ from the reference. Inspect extract_words() output.")

    header = _parse_header(text)
    f_qty, f_amt = _parse_footer(text)
    return MyntraPO(header=header, items=items,
                    footer_total_qty=f_qty, footer_total_amount=f_amt)


def _parse_footer(text: str) -> tuple:
    """Find 'Total Quantity: <n>  Grand Total: <amount>' for cross-check."""
    qty = amt = None
    if m := re.search(r'Total\s*Quantity\s*:?\s*([\d,]+)', text, re.IGNORECASE):
        qty = _to_int(m.group(1))
    if m := re.search(r'Grand\s*Total\s*:?\s*([\d,]+(?:\.\d+)?)',
                      text, re.IGNORECASE):
        amt = _to_float(m.group(1))
    return qty, amt


def myntra_po_to_dataframe(po: MyntraPO):
    """
    Convert a MyntraPO into the flat DataFrame the engine consumes.

    Emits the Myntra config's real column names — ``GTIN`` (ean_col),
    ``Quantity`` (qty_col), ``Landing Price`` (fob_col / amount factor),
    ``List price(FOB+Transport-Excise)`` (ref_fob_col), ``Mrp``,
    ``Location`` (loc_col) — so the SAME Myntra config drives both the
    Excel and PDF paths. The ``__po__`` / ``__po_date__`` / ``__exp_date__``
    / ``__state__`` synthetic columns feed the PO-resolution + Tracker
    pipeline (``po_col`` lists '__po__' first via the engine alias step;
    'PO' is provided too for belt-and-braces).
    """
    import pandas as pd

    rows = []
    for it in po.items:
        rows.append({
            '__po__':        po.header.po_number,
            'PO':            po.header.po_number,
            '__loc__':       po.header.location,
            'Location':      po.header.location,
            '__po_date__':   po.header.po_date,
            '__exp_date__':  po.header.exp_date,
            '__state__':     po.header.delivery_state,
            'SKU Code':      it.sku_code,
            'GTIN':          it.ean,            # ← ean_col
            'Vendor Article Number': it.ean,    # alias (some dumps key on this)
            'Vendor Article Name':   it.article_name,
            'HSN Code':      it.hsn_code,
            'Colour':        it.color,
            'Size':          it.size,
            'Style Id':      it.style_id,
            'Quantity':      it.qty,            # ← qty_col
            'Mrp':           it.mrp,
            'List price(FOB+Transport-Excise)': it.list_price,  # ← ref_fob_col
            'Landing Price': it.landed_price,   # ← fob_col / amount factor
            'Total':         it.total_amount,
        })
    return pd.DataFrame(rows)


def load_myntra_pdf_as_dataframe(filepath: str | Path):
    """
    One-shot: parse PDF → engine-ready DataFrame.

    Registered in ``marketplace_engine.PDF_PARSERS`` under 'myntra';
    invoked when a ``.pdf`` is uploaded for a marketplace whose config
    sets ``pdf_parser='myntra'`` (Myntra — dual-format, routed by the
    engine's extension check).
    """
    return myntra_po_to_dataframe(parse_myntra_pdf(filepath))
