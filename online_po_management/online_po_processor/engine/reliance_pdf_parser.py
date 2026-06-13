"""
Reliance Retail Limited PDF PO parser.

v2.3.1 — Reliance moved to PDF Purchase Orders ("PO. INTM_ <no> .PDF
Reliance Retail Limited.PDF"). This replaces the older Excel/pre_process
Reliance flow. Standard Sales Order output.

PDF layout (verified against PO 5000479071)
-------------------------------------------
* Page 1: header block — PO NO., Site, PO Date, vendor block, and the
  Delivery Address (the ship-to city lives here).
* Page 2 (and onward for big POs): the line-item table, then footer
  totals (Grand Total / TOTAL BASIC VALUE / TOTAL IGST / Total Order
  Value).
* Pages 3+: General Conditions of Purchase (no item tables) and a final
  "Annexure For Site Details" table — neither matches the item header,
  so they're ignored.

The line-item table (via pdfplumber.extract_tables()) STACKS several
fields into one column, separated by newlines::

    Sr.No | Article No.\\nHSN Code | EAN No.\\nVendor Article |
    Material Description | Quantity | UOM | MRP | Base Cost |
    IGST(%)\\nCESS(%)\\nCessFxdRt | IGST\\nCESS\\nCessFxdVl | Total Base Value

So the Article/HSN, EAN/VendorArticle and IGST(%) columns are split on
the newline and the relevant line taken.

Pricing
-------
``Base Cost`` is Reliance's PRE-GST taxable cost. The engine validates
it against MRP × (1 − 0.31×(1+GST)) ÷ (1+GST) — the GST-dependent
margin wired in the Reliance config (``compare_basis='cost'``,
``fob_col='Base Cost'``, ``gst_margin_discount=0.31``). Item No / MRP /
GST come from Items_March via the EAN (``from_ean``).

Emits the engine's synthetic ``__po__`` / ``__loc__`` columns (same
convention as the Avenue/FirstCry parsers) so the SO pipeline runs
format-agnostic.
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import pdfplumber


# ── Header regexes (run against page-1 extract_text) ──────────────────
_PO_RE   = re.compile(r'PO\s*NO\.?\s*:?\s*(\d+)', re.IGNORECASE)
_DATE_RE = re.compile(r'PO\s*Date\s*:?\s*(\d{2}[.\-/]\d{2}[.\-/]\d{4})',
                      re.IGNORECASE)
_SITE_RE = re.compile(r'Site\s*:?\s*([A-Z0-9]+)')
# Delivery city/state: 'CITY, State - 421302' within the Delivery Address
# block. Group 1 = city (ship-to key), 2 = state, 3 = pin.
_CITY_RE = re.compile(r'([A-Za-z][A-Za-z .]+?),\s*([A-Za-z ]+?)\s*-\s*(\d{6})')
# Reliance carries no 'PO Expiry'; the DELIVERY DATE is the delivery
# deadline the warehouse tracks, so it maps to the tracker's Exp Date.
_DELIVERY_DATE_RE = re.compile(
    r'DELIVERY\s*DATE\s*:?\s*(\d{2}[.\-/]\d{2}[.\-/]\d{4})', re.IGNORECASE)


@dataclass
class RelianceLineItem:
    sr_no:        int
    ean:          str
    article_no:   str
    hsn_code:     str
    description:  str
    qty:          float
    mrp:          float
    base_cost:    float       # PRE-GST taxable cost (the fob)
    igst_pct:     float
    total_base_value: float


@dataclass
class ReliancePOHeader:
    po_number:     str = ''
    po_date:       str = ''
    site:          str = ''
    delivery_city: str = ''   # ship-to key (→ Ship-To B2B)
    delivery_pin:  str = ''
    delivery_date: str = ''   # DELIVERY DATE → Tracker 'Exp Date'
    delivery_state: str = ''  # state in the Delivery Address (Tracker)


@dataclass
class ReliancePO:
    header: ReliancePOHeader
    items:  List[RelianceLineItem] = field(default_factory=list)


# ── number / text cleaning ────────────────────────────────────────────

def _first_line(cell: Any) -> str:
    """First newline-separated line of a stacked cell, stripped."""
    return str(cell or '').split('\n')[0].strip()


def _nth_line(cell: Any, n: int) -> str:
    parts = str(cell or '').split('\n')
    return parts[n].strip() if len(parts) > n else ''


def _to_float(token: Any) -> float:
    t = re.sub(r'[^\d.\-]', '', _first_line(token))
    if not t or t in ('-', '.'):
        return 0.0
    try:
        return float(t)
    except ValueError:
        return 0.0


def _clean_text(cell: Any) -> str:
    return re.sub(r'\s+', ' ', str(cell or '')).strip()


# ── Header ────────────────────────────────────────────────────────────

def _parse_header(text: str) -> ReliancePOHeader:
    h = ReliancePOHeader()
    flat = re.sub(r'[ \t]+', ' ', text)
    if m := _PO_RE.search(flat):
        h.po_number = m.group(1)
    if m := _DATE_RE.search(flat):
        h.po_date = m.group(1)
    if m := _SITE_RE.search(flat):
        h.site = m.group(1)
    if m := _DELIVERY_DATE_RE.search(flat):
        h.delivery_date = m.group(1)

    # Delivery city/state — search the 'Delivery Address' block so the
    # vendor's own city (AHMEDABAD) isn't picked up. The vendor address
    # uses 'Pin Code :' (no dash) so it wouldn't match _CITY_RE anyway,
    # but anchoring is safer.
    idx = text.find('Delivery Address')
    block = text[idx:idx + 400] if idx >= 0 else text
    if m := _CITY_RE.search(block):
        h.delivery_city = m.group(1).strip()
        h.delivery_state = m.group(2).strip()
        h.delivery_pin = m.group(3)
    return h


# ── Line-item table ───────────────────────────────────────────────────

def _map_columns(header_row: List[Any]) -> Optional[Dict[str, int]]:
    """
    Map the line-item header → {field: col_index}. Returns None if this
    table isn't the item table (needs MRP + Base Cost + Quantity).
    """
    norm = [re.sub(r'\s+', '', str(c or '')).lower() for c in header_row]
    m: Dict[str, int] = {}
    for i, label in enumerate(norm):
        if not label:
            continue
        if 'sr' in label and 'sr_no' not in m and label.startswith('sr'):
            m['sr_no'] = i
        if 'articleno' in label or ('article' in label and 'hsn' in label):
            m.setdefault('article_hsn', i)
        if label.startswith('eanno') or label.startswith('ean'):
            m.setdefault('ean', i)
        if 'materialdescription' in label or 'description' in label:
            m.setdefault('description', i)
        if label == 'quantity' or label.startswith('quantity'):
            m.setdefault('qty', i)
        if label == 'mrp':
            m.setdefault('mrp', i)
        if 'basecost' in label:
            m.setdefault('base_cost', i)
        if 'igst' in label and '%' in label:
            m.setdefault('igst_pct', i)
        if 'totalbasevalue' in label:
            m.setdefault('total_base_value', i)
    if all(k in m for k in ('mrp', 'base_cost', 'qty', 'ean')):
        return m
    return None


def _parse_items(tables: List[List[List[Any]]]) -> List[RelianceLineItem]:
    items: List[RelianceLineItem] = []
    seen: set = set()
    for table in tables:
        if not table or len(table) < 2:
            continue
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
            sr_raw = _first_line(cell(row, 'sr_no'))
            ean = _first_line(cell(row, 'ean'))
            # Real line: integer Sr.No AND a non-empty EAN. Footer rows
            # ('Grand Total of Qty', 'TOTAL BASIC VALUE', …) fail this.
            if not sr_raw.isdigit() or not ean:
                continue
            sr_no = int(sr_raw)
            if sr_no in seen:
                continue
            seen.add(sr_no)
            art_hsn = cell(row, 'article_hsn')
            items.append(RelianceLineItem(
                sr_no=sr_no,
                ean=ean,
                article_no=_nth_line(art_hsn, 0),
                hsn_code=_nth_line(art_hsn, 1),
                description=_clean_text(cell(row, 'description')),
                qty=_to_float(cell(row, 'qty')),
                mrp=_to_float(cell(row, 'mrp')),
                base_cost=_to_float(cell(row, 'base_cost')),
                igst_pct=_to_float(cell(row, 'igst_pct')),
                total_base_value=_to_float(cell(row, 'total_base_value')),
            ))
    return items


# ── Public entry points ───────────────────────────────────────────────

def parse_reliance_pdf(filepath: str | Path) -> ReliancePO:
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(filepath)
    with pdfplumber.open(filepath) as pdf:
        text = pdf.pages[0].extract_text() or ''
        tables: List[List[List[Any]]] = []
        for page in pdf.pages:
            tables.extend(page.extract_tables() or [])
    header = _parse_header(text)
    items = _parse_items(tables)
    if not items:
        raise ValueError(
            f"{filepath.name}: no Reliance line items found — the item "
            f"table header (MRP / Base Cost / Quantity / EAN) wasn't "
            f"recognised. Inspect page.extract_tables() to tune "
            f"_map_columns."
        )
    return ReliancePO(header=header, items=items)


def reliance_po_to_dataframe(po: ReliancePO):
    """
    Convert a ReliancePO into the engine-ready DataFrame.

    The Reliance config references: ``po_col='__po__'``,
    ``loc_col='__loc__'``, ``qty_col='Qty'``, ``ean_col='EAN'``,
    ``mrp_col='MRP'``, ``fob_col='Base Cost'``, ``hsn_col='HSN Code'``,
    ``amount_col={'multiply':['Base Cost','Qty']}``.
    """
    import pandas as pd
    rows = []
    for it in po.items:
        rows.append({
            '__po__':       po.header.po_number,
            '__loc__':      po.header.delivery_city,
            # v2.3.1: header fields replicated per row for the Tracker sheet.
            '__po_date__':  po.header.po_date,
            '__exp_date__': po.header.delivery_date,   # DELIVERY DATE
            '__state__':    po.header.delivery_state,
            'Sr No':        it.sr_no,
            'Article No':   it.article_no,
            'EAN':          it.ean,
            'HSN Code':     it.hsn_code,
            'Material Description': it.description,
            'Qty':          it.qty,
            'MRP':          it.mrp,
            'Base Cost':    it.base_cost,
            'GST Rate':     it.igst_pct,
            'Total Base Value': it.total_base_value,
            'Site':         po.header.site,
        })
    return pd.DataFrame(rows)


def load_reliance_pdf_as_dataframe(filepath: str | Path):
    """One-shot: parse PDF → engine-ready DataFrame.

    Registered in ``marketplace_engine.PDF_PARSERS`` under 'reliance'.
    """
    return reliance_po_to_dataframe(parse_reliance_pdf(filepath))
