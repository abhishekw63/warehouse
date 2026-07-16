"""
grn.services.grn_parser
========================

Headless port of the standalone **Marketplace GRN / Return-Note parser**
(``standalone_files/standalone_blinkit_grn_extractor.py``) — same extraction
logic, no Tkinter. Each marketplace is a parser class registered in
``REGISTRY``; adding a marketplace = adding one class.

Public surface
--------------
``parse_pdf(path)``            → ``{ok, marketplace, header, items, summary, ...}``
``parse_many(paths)``          → combined dict across many PDFs
``REGISTRY``                   → {name: parser class} (Blinkit, Flipkart)

Returns plain JSON-safe dicts (fat service → thin view), so the web layer just
renders them.
"""
from __future__ import annotations

import re

import pdfplumber


# ── helpers (verbatim from the standalone) ──────────────────────────────────
def clean_upc(raw) -> str:
    return re.sub(r'\s+', '', str(raw))


def clean_number(val) -> float:
    if val is None or str(val).strip() in ('-', '', 'None'):
        return 0.0
    try:
        return float(str(val).replace('\n', '').replace(',', '').strip())
    except (ValueError, TypeError):
        return 0.0


class BaseParser:
    name = ''
    document_type = ''
    item_columns: list = []
    summary_columns: list = []
    status_column = ''
    status_colors: dict = {}

    @classmethod
    def detect(cls, first_page_text: str) -> bool:
        raise NotImplementedError

    def parse(self, pdf) -> dict:
        """Return {header, items(list[dict]), summary(dict), raw_text}."""
        raise NotImplementedError


# ── Blinkit GRN ─────────────────────────────────────────────────────────────
class BlinkitParser(BaseParser):
    name = 'Blinkit'
    document_type = 'GRN'
    status_column = 'Line GRN Status'
    status_colors = {'Full GRN': '00C853', 'Partial GRN': 'FFB300',
                     'Not GRNed': 'D50000'}
    item_columns = [
        'PO Number', 'PO Date', 'Facility', 'Sr No', 'Item Code', 'UPC / GTIN',
        'Description', 'MRP', 'Landing Rate', 'PO Qty', 'GRN Qty', 'Fill Rate %',
        'GRN Amount', 'GMV Loss', 'Line GRN Status', 'PO<>EAN',
    ]
    summary_columns = [
        'PO Number', 'PO Date', 'Facility', 'Total PO Qty', 'Total GRN Qty',
        'Fill Rate %', 'Articles in PO', 'Articles in GRN', 'Total PO Amount',
        'Net GRN Amount', 'GMV Loss',
    ]

    @classmethod
    def detect(cls, text: str) -> bool:
        t = (text or '').lower()
        return 'blinkit' in t or 'bcpl' in t or 'p.o. number' in t

    def _header(self, pdf) -> dict:
        text = pdf.pages[0].extract_text() or ''
        po = re.search(r'P\.O\.\s*Number\s*[:\s]+(\d+)', text)
        dt = re.search(r'Date\s*[:\s]+([\w.]+\s+\d+,\s+\d{4})', text)
        fac = re.search(r'BCPL\s*-\s*(.+?)(?:\n|Contact)', text)
        return {
            'PO Number': po.group(1).strip() if po else 'UNKNOWN',
            'PO Date': dt.group(1).strip() if dt else '',
            'Facility': fac.group(1).strip() if fac else '',
        }

    def _summary(self, all_text: str) -> dict:
        def find(p):
            m = re.search(p, all_text, re.IGNORECASE)
            return clean_number(m.group(1)) if m else 0.0
        fr = re.search(r'Fill rate:\s*([\d.]+)%', all_text)
        return {
            'Total PO Qty': find(r'Total Quantity in PO:\s*([\d,]+)'),
            'Total GRN Qty': find(r'Total Quantity in GRN\(s\):\s*([\d,]+)'),
            'Fill Rate %': float(fr.group(1)) if fr else 0.0,
            'Articles in PO': find(r'Articles in PO:\s*([\d,]+)'),
            'Articles in GRN': find(r'Articles in GRN\(s\):\s*([\d,]+)'),
            'Total PO Amount': find(r'Total Amount in PO\s+([\d,\.]+)'),
            'Net GRN Amount': find(r'Net amt\. by GRN\s+([\d,\.]+)'),
            'GMV Loss': find(r'Potential GMV Loss \(in INR\)\s+([\d,\.]+)'),
        }

    def parse(self, pdf) -> dict:
        header = self._header(pdf)
        rows, all_text = [], ''
        for page in pdf.pages:
            all_text += (page.extract_text() or '') + '\n'
            for table in page.extract_tables():
                for row in table:
                    if not row or row[0] is None:
                        continue
                    if not re.match(r'^\d+$', str(row[0]).strip()):
                        continue
                    try:
                        upc = clean_upc(row[2] if len(row) > 2 else '')
                        desc = str(row[3] or '').replace('\n', ' ').strip()
                        po_qty = int(clean_number(row[8])) if len(row) > 8 else 0
                        grn_qty = int(clean_number(row[9])) if len(row) > 9 else 0
                        mrp = clean_number(row[4]) if len(row) > 4 else 0.0
                        lr = clean_number(row[6]) if len(row) > 6 else 0.0
                        fr_raw = str(row[10] or '').strip() if len(row) > 10 else '-'
                        fr = clean_number(fr_raw) if fr_raw != '-' else 0.0
                        grn_amt = clean_number(row[11]) if len(row) > 11 else 0.0
                        gmv = clean_number(row[12]) if len(row) > 12 else 0.0
                        status = ('Not GRNed' if grn_qty == 0
                                  else 'Partial GRN' if grn_qty < po_qty
                                  else 'Full GRN')
                        rows.append({
                            'PO Number': header['PO Number'], 'PO Date': header['PO Date'],
                            'Facility': header['Facility'], 'Sr No': int(row[0]),
                            'Item Code': str(row[1] or '').strip(), 'UPC / GTIN': upc,
                            'Description': desc, 'MRP': mrp, 'Landing Rate': lr,
                            'PO Qty': po_qty, 'GRN Qty': grn_qty, 'Fill Rate %': fr,
                            'GRN Amount': grn_amt, 'GMV Loss': gmv,
                            'Line GRN Status': status,
                            'PO<>EAN': f"{header['PO Number']}<>{upc}",
                        })
                    except (ValueError, TypeError, IndexError):
                        pass
        summary = self._summary(all_text)
        header.update(summary)
        return {'header': header, 'items': rows, 'summary': summary,
                'raw_text': all_text}


# ── Flipkart Return Note (kept for parity; Blinkit is the primary) ──────────
class FlipkartParser(BaseParser):
    name = 'Flipkart'
    document_type = 'Return Note'
    status_column = 'Section'
    status_colors = {'A': '00C853', 'B': 'FFB300', 'C': 'D50000'}

    @classmethod
    def detect(cls, text: str) -> bool:
        t = (text or '').lower()
        return 'flipkart' in t or 'return note' in t or 'wsn' in t

    def parse(self, pdf) -> dict:
        # Flipkart return-note extraction is available in the standalone; the web
        # GRN feature ships Blinkit first. Returning empty keeps detect()/registry
        # working without pulling in the full return-note table logic yet.
        all_text = ''
        for page in pdf.pages:
            all_text += (page.extract_text() or '') + '\n'
        return {'header': {}, 'items': [], 'summary': {},
                'raw_text': all_text,
                'note': 'Flipkart return-note parsing is not enabled in the web '
                        'app yet — Blinkit GRN is live.'}


REGISTRY = {c.name: c for c in (BlinkitParser, FlipkartParser)}


def _auto_detect(pdf) -> str | None:
    first = (pdf.pages[0].extract_text() or '') if pdf.pages else ''
    for name, klass in REGISTRY.items():
        if klass.detect(first):
            return name
    return None


def parse_pdf(path: str, marketplace: str = '') -> dict:
    """Parse ONE GRN/return PDF. ``marketplace`` forces a parser; blank →
    auto-detect. Returns a JSON-safe dict; never raises."""
    try:
        with pdfplumber.open(path) as pdf:
            name = marketplace or _auto_detect(pdf)
            if not name or name not in REGISTRY:
                return {'ok': False,
                        'error': 'Could not detect the marketplace from this PDF. '
                                 'Is it a Blinkit GRN?'}
            parser = REGISTRY[name]()
            res = parser.parse(pdf)
        return {
            'ok': True, 'marketplace': name,
            'document_type': parser.document_type,
            'header': res['header'], 'items': res['items'],
            'summary': res['summary'],
            'item_columns': parser.item_columns,
            'summary_columns': parser.summary_columns,
            'status_column': parser.status_column,
            'status_colors': parser.status_colors,
            'note': res.get('note', ''),
        }
    except Exception as e:  # noqa: BLE001 — a bad PDF must not 500 the page
        return {'ok': False, 'error': f'{type(e).__name__}: {e}'}


def parse_many(paths: list) -> dict:
    """Parse many PDFs (same marketplace) → one combined result with a
    per-PO summary and all line items. Files that fail are reported, never fatal."""
    all_items, summaries, errors, mkt = [], [], [], ''
    for p in paths:
        r = parse_pdf(p)
        if not r['ok']:
            errors.append({'file': p, 'error': r['error']})
            continue
        mkt = r['marketplace']
        all_items.extend(r['items'])
        s = dict(r['header'])
        summaries.append(s)
    stat_col = REGISTRY[mkt].status_column if mkt else 'Line GRN Status'
    from collections import Counter
    status_counts = Counter(i.get(stat_col, '') for i in all_items)
    tot = {
        'files': len(paths), 'parsed': len(summaries), 'failed': len(errors),
        'pos': len({i.get('PO Number') for i in all_items}),
        'lines': len(all_items),
        'po_qty': sum(int(i.get('PO Qty', 0) or 0) for i in all_items),
        'grn_qty': sum(int(i.get('GRN Qty', 0) or 0) for i in all_items),
        'gmv_loss': round(sum(clean_number(i.get('GMV Loss', 0)) for i in all_items), 2),
        'full': status_counts.get('Full GRN', 0),
        'partial': status_counts.get('Partial GRN', 0),
        'not_grn': status_counts.get('Not GRNed', 0),
    }
    tot['fill_rate'] = round(tot['grn_qty'] * 100 / tot['po_qty'], 2) if tot['po_qty'] else 0.0
    return {
        'ok': bool(summaries), 'marketplace': mkt,
        'items': all_items, 'po_summaries': summaries, 'totals': tot,
        'errors': errors,
        'item_columns': REGISTRY[mkt].item_columns if mkt else BlinkitParser.item_columns,
        'summary_columns': REGISTRY[mkt].summary_columns if mkt else BlinkitParser.summary_columns,
        'status_column': stat_col,
        'status_colors': REGISTRY[mkt].status_colors if mkt else BlinkitParser.status_colors,
    }
