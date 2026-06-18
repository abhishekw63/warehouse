"""
exporter.d365_package
=====================

Produce a **ready-to-publish D365 / Business Central "Edit in Excel" package**
by injecting our computed Header/Line rows into a *bound* template workbook —
so the operator no longer hand-copies the Headers (SO) / Lines (SO) sheets
into the connector file.

Why a template + surgery (not openpyxl)
---------------------------------------
The connector workbook isn't an ordinary spreadsheet. Alongside the two data
sheets it carries the binding that makes "Publish" push straight to D365:

  * ``xl/xmlMaps.xml``      — an XML Schema (``SalesHeaderList`` /
                              ``SalesLineList``) whose elements are the entity
                              fields.
  * ``xl/tables/table*.xml``— Excel Tables of ``tableType="xml"`` whose every
                              column binds to a schema xpath
                              (``<xmlColumnPr xpath=.../>``).
  * ``xl/connections.xml`` + ``tableSingleCells*.xml`` — the data connection
                              and the PackageCode / TableID single-cell binds.

openpyxl silently DROPS all of those parts on save, which would turn the file
into a dead spreadsheet. So we treat the template as an opaque ZIP and rewrite
ONLY the worksheet data rows + the table ``ref`` ranges, copying every other
part through **byte-for-byte**. The binding is preserved because we never
touch it.

What we change
--------------
Per sheet we keep row 1 (PackageCode / TableID metadata) and row 3 (the bound
column headers) verbatim, drop the template's sample data rows, and write our
own data starting at row 4 — as inline strings (the connector reads cell text
regardless of inline-vs-shared encoding). The owning table's ``ref`` is
stretched to ``A3:<lastcol><lastrow>`` to cover the new extent.

Column layout (must match the bound tables exactly)
---------------------------------------------------
SO Header table (sheet1, cols A–R, 18) — identical to ``headers_sheet``:
    Document Type | No. | Sell-to Customer No. | Ship-to Code | Posting Date |
    Order Date | Document Date | Invoice From Date | Invoice To Date |
    External Document No. | Location Code | Dimension Set ID | Supply Type |
    Voucher Narration | Brand | Channel | Catagory | Geography
SO Line table (sheet2, cols A–H, 8) — identical to ``lines_sheet``:
    Document Type | Document No. | Line No. | Type | No. | Location Code |
    Quantity | Unit Price
"""
from __future__ import annotations

import re
import zipfile
from datetime import datetime
from pathlib import Path
from typing import List, Optional
from xml.sax.saxutils import escape

from online_po_processor.data.models import ProcessingResult

_HDR_COLS = list('ABCDEFGHIJKLMNOPQR')          # 18 — SalesHeader
_LINE_COLS = list('ABCDEFGH')                   # 8  — SalesLine
# Transfer Order (Flipkart-TO): Transfer Header A–N (14), Transfer Line A–G (7).
_TO_HDR_COLS = list('ABCDEFGHIJKLMN')           # 14 — TransferHeader
_TO_LINE_COLS = list('ABCDEFG')                 # 7  — TransferLine
_TO_TRANSFER_FROM = 'PICK'
_TO_IN_TRANSIT = 'IN TRANSIT'
_TO_DIRECT_TRANSFER = 'false'
_LINE_NO_STEP = 10_000

# Data-cell styles lifted from the template's own sample rows so the injected
# rows look identical (header rows use s=10, line rows use s=3).
_HDR_STYLE = '10'
_LINE_STYLE = '3'


# ── value builders (mirror headers_sheet / lines_sheet exactly) ───────────

def _header_values(result: ProcessingResult) -> List[list]:
    """One 18-value row per unique PO — same logic as ``headers_sheet``."""
    today = datetime.now().strftime('%d-%m-%Y')
    loc = getattr(result, 'warehouse_code', '') or 'PICK'
    seen: set = set()
    rows: List[list] = []
    for so in result.rows:
        if so.po_number in seen:
            continue
        seen.add(so.po_number)
        rows.append([
            'Order', so.po_number, so.cust_no, so.ship_to,
            today, today, today, today, today,
            so.po_number, loc, '', 'B2B', '', '', '', '', '',
        ])
    return rows


def _line_values(result: ProcessingResult) -> List[list]:
    """One 8-value row per SORow — same logic as ``lines_sheet`` (incl. the
    forced/override Unit Price rule; blank otherwise)."""
    loc = getattr(result, 'warehouse_code', '') or 'PICK'
    override = bool(getattr(result, 'override_unit_price', False))
    rows: List[list] = []
    current_po = None
    line_no = 0
    for so in result.rows:
        if so.po_number != current_po:
            current_po = so.po_number
            line_no = 0
        line_no += _LINE_NO_STEP

        if so.forced_unit_price is not None:
            unit_price = round(so.forced_unit_price, 2)
        elif override and so.cost_price_ref is not None:
            unit_price = round(so.cost_price_ref, 2)
        else:
            unit_price = ''
        rows.append([
            'Order', so.po_number, line_no, 'Item', so.item_no,
            loc, so.qty, unit_price,
        ])
    return rows


def _header_values_to(result: ProcessingResult) -> List[list]:
    """One 14-value Transfer Header row per unique PO — mirrors the TO branch
    of ``headers_sheet`` (cols G–N are dimension/posting blanks)."""
    today = datetime.now().strftime('%d-%m-%Y')
    seen: set = set()
    rows: List[list] = []
    for so in result.rows:
        if so.po_number in seen:
            continue
        seen.add(so.po_number)
        rows.append([
            so.po_number, _TO_TRANSFER_FROM, so.ship_to or '', today,
            _TO_IN_TRANSIT, _TO_DIRECT_TRANSFER,
            '', '', '', '', '', '', '', '',
        ])
    return rows


def _line_values_to(result: ProcessingResult) -> List[list]:
    """One 7-value Transfer Line row per SORow: Document No. | Line No. |
    Item No. | Quantity | Unit of Measure | Transfer-from Bin Code |
    Transfer Price (= ``calc_price``, the post-GST cost)."""
    rows: List[list] = []
    current_po = None
    line_no = 0
    for so in result.rows:
        if so.po_number != current_po:
            current_po = so.po_number
            line_no = 0
        line_no += _LINE_NO_STEP
        price = so.calc_price if so.calc_price is not None else ''
        rows.append([
            so.po_number, line_no, so.item_no, so.qty, '', '', price,
        ])
    return rows


# ── XML surgery ───────────────────────────────────────────────────────────

def _cell_xml(col: str, row: int, value, style: str) -> str:
    """One inline-string ``<c>`` cell. Empty values emit nothing (Excel omits
    blank cells in a row)."""
    if value is None or value == '':
        return ''
    text = escape(str(value))
    return (f'<c r="{col}{row}" s="{style}" t="inlineStr">'
            f'<is><t xml:space="preserve">{text}</t></is></c>')


def _data_rows_xml(values: List[list], cols: List[str], style: str) -> str:
    out = []
    span = f'1:{len(cols)}'
    for i, vals in enumerate(values):
        r = 4 + i
        cells = ''.join(_cell_xml(cols[j], r, vals[j], style)
                        for j in range(len(vals)))
        out.append(f'<row r="{r}" spans="{span}">{cells}</row>')
    return ''.join(out)


def _keep_row(sheet_xml: str, rnum: str) -> str:
    """Extract the verbatim ``<row r="rnum">…</row>`` we must preserve."""
    m = re.search(r'<row\b[^>]*\br="' + rnum + r'"[^>]*>.*?</row>',
                  sheet_xml, re.S)
    return m.group(0) if m else ''


def _detect_style(sheet_xml: str, fallback: str) -> str:
    """The ``s="N"`` style id used by the template's own data cells (row 4),
    so injected rows inherit the same formatting across different templates.
    Falls back to the supplied default if row 4 is absent."""
    m = re.search(r'<c r="[A-Z]+4"[^>]*\bs="(\d+)"', sheet_xml)
    return m.group(1) if m else fallback


def _rewrite_sheet(sheet_xml: str, values: List[list], cols: List[str],
                   style: str, last_col: str) -> str:
    """Replace the sheet's data rows (4+) with ours, keeping rows 1 & 3, and
    stretch the ``<dimension>`` to the new extent."""
    style = _detect_style(sheet_xml, style)
    row1 = _keep_row(sheet_xml, '1')
    row3 = _keep_row(sheet_xml, '3')
    new_data = _data_rows_xml(values, cols, style)
    new_sheet_data = f'<sheetData>{row1}{row3}{new_data}</sheetData>'
    sheet_xml = re.sub(r'<sheetData>.*?</sheetData>',
                       lambda _m: new_sheet_data, sheet_xml, count=1, flags=re.S)
    last_row = 3 + len(values)
    sheet_xml = re.sub(
        r'<dimension ref="[^"]+"\s*/>',
        lambda _m: f'<dimension ref="A1:{last_col}{last_row}"/>',
        sheet_xml, count=1)
    return sheet_xml


def _rewrite_table_ref(table_xml: str, last_col: str, n_rows: int) -> str:
    """Stretch the table's ``ref`` to cover header row 3 + ``n_rows`` data."""
    last_row = 3 + n_rows
    return re.sub(
        r'(<table\b[^>]*\bref=")[^"]+(")',
        lambda m: f'{m.group(1)}A3:{last_col}{last_row}{m.group(2)}',
        table_xml, count=1)


def export_d365_package(result: ProcessingResult, template_path: str | Path,
                        output_path: str | Path) -> Path:
    """
    Inject ``result``'s Header/Line rows into the bound D365 template and write
    a ready-to-publish workbook to ``output_path``.

    Args:
        result:        the processed batch (SO mode — uses ``result.rows``).
        template_path: a bound "Edit in Excel" workbook (sheet1=Sales Header,
                       sheet2=Sales Line) for the operator's environment.
        output_path:   destination .xlsx.

    Returns:
        ``Path`` to the written file.

    Raises:
        FileNotFoundError / KeyError if the template is missing the expected
        parts (sheet1/sheet2/table1/table2) — surfaced rather than producing a
        silently-broken file.
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    # SO (Sales Order) vs TO (Transfer Order — Flipkart-TO) column layouts.
    is_to = getattr(result, 'output_type', 'so') == 'to'
    if is_to:
        headers = _header_values_to(result)
        lines = _line_values_to(result)
        hdr_cols, hdr_last = _TO_HDR_COLS, 'N'
        line_cols, line_last = _TO_LINE_COLS, 'G'
    else:
        headers = _header_values(result)
        lines = _line_values(result)
        hdr_cols, hdr_last = _HDR_COLS, 'R'
        line_cols, line_last = _LINE_COLS, 'H'

    with zipfile.ZipFile(template_path) as z:
        names = z.namelist()
        parts = {n: z.read(n) for n in names}

    for required in ('xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml',
                     'xl/tables/table1.xml', 'xl/tables/table2.xml'):
        if required not in parts:
            raise KeyError(
                f"D365 template {template_path.name} is missing {required} — "
                f"is this the bound 'Edit in Excel' workbook?")

    parts['xl/worksheets/sheet1.xml'] = _rewrite_sheet(
        parts['xl/worksheets/sheet1.xml'].decode('utf-8'),
        headers, hdr_cols, _HDR_STYLE, hdr_last).encode('utf-8')
    parts['xl/worksheets/sheet2.xml'] = _rewrite_sheet(
        parts['xl/worksheets/sheet2.xml'].decode('utf-8'),
        lines, line_cols, _LINE_STYLE, line_last).encode('utf-8')
    parts['xl/tables/table1.xml'] = _rewrite_table_ref(
        parts['xl/tables/table1.xml'].decode('utf-8'), hdr_last, len(headers)
    ).encode('utf-8')
    parts['xl/tables/table2.xml'] = _rewrite_table_ref(
        parts['xl/tables/table2.xml'].decode('utf-8'), line_last, len(lines)
    ).encode('utf-8')

    output_path.parent.mkdir(parents=True, exist_ok=True)
    # Preserve original member order (keeps [Content_Types].xml first).
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, parts[n])
    return output_path
