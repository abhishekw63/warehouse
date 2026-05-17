"""
exporter.d365_exporter
======================

Fills a Dynamics 365 sample package template with Online PO Sales Order
data, producing an Excel file the ERP team can drop straight into D365
Business Central's Data Management framework.

Strategy
--------
A D365 sample template is a real ``.xlsx`` file that internally is a ZIP
of XML parts. We don't use ``openpyxl`` for the fill because it strips
the table/dimension metadata D365 uses to accept the package. Instead we
surgically edit the XML:

1. Open the template as a ZIP, read every part into memory.
2. Extend ``xl/sharedStrings.xml`` with any new strings we'll reference
   (SO numbers, location codes, literal 'Order', 'Item', 'B2B', date).
3. Fill empty pre-styled cells in ``sheet1.xml`` (Headers) and
   ``sheet2.xml`` (Lines) — regex-based replace that keeps each cell's
   ``s="N"`` style id intact.
4. If our data exceeds the template's pre-formatted rows, inject new
   ``<row>`` elements with matching style ids before filling.
5. Remove ALL trailing empty rows and fix ``<dimension>`` +
   ``<table>`` refs so D365 doesn't complain about "extra" blank rows.
6. Write the modified parts back to a fresh ZIP.

What maps where
---------------
From each ``SORow`` we produce:

* **Headers sheet (sheet1)** — one row per unique PO:
    A: 'Order', B: po_number, C: cust_no, D: ship_to,
    E–I: today's date, J: po_number again (External Document No.),
    K: ``_ERP_LOCATION_CODE`` ('PICK' — see note below),
    M: 'B2B'

* **Lines sheet (sheet2)** — one row per ``SORow``:
    A: 'Order', B: po_number, C: line_no (10000 step),
    D: 'Item', E: item_no (numeric),
    F: ``_ERP_LOCATION_CODE`` ('PICK'), G: qty

ERP Location Code (v1.5.8)
--------------------------
The D365 Business Central configuration posts every online B2B sales
order to a single warehouse code (``PICK``) regardless of which city
the order is shipping to. This used to populate from
``SORow.mapped_location`` (derived from the Ship-To B2B mapping
sheet's ``Del Location`` column), which inadvertently let raw
facility names like "Bilaspur" or "Mumbai M11 - Feeder Warehouse"
land in the Location Code column — which the ERP doesn't recognise.

The per-marketplace mapped_location is still preserved on each
``SORow`` and surfaced in the Summary sheet so the warehouse team
can see which marketplace location generated each PO. Only the D365
import file receives the hardcoded constant.
"""

from __future__ import annotations
import logging
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set

from online_po_processor.data.models import ProcessingResult, SORow


# ── XML shaping constants ──────────────────────────────────────────────
#
# These mirror the column letters used by the standard D365 sample
# package for Sales Orders. If the template ever changes, these need
# to move accordingly — but for now both GT Mass and Online PO share
# the same template, so keeping the values hard-coded here is fine.

# Header sheet columns — A..R (18 columns wide, mirrors D365 template)
_HEADER_COLS: List[str] = list('ABCDEFGHIJKLMNOPQR')

# Line sheet columns — A..H (8 columns wide)
_LINE_COLS: List[str] = list('ABCDEFGH')

# v1.5.8: The ERP's Business Central configuration posts ALL online
# marketplace B2B orders to a single warehouse location code. The raw
# facility names from marketplace punch files ("Mumbai M11 - Feeder
# Warehouse", "Bilaspur", etc.) are operational labels — the ERP
# doesn't have those as valid location codes. Regardless of marketplace
# or which delivery city the order goes to, the D365 Sales Header's
# Location Code (col K) and Sales Line's Location Code (col F) must
# always contain this fixed value.
#
# The marketplace-specific raw/mapped locations are still preserved
# on the ``SORow`` objects and displayed in the Summary sheet for
# operational visibility — this constant only governs what lands in
# the D365 import file the ERP team actually consumes.
_ERP_LOCATION_CODE = 'PICK'

# Style ids carried on the empty template cells — these come straight
# from the D365 sample's cellXfs table. New rows we inject must use the
# same style id so formatting stays consistent.
_HEADER_STYLE_ID = '11'
_LINE_STYLE_ID = '8'


# ── v2.0.0: Transfer Order constants ───────────────────────────────────
#
# Used only when exporting Flipkart-TO results (output_type='to'). The
# D365 TO sample template has different sheet structure than the SO
# template:
#
#   sheet1 ('Transfer Header')  — 12 cols A..L, data starts row 4
#   sheet2 ('Transfer Line')    —  9 cols A..I, data starts row 4
#
# Layout per row:
#
#   Transfer Header
#     A No.                         po_number (numeric, e.g. 204345116)
#     B Transfer-from Code          'PICK' (constant, see below)
#     C Transfer-to Code            ship_to from mapping (e.g. FK_BHW_BTS)
#     D Posting Date                today (dd-mm-yyyy string)
#     E In-Transit Code             'IN TRANSIT' (constant)
#     F Direct Transfer             'false' (constant string, lowercase)
#     G–L Dimensions                EMPTY
#
#   Transfer Line
#     A Document No.                po_number (numeric)
#     B Line No.                    '10000','20000',... STRING (per dump)
#     C Item No.                    item_no (numeric)
#     D Quantity                    qty (numeric)
#     E Unit of Measure             EMPTY
#     F Qty. to Ship                EMPTY
#     G Qty. to Receive             EMPTY
#     H Dimension Set ID            EMPTY
#     I Transfer Price              calc_price (numeric, MRP × m% ÷ GST)
#
# The four hardcoded values below come straight from the manual D365
# dump (30_04_2026_14_31.xlsx) the user provided as the canonical
# spec. They were verified the same on every header row in that dump.

# 'PICK' is also the SO Location Code, but we keep a separate constant
# to make the TO-vs-SO branching explicit and self-documenting.
_TO_TRANSFER_FROM_CODE = 'PICK'

# In-Transit Code — same on every TO header in the manual dump. The
# ERP team confirmed this is fixed across all online marketplace TOs.
_TO_IN_TRANSIT_CODE = 'IN TRANSIT'

# Direct Transfer — boolean stored as lowercase string in the manual
# dump (Excel boolean cells render as TRUE/FALSE but the dump uses the
# string 'false'). Whether D365 import accepts both is untested; we
# match the manual dump exactly to avoid surprises.
_TO_DIRECT_TRANSFER = 'false'

# Header sheet columns — A..L (12 columns wide for TO header).
_TO_HEADER_COLS: List[str] = list('ABCDEFGHIJKL')

# Line sheet columns — A..I (9 columns wide for TO line).
_TO_LINE_COLS: List[str] = list('ABCDEFGHI')


class D365Exporter:
    """
    Fill a D365 Sales Order sample template with the rows from a
    ``ProcessingResult``.

    Usage::

        exporter = D365Exporter()
        out_path = exporter.export(result, template_path, output_dir)

    The exporter does NOT mutate ``result`` in any way. Output is
    written to ``<output_dir>/d365_import_<ts>.xlsx``. ``output_dir``
    is created on demand; typical choice is the same
    ``<punch_dir>/output/`` folder used by the main SO exporter.
    """

    def export(
        self,
        result: ProcessingResult,
        template_path: str,
        output_dir: Path,
    ) -> Optional[Path]:
        """
        Entry point — produce a D365-ready import workbook from
        ``result``.

        Args:
            result:         Populated ``ProcessingResult`` from the
                            engine. Must contain at least one row.
            template_path:  Filesystem path to the D365 sample template
                            .xlsx.
            output_dir:     Directory to write the output into. Created
                            if it does not already exist.

        Returns:
            Path to the generated .xlsx on success; ``None`` if the
            result was empty or a fatal template error occurred
            (errors are logged; a higher-level GUI layer is expected
            to surface them to the user).
        """
        # ── Guard: nothing to export ────────────────────────────────────
        if not result.rows:
            logging.warning("D365 export aborted — no rows in result")
            return None

        # ── v2.0.0: Branch by output_type ────────────────────────────────
        # SO is the default (existing behavior, untouched). When
        # ``output_type='to'`` the result was produced by the engine's
        # TO pipeline (currently Flipkart-TO) and must be filled into
        # a Transfer Order template — different sheet names, different
        # column layout. The TO path mirrors the SO flow at a high
        # level (read ZIP → fill sharedStrings → fill sheet1+sheet2 →
        # finalize → write ZIP) but with TO-specific writers.
        if getattr(result, 'output_type', 'so') == 'to':
            return self._export_to(result, template_path, output_dir)

        output_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime('%d-%m-%Y_%H%M%S')
        out_path = output_dir / f"d365_import_{timestamp}.xlsx"

        # Work on a copy of the template so the original is never touched.
        try:
            shutil.copy2(template_path, str(out_path))
        except (OSError, shutil.SameFileError) as e:
            logging.error("D365 copy failed: %s", e)
            return None

        # ── Core fill workflow ──────────────────────────────────────────
        try:
            zip_contents = self._read_zip(out_path)
        except (KeyError, zipfile.BadZipFile, OSError) as e:
            logging.error("D365 template read failed: %s", e)
            out_path.unlink(missing_ok=True)
            return None

        today_str = datetime.now().strftime('%d-%m-%Y')
        unique_pos = self._unique_pos(result.rows)

        # v1.9.0: warehouse code is now pulled from the result (set
        # by the GUI based on the user's dropdown choice) instead of
        # a module-level constant. Falling back to ``_ERP_LOCATION_CODE``
        # keeps backwards compatibility with any callers that
        # construct a ProcessingResult manually without going
        # through the GUI (tests, scripts).
        warehouse_code = getattr(result, 'warehouse_code', None) or _ERP_LOCATION_CODE
        logging.info("D365 export using warehouse code: %s", warehouse_code)

        # v2.1.3: read the override flag from the ``ProcessingResult``
        # field (set by the GUI's "Override Unit Price" checkbox)
        # instead of from the marketplace config dict.
        #
        # Pre-v2.1.3 this was a config-driven decision: every BlinkMP
        # run automatically overrode because BlinkMP's config had
        # ``override_unit_price=True``. v2.1.3 moves the runtime
        # decision to the GUI so operators can opt in/out per run.
        # The marketplace config flag still exists but only as a hint
        # for the GUI to pre-check the toggle on marketplace change.
        override_unit_price = bool(
            getattr(result, 'override_unit_price', False)
        )
        if override_unit_price:
            logging.info(
                "D365 export: override_unit_price ON for %s — "
                "Sales Line col H will carry Cost Price",
                result.marketplace,
            )

        # Step 1: extend sharedStrings with every literal we'll reference.
        string_map = self._build_shared_strings(
            zip_contents, unique_pos, result.rows, today_str,
            warehouse_code,
        )

        # Step 2: fill the Headers sheet (sheet1).
        zip_contents['xl/worksheets/sheet1.xml'] = self._fill_header_sheet(
            zip_contents['xl/worksheets/sheet1.xml'].decode('utf-8'),
            unique_pos, today_str, string_map, warehouse_code,
        ).encode('utf-8')

        # Step 3: fill the Lines sheet (sheet2).
        zip_contents['xl/worksheets/sheet2.xml'] = self._fill_line_sheet(
            zip_contents['xl/worksheets/sheet2.xml'].decode('utf-8'),
            result.rows, string_map, warehouse_code,
            override_unit_price,
        ).encode('utf-8')

        # Step 4: trim trailing empty rows + fix dimensions/table refs.
        self._finalize_sheets(
            zip_contents, len(unique_pos), len(result.rows),
        )

        # Step 5: write the modified ZIP back to disk.
        try:
            self._write_zip(out_path, zip_contents)
        except (OSError, zipfile.BadZipFile) as e:
            logging.error("D365 write failed: %s", e)
            out_path.unlink(missing_ok=True)
            return None

        logging.info("D365 export written to %s", out_path)
        return out_path

    # ── Read / write helpers ───────────────────────────────────────────

    @staticmethod
    def _read_zip(path: Path) -> Dict[str, bytes]:
        """
        Read every file inside an .xlsx (ZIP) into a name→bytes dict.

        Loading everything into memory is fine — template files are
        tiny (~30 KB) and we need to mutate several parts together.
        """
        contents: Dict[str, bytes] = {}
        with zipfile.ZipFile(str(path), 'r') as zf:
            for name in zf.namelist():
                contents[name] = zf.read(name)
        return contents

    @staticmethod
    def _write_zip(path: Path, contents: Dict[str, bytes]) -> None:
        """
        Rewrite an .xlsx (ZIP) from an in-memory name→bytes dict.

        Uses DEFLATE compression to stay close to the original file
        size; D365's importer doesn't care about the compression level
        but keeps output tidy.
        """
        with zipfile.ZipFile(
            str(path), 'w', zipfile.ZIP_DEFLATED,
        ) as zf:
            for name, data in contents.items():
                zf.writestr(name, data)

    # ── Row aggregation ────────────────────────────────────────────────

    @staticmethod
    def _unique_pos(rows: List[SORow]) -> List[SORow]:
        """
        Return the first ``SORow`` seen for each distinct PO number,
        preserving input order.

        The headers sheet has one row per PO, and we use the first row
        encountered as the "representative" for that PO (since
        cust_no / ship_to / mapped_location are all PO-level identical).
        """
        seen: Set[str] = set()
        unique: List[SORow] = []

        for row in rows:
            if row.po_number not in seen:
                seen.add(row.po_number)
                unique.append(row)

        return unique

    # ── sharedStrings.xml handling ─────────────────────────────────────

    @staticmethod
    def _build_shared_strings(
        zip_contents: Dict[str, bytes],
        unique_pos: List[SORow],
        all_rows: List[SORow],
        today_str: str,
        warehouse_code: str = _ERP_LOCATION_CODE,
    ) -> Dict[str, int]:
        """
        Extend ``xl/sharedStrings.xml`` with every literal we'll emit.

        Excel uses a shared-string table so duplicate text values are
        stored once and referenced by integer index from the cell
        body. We:

        1. Parse out the existing ``<t>…</t>`` values into a
           name→index map.
        2. Collect every new string we'll need (PO numbers, location
           codes, the literals ``Order`` / ``Item`` / ``B2B``, today's
           date).
        3. Append any strings not already present, continuing the index.
        4. Rewrite the full table back as bytes in ``zip_contents``.

        v1.9.0: the warehouse ERP code is now supplied by the caller
        (from the GUI's dropdown) instead of being the hardcoded
        ``_ERP_LOCATION_CODE``. Backwards-compatible default keeps
        non-GUI callers working.

        Returns:
            Dict mapping string value → 0-based shared-string index.
            Callers use this to produce ``<v>N</v>`` cell bodies.
        """
        ss_xml = zip_contents['xl/sharedStrings.xml'].decode('utf-8')

        # Pull out existing strings — preserves their index positions.
        existing = re.findall(r'<t[^>]*>([^<]*)</t>', ss_xml)
        string_map: Dict[str, int] = {s: i for i, s in enumerate(existing)}

        # Collect the set of strings our fill code will emit.
        # v1.9.0: ``warehouse_code`` replaces the previously-hardcoded
        # ``_ERP_LOCATION_CODE``. Same role — goes into col K
        # (Sales Header) and col F (Sales Line) — but now variable
        # per-run based on GUI dropdown.
        new_strings: Set[str] = {
            'Order', 'Item', 'B2B', today_str, warehouse_code,
        }

        for po_row in unique_pos:
            new_strings.add(po_row.po_number)
            if po_row.cust_no:
                new_strings.add(po_row.cust_no)
            if po_row.ship_to:
                new_strings.add(po_row.ship_to)

        for row in all_rows:
            new_strings.add(row.po_number)

        # Allocate indices for previously-unseen strings.
        next_idx = len(existing)
        for s in sorted(new_strings):
            if s not in string_map:
                string_map[s] = next_idx
                next_idx += 1

        # Rebuild the table body, slot by slot, so both preserved and
        # new strings land at the correct index.
        total_count = next_idx
        si_items: List[str] = [''] * total_count

        for s, idx in string_map.items():
            si_items[idx] = f'<si><t>{_xml_escape(s)}</t></si>'

        ss_new = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<sst xmlns="http://schemas.openxmlformats.org/'
            f'spreadsheetml/2006/main" '
            f'count="{total_count}" uniqueCount="{total_count}">'
            + ''.join(si_items)
            + '</sst>'
        )
        zip_contents['xl/sharedStrings.xml'] = ss_new.encode('utf-8')

        return string_map

    # ── Headers sheet fill (sheet1) ────────────────────────────────────

    def _fill_header_sheet(
        self,
        s1_xml: str,
        unique_pos: List[SORow],
        today_str: str,
        string_map: Dict[str, int],
        warehouse_code: str = _ERP_LOCATION_CODE,
    ) -> str:
        """
        Fill the D365 'Sales Header' sheet (sheet1).

        One row per unique PO. Rows 1–3 are template headers — data
        begins at row 4.

        Column layout (matches the D365 sample template):
            A Document Type            'Order'
            B No.                      po_number
            C Sell-to Customer No.     cust_no
            D Ship-to Code             ship_to
            E Posting Date             today
            F Order Date               today
            G Document Date            today
            H Invoice From Date        today
            I Invoice To Date          today
            J External Document No.    po_number (duplicate of B)
            K Location Code            _ERP_LOCATION_CODE ('PICK')
            L Dimension Set ID         (left empty)
            M Supply Type              'B2B'
            N-R Dimension codes        (left empty)
        """
        # v1.5.2: detect the actual data-row style id from the
        # template instead of hardcoding. The _HEADER_STYLE_ID
        # constant (11) was right for GT Mass's D365 template but
        # wrong for BL_Sample (which uses style id 5). Reading a
        # cell from the first data row (row 4 column A) gives us
        # the right value for whichever template was supplied.
        detected_style = self._detect_data_style(
            s1_xml, data_start_row=4, fallback=_HEADER_STYLE_ID,
        )

        s1_xml = self._ensure_enough_rows(
            s1_xml,
            needed=len(unique_pos),
            template_header_rows=3,
            data_start_row=4,
            columns=_HEADER_COLS,
            style_id=detected_style,
            sheet_label='Sheet1',
        )

        for i, po_row in enumerate(unique_pos):
            r = i + 4  # data rows start at row 4

            # Document Type = 'Order'
            s1_xml = _fill_cell(
                s1_xml, 'A', r, 'Order',
                is_string=True, string_map=string_map,
            )

            # No. = PO number
            s1_xml = _fill_cell(
                s1_xml, 'B', r, po_row.po_number,
                is_string=True, string_map=string_map,
            )

            # Sell-to Customer No. (from mapping)
            if po_row.cust_no:
                s1_xml = _fill_cell(
                    s1_xml, 'C', r, po_row.cust_no,
                    is_string=True, string_map=string_map,
                )

            # Ship-to Code (from mapping)
            if po_row.ship_to:
                s1_xml = _fill_cell(
                    s1_xml, 'D', r, po_row.ship_to,
                    is_string=True, string_map=string_map,
                )

            # Dates — all set to today
            for col in 'EFGHI':
                s1_xml = _fill_cell(
                    s1_xml, col, r, today_str,
                    is_string=True, string_map=string_map,
                )

            # External Document No. = PO number (again)
            s1_xml = _fill_cell(
                s1_xml, 'J', r, po_row.po_number,
                is_string=True, string_map=string_map,
            )

            # v1.5.8: Location Code is always the literal ERP
            # location code — NOT the marketplace's raw or mapped
            # location. The ERP posts every B2B sales order to this
            # single warehouse regardless of which city the order is
            # shipping to; the per-city data is preserved on the
            # SORow for operational reporting but doesn't belong in
            # this column. Hardcoded unconditionally — no guard for
            # empty values because the constant is always set.
            # v1.9.0: ``warehouse_code`` parameter (from the GUI's
            # Warehouse dropdown) replaces the previously hardcoded
            # ``_ERP_LOCATION_CODE``.
            s1_xml = _fill_cell(
                s1_xml, 'K', r, warehouse_code,
                is_string=True, string_map=string_map,
            )

            # Supply Type = 'B2B' (all online marketplace orders)
            s1_xml = _fill_cell(
                s1_xml, 'M', r, 'B2B',
                is_string=True, string_map=string_map,
            )

        return s1_xml

    # ── Lines sheet fill (sheet2) ──────────────────────────────────────

    def _fill_line_sheet(
        self,
        s2_xml: str,
        rows: List[SORow],
        string_map: Dict[str, int],
        warehouse_code: str = _ERP_LOCATION_CODE,
        override_unit_price: bool = False,
    ) -> str:
        """
        Fill the D365 'Sales Line' sheet (sheet2).

        One row per ``SORow``. Rows 1–3 are template headers — data
        begins at row 4. Line numbers increment by 10000 within each
        PO and reset when the PO changes (standard D365 convention,
        mirrors our own Lines (SO) sheet in the main exporter).

        Column layout:
            A Document Type    'Order'
            B Document No.     po_number
            C Line No.         10000, 20000, ... (numeric, resets per PO)
            D Type             'Item'
            E No.              item_no (numeric when parseable)
            F Location Code    ``warehouse_code`` (v1.9.0 — from GUI)
            G Quantity         qty (numeric)
            H Unit Price       EMPTY by default (WMS computes downstream)
                               or ``cost_price_ref`` when
                               ``override_unit_price=True`` (v1.9.1).
                               See the override_unit_price config doc
                               for the rationale (BCPL vendor record
                               has wrong margin for BlinkMP).
        """
        # v1.5.2: detect actual style id (see _fill_header_sheet).
        detected_style = self._detect_data_style(
            s2_xml, data_start_row=4, fallback=_LINE_STYLE_ID,
        )

        s2_xml = self._ensure_enough_rows(
            s2_xml,
            needed=len(rows),
            template_header_rows=3,
            data_start_row=4,
            columns=_LINE_COLS,
            style_id=detected_style,
            sheet_label='Sheet2',
        )

        current_po: Optional[str] = None
        line_no = 0

        for i, row in enumerate(rows):
            r = i + 4

            # Reset line numbering when PO changes.
            if row.po_number != current_po:
                current_po = row.po_number
                line_no = 0

            line_no += 10000

            # Document Type
            s2_xml = _fill_cell(
                s2_xml, 'A', r, 'Order',
                is_string=True, string_map=string_map,
            )

            # Document No. (PO number)
            s2_xml = _fill_cell(
                s2_xml, 'B', r, row.po_number,
                is_string=True, string_map=string_map,
            )

            # Line No. (numeric, step 10000)
            s2_xml = _fill_cell(
                s2_xml, 'C', r, line_no, is_string=False,
            )

            # Type = 'Item'
            s2_xml = _fill_cell(
                s2_xml, 'D', r, 'Item',
                is_string=True, string_map=string_map,
            )

            # No. (item_no) — numeric when it's an int-like value,
            # else fall back to shared-string write so non-numeric
            # item codes still land intact.
            item_numeric: Optional[int] = None
            try:
                item_numeric = int(row.item_no)
            except (TypeError, ValueError):
                item_numeric = None

            if item_numeric is not None:
                s2_xml = _fill_cell(
                    s2_xml, 'E', r, item_numeric, is_string=False,
                )
            else:
                # Track the value in the string map on-the-fly — we may
                # have missed adding unusual item_no strings earlier.
                val = str(row.item_no)
                if val not in string_map:
                    # Degrade gracefully: write the value literally as
                    # inline string rather than shared-string. D365
                    # still reads it, just without dedupe.
                    s2_xml = _fill_inline_string(s2_xml, 'E', r, val)
                else:
                    s2_xml = _fill_cell(
                        s2_xml, 'E', r, val,
                        is_string=True, string_map=string_map,
                    )

            # v1.5.8: Location Code is the literal ERP location
            # code — same reasoning as the Sales Header write site.
            # Every line on every sales order posts to this single
            # warehouse; marketplace facility names live on the
            # SORow for reporting but don't belong in the ERP
            # import file.
            # v1.9.0: value now comes from caller's ``warehouse_code``
            # parameter (GUI Warehouse dropdown) rather than the
            # module constant.
            s2_xml = _fill_cell(
                s2_xml, 'F', r, warehouse_code,
                is_string=True, string_map=string_map,
            )

            # Quantity (numeric)
            s2_xml = _fill_cell(
                s2_xml, 'G', r, row.qty, is_string=False,
            )

            # v1.9.1: Unit Price (col H) — empty by default, populated
            # with the engine's post-GST Cost Price per row when the
            # marketplace config sets ``override_unit_price: True``.
            # Used by BlinkMP to force the correct 75%-margin cost
            # into the ERP (which would otherwise compute using the
            # BCPL vendor's default 70% setting meant for Blinkit).
            if override_unit_price:
                cost = row.cost_price_ref
                if cost is not None and cost > 0:
                    # Round to 4 dp to match how the ERP stores cost
                    # values — anything finer is rounding dust from
                    # the MRP × margin ÷ GST-divisor calculation.
                    s2_xml = _fill_cell(
                        s2_xml, 'H', r, round(cost, 4),
                        is_string=False,
                    )
                else:
                    # Lenient: leave col H empty for this row, log
                    # so the user can spot gaps (typically a row
                    # whose master lookup failed — the validation
                    # sheet flags these separately as
                    # NOT_IN_MASTER).
                    logging.warning(
                        "D365 override_unit_price: row PO=%s item=%s "
                        "has no cost_price_ref; leaving col H empty",
                        row.po_number, row.item_no or row.ean,
                    )

        return s2_xml

    # ── v2.0.0: Transfer Order export ──────────────────────────────────
    #
    # This block is the entire TO code path. Called from ``export()``
    # via the ``output_type=='to'`` branch. Mirrors the SO pipeline
    # structurally (read → fill shared strings → fill sheet1/sheet2 →
    # finalize → write) but with TO-shaped sheet writers and
    # different shared-string content (no 'Order'/'Item'/'B2B'
    # literals; instead 'PICK','IN TRANSIT','false' + today's date).
    #
    # Reuses existing helpers wherever possible so this remains a
    # thin parallel path — not a fork:
    #
    #   _read_zip, _write_zip            (no change)
    #   _detect_data_style               (no change)
    #   _ensure_enough_rows              (no change — column-list
    #                                     parameterized)
    #   _finalize_sheets                 (no change — the SO version
    #                                     already trims trailing rows
    #                                     by count, doesn't care
    #                                     about column names)
    #   _fill_cell, _xml_escape          (no change)
    #
    # New TO-specific code:
    #   _export_to                — orchestrator
    #   _build_shared_strings_to  — TO-shaped strings
    #   _fill_to_header_sheet     — Transfer Header layout
    #   _fill_to_line_sheet       — Transfer Line layout

    def _export_to(
        self,
        result: ProcessingResult,
        template_path: str,
        output_dir: Path,
    ) -> Optional[Path]:
        """
        TO export orchestrator — counterpart of the SO branch in
        :meth:`export`. Caller has already verified ``result.rows``
        is non-empty.

        Args:
            result:        ProcessingResult with output_type='to'.
            template_path: Path to the user-supplied D365 TO template
                            (a real .xlsx, internally a ZIP). The user
                            picks this via the GUI, same UX as SO.
            output_dir:    Where to write the filled TO file.

        Returns:
            Path to the filled file on success, ``None`` on read/write
            errors (errors are logged; GUI surfaces them).
        """
        output_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime('%d-%m-%Y_%H%M%S')
        out_path = output_dir / f"d365_to_import_{timestamp}.xlsx"

        # Work on a copy of the template so the original is never touched.
        try:
            shutil.copy2(template_path, str(out_path))
        except (OSError, shutil.SameFileError) as e:
            logging.error("D365 TO copy failed: %s", e)
            return None

        try:
            zip_contents = self._read_zip(out_path)
        except (KeyError, zipfile.BadZipFile, OSError) as e:
            logging.error("D365 TO template read failed: %s", e)
            out_path.unlink(missing_ok=True)
            return None

        today_str = datetime.now().strftime('%d-%m-%Y')
        unique_pos = self._unique_pos(result.rows)

        logging.info(
            "D365 TO export: %d unique POs, %d total Transfer Lines",
            len(unique_pos), len(result.rows),
        )

        # Step 1: extend sharedStrings with TO-specific literals.
        string_map = self._build_shared_strings_to(
            zip_contents, unique_pos, result.rows, today_str,
        )

        # Step 2: fill Transfer Header (sheet1).
        zip_contents['xl/worksheets/sheet1.xml'] = self._fill_to_header_sheet(
            zip_contents['xl/worksheets/sheet1.xml'].decode('utf-8'),
            unique_pos, today_str, string_map,
        ).encode('utf-8')

        # Step 3: fill Transfer Line (sheet2).
        zip_contents['xl/worksheets/sheet2.xml'] = self._fill_to_line_sheet(
            zip_contents['xl/worksheets/sheet2.xml'].decode('utf-8'),
            result.rows, string_map,
        ).encode('utf-8')

        # Step 4: trim trailing empty rows + fix dimensions/table refs.
        # Reuses the SO finalize helper because it operates by row
        # count, not by column shape — same mechanism works for TO.
        self._finalize_sheets(
            zip_contents, len(unique_pos), len(result.rows),
        )

        # Step 5: write the modified ZIP back to disk.
        try:
            self._write_zip(out_path, zip_contents)
        except (OSError, zipfile.BadZipFile) as e:
            logging.error("D365 TO write failed: %s", e)
            out_path.unlink(missing_ok=True)
            return None

        logging.info("D365 TO export written to %s", out_path)
        return out_path

    @staticmethod
    def _build_shared_strings_to(
        zip_contents: Dict[str, bytes],
        unique_pos: List[SORow],
        all_rows: List[SORow],
        today_str: str,
    ) -> Dict[str, int]:
        """
        TO counterpart of :meth:`_build_shared_strings`.

        Adds the literals our TO writers will reference:

        * ``'PICK'`` — Transfer-from Code (col B on every header)
        * ``'IN TRANSIT'`` — In-Transit Code (col E on every header)
        * ``'false'`` — Direct Transfer flag (col F on every header)
        * ``today_str`` — Posting Date (col D on every header)
        * Each unique Transfer-to Code (FK_BHW_BTS, FK_NLFC01, …)
        * Each Line No. as a string ('10000', '20000', …) — the manual
          dump stores Line No. as a STRING in col B of Transfer Line,
          so we mirror that.

        PO numbers and item numbers go into cells as numeric values
        (not strings), so they're NOT added to shared strings — same
        rationale as the manual dump.

        Returns the {string→index} map used by :func:`_fill_cell` to
        emit ``<v>N</v>`` cell bodies for shared-string cells.
        """
        ss_xml = zip_contents['xl/sharedStrings.xml'].decode('utf-8')

        existing = re.findall(r'<t[^>]*>([^<]*)</t>', ss_xml)
        string_map: Dict[str, int] = {s: i for i, s in enumerate(existing)}

        new_strings: Set[str] = {
            _TO_TRANSFER_FROM_CODE,
            _TO_IN_TRANSIT_CODE,
            _TO_DIRECT_TRANSFER,
            today_str,
        }

        # Transfer-to Codes from each unique PO header.
        for po_row in unique_pos:
            if po_row.ship_to:
                new_strings.add(po_row.ship_to)

        # Line No. strings ('10000', '20000', …, in 10000 increments
        # within each PO). We pre-compute the maximum line number we'll
        # need across all rows so every required value is in the table.
        # Simpler: count rows per PO, generate string for each.
        from collections import Counter
        rows_per_po: Counter = Counter(r.po_number for r in all_rows)
        for count in rows_per_po.values():
            for i in range(1, count + 1):
                new_strings.add(str(i * 10000))

        # Allocate indices for previously-unseen strings.
        next_idx = len(existing)
        for s in sorted(new_strings):
            if s not in string_map:
                string_map[s] = next_idx
                next_idx += 1

        # Rebuild the table body.
        total_count = next_idx
        si_items: List[str] = [''] * total_count
        for s, idx in string_map.items():
            si_items[idx] = f'<si><t>{_xml_escape(s)}</t></si>'

        ss_new = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<sst xmlns="http://schemas.openxmlformats.org/'
            f'spreadsheetml/2006/main" '
            f'count="{total_count}" uniqueCount="{total_count}">'
            + ''.join(si_items)
            + '</sst>'
        )
        zip_contents['xl/sharedStrings.xml'] = ss_new.encode('utf-8')

        return string_map

    def _fill_to_header_sheet(
        self,
        s1_xml: str,
        unique_pos: List[SORow],
        today_str: str,
        string_map: Dict[str, int],
    ) -> str:
        """
        Fill the D365 'Transfer Header' sheet (sheet1).

        One row per unique PO. Rows 1–3 are template headers — data
        begins at row 4 (same row layout as SO template).

        Column layout (matches the manual D365 dump exactly):

            A No.                  PO number (numeric, e.g. 204345116)
            B Transfer-from Code   'PICK'  (constant)
            C Transfer-to Code     ship_to from mapping (e.g.
                                    FK_BHW_BTS). Left empty when the
                                    Ship-To B2B row exists but its
                                    Ship to col is blank (Howrah case)
                                    — the engine has already emitted a
                                    warning for the user.
            D Posting Date         today_str (dd-mm-yyyy)
            E In-Transit Code      'IN TRANSIT' (constant)
            F Direct Transfer      'false'  (constant string)
            G–L                    EMPTY (Dimension cols, optional in
                                    D365 import — the manual dump
                                    leaves them blank so we do too)
        """
        detected_style = self._detect_data_style(
            s1_xml, data_start_row=4, fallback=_HEADER_STYLE_ID,
        )

        s1_xml = self._ensure_enough_rows(
            s1_xml,
            needed=len(unique_pos),
            template_header_rows=3,
            data_start_row=4,
            columns=_TO_HEADER_COLS,
            style_id=detected_style,
            sheet_label='Sheet1',
        )

        for i, po_row in enumerate(unique_pos):
            r = i + 4

            # A: No. (PO number, numeric).
            # The Flipkart dump has int PO numbers like 204345116; the
            # engine's _coerce_po_to_str turned these into strings. For
            # the TO header we want them back as numerics so D365 sees
            # an integer No. cell exactly like the manual dump.
            try:
                po_numeric = int(po_row.po_number)
                s1_xml = _fill_cell(
                    s1_xml, 'A', r, po_numeric, is_string=False,
                )
            except (ValueError, TypeError):
                # Non-numeric PO — fall back to string. Add to map
                # on-the-fly if we missed it during shared-string build.
                val = str(po_row.po_number)
                if val not in string_map:
                    s1_xml = _fill_inline_string(s1_xml, 'A', r, val)
                else:
                    s1_xml = _fill_cell(
                        s1_xml, 'A', r, val,
                        is_string=True, string_map=string_map,
                    )

            # B: Transfer-from Code (constant 'PICK').
            s1_xml = _fill_cell(
                s1_xml, 'B', r, _TO_TRANSFER_FROM_CODE,
                is_string=True, string_map=string_map,
            )

            # C: Transfer-to Code (ship_to). Left empty when blank
            # (Howrah-style row) — engine has already warned.
            if po_row.ship_to:
                s1_xml = _fill_cell(
                    s1_xml, 'C', r, po_row.ship_to,
                    is_string=True, string_map=string_map,
                )

            # D: Posting Date.
            s1_xml = _fill_cell(
                s1_xml, 'D', r, today_str,
                is_string=True, string_map=string_map,
            )

            # E: In-Transit Code (constant 'IN TRANSIT').
            s1_xml = _fill_cell(
                s1_xml, 'E', r, _TO_IN_TRANSIT_CODE,
                is_string=True, string_map=string_map,
            )

            # F: Direct Transfer (constant 'false').
            s1_xml = _fill_cell(
                s1_xml, 'F', r, _TO_DIRECT_TRANSFER,
                is_string=True, string_map=string_map,
            )

            # G–L: Dimension cols left empty (matches manual dump).

        return s1_xml

    def _fill_to_line_sheet(
        self,
        s2_xml: str,
        rows: List[SORow],
        string_map: Dict[str, int],
    ) -> str:
        """
        Fill the D365 'Transfer Line' sheet (sheet2).

        One row per ``SORow``. Rows 1–3 are template headers — data
        begins at row 4.

        Column layout (matches the manual D365 dump exactly):

            A Document No.        PO number (numeric)
            B Line No.            '10000','20000',… as STRING (the
                                   manual dump uses strings here, not
                                   numerics — we mirror it exactly to
                                   avoid any chance of D365 parsing it
                                   differently). Resets per PO.
            C Item No.            item_no (numeric when parseable)
            D Quantity            qty (numeric)
            E Unit of Measure     EMPTY
            F Qty. to Ship        EMPTY
            G Qty. to Receive     EMPTY
            H Dimension Set ID    EMPTY
            I Transfer Price      calc_price (numeric, post-GST cost
                                   from MRP × m% ÷ GST). Rounded to
                                   the same precision as the manual
                                   dump (Python's float repr keeps
                                   ~16 sig digits — same as the dump).
                                   Left empty if calc_price is None
                                   (e.g. row had unparseable MRP).
        """
        detected_style = self._detect_data_style(
            s2_xml, data_start_row=4, fallback=_LINE_STYLE_ID,
        )

        s2_xml = self._ensure_enough_rows(
            s2_xml,
            needed=len(rows),
            template_header_rows=3,
            data_start_row=4,
            columns=_TO_LINE_COLS,
            style_id=detected_style,
            sheet_label='Sheet2',
        )

        current_po: Optional[str] = None
        line_no = 0

        for i, row in enumerate(rows):
            r = i + 4

            # Reset line numbering when PO changes.
            if row.po_number != current_po:
                current_po = row.po_number
                line_no = 0
            line_no += 10000

            # A: Document No. (PO number, numeric).
            try:
                po_numeric = int(row.po_number)
                s2_xml = _fill_cell(
                    s2_xml, 'A', r, po_numeric, is_string=False,
                )
            except (ValueError, TypeError):
                val = str(row.po_number)
                if val not in string_map:
                    s2_xml = _fill_inline_string(s2_xml, 'A', r, val)
                else:
                    s2_xml = _fill_cell(
                        s2_xml, 'A', r, val,
                        is_string=True, string_map=string_map,
                    )

            # B: Line No. — STRING in the manual dump, so we emit as
            # a shared-string reference. Pre-built into string_map by
            # _build_shared_strings_to.
            line_no_str = str(line_no)
            s2_xml = _fill_cell(
                s2_xml, 'B', r, line_no_str,
                is_string=True, string_map=string_map,
            )

            # C: Item No. — numeric when parseable.
            item_numeric: Optional[int] = None
            try:
                item_numeric = int(row.item_no)
            except (TypeError, ValueError):
                item_numeric = None

            if item_numeric is not None:
                s2_xml = _fill_cell(
                    s2_xml, 'C', r, item_numeric, is_string=False,
                )
            else:
                val = str(row.item_no)
                if val not in string_map:
                    s2_xml = _fill_inline_string(s2_xml, 'C', r, val)
                else:
                    s2_xml = _fill_cell(
                        s2_xml, 'C', r, val,
                        is_string=True, string_map=string_map,
                    )

            # D: Quantity (numeric).
            s2_xml = _fill_cell(
                s2_xml, 'D', r, row.qty, is_string=False,
            )

            # E–H: Unit of Measure / Qty.to Ship / Qty.to Receive /
            # Dimension Set ID — left empty per manual dump.

            # I: Transfer Price (numeric, post-GST cost). Skip when
            # missing (would happen if MRP couldn't be parsed; engine
            # already skips those rows but defensive guard kept here).
            price = row.calc_price
            if price is not None:
                # Python's float repr matches the manual dump's
                # precision (~16 sig digits, e.g. 330.5084745762712).
                # Don't round — D365 accepts any precision in this col.
                s2_xml = _fill_cell(
                    s2_xml, 'I', r, price, is_string=False,
                )

        return s2_xml

    # ── Row-count management ───────────────────────────────────────────

    @staticmethod
    def _detect_data_style(
        xml: str, data_start_row: int, fallback: str,
    ) -> str:
        """
        Find the style id used by data cells in a template sheet.

        The D365 sample templates reuse a single style id for every
        cell in the data region. We can detect it reliably by looking
        at any cell in ``data_start_row`` (row 4 by convention). If
        no match is found (e.g. the row is entirely missing in the
        template XML), fall back to the supplied default.

        This matters because hardcoded style ids like ``11`` and
        ``8`` were specific to the D365 template GT Mass used —
        different templates use different numbers, and using the
        wrong id produces an output file that Excel can open with a
        "repairing" dialog but that openpyxl rejects outright.

        Args:
            xml:            Sheet XML to scan.
            data_start_row: Row number containing data (usually 4).
            fallback:       Default style id to return if nothing
                            found.

        Returns:
            String style id (digits only), or ``fallback``.
        """
        # Match any <c r="X4" s="N" ...> cell; first hit wins.
        match = re.search(
            rf'<c r="[A-Z]+{data_start_row}"[^>]*s="(\d+)"', xml,
        )
        return match.group(1) if match else fallback

    @staticmethod
    def _ensure_enough_rows(
        xml: str,
        needed: int,
        template_header_rows: int,
        data_start_row: int,
        columns: Iterable[str],
        style_id: str,
        sheet_label: str,
    ) -> str:
        """
        Inject additional pre-styled ``<row>`` elements when the data
        exceeds the template's pre-formatted capacity.

        The D365 sample template comes with a fixed number of empty
        styled rows (typically ~30 for Headers, ~50 for Lines). When
        our batch has more POs/rows than that, our ``_fill_cell``
        regex would otherwise find nothing to replace and silently
        drop the extra rows. This method plugs that gap up-front.

        Args:
            xml:                   Current sheet XML.
            needed:                Number of data rows we want to
                                   write.
            template_header_rows:  How many template rows exist above
                                   ``data_start_row`` — subtracted
                                   from the total row count to get
                                   the existing data-row capacity.
            data_start_row:        1-based row number where data
                                   starts (usually 4).
            columns:               Column letters that need a cell
                                   slot per row (e.g. ``'A'..'R'`` for
                                   Headers).
            style_id:              ``s="N"`` style id matching the
                                   template's existing styled cells.
            sheet_label:           Human-readable sheet label for log
                                   messages only.

        Returns:
            Updated XML with extra rows inserted as needed.

        v1.5.2 bugfix
        -------------
        Original implementation used ``len(re.findall(r'<row r="(\\d+)"'))``
        as ``existing_rows``, then subtracted ``template_header_rows``
        to derive capacity. This broke when the template's XML skipped
        a row number (e.g. ``BL_Sample.xlsx`` Sheet1 omits ``<row r="2">``
        because that row is visually blank in the template). A 13-entry
        row list with rows numbered 1, 3, 4..14 gives a highest row
        number of 14 and actual data capacity 11 (rows 4-14), but the
        old code computed capacity = 13 - 3 = 10, mis-counting by 1
        and dropping the last PO. Now we use the maximum ``r``
        attribute across all existing ``<row>`` elements to determine
        the template's true extent, then figure capacity as
        ``max_row - data_start_row + 1``.
        """
        row_nums = [int(x) for x in re.findall(r'<row r="(\d+)"', xml)]
        max_existing_row = max(row_nums) if row_nums else 0
        existing_capacity = max(0, max_existing_row - data_start_row + 1)

        if needed <= existing_capacity:
            return xml

        logging.info(
            "D365 %s: template has %d data rows (max row=%d), "
            "need %d — injecting %d",
            sheet_label, existing_capacity, max_existing_row, needed,
            needed - existing_capacity,
        )

        # Append new rows right before </sheetData>. Each row has an
        # empty pre-styled cell in every column so _fill_cell's regex
        # can find-and-replace them later. First new row starts
        # immediately after the existing max row number — NOT after
        # the row count — so we don't create duplicate row indices
        # when the template skipped row numbers.
        columns_list = list(columns)
        new_row_count = needed - existing_capacity
        first_new_row = max_existing_row + 1

        new_rows: List[str] = []
        for offset in range(new_row_count):
            row_num = first_new_row + offset
            cells = ''.join(
                f'<c r="{col}{row_num}" s="{style_id}"/>'
                for col in columns_list
            )
            new_rows.append(
                f'<row r="{row_num}" spans="1:{len(columns_list)}" '
                f'x14ac:dyDescent="0.3">{cells}</row>'
            )

        return xml.replace('</sheetData>', ''.join(new_rows) + '</sheetData>')

    # ── Final trim ─────────────────────────────────────────────────────

    @staticmethod
    def _finalize_sheets(
        zip_contents: Dict[str, bytes],
        header_count: int,
        line_count: int,
    ) -> None:
        """
        Remove rows beyond our last data row and fix dimension/table
        refs so D365 doesn't flag "extra blank rows" on import.

        Sheets to clean:
            sheet1 — Headers. Keep rows 1..(3 + header_count).
            sheet2 — Lines.   Keep rows 1..(3 + line_count).

        Also updates ``xl/tables/table1.xml`` and ``table2.xml`` ref
        attributes to match the trimmed row range.
        """
        last_hdr = 3 + header_count
        last_line = 3 + line_count

        # ── Trim Headers sheet ──
        s1 = zip_contents['xl/worksheets/sheet1.xml'].decode('utf-8')
        s1 = _remove_rows_beyond(s1, last_hdr)
        s1 = re.sub(
            r'<dimension ref="[^"]*"/>',
            f'<dimension ref="A1:R{last_hdr}"/>',
            s1,
        )
        zip_contents['xl/worksheets/sheet1.xml'] = s1.encode('utf-8')

        # ── Trim Lines sheet ──
        s2 = zip_contents['xl/worksheets/sheet2.xml'].decode('utf-8')
        s2 = _remove_rows_beyond(s2, last_line)
        s2 = re.sub(
            r'<dimension ref="[^"]*"/>',
            f'<dimension ref="A1:H{last_line}"/>',
            s2,
        )
        zip_contents['xl/worksheets/sheet2.xml'] = s2.encode('utf-8')

        # ── Update table1 / table2 refs so the ListObject bounds match
        # the trimmed data. Without this, D365 complains the table
        # range exceeds the sheet extent and rejects the package.
        if 'xl/tables/table1.xml' in zip_contents:
            t1 = zip_contents['xl/tables/table1.xml'].decode('utf-8')
            t1 = re.sub(
                r'ref="A3:[A-Z]+\d+"',
                f'ref="A3:R{last_hdr}"',
                t1,
            )
            zip_contents['xl/tables/table1.xml'] = t1.encode('utf-8')

        if 'xl/tables/table2.xml' in zip_contents:
            t2 = zip_contents['xl/tables/table2.xml'].decode('utf-8')
            t2 = re.sub(
                r'ref="A3:[A-Z]+\d+"',
                f'ref="A3:H{last_line}"',
                t2,
            )
            zip_contents['xl/tables/table2.xml'] = t2.encode('utf-8')


# ═══════════════════════════════════════════════════════════════════════
#  Module-level XML utilities
# ═══════════════════════════════════════════════════════════════════════

def _xml_escape(s: str) -> str:
    """Minimal XML text escaping for shared-string content."""
    return (
        s.replace('&', '&amp;')
         .replace('<', '&lt;')
         .replace('>', '&gt;')
    )


def _fill_cell(
    xml: str,
    col_letter: str,
    row_num: int,
    value,
    is_string: bool,
    string_map: Optional[Dict[str, int]] = None,
) -> str:
    """
    Replace one empty pre-styled cell with a filled one.

    Template cells look like ``<c r="A4" s="11"/>`` — self-closed with
    only a style attribute. We substitute them with populated forms:

    * String: ``<c r="A4" s="11" t="s"><v>INDEX</v></c>``
      where INDEX is the 0-based shared-string index.

    * Numeric: ``<c r="A4" s="11"><v>VALUE</v></c>``
      no ``t`` attribute (implicit number).

    We use a regex anchored to the exact cell reference and capture the
    style id so it's preserved through the edit. ``count=1`` keeps the
    replace from accidentally matching a lookalike cell elsewhere.

    Args:
        xml:         Sheet XML to edit.
        col_letter:  Column letter (e.g. ``'A'``).
        row_num:     1-based row number.
        value:       Value to write (string or number).
        is_string:   If True, look up ``value`` in ``string_map`` and
                     write a shared-string reference; otherwise write
                     as a literal number.
        string_map:  Required when ``is_string=True``.

    Returns:
        Updated XML with the cell filled. If the cell reference
        doesn't appear in the XML (e.g. template had fewer rows than
        expected and row injection wasn't done), returns the XML
        unchanged — this should not normally happen because
        :meth:`_ensure_enough_rows` guards against it.

    v1.5.2 bugfix
    -------------
    The original implementation only matched empty self-closing cells
    (``<c r="A4" s="11"/>``), which is what ``_ensure_enough_rows``
    produces when it injects new rows. But real D365 sample templates
    (like the ``BL_Sample.xlsx`` provided by the user) ship with the
    first N rows already pre-populated with sample data, e.g.::

        <c r="A4" s="5" t="s"><v>28</v></c>

    Those pre-filled cells didn't match the old regex, so our fills
    silently no-op'd and the user got an output with the template's
    sample-row values leaking into Sales Header rows. Excel then
    flagged "Workbook Repaired" because the <v> indices pointed at
    shared-string slots that got renumbered during our rebuild.

    The fix: try the empty-cell regex first (fast path for injected
    rows), and if that doesn't match, try a broader regex that
    matches any ``<c r="REF" ...>...</c>`` variant and replaces it
    wholesale. Either way, the resulting cell is fully replaced with
    our desired content.
    """
    ref = f'{col_letter}{row_num}'
    style_id_default = '11' if col_letter in 'ABCDEFGHIJKLMNOPQR' else '8'

    # ── Fast path: match empty pre-styled cell (self-closing) ──
    empty_pattern = f'<c r="{ref}" s="(\\d+)"\\s*/>'

    if is_string:
        if string_map is None:
            raise ValueError(
                "string_map is required when is_string=True"
            )
        idx = string_map.get(str(value), 0)
        empty_replacement = f'<c r="{ref}" s="\\1" t="s"><v>{idx}</v></c>'
    else:
        empty_replacement = f'<c r="{ref}" s="\\1"><v>{value}</v></c>'

    new_xml, n_subs = re.subn(empty_pattern, empty_replacement, xml, count=1)
    if n_subs > 0:
        return new_xml

    # ── Slow path: cell is pre-populated (template ships with sample
    # data, e.g. <c r="A4" s="5" t="s"><v>28</v></c>). We replace the
    # whole element, preserving its style id so formatting stays
    # consistent. Two subcases: self-closing with attrs only (rare but
    # possible after openpyxl rewrites) and with content. ──
    prefilled_pattern = (
        # Capture group 1 = style id
        f'<c r="{ref}"([^>]*s="(\\d+)"[^>]*)'
        # Then either self-closing /> or open>...</c>
        f'(?:/>|>.*?</c>)'
    )

    def _replace(match: re.Match) -> str:
        style = match.group(2)
        if is_string:
            idx = string_map.get(str(value), 0)
            return f'<c r="{ref}" s="{style}" t="s"><v>{idx}</v></c>'
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'

    new_xml = re.sub(prefilled_pattern, _replace, xml, count=1)
    return new_xml


def _fill_inline_string(
    xml: str,
    col_letter: str,
    row_num: int,
    value: str,
) -> str:
    """
    Fallback: write a cell as an *inline* string rather than via the
    shared-string table.

    Inline strings are valid OOXML but slightly larger on disk. We
    use them only for the rare case where a value wasn't registered
    in the shared-string map ahead of time (e.g. an oddly-typed
    item_no on the Lines sheet).
    """
    ref = f'{col_letter}{row_num}'
    pattern = f'<c r="{ref}" s="(\\d+)"\\s*/>'
    replacement = (
        f'<c r="{ref}" s="\\1" t="inlineStr">'
        f'<is><t>{_xml_escape(value)}</t></is></c>'
    )
    return re.sub(pattern, replacement, xml, count=1)


def _remove_rows_beyond(xml: str, max_row: int) -> str:
    """
    Remove every ``<row r="N" …>…</row>`` element where ``N`` exceeds
    ``max_row``.

    Runs on a serialized XML string so we can do it without loading a
    full DOM. The regex is greedy-on-contents but safely anchored on
    the row open/close tags.

    Args:
        xml:     Sheet XML.
        max_row: Highest row number to keep. All rows with
                 ``r > max_row`` are dropped.

    Returns:
        Sheet XML with excess rows removed.
    """
    def _replacer(match: re.Match) -> str:
        row_num = int(match.group(1))
        return '' if row_num > max_row else match.group(0)

    return re.sub(
        r'<row r="(\d+)"[^>]*>.*?</row>',
        _replacer,
        xml,
        flags=re.DOTALL,
    )