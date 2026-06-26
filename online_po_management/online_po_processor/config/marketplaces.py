"""
config.marketplaces
===================

Per-marketplace configuration registry.

Each marketplace produces PO/punch files with its own column layout. To
support a new marketplace, add an entry to ``MARKETPLACE_CONFIGS`` with
the keys documented below — the rest of the pipeline is config-driven and
needs no further code changes for the common cases.

Config schema
-------------

Required for every marketplace:

``party_name`` (str)
    Must match the ``Party`` column in the Ship-To B2B mapping sheet
    exactly (case-insensitive). Used to filter the mapping registry to
    just this marketplace's locations.
``po_col`` (str | list[str])
    Column containing the PO/SO number. Use a plain string for the
    common case of a single known column name. Use a list when the
    marketplace's punch file sometimes arrives with variant headers
    — the engine will pick the first list entry that actually
    appears in the file. List order is preference order; e.g.
    ``['PO', 'PO Number']`` means "use 'PO' when present, fall back
    to 'PO Number'". Myntra uses this because its dumps sometimes
    label the column 'PO' and sometimes 'PO Number'.
``loc_col`` (str | list[str])
    Column containing the delivery location (matched against ``Del
    Location`` in the mapping sheet). Supports the same list
    fallback as ``po_col``.
``qty_col`` (str | list[str])
    Column containing the order quantity. Supports list fallback.
``item_resolution`` (str, ``'from_column'`` | ``'from_ean'``, default
``'from_column'``)
    How to determine the canonical Item No for each row:

    * ``'from_column'`` — take it directly from ``item_col``. Use when the
      marketplace already provides a pre-resolved Item No.
    * ``'from_ean'`` — look the EAN up in Items_March and use
      ``master_info['item_no']``. Use when the marketplace only provides
      EAN/GTIN (real Myntra and RK files).

``item_col`` (str, required when ``item_resolution='from_column'``)
    Column containing the Item No.
``ean_col`` (str, required when ``item_resolution='from_ean'``)
    Column containing the EAN/GTIN. Also used for price-validation lookups
    even when ``item_resolution='from_column'``.

Optional but recommended:

``fob_col`` (str)
    Column containing the marketplace price to validate against.
``ref_fob_col`` (str)
    Optional second marketplace price column shown only as a *reference*
    Diffn in Raw Data (no effect on OK/MISMATCH status).
``compare_basis`` (str, ``'landing'`` | ``'cost'``, default ``'cost'``)
    What we compare ``fob_col`` against:

    * ``'landing'`` → ``MRP × margin%`` (pre-GST). Used by Myntra because
      its "Landing Price" column is the pre-GST value.
    * ``'cost'`` → ``MRP × margin% ÷ GST`` (post-GST). Used by RK.

``compare_label`` (str, default ``'Price'``)
    Friendly label shown in the Validation sheet. E.g. ``'Landing Rate'``
    yields ``"Marketplace Landing Rate"`` and ``"Difference with Landing
    Rate"`` column headers.
``amount_col`` (str | dict | None, default omitted)
    Per-row monetary amount source. When configured, the engine reads
    this into ``SORow.amount`` and the email report sums it for the
    headline "Amount" stat.

    Accepted forms:

    1. ``str`` — single column name on the punch file::

           'amount_col': 'total_amount'            # Blink
           'amount_col': 'Total accepted cost'     # RK (qty × cost)

    2. ``{'multiply': [col_a, col_b, ...]}`` — product of columns.
       Used when the marketplace doesn't carry a pre-calculated
       amount but does carry the factors (added v1.5.7)::

           'amount_col': {'multiply': ['Landing Price', 'Quantity']}
           # Myntra — landing value pre-tax

    3. ``{'multiply': [...], 'apply_margin': True}`` — product of
       columns, then multiplied by the run's ``margin_pct`` (v1.6.0).
       Used when one conceptual factor is the derived Landing Cost
       rather than a raw column::

           'amount_col': {'multiply': ['MRP', 'Qty'], 'apply_margin': True}
           # Reliance — Landing × Qty = (MRP × margin%) × Qty

    Omit the key entirely when the marketplace has no usable amount
    source; the email aggregator treats absent ``amount`` as 0.

``hsn_col`` (str, optional, v1.6.0)
    Column on the punch file carrying the HSN/SAC code. When set,
    the engine cross-checks it against the master's ``HSN/SAC Code``
    column and records an ``hsn_check_status`` on each ``SORow``
    (``'OK'``, ``'MISMATCH'``, or ``'NOT_IN_MASTER'``). Mismatches
    produce a deduped per-item warning and surface on the Validation
    sheet. Currently Reliance only. Other marketplaces carry HSN
    columns but we don't cross-check them by default because their
    HSN data isn't contractually authoritative the way Reliance's is.

``source_sheet`` (str, default ``'Sheet1'``)
    Name of the sheet to read inside the punch workbook. Supports
    two match modes:

    * **Exact** — plain sheet name. Reliance uses ``'PO'``; the
      default ``'Sheet1'`` covers Blink/Myntra/RK.
    * **Wildcard prefix** (v1.8.0) — trailing ``*`` matches any
      sheet whose name starts with the given prefix. Zepto uses
      ``'PO_*'`` because its data sheet is literally named
      ``'PO_<random-hex>'`` (different hex every dump).

    Behavior on miss:

    * Exact miss falls back to the first sheet in the workbook and
      logs a warning (kind to files with reordered sheets).
    * Wildcard miss aborts with an explicit error — the wildcard
      implies a specific marketplace's expected sheet shape, so
      silently reading a different sheet would produce garbage.

    Only consulted when ``source_format`` is ``'excel'`` (the default).
    Ignored for PDF marketplaces.

``source_format`` (str, ``'excel'`` | ``'pdf'``, default ``'excel'``, v2.2.0)
    Selects the file-loading path for this marketplace:

    * ``'excel'`` — the original path. Engine calls ``pd.read_excel``
      on the configured ``source_sheet`` with the configured
      ``header_row``. Used by Myntra, RK, Blink, Reliance, Zepto,
      BlinkMP.
    * ``'pdf'``   — engine calls the PDF parser registered under
      ``pdf_parser`` (see below). Used by Avenue/DMart Ready.

    PDF parsers MUST return a ``pandas.DataFrame`` shaped the same
    way the marketplace's other config keys reference (``po_col``,
    ``ean_col``, etc.) — the engine's downstream logic is
    format-agnostic.

``pdf_parser`` (str, required when ``source_format='pdf'``, v2.2.0)
    Key into the engine's ``PDF_PARSERS`` registry. Each registered
    parser is a callable ``(filepath: str) -> pd.DataFrame``. To add
    a new PDF marketplace:

    1. Write a parser module (e.g. ``engine/avenue_pdf_parser.py``)
       exposing a ``load_<name>_pdf_as_dataframe(filepath)`` function.
    2. Register it in ``PDF_PARSERS`` in ``marketplace_engine.py``.
    3. Set ``source_format='pdf'`` and ``pdf_parser='<name>'`` on the
       marketplace's config.

    Current parsers:

    * ``'avenue'`` — Avenue E-Commerce Ltd / DMart Ready PO PDFs.
      Two-page layout, 2-row-per-item table, header block with PO
      number / GST / ship-to / vendor code.

``accepted_extensions`` (list[str], optional, v2.3.0)
    Explicit list of file extensions the marketplace's PO uploads can
    use, in display priority order. The list affects ONLY the GUI file
    dialog filter — the engine itself always auto-detects file type
    by extension at load time (``.csv`` → ``pd.read_csv``,
    ``.xlsx``/``.xls``/``.xlsm`` → ``pd.read_excel``, ``.pdf`` →
    registered PDF parser).

    Use when a marketplace's dashboard exports POs in MORE than one
    format. Blink is the canonical case: the new dashboard exports
    CSV, but older xlsx exports are still in circulation on operators'
    disks. Setting ``accepted_extensions: ['.csv', '.xlsx', '.xls']``
    advertises all three in the file dialog so the operator can pick
    whichever they have.

    When unset (the common case), the dialog filter falls back to
    ``source_format``:
    * ``'pdf'``   → ``*.pdf`` only
    * anything else → ``*.xlsx`` only (the historical default)

``header_row`` (int, default ``0``)
    0-indexed row that pandas treats as column headers. Reliance's
    PO sheet has a merged title on row 0 so its real headers are on
    row 1. Most marketplaces leave this at 0.


``default_margin`` (int | float, default 70)
    Default margin % pre-filled in the GUI when this marketplace is
    selected. The user can override per-run via the GUI input.
    Typically an integer (70 for Blink/RK/Myntra); Reliance uses
    the decimal ``63.42``.

``case_insensitive_cols`` (bool, default False, v1.8.1)
    When True, the engine's column-name resolution matches ``*_col``
    values against the DataFrame's actual headers **case-insensitively**
    and **whitespace-tolerantly** (surrounding whitespace is ignored,
    internal multi-space runs collapse to one). Protects against
    marketplace dashboards that ship the same semantic column under
    varying casings/spacings across dumps. Currently enabled on:

    * **Myntra** — PO header seen as ``'PO'``, ``'PO Number'``, and
      ``'Po number'`` across recent dumps.
    * **Reliance** — HSN header seen as ``'HSN'`` (older batches)
      and ``'Hsn'`` (newer batches).

    Omit/False for marketplaces with stable, contractual column
    headers (Blink/RK/Zepto today) so typo-in-template errors fail
    loud rather than silently match something unintended.

``override_unit_price`` (bool, default False, v1.9.1)
    When True, the D365 exporter populates Sales Line col H (Unit
    Price) with our computed post-GST Cost Price per row rather than
    leaving it blank. Blank is the normal default because the ERP
    auto-computes unit price from the vendor master's margin
    setting downstream — but when the ERP's recorded margin for the
    vendor doesn't match the marketplace's actual margin, the
    auto-computed cost is wrong and has to be overridden.

    Canonical example: BCPL (Blinkit + BlinkMP share a single vendor
    record in Business Central with 70% margin — Blinkit's rate).
    BlinkMP operates at 75%, so without this override every BlinkMP
    row in the D365 import would post with Blinkit's 70% cost,
    understating BlinkMP's margin. Enabling the flag stamps the
    correct cost (computed via the engine's standard
    ``MRP × margin% ÷ GST-divisor`` formula, per-item GST rate) into
    col H so the ERP records what BlinkMP actually charged us.

    Rows without a computable Cost Price (master lookup failed)
    still emit with col H blank — the engine logs a per-row warning
    and continues. Matches the lenient behavior elsewhere in the
    pipeline for missing-master rows.

``price_col`` (str | None)
    Column containing unit price for the SO Lines output (rare — both
    current marketplaces leave it None so the WMS computes it).
``template_headers`` (list[str])
    Full column list used when the user clicks "Download PO Template". If
    omitted, a minimal list is built from the required + validation cols.

PO template colour legend
-------------------------
When the user downloads a template, headers are colour-coded by role:

* **BLUE** (``#1A237E``) — Required. Script fails without these.
* **GREEN** (``#1B5E20``) — Validation. Used for price check & master
  lookup.
* **GREY** (``#9E9E9E``) — Not read by script. Kept only to mirror the
  marketplace's native file format.
"""

from __future__ import annotations
from typing import Any, Dict, List


MARKETPLACE_CONFIGS: Dict[str, Dict[str, Any]] = {
    # ────────────────────────────────────────────────────────────────────
    # MYNTRA
    # Real Myntra files have NO 'Item no' column — only 'GTIN' and
    # 'Vendor Article Number' (both carry the EAN). Item No is resolved
    # from EAN via the Items_March master.
    # ────────────────────────────────────────────────────────────────────
    'Myntra': {
        'party_name': 'Myntra',
        # v2.4.0: Myntra is the first DUAL-FORMAT marketplace — the
        # operator can feed either the dashboard Excel "punch" (the
        # historical path, unchanged) OR the PO PDF Myntra emails. We
        # keep ``source_format`` at its 'excel' default and add a
        # ``pdf_parser`` + ``accepted_extensions``: the engine routes a
        # ``.pdf`` upload to ``myntra_pdf_parser`` by file extension and
        # an ``.xlsx`` upload down the existing Excel path. The PDF parser
        # emits the SAME column names referenced below (GTIN / Quantity /
        # Landing Price / List price… / Mrp / Location), so this single
        # config drives both formats. The PDF also injects real PO dates
        # (__po_date__ / __exp_date__) the Excel punch lacks.
        #
        # 2026-06-25: REVERTED to Excel-primary. The operator now manually
        # compiles Myntra's PO PDFs into one accurate ``dump.xlsx`` (the PDF
        # parser occasionally mis-reads the line grid — e.g. two EANs merged on
        # a page wrap), so ``.xlsx`` LEADS ``accepted_extensions`` (the file
        # picker defaults to Excel; the compiled dump holds ALL POs in one file).
        # ``pdf_parser='myntra'`` is KEPT as a still-selectable FALLBACK only —
        # the same config drives both, since the PDF parser emits the same column
        # names the Excel carries (GTIN / Quantity / Landing Price / Mrp /
        # Location). Each PDF is one PO, so Myntra stays multi-file; the
        # single-dump Excel flows through process_multi unchanged.
        'pdf_parser': 'myntra',
        'accepted_extensions': ['.xlsx', '.pdf'],
        # v1.8.1: Myntra's dashboard exports have historically varied
        # the case of column headers between dumps. We've seen at
        # least three casings of the PO column header ('PO', 'PO
        # Number', 'Po number') in a two-week span. Enabling
        # case-insensitive column matching absorbs these variations
        # automatically — the engine's ``_resolve_column_aliases``
        # will find whichever header exists regardless of case when
        # this flag is True.
        'case_insensitive_cols': True,
        # v1.5.5: ``po_col`` is a list because Myntra dumps arrive
        # with either 'PO' or 'PO Number' as the header — list order
        # is the preference, so 'PO' wins when both exist. The engine
        # collapses this to a single string after loading the file
        # (see MarketplaceEngine._resolve_column_aliases).
        'po_col': ['PO', 'PO Number'],                        # [REQUIRED]
        'loc_col': 'Location',                                # [REQUIRED]
        'qty_col': 'Quantity',                                # [REQUIRED]
        'item_resolution': 'from_ean',                        # see schema
        # v1.5.8: EAN source switched from 'Vendor Article Number' to
        # 'GTIN'. Semantically GTIN is the EAN — "Vendor Article
        # Number" is Myntra's field for the seller's internal SKU
        # code, and even though Renee has historically populated it
        # with the EAN (so the two columns matched 237/242 rows in
        # the 17-04 file), relying on that is fragile — Myntra could
        # diverge them any time. GTIN is the semantically correct
        # column. Rows where GTIN is blank or unreadable follow the
        # standard missing-EAN warning flow (NOT_IN_MASTER in the
        # Validation sheet + a row-level warning).
        'ean_col': 'GTIN',                                    # [REQUIRED in this mode]
        'price_col': None,                                    # WMS computes
        'fob_col': 'Landing Price',                           # [VALIDATION]
        # ref_fob_col surfaces "what would the diff have been against List
        # price?" alongside the active diff. Reference only — has zero
        # effect on OK/MISMATCH status.
        'ref_fob_col': 'List price(FOB+Transport-Excise)',
        # v1.5.7: Myntra dumps don't carry a pre-calculated per-row
        # amount column (unlike Blink's 'total_amount' or RK's 'Total
        # accepted cost'), but they do carry the factors. We compute
        # Landing × Qty per row which is the pre-tax landing value
        # — matches what the finance team reconciles against on the
        # Myntra invoice. The multiply-spec form is handled by the
        # engine's ``_extract_amount`` helper.
        'amount_col': {'multiply': ['Landing Price', 'Quantity']},
        'default_margin': 70,
        'compare_basis': 'landing',                           # MRP × m%
        'compare_label': 'Landing Rate',
        # v2.4.4: status reflects BOTH the landing pair AND the cost pair —
        # a row is OK only when vendor Landing matches our landing (MRP×m%)
        # AND vendor CP (the 'List price' ref column) matches our CP
        # (MRP×m%÷GST). Either failing → MISMATCH. Deal/vendor-CP exception
        # rows (e.g. Goddess) are exempt.
        'also_check_cost': True,
        'template_headers': [
            'PO/PO Number', 'Location', 'SKU Id', 'Style Id', 'SKU Code',
            'HSN Code', 'Brand', 'GTIN', 'Vendor Article Number',
            'Vendor Article Name', 'Size', 'Colour', 'Mrp',
            'Credit Period', 'Margin Type', 'Agreed Margin',
            'Gross Margin', 'Quantity', 'FOB Amount',
            'List price(FOB+Transport-Excise)', 'Landing Price',
            'Estimated Delivery Date',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # RK
    # Real RK files also lack an 'Item no' column — they expose EAN as
    # 'External ID'. Same resolution mechanism as Myntra. Compare basis
    # is 'cost' because RK's Cost column is post-GST, matching our
    # MRP × 70% ÷ 1.18 to the paisa.
    # ────────────────────────────────────────────────────────────────────
    'RK': {
        'party_name': 'RK',
        'po_col': 'PO',                                       # [REQUIRED]
        'loc_col': 'Ship-to location',                        # [REQUIRED]
        'qty_col': 'Accepted quantity',                       # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'External ID',                             # [REQUIRED in this mode]
        'price_col': None,
        'fob_col': 'Cost',                                    # [VALIDATION]
        'default_margin': 70,
        'compare_basis': 'cost',                              # MRP × m% ÷ GST
        'compare_label': 'Cost',
        'amount_col': 'Total accepted cost',                  # qty × cost — for email Amount stat
        # v2.3.1: RK's 'Total accepted cost' is the PRE-GST order value
        # (qty × pre-GST Cost). Flagging it tells the Summary sheet to
        # add a 'Total Amount (Inc GST)' column and the Tracker sheet to
        # report Order Value GST-inclusive — matching how every other
        # marketplace's amount already includes tax. Inc-GST is computed
        # per line as amount × (1 + line GST) via MasterLoader.gst_divisor.
        'amount_is_pre_gst': True,
        # Tracker sheet (paste-ready ops view). Exp Date maps to RK's
        # 'Cancellation deadline' (the PO's effective expiry); PO Date to
        # 'Order date'. See exporter/sheets/tracker_sheet.py.
        'template_headers': [
            'PO', 'Vendor code', 'Order date', 'Product name',
            'External ID', 'Accepted quantity', 'Ship-to location',
            'Cost', 'Total accepted cost',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # Blink  (a.k.a. Blinkit — Blink Commerce Private Limited, "BCPL".
    # The dropdown and party_name both use "Blink" to match what's in
    # the Ship-To B2B registry — mapping lookups filter on exactly that
    # string, so the config must agree with the sheet character-for-
    # character.)
    #
    # Real Blink files carry line data on 'Sheet1' (same as all other
    # marketplaces — see note in engine.marketplace_engine). They may
    # also ship pivot / sidecar sheets (Sheet2, Sheet4) with the user's
    # own reference data; those are intentionally ignored by the engine.
    #
    # Columns on Sheet1 (25 native):
    #   po_number | facility_name | manufacturer_name |
    #   entity_vendor_legal_name | vendor_name | order_date |
    #   appointment_date | expiry_date | po_state | item_id | name |
    #   uom_text | upc | units_ordered | remaining_quantity |
    #   landing_rate | cost_price | margin_percentage | cess_value |
    #   sgst_value | igst_value | cgst_value | tax_value |
    #   total_amount | mrp
    #
    # Key specifics:
    #   • No 'Item no' column → resolve from 'upc' (EAN) via master lookup.
    #   • 'margin_percentage' = 30 is Blink's margin → ours is 70%.
    #   • 'cost_price' is POST-GST (landing ÷ 1.18), compare_basis='cost'.
    #   • 'landing_rate' is pre-GST (MRP × 70%), present but not used for
    #     validation — could be added as ref_fob_col if reference diff is
    #     wanted later.
    #
    # 'Blink RO' (reverse-order / return entity) will be added as a
    # sibling entry once we have a sample file.
    # ────────────────────────────────────────────────────────────────────
    'Blink': {
        'party_name': 'Blink',               # Must match 'Party' in mapping sheet
        # v2.3.0: Blink now exports POs in TWO formats from its vendor
        # dashboard — a CSV (newer, default download) and the historical
        # xlsx. Both share IDENTICAL column names (po_number,
        # units_ordered, cost_price, total_amount, mrp, upc,
        # facility_name, ...), so the same config drives both. The
        # engine auto-detects file type by extension at load time:
        # .csv → pandas.read_csv, .xlsx/.xls → pandas.read_excel.
        # No code path branches on this config key — it only affects
        # the GUI file picker filter, advertising what the operator
        # can upload.
        'accepted_extensions': ['.csv', '.xlsx', '.xls'],
        'po_col': 'po_number',               # [REQUIRED] int64 (e.g. 1723710027417)
        'loc_col': 'facility_name',          # [REQUIRED] e.g. 'Pune P2 - Feeder Warehouse'
        'qty_col': 'units_ordered',          # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'upc',                    # [REQUIRED in this mode]
        'price_col': None,                   # WMS computes
        'fob_col': 'cost_price',             # [VALIDATION] post-GST
        'mrp_col': 'mrp',                    # [VALIDATION] file's MRP → "Vendor MRP" vs "Our MRP"
        'default_margin': 70,                # 100 - Blink's margin_percentage(30)
        'compare_basis': 'cost',             # cost_price ≈ MRP × 70% ÷ 1.18
        'compare_label': 'Cost',
        'amount_col': 'total_amount',        # incl. taxes — for email Amount stat
        'template_headers': [
            'po_number', 'facility_name', 'manufacturer_name',
            'entity_vendor_legal_name', 'vendor_name',
            'order_date', 'appointment_date', 'expiry_date', 'po_state',
            'item_id', 'name', 'uom_text', 'upc',
            'units_ordered', 'remaining_quantity',
            'landing_rate', 'cost_price', 'margin_percentage',
            'cess_value', 'sgst_value', 'igst_value', 'cgst_value',
            'tax_value', 'total_amount', 'mrp',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # Reliance (v2.3.1 — REWRITTEN as PDF-based; replaces the old Excel/
    # pre_process flow entirely)
    #
    # Reliance Retail now sends its Seller Purchase Order as a multi-page
    # PDF ("PO. INTM_ <no> .PDF Reliance Retail Limited.PDF"). Parsed by
    # engine/reliance_pdf_parser.py (``pdf_parser='reliance'``), which
    # reads the line-item table via pdfplumber.extract_tables() and
    # exposes synthetic ``__po__`` / ``__loc__`` columns. Standard SO.
    #
    # PDF shape: page-1 header (PO NO., Site, PO Date, Delivery Address),
    # page-2+ line-item table (Sr.No | Article No./HSN | EAN/Vendor
    # Article | Material Description | Qty | UOM | MRP | Base Cost |
    # IGST(%) … | Total Base Value).
    #
    # Ship-to: the delivery city (e.g. 'BHIWANDI') is parsed from the
    # Delivery Address; the mapping's fuzzy/substring tier resolves it to
    # the Ship-To B2B key ('BHIWANDI (Reliance)', '…-Indore', etc., all
    # Cust 20015).
    #
    # ── GST-DEPENDENT PRICING (the key change) ──────────────────────────
    # Reliance's margin is 31% off the PRE-GST value, so the effective
    # "multiplier on MRP" varies with the item's GST rate:
    #     keep% = 1 − 0.31 × (1 + GST)   → 69.00% @0%, 67.45% @5%, 63.42% @18%
    # ``gst_margin_discount: 0.31`` tells the engine to compute this
    # per-item margin from the master's GST (instead of a fixed margin),
    # then ``compare_basis='cost'`` gives the post-GST cost
    #     MRP × keep% ÷ (1+GST)
    # which equals Reliance's pre-GST 'Base Cost' (the fob_col). Verified:
    # MRP 250 @18% → 250 × 0.6342 ÷ 1.18 = 134.36.
    #
    # HSN cross-check enabled (tax-compliance critical). Item No / MRP /
    # GST come from Items_March via the EAN.
    # ────────────────────────────────────────────────────────────────────
    'Reliance': {
        'party_name': 'Reliance',            # Must match 'Party' in Ship-To B2B
        'source_format': 'pdf',
        'pdf_parser':    'reliance',
        'po_col':  '__po__',                  # [REQUIRED] synthetic
        'loc_col': '__loc__',                 # [REQUIRED] synthetic (delivery city)
        'qty_col': 'Qty',                     # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'EAN',                     # [REQUIRED in this mode]
        'price_col': None,                    # WMS computes
        'mrp_col': 'MRP',                     # vendor MRP (Validation pair)
        # Reliance's PRE-GST taxable cost. Compared against our
        # GST-dependent computed cost (see gst_margin_discount).
        'fob_col': 'Base Cost',               # [VALIDATION]
        'compare_basis': 'cost',
        'compare_label': 'Cost',
        # GST-dependent margin: keep% = 1 − 0.31 × (1+GST). The GUI
        # default below is just for display (the 18% case).
        'gst_margin_discount': 0.31,
        'default_margin': 63.42,              # display only — gst_margin_discount governs
        'hsn_col': 'HSN Code',                # cross-checked vs master
        # Per-line amount = Base Cost × Qty (= the PDF's 'TOTAL BASIC VALUE',
        # PRE-GST). v2.7: flag it pre-GST so the tracker/DB Order Value is
        # grossed up by each line's GST → the PDF's 'Total Order Value'
        # (inc GST), e.g. 52,533.26 basic → 61,989.29 inc-GST.
        'amount_col': {'multiply': ['Base Cost', 'Qty']},
        'amount_is_pre_gst': True,
        # v2.7: gross up the (pre-GST) amount by the PDF's OWN per-line GST
        # rate ('GST Rate' = IGST%), not the master's gst_code — so the
        # tracker/DB Order Value matches the PDF's 'Total Order Value'.
        'gst_pct_col': 'GST Rate',
        'template_headers': [
            'Sr No', 'Article No', 'EAN', 'HSN Code',
            'Material Description', 'Qty', 'MRP', 'Base Cost',
            'GST Rate', 'Total Base Value', 'Site',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # Zepto (v1.8.0)
    # Zepto's raw PO dump bundles many POs (10-15 typical) into a
    # single data sheet, but that sheet's NAME changes every dump —
    # it's literally 'PO_<random-hex>' like 'PO_64863340b23e6c90' or
    # 'PO_c881cfb0a4fa2ebc'. We use a wildcard ``source_sheet`` to
    # match any such prefix. The engine aborts with a clear error if
    # no matching sheet is found (vs. Reliance's behavior of failing
    # over to the first sheet — that would be wrong here because
    # Zepto workbooks usually have other sheets like 'Sheet1' or
    # 'Skipped' that we must NOT read).
    #
    # Math: same as Blink — ``Unit Base Cost`` is POST-GST, so
    # ``compare_basis='cost'``. ``Total Amount`` is pre-calculated
    # on the punch so we just read it directly.
    #
    # Ignored by design (per user decision):
    #   * 'Sheet1' / 'Sheet2' — user's manual ship-to map (we use
    #     Ship-To B2B.xlsx instead, which is location-keyed and
    #     more stable than per-PO mappings).
    #   * 'Skipped' — user's own bucket for high-diff items they
    #     chose to exclude from the SO. Engine treats it as
    #     invisible.
    #   * HSN column is present but we DON'T cross-check it
    #     (Reliance-only feature).
    # ────────────────────────────────────────────────────────────────────
    'Zepto': {
        'party_name': 'Zepto',           # Must match 'Party' in mapping sheet
        'source_sheet': 'PO_*',          # v1.8.0 — wildcard prefix match
        'po_col': 'PO No.',              # [REQUIRED]
        'loc_col': 'Del Location',       # [REQUIRED] e.g. 'CHN-SS-MH-THIRUVALLUR'
        'qty_col': 'Qty',                # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'EAN',                # [REQUIRED in this mode]
        'price_col': None,               # WMS computes
        'fob_col': 'Unit Base Cost',     # [VALIDATION] post-GST (same as Blink's cost_price)
        'default_margin': 70,
        'compare_basis': 'cost',         # Unit Base Cost ≈ MRP × 70% ÷ 1.18
        'compare_label': 'Cost',
        'amount_col': 'Total Amount',    # Zepto's pre-calc: Unit Base Cost × Qty
        'template_headers': [
            'PO No.', 'PO Date', 'Status', 'Vendor Code', 'Vendor Name',
            'PO Amount', 'Del Location', 'Line No', 'SKU', 'SKU Code',
            'SKU Desc', 'Brand', 'EAN', 'HSN',
            'CGST %', 'SGST %', 'IGST %', 'CESS %',
            'MRP/RSP', 'Qty', 'Unit Base Cost', 'Landing Cost',
            'Total Amount',
            'Created By', 'ASN Quantity', 'GRN Quantity', 'PO Expiry Date',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # BlinkMP (v1.9.0) — distinct from "Blink" (Blinkit-BCPL) despite
    # similar naming.
    #
    # Blink (Blinkit-BCPL) — the quick-commerce dark stores, handles
    #                        PO number in column 'po_number',
    #                        post-GST cost in 'cost_price', 70% margin.
    # BlinkMP (BCPL-RO)   — the reorder/wholesale channel, PO column
    #                        is simply 'PO', has 'Landing Rate'
    #                        (pre-GST) instead of post-GST cost,
    #                        75% margin.
    #
    # Both ship into BCPL's warehouse network (Lucknow L4, Kundli,
    # Mumbai M10, etc.) which is why the location strings look
    # similar, but they're separate commercial relationships with
    # their own Ship-To B2B rows. Do NOT conflate them.
    # ────────────────────────────────────────────────────────────────────
    'BlinkMP': {
        # v1.9.3: ``party_name`` is 'Blink RO' (BCPL's internal name
        # for the Reorder channel) to match the existing mapping rows
        # in Ship-To B2B.xlsx. The user-facing marketplace name stays
        # 'BlinkMP' (that's the dict key), so the GUI dropdown still
        # shows "BlinkMP" — but behind the scenes the mapping lookup
        # uses the canonical ERP name. Lets the GUI label and the
        # operational spreadsheet stay independently correct.
        #
        # v2.1.2: BlinkMP now arrives in TWO valid formats because the
        # AHD (Ahmedabad) and BLR (Bengaluru) operations teams produce
        # the dump independently and use slightly different column
        # names. Both formats MUST be supported simultaneously; we
        # don't know which file we'll receive on any given day.
        #
        # AHD format (sheet 'Sheet1', 17 cols):
        #   PO | Location | Item Code | HSN Code | Product UPC |
        #   Product Description | Grammage | CGST % | SGST % | IGST % |
        #   CESS % | Additional CES | Tax Amount | Landing Rate |
        #   Quantity | MRP | Total Amount
        #
        # BLR format (sheet 'Sheet1', 17 cols):
        #   date | ro_number | location | entity_name | Item Code |
        #   HSN Code | Product UPC | Product Description | Grammage |
        #   CGST % | SGST % | IGST % | Tax Amount | LR | Quantity |
        #   MRP | Total Amount
        #
        # Engine-critical column differences (everything else either
        # matches or is absorbed by case_insensitive_cols):
        #   * AHD 'PO'           ↔  BLR 'ro_number'
        #   * AHD 'Landing Rate' ↔  BLR 'LR'
        #
        # Resolved via list-aliases on po_col + fob_col — the engine
        # already supports list aliases (Myntra uses one for po_col),
        # the resolver picks whichever entry matches an actual file
        # column. AHD's 'Location' (Title) vs BLR's 'location' (lower)
        # is handled by case_insensitive_cols=True.
        #
        # AHD-only columns (CESS %, Additional CES) and BLR-only
        # columns (date, entity_name) appear on the Raw Data sheet
        # for audit but are not consumed by the engine — the SO
        # output is identical regardless of which source format
        # arrived.
        #
        # Both files use 'Sheet1' as the data sheet (BLR's earlier
        # 'DUMP' sheet was a one-off from a single sample file —
        # confirmed not canonical). source_sheet config key is
        # therefore omitted and defaults to 'Sheet1'.
        'party_name': 'Blink RO',            # Must match 'Party' in mapping sheet
        'po_col': ['PO', 'ro_number'],       # [REQUIRED] AHD uses 'PO', BLR uses 'ro_number'
        'loc_col': 'location',               # [REQUIRED] case_insens absorbs AHD's 'Location'
        'qty_col': 'Quantity',               # [REQUIRED] same in both
        'item_resolution': 'from_ean',
        'ean_col': 'Product UPC',            # [REQUIRED in this mode] same in both
        'price_col': None,                   # WMS computes
        'fob_col': ['Landing Rate', 'LR'],   # [VALIDATION] AHD 'Landing Rate', BLR 'LR'; both pre-GST
        'default_margin': 75,                # 75% margin (higher than Blink's 70%)
        'compare_basis': 'landing',          # Landing Rate / LR IS pre-GST
        'compare_label': 'Landing Rate',     # Display label on Validation sheet (uniform across formats)
        'amount_col': 'Total Amount',        # pre-calculated by BlinkMP, same in both formats
        # v2.1.1: tolerate header case drift the same way Myntra and
        # Reliance do. Critical for AHD ('Location') vs BLR
        # ('location') — without this flag, BLR files would crash
        # with "Required column 'location' not found" because the
        # config asks for the lowercase variant. The flag also future-
        # proofs against either operations team adding stray
        # whitespace or capitalisation tweaks.
        'case_insensitive_cols': True,
        # v1.9.1 / v2.1.3: GUI default-state hint for the "Override
        # Unit Price" checkbox. When this flag is True on the selected
        # marketplace, the GUI pre-checks the override box on
        # marketplace change so the user doesn't have to remember to
        # enable it. The user can still uncheck per-run.
        #
        # Pre-v2.1.3 this flag was a runtime trigger — every BlinkMP
        # run automatically overrode Unit Price. v2.1.3 moved the
        # runtime decision to ``ProcessingResult.override_unit_price``
        # (set from the GUI checkbox state), so this flag is now
        # purely advisory: it tells the GUI "this marketplace
        # typically needs override" but doesn't enforce it.
        #
        # Why BlinkMP needs override by default
        # -------------------------------------
        # BCPL (Blinkit + BlinkMP) is registered in Business Central
        # with 70% margin because that's Blinkit's rate — but BlinkMP
        # runs at 75%. Without overriding the Sales Line Unit Price
        # the ERP auto-computes the wrong cost on BlinkMP rows (using
        # the 70% vendor default), creating accounting discrepancies.
        # Stamping our own post-GST Cost Price into col H forces the
        # ERP to record the correct figure.
        #
        # When override is on and a row has no computable Cost Price
        # (e.g. master lookup failed), col H is left empty for that
        # specific row and a log warning is emitted — matches the
        # engine's existing lenient handling for NOT_IN_MASTER rows.
        'override_unit_price': True,
        # v2.1.2: template_headers shows BLR's 17-col format. Chosen
        # over AHD's because BLR's is the more recent / forward-
        # looking shape. Either format is accepted at runtime via
        # the list-aliased po_col/fob_col, so the template only
        # affects the "Download PO Template" button output for users
        # building NEW dump files from scratch — operations teams
        # using their existing dump pipelines aren't affected.
        'template_headers': [
            'date', 'ro_number', 'location', 'entity_name',
            'Item Code', 'HSN Code', 'Product UPC', 'Product Description',
            'Grammage', 'CGST %', 'SGST %', 'IGST %',
            'Tax Amount', 'LR', 'Quantity', 'MRP', 'Total Amount',
        ],
        # v2.1.3: extra template variant for the single-workbook PO
        # template generator. The master template workbook ships with
        # ONE sheet per marketplace; BlinkMP gets TWO sheets (AHD and
        # BLR) because both formats are used in production by
        # different operations teams. ``template_headers`` above
        # populates the 'BlinkMP (BLR)' sheet; this list populates
        # the 'BlinkMP (AHD)' sheet.
        #
        # AHD format differs from BLR by:
        #   * No 'date', 'entity_name' columns
        #   * Has 'CESS %', 'Additional CES' columns instead
        #   * Uses 'PO' (not 'ro_number'), 'Landing Rate' (not 'LR')
        # Engine resolves either format at runtime via list-aliased
        # po_col / fob_col — only the template generator needs both
        # spellings.
        'template_headers_extra': [
            'PO', 'Location', 'Item Code', 'HSN Code', 'Product UPC',
            'Product Description', 'Grammage',
            'CGST %', 'SGST %', 'IGST %', 'CESS %', 'Additional CES',
            'Tax Amount', 'Landing Rate', 'Quantity', 'MRP',
            'Total Amount',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # Flipkart  (v2.1.0 — first SO Flipkart marketplace, distinct from
    #            Flipkart-TO below)
    #
    # Flipkart's punch arrives as a single-sheet workbook covering many
    # POs (51 in our reference sample) shipping to multiple Flipkart-
    # owned hubs. Unlike Flipkart-TO, this marketplace produces standard
    # Sales Orders — there IS a Sell-to Customer (Renee's Flipkart
    # Wholesale entity, 20020 in the ERP) and the Ship-to Code is the
    # specific hub (e.g. 20020_11, 20020_22). Same SO output shape as
    # Blink/RK/Myntra — no special handling needed.
    #
    # Why a separate config from Flipkart-TO
    # --------------------------------------
    # Flipkart-Branch (TO) ships into FK-owned dark warehouses for
    # internal stock movements — no commercial transaction. Flipkart
    # (this entry) ships into FK's wholesale hubs for resale on
    # Flipkart.com — that's a B2B sale to FK Wholesale and gets
    # invoiced. Two different commercial flows = two different ERP
    # document types = two configs. They share the 'Flipkart' brand
    # but nothing else operationally.
    #
    # Punch file shape (cols A-O, 15 cols):
    #   Date | PO | FSN Code | EAN | Qty | COST PRICE With GST |
    #   total_amount | description | Address | Ship TO | Item n |
    #   MRP | GST Rate | 0.77 | Diff
    #
    # The user pre-fills cols J ('Ship TO'), K ('Item n'), N (0.77),
    # O ('Diff') via XLOOKUP/formulas as a sanity check. The engine
    # IGNORES all four — Item No comes from Items_March via EAN lookup
    # (consistent with every other from_ean SO marketplace), Ship-to
    # comes from Ship-To B2B via Address lookup, validation Diff is
    # computed fresh from MRP × margin% vs the file's COST PRICE With
    # GST column.
    #
    # Cost basis: 'landing'  (pre-GST)
    # --------------------------------
    # Flipkart's 'COST PRICE With GST' column is misleadingly named —
    # despite the "With GST" suffix, the column actually stores the
    # PRE-GST landing rate (= MRP × 0.77). Verified across 298 sample
    # rows: file's own Diff column = COST_PRICE − (MRP × 0.77) = 0.00
    # for every row. So compare_basis='landing' (NOT 'cost') and we
    # match against MRP × margin%. Our Cost Price column on the
    # Validation sheet still shows the post-GST cost for reference,
    # same as Myntra (which also uses landing basis).
    #
    # Margin: 77%  (default, user can override per-run via GUI)
    # --------------------------------------------------------
    # Renee's standing agreement with Flipkart Wholesale. Verifiable in
    # the file itself — col N is the literal numeric value 0.77 used in
    # the user's reference Diff formula.
    #
    # Unmapped POs handling
    # ---------------------
    # In the 30-04-2026 sample, 6 of 51 POs ship to addresses NOT yet
    # in Ship-To B2B (Ludhiana 'Plot no. B3 to B8' → 20020_22, and a
    # Padgha CF warehouse → 20020_3). Treated like any other unmapped
    # SO row: empty Cust No / Ship-to cells, UNMAPPED status badge on
    # Summary, deduped warning on Warnings sheet — no special logic.
    # User adds the missing rows to Ship-To B2B and re-runs.
    # ────────────────────────────────────────────────────────────────────
    'Flipkart': {
        'party_name': 'Flipkart',            # Must match 'Party' in mapping sheet
        # v2.7.x: Flipkart's new vendor portal ships ONE
        # 'purchase_order_<PO>.xlsx' per PO (a two-row hierarchical header),
        # retiring the old standalone "dump generator" tool. ``file_parser``
        # routes each file through engine/flipkart_dump_parser.py →
        # compiles the dump columns below in memory; setting it also makes
        # Flipkart MULTI-FILE (``_supports_multi_file`` / ``process_multi``)
        # so the operator drops ALL of the day's PO files → one SO batch.
        # PO number comes from the FILENAME (not a cell). NOTE: with the
        # parser set, the input is the RAW portal PO files — the old
        # FL_DUMP_COMPILATION.xlsx (the standalone generator's output) is
        # retired; that two-step flow is now done in-app.
        'file_parser': 'flipkart',
        'accepted_extensions': ['.xlsx'],    # raw portal PO files
        'po_col': 'PO',                      # [REQUIRED] alphanumeric (e.g. 'FLS05608C31C')
        'loc_col': 'Address',                # [REQUIRED] full postal address
        # v2.7.x: resolve ship-to by ADDRESS (pincode + survey-no/village body
        # overlap) rather than the generic name tiers. The new portal drops
        # the 'Flipkart India Pvt. Ltd., ' prefix the Del Location entries
        # carry, so a plain substring match misses — see
        # MappingLoader.lookup_by_address.
        'loc_match': 'address',
        'qty_col': 'Qty',                    # [REQUIRED]
        'item_resolution': 'from_ean',       # Same as Blink/RK/Myntra/Reliance/Zepto
        'ean_col': 'EAN',                    # [REQUIRED in this mode] int64 in punch
        # v2.7.x: the new portal exposes the vendor MRP ('Supplier MRP'), so
        # we carry it for the Validation 'Vendor MRP' pair (199 × 0.77 =
        # 153.23 = Supplier Unit Price → the file itself proves the 77% rule).
        'mrp_col': 'MRP',
        'price_col': None,                   # WMS computes Unit Price downstream
        # [VALIDATION] file's stated landing rate (despite the historical
        # 'With GST' name, it's the PRE-GST landing = MRP × 77%).
        # v2.3.1: as of the 10-06-2026 dump the column was RENAMED from
        # 'COST PRICE With GST' → 'COST PRICE', and its values now carry a
        # ' INR' suffix ('114.73 INR'). List-alias resolves either header
        # (new name first); the engine's _extract_float strips the ' INR'
        # suffix / thousands separators so the value parses.
        'fob_col': ['COST PRICE', 'COST PRICE With GST'],
        'default_margin': 77,                # 77% landing — Renee × Flipkart Wholesale agreement
        # Compare against MRP × margin% (pre-GST landing). Our own Cost
        # Price column still appears on the Validation sheet for
        # reference (engine populates cost_price_ref unconditionally).
        # Same basis as Myntra/BlinkMP.
        'compare_basis': 'landing',
        'compare_label': 'Landing Rate',
        'amount_col': 'total_amount',        # File's pre-calculated row amount (Cost × Qty)
        # v2.3.1: template mirrors the current dump shape (COST PRICE, and
        # an 'Item' column; the old 'Item n'/'MRP'/'GST Rate' columns are
        # gone — Item No / MRP / GST come from master via EAN anyway).
        'template_headers': [
            'Date', 'PO', 'FSN Code', 'EAN', 'Item', 'Qty',
            'COST PRICE', 'total_amount', 'description',
            'Address', 'Ship TO',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # Nykaa  (Nykaa E-Retail Ltd — v2.3.1)
    #
    # Standard SO marketplace (CSV). Cust No 20012 in Ship-To B2B; each
    # 'Delivery Location' (Nykaa Warehouse BLR2, Chennai Warehouse, ...)
    # maps to a 20012_<n> ship-to code.
    #
    # File columns:
    #   PO Number | ... | SKU Code | SKU Name | ... | Unit Cost | MRP |
    #   PO Qty | ... | Tax Amount | HSN Code | EAN Code | Delivery Location
    #
    # Item No / MRP / GST come from Items_March via 'EAN Code'
    # (item_resolution='from_ean'). 'SKU Code' (RENEE000…) is Nykaa's
    # internal code, not used.
    #
    # PRICE BASIS — 'cost' (post-GST). The file's 'Unit Cost' is the
    # post-GST cost: Unit Cost == MRP × keep% ÷ (1 + GST). Verified across
    # the 10-06-2026 dump (e.g. MRP 450 → 251.68 == 450 × 0.66 ÷ 1.18).
    #
    # ── Per-line margin split (margin_rules) ────────────────────────────
    # Nykaa's discount depends on product category (two Ship-To B2B rows,
    # both Cust 20012):
    #   * Perfume/Fragrance → 31% off MRP → keep 69%
    #   * Cosmetics (default) → 34% off MRP → keep 66%
    # So landing/cost is computed PER LINE.
    #
    # DECISION is by NAME only — the same manual lookup the operator does:
    # a line is Perfume/Fragrance when its 'SKU Name' contains 'perfume'
    # or 'fragra'. (Drop/extend the ``contains`` list to change this.)
    #
    # HSN is a CROSS-CHECK, not a decider (``flag_hsn_conflicts``): when a
    # row's HSN points to a different category than the name chose — e.g.
    # a 'Body Mist' (HSN 3303 = perfume) whose name has no 'perfume'/
    # 'fragra', or an 'Eau De Parfum' carrying cosmetics HSN 3304 — a
    # '⚠ AMBIGUOUS (name vs HSN)' warning is logged for review. This is a
    # deliberate stopgap until the exact perfume lookup is finalised; the
    # NAME still governs the margin so behaviour matches the manual method.
    #
    # NOTE: the current dump applies a flat 66% even to perfumes, so
    # perfume lines show a Cost MISMATCH against our 69% calculation —
    # exactly the discrepancy this surfaces. Each rule application (and
    # each name-vs-HSN conflict) is logged on the Warnings sheet.
    # ────────────────────────────────────────────────────────────────────
    'Nykaa': {
        'party_name': 'Nykaa',               # Must match Ship-To B2B 'Party'
        'po_col': 'PO Number',               # [REQUIRED]
        'loc_col': 'Delivery Location',      # [REQUIRED] → Ship-To B2B
        'qty_col': 'PO Qty',                 # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'EAN Code',               # [REQUIRED in this mode]
        'price_col': None,                   # WMS computes Unit Price
        'fob_col': 'Unit Cost',              # [VALIDATION] vendor post-GST cost
        'compare_basis': 'cost',             # MRP × keep% ÷ (1+GST)
        'compare_label': 'Cost',
        'mrp_col': 'MRP',                    # vendor MRP (Validation pair)
        'default_margin': 66,                # GUI display; margin_rules govern
        # Per-line keep% by product category (see header note).
        'margin_rules': {
            'rules': [
                {
                    'label': 'Perfume/Fragrance',
                    'keep_pct': 69,                       # 31% off MRP
                    # DECIDER: description keywords (your manual search).
                    'contains': ['perfume', 'fragra'],
                    'contains_column': 'SKU Name',
                    # v2.4.6: a 'HAIR PERFUME'/'HAIR FRAGRANCE' is a HAIR
                    # product (e.g. RENEE Caramel Crush Hair Perfume), NOT a
                    # fragrance — exclude it so it gets the regular Cosmetics
                    # rate (66%), not 69%. ``excludes`` wins over ``contains``.
                    'excludes': ['hair'],
                    # CROSS-CHECK only (does not decide the margin): flags
                    # rows whose HSN disagrees with the name decision.
                    'hsn_prefix': ['3303'],
                },
            ],
            'default_keep_pct': 66,           # Cosmetics — 34% off MRP
            'default_label': 'Cosmetics',
            'flag_hsn_conflicts': True,       # ⚠ highlight name-vs-HSN clashes
        },
        # Per-line amount = Unit Cost × Qty (file's 'PO Amount' is the
        # whole-PO total repeated on every line, not per-line).
        'amount_col': {'multiply': ['Unit Cost', 'PO Qty']},
        # v2.4.0: the Tracker's Order Value uses this per-PO total verbatim
        # (GST-inclusive; equals the Nykaa portal's 'PO Amount' to the rupee).
        'po_total_col': 'PO Amount',
        # PO Date / Exp Date for the tracker come from these columns.
        'po_date_col': 'PO Release Date',
        'exp_date_col': 'Expiry Date',
        'hsn_col': 'HSN Code',               # 8-digit, cross-checked vs master
        'template_headers': [
            'PO Number', 'Vendor Code', 'Vendor Name', 'PO Release Date',
            'Expected Delivery Date', 'Expiry Date', 'PO Amount',
            'PO Line No', 'SKU Code', 'SKU Name', 'Vendor SKU Code',
            'Unit Cost', 'MRP', 'PO Qty', 'Received Qty', 'Tax Amount',
            'HSN Code', 'EAN Code', 'Delivery Location',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # PURPLLE  (v2.4.0)
    # The dump is a SAP/ERP export saved as '.XLS' but actually TAB-
    # separated text, so it's routed through a custom parser
    # (``file_parser='purplle'``, engine/purplle_parser.py) that reads the
    # TSV and cleans the zero-padded/quote-suffixed EAN
    # ('000008904473100590'' → '8904473100590'). Standard SO pipeline
    # otherwise. Pricing: STRAIGHT 70% — verified to the paisa, the file's
    # 'Price' = MRP × 0.70 ÷ (1+GST) (post-GST cost, like Blink's). Single
    # destination: the Thane address → Cust 20021 / Ship-to 20021_16.
    # ────────────────────────────────────────────────────────────────────
    'Purplle': {
        'party_name': 'Purplle',             # Must match Ship-To B2B 'Party'
        'file_parser': 'purplle',            # tab-separated '.XLS' reader
        'accepted_extensions': ['.xls', '.XLS', '.csv', '.txt'],
        'po_col': 'PO Document Number',      # [REQUIRED]
        'loc_col': 'Address',                # [REQUIRED] → Ship-To B2B (Thane)
        'qty_col': 'Qty',                    # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'EAN Number',             # cleaned by the parser
        'price_col': None,                   # WMS computes Unit Price
        'fob_col': 'Price',                  # [VALIDATION] post-GST cost
        'compare_basis': 'cost',             # Price ≈ MRP × 70% ÷ (1+GST)
        'compare_label': 'Cost',
        'mrp_col': 'MRP',                    # vendor MRP (Validation pair)
        'default_margin': 70,                # STRAIGHT 70% (30% margin)
        # Order Value (Summary/Tracker/DB): Price × Qty grossed by GST →
        # MRP × 70% × Qty (landing), i.e. GST-inclusive per the thumb rule.
        'amount_col': {'multiply': ['Price', 'Qty']},
        'amount_is_pre_gst': True,
        # Tracker dates.
        'po_date_col': 'PO Date',
        'exp_date_col': 'Expiry Date',
        'template_headers': [
            'PO Document Number', 'Item No', 'EAN Number', 'Sku',
            'Material long text', 'MRP', 'Price', 'Qty', 'Plant', 'Address',
            'Storage location', 'Purchasing Group', 'PO Date', 'Expiry Date',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # SWIGGY  (v2.4.0)
    # Flat CSV ('PO_<id>.csv'). The dump carries only a Swiggy 'SkuCode'
    # (no EAN/Item No), so item_resolution='from_swiggy_sku' recovers the EAN
    # from the master's 'Swiggy' sheet (SkuCode→EAN), then resolves it like
    # any from_ean marketplace. Pricing: STRAIGHT 80% — verified, the dump's
    # 'UnitBasedCost' = MRP × 0.80 ÷ (1+GST). Deal SKUs (master's 'Swiggy
    # Deal SKUs' sheet) carry an explicit negotiated cost that overrides the
    # 80% calc in validation. Order value = 'PoLineValueWithTax' (the dump's
    # own GST-inclusive line total, summed — like Blinkit).
    # ────────────────────────────────────────────────────────────────────
    'Swiggy': {
        'party_name': 'Swiggy',              # Must match Ship-To B2B 'Party'
        'accepted_extensions': ['.csv'],
        'po_col': 'PoNumber',                # [REQUIRED]
        'loc_col': 'FacilityName',           # [REQUIRED] → Ship-To B2B
        'qty_col': 'OrderedQty',             # [REQUIRED]
        # v2.4.4: Swiggy PO-status REVIEW. The dump's 'Status' column carries
        # CONFIRMED plus EXPIRED / COMPLETED / CANCELLED (and PENDING on some
        # exports). Non-CONFIRMED lines are NOT dropped — they're pasted as-is
        # and FLAGGED (one named warning per PO) so the operator can manually
        # audit and remove them (a status can be wrongly given). ``status_keep``
        # = the states that need no review.
        'status_col': 'Status',
        'status_keep': ['CONFIRMED'],
        'item_resolution': 'from_swiggy_sku',
        'sku_col': 'SkuCode',                # → master 'Swiggy' sheet → EAN
        'price_col': None,                   # WMS computes Unit Price
        'fob_col': 'UnitBasedCost',          # [VALIDATION] post-GST cost
        'compare_basis': 'cost',             # UnitBasedCost ≈ MRP × 80% ÷ (1+GST)
        'compare_label': 'Cost',
        'mrp_col': 'Mrp',                    # vendor MRP (Validation pair)
        'default_margin': 80,                # STRAIGHT 80% (20% margin)
        'amount_col': 'PoLineValueWithTax',  # GST-inclusive line total (sum)
        'po_date_col': 'PoCreatedAt',
        'exp_date_col': 'PoExpiryDate',
        'template_headers': [
            'PoNumber', 'FacilityId', 'FacilityName', 'City', 'PoCreatedAt',
            'SkuCode', 'SkuDescription', 'OrderedQty', 'Tax',
            'PoLineValueWithoutTax', 'PoLineValueWithTax', 'Mrp',
            'UnitBasedCost', 'ExpectedDeliveryDate', 'PoExpiryDate',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # Flipkart-TO  (v2.0.0 — first Transfer Order marketplace)
    #
    # Distinct from the 'Flipkart' SO entry above. Flipkart Branch
    # dumps go to FK-owned warehouses (Bhiwandi BTS, Sanpka 01, etc.)
    # via Transfer Orders, not Sales Orders. The downstream ERP record
    # has no Sell-to Customer — only a Transfer-to Code (FK_BHW_BTS,
    # FK_SANPKA1, ...).
    #
    # The 'output_type': 'to' flag tells the engine to take the TO
    # code path, which:
    #
    #   * Aggregates rows by (PO, Item No) so duplicate FSNs for the
    #     same Item No collapse to a single Transfer Line with summed
    #     qty (cleaner D365 import; same total qty per Item No).
    #   * Emits a Transfer Order package (Headers (TO) + Lines (TO))
    #     instead of an SO package.
    #
    # The file's optional 'Ship to' col is IGNORED — that column is
    # added by the user as a sanity-check, not authoritative.
    # Transfer-to Code is always resolved fresh from Ship-To B2B
    # (Party=Flipkart-TO, Del Location=row's Location → Ship to col).
    #
    # When a Ship-To B2B row exists for the location but its Ship to
    # column is blank (currently 'Howrah'), the engine emits a
    # WARNING and writes an empty Transfer-to cell so the user spots
    # it on D365 import preview.
    #
    # ── v2.1.4: structural change to dump format ────────────────────────
    # Pre-v2.1.4 the Flipkart Branch dump was self-contained (carried
    # Item No, MRP, GST Code per row). As of 2026-05-06 the dump only
    # ships 7 columns:
    #
    #   1. Po Number   (sometimes with trailing space — case_insensitive_cols
    #                   absorbs it)
    #   2. Location
    #   3. FSN         (Flipkart's internal SKU code, display only)
    #   4. SKU Id      (= EAN; used for Items_March master lookup)
    #   5. Product Name
    #   6. Cost Price  (Flipkart's stated cost — used for REFERENCE
    #                   comparison only; engine still calculates its
    #                   own Transfer Price from master)
    #   7. Quantity Sent
    #
    # That structural change drove three config-level changes:
    #
    #   * ``item_resolution`` switched from ``'from_column'`` to
    #     ``'from_ean'``. Item No / MRP / GST Code now come from
    #     Items_March via SKU Id (= EAN) — same path Blink uses.
    #
    #   * ``mrp_col``, ``gst_col``, ``item_col`` removed. Those
    #     columns no longer exist in the file.
    #
    #   * ``compare_mode='reference_only'`` added. Flipkart's stated
    #     Cost Price (col 6) is logged + shown on the Validation
    #     sheet, but never blocks — Transfer Price always uses the
    #     engine's calculated value (MRP × 60% ÷ GST). Differences
    #     are written to the Warnings sheet for audit but the row's
    #     ``validation_status`` stays 'OK' instead of 'MISMATCH'.
    #
    # ────────────────────────────────────────────────────────────────────
    'Flipkart-TO': {
        'party_name': 'Flipkart-TO',         # Must match Ship-To B2B 'Party' col exactly
        'output_type': 'to',                  # v2.0.0 — flips engine + exporter to TO pipeline

        # Required columns (v2.1.4 standard 7-col format)
        'po_col': 'Po Number',                # → D365 Transfer Header 'No.' column
        'loc_col': 'Location',                # → looked up in Ship-To B2B
        'qty_col': 'Quantity Sent',           # → Transfer Line Quantity (col D)

        # v2.1.4: Item resolution moved to master lookup via EAN.
        # SKU Id column contains EAN-format integers (e.g. 8906121646603).
        # Engine resolves item_no, mrp, gst_code from Items_March via
        # this EAN. Rows whose EAN isn't in master become
        # NOT_IN_MASTER (warned, but row continues with item_no=
        # f'?EAN:{ean}' placeholder — same behaviour as Blink).
        'item_resolution': 'from_ean',
        'ean_col': 'SKU Id',

        # v2.1.4: Cost Price column on the file is Flipkart's stated
        # cost. We compare against engine's calculated value and log
        # diffs as warnings — see compare_mode='reference_only' below.
        'fob_col': 'Cost Price',
        'compare_basis': 'cost',
        'compare_label': 'Cost (Flipkart)',

        # v2.1.4: NEW flag. When set, the engine still computes diffn
        # but downgrades MISMATCH → OK on the validation sheet, and
        # writes one warning per mismatched row. Transfer Price
        # written to D365 always uses the engine's calculated value
        # (MRP × margin% ÷ (1+GST)) regardless of diff. Used for
        # marketplaces where the source's price is informational only
        # but worth tracking for audit.
        'compare_mode': 'reference_only',

        'default_margin': 60,                 # 60% — Flipkart Branch agreement

        # No revenue tracking (TOs aren't revenue):
        # 'amount_col' omitted

        # v2.1.4: tolerate header drift the same way Myntra / Reliance
        # / BlinkMP do. Critical for the trailing-space typo seen in
        # the 06-05-2026 dump ('Po Number ' instead of 'Po Number');
        # this flag normalises header names before matching.
        'case_insensitive_cols': True,

        'source_sheet': 'Sheet1',
        'header_row': 0,

        # v2.1.4: template_headers refreshed to match the new 7-col
        # canonical format. Used by the 'Download All Templates'
        # button — the user gets a blank template that mirrors what
        # they'll actually receive. The 'Cost Price' column is kept
        # because it's now the reference-comparison source (was
        # purely informational pre-v2.1.4).
        'template_headers': [
            'Po Number', 'Location', 'FSN', 'SKU Id', 'Product Name',
            'Cost Price', 'Quantity Sent',
        ],

        # ── v2.3.1: bulk consignment input mode ─────────────────────────
        # Flipkart-TO now accepts its punch in TWO shapes; the operator
        # chooses which in the GUI ("Flipkart-TO Input Mode" radios):
        #
        #   1. CONSOLIDATED (historical, default) — the operator manually
        #      consolidates Flipkart's per-PO consignment exports into a
        #      single 7-column dump (see ``template_headers`` above) and
        #      feeds that ONE file through the normal ``engine.process``
        #      path. The 'Po Number' and 'Location' columns are real
        #      columns the operator filled in by hand. Nothing about this
        #      mode changed.
        #
        #   2. CONSIGNMENTS (new) — the operator instead hands us the raw
        #      per-PO exports directly, named
        #      ``Consignment_Details_<PO>_<date>.csv`` (one file per PO),
        #      and the engine builds the consolidated dump itself via
        #      ``engine.process_consignments``. Two values are synthesized
        #      because the raw CSV doesn't carry them:
        #        * 'Po Number'  ← parsed from each FILENAME
        #          (``filename_po_regex`` capture group 1).
        #        * 'Location'   ← left EMPTY on purpose. Ship-to/bill-to
        #          resolution for these consignments is still being
        #          worked out; until then every row maps to a blank
        #          Transfer-to Code. When the rule is known, populate it
        #          in ``engine.process_consignments`` — no other change
        #          needed.
        #      Every OTHER column the dump needs (``SKU Id`` = EAN,
        #      ``Quantity Sent``, ``Cost Price``, ``Product Name``,
        #      ``FSN``) is already present in the raw consignment CSV
        #      under these exact names, so the assembled dump is
        #      shape-identical to mode 1 and flows through the same TO
        #      pipeline (aggregate by (PO, Item No), master lookup by EAN,
        #      engine-computed Transfer Price).
        #
        # The raw consignment CSV carries ~20 columns (Product Name, FSN,
        # SKU Id, Brand, ..., Quantity Sent, ..., Cost Price, dimensions);
        # the extra columns are simply ignored by the pipeline.
        'consignment_mode': {
            'enabled': True,
            # v2.3.1: Flipkart-TO historically also accepts a single
            # pre-consolidated dump, so the GUI offers BOTH the
            # consolidated and the bulk-consignment input modes (two
            # radios). Marketplaces without a consolidated format (e.g.
            # Meesho-TO) set this False and are bulk-consignment-only.
            'consolidated_option': True,
            # Capture-group-1 regex applied to each filename's basename.
            # 'Consignment_Details_204535220_09-06-2026.csv' → '204535220'.
            'filename_po_regex': r'Consignment_Details_([^_]+)_',
            # Multi-select CSV picker in the GUI for this mode.
            'accepted_extensions': ['.csv'],
            # v2.4.1: Auto mode auto-detects the Consignment Visibility
            # Report dropped in the folder by this filename pattern
            # (e.g. 'Consignment_Visibility_Report_f968…_09-06-2026.csv')
            # and threads it into process_consignments for Location
            # resolution — so Auto matches the manual GUI (which took the
            # report as a separate pick). The per-PO 'Consignment_Details_'
            # files don't match this, so the two never collide.
            'visibility_filename_regex': r'Consignment_Visibility_Report',

            # ── Location source: Consignment Visibility Report ──────────
            # The raw consignment CSVs carry no Location, so the operator
            # optionally supplies Flipkart's Consignment Visibility Report
            # (a per-run pick in the GUI). The engine joins it to each
            # consignment on PO and writes the report's Warehouse Id as
            # the row's Location.
            #
            # Column names in the visibility report:
            'visibility_po_col': 'Consignment Id',   # = PO number
            'visibility_loc_col': 'Warehouse Id',    # = destination code

            # ── PROVISIONAL Warehouse-Id → Ship-To B2B bridge ──────────
            # The report's 'Warehouse Id' is a machine code (e.g.
            # 'malur_bts') that does NOT match the friendly Ship-To B2B
            # location names ('Malur BTS Warehouse'), so it can't resolve
            # to a Transfer-to Code on its own.
            #
            # TEMPORARY measure (until exact ship-to codes are wired into
            # the report or Ship-To B2B): this internal alias map
            # translates each machine code to its friendly Ship-To B2B
            # location NAME. In _process_to the engine looks the mapping
            # up under the friendly name (so the real Transfer-to Code
            # resolves from the spreadsheet — the code stays single-sourced
            # there) but KEEPS the raw machine code as the row's Location.
            # The Summary therefore shows both 'Location (Raw)' (the
            # machine code that came in) and 'Location (Mapped)' (the
            # deciphered friendly name), and auto-highlights the pair as a
            # fuzzy match because they differ.
            #
            # Every PO resolved through this map is also FLAGGED as a
            # PROVISIONAL / fuzzy match on the Warnings sheet so the
            # operator verifies it before import. Keys are the lowercase
            # machine codes exactly as they appear in 'Warehouse Id';
            # values are the EXACT Ship-To B2B location keys.
            #
            # Pairings confirmed by the operator (2026-06-09). DELETE this
            # whole map once the report/Ship-To B2B carries exact codes.
            'warehouse_aliases': {
                'hyderabad_medchal_01': 'Hyderabad Medchal 01',  # FK_HYD_MED
                'chn_tvr_04':           'Chennai 04 Warehouse',   # FK_CHN_04
                'ulub_bts':             'kolkata uluberia bts',   # FK_ULB_BTS
                'ahm_khe_wh_nl_01nl':   'Ahmedabad_01',           # FK_AMD_01
                'jai_san_wh_nl_01nl':   'Jaipur Sanganer NLFC 01',# FK_NLFC01
                'malur_bts':            'Malur BTS Warehouse',    # FK_MAL_BTS
                'gur_san_wh_nl_01nl':   'sanpka 01',              # FK_SANPKA1
                'bhi_vas_wh_nl_01nl':   'Bhiwandi BTS',           # FK_BHW_BTS
                'nag_wad_wh_nl_02nl':   'Nagpur 3pl 01',          # FK_NAGPUR
                'dha_kot_wh_nl_01nl':   'Hubli_01',               # FK_DHARWAD
            },
        },
    },

    # ────────────────────────────────────────────────────────────────────
    # Meesho-TO  (v2.3.1 — second Transfer Order marketplace)
    #
    # Same TO pipeline as Flipkart-TO (output_type='to': aggregate by
    # (PO, Item No), Item No / MRP / GST from Items_March via EAN, engine
    # computes Transfer Price), but a DIFFERENT source file shape.
    #
    # Source files: Meesho exports ONE CSV per order, named
    # ``order-line-items-<PO>.csv`` (e.g. order-line-items-64415.csv).
    # There is no consolidated dump, so Meesho-TO is BULK-CONSIGNMENT-ONLY
    # — it always goes through ``engine.process_consignments`` (the GUI
    # shows no consolidated/consignment toggle for it).
    #
    # File columns (11):
    #   skuCode            — Meesho's internal listing SKU (e.g.
    #                        '57nb-498q8d-4n'); display only, NOT the EAN.
    #   description        — product title.
    #   styleCode          — THE EAN (e.g. 8906121644654) used for the
    #                        Items_March master lookup. A few combo SKUs
    #                        carry a style string instead (e.g.
    #                        'RENEEFACEBRIGHTCOMBO'); those resolve too
    #                        when present in master, else NOT_IN_MASTER.
    #   size / color       — ignored.
    #   sellingPricePerUnit— Meesho's SELLING price. Deliberately NOT used
    #                        as a cost reference: comparing a selling price
    #                        to the engine's computed Transfer Price (a
    #                        cost) would emit misleading diffs. So no
    #                        ``fob_col`` — Transfer Price comes purely from
    #                        master (MRP × margin ÷ (1+GST)).
    #   orderedQty         — quantity → Transfer Line Qty.
    #   asnCreatedQty / packedQty / dispatchedQty / pendingQty
    #                      — fulfillment counters; ignored.
    #
    # PO number: parsed from the FILENAME (no PO column), same mechanism
    # as Flipkart-TO consignment mode.
    #
    # Location: LEFT EMPTY for now. Meesho's Transfer-to codes are 'MS_'-
    # prefixed in Ship-To B2B, but how each order's destination warehouse
    # is fetched is still TBD. Until then Location is blank → empty
    # Transfer-to Code + one "unmapped" warning per PO. When the source is
    # known, add it to ``process_consignments`` (and, if it arrives as a
    # machine code, a ``warehouse_aliases`` map here) exactly like
    # Flipkart-TO.
    #
    # Margin: 60% is a PLACEHOLDER copied from Flipkart-TO — CONFIRM the
    # real Meesho TO margin. It's overridable per run in the GUI, but this
    # default should be corrected once known.
    # ────────────────────────────────────────────────────────────────────
    'Meesho-TO': {
        'party_name': 'Meesho-TO',           # Must match Ship-To B2B 'Party'
        'output_type': 'to',                  # TO pipeline (Headers/Lines (TO))

        # PO is synthesized from the filename; Location injected blank.
        'po_col': 'Po Number',
        'loc_col': 'Location',
        'qty_col': 'orderedQty',

        # Item No / MRP / GST resolved from master via the EAN in styleCode.
        'item_resolution': 'from_ean',
        'ean_col': 'styleCode',

        # No fob_col — see note above (sellingPricePerUnit is a selling
        # price, not a cost). Transfer Price is engine-computed from master.

        'default_margin': 60,                 # PLACEHOLDER — confirm Meesho margin

        # Tolerate header case/whitespace drift like the other CSV configs.
        'case_insensitive_cols': True,
        'header_row': 0,

        'consignment_mode': {
            'enabled': True,
            # Bulk-consignment-ONLY: no consolidated single-file format,
            # so the GUI skips the mode toggle and always uses the
            # consignment flow for Meesho-TO.
            'consolidated_option': False,
            # 'order-line-items-64415.csv' → '64415'.
            'filename_po_regex': r'order-line-items-(\d+)',
            'accepted_extensions': ['.csv'],

            # v2.4.0: Location comes from the FILENAME. The operator renames
            # each file to include the destination city token (the suffix of
            # the Meesho Ship-To B2B code: MS_BLR → 'blr', MS_GGN → 'ggn',
            # MS_KOL → 'kol'), e.g. 'order-line-items-64415-blr.csv'. The
            # engine scans the filename for any known token and resolves it to
            # that city's Transfer-to Code via Ship-To B2B — so adding a new
            # city only needs a new Ship-To B2B row, no code change.
            'filename_loc_from_shipto': True,
            # v2.4.0: Meesho files carry no dates → PO Date = today,
            # Exp Date = today + 7 days (for the tracker).
            'po_date_today': True,
            'exp_date_offset_days': 7,
        },

        # Used by 'Download PO Template' — mirrors Meesho's export columns.
        'template_headers': [
            'skuCode', 'description', 'styleCode', 'size', 'color',
            'sellingPricePerUnit', 'orderedQty', 'asnCreatedQty',
            'packedQty', 'dispatchedQty', 'pendingQty',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # BIG BASKET (Innovative Retail Concepts Pvt Ltd / Bigbasket)
    # v2.7: each PO arrives as its OWN '<PO>.xlsx' with a multi-row HEADER
    # block (warehouse / supplier / GSTINs / PO Number / PO Date / Expiry)
    # ABOVE the line-item table — so a flat read can't find the columns.
    # ``file_parser='bigbasket'`` routes it through a custom parser that
    # pulls the PO/dates/warehouse from the preamble and the lines from the
    # table, emitting the engine's synthetic __po__/__loc__/__po_date__/
    # __exp_date__ columns. One file per PO → pick many (multi-file).
    # Pricing: cost = MRP × 70% ÷ (1+GST) (straight 70% multiplier / 30%
    # margin), compared to the file's per-unit pre-GST 'Basic Cost'.
    # ────────────────────────────────────────────────────────────────────
    'Bigbasket': {
        'party_name': 'Bigbasket',            # matches Ship-To B2B party (Cust 20007)
        'file_parser': 'bigbasket',           # custom-layout Excel parser
        'po_col': '__po__',                   # from the parser (PO Number)
        'loc_col': '__loc__',                 # warehouse code = exact Del Location
        'qty_col': 'Quantity',                # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'EAN/UPC Code',            # [REQUIRED in this mode]
        'price_col': None,                    # WMS computes
        'fob_col': 'Basic Cost',              # per-unit PRE-GST cost [VALIDATION]
        'compare_basis': 'cost',              # CP = MRP × m% ÷ (1+GST)
        'compare_label': 'Cost',
        'default_margin': 70,                 # MRP × 70% (30% margin) — straight
        'hsn_col': 'HSN Code',                # cross-checked vs master
        # Per-line amount = Basic Cost × Qty (pre-GST 'TOTAL BASIC VALUE');
        # grossed up by the PO's own GST% → inc-GST order value for the
        # tracker/DB (same mechanism as Reliance).
        'amount_col': {'multiply': ['Basic Cost', 'Quantity']},
        'amount_is_pre_gst': True,
        'gst_pct_col': 'GST%',
        'template_headers': [
            'S.No', 'HSN Code', 'SKU Code', 'Description', 'EAN/UPC Code',
            'Case Quantity', 'Quantity', 'Basic Cost', 'SGST%', 'SGST',
            'CGST%', 'CGST', 'IGST%', 'IGST', 'GST%', 'GST Amount',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # Dmart (Avenue E-Commerce Ltd, brand 'DMart Ready', CIN
    # U74120MH2014PLC259234). The FIRST PDF-source marketplace in this
    # registry (v2.2.0).
    #
    # NAMING — this marketplace is referenced by THREE different names
    # depending on context, and operators need to know which is which:
    #
    #   * 'Avenue E-Commerce Ltd' — the legal entity printed on the PO
    #     header. The vendor relationship and GST registrations all
    #     reference this name.
    #   * 'DMart Ready' / 'DMart' — the consumer-facing brand. Avenue
    #     operates the DMart Ready quick-commerce service.
    #   * 'Dmart' — what the operator's Ship-To B2B sheet uses as the
    #     'Party' column value. This is what the engine matches against
    #     and what the GUI dropdown displays.
    #
    # So the dict key here MUST be 'Dmart' (matches Ship-To B2B), but the
    # PDF parser inspects 'Avenue E-Commerce' in the header text since
    # that's what's actually printed on the PO.
    #
    # PDF format — NOT Excel. Avenue emails the PO as a 2-page PDF
    # attachment with this layout:
    #   * Page 1: Header block (PO number, GST, dates, vendor, ship-to)
    #     then a table header then ~12 line items.
    #   * Page 2: ~4 more line items, then 'Total <qty> <value>' footer
    #     row, then Amount-in-Words, then Terms.
    #
    # Each line item occupies TWO consecutive text rows after PDF text
    # extraction (description spans 2 visible rows because the table has
    # a 2-row header — '%' values on row 1, '<tax> V' values on row 2 —
    # and the description text wraps). The PDF parser
    # (engine/avenue_pdf_parser.py) handles this — see its module docstring.
    #
    # The parser exposes the data via synthetic columns ``__po__`` /
    # ``__loc__`` (same pattern Reliance's pre_process hook uses) so
    # this config plugs into the existing engine logic without
    # awareness that the source was a PDF.
    #
    # Multiplier verification (Renee_AE00.pdf, 16 items):
    #   For every row, Landed Price == MRP × 0.45 to the paisa,
    #   regardless of whether the GST rate is 5% or 18%. So the
    #   engine's existing ``compare_basis='landing'`` formula
    #   (``MRP × margin%``) matches Avenue's Landed Price directly
    #   when ``default_margin=45``.
    #
    # Ship-to: Dmart has 33 FC locations across India in the operator's
    # Ship-To B2B sheet (Andheri E FC, Belapur FC, ... IFC- Kukse
    # Bhiwandi, ... Yelahanka FC). The Renee_AE00.pdf sample ships to
    # 'IFC- Kukse Bhiwandi' (Cust No 20001, Ship-to 20001_34). The PDF
    # parser canonicalizes the extracted text to that EXACT string
    # (including the space after 'IFC-' AND between 'Kukse' and
    # 'Bhiwandi') so the mapping lookup matches without fuzzy logic.
    #
    # HSN cross-check is enabled (like Reliance) — the PDF carries an
    # 8-digit HSN per item (split across two halves in the PDF, joined
    # by the parser) and Avenue's PO contractually requires HSN to
    # match Items_March master, or the ERP posts wrong GST.
    # ────────────────────────────────────────────────────────────────────
    'Dmart': {
        'party_name': 'Dmart',               # Must match 'Party' in Ship-To B2B
        # v2.2.0: PDF source — engine calls PDF_PARSERS['avenue']
        # (parser key matches the PDF format's origin: Avenue
        # E-Commerce Ltd, the parent company that issues the PO).
        # See module docstring at top of marketplaces.py for the full
        # PDF-parser registration mechanism.
        'source_format': 'pdf',
        'pdf_parser':    'avenue',
        # Synthetic columns set by the Avenue PDF parser. Same pattern
        # Reliance uses for its merged-cell title. ``__po__`` and
        # ``__loc__`` are populated identically on every line item row.
        'po_col':  '__po__',                  # [REQUIRED] synthetic
        'loc_col': '__loc__',                 # [REQUIRED] synthetic
        'qty_col': 'PO Qty',                  # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'EAN',                     # [REQUIRED in this mode]
        'price_col': None,                    # WMS computes
        # Dmart's 'Landed Price' column == MRP × 45% (post-GST,
        # verified across rows at both 5% and 18% GST rates). Matches
        # the engine's ``'landing'`` formula exactly when margin=45.
        'fob_col': 'Landed Price',            # [VALIDATION] vendor Landing (post-GST)
        # v2.4.5: the PDF also carries vendor MRP + vendor CP — surface both.
        #   * ``mrp_col='MRP'`` → Validation 'Vendor MRP' (vs Our MRP).
        #   * ``ref_fob_col='Basic Price'`` → Validation 'Vendor CP'. Basic
        #     Price = Landed ÷ (1+GST) = the PRE-GST cost, which equals our CP
        #     (MRP×45%÷(1+GST)) for a correctly-priced line.
        'mrp_col': 'MRP',
        'ref_fob_col': 'Basic Price',
        'default_margin': 45,                 # Dmart's operating margin
        'compare_basis': 'landing',           # Landing shown (MRP × margin%)
        'compare_label': 'Landed Price',
        # v2.4.5: FINALIZE status on the CP pair, not landing. Vendor Landing /
        # Our LR are still shown and amber-highlighted on any diff, but the
        # row's OK/MISMATCH is decided by |Vendor CP − Our CP| — the operator's
        # reconciliation basis for Dmart.
        'status_basis': 'cost',
        # Total Value is pre-calculated by Avenue (Landed × Qty, with
        # paise-level rounding). Matched within ₹0.5 against the PDF's
        # footer 'Total <qty> <value>' line on the Renee_AE00.pdf
        # sample (143,846.98 to the paisa).
        'amount_col': 'Total Value',
        # Dmart prints an 8-digit HSN per line (4-digit prefix on
        # row 1 + 4-digit suffix on row 2, joined by parser). Worth
        # cross-checking against the master because Dmart's POs
        # contractually require correct HSN — wrong HSN → ERP posts
        # wrong GST → tax authority dispute risk. Same enforcement
        # rationale as Reliance.
        'hsn_col': 'HSN Code',
        'template_headers': [
            'Sr No', 'EAN', 'Article No', 'HSN Code', 'Description', 'UOM',
            'PO Qty', 'MRP', 'Basic Price', 'Landed Price', 'Total Value',
            'CGST %', 'SGST %', 'CESS %', 'IGST %', 'UGST %',
            'CGST V', 'SGST V', 'CESS V', 'IGST V', 'UGST V',
            'GST Rate',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # FirstCry  (Siddeshwari Trading / firstcry.com — v2.3.1)
    #
    # The SECOND PDF-source marketplace (after Dmart) but a STANDARD Sales
    # Order, not a Transfer Order — same SO output shape as Blink / Flipkart
    # / Reliance. Parsed by engine/firstcry_pdf_parser.py
    # (``pdf_parser='firstcry'``), which reads the bordered line-item table
    # via pdfplumber.extract_tables() and exposes synthetic ``__po__`` /
    # ``__loc__`` columns (same convention as the Avenue/Dmart parser), so
    # the engine's SO logic runs format-agnostic.
    #
    # PDF shape (reference PO 'pin212260609ity61c4'):
    #   Header: PONO, PO Date, Vendor Name (Renee = us), Delivered To
    #           (the FirstCry hub — ship-to), buyer GST, Currency.
    #   Table : Sr No | FCID | HSN Code | Manufacturer | Style Code |
    #           Image | Product Name | Description | Brand | Color |
    #           MRP | Base Cost | Tax | Landing Rate |
    #           Billed Qty | Free Qty | Total Qty | Total Amount
    #
    # EAN source: the 'Manufacturer' column (NOT 'Style Code'). Verified
    # against Items_March — Manufacturer holds the GTIN on every row and
    # always resolves; Style Code is unreliable (FCID / blank / a
    # different EAN on several rows). Item No / MRP / GST come from master
    # via this EAN — same ``from_ean`` path as Blink/Flipkart/Dmart.
    #
    # Cost basis: 'landing'. Landing Rate == MRP × 0.60 to the paisa
    # across the sample, so the engine's MRP × margin% formula matches the
    # file's Landing Rate exactly at default_margin=60.
    #
    # Ship-to: 'Delivered To' names the FirstCry hub
    # ('Siddeshwari Trading-Hoskote'). The operator's Ship-To B2B must
    # carry a FirstCry party row whose 'Del Location' matches it (exact or
    # via the mapping's fuzzy/substring tiers). CONFIRM the exact 'Party'
    # spelling and 'Del Location' key used in Ship-To B2B.
    #
    # HSN cross-check enabled (8-digit per line) — same rationale as
    # Dmart/Reliance.
    # ────────────────────────────────────────────────────────────────────
    'Firstcry': {
        'party_name': 'Firstcry',            # Must match Ship-To B2B 'Party' — CONFIRM
        'source_format': 'pdf',
        'pdf_parser':    'firstcry',
        # Synthetic columns injected by the FirstCry PDF parser.
        'po_col':  '__po__',                  # [REQUIRED] synthetic
        'loc_col': '__loc__',                 # [REQUIRED] synthetic (Delivered To NAME)
        # v2.4.6: ship-to ADDRESS, tried as an EXACT match BEFORE the name/
        # fuzzy tiers — so one buyer name with several ship-tos (OM ENTERPRISES
        # → 20493_1 vs the Pune Survey-27/1B address → 20493_2) resolves right.
        # Add the exact PDF address as a Del Location to target a specific
        # ship-to; otherwise resolution falls back to the Delivered-To name.
        'loc_addr_col': '__loc_address__',
        'qty_col': 'Total Qty',               # [REQUIRED] Billed + Free
        'item_resolution': 'from_ean',
        'ean_col': 'Manufacturer',            # [REQUIRED] the GTIN column
        'price_col': None,                    # WMS computes Unit Price
        # Landing Rate == MRP × 60% (post-GST landed cost). Matches the
        # engine's 'landing' formula exactly at margin=60.
        'fob_col': 'Landing Rate',            # [VALIDATION] vendor landing
        'default_margin': 60,                 # Renee × FirstCry agreement
        'compare_basis': 'landing',           # MRP × margin% (no GST division)
        'compare_label': 'Landing Rate',
        # v2.3.1: FirstCry's PDF also carries the vendor's MRP and a
        # pre-GST 'Base Cost'. Wire them so the Validation sheet can show
        # Vendor MRP vs Our MRP and Vendor CP vs Our CP side by side (and
        # compute a Cost-Price diff, not just the Landing-Rate diff).
        #   mrp_col     → vendor MRP (Vendor MRP column)
        #   ref_fob_col → vendor cost price; ref_diffn = Base Cost − our CP
        'mrp_col': 'MRP',                     # vendor MRP
        'ref_fob_col': 'Base Cost',           # vendor cost price (pre-GST)
        'amount_col': 'Total Amount',         # file's pre-calc row amount
        'hsn_col': 'HSN Code',                # 8-digit, cross-checked vs master
        'template_headers': [
            'Sr No', 'FCID', 'HSN Code', 'Manufacturer', 'Style Code',
            'Product Name', 'Brand Name', 'Color', 'MRP', 'Base Cost',
            'Tax', 'Landing Rate', 'Billed Qty', 'Free Qty', 'Total Qty',
            'Total Amount',
        ],
    },
}


# Marketplace names for the GUI dropdown (insertion order = display order).
MARKETPLACE_NAMES: List[str] = list(MARKETPLACE_CONFIGS.keys())


# ════════════════════════════════════════════════════════════════════════
# Warehouse codes (v1.9.0)
# ════════════════════════════════════════════════════════════════════════
#
# Purpose
# -------
# The D365 Sales Header (col K) and Sales Line (col F) require a valid
# Business Central location code on every row. Until v1.9.0 we hardcoded
# ``PICK`` (Ahmedabad's warehouse code) because it was the only warehouse
# in use. As of v1.9.0 the user can pick between warehouses at run-time
# via a GUI dropdown next to the Margin input.
#
# The dropdown shows the short friendly code (AHD, BLR); the engine
# substitutes the full ERP code (``PICK``, ``DS_BL_OFF1``) into the D365
# file. Users don't have to remember which cryptic string maps to what
# city — they pick by city name and the translation happens under the
# hood.
#
# Adding a new warehouse
# ----------------------
# 1. Add an entry here, e.g. ``'MUM': 'DS_MUM_WH1'``.
# 2. That's it — the GUI dropdown auto-populates from this dict, and
#    D365 export auto-uses the mapped value when the user selects it.
#
# The FIRST entry is treated as the GUI's default selection. Keep AHD
# at the top unless the operational default changes.
#
# Dict values are the literal strings Business Central expects. Get them
# from the ERP team — typos here silently produce D365 imports the ERP
# will reject.
# ════════════════════════════════════════════════════════════════════════

WAREHOUSE_CODES: Dict[str, str] = {
    'AHD': 'PICK',          # Ahmedabad (default) — PICK warehouse
    'BLR': 'DS_BL_OFF1',    # Bangalore — Bangalore office 1 warehouse
    # Add more warehouses as needed, e.g.:
    #   'MUM': 'DS_MUM_WH1',
    #   'DEL': 'DS_DEL_WH1',
}

WAREHOUSE_DISPLAY_NAMES: List[str] = list(WAREHOUSE_CODES.keys())
DEFAULT_WAREHOUSE: str = WAREHOUSE_DISPLAY_NAMES[0]  # 'AHD'