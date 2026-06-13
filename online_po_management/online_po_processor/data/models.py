"""
data.models
===========

Pure-data classes that flow through the pipeline. No I/O, no business
logic — only field definitions and lightweight container behaviour
provided by ``@dataclass``.

Two types::

    SORow             — one row per ordered item line
    ProcessingResult  — what the engine returns to the exporter

The exporter consumes ``ProcessingResult.rows`` (a list of ``SORow``)
together with marketplace metadata to render the output workbook.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any, List, Optional, Tuple


@dataclass
class SORow:
    """
    Single line item destined for the SO Lines (and contributing to the
    SO Header for its PO).

    Field groups
    ------------
    Identity (always populated):
        ``po_number``, ``location``, ``item_no``, ``qty``

    Mapping resolution (populated by MarketplaceEngine after Ship-To
    lookup):
        ``cust_no``, ``ship_to``, ``mapped``, ``mapped_location``

    Master lookup + price validation (populated when MasterLoader is
    available and the EAN/Item No can be resolved):
        ``ean``, ``description``, ``mrp``, ``gst_code``,
        ``cost_price_ref`` (always the post-GST naked CP),
        ``calc_price`` (what's used for the active diff — depends on
        the marketplace's ``compare_basis``),
        ``fob_price`` (marketplace's price column),
        ``ref_fob_price`` (optional reference marketplace price),
        ``diffn`` = ``fob_price - calc_price``,
        ``ref_diffn`` = ``ref_fob_price - cost_price_ref``,
        ``validation_status`` ∈ ``{'OK', 'MISMATCH', 'NOT_IN_MASTER',
        'NO_PRICE', ''}``

    Raw input pass-through:
        ``unit_price`` — only set when ``price_col`` is configured (rare;
        both current marketplaces leave Unit Price blank for the WMS to
        compute downstream).
    """

    # Identity
    po_number: str
    location: str
    item_no: Any
    qty: int

    # Pass-through input
    unit_price: Optional[float] = None

    # v1.5.1: Marketplace-native row amount.
    # Some marketplaces (Blink, RK) report a per-row monetary value on
    # the punch file itself — this is the figure the marketplace will
    # invoice or settle for. We surface it here so the email report can
    # sum it for the Amount stat in the headline grid.
    #
    # Source column is declared per-marketplace via ``amount_col`` in
    # ``MARKETPLACE_CONFIGS``:
    #     Blink → 'total_amount'           (sum incl. taxes)
    #     RK    → 'Total accepted cost'    (qty × cost)
    #     Myntra → not configured yet      (stays None → email shows ₹0)
    #
    # When ``amount_col`` is missing or the cell is blank/NaN, this
    # stays ``None``. The email's aggregator treats None as 0 so the
    # Amount stat always renders consistently across marketplaces.
    amount: Optional[float] = None

    # Mapping resolution
    cust_no: str = ''
    ship_to: str = ''
    mapped: bool = False
    mapped_location: str = ''   # canonical mapping key matched to (may
                                # differ from raw `location` due to
                                # case-insensitive or fuzzy match)

    # Master lookup
    ean: str = ''
    description: str = ''       # from Items_March 'Description'

    # Marketplace prices
    fob_price: Optional[float] = None
    ref_fob_price: Optional[float] = None  # purely for reference Diffn
    # v2.3.1: the MRP the marketplace stated in its file (when the file
    # carries an MRP column, via config ``mrp_col``). Distinct from
    # ``mrp`` below, which is OUR master MRP. Lets the Validation sheet
    # show Vendor MRP vs Our MRP side by side. None when the file has no
    # MRP column.
    vendor_mrp: Optional[float] = None

    # v2.3.1: the margin (keep%) actually used to compute THIS row's
    # landing/cost. Normally the run margin, but per-line ``margin_rules``
    # (e.g. Nykaa's Perfume/Fragrance vs Cosmetics split) can override it.
    # The Validation sheet uses this for the row's Our Landing so the
    # display matches what was computed. None ⇒ the run margin was used.
    applied_margin_pct: Optional[float] = None

    # Our calculations
    calc_price: Optional[float] = None      # value used for the active diff
                                            # (basis='cost'    → MRP×m%÷GST,
                                            #  basis='landing' → MRP×m%)
    cost_price_ref: Optional[float] = None  # ALWAYS post-GST cost price
                                            # (the "naked CP"), shown for
                                            # reference even when basis
                                            # ≠ 'cost'

    # Diffs
    diffn: Optional[float] = None     # active: fob_price - calc_price
    ref_diffn: Optional[float] = None  # reference: ref_fob_price - cost_price_ref

    # Master attributes (raw, for display)
    mrp: Optional[float] = None
    gst_code: str = ''

    # v1.6.0: HSN cross-check (currently used by Reliance only).
    # Some marketplaces carry an HSN/SAC Code on their punch file for
    # each line — Reliance does, and discrepancies between the
    # marketplace's HSN and our master's HSN are a tax/compliance
    # risk (wrong HSN → wrong GST rate → ERP posts incorrect tax).
    # When the marketplace config declares ``hsn_col``, the engine
    # populates these three fields and surfaces mismatches on the
    # Validation sheet.
    #
    # All three stay at their defaults when the marketplace has no
    # ``hsn_col`` configured. The Validation sheet only adds its HSN
    # columns when at least one row has a non-empty
    # ``hsn_check_status``, so other marketplaces' output is
    # unaffected.
    hsn_punch: str = ''          # raw HSN from the marketplace file
    hsn_master: str = ''         # what Items_March.xlsx carries for this item
    hsn_check_status: str = ''   # '' | 'OK' | 'MISMATCH' | 'NOT_IN_MASTER'

    # v1.7.0: Multi-file upload traceability (Reliance-only today).
    # When the user uploads multiple Reliance PO files in one batch,
    # the engine tags every row produced from each file with the
    # PO number and location parsed from that file's title — the same
    # '5000466441  BHIWANDI (Reliance)' string Reliance puts on row 0
    # of the punch. The Raw Data sheet surfaces this as a leading
    # "Source" column so you can tell at a glance which file each
    # row came from when 5 POs are combined in one workbook.
    #
    # For single-file runs these stay at their defaults and the
    # Raw Data sheet omits the Source column (backward-compatible
    # with all Blink/Myntra/RK output layouts).
    source_po: str = ''          # e.g. '5000466441' (same as po_number for Reliance)
    source_location: str = ''    # e.g. 'BHIWANDI (Reliance)' (same as location for Reliance)

    # Validation status
    validation_status: str = ''


@dataclass
class ProcessingResult:
    """
    Output of ``MarketplaceEngine.process()``. Consumed by ``SOExporter``.

    Holds the rendered SORows plus the metadata the exporter needs to
    label sheets correctly (compare_basis affects Validation column
    headers, etc.).
    """

    # Per-row results
    rows: List[SORow] = field(default_factory=list)

    # Warnings to surface in the GUI log + Warnings sheet:
    # tuples of (po, location, message). PO and location may be empty
    # strings for global warnings.
    warnings: List[Tuple[str, str, str]] = field(default_factory=list)

    # Marketplace context
    marketplace: str = ''
    input_file: str = ''            # basename, for display
    input_file_path: str = ''       # full path — used to locate the
                                    # output/ folder next to the input
    margin_pct: float = 0.70        # decimal (0.70 = 70%)
    compare_basis: str = 'cost'     # 'landing' | 'cost'
    compare_label: str = 'Price'    # friendly label shown in Validation

    # v1.7.0: Multi-file upload counter. Stays at 1 for the ordinary
    # single-file flow; set by ``MarketplaceEngine.process_multi`` to
    # the number of files aggregated when a batch upload was run
    # (currently Reliance only). The email builder and SO filename
    # generator use this to label the output appropriately —
    # "Reliance_5PO" vs "Reliance_SO" for example.
    input_files_count: int = 1

    # v1.9.0: Warehouse ERP location code used for D365 export.
    # Before v1.9.0 this was hardcoded to 'PICK' in the exporter.
    # The GUI now exposes a Warehouse dropdown (AHD/BLR/...) and
    # stamps the selected warehouse's ERP code onto the result so
    # D365Exporter can use it uniformly on Sales Header col K and
    # Sales Line col F. Stays at 'PICK' when unset for backwards
    # compatibility with any code paths that construct a
    # ProcessingResult directly without going through the GUI.
    warehouse_code: str = 'PICK'
    warehouse_display: str = 'AHD'   # friendly label, for Summary footer + email

    # Original DataFrame (for the Raw Data sheet)
    raw_df: Any = None

    # Runtime metrics (populated by the GUI after engine+exporter complete).
    # Used by the email report footer to show "Processing: X.XX seconds"
    # and by the D365 export popup to show timing. None until measured.
    elapsed_seconds: Optional[float] = None

    # v2.0.0: Output type — 'so' (default) or 'to'.
    #
    # Set by ``MarketplaceEngine.process()`` when the marketplace's
    # config declares ``output_type='to'`` (currently Flipkart-TO).
    # Drives downstream branching:
    #   * ``D365Exporter`` fills 'Transfer Header'/'Transfer Line'
    #     sheets (vs. 'Sales Header'/'Sales Line' for SO).
    #   * Sheet writers (Validation, Summary, Headers, Lines) drop
    #     SO-specific columns (price compare, customer no.) when 'to'.
    #   * Email builder swaps the headline metric (Amount → Total Qty).
    #
    # Stays 'so' for all existing marketplaces (Blink/Myntra/Reliance/
    # etc.) — they don't set ``output_type`` in their configs, so the
    # engine never overrides this default.
    output_type: str = 'so'

    # v2.1.3: Runtime flag — write the engine-computed Cost Price into
    # the Lines (SO) ``Unit Price`` column (col 8) and the D365 Sales
    # Line ``Unit Price`` column (col H), instead of leaving them blank
    # for downstream WMS / vendor-master defaulting.
    #
    # Driven by the GUI's "Override Unit Price" checkbox, NOT by any
    # config flag. The marketplace config's legacy
    # ``override_unit_price=True`` (currently only on BlinkMP) is now a
    # GUI default-state hint — it pre-checks the box on marketplace
    # change but doesn't enforce override at runtime.
    #
    # Why a runtime field instead of a config-driven one
    # ---------------------------------------------------
    # Operations occasionally need BlinkMP to NOT override (e.g. when
    # validating a vendor-master refresh, or processing a one-off batch
    # under the old margin agreement). Conversely, they sometimes want
    # to force-override on a marketplace that doesn't normally need it
    # (one-off price corrections). A per-run toggle covers both cases
    # without config edits and without committing to a permanent change.
    #
    # Stays ``False`` for runs where the GUI doesn't set it (older code
    # paths, tests, or marketplaces processed via the API directly).
    override_unit_price: bool = False

    # v1.5.6: Alias-resolved marketplace config.
    #
    # The engine's ``_resolve_column_aliases`` picks a concrete header
    # name for every list-valued ``*_col`` config entry based on what
    # actually appears in the uploaded punch file. Downstream consumers
    # (raw_data_sheet, validation_sheet, etc.) must use these resolved
    # names rather than re-reading ``MARKETPLACE_CONFIGS`` — otherwise
    # they hit the unresolved list values and crash with errors like
    # ``TypeError: unhashable type: 'list'`` when doing
    # ``col_name in df.columns``.
    #
    # Populated by :meth:`MarketplaceEngine.process` as soon as the
    # file is read; stays ``None`` if processing aborted before that
    # point (e.g. file couldn't be opened).
    resolved_config: Optional[dict] = None

    # v2.4.0: dedup-skip. When DEDUP_SKIP_ENABLED, ``apply_dedup`` removes
    # already-uploaded POs from ``rows`` (so they don't reach Headers/Lines)
    # and records a summary of each removed PO here. Consumed by the
    # "Skipped" output sheet and logged to the ``dedup_skips`` table.
    # Each entry: {segment, marketplace, marketplace_label, po, location,
    # qty, order_value, first_seen}. Empty when nothing was skipped.
    skipped_orders: List[dict] = field(default_factory=list)