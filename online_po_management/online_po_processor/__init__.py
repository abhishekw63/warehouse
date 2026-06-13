"""
online_po_processor
===================

Marketplace PO/punch file → ERP-importable Sales Order generator.

This package replaces the single-file ``standalone_po_processing.py`` script
(now retained as ``legacy_standalone.py`` for fallback). Logic is identical
through v1.4.0 — only the file layout changed.

Layout overview
---------------
::

    online_po_processor/
        config/      → constants, marketplace registry, paths/history helpers
        data/        → pure-data classes (SORow, ProcessingResult) and loaders
        engine/      → MarketplaceEngine — turns a punch file into result rows
        exporter/    → SOExporter + D365Exporter + per-sheet writers
        emailer/     → HTML report builder + SMTP sender
        gui/         → Tkinter UI (OnlinePOApp + dialogs)
        utils/       → cross-platform helpers (open_file)
        app.py       → bootstrap: expiry check + main() entry point

Quick start
-----------
The intended entry point is the top-level ``main.py`` in the project root::

    python main.py

That file does nothing more than::

    from online_po_processor.app import main
    main()

Public re-exports
-----------------
The most commonly imported names are exposed at package level for
convenience and to mirror what the legacy single-file module exported:

    >>> from online_po_processor import (
    ...     OnlinePOApp,
    ...     MARKETPLACE_CONFIGS,
    ...     SORow,
    ...     ProcessingResult,
    ... )

Changelog
---------
    v2.3.1 — Two changes:
             1) Location Code on Headers (SO) + Lines (SO) now follows
                the selected dispatch warehouse (AHD/BLR/...) via
                ``result.warehouse_code`` — the same value the Summary
                footer reports — instead of the hardcoded 'PICK'. Falls
                back to 'PICK' when no warehouse is selected.
             2) Flipkart-TO bulk consignment input mode. Alongside the
                historical single consolidated dump, the operator can
                now hand in Flipkart's raw per-PO consignment CSVs
                (Consignment_Details_<PO>_<date>.csv) directly. The new
                ``MarketplaceEngine.process_consignments`` assembles the
                consolidated dump in memory — PO read from each filename —
                then runs the standard TO pipeline. Mode is chosen via a
                GUI selector; the consolidated path is unchanged.
                Location is supplied by an optional Consignment Visibility
                Report (PO → Warehouse Id). The machine-code Warehouse Id
                is bridged to a Transfer-to Code via a PROVISIONAL
                internal alias map (config ``warehouse_aliases``) that
                translates it to the friendly Ship-To B2B name; every
                aliased PO is flagged as a fuzzy/provisional match on the
                Warnings sheet until exact codes are wired in. Without the
                report (or for a PO absent from it) Location is empty.
             3) Meesho-TO marketplace (second Transfer Order channel).
                Reuses the consignment pipeline with a different file
                shape (``order-line-items-<PO>.csv``: PO from filename,
                EAN from ``styleCode``, qty from ``orderedQty``).
                Bulk-consignment-ONLY (no consolidated dump); Location
                left empty until its destination-warehouse source is
                decided; margin defaults to a 60% placeholder pending
                confirmation. The GUI's input-mode selector is now fully
                config-driven (any marketplace opts in via a
                ``consignment_mode`` block), not hard-coded to Flipkart-TO.
             4) FirstCry marketplace (second PDF-source channel, standard
                SO). New engine/firstcry_pdf_parser.py reads the bordered
                line-item table via pdfplumber.extract_tables() with
                name-based column mapping, exposing __po__/__loc__ like the
                Avenue/Dmart parser. EAN comes from the 'Manufacturer'
                column (verified vs master); Landing Rate = MRP × 60%
                (compare_basis 'landing', margin 60). Ship-to resolves from
                'Delivered To' against existing Firstcry Ship-To B2B rows.
             5) Validation sheet now shows Vendor vs Our side by side for
                MRP / Landing / CP (amber-flagged on mismatch); FirstCry
                also diffs Cost Price (ref_fob_col 'Base Cost'). HSN
                mismatch downgraded from red error to amber alert.
             6) Per-line margin rules + Nykaa marketplace. Generic config
                ``margin_rules`` computes landing/cost PER LINE by product
                category (substring and/or HSN-prefix match); each
                non-default match is logged on the Warnings sheet. Nykaa
                (CSV, compare_basis 'cost' vs 'Unit Cost') uses it for its
                Perfume/Fragrance (keep 69%) vs Cosmetics (keep 66%) split.
                Adds SORow.applied_margin_pct.
             7) Reliance rewritten as PDF-based (replaces the old Excel/
                pre_process flow). New engine/reliance_pdf_parser.py reads
                the multi-page PO via pdfplumber.extract_tables(); ship-to
                is the delivery city. GST-DEPENDENT margin: new config
                ``gst_margin_discount`` makes the per-item keep% =
                1 − discount × (1+GST), reproducing Reliance's table
                (69%/67.45%/63.42% at 0/5/18% GST). MasterLoader gains a
                shared ``gst_divisor`` helper.
"""

# Aligned to the inline vX.Y.Z changelog tag series (latest: v2.3.1).
# Was historically out of sync at "1.9.3" while inline tags had reached
# v2.3.0; bumped here per the v2.3.1 Location Code fix.
__version__ = "2.3.1"
__all__ = [
    "__version__",
    # Re-exports for code that used to ``import standalone_po_processing as opp``
    "OnlinePOApp",
    "MARKETPLACE_CONFIGS",
    "MARKETPLACE_NAMES",
    "WAREHOUSE_CODES",
    "WAREHOUSE_DISPLAY_NAMES",
    "DEFAULT_WAREHOUSE",
    "SORow",
    "ProcessingResult",
    "MasterLoader",
    "MappingLoader",
    "MarketplaceEngine",
    "SOExporter",
    "D365Exporter",
    "EmailBuilder",
    "EmailSender",
    "get_email_config",
    "main",
]

# --- public re-exports (kept thin; all real code lives in submodules) -------
from online_po_processor.config.email_config import get_email_config
from online_po_processor.config.marketplaces import (
    DEFAULT_WAREHOUSE,
    MARKETPLACE_CONFIGS,
    MARKETPLACE_NAMES,
    WAREHOUSE_CODES,
    WAREHOUSE_DISPLAY_NAMES,
)
from online_po_processor.data.mapping_loader import MappingLoader
from online_po_processor.data.master_loader import MasterLoader
from online_po_processor.data.models import ProcessingResult, SORow
from online_po_processor.emailer import EmailBuilder, EmailSender
from online_po_processor.engine.marketplace_engine import MarketplaceEngine
from online_po_processor.exporter.d365_exporter import D365Exporter
from online_po_processor.exporter.so_exporter import SOExporter
from online_po_processor.gui.app_window import OnlinePOApp
from online_po_processor.app import main