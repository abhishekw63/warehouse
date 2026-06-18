"""
engine.marketplace_engine
=========================

Turns a marketplace punch file into a list of ``SORow`` rows ready for
the exporter. This is the heart of the pipeline:

#. Read the punch file (Excel).
#. Validate the columns we need exist (depends on ``item_resolution``).
#. For each row:
   * Parse identity (PO, location, qty).
   * Resolve EAN, then Item No (either from ``item_col`` or by EAN→master
     lookup).
   * Look up MRP / GST / Description from the master.
   * Compute our calculated price (``calc_price``) per the marketplace's
     ``compare_basis`` (landing or cost).
   * Compute the diff against the marketplace's quoted price; flag
     mismatches.
   * Look up the delivery location in the mapping registry.
#. Return a ``ProcessingResult``.

The engine never raises on per-row data problems — it appends to
``result.warnings`` so the GUI can surface them. The only fatal failures
return early after a warning: missing required columns, missing master
when EAN-resolution requires it, etc.
"""

from __future__ import annotations
import logging
import os
import re
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

import pandas as pd

from online_po_processor.data.models import SORow, ProcessingResult
from online_po_processor.data.master_loader import MasterLoader
from online_po_processor.data.mapping_loader import MappingLoader
# v2.2.0: PDF parser dispatch. Each callable takes a filepath and
# returns a ``pandas.DataFrame`` shaped the same way the engine would
# expect from an Excel read — i.e. with the column names referenced by
# the marketplace's config (``po_col``, ``ean_col``, ``qty_col``, etc.).
# New PDF marketplaces add an entry to ``PDF_PARSERS`` and set
# ``source_format='pdf'`` + ``pdf_parser='<key>'`` in their config.
from online_po_processor.engine.avenue_pdf_parser import (
    load_avenue_pdf_as_dataframe,
)
from online_po_processor.engine.firstcry_pdf_parser import (
    load_firstcry_pdf_as_dataframe,
)
from online_po_processor.engine.reliance_pdf_parser import (
    load_reliance_pdf_as_dataframe,
)
from online_po_processor.engine.myntra_pdf_parser import (
    load_myntra_pdf_as_dataframe,
)
from online_po_processor.engine.bigbasket_parser import (
    load_bigbasket_excel_as_dataframe,
)
from online_po_processor.engine.purplle_parser import (
    load_purplle_export_as_dataframe,
)
from online_po_processor.engine.flipkart_dump_parser import (
    load_flipkart_po_as_dataframe,
)


# Registry of PDF parsers — keyed by the ``pdf_parser`` config value.
# Each parser is a callable: ``(filepath: str) -> pd.DataFrame``.
PDF_PARSERS: Dict[str, Callable[[str], pd.DataFrame]] = {
    'avenue': load_avenue_pdf_as_dataframe,
    'firstcry': load_firstcry_pdf_as_dataframe,
    'reliance': load_reliance_pdf_as_dataframe,
    # v2.4.0: Myntra is dual-format (Excel + PDF) — see the extension
    # routing in process() and 'pdf_parser'/'accepted_extensions' in its
    # config. The Excel-only / PDF-only marketplaces are unaffected.
    'myntra': load_myntra_pdf_as_dataframe,
    # v2.7: Big Basket — a custom EXCEL parser (preamble + table), routed
    # via the config's ``file_parser='bigbasket'`` key (not extension).
    'bigbasket': load_bigbasket_excel_as_dataframe,
    # Purplle: tab-separated '.XLS' (SAP export) — routed by file_parser.
    'purplle': load_purplle_export_as_dataframe,
    # v2.7.x: Flipkart — the new portal emits ONE 'purchase_order_<PO>.xlsx'
    # per PO (two-row header). Routed via ``file_parser='flipkart'``; the
    # operator drops all of the day's PO files → process_multi compiles them
    # into one batch in memory (replaces the old standalone dump generator).
    'flipkart': load_flipkart_po_as_dataframe,
}


# Set of GST codes we know how to handle. Anything outside this set
# triggers a warning and falls back to 18% in MasterLoader.
# v1.9.1: added 'G-0' and 'G-0-S' as recognised 0% codes — the
# cost calculation already handled them in calc_cost_price, but the
# engine was incorrectly flagging them as "unknown GST code" in a
# warning because they weren't listed here.
_KNOWN_GST_CODES = frozenset({
    '0-G', 'G-0', 'G-0-S',
    'G-3', 'G-3-S',
    'G-5', 'G-5-S',
    'G-12', 'G-12-S',
    'G-18', 'G-18-S',
    '',
})


class MarketplaceEngine:
    """
    Apply per-marketplace config rules to a punch file.

    Args:
        mapping: Loaded ``MappingLoader`` for the selected marketplace.
        master:  Loaded ``MasterLoader``, or ``None``. Required when the
                 marketplace's ``item_resolution`` is ``'from_ean'`` —
                 otherwise the engine has no way to derive Item No. When
                 ``None`` and resolution is ``'from_column'``, price
                 validation is silently disabled (rows still pass through).
    """

    # Threshold for flagging price mismatches (rupees). Diffs at or below
    # this are treated as rounding noise → status='OK'. Above this →
    # 'MISMATCH' with a warning row.
    DIFFN_THRESHOLD: float = 1.0

    def __init__(self, mapping: MappingLoader,
                 master: Optional[MasterLoader] = None) -> None:
        self.mapping = mapping
        self.master = master

    # ── Public entry point ─────────────────────────────────────────────

    def process_multi(
        self,
        filepaths: List[str],
        config: Dict[str, Any],
        margin_pct: float = 0.70,
    ) -> ProcessingResult:
        """
        Batch-process multiple punch files into ONE combined result.

        Currently used exclusively by Reliance where the user receives
        one PO file per order and wants to process 5 POs (5 files) as
        a single SO batch. Other marketplaces don't call this because
        they consolidate POs inside a single file natively (Blink's
        35 POs in one punch, Myntra's 4 POs in one dump, etc.).

        Per-file failures are isolated: if file 3 of 5 has a bad
        title row or a missing sheet, files 1/2/4/5 still produce
        SORows in the combined output and file 3's error appears in
        the Warnings sheet (prefixed with the bad file's basename
        so the user can tell which upload caused it). A single bad
        file never aborts the whole batch.

        The returned ``ProcessingResult``:
          * ``rows`` is the concatenation of all files' rows.
          * ``warnings`` is all files' warnings, each prefixed with
            ``[<filename>]`` when the warning came from a specific
            file (vs a batch-level warning which has no prefix).
          * ``input_file`` is the basename of the FIRST file (for
            display in the email banner header).
          * ``input_file_path`` is the full path of the FIRST file
            (so SOExporter writes output next to the batch's first
            file — typical case: all files live in the same folder).
          * ``input_files_count`` is ``len(filepaths)``.
          * ``raw_df`` is the vertical concatenation of all files'
            DataFrames with source columns preserved.

        Args:
            filepaths:  List of Excel file paths. Must not be empty.
            config:     Marketplace config (Reliance's entry).
            margin_pct: Run margin as decimal.

        Returns:
            Combined ``ProcessingResult``. Callers that want to know
            which files contributed can inspect each row's
            ``source_po``/``source_location``.
        """
        if not filepaths:
            # Return an empty result with a batch-level warning so the
            # GUI can show "no files selected" rather than crash.
            empty = ProcessingResult(marketplace=config['party_name'])
            empty.warnings.append((
                '', '',
                "process_multi called with empty file list — nothing "
                "to process."
            ))
            return empty

        if len(filepaths) == 1:
            # Single-file batch is just a delegated single-file run.
            # No tagging needed, no concatenation overhead.
            r = self.process(filepaths[0], config, margin_pct)
            r.input_files_count = 1
            return r

        # ── Multi-file path ────────────────────────────────────────────
        logging.info("process_multi: starting batch of %d files",
                     len(filepaths))

        combined = ProcessingResult(
            marketplace=config['party_name'],
            input_file=os.path.basename(filepaths[0]),
            input_file_path=filepaths[0],
            margin_pct=margin_pct,
            compare_basis=config.get('compare_basis', 'cost'),
            compare_label=config.get('compare_label', 'Price'),
            input_files_count=len(filepaths),
        )

        per_file_dfs: List[pd.DataFrame] = []

        for idx, fp in enumerate(filepaths, start=1):
            basename = os.path.basename(fp)
            logging.info("process_multi: [%d/%d] processing %s",
                         idx, len(filepaths), basename)
            try:
                sub = self.process(fp, config, margin_pct)
            except Exception as e:  # noqa: BLE001
                # Never let a single file crash the batch. Log the
                # failure, tag it with the filename, move on.
                logging.exception(
                    "process_multi: file %s raised an exception",
                    basename,
                )
                combined.warnings.append((
                    '', '',
                    f"[{basename}] Failed to process: {e}"
                ))
                continue

            # Merge: rows pass through untouched — their source_po and
            # source_location were already tagged by _process_row.
            combined.rows.extend(sub.rows)

            # v2.4.2: carry per-file Master-Exception records into the
            # combined result so the Exceptions sheet is populated for
            # MULTI-FILE marketplaces too (Myntra Goddess vendor-CP, Flipkart,
            # …). Without this they were recorded on each sub-result and lost
            # at merge time — the sheet came up empty despite rows being
            # flagged/highlighted.
            if getattr(sub, 'exceptions_applied', None):
                combined.exceptions_applied.extend(sub.exceptions_applied)

            # v2.4.3: the registry is the SAME full list for every file (it's
            # the master's, not per-file), so copy it once.
            if (not combined.exception_registry
                    and getattr(sub, 'exception_registry', None)):
                combined.exception_registry = sub.exception_registry

            # Prefix every warning with the filename so the user can
            # tell which upload caused which warning in a combined
            # batch. Batch-level warnings (empty PO + empty location
            # tuple components from process_multi itself) stay
            # unprefixed.
            for po, loc, msg in sub.warnings:
                combined.warnings.append((po, loc, f"[{basename}] {msg}"))

            # Keep the file's resolved config on the combined result.
            # All files in a batch share the same marketplace config so
            # the last one wins (they're all equivalent anyway).
            if sub.resolved_config is not None:
                combined.resolved_config = sub.resolved_config

            # Accumulate raw DataFrames for later concatenation. Tag
            # each with a __source_file__ column so Raw Data can still
            # distinguish them if someone wants to filter. Source PO
            # and location live on SORows themselves (more precise).
            if sub.raw_df is not None and not sub.raw_df.empty:
                df_tagged = sub.raw_df.copy()
                df_tagged['__source_file__'] = basename
                per_file_dfs.append(df_tagged)

        # Concatenate raw_df across files. ``sort=False`` keeps columns
        # in the order of the first file; missing columns in later
        # files just become NaN in the combined frame.
        if per_file_dfs:
            combined.raw_df = pd.concat(
                per_file_dfs, ignore_index=True, sort=False,
            )
        else:
            combined.raw_df = pd.DataFrame()

        logging.info(
            "process_multi: batch done — %d total rows from %d files, "
            "%d warnings",
            len(combined.rows), len(filepaths), len(combined.warnings),
        )
        return combined

    def process(self, filepath: str, config: Dict[str, Any],
                margin_pct: float = 0.70) -> ProcessingResult:
        """
        Read ``filepath`` and produce a ``ProcessingResult``.

        Args:
            filepath:   Path to the punch/PO Excel file.
            config:     Entry from
                        :data:`~online_po_processor.config.marketplaces.MARKETPLACE_CONFIGS`.
            margin_pct: Margin as decimal (e.g. ``0.70`` for 70%).

        Returns:
            Always returns a result, even if the read failed or all rows
            were skipped. Inspect ``result.warnings`` and
            ``len(result.rows)`` to detect problems.
        """
        result = ProcessingResult(
            marketplace=config['party_name'],
            input_file=os.path.basename(filepath),
            input_file_path=str(filepath),  # used for output/ folder location
            compare_basis=config.get('compare_basis', 'cost'),
            compare_label=config.get('compare_label', 'Price'),
            margin_pct=margin_pct,
        )

        # ── Read file ───────────────────────────────────────────────────
        # v2.2.0: Marketplace files arrive in two formats — Excel (the
        # original and most common: Blink, Myntra, RK, Reliance, Zepto,
        # BlinkMP) and PDF (Avenue/DMart Ready, added v2.2.0). The
        # config's ``source_format`` key selects the load path. Default
        # ``'excel'`` preserves all pre-existing behaviour exactly.
        source_format = config.get('source_format', 'excel')

        # v2.4.0: dual-format marketplaces (Myntra) keep their regular
        # ``source_format='excel'`` AND register a ``pdf_parser`` so they
        # also accept the PO PDF. Route by the *uploaded file's*
        # extension: a ``.pdf`` whose config has a registered parser goes
        # through the PDF path; everything else uses the configured
        # ``source_format``. Single-format marketplaces (Excel-only and
        # PDF-only) are unaffected — Excel-only configs have no
        # ``pdf_parser``, PDF-only configs already set source_format='pdf'.
        if (os.path.splitext(filepath)[1].lower() == '.pdf'
                and config.get('pdf_parser')):
            source_format = 'pdf'

        # v2.7: a ``file_parser`` config key routes ANY file (e.g. Big
        # Basket's custom-layout Excel) through a registered parser →
        # DataFrame, regardless of extension. ``pdf_parser`` is the
        # extension-based special case of the same mechanism.
        parser_key = (config.get('pdf_parser') if source_format == 'pdf'
                      else config.get('file_parser'))

        if parser_key:
            df = self._load_pdf(filepath, config, result, parser_key=parser_key)
            if df is None:
                # _load_pdf has already appended an abort warning.
                return result
            # Parsers produce a DataFrame with the same columns the engine's
            # downstream logic expects (po_col / loc_col / qty_col /
            # ean_col / etc.). No sheet or header_row dance needed.
            result.raw_df = df
            logging.info("Read %d rows from %s",
                         len(df), os.path.basename(filepath))
            # Skip the Excel-specific sheet resolution + pre_process
            # hook chain; jump straight to column alias resolution.
            # This is done by falling through to the existing
            # ``_resolve_column_aliases`` call below (which is shared).

        else:
            # v2.3.0: CSV early-return — some marketplaces (Blink) now
            # export their dump in CSV format alongside the historical
            # xlsx. We auto-detect by file extension and read CSV
            # directly without going through the Excel sheet-selection
            # / pre_process machinery (which is meaningless for CSVs).
            #
            # The downstream column-resolution + mapping + validation
            # pipeline is shape-agnostic — it operates on a DataFrame
            # regardless of whether it came from CSV, Excel, or PDF.
            # So a Blink CSV with the same column names as the Blink
            # xlsx (`po_number`, `units_ordered`, `cost_price`,
            # `total_amount`, `mrp`, `upc`, `facility_name`, ...) plugs
            # in without any config change. Column-name *changes*
            # between the CSV and xlsx exports would still need
            # ``case_insensitive_cols`` / column aliases.
            ext = os.path.splitext(filepath)[1].lower()
            if ext == '.csv':
                try:
                    # ``header`` follows the same config key Excel uses,
                    # so a marketplace can opt into header_row=1 (etc.)
                    # for both formats consistently.
                    df = pd.read_csv(
                        filepath, header=config.get('header_row', 0),
                    )
                except Exception as e:  # noqa: BLE001 — surface ANY CSV-read error
                    result.warnings.append((
                        '', '',
                        f"Cannot read CSV: {e}"
                    ))
                    return result
                logging.info("Read %d rows from CSV %s",
                             len(df), os.path.basename(filepath))
                result.raw_df = df
                # No sheet selection. No pre_process hook. CSVs are
                # always flat — any per-marketplace pre-processing
                # that exists today (Reliance's merged-cell title parse)
                # is specific to Excel-only marketplaces.
                # Fall through to the shared column-resolution code.

            else:
                # v1.4.2: Most marketplace punch files carry data on 'Sheet1'.
                # v1.6.0: Reliance is the exception — its file is the raw PO
                # attachment that has 6 sheets, with clean flat data on a
                # sheet literally called 'PO' (the other sheets are messy
                # auto-generated renderings of the same data). So the config
                # can now override the sheet name via ``source_sheet``.
                #
                # Other sheets in the workbook are user-side pivots / manual
                # calc / sidecars that must NOT be read by the script.
                # Previously we defaulted to pandas' "first sheet" behavior
                # which silently latched onto Sheet2/Sheet4 when a pivot sheet
                # came first — that's now fixed across the board.
                #
                # If the target sheet doesn't exist, we fall back to the first
                # sheet and log a warning so the user can spot the issue
                # rather than getting a cryptic KeyError downstream.
                try:
                    available_sheets = pd.ExcelFile(filepath).sheet_names
                except Exception as e:  # noqa: BLE001
                    result.warnings.append((
                        '', '', f"Cannot open file: {e}"
                    ))
                    return result

                # Per-marketplace sheet override. Values:
                #   * Omitted / ``'Sheet1'`` — most marketplaces
                #   * ``'PO'`` — Reliance (exact sheet name)
                #   * ``'PO_*'`` — Zepto (prefix match) (v1.8.0)
                #
                # Zepto's dumps put the data on a sheet whose name varies per
                # export — literally ``PO_<random-hex>`` like
                # ``PO_64863340b23e6c90`` or ``PO_c881cfb0a4fa2ebc``. The
                # wildcard lets us match any of them without reconfiguring
                # per file. If multiple matching sheets exist we take the
                # first; duplicates aren't expected and would indicate a
                # malformed dump.
                target_sheet = config.get('source_sheet', 'Sheet1')

                sheet_to_read = self._resolve_source_sheet(
                    target_sheet, available_sheets, result,
                )
                if sheet_to_read is None:
                    # _resolve_source_sheet has appended an abort warning.
                    return result

                # Header row is configurable too — Reliance's 'PO' sheet has
                # its header on row 1 (the title merged-cell occupies row 0),
                # while everyone else is 0-indexed.
                header_row = config.get('header_row', 0)

                try:
                    df = pd.read_excel(
                        filepath, sheet_name=sheet_to_read, header=header_row,
                    )
                except Exception as e:  # noqa: BLE001 — we want to surface ANY read error
                    result.warnings.append((
                        '', '',
                        f"Cannot read sheet {sheet_to_read!r}: {e}"
                    ))
                    return result

                logging.info("Read %d rows from %s",
                             len(df), os.path.basename(filepath))

                # ── Marketplace-specific pre-processor ──────────────────────────
                result.raw_df = df

        # v2.3.1: hand off to the shared post-load pipeline. Everything
        # from column-alias resolution onward is identical regardless of
        # whether the DataFrame came from a single Excel/CSV/PDF file
        # (this method) or from the in-memory consolidated dump that
        # ``process_consignments`` builds from many Flipkart consignment
        # CSVs. Factoring it out keeps the two entry points in lock-step.
        return self._process_loaded_dataframe(df, config, margin_pct, result)

    def _process_loaded_dataframe(
        self,
        df: 'pd.DataFrame',
        config: Dict[str, Any],
        margin_pct: float,
        result: ProcessingResult,
    ) -> ProcessingResult:
        """
        Shared post-load pipeline: column-alias resolution → required-
        column validation → SO/TO branch.

        v2.3.1: extracted verbatim from the back half of :meth:`process`
        so it can be reused by :meth:`process_consignments`, which feeds
        an in-memory consolidated DataFrame (built from many Flipkart
        consignment CSVs) through exactly the same logic. The behaviour
        is unchanged for every existing caller — :meth:`process` simply
        calls this instead of inlining the steps.

        Pre-conditions (the caller must have already done these):
          * ``result.raw_df`` is set to ``df`` (so the Raw Data sheet has
            the source frame even if validation fails early here).
          * ``result`` carries the marketplace metadata
            (``marketplace`` / ``input_file`` / ``compare_*`` / margin).

        Args:
            df:         The loaded (or assembled) source DataFrame.
            config:     Marketplace config — column keys may still be in
                        list-alias form; this method resolves them.
            margin_pct: Decimal margin for the run.
            result:     ProcessingResult to populate in place.

        Returns:
            The same ``result`` with ``rows`` populated (and any
            warnings appended). Returns early — with whatever rows exist
            so far, i.e. none — on a fatal column problem.
        """
        # ── v1.5.5: Resolve column aliases against actual headers ──────
        # Marketplace configs may declare a column key as a LIST of
        # acceptable names when the marketplace's punch file sometimes
        # arrives with different headers for the same field. Myntra is
        # the canonical case: the PO column is sometimes labeled 'PO'
        # and sometimes 'PO Number' depending on which dashboard
        # exported the dump. We pick the first name that exists in the
        # DataFrame and collapse the list back to a single string, so
        # the rest of the pipeline sees a normal scalar config with no
        # awareness that aliases ever existed.
        #
        # Works for any column key that could reasonably have variant
        # names: po_col, loc_col, qty_col, ean_col, item_col, fob_col,
        # amount_col, etc. Non-list values pass through untouched
        # (backward-compatible).
        config = self._resolve_column_aliases(config, df.columns, result)
        if config is None:
            # _resolve_column_aliases already appended a warning
            return result

        # v1.5.6: stash the resolved config on the result so
        # downstream exporter sheets can read alias-resolved column
        # names directly instead of hitting the original module-level
        # MARKETPLACE_CONFIGS (which still has list values). Without
        # this, raw_data_sheet crashed with "unhashable type: 'list'"
        # when doing ``col in df.columns`` against the unresolved list.
        result.resolved_config = config

        # v2.4.3: carry the full cross-marketplace exception registry onto the
        # result so the Exceptions sheet can list ALL marketplaces' exceptions
        # (highlighting this marketplace's own). Independent of which fired.
        if self.master is not None:
            result.exception_registry = getattr(
                self.master, 'exception_registry', []) or []

        # ── Required-column validation ──────────────────────────────────
        if not self._validate_required_columns(df, config, result):
            return result

        # Optional columns — log misses but keep going
        price_col = self._validate_optional_column(df, config, 'price_col')
        self._validate_optional_column(df, config, 'fob_col',
                                        log_warn_to_result=result)
        self._validate_optional_column(df, config, 'ean_col',
                                        log_warn_to_result=result)

        # ── v2.0.0: Branch by output_type ────────────────────────────────
        # Marketplaces with ``output_type='to'`` (currently Flipkart-TO)
        # produce Transfer Orders, not Sales Orders. The TO pipeline is
        # structurally simpler:
        #
        #   * No Items_March master lookup — Item No, MRP, and GST come
        #     from the file's own columns
        #   * No price validation — there's no fob_col to compare against
        #   * No HSN cross-check — no master to compare against
        #   * Rows with same Item No within a PO are aggregated (qty
        #     summed) so D365 sees one Transfer Line per Item No
        #   * Mapping returns Transfer-to Code (in the 'ship_to' field
        #     of the Ship-To B2B sheet); 'cust_no' stays empty because
        #     TOs have no customer
        #
        # The SO marketplaces above this branch are completely
        # unaffected — they don't set ``output_type``, so it defaults
        # to 'so' and they hit the per-row processing loop below as
        # before.
        if config.get('output_type') == 'to':
            result.output_type = 'to'
            return self._process_to(df, config, margin_pct, result)

        # ── Per-row processing (SO path — unchanged) ────────────────────
        warned_keys: Set[Tuple] = set()  # dedupe warnings (e.g. one per PO)
        item_resolution = config.get('item_resolution', 'from_column')
        compare_basis = config.get('compare_basis', 'cost')
        compare_label = config.get('compare_label', 'Price')

        # v2.4.4: PO-status REVIEW (Swiggy). When the config names a
        # ``status_col`` + ``status_keep`` set, every line whose PO is in any
        # other state (EXPIRED / COMPLETED / CANCELLED / PENDING) is KEPT in the
        # output but FLAGGED — one named warning per such PO — so the operator
        # can manually audit and remove it. We do NOT auto-drop: a status can
        # be wrongly given, so the human makes the final call (golden rule:
        # nothing skipped silently, and nothing skipped at all here).
        self._flag_po_status(df, config, result)

        for _, row in df.iterrows():
            so_row = self._process_row(
                row=row,
                df=df,
                config=config,
                item_resolution=item_resolution,
                compare_basis=compare_basis,
                compare_label=compare_label,
                margin_pct=margin_pct,
                price_col=price_col,
                result=result,
                warned_keys=warned_keys,
            )
            if so_row is not None:
                result.rows.append(so_row)

        logging.info("Processed %d items across %d PO(s)",
                     len(result.rows),
                     len({r.po_number for r in result.rows}))
        return result

    # ── Flipkart-TO bulk consignment mode (v2.3.1) ──────────────────────

    @staticmethod
    def _extract_po_from_filename(
        filepath: str,
        regex: str = r'Consignment_Details_([^_]+)_',
    ) -> str:
        """
        Pull the PO number out of a Flipkart consignment filename.

        Flipkart exports one consignment CSV per PO, named
        ``Consignment_Details_<PO>_<DD-MM-YYYY>.csv`` — e.g.
        ``Consignment_Details_204535220_09-06-2026.csv`` → ``204535220``.
        The PO is NOT a column inside the file, so the bulk-consignment
        flow recovers it from the filename instead.

        Args:
            filepath: Full path or bare filename of the consignment CSV.
            regex:    Capture-group-1 pattern applied to the basename.
                      Overridable via the marketplace config's
                      ``consignment_mode['filename_po_regex']`` so the
                      naming convention can change without code edits.

        Returns:
            The extracted PO string (stripped), or ``''`` when neither
            the configured pattern nor the digit-run fallback matches —
            the caller treats an empty return as "skip this file with a
            warning".
        """
        name = os.path.basename(filepath)
        m = re.search(regex, name)
        if m:
            return m.group(1).strip()
        # Defensive fallback: if the strict pattern drifts (renamed
        # export, extra prefix), take the longest run of digits in the
        # stem — Flipkart PO numbers are long integers, so the longest
        # digit run is overwhelmingly likely to be the PO rather than a
        # date fragment (which is split by '-' into <=4-digit pieces).
        stem = os.path.splitext(name)[0]
        digit_runs = re.findall(r'\d+', stem)
        return max(digit_runs, key=len) if digit_runs else ''

    def _load_consignment_location_lookup(
        self,
        path: str,
        config: Dict[str, Any],
        result: ProcessingResult,
    ) -> Dict[str, Dict[str, str]]:
        """
        Build a ``{PO -> {wh, po_date, exp_date}}`` lookup from Flipkart's
        Consignment Visibility Report (v2.3.1; dates added v2.4.0).

        The raw consignment CSVs carry no Location, so when the operator
        supplies the visibility report we recover each PO's destination
        warehouse from it. The report has one row per (PO, Super
        Category) but the Warehouse Id is identical across a PO's rows,
        so we collapse to one entry per PO.

        The returned value is the RAW ``Warehouse Id`` machine code
        (e.g. ``malur_bts``). That string becomes the row's Location and
        is resolved to a Transfer-to Code by the normal Ship-To B2B
        mapping lookup — which is why the operator must add these machine
        codes as alias rows in Ship-To B2B (the chosen bridge). A PO that
        maps here but whose Warehouse Id isn't in Ship-To B2B still gets
        the usual "location not found" warning, pointing them at the
        missing alias.

        Failures are soft: an unreadable report or missing columns logs a
        warning and returns ``{}`` (every Location falls back to empty),
        never aborting the run.

        Args:
            path:   Path to the visibility report CSV.
            config: Flipkart-TO config (column names read from
                    ``consignment_mode``).
            result: ProcessingResult for warning accumulation.

        Returns:
            ``{po_str: warehouse_id_str}`` — possibly empty.
        """
        cmode = config.get('consignment_mode', {})
        po_col = cmode.get('visibility_po_col', 'Consignment Id')
        wh_col = cmode.get('visibility_loc_col', 'Warehouse Id')
        # v2.4.0: the report also carries the dates the tracker wants —
        # 'Creation Date' (→ PO Date) and 'Scheduled Pick Up Date'
        # (→ Exp Date). Optional: a report missing them just leaves those
        # tracker cells blank, exactly as before.
        po_date_col = cmode.get('visibility_po_date_col', 'Creation Date')
        exp_date_col = cmode.get('visibility_exp_date_col',
                                 'Scheduled Pick Up Date')

        try:
            df = pd.read_csv(path)
        except Exception as e:  # noqa: BLE001
            result.warnings.append((
                '', '',
                f"Could not read the location report "
                f"({os.path.basename(path)}): {e}. Locations left empty."
            ))
            return {}

        missing = [c for c in (po_col, wh_col) if c not in df.columns]
        if missing:
            result.warnings.append((
                '', '',
                f"Location report is missing column(s) {missing} "
                f"(expected '{po_col}' and '{wh_col}'). Locations left "
                f"empty."
            ))
            return {}

        have_po_date = po_date_col in df.columns
        have_exp_date = exp_date_col in df.columns

        # v2.4.0: each entry is {wh, po_date, exp_date}. wh drives Location
        # (resolved to a Transfer-to Code downstream); the two dates are
        # injected as __po_date__/__exp_date__ so build_tracker_rows fills
        # the PO Date / Exp Date columns.
        lookup: Dict[str, Dict[str, str]] = {}
        conflicts: Set[str] = set()
        for _, row in df.iterrows():
            po = self._coerce_po_to_str(row[po_col])
            wh_raw = row[wh_col]
            if pd.isna(wh_raw):
                continue
            wh = str(wh_raw).strip()
            # Skip blanks and Flipkart's literal 'N/A' placeholders.
            if not po or wh == '' or wh.upper() == 'N/A':
                continue
            if (po in lookup and lookup[po]['wh'] != wh
                    and po not in conflicts):
                conflicts.add(po)
                result.warnings.append((
                    po, '',
                    f"Location report lists PO {po} under multiple "
                    f"Warehouse Ids ('{lookup[po]['wh']}' and '{wh}'). Using "
                    f"the first ('{lookup[po]['wh']}')."
                ))
                continue
            if po in lookup:
                continue
            po_date = ''
            if have_po_date and not pd.isna(row[po_date_col]):
                po_date = str(row[po_date_col]).strip()
            exp_date = ''
            if have_exp_date and not pd.isna(row[exp_date_col]):
                exp_date = str(row[exp_date_col]).strip()
            lookup[po] = {'wh': wh, 'po_date': po_date, 'exp_date': exp_date}

        if not have_po_date or not have_exp_date:
            miss = []
            if not have_po_date:
                miss.append(f"'{po_date_col}' (PO Date)")
            if not have_exp_date:
                miss.append(f"'{exp_date_col}' (Exp Date)")
            result.warnings.append((
                '', '',
                f"Visibility report has no {', '.join(miss)} column — "
                f"those tracker date cells will be blank."
            ))

        logging.info(
            "consignment location report: %d PO->Warehouse mappings from %s",
            len(lookup), os.path.basename(path),
        )
        return lookup

    def process_consignments(
        self,
        filepaths: List[str],
        config: Dict[str, Any],
        margin_pct: float = 0.70,
        visibility_report_path: Optional[str] = None,
    ) -> ProcessingResult:
        """
        Flipkart-TO BULK CONSIGNMENT mode — build the consolidated dump
        from many raw consignment CSVs, then run the standard TO pipeline.

        Background
        ----------
        Historically the operator hand-consolidated Flipkart's per-PO
        consignment exports into ONE 7-column dump (``Po Number |
        Location | FSN | SKU Id | Product Name | Cost Price | Quantity
        Sent``) and fed that single file in via :meth:`process`. That
        path still exists unchanged — see the GUI's "Consolidated dump"
        mode.

        This method automates the manual consolidation. The operator
        instead hands us the raw exports directly
        (``Consignment_Details_<PO>_<date>.csv``, one per PO) and we
        assemble the dump in memory:

          * **PO number** comes from each file's NAME (the CSV has no PO
            column) via :meth:`_extract_po_from_filename`.
          * **Location** comes from the optional Consignment Visibility
            Report (``visibility_report_path``) — a Flipkart export that
            lists ``Consignment Id`` (= PO) and ``Warehouse Id`` (the
            destination warehouse machine code, e.g. ``malur_bts``). We
            join on PO and write the RAW Warehouse Id as the row's
            Location, so the Summary's 'Location (Raw)' shows exactly what
            came in. The machine-code → friendly-name aliasing and
            Transfer-to Code resolution happen later in :meth:`_process_to`
            (via the PROVISIONAL ``warehouse_aliases`` config map), which
            records the deciphered friendly name as 'Location (Mapped)'
            and flags it as a fuzzy/provisional match. When no report is
            supplied or a PO isn't in it, Location is EMPTY (blank
            Transfer-to Code + a warning).
          * Every other column the dump needs (``SKU Id`` = EAN,
            ``Quantity Sent``, ``Cost Price``, ``Product Name``, ``FSN``)
            is ALREADY present in the consignment CSV, so it passes
            straight through — no per-column mapping required.

        The assembled DataFrame is the exact shape the consolidated dump
        has, so it flows through :meth:`_process_loaded_dataframe` →
        :meth:`_process_to` identically: rows aggregate by (PO, Item No),
        Item No / MRP / GST come from Items_March via the EAN, and the
        engine computes Transfer Price itself.

        Per-file failures are isolated (mirrors :meth:`process_multi`):
        an unreadable CSV or a filename with no recoverable PO is skipped
        with a ``[<filename>]``-prefixed warning; the rest of the batch
        still produces rows.

        Args:
            filepaths:  Consignment CSV paths. Must not be empty.
            config:     The Flipkart-TO config (``output_type='to'``).
            margin_pct: Decimal margin (0.60 for Flipkart-TO).

        Returns:
            A combined ``ProcessingResult`` whose ``raw_df`` is the
            assembled dump and whose ``rows`` are the aggregated
            Transfer Lines across every consignment file.
        """
        if not filepaths:
            empty = ProcessingResult(marketplace=config['party_name'])
            empty.warnings.append((
                '', '',
                "process_consignments called with empty file list — "
                "nothing to process."
            ))
            return empty

        cmode = config.get('consignment_mode', {})
        po_col = config['po_col']     # 'Po Number' — we synthesize it
        loc_col = config['loc_col']   # 'Location'  — we leave it blank
        po_regex = cmode.get(
            'filename_po_regex', r'Consignment_Details_([^_]+)_',
        )

        result = ProcessingResult(
            marketplace=config['party_name'],
            input_file=os.path.basename(filepaths[0]),
            input_file_path=filepaths[0],
            margin_pct=margin_pct,
            compare_basis=config.get('compare_basis', 'cost'),
            compare_label=config.get('compare_label', 'Price'),
            input_files_count=len(filepaths),
        )
        # v2.4.3: full exception registry for the Exceptions sheet (TO mode).
        if self.master is not None:
            result.exception_registry = getattr(
                self.master, 'exception_registry', []) or []

        # v2.3.1: build the PO -> Warehouse Id lookup from the optional
        # visibility report. Empty dict when no report is supplied, in
        # which case every Location stays blank (original behaviour).
        loc_lookup: Dict[str, Dict[str, str]] = {}
        if visibility_report_path:
            loc_lookup = self._load_consignment_location_lookup(
                visibility_report_path, config, result,
            )

        # v2.4.0 (Meesho): Location comes from the FILENAME. Build a {city
        # token → Ship-To B2B Del Location name} map from the loaded mapping
        # by taking the suffix of each Transfer-to Code (MS_BLR → 'BLR'). The
        # filename is scanned for any token; the matched Del Location name is
        # injected as the row's Location so the normal Ship-To B2B resolution
        # in _process_to produces the right Transfer-to Code.
        cmode = config.get('consignment_mode', {})
        token_map: Dict[str, str] = {}
        if cmode.get('filename_loc_from_shipto'):
            for info in self.mapping.mappings.values():
                code = str(info.get('ship_to', '') or '')
                tok = code.split('_')[-1].strip().upper() if '_' in code else ''
                if tok:
                    # token (BLR) → the short Transfer-to Code (MS_BLR), which
                    # becomes the row's Location and resolves to itself via the
                    # mapping's by-ship-to-code index — short and self-evident.
                    token_map.setdefault(tok, code)

        # v2.4.0 (Meesho): synthetic dates — PO Date = today, Exp Date =
        # today + N days. Files carry no dates; the tracker still wants them.
        synth_po_date = ''
        synth_exp_date = ''
        if cmode.get('po_date_today') or cmode.get('exp_date_offset_days') is not None:
            from datetime import date, timedelta
            _today = date.today()
            if cmode.get('po_date_today'):
                synth_po_date = _today.isoformat()
            _off = cmode.get('exp_date_offset_days')
            if _off is not None:
                synth_exp_date = (_today + timedelta(days=int(_off))).isoformat()

        logging.info("process_consignments: assembling dump from %d files",
                     len(filepaths))

        per_file_dfs: List[pd.DataFrame] = []
        for fp in filepaths:
            basename = os.path.basename(fp)

            po = self._extract_po_from_filename(fp, po_regex)
            if not po:
                result.warnings.append((
                    '', '',
                    f"[{basename}] Could not extract a PO number from the "
                    f"filename (expected 'Consignment_Details_<PO>_<date>"
                    f".csv'). File skipped."
                ))
                continue

            try:
                df = pd.read_csv(fp, header=config.get('header_row', 0))
            except Exception as e:  # noqa: BLE001 — surface ANY CSV-read error
                result.warnings.append((
                    po, '', f"[{basename}] Cannot read CSV: {e}"
                ))
                continue

            if df.empty:
                result.warnings.append((
                    po, '', f"[{basename}] No data rows — skipped."
                ))
                continue

            # Inject the two columns the consolidated dump carries but the
            # raw consignment CSV lacks. Everything else (SKU Id /
            # Quantity Sent / Cost Price / Product Name / FSN) is already
            # in the file under the names the Flipkart-TO config expects.
            #
            # Location = the PO's RAW Warehouse Id from the visibility
            # report (machine code, e.g. 'malur_bts'). Kept RAW on purpose
            # so the Summary's 'Location (Raw)' shows exactly what came in;
            # the machine-code → friendly-name aliasing + Transfer-to Code
            # resolution happens in _process_to, which records the
            # deciphered friendly name as 'Location (Mapped)'. Blank when
            # no report was supplied or the PO is absent from it (warn so
            # the operator knows that PO's Transfer-to Code will be empty).
            vis = loc_lookup.get(po) or {}
            wh_id = vis.get('wh', '')
            if loc_lookup and not wh_id:
                result.warnings.append((
                    po, '',
                    f"[{basename}] PO {po} not found in the location "
                    f"report — Location left empty (Transfer-to Code will "
                    f"be blank)."
                ))

            # v2.4.0 (Meesho): derive Location from a city token in the
            # filename (e.g. '…-blr.csv' → MS_BLR's Del Location).
            if token_map and not wh_id:
                base_up = basename.upper()
                hit = None
                for tok in sorted(token_map, key=len, reverse=True):
                    if re.search(r'(?<![A-Z0-9])' + re.escape(tok)
                                 + r'(?![A-Z0-9])', base_up):
                        hit = token_map[tok]
                        break
                if hit:
                    wh_id = hit
                else:
                    result.warnings.append((
                        po, '',
                        f"[{basename}] No ship-to city token "
                        f"({', '.join(sorted(token_map))}) found in the "
                        f"filename — Location left empty (Transfer-to Code "
                        f"will be blank). Rename the file to include the "
                        f"city, e.g. 'order-line-items-{po}-blr.csv'."))

            df[po_col] = po          # PO recovered from the filename
            df[loc_col] = wh_id      # RAW Warehouse Id (or '' if unknown)
            # v2.4.0: inject the dates so the tracker's PO Date / Exp Date
            # fill in. From the visibility report when present (Flipkart-TO),
            # else the synthetic today / today+N (Meesho). __po_date__ /
            # __exp_date__ are the first date candidates build_tracker_rows
            # looks for.
            df['__po_date__'] = vis.get('po_date', '') or synth_po_date
            df['__exp_date__'] = vis.get('exp_date', '') or synth_exp_date
            # Tag the source file so Raw Data can distinguish rows.
            df['__source_file__'] = basename
            per_file_dfs.append(df)
            logging.info("process_consignments: %s -> PO %s, loc %r (%d rows)",
                         basename, po, wh_id, len(df))

        if not per_file_dfs:
            result.warnings.append((
                '', '',
                "No readable consignment files produced any rows — check "
                "the filenames match 'Consignment_Details_<PO>_<date>.csv' "
                "and that the CSVs aren't empty."
            ))
            result.raw_df = pd.DataFrame()
            return result

        # Concatenate into the in-memory "manual dump". ``sort=False``
        # preserves the first file's column order; differing columns in
        # later files just become NaN (harmless — the pipeline only reads
        # the configured columns).
        combined_df = pd.concat(
            per_file_dfs, ignore_index=True, sort=False,
        )
        result.raw_df = combined_df
        logging.info(
            "process_consignments: dump assembled — %d rows from %d/%d "
            "files",
            len(combined_df), len(per_file_dfs), len(filepaths),
        )

        # Run the assembled dump through the same pipeline a single
        # consolidated file would take.
        return self._process_loaded_dataframe(
            combined_df, config, margin_pct, result,
        )

    # ── Column validation helpers ──────────────────────────────────────

    @staticmethod
    def _resolve_column_aliases(
        config: Dict[str, Any],
        df_columns,
        result: ProcessingResult,
    ) -> Optional[Dict[str, Any]]:
        """
        Normalize every ``*_col`` config key against actual DataFrame
        headers. Handles three forms of mismatch:

        1. **List alias** — the config value is a list of candidate
           header names, e.g. ``'po_col': ['PO', 'PO Number']``. We
           pick the first entry that matches. List order IS preference
           order.

        2. **Case / whitespace drift** (v1.8.1, opt-in via
           ``config['case_insensitive_cols'] = True``) — the config
           value is a plain string, but the actual header in the file
           differs only in case or surrounding/internal whitespace.
           E.g. config says ``'HSN'`` but the file has ``'Hsn'``, or
           config says ``'PO Number'`` but the file has
           ``'Po  Number'`` (double space). The resolver finds these
           via a lowercase + whitespace-collapsed match and substitutes
           the file's actual header string into the config (because
           downstream pandas indexing needs exact match).

           Rationale: marketplace dashboards occasionally reformat their
           exports — Myntra shipped three casings of its PO header in
           two weeks; Reliance shipped ``HSN`` vs ``Hsn`` across
           batches. Without this flag, every drift forces a code
           update. With it, the engine absorbs drift automatically.
           The flag is opt-in per marketplace so stable-header
           marketplaces (Blink/RK/Zepto) still fail LOUDLY on a real
           mistake rather than silently matching something unintended.

        3. **Not found at all** — emits a warning. For required
           columns (po/loc/qty/ean/item) this aborts the run; for
           optional columns the key is set to None and the pipeline
           continues.

        The returned config is a shallow copy ready for the rest of the
        pipeline (which expects scalar, exact-match column names).

        Args:
            config:     Original marketplace config dict.
            df_columns: Pandas columns of the loaded punch file
                        (``df.columns``).
            result:     ProcessingResult for appending warnings about
                        unresolvable columns.

        Returns:
            New config dict with all ``*_col`` entries normalized, or
            ``None`` if a required column can't be resolved (caller
            should return the result immediately; warning already
            appended).
        """
        resolved = dict(config)
        available_list = list(df_columns)
        available_set = set(available_list)
        case_insensitive = bool(config.get('case_insensitive_cols'))

        # v1.8.1: build a lowercase+whitespace-normalized lookup so we
        # can find headers by "semantic" equality. Example entries:
        #   'hsn' -> 'HSN'
        #   'po number' -> 'PO Number'
        # We canonicalize by lowercasing, stripping edges, and
        # collapsing internal multi-space runs.
        def _normalize(s: Any) -> str:
            if s is None:
                return ''
            return ' '.join(str(s).split()).lower()

        lower_lookup: Dict[str, str] = {}
        if case_insensitive:
            for actual in available_list:
                lower_lookup.setdefault(_normalize(actual), str(actual))

        def _find(name: str) -> Optional[str]:
            """Return the file's actual column for ``name``, or None.

            Exact match first (always); case-insensitive fallback
            only when the marketplace opts in.
            """
            if name in available_set:
                return name
            if case_insensitive:
                return lower_lookup.get(_normalize(name))
            return None

        required_keys = {'po_col', 'loc_col', 'qty_col',
                         'ean_col', 'item_col'}

        for key, value in list(config.items()):
            if not key.endswith('_col'):
                continue

            # ── List alias path ────────────────────────────────────────
            if isinstance(value, list):
                chosen: Optional[str] = None
                for candidate in value:
                    hit = _find(candidate)
                    if hit is not None:
                        chosen = hit
                        break

                if chosen is not None:
                    resolved[key] = chosen
                    logging.info(
                        "Column alias: %s = %r (from options %r)",
                        key, chosen, value,
                    )
                else:
                    if key in required_keys:
                        result.warnings.append((
                            '', '',
                            f"Required column '{key}' not found — tried "
                            f"{value!r}, but none exist in the punch "
                            f"file. Available columns: "
                            f"{available_list[:15]}..."
                        ))
                        return None
                    resolved[key] = None
                    logging.info(
                        "Column alias: %s = None (none of %r found)",
                        key, value,
                    )
                continue

            # ── Scalar path: apply case-insensitive lookup if opted in ─
            if isinstance(value, str) and value:
                hit = _find(value)
                if hit is not None and hit != value:
                    # Found it, but under a different casing/spacing.
                    # Substitute the file's actual header so downstream
                    # pandas indexing works.
                    resolved[key] = hit
                    logging.info(
                        "Column case-fold: %s = %r (config said %r)",
                        key, hit, value,
                    )
                # If hit == value, nothing to do. If hit is None, let
                # the validator complain downstream — this keeps
                # behavior identical to pre-v1.8.1 for marketplaces
                # without case_insensitive_cols, and gives a specific
                # "required column missing" message via
                # _validate_required_columns for those with it.

        return resolved

    def _load_pdf(
        self,
        filepath: str,
        config: Dict[str, Any],
        result: ProcessingResult,
        parser_key: Optional[str] = None,
    ) -> Optional[pd.DataFrame]:
        """
        Load a PDF marketplace file via the configured PDF parser.

        Dispatches on the config's ``pdf_parser`` key (e.g. ``'avenue'``)
        to look up a callable in ``PDF_PARSERS``. Each parser is
        responsible for returning a DataFrame with the same column
        shape an Excel read would produce — so the engine's downstream
        column-resolution / mapping / validation code runs without
        knowing whether the source was Excel or PDF.

        Args:
            filepath:   Path to the marketplace PDF file.
            config:     Entry from ``MARKETPLACE_CONFIGS`` (the active
                        marketplace's config dict).
            result:     ProcessingResult being built — warnings are
                        appended to ``result.warnings`` on failure.

        Returns:
            The parsed DataFrame on success, or ``None`` on any failure
            (file read error, no parser registered, parser raised). The
            caller treats ``None`` as a signal to abort and surface the
            warnings.
        """
        parser_key = (parser_key or config.get('pdf_parser')
                      or config.get('file_parser'))
        if not parser_key:
            result.warnings.append((
                '', '',
                f"Marketplace config requires a parser but none "
                f"('pdf_parser'/'file_parser') was provided — cannot load "
                f"{os.path.basename(filepath)}."
            ))
            return None

        parser = PDF_PARSERS.get(parser_key)
        if parser is None:
            result.warnings.append((
                '', '',
                f"No PDF parser registered under {parser_key!r} — "
                f"available parsers: {sorted(PDF_PARSERS.keys())}"
            ))
            return None

        try:
            df = parser(filepath)
        except Exception as e:  # noqa: BLE001 — any parser failure → warning
            logging.exception("PDF parser %r failed on %s", parser_key, filepath)
            result.warnings.append((
                '', '',
                f"PDF parse failed ({parser_key} parser): {e}"
            ))
            return None

        if df is None or df.empty:
            result.warnings.append((
                '', '',
                f"PDF parser {parser_key!r} returned no rows from "
                f"{os.path.basename(filepath)} — file may be empty, "
                f"corrupted, or in an unexpected format."
            ))
            return None

        return df

    def _resolve_source_sheet(
        self,
        target: str,
        available: List[str],
        result: ProcessingResult,
    ) -> Optional[str]:
        """
        Pick the right sheet to read based on the config's ``source_sheet``.

        Supports two match modes:
            * **Exact match** — ``target`` equals a sheet name.
              Used by most marketplaces (``'Sheet1'`` default, ``'PO'``
              for Reliance).
            * **Wildcard prefix match** — ``target`` ends with ``'*'``.
              Strips the ``*`` and finds any sheet whose name starts
              with the remaining prefix. Used by Zepto because its
              data sheet is named ``'PO_<random-hex>'`` which changes
              every dump (e.g. ``PO_64863340b23e6c90``).

        Behavior on miss:
            * **Exact miss** — falls back to the first available
              sheet, emits a warning so the user sees what happened.
              This has historically been kind to users whose files
              have unexpected sheet ordering (e.g. a user-added
              pivot sheet sitting before 'Sheet1').
            * **Wildcard miss** — aborts with a clear error. The
              wildcard implies a specific marketplace's data format;
              falling back silently would produce nonsense output
              by running the engine against whatever sheet happens
              to come first.

        Args:
            target:     Value of ``config['source_sheet']``.
            available:  List of sheet names in the workbook.
            result:     For appending warnings/errors.

        Returns:
            Chosen sheet name, or ``None`` to abort processing
            (wildcard miss only).
        """
        # ── Wildcard mode ───────────────────────────────────────────────
        if target.endswith('*'):
            prefix = target[:-1]
            matches = [s for s in available if s.startswith(prefix)]

            if len(matches) == 1:
                return matches[0]

            if not matches:
                result.warnings.append((
                    '', '',
                    f"No sheet starting with {prefix!r} found in file. "
                    f"Available sheets: {available}. "
                    f"This marketplace requires its data sheet — check "
                    f"that the upload is a complete, untouched dump."
                ))
                return None

            # Multiple matches: take the first but warn.
            chosen = matches[0]
            result.warnings.append((
                '', '',
                f"Multiple sheets match {target!r}: {matches}. Using "
                f"{chosen!r} (first match). Verify this is the correct "
                f"data sheet."
            ))
            return chosen

        # ── Exact mode (original behavior) ──────────────────────────────
        if target in available:
            return target

        sheet_to_read = available[0]
        result.warnings.append((
            '', '',
            f"'{target}' not found in file — falling back to "
            f"'{sheet_to_read}'. Available sheets: {available}"
        ))
        logging.warning("'%s' missing; reading '%s' instead",
                         target, sheet_to_read)
        return sheet_to_read

    def _validate_required_columns(self, df: pd.DataFrame,
                                    config: Dict[str, Any],
                                    result: ProcessingResult) -> bool:
        """
        Confirm the required columns exist for this marketplace.

        Required set depends on ``item_resolution``:
          * ``from_column`` → po, loc, item, qty
          * ``from_ean``    → po, loc, ean, qty   (item_col may be absent)

        On failure, appends a warning to ``result`` and returns False.
        Caller should ``return result`` immediately.
        """
        item_resolution = config.get('item_resolution', 'from_column')

        required_cols: Dict[str, str] = {
            'po': config['po_col'],
            'loc': config['loc_col'],
            'qty': config['qty_col'],
        }

        if item_resolution == 'from_swiggy_sku':
            sku_required = config.get('sku_col')
            if not sku_required:
                result.warnings.append((
                    '', '',
                    "Config error: item_resolution='from_swiggy_sku' requires "
                    "sku_col."))
                return False
            required_cols['sku'] = sku_required
        elif item_resolution == 'from_ean':
            ean_required = config.get('ean_col')
            if not ean_required:
                result.warnings.append((
                    '', '',
                    "Config error: item_resolution='from_ean' requires ean_col."))
                return False
            required_cols['ean'] = ean_required
        else:  # 'from_column' (default)
            item_required = config.get('item_col')
            if not item_required:
                result.warnings.append((
                    '', '',
                    "Config error: item_resolution='from_column' requires item_col."))
                return False
            required_cols['item'] = item_required

        for _key, col_name in required_cols.items():
            if col_name not in df.columns:
                result.warnings.append((
                    '', '',
                    f"Required column '{col_name}' not found. "
                    f"Available: {list(df.columns)[:15]}..."))
                return False

        return True

    @staticmethod
    def _validate_optional_column(
        df: pd.DataFrame,
        config: Dict[str, Any],
        config_key: str,
        log_warn_to_result: Optional[ProcessingResult] = None,
    ) -> Optional[str]:
        """
        Check if an optional column exists. Returns its name if present,
        ``None`` otherwise. If absent and ``log_warn_to_result`` is given,
        appends a warning row.
        """
        col_name = config.get(config_key)
        if col_name and col_name in df.columns:
            return col_name

        if col_name:
            # Configured but missing — log it
            logging.warning("%s column '%s' not found — skipping",
                            config_key, col_name)
            if log_warn_to_result is not None:
                log_warn_to_result.warnings.append((
                    '', '',
                    f"Column '{col_name}' (config key '{config_key}') not "
                    f"found in file — that feature will be skipped. "
                    f"Available: {list(df.columns)[:10]}..."))
        return None

    # ── TO output pipeline (v2.0.0) ────────────────────────────────────

    def _process_to(self, df: 'pd.DataFrame', config: Dict[str, Any],
                     margin_pct: float,
                     result: ProcessingResult) -> ProcessingResult:
        """
        Build SORows for the Transfer Order pipeline.

        Called from :meth:`process` when the marketplace config sets
        ``output_type='to'`` (currently only Flipkart-TO).

        v2.1.4 redesign â switched from in-band item/MRP/GST resolution
        to master-lookup via EAN. The Flipkart Branch dump shrank from
        13 self-contained columns to 7 (the source no longer ships
        Item No / MRP / GST), so this method now mirrors Blink's
        ``from_ean`` resolution path.

        Differences from the SO row-builder (:meth:`_process_row`):

        1. **Aggregation.** Multiple rows in the same PO with the
           same Item No are collapsed into a single ``SORow`` whose
           ``qty`` is the sum. The Flipkart Branch dump has separate
           rows per FSN, but D365 only cares about Item No. The
           aggregation key is ``(po, item_no)`` exactly.

        2. **Reference-only price comparison** (``compare_mode='reference_only'``
           in config). The engine reads ``fob_col`` per row (Flipkart's
           stated Cost Price) and computes ``diffn`` against the
           engine's calculated value, but never marks the row as
           MISMATCH. Rows where ``|diffn| > 0.01`` get a one-line
           audit warning written to the Warnings sheet. Transfer
           Price written to D365 always uses the engine's calculated
           value regardless of diff.

        3. **No customer.** ``cust_no`` stays empty because TOs have
           no customer; ``ship_to`` holds the Transfer-to Code (e.g.
           'FK_BHW_BTS') from Ship-To B2B.

        Master miss handling (``NOT_IN_MASTER``) mirrors Blink:
        the row IS emitted with ``item_no = f'?EAN:{ean}'`` and
        ``validation_status = 'NOT_IN_MASTER'`` so the operator
        can spot the gap in Items_March on the Validation sheet.
        Transfer Price stays ``None`` for those rows (D365 will
        default-fill from the vendor master).

        Args:
            df:         The Flipkart dump's data frame after column
                         alias resolution + required-column validation
                         (both already done by the caller).
            config:     Resolved marketplace config dict.
            margin_pct: Decimal margin (0.60 for Flipkart-TO).
            result:     ProcessingResult to populate. Already has
                         marketplace metadata + ``output_type='to'``
                         set by the caller.

        Returns:
            The same ``result`` instance with ``rows`` populated and
            any warnings appended.
        """
        po_col = config['po_col']
        loc_col = config['loc_col']
        qty_col = config['qty_col']
        ean_col = config['ean_col']
        fob_col = config.get('fob_col')         # optional reference comparison
        compare_mode = config.get('compare_mode')  # e.g. 'reference_only'
        party_name = config['party_name']

        # v2.3.1: optional Warehouse-Id → friendly-name alias map (e.g.
        # Flipkart-TO consignments, where the visibility report's Location
        # is a machine code like 'malur_bts'). When a row's Location
        # matches an alias key we look the mapping up under the friendly
        # name so the Transfer-to Code resolves, BUT keep the raw machine
        # code as the row's Location — so the Summary shows both 'what
        # came in' (Raw) and 'what we deciphered' (Mapped). Keys are
        # lowercased so report casing doesn't matter. Aliased rows are
        # flagged PROVISIONAL (deduped) — a temporary fuzzy bridge until
        # exact ship-to codes are wired in. Empty for marketplaces with no
        # alias map (Meesho-TO) or the consolidated dump (friendly names
        # already), so this is a no-op there.
        loc_aliases = {
            str(k).strip().lower(): v
            for k, v in config.get('consignment_mode', {})
                              .get('warehouse_aliases', {}).items()
        }
        provisional_aliases: Dict[str, str] = {}

        # v2.1.4: master is required now. Pre-v2.1.4 TO mode bypassed
        # master entirely; we now need it to resolve item_no / mrp /
        # gst_code via the EAN. Fail loudly if it isn't loaded so the
        # operator gets a clear hint instead of silent NOT_IN_MASTER
        # on every row.
        if not self.master:
            result.warnings.append(('', '',
                "TO mode now requires Items_March master to be loaded "
                "(v2.1.4 — Item No / MRP / GST come from master via "
                f"EAN in '{ean_col}'). No master loaded → no rows "
                "produced. Load the Items Master and re-run."))
            return result

        # Aggregator: (po_str, item_no) → accumulated SORow draft.
        aggregator: Dict[Tuple[str, Any], Dict[str, Any]] = {}

        # Warning de-dupe sets.
        warned_locations: Set[Tuple[str, str]] = set()
        skipped_no_ean = 0
        skipped_no_qty = 0
        ref_warning_count = 0

        # Tolerance for reference-only diff warnings. Below this we
        # don't bother emitting an audit line — sub-paisa rounding
        # noise from MRP × margin / (1 + GST) isn't actionable.
        REF_TOLERANCE = 0.01

        for _, row in df.iterrows():
            po_raw = row[po_col]
            loc_raw = row[loc_col]
            qty_raw = row[qty_col]
            ean_raw = row[ean_col]
            fob_raw = row[fob_col] if fob_col else None

            # Skip totally blank rows (pandas trailing NaN rows).
            if pd.isna(po_raw) and pd.isna(loc_raw) and pd.isna(qty_raw):
                continue

            po_str = self._coerce_po_to_str(po_raw)
            loc_str = '' if pd.isna(loc_raw) else str(loc_raw).strip()

            # ── EAN — required ──────────────────────────────
            if pd.isna(ean_raw) or str(ean_raw).strip() == '':
                # v2.3.1 thumb rule: log every dropped line with identity,
                # not just the roll-up count below.
                skipped_no_ean += 1
                result.warnings.append((
                    po_str, loc_str,
                    f"Row {row.name}: missing {ean_col} (EAN) — NOT written "
                    f"to the Transfer Order (PO {po_str})."))
                continue
            # EANs come in as int64 from openpyxl. str(int(...)) handles
            # the common case; fall back to bare str() for non-numeric
            # EAN-like values.
            try:
                ean = str(int(float(ean_raw)))
            except (ValueError, TypeError):
                ean = str(ean_raw).strip()

            # ── Qty — required, must be > 0 ────────────────────────
            try:
                qty = int(float(qty_raw)) if not pd.isna(qty_raw) else 0
            except (ValueError, TypeError):
                qty = 0
            if qty <= 0:
                skipped_no_qty += 1
                result.warnings.append((
                    po_str, loc_str,
                    f"Row {row.name}: qty is {qty} (raw={qty_raw!r}) — NOT "
                    f"written to the Transfer Order (PO {po_str}, EAN "
                    f"{ean})."))
                continue

            # ── Master lookup via EAN ──────────────────────────
            master_info = self.master.lookup(ean)

            if master_info is None:
                # NOT_IN_MASTER — emit row with placeholder item_no so
                # the operator can spot the gap on the Validation
                # sheet. No Transfer Price (calc_price=None) — D365
                # will default-fill from vendor master if possible.
                item_no: Any = f'?EAN:{ean}'
                mrp: Optional[float] = None
                gst_code = ''
                description = ''
                calc_price: Optional[float] = None
                validation_status = 'NOT_IN_MASTER'
            else:
                resolved_item = master_info.get('item_no', '')
                try:
                    item_no = int(resolved_item)
                except (ValueError, TypeError):
                    item_no = str(resolved_item).strip()

                mrp_master = master_info.get('mrp')
                gst_master = master_info.get('gst_code', '') or ''
                try:
                    mrp = float(mrp_master) if mrp_master is not None else None
                except (ValueError, TypeError):
                    mrp = None
                gst_code = str(gst_master).strip()
                description = str(master_info.get('description', '') or '')

                # Compute Transfer Price from master MRP + GST.
                if mrp is not None and gst_code:
                    calc_price = MasterLoader.calc_cost_price(
                        mrp, gst_code, margin_pct,
                    )
                else:
                    calc_price = None

                validation_status = 'OK'

            # ── Read Flipkart's stated cost (reference) ──────────────────
            fob_price: Optional[float] = None
            if fob_raw is not None and not pd.isna(fob_raw):
                try:
                    fob_price = float(fob_raw)
                except (ValueError, TypeError):
                    fob_price = None

            # ── Reference-only comparison (v2.1.4) ───────────────────────
            # When compare_mode is 'reference_only', we compute diffn
            # but DON'T promote OK → MISMATCH. Diffs above tolerance
            # still get a one-line audit warning so the operator can
            # spot pricing drift.
            #
            # Dedup: the same SKU often appears in many POs at the same
            # diff (Flipkart's pricing tier for that EAN drifted from
            # ours). One warning per (item_no, rounded diff) keeps the
            # Warnings sheet scannable. The total count is still
            # reported in the roll-up at the bottom.
            diffn: Optional[float] = None
            if (calc_price is not None and fob_price is not None):
                diffn = round(fob_price - calc_price, 4)
                if (compare_mode == 'reference_only'
                        and abs(diffn) > REF_TOLERANCE):
                    ref_warning_count += 1
                    warn_key = ('ref_diff', str(item_no), round(diffn, 2))
                    if warn_key not in warned_locations:
                        warned_locations.add(warn_key)
                        result.warnings.append((po_str, loc_str,
                            f"Cost diff (reference only): item {item_no} "
                            f"EAN {ean} — Flipkart stated ₹{fob_price:.2f}, "
                            f"engine calculated ₹{calc_price:.2f} "
                            f"(diff ₹{diffn:+.2f}). Engine value used "
                            f"in Transfer Price."))

            # ── Mapping lookup (Ship-To B2B) ────────────────────────
            # v2.3.1: when the Location is a known machine-code alias,
            # resolve the mapping under its friendly name but KEEP the raw
            # machine code as ``loc_str`` (the row's Location). The matched
            # friendly name comes back as ``mapped_loc`` so Raw and Mapped
            # differ on the Summary (auto-highlighted as a fuzzy match).
            lookup_loc = loc_str
            if loc_aliases and loc_str:
                alias = loc_aliases.get(loc_str.strip().lower())
                if alias:
                    lookup_loc = alias
            cust_no, ship_to, mapped, mapped_loc = self._resolve_mapping(
                lookup_loc, po_str, party_name, warned_locations, result,
            )
            # Flag the provisional alias (deduped) once it actually maps.
            if lookup_loc != loc_str and mapped and loc_str not in provisional_aliases:
                provisional_aliases[loc_str] = mapped_loc or lookup_loc
                result.warnings.append((po_str, loc_str,
                    f"PROVISIONAL location alias (verify): '{loc_str}' "
                    f"deciphered to Ship-To B2B location "
                    f"'{mapped_loc or lookup_loc}' → Transfer-to {ship_to}. "
                    f"Fuzzy bridge until exact ship-to codes are wired."))
            if mapped and not ship_to:
                key = ('blank_transfer_to', loc_str)
                if key not in warned_locations:
                    warned_locations.add(key)
                    result.warnings.append((po_str, loc_str,
                        f"Ship-To B2B has '{loc_str}' for Party "
                        f"'{party_name}' but its 'Ship to' (Transfer-to "
                        f"Code) is blank. Transfer Order will export "
                        f"with EMPTY Transfer-to Code — fix in D365 "
                        f"import preview or update Ship-To B2B."))

            # ── Tracker amount (v2.4.0) ─────────────────────────────────
            # TOs carry no ``amount_col``, but the operator wants a per-
            # consignment Total Amount in the tracker. Two figures exist:
            #
            #   OUR amount  = engine Transfer Price grossed back up by GST
            #                 × qty  =  (MRP × margin ÷ (1+GST)) × (1+GST)
            #                 × qty  =  landing (MRP × margin) × qty.
            #                 This is the GST-inclusive value D365 records
            #                 (Total Amount Incl. GST) — what we want in the
            #                 tracker's Order Value per the operator.
            #   VENDOR amount = Flipkart's stated Cost Price × qty. Equals
            #                 the portal's "Amount" figure (verified to the
            #                 paisa). Kept only for the reference log so the
            #                 operator can cross-check against the portal.
            #
            # NOT_IN_MASTER lines have no calc_price → our amount is 0 for
            # them (no MRP/GST to value the transfer); the vendor amount
            # still accrues so the reference total stays complete.
            our_line = (calc_price * MasterLoader.gst_divisor(gst_code) * qty
                        if calc_price is not None else 0.0)
            vendor_line = (fob_price * qty) if fob_price is not None else 0.0

            # ── Aggregate into the (PO, Item No) bucket ──────────────
            agg_key = (po_str, item_no)
            entry = aggregator.get(agg_key)
            if entry is None:
                aggregator[agg_key] = {
                    'po_number': po_str,
                    'location': loc_str,
                    'item_no': item_no,
                    'qty': qty,
                    'cust_no': cust_no,
                    'ship_to': ship_to,
                    'mapped': mapped,
                    'mapped_location': mapped_loc,
                    'ean': ean,
                    'description': description,
                    'mrp': mrp,
                    'gst_code': gst_code,
                    'calc_price': calc_price,
                    'fob_price': fob_price,
                    'diffn': diffn,
                    'validation_status': validation_status,
                    'amount': our_line,
                    'vendor_amount': vendor_line,
                }
            else:
                entry['qty'] += qty
                entry['amount'] += our_line
                entry['vendor_amount'] += vendor_line
                if loc_str != entry['location']:
                    key = ('multi_loc', po_str, item_no)
                    if key not in warned_locations:
                        warned_locations.add(key)
                        result.warnings.append((po_str, loc_str,
                            f"Item No {item_no} appears in multiple "
                            f"locations within PO {po_str}: "
                            f"'{entry['location']}' + '{loc_str}'. "
                            f"Aggregating qty under first-seen "
                            f"location."))

        # ── Emit aggregated SORows ──────────────────────────────────────
        for entry in aggregator.values():
            so_row = SORow(
                po_number=entry['po_number'],
                location=entry['location'],
                item_no=entry['item_no'],
                qty=entry['qty'],
                cust_no=entry['cust_no'],
                ship_to=entry['ship_to'],
                mapped=entry['mapped'],
                mapped_location=entry['mapped_location'],
                ean=entry['ean'],
                description=entry['description'],
                mrp=entry['mrp'],
                gst_code=entry['gst_code'],
                fob_price=entry['fob_price'],
                calc_price=entry['calc_price'],
                cost_price_ref=entry['calc_price'],
                diffn=entry['diffn'],
                validation_status=entry['validation_status'],
                amount=entry['amount'],
            )
            result.rows.append(so_row)

        # ── v2.4.0: per-PO vendor vs our total (reference log) ───────────
        # The tracker's Order Value uses OUR amount (landing × qty, incl
        # GST). Flipkart's portal shows the VENDOR amount (stated Cost
        # Price × qty). They legitimately differ, so we log both per PO —
        # the operator can reconcile the tracker against the portal at a
        # glance without re-opening each consignment.
        po_totals: Dict[str, Dict[str, float]] = {}
        for entry in aggregator.values():
            t = po_totals.setdefault(entry['po_number'], {'our': 0.0, 'vendor': 0.0})
            t['our'] += entry['amount']
            t['vendor'] += entry['vendor_amount']
        for po_str in sorted(po_totals):
            t = po_totals[po_str]
            logging.info(
                "Flipkart-TO PO %s: our total (incl GST) = %.2f, "
                "vendor/portal total = %.2f",
                po_str, t['our'], t['vendor'],
            )
            result.warnings.append((
                po_str, '',
                f"Amount reference — Tracker Order Value (ours, incl GST) = "
                f"₹{t['our']:,.2f}; Flipkart portal 'Amount' (vendor "
                f"Cost Price × qty) = ₹{t['vendor']:,.2f}."))

        # Roll-up warnings.
        if provisional_aliases:
            pairs = ', '.join(
                f"{k}→{v}" for k, v in sorted(provisional_aliases.items())
            )
            result.warnings.append(('', '',
                f"{len(provisional_aliases)} location(s) resolved via "
                f"PROVISIONAL aliases (fuzzy bridge, not yet confirmed "
                f"exact codes): {pairs}. Raw vs Mapped are highlighted on "
                f"the Summary — verify the Transfer-to Codes before "
                f"D365 import."))
        if skipped_no_ean:
            result.warnings.append(('', '',
                f"Skipped {skipped_no_ean} row(s) with missing/blank "
                f"'{ean_col}'. These rows do NOT appear in the output."))
        if skipped_no_qty:
            result.warnings.append(('', '',
                f"Skipped {skipped_no_qty} row(s) with qty=0 or invalid "
                f"qty in '{qty_col}'."))
        if ref_warning_count:
            unique_diffs = sum(
                1 for k in warned_locations
                if isinstance(k, tuple) and k and k[0] == 'ref_diff'
            )
            result.warnings.append(('', '',
                f"Reference-only diffs: {ref_warning_count} row(s) "
                f"had |diff| > \u20b9{REF_TOLERANCE:.2f} between "
                f"Flipkart's stated Cost and engine-calculated values. "
                f"Showing {unique_diffs} unique (item, diff) line(s) "
                f"above — same SKU/diff in multiple POs is collapsed. "
                f"Transfer Price uses engine values; file values are "
                f"on the Validation sheet for reference."))
            logging.info(
                "TO mode (reference_only): %d row-level diffs, "
                "%d unique (item, diff) warnings logged",
                ref_warning_count, unique_diffs,
            )

        logging.info("TO mode: emitted %d aggregated rows across %d PO(s)",
                     len(result.rows),
                     len({r.po_number for r in result.rows}))
        return result

    @staticmethod
    def _coerce_po_to_str(po_raw: Any) -> str:
        """
        Convert a raw PO cell value to a clean string PO number.

        Flipkart dumps store PO numbers as int64 (e.g. ``204345116``).
        Casting via ``str()`` would give ``'204345116'`` which is fine,
        but going through float first (``str(204345116.0)`` →
        ``'204345116.0'``) breaks D365 imports — those trailing
        zeros become part of the document number. Use int coercion
        when the value is numeric to avoid that.
        """
        if pd.isna(po_raw):
            return ''
        if isinstance(po_raw, (int,)) or (
            isinstance(po_raw, float) and po_raw.is_integer()
        ):
            return str(int(po_raw))
        return str(po_raw).strip()

    # ── Per-row processing (SO mode) ────────────────────────────────────

    def _resolve_row_margin(
        self,
        row: pd.Series,
        config: Dict[str, Any],
        run_margin_pct: float,
        item_no: Any,
        po: str,
        result: ProcessingResult,
        warned_keys: Set[Tuple],
    ) -> Tuple[float, Optional[str]]:
        """
        Resolve the per-LINE margin (keep%) for this row from the
        marketplace's optional ``margin_rules`` config, and LOG the
        exception when a non-default rule matched.

        v2.3.1 — built for Nykaa, whose landing/cost differs by product
        category: Perfume/Fragrance keeps 69% of MRP (31% off) while
        Cosmetics keep 66% (34% off). The mechanism is generic so other
        marketplaces can add their own per-line splits purely via config.

        ``config['margin_rules']`` shape::

            {
              'rules': [
                {'label': 'Perfume/Fragrance',
                 'keep_pct': 69,                 # engine margin = 0.69
                 'contains': ['perfume','fragra'],     # substring test...
                 'contains_column': 'SKU Name',        # ...on this column
                 'hsn_prefix': ['3303']},              # HSN cross-check only
              ],
              'default_keep_pct': 66,            # used when no rule matches
              'default_label': 'Cosmetics',
              'flag_hsn_conflicts': True,        # highlight name-vs-HSN clashes
            }

        DECISION — the per-line margin is decided ONLY by ``contains``
        (the description keywords), mirroring the operator's manual
        search. ``hsn_prefix`` does NOT decide the margin; it's a
        cross-check signal.

        CONFLICT HIGHLIGHT — when ``flag_hsn_conflicts`` is set, the row's
        HSN is compared against the rules' ``hsn_prefix`` lists. If the
        category the HSN points to differs from the one the NAME chose
        (e.g. a Body Mist whose name lacks 'perfume'/'fragra' but whose
        HSN is 3303), a ``⚠ AMBIGUOUS`` warning is logged so the operator
        can review it while the exact perfume lookup is being finalised.
        The NAME decision still governs the margin.

        Returns ``(margin_pct_decimal, label)``. With no ``margin_rules``
        the run margin is returned unchanged (other marketplaces are
        unaffected). Each applied non-default rule and each conflict is
        logged once per item on the Warnings sheet (audit trail).
        """
        rules_cfg = config.get('margin_rules')
        if not rules_cfg:
            return run_margin_pct, None
        rules = rules_cfg.get('rules', [])

        # Row HSN (for the cross-check) — same coercion as _check_hsn.
        hsn_col = config.get('hsn_col')
        hsn_val = ''
        if hsn_col and hsn_col in row.index and pd.notna(row[hsn_col]):
            try:
                hsn_val = str(int(float(row[hsn_col])))
            except (ValueError, TypeError):
                hsn_val = str(row[hsn_col]).strip()

        # ── DECISION: description keywords only ─────────────────────────
        chosen_rule = None
        for rule in rules:
            col = rule.get('contains_column')
            needles = rule.get('contains') or []
            if col and needles and col in row.index and pd.notna(row[col]):
                text = str(row[col]).lower()
                if any(str(nd).lower() in text for nd in needles):
                    chosen_rule = rule
                    break

        if chosen_rule is not None:
            keep = float(chosen_rule['keep_pct']) / 100.0
            label = chosen_rule.get('label', 'rule')
            key = ('margin_rule', label, str(item_no))
            if key not in warned_keys:
                warned_keys.add(key)
                default_keep = rules_cfg.get('default_keep_pct')
                result.warnings.append((
                    po, '',
                    f"ℹ Margin rule '{label}' applied to Item "
                    f"{item_no} (matched by name): keep "
                    f"{chosen_rule['keep_pct']}% of MRP"
                    + (f" (vs default {default_keep}%)"
                       if default_keep is not None else "")
                    + "."
                ))
        else:
            default_keep = rules_cfg.get('default_keep_pct')
            keep = (float(default_keep) / 100.0
                    if default_keep is not None else run_margin_pct)
            label = rules_cfg.get('default_label')

        # ── CONFLICT HIGHLIGHT: HSN signal vs the name decision ─────────
        if rules_cfg.get('flag_hsn_conflicts') and hsn_val:
            hsn_rule = None
            for rule in rules:
                if any(hsn_val.startswith(str(p))
                       for p in (rule.get('hsn_prefix') or [])):
                    hsn_rule = rule
                    break
            hsn_label = (hsn_rule.get('label') if hsn_rule
                         else rules_cfg.get('default_label'))
            if hsn_label != label:
                ckey = ('margin_conflict', str(item_no))
                if ckey not in warned_keys:
                    warned_keys.add(ckey)
                    result.warnings.append((
                        po, '',
                        f"⚠ AMBIGUOUS margin (name vs HSN): Item {item_no} "
                        f"— name → '{label}' ({int(round(keep * 100))}%), "
                        f"but HSN {hsn_val} → '{hsn_label}'. Applied the "
                        f"NAME decision; REVIEW (perfume lookup not yet "
                        f"finalised)."
                    ))

        return keep, label

    def _flag_po_status(self, df, config: Dict[str, Any],
                        result: ProcessingResult) -> None:
        """
        FLAG (but never drop) rows whose PO status is outside ``status_keep``.

        Opt-in: active only when the config sets both ``status_col`` (the
        dump's status column, e.g. Swiggy 'Status') and ``status_keep`` (the
        states that need no review, e.g. ``['CONFIRMED']``). Every line in any
        other state — EXPIRED / COMPLETED / CANCELLED / PENDING — is KEPT in the
        output (pasted as-is) and flagged with ONE named warning per such PO,
        so the operator can manually audit and remove it. We deliberately do
        NOT auto-drop: a status may be wrongly given, so the human decides.

        Mutates ``result.warnings`` only; the DataFrame is untouched.
        """
        status_col = config.get('status_col')
        keep = config.get('status_keep')
        if not (status_col and keep):
            return

        # Case-insensitive column match (dumps occasionally vary header case).
        actual = next((c for c in df.columns
                       if str(c).strip().lower() == str(status_col).strip().lower()),
                      None)
        if actual is None:
            result.warnings.append((
                '', '',
                f"Status review: column '{status_col}' not found — cannot flag "
                f"PO statuses. Columns: {list(df.columns)[:12]}…"))
            return

        keep_norm = {str(s).strip().upper() for s in keep}
        col = df[actual].astype(str).str.strip().str.upper()
        bad = ~col.isin(keep_norm)
        if not bool(bad.any()):
            return

        states = col[bad].value_counts().to_dict()
        po_col = config.get('po_col')
        # GOLDEN RULE — nothing skipped silently (and here, nothing skipped at
        # all). Name EVERY non-confirmed PO so the operator can see and
        # rectify it; the lines stay in the output for manual audit.
        n_pos = 0
        if po_col and po_col in df.columns:
            sub = df.loc[bad, [po_col]].copy()
            sub['__st__'] = col[bad]
            for po_no, grp in sub.groupby(po_col):
                sts = ', '.join(sorted(set(grp['__st__'].astype(str))))
                n_pos += 1
                result.warnings.append((
                    str(po_no), '',
                    f"PO STATUS {sts} — NOT {sorted(keep_norm)}; KEPT in output "
                    f"for manual review — remove if it should not be punched "
                    f"({len(grp)} line(s))."))
        result.warnings.append((
            '', '',
            f"Status review: {int(bad.sum())} line(s) across {n_pos} PO(s) are "
            f"not {sorted(keep_norm)} (states {states}). They were NOT dropped "
            f"— pasted as-is for manual audit."))
        logging.info("Status review (%s): flagged %d line(s) across %d POs %s",
                     config.get('party_name'), int(bad.sum()), n_pos, states)

    def _process_row(
        self,
        row: pd.Series,
        df: pd.DataFrame,
        config: Dict[str, Any],
        item_resolution: str,
        compare_basis: str,
        compare_label: str,
        margin_pct: float,
        price_col: Optional[str],
        result: ProcessingResult,
        warned_keys: Set[Tuple],
    ) -> Optional[SORow]:
        """
        Build a single SORow from one DataFrame row, or return None if
        the row should be skipped.

        Skips on: missing PO, qty ≤ 0, missing item value (mode-dependent).

        Side-effect: appends to ``result.warnings`` for unmappable
        locations, GST code surprises, price mismatches, etc.
        """
        # ── Identity: PO, location, qty ─────────────────────────────────
        po = str(row[config['po_col']]).strip()
        location = (str(row[config['loc_col']]).strip()
                    if pd.notna(row[config['loc_col']]) else '')
        qty_raw = row[config['qty_col']]

        # Parse quantity early — used both for the skip-logging below and
        # the master/mapping work further down.
        try:
            qty = int(float(qty_raw)) if pd.notna(qty_raw) else 0
        except (ValueError, TypeError):
            qty = 0

        # Thumb rule (v2.3.1): never SILENTLY drop a line that carries
        # real content. A row with no PO but with an EAN or a quantity is
        # a data problem, not a spacer — log it (with EAN/qty/row identity)
        # so it can't vanish unseen. Truly-blank rows (no PO, no EAN, no
        # qty — pandas trailing rows) are skipped quietly.
        if po.lower() == 'nan' or po == '':
            ean_dbg = self._extract_ean(row, df, config)
            if ean_dbg or qty > 0:
                key = ('NO_PO', str(row.name), ean_dbg)
                if key not in warned_keys:
                    warned_keys.add(key)
                    result.warnings.append((
                        '', '',
                        f"Row {row.name}: NO PO number — NOT written to the "
                        f"SO (EAN={ean_dbg or 'n/a'}, qty={qty}). Fix the PO "
                        f"cell in the source and re-run."))
            return None

        # Zero / invalid quantity — log (don't silently drop) so a real
        # item line never disappears without explanation.
        if qty <= 0:
            ean_dbg = self._extract_ean(row, df, config)
            key = ('ZERO_QTY', po, ean_dbg, str(row.name))
            if key not in warned_keys:
                warned_keys.add(key)
                result.warnings.append((
                    po, '',
                    f"Row {row.name}: quantity is {qty} (raw={qty_raw!r}) "
                    f"for PO {po}, EAN {ean_dbg or 'n/a'} — NOT written to "
                    f"the SO. Verify the qty in the source."))
            return None

        # ── Extract EAN (needed before item resolution for from_ean) ────
        ean = self._extract_ean(row, df, config)

        # v2.4.0 (Swiggy): the dump carries only a SkuCode (no EAN). Recover
        # the EAN from the master's 'Swiggy' sheet (SkuCode→EAN) so the rest
        # of the row flows through the standard from_ean path (master lookup,
        # validation, deal override all key off this EAN).
        if item_resolution == 'from_swiggy_sku' and not ean:
            sku_raw = row.get(config.get('sku_col')) if config.get('sku_col') else None
            if sku_raw is not None and not pd.isna(sku_raw):
                ean = self.master.swiggy_sku.get(
                    MasterLoader._clean_code(sku_raw), '') if self.master else ''

        # ── Resolve Item No per the marketplace's strategy ──────────────
        item_no = self._resolve_item_no(
            row=row, ean=ean, po=po, config=config,
            item_resolution=item_resolution,
            warned_keys=warned_keys, result=result,
        )
        if item_no is None:
            return None  # already warned inside the helper

        # ── v2.3.1: per-line margin (keep%) ─────────────────────────────
        # Marketplaces with ``margin_rules`` (Nykaa) compute landing/cost
        # per product category (Perfume/Fragrance vs Cosmetics). For every
        # other marketplace this returns the run margin unchanged. The
        # resolved rate drives BOTH the amount math and the master
        # validation below, and is stamped on the SORow for the
        # Validation sheet.
        row_margin, _margin_label = self._resolve_row_margin(
            row=row, config=config, run_margin_pct=margin_pct,
            item_no=item_no, po=po, result=result, warned_keys=warned_keys,
        )

        # ── Pull pass-through unit price (rare, both current MPs leave
        # this None so the WMS computes it downstream) ──────────────────
        unit_price = self._extract_float(row, price_col) if price_col else None

        # ── v1.5.1: Pull marketplace-native row amount when configured.
        # ``amount_col`` is optional in MARKETPLACE_CONFIGS — when
        # absent (fully unconfigured), ``amount`` stays None and the
        # email aggregator treats that as 0 for the headline stat.
        #
        # Accepted ``amount_col`` forms:
        #   1. ``str`` — single column name.
        #        Blink: 'total_amount', RK: 'Total accepted cost'.
        #   2. ``{'multiply': [col_a, col_b, ...]}`` — product of
        #        columns (v1.5.7).
        #        Myntra: ``['Landing Price', 'Quantity']`` → Landing × Qty.
        #   3. ``{'multiply': [...], 'apply_margin': True}`` — product
        #        of columns, then multiplied by the run's margin%
        #        (v1.6.0). Used when one of the "factors" is the
        #        derived landing value rather than a punch column.
        #        Reliance: ``['MRP', 'Qty'] + apply_margin=True`` →
        #        MRP × Qty × margin% = Landing × Qty.
        amount = self._extract_amount(
            row, config.get('amount_col'), df, row_margin,
        )

        # ── Marketplace prices: active (validation) + optional reference ─
        fob_price = self._extract_float(row, config.get('fob_col'),
                                         only_if_in_df=df)
        ref_fob_price = self._extract_float(row, config.get('ref_fob_col'),
                                             only_if_in_df=df)
        # v2.3.1: the marketplace's stated MRP, when the file carries one
        # (config ``mrp_col``). Used for the Validation sheet's Vendor MRP
        # vs Our MRP pair. None for marketplaces whose file has no MRP.
        vendor_mrp = self._extract_float(row, config.get('mrp_col'),
                                          only_if_in_df=df)

        # ── Master lookup + price validation ────────────────────────────
        (mrp, gst_code, description, cost_price_ref, calc_price,
         diffn, ref_diffn, validation_status,
         row_margin) = self._validate_against_master(
            ean=ean,
            item_no=item_no,
            po=po,
            margin_pct=row_margin,
            gst_margin_discount=config.get('gst_margin_discount'),
            compare_basis=compare_basis,
            compare_label=compare_label,
            fob_price=fob_price,
            ref_fob_price=ref_fob_price,
            warned_keys=warned_keys,
            result=result,
        )

        # ── Mapping lookup ──────────────────────────────────────────────
        cust_no, ship_to, mapped, mapped_location = self._resolve_mapping(
            location=location, po=po, party_name=config['party_name'],
            warned_keys=warned_keys, result=result,
        )

        # ── v1.6.0: HSN cross-check (opt-in via ``hsn_col``) ────────────
        # When the marketplace config sets ``hsn_col``, we read the
        # HSN from the punch and compare it against the master's HSN
        # for this item. Reliance is currently the only marketplace
        # with this enabled — Blink/Myntra/RK's configs leave
        # ``hsn_col`` unset and skip this block entirely.
        #
        # A mismatch isn't fatal — the row still flows through to
        # the SO and email. But it lands a per-row warning so the
        # user can audit before posting to the ERP, and the
        # Validation sheet gains an HSN Check column.
        hsn_punch, hsn_master, hsn_check_status = self._check_hsn(
            row=row, ean=ean, item_no=item_no, po=po, config=config,
            warned_keys=warned_keys, result=result,
        )

        # Source-tagging fields (Raw Data "Source" column). No marketplace
        # populates these today — multi-file batches are traced via the
        # combined raw_df's ``__source_file__`` column instead — but the
        # SORow fields are kept for forward compatibility.
        source_po = ''
        source_location = ''

        # v2.7: per-line GST% straight from the punch/PDF (config
        # 'gst_pct_col', e.g. Reliance 'GST Rate'). Used for the
        # GST-inclusive order value so it matches the PDF's Total Order
        # Value (the PDF's GST is authoritative over the master's code).
        gst_rate_pct = None
        _gpc = config.get('gst_pct_col')
        if _gpc:
            gst_rate_pct = self._extract_float(row, _gpc, only_if_in_df=df)

        # v2.4.0: Master-Exceptions per-line markers (highlight + Lines price).
        # exception_label drives the row highlight on Validation/Lines;
        # forced_unit_price writes the vendor cost into the D365 Lines Unit
        # Price for 'Use Vendor CP' rows (else blank/WMS as before).
        exception_label = ''
        forced_unit_price = None
        if self.master:
            mp = getattr(result, 'marketplace', '')
            ean_c = MasterLoader._clean_code(ean) if ean else ''
            if fob_price is not None and self.master.use_vendor_cp(
                    ean, item_no, marketplace=mp):
                exception_label = 'Vendor CP (deal)'
                forced_unit_price = fob_price
            elif self.master.price_override(ean, item_no, marketplace=mp):
                exception_label = 'Price override'
            elif ean_c and ean_c not in self.master.master \
                    and self.master.exceptions.get(ean_c):
                exception_label = 'EAN remap'

        return SORow(
            po_number=po,
            location=location,
            item_no=item_no,
            qty=qty,
            unit_price=unit_price,
            forced_unit_price=forced_unit_price,
            exception_label=exception_label,
            amount=amount,
            cust_no=cust_no,
            ship_to=ship_to,
            mapped=mapped,
            mapped_location=mapped_location,
            ean=ean,
            description=description,
            fob_price=fob_price,
            ref_fob_price=ref_fob_price,
            vendor_mrp=vendor_mrp,
            applied_margin_pct=row_margin,
            calc_price=calc_price,
            cost_price_ref=cost_price_ref,
            diffn=diffn,
            ref_diffn=ref_diffn,
            mrp=mrp,
            gst_code=gst_code,
            gst_rate_pct=gst_rate_pct,
            hsn_punch=hsn_punch,
            hsn_master=hsn_master,
            hsn_check_status=hsn_check_status,
            source_po=source_po,
            source_location=source_location,
            validation_status=validation_status,
        )

    def _check_hsn(
        self,
        row: pd.Series,
        ean: str,
        item_no: Any,
        po: str,
        config: Dict[str, Any],
        warned_keys: Set[Tuple],
        result: ProcessingResult,
    ) -> Tuple[str, str, str]:
        """
        Compare the punch's HSN against the master's HSN for this item.

        Only runs when the marketplace config has ``hsn_col`` set
        (currently Reliance only). For other marketplaces returns
        three empty strings so the Validation sheet skips its HSN
        columns.

        Status values:
            * ``''`` — not applicable (marketplace didn't opt in).
            * ``'OK'`` — punch HSN matches master HSN.
            * ``'MISMATCH'`` — both known, but they differ. Warning
              emitted (once per item_no + HSN pair via
              ``warned_keys`` so we don't flood the log when the
              same SKU appears many times).
            * ``'NOT_IN_MASTER'`` — master has no HSN for this item.
              User needs to update ``Items_March.xlsx``.

        Args:
            row:          The punch-file row (pandas Series).
            ean:          Resolved EAN (from ``_extract_ean``).
            item_no:      Resolved Item No from master lookup.
            po:           PO number (for warning attribution).
            config:       Marketplace config.
            warned_keys:  Dedup set, shared across rows of this run.
            result:       For appending warnings.

        Returns:
            ``(hsn_punch, hsn_master, hsn_check_status)`` tuple.
        """
        hsn_col = config.get('hsn_col')
        if not hsn_col:
            return ('', '', '')

        # Read punch HSN. Normalise: Excel often stores HSN codes as
        # floats (e.g. 33049990.0), so we strip the trailing .0 for a
        # clean string comparison.
        hsn_raw = row.get(hsn_col) if hsn_col in row.index else None
        hsn_punch = ''
        if hsn_raw is not None and pd.notna(hsn_raw):
            try:
                hsn_punch = str(int(float(hsn_raw)))
            except (ValueError, TypeError):
                hsn_punch = str(hsn_raw).strip()

        # Pull master HSN. The master lookup was already done by
        # _validate_against_master to build validation_status, but
        # it didn't surface hsn because callers that don't need it
        # shouldn't pay for the dict entry. Re-look up here (same
        # key priority as the main resolution: EAN first, Item No
        # fallback).
        hsn_master = ''
        if self.master is not None:
            entry = self.master.lookup(ean) if ean else None
            if entry is None and item_no:
                entry = self.master.lookup(str(item_no))
            if entry:
                hsn_master = entry.get('hsn', '') or ''

        # Decide status.
        if not hsn_master:
            status = 'NOT_IN_MASTER'
        elif hsn_punch == hsn_master:
            status = 'OK'
        else:
            status = 'MISMATCH'
            # Dedup warning by (item_no, punch_hsn, master_hsn). One
            # mismatched SKU across 50 POs shouldn't create 50
            # warning rows.
            warn_key = ('hsn_mismatch', str(item_no), hsn_punch, hsn_master)
            if warn_key not in warned_keys:
                warned_keys.add(warn_key)
                result.warnings.append((
                    po, '',
                    f"HSN mismatch on Item {item_no}: "
                    f"marketplace sent '{hsn_punch}' but master has "
                    f"'{hsn_master}'. Verify the correct HSN with "
                    f"the master data team before posting."
                ))

        return (hsn_punch, hsn_master, status)

    # ── Row-level extraction helpers ───────────────────────────────────

    @staticmethod
    def _extract_ean(row: pd.Series, df: pd.DataFrame,
                     config: Dict[str, Any]) -> str:
        """
        Read the EAN cell as a clean string, handling the float64 case
        where pandas reads ``8906121642599`` as ``8906121642599.0`` and
        we'd otherwise pass that ``.0`` into a master lookup.
        """
        ean_col = config.get('ean_col')
        if not (ean_col and ean_col in df.columns):
            return ''

        ean_raw = row[ean_col]
        if not pd.notna(ean_raw):
            return ''

        # Numeric EAN — coerce through int() to drop the trailing .0,
        # then str() for the lookup key.
        if isinstance(ean_raw, (int, float)):
            try:
                return str(int(ean_raw))
            except (ValueError, OverflowError):
                return str(ean_raw).strip()
        return str(ean_raw).strip()

    @staticmethod
    def _extract_float(row: pd.Series, col: Optional[str],
                       only_if_in_df: Optional[pd.DataFrame] = None,
                       ) -> Optional[float]:
        """
        Read ``row[col]`` as ``float | None``. Defensive against missing
        column, NaN, or non-numeric strings.
        """
        if not col:
            return None
        if only_if_in_df is not None and col not in only_if_in_df.columns:
            return None
        try:
            v = row[col]
        except KeyError:
            return None
        if not pd.notna(v):
            return None
        try:
            return float(v)
        except (ValueError, TypeError):
            # v2.3.1: tolerate currency-suffixed / thousands-separated
            # numeric strings. Flipkart's dump started shipping its cost
            # column as '114.73 INR' (string + ' INR' suffix) instead of
            # a plain number; '1,606.22' and '₹114.73' are also handled.
            # Pull the first numeric token out of the string and parse it.
            if isinstance(v, str):
                m = re.search(r'-?\d[\d,]*(?:\.\d+)?', v)
                if m:
                    try:
                        return float(m.group(0).replace(',', ''))
                    except ValueError:
                        return None
            return None

    @classmethod
    def _extract_amount(
        cls,
        row: pd.Series,
        spec: Any,
        df: pd.DataFrame,
        margin_pct: float = 0.0,
    ) -> Optional[float]:
        """
        Resolve a row's ``amount`` per the marketplace's ``amount_col``
        config spec.

        Accepted spec forms:
            * ``None`` / missing     → returns ``None`` (no amount
              configured — email will show ₹0 in the Amount stat).
            * ``str``                → reads that column as a float.
            * ``dict`` with:
                - ``'multiply'`` → iterable of column names whose
                  values are multiplied together. Any missing or
                  non-numeric factor collapses the product to
                  ``None`` for that row.
                - ``'apply_margin'`` (optional, bool, v1.6.0) → if
                  True, the final product is additionally multiplied
                  by the run's ``margin_pct``. Used when one of the
                  conceptual "factors" is the derived Landing Cost
                  (e.g. Reliance: Landing × Qty = (MRP × margin) × Qty
                  = MRP × Qty × margin%).

        Returning ``None`` on any error keeps the pipeline resilient
        — a single unparseable cell shouldn't abort the whole export.
        The failure just contributes 0 to the headline Amount stat,
        which the recipient can reconcile against the marketplace's
        own invoice if needed.

        Args:
            row:        Pandas Series for the current punch-file row.
            spec:       Whatever was in ``config.get('amount_col')``.
            df:         DataFrame (used to short-circuit when a named
                        column isn't present, avoiding per-row
                        KeyErrors).
            margin_pct: The run's margin as decimal (e.g. 0.6342).
                        Only consulted when
                        ``spec['apply_margin'] is True``.

        Returns:
            Computed ``float`` or ``None``.
        """
        if spec is None:
            return None

        # Simple column name — existing v1.5.1 behavior.
        if isinstance(spec, str):
            return cls._extract_float(row, spec, only_if_in_df=df)

        # v1.5.7+: multiply-spec for marketplaces that don't carry a
        # pre-calculated amount column but do carry the factors.
        if isinstance(spec, dict) and 'multiply' in spec:
            factors = spec['multiply']
            if not factors:
                return None
            product = 1.0
            for col in factors:
                v = cls._extract_float(row, col, only_if_in_df=df)
                if v is None:
                    # Any missing/invalid factor → no amount for this
                    # row. Callers treat None as 0 when aggregating.
                    return None
                product *= v

            # v1.6.0: apply runtime margin% as an additional factor
            # when requested. Reliance uses this because its
            # "Landing Cost" isn't a column on the punch — it's
            # derived from MRP × margin% at runtime.
            if spec.get('apply_margin'):
                product *= margin_pct

            return product

        # Unknown spec shape — log once at debug level (silent in
        # production) and skip.
        logging.debug("Unknown amount_col spec shape: %r", spec)
        return None

    def _resolve_item_no(self, row: pd.Series, ean: str, po: str,
                          config: Dict[str, Any], item_resolution: str,
                          warned_keys: Set[Tuple],
                          result: ProcessingResult) -> Any:
        """
        Resolve the canonical Item No based on ``item_resolution``.

        ``from_column`` path: read from ``item_col``.
        ``from_ean`` path: look the EAN up in the master and use
        ``master_info['item_no']``. EAN not in master → emit row with
        ``item_no = ean`` so it still appears (NOT_IN_MASTER on Validation).

        GOLDEN RULE (v2.4.4): a row that carries a PO + qty is NEVER dropped,
        even when its EAN / Item No is missing — it's KEPT with a BLANK Item No
        (NOT_IN_MASTER) and a per-PO warning, so the operator can see it and
        fill it in. Only returns ``None`` when there is genuinely nothing to
        write (and that's already logged upstream).
        """
        # v2.4.0 (Swiggy): from_swiggy_sku behaves like from_ean once the EAN
        # has been recovered from the SkuCode (done in _process_row). A SkuCode
        # that didn't resolve to an EAN surfaces the SkuCode as the placeholder
        # so the line still appears (NOT_IN_MASTER) — never silently dropped.
        if item_resolution == 'from_swiggy_sku':
            if not ean:
                sku_raw = row.get(config.get('sku_col')) if config.get('sku_col') else None
                placeholder = (MasterLoader._clean_code(sku_raw)
                               if sku_raw is not None and not pd.isna(sku_raw) else '')
                key = ('NO_SWIGGY_SKU', po, placeholder)
                if key not in warned_keys:
                    warned_keys.add(key)
                    result.warnings.append((
                        po, '',
                        f"Swiggy SkuCode '{placeholder or 'n/a'}' not found in "
                        f"the master 'Swiggy' sheet (SkuCode→EAN) for PO {po} — "
                        f"row KEPT with the SkuCode as a placeholder Item No "
                        f"(blank if no SkuCode) for manual review. Add it to the "
                        f"Swiggy sheet."))
                return placeholder   # keep the line (golden rule), even if ''
            item_resolution = 'from_ean'   # fall through to the EAN path

        if item_resolution == 'from_ean':
            if not ean:
                key = ('NO_EAN', po)
                if key not in warned_keys:
                    warned_keys.add(key)
                    result.warnings.append((
                        po, '',
                        f"PO {po}: a line has qty but NO EAN (ean_col "
                        f"'{config.get('ean_col')}' is empty) — KEPT in the "
                        f"output with a BLANK Item No (NOT_IN_MASTER) for "
                        f"manual review; fill the EAN/Item before import."
                    ))
                return ''   # golden rule: keep the line, don't drop it

            if not self.master:
                key = ('NO_MASTER', 'global')
                if key not in warned_keys:
                    warned_keys.add(key)
                    result.warnings.append((
                        '', '',
                        "Cannot resolve Item No: item_resolution='from_ean' "
                        "requires the Items_March master to be loaded."
                    ))
                return None

            master_info = self.master.lookup(ean)
            if not master_info:
                # Surface the row anyway with EAN as the visible item value
                return ean

            resolved = master_info.get('item_no', '')
            try:
                return int(resolved)
            except (ValueError, TypeError):
                return str(resolved).strip()

        # 'from_column' (default)
        item_raw = row[config['item_col']]
        if pd.isna(item_raw):
            # Golden rule: keep the line (blank Item No) + flag — never drop a
            # qty-bearing row just because its Item No cell is empty.
            key = ('NO_ITEM', po)
            if key not in warned_keys:
                warned_keys.add(key)
                result.warnings.append((
                    po, '',
                    f"PO {po}: a line has qty but NO Item No (item_col "
                    f"'{config.get('item_col')}' is empty) — KEPT in the "
                    f"output with a BLANK Item No for manual review; fill it "
                    f"before import."
                ))
            return ''   # keep the line, don't drop it
        try:
            return int(item_raw)
        except (ValueError, TypeError):
            return str(item_raw).strip()

    def _validate_against_master(
        self, ean: str, item_no: Any, po: str, margin_pct: float,
        compare_basis: str, compare_label: str,
        fob_price: Optional[float], ref_fob_price: Optional[float],
        warned_keys: Set[Tuple], result: ProcessingResult,
        gst_margin_discount: Optional[float] = None,
    ) -> Tuple[Optional[float], str, str,
               Optional[float], Optional[float],
               Optional[float], Optional[float], str, float]:
        """
        Run the master lookup and compute calc/cost/diffs/status.

        v2.3.1: when ``gst_margin_discount`` is set (Reliance), the per-
        item margin becomes GST-DEPENDENT: ``1 − discount × (1+GST)``,
        using the master's GST. This reproduces Reliance's pricing table
        (31% off pre-GST → keep 69%/67.45%/63.42% at 0/5/18% GST). The
        passed ``margin_pct`` is then ignored for this row.

        Returns (in order)::

            mrp, gst_code, description,
            cost_price_ref, calc_price,
            diffn, ref_diffn, validation_status,
            applied_margin       # the margin actually used (for the row)
        """
        mrp: Optional[float] = None
        gst_code: str = ''
        description: str = ''
        cost_price_ref: Optional[float] = None
        calc_price: Optional[float] = None
        diffn: Optional[float] = None
        ref_diffn: Optional[float] = None
        validation_status: str = ''
        applied_margin: float = margin_pct

        if not self.master:
            return (mrp, gst_code, description, cost_price_ref,
                    calc_price, diffn, ref_diffn, validation_status,
                    applied_margin)

        # Try EAN first, then fall back to Item No.
        master_info = self.master.lookup(ean) if ean else None
        if not master_info:
            master_info = self.master.lookup(str(item_no))

        if not master_info:
            # v2.3.1: the row IS still written (item_no = the EAN
            # placeholder, status NOT_IN_MASTER, highlighted orange on the
            # Validation/Summary sheets). Also log it on the Warnings sheet
            # — "not in our dump/master" should never be silent. Deduped
            # per EAN so one missing SKU across many POs = one line.
            key = ('NOT_IN_MASTER', str(ean or item_no))
            if key not in warned_keys:
                warned_keys.add(key)
                result.warnings.append((
                    po, '',
                    f"NOT IN MASTER: EAN {ean or 'n/a'} (Item {item_no}) is "
                    f"not in Items_March — the row IS written with the EAN "
                    f"as a placeholder and no MRP/GST/price. Add it to the "
                    f"master for full validation."))
            return (mrp, gst_code, description, cost_price_ref,
                    calc_price, diffn, ref_diffn, 'NOT_IN_MASTER',
                    applied_margin)

        mrp = master_info['mrp']
        gst_code = master_info['gst_code']
        description = master_info.get('description', '')

        # v2.4.4: tracks whether a deal/exception modified this row's pricing
        # (price override / Swiggy deal / vendor-CP). The dual landing+cost
        # check (``also_check_cost``) skips such rows — their CP is
        # intentionally non-standard, so it must not trip a cost MISMATCH.
        _pricing_exception = False

        # v2.4.0: central PRICE override (exceptions file). When this EAN/Item
        # has a deal price recorded for this marketplace, the expected MRP
        # and/or landing% come from the override instead of the master — so a
        # legit negotiated price (e.g. Blinkit EPISENSE: 24% off a 899 deal
        # MRP) validates as OK rather than MISMATCH. Recorded in
        # result.exceptions_applied for the Exceptions log. Wins over the
        # Reliance-style gst_margin_discount for this row.
        override = self.master.price_override(
            ean, item_no, marketplace=getattr(result, 'marketplace', ''))
        if override:
            _pricing_exception = True
            ov_mrp = override.get('mrp')
            ov_margin = override.get('margin_pct')
            if ov_mrp is not None:
                mrp = ov_mrp
            if ov_margin is not None:
                applied_margin = ov_margin
                gst_margin_discount = None     # override wins
            result.exceptions_applied.append({
                'type': 'price_override', 'po': po,
                'ean': ean, 'item_no': str(item_no),
                'detail': (f"deal MRP {ov_mrp if ov_mrp is not None else '—'}, "
                           f"landing {round(ov_margin*100,2) if ov_margin is not None else '—'}%"
                           f"{' ('+override['marketplace']+')' if override.get('marketplace') else ''}"),
            })

        # v2.4.0: record item-alias remaps that resolved this row (the EAN
        # wasn't in the master verbatim but the exceptions file mapped it to a
        # key that is — FirstCry's '…885' → '…885_1'). Only when the alias was
        # actually needed (EAN not a direct master key).
        ean_clean = MasterLoader._clean_code(ean) if ean else ''
        if (ean_clean and ean_clean not in self.master.master
                and self.master.exceptions.get(ean_clean)):
            result.exceptions_applied.append({
                'type': 'item_alias', 'po': po, 'ean': ean,
                'item_no': str(item_no),
                'detail': f"EAN {ean} → master {self.master.exceptions[ean_clean]}",
            })

        # v2.3.1: GST-dependent margin (Reliance). The keep% depends on
        # the item's GST: keep = 1 − discount × (1+GST). Replaces the run
        # margin for this row. For everyone else applied_margin stays the
        # passed margin_pct (Nykaa's per-line value or the run margin).
        if gst_margin_discount is not None:
            applied_margin = 1.0 - gst_margin_discount * \
                MasterLoader.gst_divisor(gst_code)
        margin_pct = applied_margin

        # Warn on unknown GST code (still computes, defaulting to 18%)
        gst_upper = str(gst_code).strip().upper()
        if gst_upper not in _KNOWN_GST_CODES and gst_upper != 'NAN':
            key = ('GST', gst_upper)
            if key not in warned_keys:
                warned_keys.add(key)
                result.warnings.append((
                    po, str(item_no),
                    f"Unknown GST code '{gst_code}' for Item {item_no} — "
                    f"defaulting to 18%. Please verify in Items_March."
                ))
                logging.warning("Unknown GST code '%s' for Item %s",
                                gst_code, item_no)

        # Always compute the post-GST cost price for the "naked CP"
        # column shown in the Validation sheet.
        cost_price_ref = MasterLoader.calc_cost_price(mrp, gst_code, margin_pct)

        # v2.4.0 (Swiggy): deal-SKU override. For EANs in the 'Swiggy Deal
        # SKUs' sheet the expected cost is the sheet's explicit 'Cost after
        # GST' (a negotiated deal price), NOT MRP×80%÷(1+GST). Override the
        # expected CP so the deal validates OK and the row's pricing is right.
        # v2.4.1 (BUGFIX): this is a SWIGGY-ONLY negotiated price — it must NOT
        # leak into other marketplaces. Previously any marketplace whose punch
        # carried a deal-sheet EAN (e.g. Blinkit's Villain combo 8906121643282)
        # had its expected CP wrongly clamped to Swiggy's deal cost (355.42),
        # producing a false MISMATCH. Gate on the active marketplace.
        _mp_norm = ''.join(str(getattr(result, 'marketplace', '')).split()).lower()
        if (_mp_norm == 'swiggy'
                and getattr(self.master, 'swiggy_deals', None) and ean_clean):
            sdeal = self.master.swiggy_deals.get(ean_clean)
            cag = sdeal.get('cost_after_gst') if sdeal else None
            if cag is not None:
                _pricing_exception = True
                cost_price_ref = cag
                if sdeal.get('mrp') is not None:
                    mrp = sdeal['mrp']
                result.exceptions_applied.append({
                    'type': 'price_override', 'po': po,
                    'ean': ean, 'item_no': str(item_no),
                    'detail': f"Swiggy deal SKU — expected CP {cag} (sheet)",
                })

        # Reference diff (vs naked CP) — display-only, always vs post-GST
        # because the reference column itself is post-GST (e.g. Myntra's
        # List price).
        if cost_price_ref is not None and ref_fob_price is not None:
            ref_diffn = ref_fob_price - cost_price_ref

        # Pick what we ACTUALLY compare against, based on basis.
        if compare_basis == 'landing':
            calc_price = MasterLoader.calc_landing_price(mrp, margin_pct)
        else:  # 'cost' (default)
            calc_price = cost_price_ref

        # v2.4.0: 'Use Vendor CP' exception (Master Exceptions). The vendor's
        # stated cost is authoritative for this EAN+marketplace (e.g. Myntra's
        # RENEE Goddess perfume) — accept it: expected == vendor so the row
        # validates OK instead of MISMATCH, and the Validation 'Our CP' shows
        # the vendor figure. The Lines Unit Price overwrite happens in
        # _process_row (forced_unit_price).
        if (fob_price is not None and ean_clean
                and self.master.use_vendor_cp(
                    ean, item_no, marketplace=getattr(result, 'marketplace', ''))):
            _pricing_exception = True
            calc_price = fob_price
            cost_price_ref = fob_price
            result.exceptions_applied.append({
                'type': 'vendor_cp', 'po': po,
                'ean': ean, 'item_no': str(item_no),
                'detail': f"vendor CP {fob_price} accepted (deal) — written to Lines",
            })

        # Compute active diff + status
        if calc_price is not None and fob_price is not None:
            diffn = fob_price - calc_price
            if abs(diffn) <= self.DIFFN_THRESHOLD:
                validation_status = 'OK'
            else:
                validation_status = 'MISMATCH'
                key = ('VALIDATION', str(item_no))
                if key not in warned_keys:
                    warned_keys.add(key)
                    result.warnings.append((
                        po, str(item_no),
                        f"{compare_label} mismatch: Item {item_no}, "
                        f"Marketplace={fob_price:.2f}, "
                        f"Calculated={calc_price:.2f}, "
                        f"Diff={diffn:.2f}"
                    ))

            # v2.4.4: DUAL landing+cost check (opt-in via ``also_check_cost``,
            # Myntra). The landing check above governs status by default; with
            # this flag a row is OK only when BOTH the landing pair AND the
            # cost pair (vendor CP = ``ref_fob_price`` vs our CP =
            # ``cost_price_ref``) agree — so either failing → MISMATCH. Rows
            # whose pricing came from a deal/exception are skipped (their CP is
            # intentionally non-standard). Only downgrades an otherwise-OK row;
            # a landing failure is already MISMATCH.
            if (validation_status == 'OK'
                    and (result.resolved_config or {}).get('also_check_cost')
                    and not _pricing_exception
                    and ref_fob_price is not None
                    and cost_price_ref is not None
                    and abs(ref_fob_price - cost_price_ref) > self.DIFFN_THRESHOLD):
                validation_status = 'MISMATCH'
                key = ('VALIDATION_CP', str(item_no))
                if key not in warned_keys:
                    warned_keys.add(key)
                    result.warnings.append((
                        po, str(item_no),
                        f"Cost mismatch: Item {item_no}, "
                        f"Vendor CP={ref_fob_price:.2f}, "
                        f"Our CP={cost_price_ref:.2f}, "
                        f"Diff={ref_fob_price - cost_price_ref:.2f} "
                        f"(landing matched)"
                    ))
        else:
            validation_status = 'NO_PRICE'

        return (mrp, gst_code, description, cost_price_ref,
                calc_price, diffn, ref_diffn, validation_status,
                applied_margin)

    def _resolve_mapping(self, location: str, po: str, party_name: str,
                          warned_keys: Set[Tuple],
                          result: ProcessingResult,
                          ) -> Tuple[str, str, bool, str]:
        """
        Look up the location in the mapping registry.

        Returns (cust_no, ship_to, mapped_bool, mapped_location_str).
        On miss, appends a warning (deduped per (po, location)) and
        returns blanks plus mapped=False.
        """
        # v2.7.x: address-based marketplaces (Flipkart — loc_col is a full
        # postal address) opt into the pincode+body-overlap resolver, which
        # matches the new portal's prefix-less Shipped-To address against the
        # Ship-To B2B Del Location entries. Everyone else uses the generic
        # name-based tiers unchanged.
        if (result.resolved_config or {}).get('loc_match') == 'address':
            mapping_result = self.mapping.lookup_by_address(location)
        else:
            mapping_result = self.mapping.lookup(location)

        if mapping_result:
            return (
                mapping_result['cust_no'],
                mapping_result['ship_to'],
                True,
                mapping_result.get('matched_key', location),
            )

        # Unmapped — warn once per (po, location)
        key = (po, location)
        if key not in warned_keys:
            warned_keys.add(key)
            result.warnings.append((
                po, location,
                f"Location '{location}' not found in mapping for {party_name}. "
                f"Cust No and Ship-to left empty."
            ))
        return ('', '', False, '')