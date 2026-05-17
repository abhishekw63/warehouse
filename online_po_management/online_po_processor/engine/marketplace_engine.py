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
from typing import Any, Dict, Optional, Set, Tuple

import pandas as pd

from online_po_processor.data.models import SORow, ProcessingResult
from online_po_processor.data.master_loader import MasterLoader
from online_po_processor.data.mapping_loader import MappingLoader


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
        # Some marketplaces carry the PO number and/or Location in a
        # header/title position rather than per-data-row. For those, a
        # pre-processor hook parses the out-of-band values and injects
        # synthetic columns so the rest of the pipeline sees a normal
        # wide-format DataFrame.
        #
        # Currently used by Reliance (parses row 0's merged title
        # like "5000466441  BHIWANDI (Reliance)" into per-row
        # ``__po__`` and ``__loc__`` columns).
        pre_process = config.get('pre_process')
        if pre_process == 'reliance_po_sheet':
            df = self._preprocess_reliance(
                filepath, sheet_to_read, df, config, result,
            )
            if df is None:
                return result  # warning already appended

        result.raw_df = df

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

    def _preprocess_reliance(
        self,
        filepath: str,
        sheet_name: str,
        df: pd.DataFrame,
        config: Dict[str, Any],
        result: ProcessingResult,
    ) -> Optional[pd.DataFrame]:
        """
        Inject synthetic ``__po__`` and ``__loc__`` columns for Reliance.

        Reliance's raw PO attachment doesn't carry the PO number or
        delivery location in a data column — they appear as a merged
        "title" cell on row 0 of the PO sheet, formatted like::

            5000466441  BHIWANDI (Reliance)

        (The first token is the 10-digit PO number; everything after
        the first whitespace gap is the delivery location label which
        must match an entry in the Ship-To B2B mapping sheet.)

        To fit the rest of the pipeline without special-casing
        downstream, we:

        1. Re-read row 0 raw (the header=1 read we already did
           skipped it).
        2. Scan that row's cells for the first non-empty string.
        3. Regex-split into PO number + location.
        4. Inject ``__po__`` and ``__loc__`` on every data row so the
           engine can point ``po_col='__po__'`` / ``loc_col='__loc__'``
           and otherwise run normally.

        A single-PO assumption is hard-coded. Reliance's B2B ordering
        system always emits one PO per attachment file; if a future
        format change merges multiple POs, the regex will match the
        first title and silently mislabel the rest — so we also warn
        if there's any sign of a second title row.

        Args:
            filepath:    The Reliance file on disk.
            sheet_name:  Name of the sheet the data was read from
                         (typically 'PO').
            df:          Data already loaded with header=1 (i.e.
                         title row skipped, proper headers used).
            config:      Marketplace config (for future extension).
            result:      For appending abort warnings.

        Returns:
            DataFrame with ``__po__`` and ``__loc__`` injected, or
            ``None`` if the title cannot be parsed (caller should
            ``return result`` immediately).
        """
        # Re-read row 0 as raw (header=None), taking only the first row.
        try:
            title_df = pd.read_excel(
                filepath, sheet_name=sheet_name, header=None, nrows=1,
            )
        except Exception as e:  # noqa: BLE001
            result.warnings.append((
                '', '',
                f"Reliance pre-process: cannot read title row of "
                f"sheet {sheet_name!r}: {e}",
            ))
            return None

        # Find the first non-empty string cell in row 0. The title is
        # typically in a merged cell somewhere around columns D-E but
        # we scan the whole row rather than hard-code a position.
        title_text = None
        for cell in title_df.iloc[0].tolist():
            if cell is not None and pd.notna(cell):
                text = str(cell).strip()
                if text:
                    title_text = text
                    break

        if not title_text:
            result.warnings.append((
                '', '',
                f"Reliance pre-process: title row of sheet "
                f"{sheet_name!r} is empty. Expected a cell like "
                f"'5000466441  BHIWANDI (Reliance)' with the PO "
                f"number and location. Cannot identify this PO.",
            ))
            return None

        # Parse: first whitespace-delimited token = PO number, rest
        # = location. We use split(maxsplit=1) so locations with
        # multiple words + parentheses ("BHIWANDI (Reliance)") stay
        # intact.
        parts = title_text.split(maxsplit=1)
        if len(parts) < 2:
            result.warnings.append((
                '', '',
                f"Reliance pre-process: title cell {title_text!r} "
                f"doesn't contain both a PO number and a location. "
                f"Expected something like '5000466441  BHIWANDI "
                f"(Reliance)'.",
            ))
            return None

        po_number, location = parts[0].strip(), parts[1].strip()
        logging.info(
            "Reliance pre-process: parsed title %r → PO=%r loc=%r",
            title_text, po_number, location,
        )

        # Inject synthetic columns. Every data row inherits the same
        # PO and location — correct because Reliance sends one PO
        # per file.
        df = df.copy()
        df['__po__'] = po_number
        df['__loc__'] = location

        # Drop rows where Qty is blank — those are separator/total
        # rows in Reliance's layout (unlikely on the clean PO sheet
        # but be defensive; also drops any trailing blank rows pandas
        # might have loaded).
        qty_col_name = config.get('qty_col', 'Qty')
        if qty_col_name in df.columns:
            before = len(df)
            df = df[df[qty_col_name].notna()].reset_index(drop=True)
            if len(df) < before:
                logging.info(
                    "Reliance pre-process: dropped %d blank-Qty rows",
                    before - len(df),
                )

        return df

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

        if item_resolution == 'from_ean':
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
                skipped_no_ean += 1
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
            cust_no, ship_to, mapped, mapped_loc = self._resolve_mapping(
                loc_str, po_str, party_name, warned_locations, result,
            )
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
                }
            else:
                entry['qty'] += qty
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
            )
            result.rows.append(so_row)

        # Roll-up warnings.
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

        # Skip rows with no PO number
        if po.lower() == 'nan':
            return None

        # Parse quantity early — a zero-qty row contributes nothing,
        # avoid running master/mapping lookups for it.
        try:
            qty = int(float(qty_raw)) if pd.notna(qty_raw) else 0
        except (ValueError, TypeError):
            qty = 0
        if qty <= 0:
            return None

        # ── Extract EAN (needed before item resolution for from_ean) ────
        ean = self._extract_ean(row, df, config)

        # ── Resolve Item No per the marketplace's strategy ──────────────
        item_no = self._resolve_item_no(
            row=row, ean=ean, po=po, config=config,
            item_resolution=item_resolution,
            warned_keys=warned_keys, result=result,
        )
        if item_no is None:
            return None  # already warned inside the helper

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
            row, config.get('amount_col'), df, margin_pct,
        )

        # ── Marketplace prices: active (validation) + optional reference ─
        fob_price = self._extract_float(row, config.get('fob_col'),
                                         only_if_in_df=df)
        ref_fob_price = self._extract_float(row, config.get('ref_fob_col'),
                                             only_if_in_df=df)

        # ── Master lookup + price validation ────────────────────────────
        (mrp, gst_code, description, cost_price_ref, calc_price,
         diffn, ref_diffn, validation_status) = self._validate_against_master(
            ean=ean,
            item_no=item_no,
            po=po,
            margin_pct=margin_pct,
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

        # ── v1.7.0: Source tagging for batch/multi-file traceability ────
        # When the marketplace went through a ``pre_process`` hook
        # (currently Reliance), the engine parsed the PO number and
        # location from the file's title row rather than from
        # per-row columns. Stamp those onto the SORow so the Raw
        # Data sheet can show a "Source" column with values like
        # '5000466441 BHIWANDI (Reliance)' — crucial when the user
        # uploads 5 Reliance PO files at once and wants to see
        # which row came from which file. Stays blank for
        # Blink/Myntra/RK, whose Raw Data sheet doesn't need it.
        if config.get('pre_process'):
            source_po = po
            source_location = location
        else:
            source_po = ''
            source_location = ''

        return SORow(
            po_number=po,
            location=location,
            item_no=item_no,
            qty=qty,
            unit_price=unit_price,
            amount=amount,
            cust_no=cust_no,
            ship_to=ship_to,
            mapped=mapped,
            mapped_location=mapped_location,
            ean=ean,
            description=description,
            fob_price=fob_price,
            ref_fob_price=ref_fob_price,
            calc_price=calc_price,
            cost_price_ref=cost_price_ref,
            diffn=diffn,
            ref_diffn=ref_diffn,
            mrp=mrp,
            gst_code=gst_code,
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

        ``from_column`` path: read from ``item_col``. NaN → skip row.
        ``from_ean`` path: look the EAN up in the master and use
        ``master_info['item_no']``. Empty EAN → skip with warning.
        EAN not in master → emit row with ``item_no = ean`` so it still
        appears (in NOT_IN_MASTER state in the validation sheet).

        Returns ``None`` if the row should be skipped (warnings already
        appended where appropriate).
        """
        if item_resolution == 'from_ean':
            if not ean:
                key = ('NO_EAN', po)
                if key not in warned_keys:
                    warned_keys.add(key)
                    result.warnings.append((
                        po, '',
                        f"Row skipped: ean_col '{config.get('ean_col')}' is "
                        f"empty for PO {po}. item_resolution='from_ean' "
                        f"requires a non-empty EAN."
                    ))
                return None

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
            return None
        try:
            return int(item_raw)
        except (ValueError, TypeError):
            return str(item_raw).strip()

    def _validate_against_master(
        self, ean: str, item_no: Any, po: str, margin_pct: float,
        compare_basis: str, compare_label: str,
        fob_price: Optional[float], ref_fob_price: Optional[float],
        warned_keys: Set[Tuple], result: ProcessingResult,
    ) -> Tuple[Optional[float], str, str,
               Optional[float], Optional[float],
               Optional[float], Optional[float], str]:
        """
        Run the master lookup and compute calc/cost/diffs/status.

        Returns (in order)::

            mrp, gst_code, description,
            cost_price_ref, calc_price,
            diffn, ref_diffn, validation_status
        """
        mrp: Optional[float] = None
        gst_code: str = ''
        description: str = ''
        cost_price_ref: Optional[float] = None
        calc_price: Optional[float] = None
        diffn: Optional[float] = None
        ref_diffn: Optional[float] = None
        validation_status: str = ''

        if not self.master:
            return (mrp, gst_code, description, cost_price_ref,
                    calc_price, diffn, ref_diffn, validation_status)

        # Try EAN first, then fall back to Item No.
        master_info = self.master.lookup(ean) if ean else None
        if not master_info:
            master_info = self.master.lookup(str(item_no))

        if not master_info:
            return (mrp, gst_code, description, cost_price_ref,
                    calc_price, diffn, ref_diffn, 'NOT_IN_MASTER')

        mrp = master_info['mrp']
        gst_code = master_info['gst_code']
        description = master_info.get('description', '')

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
        else:
            validation_status = 'NO_PRICE'

        return (mrp, gst_code, description, cost_price_ref,
                calc_price, diffn, ref_diffn, validation_status)

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