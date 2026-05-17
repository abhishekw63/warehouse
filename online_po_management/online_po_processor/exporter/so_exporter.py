"""
exporter.so_exporter
====================

Orchestrates writing the output workbook.

The class itself is intentionally thin: it decides WHERE to write the
file, creates the workbook, and delegates the HOW of each sheet to the
per-sheet modules in :mod:`online_po_processor.exporter.sheets`. This
keeps each sheet's logic independent and easy to change without
touching the others.

Output location (v1.3.4)
------------------------
The output file is written next to the **input punch file** in an
``output/`` subfolder that's auto-created::

    D:\\PO\\Myntra\\April\\Myntra_Punch_17-04-2026.xlsx      ← input
    D:\\PO\\Myntra\\April\\output\\myntra_so_19-04-2026_*.xlsx ← output

Each month's batch lives next to its inputs instead of piling up in a
single global folder.

If ``result.input_file_path`` is empty (shouldn't happen in normal use)
the exporter falls back to the script directory's ``output_online/``
folder — defensive only, won't trigger during a normal run.

Filename format
---------------
``<marketplace_slug>_so_<DD-MM-YYYY>_<HHMMSS>.xlsx``

The timestamp means repeat runs never clobber prior outputs.
"""

from __future__ import annotations
import logging
import time
from datetime import datetime
from pathlib import Path
from tkinter import messagebox
from typing import Optional

from openpyxl import Workbook

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter.sheets import (
    headers_sheet, lines_sheet, raw_data_sheet,
    summary_sheet, validation_sheet, warnings_sheet,
)


class SOExporter:
    """
    Write a workbook from a :class:`ProcessingResult`.

    Stateless — instances hold nothing, so reusing the same instance
    across runs is safe (and cheap).
    """

    def export(self, result: ProcessingResult,
                start_time: Optional[float] = None) -> Optional[Path]:
        """
        Render ``result`` to an .xlsx file on disk.

        v2.1.0: ``start_time`` parameter added. When supplied (typically
        a ``time.time()`` snapshot the GUI captured before any work
        began), the exporter computes the full pipeline elapsed time
        right before saving and stamps it onto ``result.elapsed_seconds``.
        That makes the duration visible to:

          * The Summary sheet's footer (read inline via
            ``result.elapsed_seconds`` during ``summary_sheet.write``).
          * The email report (read after this method returns).

        When ``start_time`` is None, ``result.elapsed_seconds`` is left
        untouched so older callers and tests behave exactly as before
        (Summary footer omits the duration segment).

        Sheets are written in order — Summary in particular reads
        ``result.elapsed_seconds`` at write-time, so we set it BEFORE
        the sheet writers run, not after. The actual file-save step
        (a few hundred ms for typical workbook sizes) isn't reflected
        in the printed duration; that's an acceptable trade-off for
        having the value visible in the file at all.

        Args:
            result:     Fully-populated result from
                        :meth:`MarketplaceEngine.process`.
            start_time: Optional ``time.time()`` snapshot taken before
                        the pipeline started. Used to compute and stamp
                        the run duration.

        Returns:
            ``Path`` to the saved file on success, ``None`` when there
            were no rows to write (a user-facing warning dialog is
            shown in that case).
        """
        if not result.rows:
            # No rows == nothing to import. Better to tell the user
            # than to silently produce an empty workbook.
            messagebox.showwarning(
                "No Data",
                "No valid rows found.\nNothing to export.",
            )
            return None

        # v2.1.0: compute elapsed BEFORE the sheet writers run so the
        # Summary footer can see the value. The few hundred ms spent
        # writing the workbook itself isn't reflected, but that's
        # negligible compared to engine + master-load time on a
        # typical batch (a few seconds).
        if start_time is not None:
            result.elapsed_seconds = time.time() - start_time

        file_path = self._resolve_output_path(result)

        wb = Workbook()
        # Workbook() auto-creates an empty 'Sheet' — remove it before we
        # add our own named sheets so the final book has exactly the
        # tabs we want.
        wb.remove(wb.active)

        # Sheet order matters for the user's reading flow:
        #   Headers (SO)  → ERP import (top tab = what you act on)
        #   Lines (SO)    → ERP import
        #   Summary       → human verification, per-PO
        #   Validation    → human verification, per-item price check
        #   Warnings      → only present if there are issues to fix
        #   Raw Data      → audit trail at the bottom
        headers_sheet.write(wb, result)
        lines_sheet.write(wb, result)
        summary_sheet.write(wb, result)
        validation_sheet.write(wb, result)
        warnings_sheet.write(wb, result)
        raw_data_sheet.write(wb, result)

        wb.save(str(file_path))
        logging.info("Output saved: %s", file_path)
        return file_path

    # ── Internal helpers ──────────────────────────────────────────────

    @staticmethod
    def _resolve_output_path(result: ProcessingResult) -> Path:
        """
        Compute the full output path and ensure the parent folder exists.

        Prefers ``<punch-dir>/output/`` so each batch's output lives
        next to the input it came from. For v1.7.0 multi-file batches,
        the "punch dir" is the parent of the FIRST file (typical case:
        all the files in a Reliance batch live in the same folder, so
        first-file's parent is the right destination).

        Filename conventions:
            * Single-file run:
                ``<marketplace>_so_<DD-MM-YYYY>_<HHMMSS>.xlsx``
                e.g. ``reliance_so_22-04-2026_154523.xlsx``
            * Multi-file batch (v1.7.0, Reliance only):
                ``<marketplace>_<N>PO_<DD-MM-YYYY>_<HHMMSS>.xlsx``
                e.g. ``reliance_5PO_22-04-2026_154523.xlsx``
              The count makes it obvious at a glance that this output
              covers multiple POs without forcing the user to open
              the workbook to find out.

        The timestamp means repeat runs never clobber prior outputs.
        """
        if result.input_file_path:
            output_folder = Path(result.input_file_path).parent / 'output'
        else:
            # Defensive fallback — the engine always populates
            # input_file_path, so hitting this branch indicates either
            # a programming error or tests that bypass the engine.
            output_folder = Path('output_online')

        output_folder.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime('%d-%m-%Y_%H%M%S')
        marketplace_slug = result.marketplace.lower().replace(' ', '_')

        # v1.7.0: multi-file batches get a count-based filename so
        # the batch size is visible without opening the workbook.
        # v2.0.0: TO marketplaces use a '_to_' slug instead of '_so_'
        # so the filename signals which import shape it carries (the
        # ERP team consumes TO and SO files via different D365 import
        # frames, so the distinction matters).
        n_files = getattr(result, 'input_files_count', 1) or 1
        kind_slug = 'to' if getattr(result, 'output_type', 'so') == 'to' else 'so'
        if n_files > 1:
            stem = f'{marketplace_slug}_{n_files}PO_{timestamp}.xlsx'
        else:
            stem = f'{marketplace_slug}_{kind_slug}_{timestamp}.xlsx'
        return output_folder / stem