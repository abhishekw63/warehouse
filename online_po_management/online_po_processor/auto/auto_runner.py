"""
auto.auto_runner
================

Headless batch runner for **AUTO mode** (v2.4.0).

Manual mode (the Tkinter GUI in ``gui/app_window.py``) is the default and
is left completely untouched — this module is a SEPARATE path that drives
the SAME engine + exporter without any GUI. If Auto ever misbehaves,
Manual is the open door.

Folder convention
-----------------
::

    Dump/
      Online/
        Reliance/   ← drop Reliance PO PDFs here
        RK/         ← drop the RK dump here
        Zepto/      ← drop the Zepto dump here
        ... (one folder per marketplace) ...

Each subfolder is NAMED for a marketplace, so there is no detection
guesswork — the folder name IS the marketplace. Every file in a folder is
processed with that marketplace's config, its per-marketplace
``default_margin``, and the run's chosen dispatch warehouse. Outputs land
in ``<marketplace>/output/`` (coupled next to the input), exactly like
Manual mode (the exporter already writes there).

Per-marketplace batching mirrors the GUI's ``generate()``:

* **Consignment TO** (Meesho-TO, Flipkart-TO when the folder holds
  ``Consignment_Details_*`` / ``order-line-items-*`` CSVs) → one combined
  ``process_consignments`` run.
* **PDF** (Reliance, FirstCry, Dmart) → all PDFs combined via
  ``process_multi`` into one SO batch.
* **Excel/CSV single-dump** (Zepto, RK, Blink, Myntra, BlinkMP, Flipkart,
  Nykaa, plus Flipkart-TO's consolidated-dump form) → one run per file.

Nothing here prompts the user. It runs unattended and returns a flat list
of :class:`MarketplaceRun` records the caller (a thin Auto UI or CLI) can
print. One marketplace/file failing never stops the rest — the failure is
captured on its record and the loop continues.

The ``no-rows`` guard matters: ``SOExporter.export`` shows a Tk warning
dialog when a result has zero rows — so the runner NEVER calls export on
an empty result (it records ``status='no_rows'`` instead), keeping Auto
mode fully headless.
"""

from __future__ import annotations

import logging
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, List, Optional

from online_po_processor.config.marketplaces import (
    DEFAULT_WAREHOUSE,
    MARKETPLACE_CONFIGS,
    WAREHOUSE_CODES,
)
from online_po_processor.data.mapping_loader import MappingLoader
from online_po_processor.data.master_loader import MasterLoader
from online_po_processor.data.models import ProcessingResult
from online_po_processor.engine.marketplace_engine import MarketplaceEngine
from online_po_processor.exporter.so_exporter import SOExporter


# Input file extensions Auto mode will pick up from a marketplace folder.
# Everything else (and the 'output/' subfolder, Excel lock files, hidden
# files) is ignored.
_INPUT_EXTS = {'.xlsx', '.xls', '.xlsm', '.csv', '.pdf'}


@dataclass
class MarketplaceRun:
    """
    Outcome of processing ONE batch for a marketplace.

    A marketplace folder can produce several batches (e.g. two Excel
    dumps → two single-file runs), so the runner returns a flat list of
    these. ``result`` is retained so a later consolidation step can
    combine Headers + Lines across all successful runs.

    ``status``:
      * ``'ok'``       — processed and exported (``output_path`` set)
      * ``'no_files'`` — the marketplace folder was empty/absent
      * ``'no_rows'``  — files processed but nothing extracted (NOT
                         exported — kept headless; see module docstring)
      * ``'error'``    — an exception was caught (see ``error``)
    """

    marketplace: str
    status: str
    input_files: List[str] = field(default_factory=list)
    warehouse: str = ''
    output_path: Optional[str] = None
    rows: int = 0
    pos: int = 0
    qty: int = 0
    skipped: int = 0          # already-uploaded POs removed from output
    warnings: int = 0
    error: str = ''
    result: Optional[ProcessingResult] = None


def _input_files(folder: Path) -> List[Path]:
    """Sorted list of processable input files directly in ``folder``."""
    out: List[Path] = []
    for p in sorted(folder.iterdir()):
        if p.is_dir():
            continue                       # skips output/ and any subfolders
        if p.name.startswith('~$') or p.name.startswith('.'):
            continue                       # Excel lock / hidden files
        if p.suffix.lower() not in _INPUT_EXTS:
            continue
        out.append(p)
    return out


class AutoRunner:
    """
    Headless batch processor over a ``Dump/Online/`` tree.

    Loads the master once, then for each marketplace folder loads that
    marketplace's mapping, runs the appropriate engine path, and exports
    — collecting a :class:`MarketplaceRun` per batch.
    """

    def __init__(
        self,
        master_path: str,
        mapping_path: str,
        warehouse: str = DEFAULT_WAREHOUSE,
        warehouse_map: Optional[dict] = None,
        log: Optional[Callable[[str], None]] = None,
    ) -> None:
        self.master_path = master_path
        self.mapping_path = mapping_path
        # ``warehouse`` is the fallback dispatch warehouse; ``warehouse_map``
        # ({marketplace: 'AHD'|'BLR'}) overrides it per marketplace so the
        # operator can route, say, Zepto from BLR and the rest from AHD in
        # one run. A marketplace absent from the map uses ``warehouse``.
        self.warehouse = warehouse
        self.warehouse_map = warehouse_map or {}
        self._log_fn = log or (lambda m: logging.info("%s", m))
        self._master: Optional[MasterLoader] = None

    def _warehouse_for(self, marketplace: str) -> str:
        return self.warehouse_map.get(marketplace, self.warehouse)

    # ── logging ────────────────────────────────────────────────────────
    def log(self, msg: str) -> None:
        self._log_fn(msg)

    # ── public entry point ─────────────────────────────────────────────
    def run(
        self,
        online_root: str,
        marketplaces: Optional[List[str]] = None,
    ) -> List[MarketplaceRun]:
        """
        Process every marketplace folder under ``online_root``.

        Args:
            online_root:  Path to ``Dump/Online``.
            marketplaces: Optional explicit subset (config keys). When
                          None, every configured marketplace that has a
                          matching subfolder is processed.

        Returns:
            Flat list of :class:`MarketplaceRun` (one per batch).
        """
        root = Path(online_root)
        self._load_master()

        names = marketplaces or [
            mp for mp in MARKETPLACE_CONFIGS if (root / mp).is_dir()
        ]
        self.log(
            f"AUTO run: {len(names)} marketplace folder(s) under {root} "
            f"| default warehouse {self.warehouse} "
            f"({WAREHOUSE_CODES.get(self.warehouse, 'PICK')})"
        )

        runs: List[MarketplaceRun] = []
        for mp in names:
            runs.extend(self._run_one(root / mp, mp))
        return runs

    # ── master ─────────────────────────────────────────────────────────
    def _load_master(self) -> None:
        master = MasterLoader()
        n = master.load(self.master_path)
        self.log(f"Master loaded: {n:,} items")
        self._master = master

    # ── one marketplace folder → 1+ batches ────────────────────────────
    def _run_one(self, folder: Path, marketplace: str) -> List[MarketplaceRun]:
        if not folder.is_dir():
            return [MarketplaceRun(marketplace, 'no_files')]

        files = _input_files(folder)
        if not files:
            self.log(f"[{marketplace}] no files — skipped")
            return [MarketplaceRun(marketplace, 'no_files')]

        config = MARKETPLACE_CONFIGS[marketplace]
        paths = [str(f) for f in files]
        cons = config.get('consignment_mode')

        # Decide the batch split: list of (paths, mode, visibility_path).
        # 'mode' picks the engine call in _run_batch; 'visibility_path' is
        # only meaningful for consignment mode (else None). This mirrors
        # the GUI's generate().
        batches: List[tuple] = []
        if cons and cons.get('enabled'):
            # v2.4.1: a Consignment Visibility Report dropped in the folder
            # alongside the per-PO consignment files is auto-detected by
            # filename and threaded into process_consignments — so Auto
            # mode resolves Locations exactly like the manual GUI (which
            # took the report as a separate pick). Detection is by the
            # ``visibility_filename_regex`` config key; the report itself
            # is NOT treated as a consignment / consolidated dump.
            vis_re = cons.get('visibility_filename_regex')
            vis_files = [p for p in paths
                         if vis_re and re.search(vis_re, Path(p).name,
                                                 re.IGNORECASE)]
            vis_path = vis_files[0] if vis_files else None
            if vis_files and len(vis_files) > 1:
                self.log(f"[{marketplace}] {len(vis_files)} visibility "
                         f"reports found — using {Path(vis_path).name}")
            rest_paths = [p for p in paths if p not in vis_files]

            if cons.get('consolidated_option'):
                # Flipkart-TO: consignment CSVs are recognised by filename;
                # anything else left over is a consolidated dump.
                regex = cons.get('filename_po_regex')
                cfiles = [p for p in rest_paths
                          if regex and re.search(regex, Path(p).name)]
                if cfiles:
                    batches.append((cfiles, 'consignment', vis_path))
                    rest = [p for p in rest_paths if p not in cfiles]
                    batches += [([p], 'single', None) for p in rest]
                else:
                    batches += [([p], 'single', None) for p in rest_paths]
            else:
                # Meesho-TO: consignment-only.
                batches.append((rest_paths, 'consignment', vis_path))
        elif config.get('source_format') == 'pdf':
            # PDF POs are one-per-file → combine into one SO batch.
            batches.append((paths, 'multi', None))
        elif config.get('pdf_parser'):
            # v2.4.1: dual-format (Myntra) — a folder of PO PDFs combines
            # into one SO batch (like PDF-only marketplaces); any non-PDF
            # (Excel punch) files stay one-run-per-file.
            pdfs = [p for p in paths if p.lower().endswith('.pdf')]
            others = [p for p in paths if not p.lower().endswith('.pdf')]
            if pdfs:
                batches.append((pdfs, 'multi', None))
            batches += [([p], 'single', None) for p in others]
        else:
            # Excel/CSV single-dump marketplaces: one run per file so a
            # second dump in the folder is never silently ignored.
            batches += [([p], 'single', None) for p in paths]

        return [
            self._run_batch(marketplace, config, bpaths, mode, vis)
            for bpaths, mode, vis in batches
        ]

    # ── one batch → process + export ───────────────────────────────────
    def _run_batch(
        self, marketplace: str, config: dict,
        paths: List[str], mode: str,
        visibility_report_path: Optional[str] = None,
    ) -> MarketplaceRun:
        run = MarketplaceRun(
            marketplace=marketplace, status='ok',
            input_files=[Path(p).name for p in paths],
        )
        margin = config.get('default_margin', 70) / 100.0
        try:
            # Fresh mapping per batch, filtered to this marketplace's party.
            mapping = MappingLoader()
            warns: list = []
            loc = mapping.load(self.mapping_path, config['party_name'], warns)
            if loc == 0:
                raise RuntimeError(
                    f"no mapping locations for party '{config['party_name']}'"
                )
            engine = MarketplaceEngine(mapping, master=self._master)

            if mode == 'consignment':
                if visibility_report_path:
                    self.log(f"[{marketplace}] using visibility report: "
                             f"{Path(visibility_report_path).name}")
                result = engine.process_consignments(
                    paths, config, margin_pct=margin,
                    visibility_report_path=visibility_report_path,
                )
            elif mode == 'multi':
                result = (engine.process_multi(paths, config, margin_pct=margin)
                          if len(paths) > 1
                          else engine.process(paths[0], config, margin_pct=margin))
            else:  # 'single'
                result = engine.process(paths[0], config, margin_pct=margin)

            run.result = result
            run.warnings = len(result.warnings) if result else 0

            if not result or not result.rows:
                # Guard: never call export() on an empty result (it would
                # pop a Tk dialog). Record and move on.
                run.status = 'no_rows'
                self.log(f"[{marketplace}] processed but 0 rows extracted")
                return run

            # Stamp the run-level metadata the exporter / sheets expect.
            # Warehouse is resolved per marketplace (warehouse_map override
            # → fallback default), matching the operator's per-marketplace
            # AHD/BLR selection in the Auto screen.
            wh = self._warehouse_for(marketplace)
            run.warehouse = wh
            result.margin_pct = margin
            result.warehouse_display = wh
            result.warehouse_code = WAREHOUSE_CODES.get(wh, 'PICK')
            result.override_unit_price = bool(
                config.get('override_unit_price', False)
            )

            # v2.4.0: dedup-skip — remove already-uploaded POs from this
            # result BEFORE export, so Headers/Lines never re-send them.
            from online_po_processor.auto.history_db import apply_dedup
            skipped = apply_dedup(result)
            run.skipped = len(skipped)
            if skipped:
                self.log(f"[{marketplace}] {len(skipped)} already-uploaded "
                         f"PO(s) removed (Skipped POs sheet)")

            out = SOExporter().export(result, start_time=time.time())
            run.output_path = str(out) if out else None
            run.rows = len(result.rows)
            run.pos = len({r.po_number for r in result.rows})
            run.qty = sum(r.qty for r in result.rows)
            self.log(
                f"[{marketplace}] {wh} | {run.rows} items / {run.pos} PO(s) / "
                f"qty {run.qty}"
                + (f" / {run.warnings} warn" if run.warnings else "")
                + f"  ->  {Path(out).name if out else 'NO OUTPUT'}"
            )
        except Exception as e:  # noqa: BLE001 — one bad batch must not abort the run
            run.status = 'error'
            run.error = f"{type(e).__name__}: {e}"
            self.log(f"[{marketplace}] ERROR — {run.error}")
        return run


def summarize(runs: List[MarketplaceRun]) -> str:
    """Build a compact human-readable report from a run list."""
    ok = [r for r in runs if r.status == 'ok']
    empty = [r for r in runs if r.status == 'no_rows']
    nofiles = [r for r in runs if r.status == 'no_files']
    errors = [r for r in runs if r.status == 'error']

    lines = ["", "═" * 60, "AUTO RUN SUMMARY", "═" * 60]
    for r in ok:
        lines.append(
            f"  ✓ {r.marketplace:12} {r.warehouse:3} {r.rows:5} items  "
            f"{r.pos:3} PO  qty {r.qty:7}"
            + (f"  ⚠{r.warnings}" if r.warnings else "")
        )
    for r in empty:
        lines.append(f"  · {r.marketplace:12} 0 rows ({', '.join(r.input_files)})")
    for r in errors:
        lines.append(f"  ✗ {r.marketplace:12} ERROR — {r.error}")
    lines.append("-" * 60)
    lines.append(
        f"  {len(ok)} ok · {len(empty)} empty · {len(errors)} error · "
        f"{len(nofiles)} no-files"
    )
    tot_items = sum(r.rows for r in ok)
    tot_qty = sum(r.qty for r in ok)
    lines.append(f"  TOTAL: {tot_items} items, {tot_qty} qty across "
                 f"{len(ok)} batch(es)")
    lines.append("═" * 60)
    return "\n".join(lines)
