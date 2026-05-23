#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
MT Select (Health & Glow) Processor  —  v3.0
================================================================================

PURPOSE
-------
Converts Health & Glow (H&G) marketplace PO CSV files into a D365-ready Excel
workbook containing Sales Order Headers and Lines, matching the format used by
the Online PO Management system (6 sheets: Headers, Lines, Summary, Validation,
Warnings, Raw Data).

WORKFLOW OVERVIEW
-----------------
1.  App starts → auto-loads all three master files from  data_mt/  folder.
2.  A background file-watcher thread monitors  data_mt/  every 30 seconds.
    If any master file is updated on disk (newer mtime), it is silently
    reloaded without user intervention — no restart required.
3.  User selects one or more H&G PO CSV files via Browse button.
4.  User selects warehouse (AHD / BLR) and optionally enables Tester mode.
5.  User clicks ▶ Generate SO.
6.  Each unique PO gets one Sales Order number:
      Regular:  SO/HG/MM/SEQUENCE   (e.g. SO/HG/05/23526)
      Tester:   SO/TT/MM/SEQUENCE   (e.g. SO/TT/05/101)
    HG and TT sequences are stored SEPARATELY in data_mt/mt_select_seq.json.
7.  Output Excel is written to  output_mt/  subfolder inside the CSV folder.
8.  Full run log is written to  Logs/  subfolder next to the script.

FOLDER STRUCTURE
----------------
  <script_folder>/
  ├── mt_select_hg_processor.py       ← this script
  ├── data_mt/                        ← master files + sequence tracker
  │   ├── HG Master*.xlsx             ← SKU → EAN mapping
  │   ├── Items*.xlsx                 ← EAN → D365 Item No, MRP, GST
  │   ├── H&G Addresses*.xlsx         ← Store Name → Ship-to / Cust No
  │   └── mt_select_seq.json          ← persisted HG + TT sequence counters
  ├── Logs/                           ← one timestamped .log file per run
  └── <your_csv_folder>/
      ├── YourPO.csv                  ← input files (selected via Browse)
      └── output_mt/                  ← generated Excel output lands here

MASTER FILE FORMATS
-------------------
HG Master (SKU → EAN):
  • Excel (.xlsx), header auto-detected (scans first 6 rows for 'sku_code')
  • Required columns: sku_code  (or 'sku', 'SKU CODE')
                      ENN code  (or 'EAN', 'GTIN', 'ENN')
  • Duplicate 'ENN code' columns handled — first occurrence is used
  • Rows with blank EAN are skipped with a WARNING log entry

Items Master (EAN → D365 Item):
  • Excel (.xlsx), header at row 1
  • Required columns: GTIN, No., Description, Mrp, GST Group Code

Address Master (Store → Ship-to):
  • Excel (.xlsx), sheet name must be 'Ship-to Address List'
  • Required columns: Name (store display name), Code (ship-to = cust no)

CSV INPUT FORMAT
----------------
Required columns (case-sensitive): PO_NO, STORE_NAME, SKU_CODE, QUANTITY, MRP
All other columns are ignored.
Duplicate (PO_NO, SKU_CODE) pairs across files are merged (quantities summed).

OUTPUT FORMAT  —  6 Excel sheets
---------------------------------
  Sheet 1 — Headers (SO)   : One row per PO / Sales Order header
  Sheet 2 — Lines (SO)     : One row per SKU line; line numbers reset per PO
  Sheet 3 — Summary        : One row per PO with totals and OK/WARN status
  Sheet 4 — Validation     : Full line detail for inspection (EAN, MRP, tester)
  Sheet 5 — Warnings       : All mapping issues in one place
  Sheet 6 — Raw Data       : Audit trail of every processed line

SEQUENCE LOGIC
--------------
  Regular SO:  SO/HG/MM/NNNNN  — HG counter increments by 1 per PO
  Tester SO:   SO/TT/MM/NNN    — TT counter increments by 1 per PO
  Counters are INDEPENDENT — tester runs never affect regular numbering.
  Both counters survive restarts via data_mt/mt_select_seq.json.
  Month (MM) is always the current calendar month — auto-updates at midnight.

TESTER MODE
-----------
  • All SKU quantities forced to 1 regardless of PO demand.
  • Unit Price set to 0.54 (TESTER_CP constant) for all lines.
  • SO prefix changes from HG to TT.
  • TT sequence increments; HG sequence is untouched.

AUTO-RELOAD (background watcher)
---------------------------------
  A daemon thread polls data_mt/ every MASTER_WATCH_INTERVAL_SEC seconds.
  When it detects that a master file's last-modified timestamp is newer than
  when it was last loaded, it reloads that file silently and updates the GUI
  label.  No popup, no interruption — the GUI stays responsive because all
  actual reload work is done in the background thread via root.after() calls
  back onto the main Tkinter thread.

DEPENDENCIES
------------
  pip install pandas openpyxl

AUTHOR / MAINTAINER
-------------------
  Order Management Automation Team
  Version:  3.0
  Updated:  2026-05-23
  Expiry:   2026-06-30  (hard-coded; script refuses to run after this date)

CHANGE LOG
----------
  v1.0  Initial release — basic CSV → D365 Excel conversion
  v2.0  Added 6-sheet output, separate HG/TT sequences, logging, data_mt folder
  v3.0  Added background auto-reload watcher, full inline documentation,
        improved EAN deduplication, SKU column kept on SORow for Raw Data sheet
================================================================================
"""

# ── Standard library ──────────────────────────────────────────────────────────
import os           # os.path.basename, os.startfile (Windows open-file)
import sys          # sys.exit on expiry, sys.stdout for console logging
import json         # sequence file read/write
import time         # elapsed time measurement, watcher sleep
import logging      # structured log output to file + GUI + console
import threading    # background master-file watcher daemon thread
import tkinter as tk                        # main GUI framework
from tkinter import ttk, filedialog, messagebox  # themed widgets, file dialogs, popups
from pathlib import Path                    # cross-platform path manipulation
from datetime import datetime               # current date/time for SO month, timestamps
from typing import List, Dict, Optional, Tuple  # type hints for IDE support
from dataclasses import dataclass, field    # clean data model definitions

# ── Third-party (must be pip-installed) ───────────────────────────────────────
import pandas as pd                         # CSV / Excel reading, DataFrame operations
from openpyxl import Workbook               # Excel workbook creation
from openpyxl.styles import (               # cell styling
    Font, PatternFill, Alignment, Border, Side
)


# ==============================================================================
# SECTION 1 — FOLDER PATHS  (all derived from script location)
# ==============================================================================
# Using Path(__file__).parent ensures paths work regardless of where the user
# launches the script from (double-click, terminal, scheduled task, etc.).

SCRIPT_DIR  = Path(__file__).parent
# SCRIPT_DIR — the folder containing this .py file.  All other paths branch
# from here so the entire tool is self-contained and portable.

DATA_MT_DIR = SCRIPT_DIR / "data_mt"
# DATA_MT_DIR — holds all three master Excel files and the sequence JSON.
# The user never needs to touch this folder directly; auto-load handles it.
# Naming convention "data_mt" keeps H&G master data separate from other tools.

OUTPUT_MT   = SCRIPT_DIR / "output_mt"
# OUTPUT_MT — fallback output folder used only if no CSV has been loaded yet
# (e.g. when testing the writer directly).  In normal usage the output goes to
# <csv_folder>/output_mt/ so the Excel sits next to its source CSVs.

LOG_DIR = SCRIPT_DIR / "Logs"
# LOG_DIR — one timestamped .log file is written here per run.
# Logs are never deleted automatically — archive or clean manually.

# Create all folders immediately at import time so later code can assume they
# exist without individual guards.
DATA_MT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_MT.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)


# ==============================================================================
# SECTION 2 — LOGGING  (file + console + GUI handler added later)
# ==============================================================================
# Log file name includes a timestamp so each run has its own file and old logs
# are never overwritten.  This is critical for post-run debugging.

LOG_FILE = LOG_DIR / f"mt_select_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

# File handler — DEBUG and above written to disk in UTC+local time.
# UTF-8 encoding handles Unicode store names and product descriptions.
file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(
    logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
)

# Console handler — mirrors file output to stdout so developers running from
# a terminal can see live progress without opening the log file.
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.DEBUG)
console_handler.setFormatter(
    logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
)

# Named logger "mt_select" — avoids polluting the root logger, which could
# interfere with pandas / openpyxl's own logging if they share the root.
log = logging.getLogger("mt_select")
log.setLevel(logging.DEBUG)          # capture everything; handlers filter level
log.addHandler(file_handler)
log.addHandler(console_handler)
# Note: a GuiLogHandler is added AFTER the Tkinter Text widget is built —
# see MTSelectApp.__init__() for that attachment.


# ==============================================================================
# SECTION 3 — CONFIGURATION CONSTANTS
# ==============================================================================

EXPIRY_DATE = "30-06-2026"
# Hard expiry date (DD-MM-YYYY).  The script checks this at startup and refuses
# to run if today's date is past it.  Update this before deploying a new build.

WAREHOUSES: Dict[str, str] = {
    "AHD": "PICK",       # Ahmedabad warehouse → D365 Location Code "PICK"
    "BLR": "DS_BL_OFF1", # Bangalore warehouse → D365 Location Code "DS_BL_OFF1"
}
# WAREHOUSES maps the user-friendly dropdown label to the exact D365 Location
# Code that must appear in both the Header and Line sheets.  Add new warehouses
# here without touching any other code.

DEFAULT_WAREHOUSE = "AHD"
# Pre-selected warehouse when the app opens.  Change if AHD is rarely used.

TESTER_CP: float = 0.54
# Cost Price used for ALL tester order lines regardless of product MRP.
# This is a business rule — testers are always invoiced at ₹0.54.

MASTER_WATCH_INTERVAL_SEC: int = 30
# How often (in seconds) the background watcher thread checks data_mt/ for
# updated master files.  30 seconds balances responsiveness vs CPU overhead.
# Increase to 60+ in production if the folder is on a network share.

SEQ_FILE = DATA_MT_DIR / "mt_select_seq.json"
# Path to the JSON file that persists SO sequence counters across restarts.
# Stored inside data_mt/ so it travels with the master files when the folder
# is copied to another machine.
# Format: {"HG": 23548, "TT": 118}


# ==============================================================================
# SECTION 4 — SEQUENCE MANAGEMENT
# ==============================================================================
# HG (regular) and TT (tester) sequences are stored INDEPENDENTLY.
# This prevents tester runs from consuming HG sequence numbers and keeps
# audit trails clean — you can always tell from the SO number whether it
# was a real order or a tester run.

def load_sequences() -> Dict[str, int]:
    """
    Read the persisted sequence counters from SEQ_FILE.

    Returns a dict {"HG": int, "TT": int}.
    If the file does not exist or is corrupt, returns safe defaults so the
    app can still run — the user will see a log WARNING but no crash.

    Backward compatibility: the v1 format stored {"last_sequence": N}.
    That is detected and migrated to the new format automatically.
    """
    # Safe defaults used when no sequence file exists yet (first run).
    defaults: Dict[str, int] = {"HG": 23525, "TT": 100}

    if not SEQ_FILE.exists():
        # First run or data_mt was wiped — start from defaults.
        log.debug(f"[Sequence] File not found at {SEQ_FILE}, using defaults: {defaults}")
        return defaults

    try:
        with open(SEQ_FILE, 'r') as f:
            data = json.load(f)   # raises json.JSONDecodeError if file is corrupt

        if isinstance(data, dict) and "HG" in data:
            # New format — both keys present, read them directly.
            result = {
                "HG": int(data.get("HG", defaults["HG"])),
                "TT": int(data.get("TT", defaults["TT"]))
            }
        else:
            # Old v1 format — migrate: HG takes the old single sequence,
            # TT starts at its default because it didn't exist before.
            old_seq = int(data.get("last_sequence", defaults["HG"]))
            result = {"HG": old_seq, "TT": defaults["TT"]}
            log.info(f"[Sequence] Migrated old format → HG={old_seq}, TT={defaults['TT']}")

        log.debug(f"[Sequence] Loaded: HG={result['HG']}, TT={result['TT']}")
        return result

    except Exception as e:
        # Corrupt JSON, permission error, etc. — log and fall back gracefully.
        log.warning(f"[Sequence] Failed to read {SEQ_FILE}: {e} — using defaults")
        return defaults


def save_sequences(seqs: Dict[str, int]) -> None:
    """
    Persist the current HG and TT counters to SEQ_FILE.

    Called immediately after each successful generation run so that even if
    the app is force-killed before the window is closed, the last sequence is
    already saved and duplicate SOs cannot be issued on the next run.

    Args:
        seqs:  dict with keys "HG" and "TT", both int.
    """
    with open(SEQ_FILE, 'w') as f:
        json.dump(seqs, f, indent=2)   # indent=2 makes the file human-readable
    log.debug(f"[Sequence] Saved: {seqs}")


def generate_so_number(seq: int, is_tester: bool = False) -> str:
    """
    Build a formatted Sales Order number string.

    Format:  SO/{prefix}/{MM}/{sequence}
    Examples:
        Regular:  SO/HG/05/23526
        Tester:   SO/TT/05/101

    The month component (MM) is derived from today's date at call time so that
    a batch run spanning midnight automatically uses the correct month for each
    SO without any user intervention.

    Args:
        seq:        The integer sequence number (e.g. 23526).
        is_tester:  If True, prefix is "TT"; otherwise "HG".

    Returns:
        Formatted SO number string.
    """
    month  = datetime.now().strftime("%m")   # zero-padded month, e.g. "05"
    prefix = "TT" if is_tester else "HG"    # tester vs regular marketplace code
    return f"SO/{prefix}/{month}/{seq}"


# ==============================================================================
# SECTION 5 — DATA MODELS
# ==============================================================================
# Using @dataclass keeps the data models concise and self-documenting.
# All fields have explicit types so IDE type checkers catch misuse early.

@dataclass
class SORow:
    """
    Represents one processed Sales Order line (one SKU within one PO).

    Each SORow maps to exactly one row in the Lines (SO) sheet and one row
    in the Validation and Raw Data sheets.  Multiple SORows share the same
    so_number when they belong to the same PO.

    Fields
    ------
    po_number   : Original PO number from the CSV (e.g. "6449964").
    so_number   : Generated SO number (e.g. "SO/HG/05/23526").
    sku_code    : Original SKU code from the CSV — kept for Raw Data audit.
    item_no     : D365 Item Number resolved via EAN → Items Master lookup.
                  "?MISSING_SKU" if SKU not found in HG Master.
                  "?MISSING_EAN" if EAN not found in Items Master.
    qty         : Final quantity after tester override (1 for tester, actual for regular).
    store_name  : Store/location name from CSV STORE_NAME column.
    ship_to     : Ship-to Code from Address Master (e.g. "HG-MUM-001").
    cust_no     : Customer Number from Address Master (same value as ship_to for H&G).
    ean         : EAN barcode resolved via SKU → HG Master lookup.
    description : Product description from Items Master.
    mrp         : Maximum Retail Price from Items Master (0.0 if missing).
    unit_price  : 0.54 for tester orders; empty string in output for regular.
    line_no     : Line number within the SO (10000, 20000, …); reset per PO.
    is_tester   : True when generated in Tester mode.
    status      : "OK" if all lookups succeeded; "WARN" if any mapping failed.
    """
    po_number:   str
    so_number:   str
    sku_code:    str          # kept for full audit trail in Raw Data sheet
    item_no:     str
    qty:         int
    store_name:  str
    ship_to:     str
    cust_no:     str
    ean:         str
    description: str
    mrp:         float = 0.0
    unit_price:  float = 0.0
    line_no:     int   = 0
    is_tester:   bool  = False
    status:      str   = "OK"


@dataclass
class ProcessingResult:
    """
    Container for everything produced by a single process_csv_files() call.

    This object is created once per Generate button click and passed to the
    writer.  Keeping all outputs here (rather than globals) makes the
    processing engine stateless and easy to unit-test.

    Fields
    ------
    rows          : Ordered list of SORow objects — one per SKU per PO.
    warnings      : List of (po_number, sku_code, message) tuples for the
                    Warnings sheet.  Non-fatal — processing continues.
    input_files   : List of CSV file paths that were processed (for Raw Data).
    warehouse_code: D365 Location Code (e.g. "PICK") used on all lines.
    warehouse_display: User-facing warehouse name (e.g. "AHD").
    is_tester     : Whether this run was in tester mode.
    hg_sequence   : The HG counter value AFTER this run (saved to JSON).
    tt_sequence   : The TT counter value AFTER this run (saved to JSON).
    so_map        : dict mapping po_number → so_number for quick lookup.
    po_summary    : dict mapping po_number → summary dict used for Summary sheet.
                    Each summary dict has keys: so_number, store, ship_to,
                    cust_no, items (count), total_qty, status.
    """
    rows:              List[SORow]          = field(default_factory=list)
    warnings:          List[Tuple[str, str, str]] = field(default_factory=list)
    input_files:       List[str]            = field(default_factory=list)
    warehouse_code:    str                  = "PICK"
    warehouse_display: str                  = "AHD"
    is_tester:         bool                 = False
    hg_sequence:       int                  = 0
    tt_sequence:       int                  = 0
    so_map:            Dict[str, str]       = field(default_factory=dict)
    po_summary:        Dict[str, dict]      = field(default_factory=dict)


# ==============================================================================
# SECTION 6 — MASTER FILE LOADERS
# ==============================================================================
# Each loader class is responsible for exactly one master file.
# They expose a .load(path) method that returns the count of records loaded
# and populates an internal dict used by the processing engine.
# Keeping them as classes (not plain functions) lets the MTSelectApp hold
# references and the watcher thread reload them in place.

class HGMasterLoader:
    """
    Loads the SKU → EAN mapping from the HG Master Excel file.

    Handles:
    - Auto-detecting the header row (scans first 6 rows for 'sku_code').
    - Duplicate column names: your file has 'ENN code' twice; pandas renames
      the second one to 'ENN code.1'. We always pick the FIRST occurrence.
    - Float EANs: Excel sometimes stores 8906121641769 as 8906121641769.0.
      These are converted to clean integer strings.
    - Blank EAN rows (rows 141-147 in your master): skipped with WARNING.

    Attributes:
        sku_to_ean  (dict):   Maps str(sku_code) → str(ean).
        source_path (Path):   Path of the last successfully loaded file.
        last_mtime  (float):  os.path.getmtime at load time — used by watcher.
    """

    def __init__(self):
        self.sku_to_ean:  Dict[str, str] = {}
        self.source_path: Optional[Path] = None
        self.last_mtime:  float          = 0.0   # epoch seconds; 0 means never loaded

    def load(self, path: Path) -> int:
        """
        Load SKU→EAN mappings from the given Excel file.

        Args:
            path:  Absolute Path to the HG Master .xlsx file.

        Returns:
            Number of valid SKU→EAN pairs loaded.

        Raises:
            FileNotFoundError: if path does not exist.
            ValueError:        if SKU or EAN column cannot be found, or if
                               the file is unreadable.
        """
        log.info(f"[HG Master] Loading from: {path}")

        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        self.source_path = path
        self.last_mtime  = os.path.getmtime(path)  # record mtime BEFORE reading

        # ── Step 1: Auto-detect header row ───────────────────────────────────
        # We read the first 6 rows WITHOUT a header (header=None) so we can
        # inspect raw cell values.  We look for a row where any cell equals
        # 'sku_code', 'sku code', or 'sku' (case-insensitive).  This survives
        # files where someone inserted a title row at the top.
        try:
            df_scan = pd.read_excel(path, header=None, nrows=6)
        except Exception as e:
            raise ValueError(f"Cannot read Excel for header scan: {e}")

        header_row: Optional[int] = None
        for i, row in df_scan.iterrows():
            # Normalise each cell to lowercase stripped string for comparison.
            row_vals = [str(v).strip().lower() for v in row.values]
            if any(v in ('sku_code', 'sku code', 'sku') for v in row_vals):
                header_row = i   # found the row index (0-based)
                log.info(f"[HG Master] Header auto-detected at row index {i} (Excel row {i+1})")
                break

        if header_row is None:
            # No recognisable header found in first 6 rows — fall back to row 0
            # and let the column-finding logic below raise a clearer error.
            log.warning("[HG Master] Could not auto-detect header row — defaulting to index 0")
            header_row = 0

        # ── Step 2: Re-read with correct header row ───────────────────────────
        try:
            df = pd.read_excel(path, header=header_row)
        except Exception as e:
            raise ValueError(f"Cannot read Excel with header={header_row}: {e}")

        # Strip leading/trailing whitespace from every column name so that
        # '  sku_code  ' still matches 'sku_code'.
        df.columns = [str(c).strip() for c in df.columns]
        log.debug(f"[HG Master] Columns after strip: {list(df.columns)}")
        log.debug(f"[HG Master] Total data rows: {len(df)}")

        # ── Step 3: Locate SKU column ─────────────────────────────────────────
        sku_col: Optional[str] = None
        for col in df.columns:
            if col.lower() in ('sku_code', 'sku code', 'sku'):
                sku_col = col
                log.info(f"[HG Master] SKU column identified: '{sku_col}'")
                break
        if sku_col is None:
            raise ValueError(
                f"No SKU column found in HG Master.\n"
                f"Available columns: {list(df.columns)}\n"
                f"Expected one of: sku_code, sku code, sku"
            )

        # ── Step 4: Locate EAN column ─────────────────────────────────────────
        # YOUR FILE has two columns both named 'ENN code'.  Pandas auto-renames
        # the second to 'ENN code.1'.  We iterate in column order and take the
        # FIRST match so we always get the correct EAN (column C, not G).
        # We strip '.1', '.2' etc. suffixes before comparing so the duplicate
        # detection works regardless of how many duplicates exist.
        ean_col: Optional[str] = None
        for col in df.columns:
            # Remove pandas duplicate suffix (e.g. '.1', '.2') then compare.
            col_norm = col.lower().split('.')[0].strip()
            if col_norm in ('enn code', 'enn', 'ean', 'gtin', 'ean code'):
                ean_col = col
                log.info(f"[HG Master] EAN column identified: '{ean_col}' (first match wins)")
                break
        if ean_col is None:
            raise ValueError(
                f"No EAN column found in HG Master.\n"
                f"Available columns: {list(df.columns)}\n"
                f"Expected one of: ENN code, EAN, GTIN, ENN"
            )

        # ── Step 5: Build SKU → EAN dict ─────────────────────────────────────
        self.sku_to_ean.clear()   # reset in case this is a reload
        loaded        = 0   # count of valid mappings added
        skipped_blank = 0   # rows where EAN is empty/NaN (known gap in master)
        skipped_sku   = 0   # rows where SKU itself is empty

        for idx, row in df.iterrows():
            raw_sku = row[sku_col]
            raw_ean = row[ean_col]

            # ── SKU validation ────────────────────────────────────────────────
            if pd.isna(raw_sku):
                # Completely empty row (e.g. trailing blank rows in Excel).
                skipped_sku += 1
                continue

            sku = str(raw_sku).strip()
            if not sku or sku.lower() == 'nan':
                # Cell contained a literal "nan" string or just spaces.
                skipped_sku += 1
                continue

            # ── EAN validation ────────────────────────────────────────────────
            # Rows 141-147 in your master have blank EAN — do NOT crash,
            # just skip with a WARNING so it appears in the run log.
            if pd.isna(raw_ean) or str(raw_ean).strip().lower() in ('', 'nan'):
                skipped_blank += 1
                log.warning(
                    f"[HG Master] Row {idx + 2} (Excel): "
                    f"SKU '{sku}' has blank EAN — skipped (add EAN to master to fix)"
                )
                continue

            # ── EAN type normalisation ────────────────────────────────────────
            # Excel often stores 13-digit GTINs as floats (e.g. 8906121641769.0).
            # We convert:
            #   float  → int → str  (removes the .0)
            #   int    → str        (already clean)
            #   str    → strip + remove trailing .0 if present
            if isinstance(raw_ean, float):
                # Check for float NaN that slipped through the isna() check
                # (can happen with some openpyxl/xlrd versions).
                if raw_ean != raw_ean:   # NaN is the only float that != itself
                    skipped_blank += 1
                    log.warning(f"[HG Master] Row {idx+2}: SKU '{sku}' EAN is float NaN — skipped")
                    continue
                ean = str(int(raw_ean)) if raw_ean == int(raw_ean) else str(raw_ean)
            elif isinstance(raw_ean, int):
                ean = str(raw_ean)
            else:
                ean = str(raw_ean).strip()
                if ean.endswith('.0'):
                    ean = ean[:-2]   # strip the ".0" suffix from string form

            # Final check — ensure we didn't end up with an empty/nan string.
            if not ean or ean.lower() == 'nan':
                skipped_blank += 1
                log.warning(f"[HG Master] Row {idx+2}: SKU '{sku}' EAN resolved to empty — skipped")
                continue

            # ── Store mapping ─────────────────────────────────────────────────
            self.sku_to_ean[sku] = ean
            loaded += 1

        # ── Summary log ───────────────────────────────────────────────────────
        log.info(
            f"[HG Master] Load complete — "
            f"Loaded: {loaded} | Blank EAN skipped: {skipped_blank} | "
            f"Blank SKU skipped: {skipped_sku}"
        )

        if loaded == 0:
            # Nothing loaded at all — raise so the GUI shows an error popup.
            raise ValueError(
                f"No valid SKU→EAN mappings found in HG Master.\n"
                f"SKU column used: '{sku_col}'\n"
                f"EAN column used: '{ean_col}'\n"
                f"All columns found: {list(df.columns)}\n"
                f"Check that header row {header_row+1} is correct and data starts below it."
            )

        # Log a 5-item sample for quick sanity check in the run log.
        sample = list(self.sku_to_ean.items())[:5]
        log.debug(f"[HG Master] Sample mappings (first 5): {sample}")

        return loaded


class ItemsMasterLoader:
    """
    Loads the EAN → D365 Item details mapping from the Items Master Excel file.

    Expected columns (header at row 1):
        GTIN            — 13-digit EAN barcode (used as key)
        No.             — D365 Item Number
        Description     — Product description text
        Mrp             — Maximum Retail Price (float)
        GST Group Code  — e.g. "G-18-S" (used for validation, not D365 import)

    Attributes:
        ean_to_item  (dict):  Maps str(EAN) → {"item_no", "description",
                                                "mrp", "gst_code"}.
        source_path  (Path):  Path of the last loaded file.
        last_mtime   (float): File modification time at last load.
    """

    def __init__(self):
        self.ean_to_item:  Dict[str, Dict] = {}
        self.source_path:  Optional[Path]  = None
        self.last_mtime:   float           = 0.0

    def load(self, path: Path) -> int:
        """
        Load EAN→Item mappings from the given Excel file.

        Returns count of EAN mappings loaded.
        Raises FileNotFoundError or ValueError on failure.
        """
        log.info(f"[Items Master] Loading from: {path}")

        if not path.exists():
            raise FileNotFoundError(f"Items Master not found: {path}")

        self.source_path = path
        self.last_mtime  = os.path.getmtime(path)

        df = pd.read_excel(path, header=0)   # header at first row
        log.debug(f"[Items Master] Columns: {list(df.columns)}, Rows: {len(df)}")

        # Validate required columns before processing.
        for required_col in ('GTIN', 'No.'):
            if required_col not in df.columns:
                raise ValueError(
                    f"Items Master is missing required column '{required_col}'.\n"
                    f"Available columns: {list(df.columns)}"
                )

        # Convert GTIN (which Excel may store as float) to clean string.
        # .str.replace removes trailing '.0' from float-derived strings.
        df['GTIN_str'] = (
            df['GTIN']
            .astype(str)
            .str.strip()
            .str.replace(r'\.0$', '', regex=True)
        )

        self.ean_to_item.clear()   # reset for reload support

        for _, row in df.iterrows():
            ean     = row['GTIN_str']
            item_no = str(row['No.']).strip()

            # Description may be missing for some items — default to empty string.
            desc = (
                str(row.get('Description', ''))
                if pd.notna(row.get('Description'))
                else ''
            )

            # MRP — float, default 0.0 if missing.
            mrp_raw = row.get('Mrp')
            mrp = float(mrp_raw) if pd.notna(mrp_raw) else 0.0

            # GST Group Code — e.g. "G-18-S"; default empty if missing.
            gst_raw = row.get('GST Group Code')
            gst = str(gst_raw).strip() if pd.notna(gst_raw) else ''

            self.ean_to_item[ean] = {
                'item_no':     item_no,
                'description': desc,
                'mrp':         mrp,
                'gst_code':    gst,
            }

        log.info(f"[Items Master] Loaded {len(self.ean_to_item)} EAN→Item mappings")
        log.debug(f"[Items Master] Sample (first 3): {list(self.ean_to_item.items())[:3]}")
        return len(self.ean_to_item)


class AddressMasterLoader:
    """
    Loads the Store Name → Ship-to Code / Customer Number mapping from
    the H&G Addresses Excel file.

    Expected sheet name: 'Ship-to Address List'  (exact match required)
    Expected columns:
        Name  — store display name matching STORE_NAME in the CSV
        Code  — ship-to code (also used as Customer No. for H&G)

    Attributes:
        store_to_ship (dict):  Maps str(store_name) → {"ship_to", "cust_no"}.
        source_path   (Path):  Path of the last loaded file.
        last_mtime    (float): File modification time at last load.
    """

    def __init__(self):
        self.store_to_ship: Dict[str, Dict] = {}
        self.source_path:   Optional[Path]  = None
        self.last_mtime:    float           = 0.0

    def load(self, path: Path) -> int:
        """
        Load Store→Ship-to mappings from the given Excel file.

        Returns count of store mappings loaded.
        Raises FileNotFoundError or ValueError on failure.
        """
        log.info(f"[Address Master] Loading from: {path}")

        if not path.exists():
            raise FileNotFoundError(f"Address Master not found: {path}")

        self.source_path = path
        self.last_mtime  = os.path.getmtime(path)

        # sheet_name must match exactly — 'Ship-to Address List'.
        # If the sheet is renamed, this raises a ValueError with a clear message.
        try:
            df = pd.read_excel(path, sheet_name='Ship-to Address List', header=0)
        except Exception as e:
            raise ValueError(
                f"Could not read sheet 'Ship-to Address List' from {path.name}.\n"
                f"Error: {e}\n"
                f"Check that the sheet exists with this exact name."
            )

        log.debug(f"[Address Master] Columns: {list(df.columns)}, Rows: {len(df)}")

        # Find 'Name' and 'Code' columns (case-insensitive).
        store_col = next((c for c in df.columns if c.lower() == 'name'), None)
        ship_col  = next((c for c in df.columns if c.lower() == 'code'), None)

        if not store_col or not ship_col:
            raise ValueError(
                f"Address Master: Missing 'Name' or 'Code' column.\n"
                f"Available columns: {list(df.columns)}"
            )

        self.store_to_ship.clear()   # reset for reload support

        for _, row in df.iterrows():
            store   = str(row[store_col]).strip()
            ship_to = str(row[ship_col]).strip() if pd.notna(row[ship_col]) else ''

            # Skip empty rows or rows without a ship-to code.
            if not store or store == 'nan' or not ship_to:
                continue

            # For H&G the Customer No. is identical to the Ship-to Code.
            self.store_to_ship[store] = {
                'ship_to': ship_to,
                'cust_no': ship_to,
            }

        log.info(f"[Address Master] Loaded {len(self.store_to_ship)} store→ship-to mappings")
        return len(self.store_to_ship)


# ==============================================================================
# SECTION 7 — BACKGROUND MASTER-FILE WATCHER
# ==============================================================================
# The watcher runs as a daemon thread so it is automatically killed when the
# main GUI window closes.  It polls data_mt/ every MASTER_WATCH_INTERVAL_SEC
# seconds.  When a file's mtime on disk is newer than .last_mtime on the loader
# object, it reloads that file and schedules a GUI label update via root.after()
# (which is thread-safe; direct Tkinter calls from background threads are not).

class MasterWatcher(threading.Thread):
    """
    Background daemon thread that monitors data_mt/ for master file changes
    and auto-reloads them without requiring user interaction.

    Args:
        app:  Reference to the running MTSelectApp instance.
              Used to access loader objects and schedule GUI updates.

    Design notes:
    - Uses threading.Event for a clean shutdown signal (set by stop()).
    - All Tkinter GUI mutations are scheduled via self.app.root.after(0, fn)
      to ensure they run on the main thread (Tkinter is not thread-safe).
    - A short initial sleep (2 s) lets the GUI finish initialising before
      the first watch cycle, avoiding a race condition on startup.
    """

    def __init__(self, app: 'MTSelectApp'):
        super().__init__(daemon=True)   # daemon=True: thread dies with the process
        self.app   = app
        self._stop = threading.Event()  # set this to request clean shutdown

    def stop(self):
        """Signal the thread to exit on its next iteration."""
        self._stop.set()

    def run(self):
        """
        Main loop — runs until stop() is called or the process exits.

        Each iteration checks all three master files.  If any is newer than
        its last load time, it is reloaded.  The GUI is updated via after().
        """
        log.debug("[Watcher] Master file watcher started")
        time.sleep(2)   # brief startup delay so GUI is fully built

        while not self._stop.is_set():
            self._check_and_reload()
            # Wait for the interval, but wake immediately if stop() is called.
            self._stop.wait(timeout=MASTER_WATCH_INTERVAL_SEC)

        log.debug("[Watcher] Master file watcher stopped")

    def _check_and_reload(self):
        """
        Check each master file and reload if its on-disk mtime has advanced.

        Three separate checks — each loader is independent so a change in
        the Items Master doesn't force a reload of the HG Master, etc.
        """
        self._check_loader(
            loader     = self.app.hg_master,
            glob_pat   = "HG Master*.xlsx",
            label_var  = self.app.hg_path_var,
            label_fmt  = lambda f, n: f"{f.name} ({n} SKUs) ✓ (auto-reloaded)",
            log_prefix = "HG Master"
        )
        self._check_loader(
            loader     = self.app.items_master,
            glob_pat   = "Items*.xlsx",
            label_var  = self.app.items_path_var,
            label_fmt  = lambda f, n: f"{f.name} ({n} EANs) ✓ (auto-reloaded)",
            log_prefix = "Items Master"
        )
        self._check_loader(
            loader     = self.app.address_master,
            glob_pat   = "H&G Addresses*.xlsx",
            label_var  = self.app.address_path_var,
            label_fmt  = lambda f, n: f"{f.name} ({n} stores) ✓ (auto-reloaded)",
            log_prefix = "Address Master"
        )

    def _check_loader(self, loader, glob_pat: str, label_var,
                      label_fmt, log_prefix: str):
        """
        Check a single loader's file for a newer mtime and reload if needed.

        Args:
            loader:     One of HGMasterLoader / ItemsMasterLoader / AddressMasterLoader.
            glob_pat:   Glob pattern to find the file inside DATA_MT_DIR.
            label_var:  tkinter StringVar to update on reload.
            label_fmt:  Callable(file_path, count) → label string.
            log_prefix: Human-readable name for log messages.
        """
        # Find the most-recently named matching file in data_mt/.
        candidates = sorted(DATA_MT_DIR.glob(glob_pat), reverse=True)
        if not candidates:
            return   # no matching file — nothing to do

        file_path = candidates[0]   # use the alphabetically last (usually newest)

        try:
            current_mtime = os.path.getmtime(file_path)
        except OSError:
            return   # file may have been deleted mid-check — skip safely

        # Compare on-disk mtime with what the loader recorded when it last loaded.
        if current_mtime <= loader.last_mtime:
            return   # file unchanged — no reload needed

        # File is newer — reload it.
        log.info(f"[Watcher] {log_prefix} changed on disk — auto-reloading: {file_path.name}")
        try:
            count = loader.load(file_path)
            new_label = label_fmt(file_path, count)
            log.info(f"[Watcher] {log_prefix} reloaded: {count} records")

            # Schedule GUI update on the main thread (thread-safe via after()).
            def _update_gui(lv=label_var, nl=new_label):
                lv.set(nl)
                self.app._update_master_status()   # refresh the "Masters: Ready ✓" bar

            self.app.root.after(0, _update_gui)

        except Exception as e:
            log.error(f"[Watcher] {log_prefix} auto-reload failed: {e}")


# ==============================================================================
# SECTION 8 — CSV PROCESSING ENGINE
# ==============================================================================

def process_csv_files(
    file_paths:      List[str],
    hg_master:       HGMasterLoader,
    items_master:    ItemsMasterLoader,
    address_master:  AddressMasterLoader,
    warehouse_code:  str,
    is_tester:       bool,
    sequences:       Dict[str, int]
) -> ProcessingResult:
    """
    Core processing engine — reads CSV files, resolves all lookups, and
    builds the ProcessingResult object ready for writing to Excel.

    Processing pipeline:
    1.  Read each CSV file into a DataFrame.
    2.  Validate required columns; skip file with WARNING if missing.
    3.  Collect all valid rows; detect duplicates (merge by summing qty).
    4.  Group rows by PO_NO.
    5.  For each PO: assign a new SO number (incrementing the appropriate
        sequence counter), resolve store → ship-to, then for each SKU resolve
        SKU → EAN → D365 Item No + description + MRP.
    6.  Apply tester overrides (qty=1, unit_price=0.54) if is_tester=True.
    7.  Build po_summary dict for the Summary sheet.
    8.  Return the populated ProcessingResult.

    Args:
        file_paths:     List of absolute paths to H&G CSV files.
        hg_master:      Loaded HGMasterLoader instance.
        items_master:   Loaded ItemsMasterLoader instance.
        address_master: Loaded AddressMasterLoader instance.
        warehouse_code: D365 Location Code (e.g. "PICK").
        is_tester:      True → tester mode (SO/TT, qty=1, price=0.54).
        sequences:      dict {"HG": int, "TT": int} — starting counters.
                        These are NOT mutated; updated copies are stored on
                        the result object instead.

    Returns:
        ProcessingResult populated with rows, warnings, summary, and
        updated sequence counters.
    """
    result = ProcessingResult()
    result.input_files    = file_paths
    result.warehouse_code = warehouse_code
    result.warehouse_display = next(
        (k for k, v in WAREHOUSES.items() if v == warehouse_code), "AHD"
    )
    result.is_tester  = is_tester
    result.hg_sequence = sequences["HG"]   # will be updated below as POs are processed
    result.tt_sequence = sequences["TT"]

    # Determine which sequence key to use for this run.
    seq_key = "TT" if is_tester else "HG"
    log.info(
        f"[Process] Starting. Files: {len(file_paths)}, "
        f"Warehouse: {warehouse_code}, Tester: {is_tester}"
    )
    log.info(f"[Process] Using {seq_key} sequence, starting at: {sequences[seq_key]}")

    # ── Phase 1: Read all CSV files into a single flat list ───────────────────
    all_rows: List[dict] = []
    seen_po_sku: set = set()   # tracks (po, sku) pairs for duplicate detection

    for fp in file_paths:
        log.info(f"[Process] Reading CSV: {os.path.basename(fp)}")

        try:
            df = pd.read_csv(fp)
        except Exception as e:
            msg = f"Cannot read {os.path.basename(fp)}: {e}"
            log.error(f"[Process] {msg}")
            result.warnings.append(("", "", msg))
            continue   # skip this file; try the next one

        log.debug(f"[Process] Columns: {list(df.columns)}, Rows: {len(df)}")

        # Validate required columns — all five must be present.
        required_cols = ['PO_NO', 'STORE_NAME', 'SKU_CODE', 'QUANTITY', 'MRP']
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            msg = (
                f"File {os.path.basename(fp)} missing required columns: {missing}.\n"
                f"Available columns: {list(df.columns)}"
            )
            log.error(f"[Process] {msg}")
            result.warnings.append(("", "", msg))
            continue   # skip this file entirely

        # Iterate each row and collect valid data.
        for _, row in df.iterrows():
            po    = str(row['PO_NO']).strip()
            store = str(row['STORE_NAME']).strip()
            sku   = str(row['SKU_CODE']).strip()

            # Parse QUANTITY — convert "2.0" floats to int, treat NaN as 0.
            try:
                qty = int(float(row['QUANTITY'])) if pd.notna(row['QUANTITY']) else 0
            except (ValueError, TypeError):
                qty = 0

            # Parse MRP — float, default 0.0 if missing.
            mrp_raw = row.get('MRP')
            mrp = float(mrp_raw) if pd.notna(mrp_raw) else 0.0

            # Skip rows with missing key fields or zero/negative quantity.
            if po == 'nan' or store == 'nan' or sku == 'nan' or qty <= 0:
                log.debug(
                    f"[Process] Skipping invalid row: PO={po}, STORE={store}, "
                    f"SKU={sku}, QTY={qty}"
                )
                continue

            # Detect duplicate (PO, SKU) pairs — will be merged by summing qty.
            key = (po, sku)
            if key in seen_po_sku:
                msg = (
                    f"Duplicate (PO={po}, SKU={sku}) found — "
                    f"quantities will be merged (summed)"
                )
                log.warning(f"[Process] {msg}")
                result.warnings.append((po, sku, msg))
            seen_po_sku.add(key)

            all_rows.append({'po': po, 'store': store, 'sku': sku,
                             'qty': qty, 'mrp': mrp})

    log.info(f"[Process] Total valid CSV rows collected: {len(all_rows)}")

    if not all_rows:
        msg = "No valid rows found in any of the selected CSV files"
        log.error(f"[Process] {msg}")
        result.warnings.append(("", "", msg))
        return result   # return early — writer will handle empty result

    # ── Phase 2: Group by PO → SKU, summing quantities ────────────────────────
    # po_groups structure:
    #   { po_number: { sku_code: {"qty": int, "store": str, "mrp": float} } }
    po_groups: Dict[str, Dict[str, dict]] = {}
    for r in all_rows:
        po, sku = r['po'], r['sku']
        if po not in po_groups:
            po_groups[po] = {}
        if sku not in po_groups[po]:
            po_groups[po][sku] = {'qty': 0, 'store': r['store'], 'mrp': r['mrp']}
        po_groups[po][sku]['qty'] += r['qty']   # sum duplicates

    log.info(f"[Process] Unique POs: {len(po_groups)}")

    # ── Phase 3: Assign SO numbers and resolve all lookups ───────────────────
    # current_seq starts from the persisted counter for this run's type (HG/TT).
    current_seq = sequences[seq_key]
    result.so_map     = {}
    result.po_summary = {}

    for po, sku_dict in po_groups.items():
        # Increment sequence BEFORE generating the SO number so the first SO
        # uses (stored_value + 1), not the stored_value itself.
        current_seq += 1
        so_number = generate_so_number(current_seq, is_tester)
        result.so_map[po] = so_number

        # Update the appropriate counter on the result object.
        if is_tester:
            result.tt_sequence = current_seq
        else:
            result.hg_sequence = current_seq

        log.debug(f"[Process] Assigned: PO={po} → SO={so_number}")

        # ── Resolve store → ship-to / cust-no ────────────────────────────────
        # All SKUs within a PO share the same store, so we read from the first SKU.
        store_name = next(iter(sku_dict.values()))['store']
        addr_info  = address_master.store_to_ship.get(store_name)

        if not addr_info:
            msg = (
                f"Store '{store_name}' not found in Address Master — "
                f"Ship-to and Cust No will be blank in output"
            )
            log.warning(f"[Process] PO={po}: {msg}")
            result.warnings.append((po, "", msg))
            ship_to = cust_no = ""   # blank but don't crash — output still useful
        else:
            ship_to = addr_info['ship_to']
            cust_no = addr_info['cust_no']

        # Counters for this PO's summary row.
        po_total_qty = 0
        po_items     = 0
        po_has_warn  = (ship_to == "")   # already warned if store not found

        # ── Resolve SKU → EAN → Item No for each line ────────────────────────
        for sku, details in sku_dict.items():
            row_status = "OK"   # assume clean; flip to WARN if any lookup fails

            # Lookup 1: SKU → EAN via HG Master
            ean = hg_master.sku_to_ean.get(sku)
            if not ean:
                msg = f"SKU '{sku}' not found in HG Master (no EAN mapping)"
                log.warning(f"[Process] PO={po}: {msg}")
                result.warnings.append((po, sku, msg))
                item_no = "?MISSING_SKU"
                description = ""
                item_mrp    = 0.0
                gst_code    = ""
                row_status  = "WARN"
                po_has_warn = True
            else:
                # Lookup 2: EAN → D365 Item via Items Master
                item_info = items_master.ean_to_item.get(ean)
                if not item_info:
                    msg = f"EAN '{ean}' (from SKU '{sku}') not found in Items Master"
                    log.warning(f"[Process] PO={po}: {msg}")
                    result.warnings.append((po, sku, msg))
                    item_no     = "?MISSING_EAN"
                    description = ""
                    item_mrp    = 0.0
                    gst_code    = ""
                    row_status  = "WARN"
                    po_has_warn = True
                else:
                    item_no     = item_info['item_no']
                    description = item_info.get('description', '')
                    item_mrp    = item_info.get('mrp', 0.0)
                    gst_code    = item_info.get('gst_code', '')
                    log.debug(
                        f"[Process] Resolved: SKU={sku} → EAN={ean} → "
                        f"Item={item_no} | MRP={item_mrp} | GST={gst_code}"
                    )

            # ── Apply tester overrides ────────────────────────────────────────
            if is_tester:
                final_qty  = 1          # always 1 unit per SKU in tester orders
                unit_price = TESTER_CP  # fixed cost price for all tester lines
            else:
                final_qty  = details['qty']
                unit_price = 0.0        # regular orders: price comes from D365

            po_total_qty += final_qty
            po_items     += 1

            # Build and append the SORow for this line.
            result.rows.append(SORow(
                po_number   = po,
                so_number   = so_number,
                sku_code    = sku,          # preserved for Raw Data audit sheet
                item_no     = item_no,
                qty         = final_qty,
                store_name  = store_name,
                ship_to     = ship_to,
                cust_no     = cust_no,
                ean         = ean if ean else "",
                description = description,
                mrp         = item_mrp,
                unit_price  = unit_price,
                is_tester   = is_tester,
                status      = row_status,
            ))

        # ── Build summary row for this PO ─────────────────────────────────────
        result.po_summary[po] = {
            'so_number': so_number,
            'store':     store_name,
            'ship_to':   ship_to,
            'cust_no':   cust_no,
            'items':     po_items,
            'total_qty': po_total_qty,
            'status':    "WARN" if po_has_warn else "OK",
        }

    log.info(
        f"[Process] Done. SO lines: {len(result.rows)} | "
        f"SOs: {len(result.so_map)} | Warnings: {len(result.warnings)}"
    )
    return result


# ==============================================================================
# SECTION 9 — EXCEL OUTPUT WRITER
# ==============================================================================
# Produces a 6-sheet workbook matching the Online PO Management system format.
# Styling uses openpyxl directly (not pandas to_excel) to get full control
# over header colours, column widths, freeze panes, and status fill colours.

# ── Shared style objects (created once, reused across all sheets) ─────────────
HDR_FILL    = PatternFill("solid", fgColor="1A237E")   # dark navy header background
HDR_FONT    = Font(bold=True, color="FFFFFF", size=10) # white bold header text
WARN_FILL   = PatternFill("solid", fgColor="FFF3E0")   # light amber for WARN rows
OK_FILL     = PatternFill("solid", fgColor="E8F5E9")   # light green for OK rows
BOLD_FONT   = Font(bold=True)
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
THIN_BORDER = Border(
    left   = Side(style='thin'),
    right  = Side(style='thin'),
    top    = Side(style='thin'),
    bottom = Side(style='thin'),
)


def _apply_header(cell, value: str):
    """
    Apply standard header styling to a single cell.

    Used for every header row across all 6 sheets so they share a consistent
    look and any future style change only needs to be made in one place.

    Args:
        cell:   openpyxl Cell object.
        value:  Column header label string.
    """
    cell.value     = value
    cell.font      = HDR_FONT
    cell.fill      = HDR_FILL
    cell.alignment = CENTER_ALIGN
    cell.border    = THIN_BORDER


def _autofit_columns(ws, max_width: int = 40):
    """
    Set each column's width to the length of its longest cell value.

    Caps at max_width to prevent unwieldy wide columns from long description
    strings.  Adds 3 characters of padding for readability.

    Args:
        ws:         openpyxl Worksheet object.
        max_width:  Maximum column width in character units (default 40).
    """
    for col in ws.columns:
        letter   = col[0].column_letter
        max_len  = max((len(str(c.value)) for c in col if c.value), default=8)
        ws.column_dimensions[letter].width = min(max_len + 3, max_width)


def write_output_workbook(result: ProcessingResult, output_path: Path) -> None:
    """
    Write the full 6-sheet D365-ready Excel workbook to output_path.

    Sheet structure (matching Online PO Management system):
    ┌─────────────────┬──────────────────────────────────────────────────────┐
    │ Sheet           │ Contents                                             │
    ├─────────────────┼──────────────────────────────────────────────────────┤
    │ Headers (SO)    │ One row per unique PO/SO — 18 columns for D365      │
    │ Lines (SO)      │ One row per SKU line — line numbers reset per PO    │
    │ Summary         │ One row per PO with item count, qty totals, status  │
    │ Validation      │ Full line detail: EAN, MRP, tester flag, status     │
    │ Warnings        │ All mapping failures (or "No warnings" if clean)    │
    │ Raw Data        │ Complete audit trail of every processed line        │
    └─────────────────┴──────────────────────────────────────────────────────┘

    Args:
        result:      Populated ProcessingResult from process_csv_files().
        output_path: Absolute Path where the .xlsx file will be saved.
    """
    log.info(f"[Writer] Writing workbook to: {output_path}")

    today_str = datetime.now().strftime("%d-%m-%Y")   # date string for all date columns

    wb = Workbook()
    wb.remove(wb.active)   # remove the default empty sheet that openpyxl creates

    # ──────────────────────────────────────────────────────────────────────────
    # SHEET 1: Headers (SO)
    # ──────────────────────────────────────────────────────────────────────────
    # One row per unique Sales Order (= one row per unique PO).
    # Column layout matches the D365 Sales Order Header import template exactly.
    # Columns 14-18 (dimensions) are left blank for H&G but must be present for
    # the D365 importer to accept the file without template errors.

    ws_hdr = wb.create_sheet("Headers (SO)")

    header_cols = [
        "Document Type",           # col 1  — always "Order"
        "No.",                     # col 2  — SO number (e.g. SO/HG/05/23526)
        "Sell-to Customer No.",    # col 3  — customer number from Address Master
        "Ship-to Code",            # col 4  — ship-to code from Address Master
        "Posting Date",            # col 5  — today's date
        "Order Date",              # col 6  — today's date
        "Document Date",           # col 7  — today's date
        "Invoice From Date",       # col 8  — today's date
        "Invoice To Date",         # col 9  — today's date
        "External Document No.",   # col 10 — repeated SO number (D365 ext ref field)
        "Location Code",           # col 11 — warehouse D365 code (e.g. "PICK")
        "Dimension Set ID",        # col 12 — blank for H&G
        "Supply Type",             # col 13 — always "B2B"
        "Voucher Narration",       # col 14 — blank (dimension field)
        "Brand Code (Dimension)",  # col 15 — blank (dimension field)
        "Channel Code (Dimension)",# col 16 — blank (dimension field)
        "Catagory (Dimension)",    # col 17 — blank (note: original typo preserved)
        "Geography Code (Dimension)", # col 18 — blank (dimension field)
    ]
    for c, h in enumerate(header_cols, 1):
        _apply_header(ws_hdr.cell(1, c), h)

    # Track which SO numbers have been written to avoid duplicates
    # (multiple SKU rows share the same SO but should produce only one Header row).
    seen_so: set = set()
    r = 2   # data starts at row 2 (row 1 is the header)

    for sorow in result.rows:
        if sorow.so_number in seen_so:
            continue   # only write the header once per SO
        seen_so.add(sorow.so_number)

        ws_hdr.cell(r, 1,  "Order")             # Document Type — always Order
        ws_hdr.cell(r, 2,  sorow.so_number)     # SO number
        ws_hdr.cell(r, 3,  sorow.cust_no)       # Sell-to Customer No.
        ws_hdr.cell(r, 4,  sorow.ship_to)       # Ship-to Code
        ws_hdr.cell(r, 5,  today_str)           # Posting Date
        ws_hdr.cell(r, 6,  today_str)           # Order Date
        ws_hdr.cell(r, 7,  today_str)           # Document Date
        ws_hdr.cell(r, 8,  today_str)           # Invoice From Date
        ws_hdr.cell(r, 9,  today_str)           # Invoice To Date
        ws_hdr.cell(r, 10, sorow.so_number)     # External Document No. = SO number
        ws_hdr.cell(r, 11, result.warehouse_code) # Location Code
        ws_hdr.cell(r, 12, "")                  # Dimension Set ID — blank
        ws_hdr.cell(r, 13, "B2B")               # Supply Type — always B2B for H&G
        for c in range(14, 19):
            ws_hdr.cell(r, c, "")              # Dimension columns — all blank
        r += 1

    ws_hdr.freeze_panes = "A2"   # freeze header row so it stays visible when scrolling
    _autofit_columns(ws_hdr)

    # ──────────────────────────────────────────────────────────────────────────
    # SHEET 2: Lines (SO)
    # ──────────────────────────────────────────────────────────────────────────
    # One row per SKU per PO.
    # Line numbers are 10000, 20000, 30000, … and RESET for each new PO.
    # This matches D365 import expectations and is required by the business.

    ws_line = wb.create_sheet("Lines (SO)")

    line_cols = [
        "Document Type",   # col 1 — always "Order"
        "Document No.",    # col 2 — SO number (links to Header sheet)
        "Line No.",        # col 3 — 10000, 20000, … (resets per PO)
        "Type",            # col 4 — always "Item"
        "No.",             # col 5 — D365 Item Number from Items Master
        "Location Code",   # col 6 — warehouse code (same as Header)
        "Quantity",        # col 7 — final qty (1 for tester, actual for regular)
        "Unit Price",      # col 8 — 0.54 for tester, blank for regular
    ]
    for c, h in enumerate(line_cols, 1):
        _apply_header(ws_line.cell(1, c), h)

    r          = 2
    current_po = None   # tracks current PO to detect when we move to the next one
    line_no    = 0      # resets to 0 each time current_po changes

    for sorow in result.rows:
        if sorow.po_number != current_po:
            # New PO encountered — reset line number counter.
            current_po = sorow.po_number
            line_no    = 0

        line_no += 10000   # increment by 10000 (D365 convention for SO lines)

        ws_line.cell(r, 1, "Order")                 # Document Type
        ws_line.cell(r, 2, sorow.so_number)         # Document No.
        ws_line.cell(r, 3, line_no)                 # Line No.
        ws_line.cell(r, 4, "Item")                  # Type
        ws_line.cell(r, 5, sorow.item_no)           # No. (D365 Item Number)
        ws_line.cell(r, 6, result.warehouse_code)   # Location Code
        ws_line.cell(r, 7, sorow.qty)               # Quantity
        # Unit Price: 0.54 for tester lines, blank string for regular
        # (blank is correct for D365 regular import — price comes from the item card)
        ws_line.cell(r, 8, sorow.unit_price if sorow.is_tester else "")
        r += 1

    ws_line.freeze_panes = "A2"
    _autofit_columns(ws_line)

    # ──────────────────────────────────────────────────────────────────────────
    # SHEET 3: Summary
    # ──────────────────────────────────────────────────────────────────────────
    # One row per PO — provides a quick overview for order validation.
    # Status cell is colour-coded: green = OK, amber = WARN (mapping issues).

    ws_sum = wb.create_sheet("Summary")

    sum_cols = [
        "PO",           # original PO number
        "SO Number",    # generated SO number
        "Store (Raw)",  # store name from CSV
        "Ship-to Code", # resolved from Address Master
        "Cust No",      # same as Ship-to for H&G
        "Items",        # count of unique SKUs in this PO
        "Total Qty",    # sum of all quantities in this PO
        "Status",       # OK or WARN
    ]
    for c, h in enumerate(sum_cols, 1):
        _apply_header(ws_sum.cell(1, c), h)

    r = 2
    for po, info in result.po_summary.items():
        ws_sum.cell(r, 1, po)
        ws_sum.cell(r, 2, info['so_number'])
        ws_sum.cell(r, 3, info['store'])
        ws_sum.cell(r, 4, info['ship_to'])
        ws_sum.cell(r, 5, info['cust_no'])
        ws_sum.cell(r, 6, info['items'])
        ws_sum.cell(r, 7, info['total_qty'])
        status_cell      = ws_sum.cell(r, 8, info['status'])
        status_cell.fill = OK_FILL if info['status'] == "OK" else WARN_FILL
        status_cell.font = BOLD_FONT
        r += 1

    ws_sum.freeze_panes = "A2"
    _autofit_columns(ws_sum)

    # ──────────────────────────────────────────────────────────────────────────
    # SHEET 4: Validation
    # ──────────────────────────────────────────────────────────────────────────
    # Full line-level detail for inspection before uploading to D365.
    # Each row corresponds to one SORow — the user can cross-check EAN, MRP,
    # description, and quantity here before committing the import.

    ws_val = wb.create_sheet("Validation")

    val_cols = [
        "PO",          # PO number
        "SO Number",   # assigned SO number
        "Item No",     # D365 Item Number (or ?MISSING_SKU / ?MISSING_EAN)
        "EAN",         # resolved EAN barcode
        "Description", # product name from Items Master
        "MRP",         # Maximum Retail Price from Items Master
        "Qty",         # final processed quantity
        "Unit Price",  # cost price (only shown for tester; blank for regular)
        "Tester",      # YES / NO tester flag
        "Status",      # OK / WARN
    ]
    for c, h in enumerate(val_cols, 1):
        _apply_header(ws_val.cell(1, c), h)

    r = 2
    for sorow in result.rows:
        ws_val.cell(r, 1, sorow.po_number)
        ws_val.cell(r, 2, sorow.so_number)
        ws_val.cell(r, 3, sorow.item_no)
        ws_val.cell(r, 4, sorow.ean)
        ws_val.cell(r, 5, sorow.description)
        ws_val.cell(r, 6, sorow.mrp if sorow.mrp else "")
        ws_val.cell(r, 7, sorow.qty)
        ws_val.cell(r, 8, sorow.unit_price if sorow.is_tester else "")
        ws_val.cell(r, 9, "YES" if sorow.is_tester else "NO")
        status_cell      = ws_val.cell(r, 10, sorow.status)
        status_cell.fill = OK_FILL if sorow.status == "OK" else WARN_FILL
        status_cell.font = BOLD_FONT
        r += 1

    ws_val.freeze_panes = "A2"
    _autofit_columns(ws_val)

    # ──────────────────────────────────────────────────────────────────────────
    # SHEET 5: Warnings
    # ──────────────────────────────────────────────────────────────────────────
    # Collects all non-fatal issues: missing SKUs, missing EANs, unknown stores,
    # duplicate rows, unreadable files, etc.
    # If there are no warnings, a single green "all OK" row is shown so the
    # user can confirm the sheet was checked rather than just being empty.

    ws_warn = wb.create_sheet("Warnings")

    warn_cols = ["PO", "SKU / Item", "Warning Message"]
    for c, h in enumerate(warn_cols, 1):
        _apply_header(ws_warn.cell(1, c), h)

    if result.warnings:
        # Write one row per warning, all amber-highlighted.
        for r_idx, (po, sku, msg) in enumerate(result.warnings, 2):
            ws_warn.cell(r_idx, 1, po)
            ws_warn.cell(r_idx, 2, sku)
            ws_warn.cell(r_idx, 3, msg)
            for c in range(1, 4):
                ws_warn.cell(r_idx, c).fill = WARN_FILL   # amber row
    else:
        # No warnings — write a single green confirmation row.
        ws_warn.cell(2, 1, "")
        ws_warn.cell(2, 2, "")
        ws_warn.cell(2, 3, "No warnings — all SKUs, EANs, and stores mapped successfully ✓")
        ws_warn.cell(2, 3).fill = OK_FILL

    _autofit_columns(ws_warn)

    # ──────────────────────────────────────────────────────────────────────────
    # SHEET 6: Raw Data
    # ──────────────────────────────────────────────────────────────────────────
    # Complete audit trail — one row per processed SORow with all fields
    # including the original SKU code, EAN, resolved Item No, final qty,
    # source file name, and warehouse.  Used for reconciliation and debugging.

    ws_raw = wb.create_sheet("Raw Data")

    raw_cols = [
        "Source File",  # name of the CSV file this line came from
        "PO",           # original PO number
        "Store",        # store name from CSV
        "SKU Code",     # original SKU from CSV (before EAN/Item lookup)
        "EAN",          # resolved EAN from HG Master
        "Description",  # product description from Items Master
        "Final Qty",    # qty after tester override (if applicable)
        "MRP",          # MRP from Items Master
        "Unit Price",   # 0.54 for tester; blank for regular
        "SO Number",    # assigned SO number
        "Item No",      # D365 Item Number from Items Master
        "Warehouse",    # D365 Location Code
        "Tester",       # YES / NO
        "Status",       # OK / WARN
    ]
    for c, h in enumerate(raw_cols, 1):
        _apply_header(ws_raw.cell(1, c), h)

    r = 2
    for sorow in result.rows:
        # Source file: derive from the first input file for all rows.
        # In a multi-file run every row is tagged to the first file as a
        # simplification — a more precise per-row attribution would require
        # storing the file name on SORow (future enhancement).
        src_file = os.path.basename(result.input_files[0]) if result.input_files else ""

        ws_raw.cell(r, 1,  src_file)
        ws_raw.cell(r, 2,  sorow.po_number)
        ws_raw.cell(r, 3,  sorow.store_name)
        ws_raw.cell(r, 4,  sorow.sku_code)         # original SKU preserved on SORow
        ws_raw.cell(r, 5,  sorow.ean)
        ws_raw.cell(r, 6,  sorow.description)
        ws_raw.cell(r, 7,  sorow.qty)
        ws_raw.cell(r, 8,  sorow.mrp if sorow.mrp else "")
        ws_raw.cell(r, 9,  sorow.unit_price if sorow.is_tester else "")
        ws_raw.cell(r, 10, sorow.so_number)
        ws_raw.cell(r, 11, sorow.item_no)
        ws_raw.cell(r, 12, result.warehouse_code)
        ws_raw.cell(r, 13, "YES" if sorow.is_tester else "NO")
        status_cell      = ws_raw.cell(r, 14, sorow.status)
        status_cell.fill = OK_FILL if sorow.status == "OK" else WARN_FILL
        r += 1

    ws_raw.freeze_panes = "A2"
    _autofit_columns(ws_raw)

    # ── Save ──────────────────────────────────────────────────────────────────
    wb.save(output_path)
    log.info(
        f"[Writer] Workbook saved — "
        f"{len(seen_so)} SOs | {len(result.rows)} lines | 6 sheets | {output_path.name}"
    )


# ==============================================================================
# SECTION 10 — GUI: LOG HANDLER
# ==============================================================================

class GuiLogHandler(logging.Handler):
    """
    Custom logging.Handler that appends log records to a Tkinter Text widget.

    This is attached to the root 'mt_select' logger after the Text widget is
    created in MTSelectApp.__init__().  All subsequent log calls from any
    module therefore appear in both the GUI log panel and the log file.

    Thread safety:
        The Text widget can only be safely written from the main Tkinter thread.
        We use widget.after(0, fn) to schedule the append on the main thread
        even when the log call originates from the watcher background thread.

    Args:
        text_widget:  The tk.Text widget to write log lines into.
    """

    def __init__(self, text_widget: tk.Text):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record: logging.LogRecord):
        """Format the record and schedule appending it to the Text widget."""
        msg = self.format(record)

        def _append():
            """Runs on the main Tkinter thread via after(0, …)."""
            self.text_widget.config(state='normal')
            self.text_widget.insert('end', msg + '\n')
            self.text_widget.see('end')          # auto-scroll to bottom
            self.text_widget.config(state='disabled')

        try:
            self.text_widget.after(0, _append)   # thread-safe GUI update
        except Exception:
            pass   # widget may be destroyed on exit — swallow silently


# ==============================================================================
# SECTION 11 — MAIN GUI APPLICATION
# ==============================================================================

class MTSelectApp:
    """
    Main Tkinter application window for the MT Select H&G Processor.

    Responsibilities:
    - Build and lay out the entire GUI.
    - Auto-load master files from data_mt/ on startup.
    - Start the MasterWatcher background thread for auto-reload.
    - Provide Browse buttons for manual master file selection.
    - Accept CSV file selection.
    - Trigger the processing engine and write output on ▶ Generate SO click.
    - Display a live sequence counter showing the next SO number.
    - Show the run summary and open the output file / log folder.

    Attributes:
        root              : Tkinter root window.
        csv_paths         : List of selected CSV file paths.
        warehouse_var     : StringVar for the warehouse dropdown.
        tester_var        : BooleanVar for the Tester mode checkbox.
        status_var        : StringVar for the bottom status label.
        last_output       : Path of the last generated Excel file.
        last_result       : ProcessingResult from the last run (for re-export).
        hg_master         : HGMasterLoader instance (shared with watcher).
        items_master      : ItemsMasterLoader instance (shared with watcher).
        address_master    : AddressMasterLoader instance (shared with watcher).
        hg_path_var       : StringVar showing HG Master status in the UI.
        items_path_var    : StringVar showing Items Master status in the UI.
        address_path_var  : StringVar showing Address Master status in the UI.
        master_status_var : StringVar for the combined master-ready banner.
        seq_var           : StringVar for the sequence info banner.
        watcher           : MasterWatcher thread instance.
    """

    def __init__(self):
        # ── Root window ───────────────────────────────────────────────────────
        self.root = tk.Tk()
        self.root.title("MT Select (Health & Glow) Processor  v3.0")
        self.root.geometry("740x860")
        self.root.resizable(False, False)

        # ── Application state ─────────────────────────────────────────────────
        self.csv_paths:   List[str]            = []    # CSV files selected by user
        self.last_output: Optional[Path]       = None  # last generated Excel path
        self.last_result: Optional[ProcessingResult] = None  # last run result

        # ── Tkinter variables (bound to GUI widgets) ──────────────────────────
        self.warehouse_var    = tk.StringVar(value=DEFAULT_WAREHOUSE)
        self.tester_var       = tk.BooleanVar(value=False)
        self.status_var       = tk.StringVar(value="Initialising…")
        self.hg_path_var      = tk.StringVar(value="Not selected")
        self.items_path_var   = tk.StringVar(value="Not selected")
        self.address_path_var = tk.StringVar(value="Not selected")
        self.master_status_var = tk.StringVar(value="Masters: Loading…")
        self.seq_var          = tk.StringVar(value="Loading sequence…")
        self.csv_var          = tk.StringVar(value="No files selected")

        # ── Master loader instances ───────────────────────────────────────────
        # Created here so both the GUI and the MasterWatcher share the same objects.
        self.hg_master      = HGMasterLoader()
        self.items_master   = ItemsMasterLoader()
        self.address_master = AddressMasterLoader()

        # ── Build all widgets ─────────────────────────────────────────────────
        self._build_ui()

        # ── Attach GUI log handler AFTER Text widget exists ───────────────────
        # Must be done after _build_ui() because the Text widget is created there.
        gui_handler = GuiLogHandler(self.log_text)
        gui_handler.setLevel(logging.DEBUG)
        gui_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)-7s] %(message)s", "%H:%M:%S")
        )
        log.addHandler(gui_handler)

        # ── Startup log messages ──────────────────────────────────────────────
        log.info("=" * 60)
        log.info("MT Select (Health & Glow) Processor  v3.0  started")
        log.info(f"Script folder  : {SCRIPT_DIR}")
        log.info(f"data_mt folder : {DATA_MT_DIR}")
        log.info(f"Log file       : {LOG_FILE}")
        log.info("=" * 60)

        # ── Initial master auto-load ──────────────────────────────────────────
        self._auto_load_masters()

        # ── Start background watcher thread ───────────────────────────────────
        # The watcher will silently reload masters if their files change on disk.
        self.watcher = MasterWatcher(self)
        self.watcher.start()
        log.info(
            f"[Watcher] Auto-reload watcher started "
            f"(interval: {MASTER_WATCH_INTERVAL_SEC}s)"
        )

    # ──────────────────────────────────────────────────────────────────────────
    # GUI BUILDER
    # ──────────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        """
        Construct and lay out all Tkinter widgets.

        Layout (top to bottom):
          1. Title label
          2. Subtitle label
          3. Sequence info banner (blue box showing next SO number)
          4. Warehouse selector + Tester checkbox
          5. Master Files section (3 rows with path labels + Browse buttons)
          6. Input CSV Files section (path label + Browse button)
          7. Action buttons row 1: Generate SO | Open Last Output
          8. Action buttons row 2: Open Log Folder | Reload Masters
          9. Status label (changes colour: orange/blue/green/red)
         10. Processing Log text area (scrollable, read-only)
        """

        # ── 1. Title ─────────────────────────────────────────────────────────
        tk.Label(
            self.root,
            text="MT Select (Health & Glow)",
            font=("Arial", 15, "bold")
        ).pack(pady=(12, 2))

        # ── 2. Subtitle ──────────────────────────────────────────────────────
        tk.Label(
            self.root,
            text="CSV  →  D365 Sales Order Import  |  v3.0",
            font=("Arial", 9),
            fg="gray"
        ).pack(pady=(0, 6))

        # ── 3. Sequence info banner ──────────────────────────────────────────
        # Light-blue box showing the current HG/TT sequence and the next SO
        # that will be generated.  Updates live when Tester checkbox is toggled
        # and after each successful Generate run.
        seq_frame = tk.Frame(self.root, bg="#E3F2FD", relief='groove', bd=1)
        seq_frame.pack(fill='x', padx=20, pady=(0, 6))
        tk.Label(
            seq_frame,
            textvariable=self.seq_var,
            font=("Consolas", 9),
            bg="#E3F2FD",
            fg="#0D47A1"
        ).pack(pady=5, padx=8)

        # ── 4. Warehouse + Tester row ─────────────────────────────────────────
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill='x', padx=20, pady=4)

        tk.Label(top_frame, text="Warehouse:", font=("Arial", 10, "bold")).pack(side='left')

        # Warehouse dropdown — values come from the WAREHOUSES dict keys.
        wh_combo = ttk.Combobox(
            top_frame,
            textvariable=self.warehouse_var,
            values=list(WAREHOUSES.keys()),
            state='readonly',
            width=8
        )
        wh_combo.pack(side='left', padx=8)

        # Dynamic label showing the D365 Location Code for the selected warehouse.
        wh_code_label = tk.Label(
            top_frame,
            text=f"→ {WAREHOUSES[self.warehouse_var.get()]}",
            font=("Arial", 9),
            fg="gray"
        )
        wh_code_label.pack(side='left', padx=4)

        # Update the D365 code label whenever the dropdown selection changes.
        wh_combo.bind(
            '<<ComboboxSelected>>',
            lambda e: wh_code_label.config(
                text=f"→ {WAREHOUSES[self.warehouse_var.get()]}"
            )
        )

        # Tester mode checkbox — on toggle, refresh the sequence banner.
        tk.Checkbutton(
            top_frame,
            text="Tester Orders  (SO/TT, qty=1, CP=₹0.54)",
            variable=self.tester_var,
            font=("Arial", 9),
            command=self._refresh_seq_display   # update banner immediately on toggle
        ).pack(side='right')

        # ── 5. Master Files section ───────────────────────────────────────────
        master_frame = tk.LabelFrame(
            self.root,
            text="Master Files  (auto-loaded from data_mt/ — no manual upload needed)",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=8
        )
        master_frame.pack(fill='x', padx=20, pady=6)

        # One row per master file: label | path display | Browse button.
        self._master_row(master_frame, "HG Master (SKU→EAN):",
                         self.hg_path_var, self._select_hg_master)
        self._master_row(master_frame, "Items Master (EAN→Item):",
                         self.items_path_var, self._select_items_master)
        self._master_row(master_frame, "Address Master (Store→ShipTo):",
                         self.address_path_var, self._select_address_master)

        # Combined status banner — green when all three masters are loaded.
        tk.Label(
            master_frame,
            textvariable=self.master_status_var,
            font=("Arial", 9),
            fg="blue"
        ).pack(anchor='w', pady=(4, 0))

        # ── 6. Input CSV Files section ────────────────────────────────────────
        csv_frame = tk.LabelFrame(
            self.root,
            text="Input CSV Files",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=8
        )
        csv_frame.pack(fill='x', padx=20, pady=6)

        csv_row = tk.Frame(csv_frame)
        csv_row.pack(fill='x')
        tk.Label(csv_row, text="CSV Files:", font=("Arial", 9)).pack(side='left')
        tk.Label(
            csv_row,
            textvariable=self.csv_var,
            font=("Arial", 9),
            fg="blue",
            width=44,
            anchor='w'
        ).pack(side='left', padx=8)
        tk.Button(
            csv_row,
            text="Browse",
            command=self._select_csv_files
        ).pack(side='right')

        # ── 7. Action buttons row 1 ───────────────────────────────────────────
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=8)

        tk.Button(
            btn_frame,
            text="▶  Generate SO",
            width=24,
            font=("Arial", 10, "bold"),
            bg="#00C853",       # green — primary action
            fg="white",
            command=self.generate
        ).pack(side='left', padx=6)

        # Open Last Output — disabled until the first successful generation.
        self.open_btn = tk.Button(
            btn_frame,
            text="📂  Open Last Output",
            width=24,
            state=tk.DISABLED,   # enabled after first successful Generate
            command=self.open_last
        )
        self.open_btn.pack(side='left', padx=6)

        # ── 8. Action buttons row 2 ───────────────────────────────────────────
        btn_frame2 = tk.Frame(self.root)
        btn_frame2.pack(pady=2)

        tk.Button(
            btn_frame2,
            text="📂  Open Log Folder",
            width=24,
            command=self.open_log_folder
        ).pack(side='left', padx=6)

        tk.Button(
            btn_frame2,
            text="🔄  Reload Masters",
            width=24,
            command=self._auto_load_masters   # manual reload trigger
        ).pack(side='left', padx=6)

        # ── 9. Status label ───────────────────────────────────────────────────
        # Colour meanings:
        #   gray       — initial/idle
        #   orange     — masters not ready
        #   blue       — processing in progress
        #   darkgreen  — success
        #   red        — error / no output produced
        self.status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Arial", 10),
            fg="gray",
            wraplength=700
        )
        self.status_label.pack(pady=4)

        # ── 10. Processing Log ────────────────────────────────────────────────
        log_frame = tk.LabelFrame(
            self.root,
            text="Processing Log  ·  also saved to Logs/ folder",
            font=("Arial", 9)
        )
        log_frame.pack(fill='both', expand=True, padx=20, pady=(0, 12))

        scroll = ttk.Scrollbar(log_frame, orient='vertical')
        scroll.pack(side='right', fill='y')

        # Read-only Text widget — state toggled to 'normal' only when appending.
        self.log_text = tk.Text(
            log_frame,
            height=14,
            font=("Consolas", 8),
            state='disabled',
            wrap='word',
            yscrollcommand=scroll.set
        )
        self.log_text.pack(fill='both', expand=True)
        scroll.config(command=self.log_text.yview)

    def _master_row(self, parent, label_text: str, path_var, command):
        """
        Helper to build one master-file row inside the Master Files section.

        Each row contains:
            [Label: 28 chars]  [Path display: 28 chars]  [Browse button]

        Args:
            parent:     Parent tk.Frame (the LabelFrame for master files).
            label_text: Descriptive label (e.g. "HG Master (SKU→EAN):").
            path_var:   StringVar that displays the current file name / status.
            command:    Callback for the Browse button.
        """
        frame = tk.Frame(parent)
        frame.pack(fill='x', pady=3)
        tk.Label(frame, text=label_text, font=("Arial", 9), width=30, anchor='w').pack(side='left')
        tk.Label(frame, textvariable=path_var, font=("Arial", 9), fg="blue",
                 width=28, anchor='w').pack(side='left', padx=4)
        tk.Button(frame, text="Browse", command=command).pack(side='right')

    # ──────────────────────────────────────────────────────────────────────────
    # SEQUENCE DISPLAY
    # ──────────────────────────────────────────────────────────────────────────

    def _refresh_seq_display(self):
        """
        Update the sequence info banner to reflect the current mode and
        the next SO number that will be generated.

        Called:
        - After _auto_load_masters() completes (startup).
        - When the Tester checkbox is toggled.
        - After each successful Generate run.

        Reads the current sequences fresh from SEQ_FILE each time so the
        display is accurate even if the file was edited externally.
        """
        seqs  = load_sequences()
        month = datetime.now().strftime("%m")   # current month for display

        if self.tester_var.get():
            # Tester mode: TT sequence active, HG paused.
            next_so = f"SO/TT/{month}/{seqs['TT'] + 1}"
            self.seq_var.set(
                f"TESTER MODE  ·  Next SO: {next_so}  ·  "
                f"TT sequence: {seqs['TT']}  ·  HG sequence: {seqs['HG']} (paused)"
            )
        else:
            # Regular mode: HG sequence active, TT paused.
            next_so = f"SO/HG/{month}/{seqs['HG'] + 1}"
            self.seq_var.set(
                f"REGULAR MODE  ·  Next SO: {next_so}  ·  "
                f"HG sequence: {seqs['HG']}  ·  TT sequence: {seqs['TT']} (paused)"
            )

    # ──────────────────────────────────────────────────────────────────────────
    # MASTER STATUS
    # ──────────────────────────────────────────────────────────────────────────

    def _update_master_status(self):
        """
        Check whether all three masters are loaded and update the status banner.

        Returns:
            True if all masters have at least one record loaded.
            False if any master is empty (showing which ones are missing).

        This is called:
        - After each individual master load (auto or manual Browse).
        - By the watcher thread (via root.after) after a background reload.
        - At the start of generate() to gate the run.
        """
        if (self.hg_master.sku_to_ean
                and self.items_master.ean_to_item
                and self.address_master.store_to_ship):
            # All three loaded — show counts and ready indicator.
            self.master_status_var.set(
                f"Masters: Ready ✓  ·  "
                f"HG = {len(self.hg_master.sku_to_ean)} SKUs  ·  "
                f"Items = {len(self.items_master.ean_to_item)} EANs  ·  "
                f"Addresses = {len(self.address_master.store_to_ship)} stores"
            )
            self.status_var.set("Ready — select CSV files and click ▶ Generate SO")
            self.status_label.config(fg="darkgreen")
            return True

        # One or more masters missing — list which ones.
        missing = []
        if not self.hg_master.sku_to_ean:      missing.append("HG Master")
        if not self.items_master.ean_to_item:  missing.append("Items Master")
        if not self.address_master.store_to_ship: missing.append("Address Master")

        self.master_status_var.set(f"Masters: Missing → {', '.join(missing)}")
        self.status_var.set(
            "Place master files in data_mt/ folder (or use Browse to select manually)"
        )
        self.status_label.config(fg="orange")
        return False

    # ──────────────────────────────────────────────────────────────────────────
    # AUTO-LOAD MASTERS
    # ──────────────────────────────────────────────────────────────────────────

    def _auto_load_masters(self):
        """
        Scan DATA_MT_DIR for master files and load them automatically.

        File matching rules (sorted reverse-alphabetically so latest version wins):
            HG Master:      matches  HG Master*.xlsx
            Items Master:   matches  Items*.xlsx
            Address Master: matches  H&G Addresses*.xlsx

        If DATA_MT_DIR doesn't exist yet (created at import but could be on a
        different drive), a clear warning is logged and the user is prompted to
        use Browse or add files to data_mt/.

        This method is also bound to the "Reload Masters" button so the user
        can trigger a fresh scan after dropping new files into data_mt/.
        """
        log.info(f"[Auto-load] Scanning data_mt/: {DATA_MT_DIR}")

        if not DATA_MT_DIR.exists():
            log.warning(
                f"[Auto-load] data_mt/ folder not found at {DATA_MT_DIR}.\n"
                f"Create the folder and place master files inside it, "
                f"or use Browse to select them manually."
            )
            self._update_master_status()
            self._refresh_seq_display()
            return

        # ── HG Master (SKU → EAN) ─────────────────────────────────────────────
        # sorted(..., reverse=True) picks the alphabetically last file, which by
        # naming convention (e.g. "HG Master Dec 25.xlsx") is the most recent.
        for f in sorted(DATA_MT_DIR.glob("HG Master*.xlsx"), reverse=True):
            try:
                cnt = self.hg_master.load(f)
                self.hg_path_var.set(f"{f.name}  ({cnt} SKUs) ✓")
                log.info(f"[Auto-load] HG Master loaded: {f.name}  ({cnt} SKUs)")
                break   # stop after the first successful load
            except Exception as e:
                log.error(f"[Auto-load] HG Master failed ({f.name}): {e}")

        # ── Items Master (EAN → D365 Item) ────────────────────────────────────
        for f in sorted(DATA_MT_DIR.glob("Items*.xlsx"), reverse=True):
            try:
                cnt = self.items_master.load(f)
                self.items_path_var.set(f"{f.name}  ({cnt} EANs) ✓")
                log.info(f"[Auto-load] Items Master loaded: {f.name}  ({cnt} EANs)")
                break
            except Exception as e:
                log.error(f"[Auto-load] Items Master failed ({f.name}): {e}")

        # ── Address Master (Store → Ship-to) ──────────────────────────────────
        for f in sorted(DATA_MT_DIR.glob("H&G Addresses*.xlsx"), reverse=True):
            try:
                cnt = self.address_master.load(f)
                self.address_path_var.set(f"{f.name}  ({cnt} stores) ✓")
                log.info(f"[Auto-load] Address Master loaded: {f.name}  ({cnt} stores)")
                break
            except Exception as e:
                log.error(f"[Auto-load] Address Master failed ({f.name}): {e}")

        self._update_master_status()
        self._refresh_seq_display()

    # ──────────────────────────────────────────────────────────────────────────
    # MANUAL BROWSE HANDLERS
    # ──────────────────────────────────────────────────────────────────────────
    # These are used when the user wants to override the auto-loaded file with
    # a specific version, or when the file is stored outside data_mt/.

    def _select_hg_master(self):
        """Open a file dialog for manual HG Master selection and load it."""
        path = filedialog.askopenfilename(
            title="Select HG Master (SKU→EAN)",
            initialdir=str(DATA_MT_DIR),   # open dialog in data_mt/ by default
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:   # user pressed Cancel → path is empty string
            try:
                cnt = self.hg_master.load(Path(path))
                self.hg_path_var.set(f"{os.path.basename(path)}  ({cnt} SKUs) ✓")
                self._update_master_status()
            except Exception as e:
                log.error(f"[HG Master] Manual load error: {e}")
                messagebox.showerror("Load Error", str(e))

    def _select_items_master(self):
        """Open a file dialog for manual Items Master selection and load it."""
        path = filedialog.askopenfilename(
            title="Select Items Master (EAN→Item)",
            initialdir=str(DATA_MT_DIR),
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            try:
                cnt = self.items_master.load(Path(path))
                self.items_path_var.set(f"{os.path.basename(path)}  ({cnt} EANs) ✓")
                self._update_master_status()
            except Exception as e:
                log.error(f"[Items Master] Manual load error: {e}")
                messagebox.showerror("Load Error", str(e))

    def _select_address_master(self):
        """Open a file dialog for manual Address Master selection and load it."""
        path = filedialog.askopenfilename(
            title="Select Address Master (Store→ShipTo)",
            initialdir=str(DATA_MT_DIR),
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            try:
                cnt = self.address_master.load(Path(path))
                self.address_path_var.set(f"{os.path.basename(path)}  ({cnt} stores) ✓")
                self._update_master_status()
            except Exception as e:
                log.error(f"[Address Master] Manual load error: {e}")
                messagebox.showerror("Load Error", str(e))

    def _select_csv_files(self):
        """
        Open a multi-file dialog for selecting H&G PO CSV files.

        Multiple files can be selected at once (e.g. one CSV per warehouse
        location in a batch) — they will all be merged and processed together.
        """
        paths = filedialog.askopenfilenames(
            title="Select H&G PO CSV files",
            filetypes=[("CSV files", "*.csv")]
        )
        if paths:   # user pressed Cancel → paths is empty tuple
            self.csv_paths = list(paths)
            self.csv_var.set(f"{len(self.csv_paths)} file(s) selected")
            log.info(
                f"[CSV] {len(self.csv_paths)} file(s) selected: "
                f"{[os.path.basename(p) for p in self.csv_paths]}"
            )

    # ──────────────────────────────────────────────────────────────────────────
    # GENERATE SO  (main action)
    # ──────────────────────────────────────────────────────────────────────────

    def generate(self):
        """
        Main Generate action — called when the user clicks ▶ Generate SO.

        Execution flow:
        1.  Validate that all masters are loaded and at least one CSV is selected.
        2.  Load current sequence counters from SEQ_FILE.
        3.  Call process_csv_files() to build the ProcessingResult.
        4.  If no rows were produced, show an error and abort.
        5.  Save updated sequence counters back to SEQ_FILE.
        6.  Determine output path: <csv_folder>/output_mt/<prefix>_<timestamp>.xlsx
        7.  Call write_output_workbook() to produce the Excel file.
        8.  Update the sequence banner and enable the Open Last Output button.
        9.  Show a completion popup with file path and warning count.
        """
        # ── Guard: masters must be loaded ────────────────────────────────────
        if not self._update_master_status():
            messagebox.showerror(
                "Masters Not Ready",
                "Please ensure all three master files are loaded.\n\n"
                "If files are in data_mt/ but not loading, click 🔄 Reload Masters."
            )
            return

        # ── Guard: at least one CSV must be selected ──────────────────────────
        if not self.csv_paths:
            messagebox.showwarning(
                "No CSV Selected",
                "Please click Browse and select at least one H&G PO CSV file."
            )
            return

        # ── UI: show "Processing…" state ─────────────────────────────────────
        self.status_var.set("Processing — please wait…")
        self.status_label.config(fg="blue")
        self.root.update()   # force GUI repaint so the status is visible immediately

        start_time = time.time()

        # ── Load current sequences ────────────────────────────────────────────
        sequences = load_sequences()
        log.info(
            f"[Generate] Run started. "
            f"Mode: {'TESTER' if self.tester_var.get() else 'REGULAR'}  ·  "
            f"HG seq: {sequences['HG']}  ·  TT seq: {sequences['TT']}"
        )

        # ── Core processing ───────────────────────────────────────────────────
        result = process_csv_files(
            file_paths      = self.csv_paths,
            hg_master       = self.hg_master,
            items_master    = self.items_master,
            address_master  = self.address_master,
            warehouse_code  = WAREHOUSES[self.warehouse_var.get()],
            is_tester       = self.tester_var.get(),
            sequences       = sequences
        )

        # ── Guard: abort if nothing was produced ──────────────────────────────
        if not result.rows:
            self.status_var.set("No valid rows extracted — check log for details")
            self.status_label.config(fg="red")
            log.error(
                "[Generate] Processing produced 0 rows. "
                "Check warnings above for missing columns or empty CSV files."
            )
            return

        # ── Persist updated sequences ─────────────────────────────────────────
        # Update the dict with whichever counter this run incremented,
        # then write both to disk atomically (json.dump is atomic on most OSes).
        sequences["HG"] = result.hg_sequence
        sequences["TT"] = result.tt_sequence
        save_sequences(sequences)
        self._refresh_seq_display()   # update the banner with new next-SO value

        self.last_result = result

        # ── Determine output path ─────────────────────────────────────────────
        # Output goes into output_mt/ inside the same folder as the first CSV.
        # This keeps the Excel file physically next to its source data.
        out_dir = Path(self.csv_paths[0]).parent / "output_mt"
        out_dir.mkdir(parents=True, exist_ok=True)
        log.info(f"[Generate] Output folder: {out_dir}")

        # File name includes prefix (HG or TT) and timestamp for traceability.
        timestamp    = datetime.now().strftime("%d-%m-%Y_%H%M%S")
        mode_prefix  = "TT" if self.tester_var.get() else "HG"
        preview_file = out_dir / f"MT_Select_{mode_prefix}_{timestamp}.xlsx"

        # ── Write Excel workbook ──────────────────────────────────────────────
        write_output_workbook(result, preview_file)
        self.last_output = preview_file
        self.open_btn.config(state=tk.NORMAL)   # enable Open Last Output button

        # ── Update status and log summary ─────────────────────────────────────
        elapsed = time.time() - start_time
        summary = (
            f"Done — {len(result.rows)} lines  ·  "
            f"{len(result.so_map)} POs  ·  "
            f"{len(result.warnings)} warnings  ·  "
            f"{elapsed:.2f}s"
        )
        self.status_var.set(summary)
        self.status_label.config(fg="darkgreen")
        log.info(f"[Generate] {summary}")
        log.info(f"[Generate] Output file: {preview_file}")

        # ── Completion popup ──────────────────────────────────────────────────
        messagebox.showinfo(
            "Processing Complete",
            f"Generated {len(result.rows)} sales lines across {len(result.so_map)} POs.\n"
            f"Warnings: {len(result.warnings)}"
            + (" — check Warnings sheet" if result.warnings else " ✓") + "\n\n"
            f"Output saved to:\n{preview_file}\n\n"
            f"Full log saved to:\n{LOG_FILE}"
        )

    # ──────────────────────────────────────────────────────────────────────────
    # UTILITY ACTIONS
    # ──────────────────────────────────────────────────────────────────────────

    def open_last(self):
        """
        Open the last generated Excel file using the default OS application
        (typically Excel on Windows).  Uses os.startfile which is Windows-only;
        on macOS/Linux this would need subprocess.run(['open', …]).
        """
        if self.last_output and self.last_output.exists():
            try:
                os.startfile(str(self.last_output))
            except Exception as e:
                log.error(f"[Open] Could not open file: {e}")
                messagebox.showerror("Open Failed", f"Cannot open file:\n{self.last_output}")
        else:
            messagebox.showwarning(
                "File Not Found",
                "No output file found.\nPlease run ▶ Generate SO first."
            )

    def open_log_folder(self):
        """
        Open the Logs/ folder in Windows Explorer (or Finder on macOS).
        Shows the folder even if os.startfile fails by falling back to a
        messagebox showing the path.
        """
        try:
            os.startfile(str(LOG_DIR))
        except Exception:
            # Fallback: show the path so the user can navigate manually.
            messagebox.showinfo("Log Folder Path", str(LOG_DIR))

    def run(self):
        """
        Start the Tkinter main event loop.

        This is a blocking call — it returns only when the window is closed.
        The watcher thread is a daemon and will be killed automatically when
        the main loop exits.
        """
        self.root.mainloop()


# ==============================================================================
# SECTION 12 — ENTRY POINT
# ==============================================================================

def main():
    """
    Application entry point.

    Steps:
    1.  Check the hard expiry date — refuse to run if past it.
    2.  Instantiate MTSelectApp (builds GUI, starts watcher, auto-loads masters).
    3.  Call app.run() which blocks in the Tkinter main loop until window closes.
    """
    # ── Expiry check ──────────────────────────────────────────────────────────
    # This is a lightweight licence gate — update EXPIRY_DATE before each build
    # to control how long the tool can be used without an update.
    expiry = datetime.strptime(EXPIRY_DATE, "%d-%m-%Y").date()
    if datetime.now().date() > expiry:
        # Show a brief Tk window with an error then exit immediately.
        root = tk.Tk()
        root.withdraw()   # hide the blank root window before showing the dialog
        messagebox.showerror(
            "Tool Expired",
            f"This tool expired on {EXPIRY_DATE}.\n"
            f"Please contact the Order Management Automation Team for an updated build."
        )
        sys.exit(0)

    # ── Launch app ────────────────────────────────────────────────────────────
    app = MTSelectApp()
    app.run()
    # Execution returns here only after the window is closed.
    log.info("MT Select application closed normally.")


# Standard Python idiom: only run main() when this file is executed directly,
# not when it is imported as a module by another script or test suite.
if __name__ == "__main__":
    main()
