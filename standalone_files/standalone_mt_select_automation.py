#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
MT Select (Health & Glow) Processor  —  v4.0
================================================================================

PURPOSE
-------
Converts Health & Glow (H&G) marketplace PO CSV files into a D365-ready Excel
workbook.  Every run produces TWO sets of Sales Orders:
  1. Regular SOs  (SO/HG/MM/NNNNN)  — always generated
  2. Tester SOs   (SO/HG/TT/NNNNN)  — generated IN ADDITION when the Tester
                                       checkbox is ticked

Both sets are written into the SAME output workbook so the user can review and
upload them together or separately.

================================================================================
SO NUMBER FORMAT  (v4.0 rules)
================================================================================

REGULAR:
  SO/HG/{MM}/{DDMYY}   where DDMYY = day(2d) + month(no pad) + year(2d)
  Example: 24-May-2026  →  first SO = SO/HG/05/24526
           next SO      = SO/HG/05/24527
           next         = SO/HG/05/24528  (increment the last 5-digit block by 1)

  Key rule: The FIRST SO of each day starts from today's DDMYY.
            If today's stored sequence is already higher because the tool ran
            earlier today, continue from it to prevent duplicate SO numbers.
            On a new date, start again from that date's DDMYY benchmark.

TESTER:
  SO/HG/TT/{DDMYY+offset}  — uses the SAME sequence value as the regular SO.
  Tester SOs are generated FOR EVERY PO alongside their paired regular SO.
  Example: regular SO/HG/05/24526 pairs with tester SO/HG/TT/24526.

  Example: 5 POs  →  5 regular SOs  +  5 tester SOs  = 10 SOs in the output.

SEQUENCE PERSISTENCE:
  The shared counter and its calendar date are saved in data_mt/mt_select_seq.json.
  Format: {"date": "2026-05-24", "HG": 24530, "TT": 24530}
  "TT" is mirrored for backward compatibility; it is not an independent counter.
  On a new calendar date, the counter resets to DDMYY-1 so the first generated
  regular/tester pair receives today's DDMYY benchmark.

================================================================================
WORKFLOW
================================================================================
1.  App starts → auto-loads 3 master files from data_mt/ folder.
2.  Background watcher thread monitors data_mt/ every 30 s; silently reloads
    any master file that has been updated on disk (newer mtime).
3.  User selects CSV file(s) → chooses warehouse → optionally ticks Tester.
4.  Click ▶ Generate SO:
      a. Each unique PO gets one regular SO.
      b. If Tester ticked, the SAME PO also gets one tester SO (qty=1, CP=0.54).
      c. Both sets are written to a single Excel workbook (9 sheets).
5.  Output saved to  <csv_folder>/output_mt/  next to the source CSVs.
6.  Log saved to  Logs/  next to the script.

================================================================================
FOLDER STRUCTURE
================================================================================
  <script_folder>/
  ├── mt_select_hg_processor.py       ← this script
  ├── data_mt/                        ← master files + sequence tracker
  │   ├── HG Master*.xlsx             ← SKU → EAN mapping
  │   ├── Items*.xlsx                 ← EAN → D365 Item No, MRP, GST
  │   ├── H&G Addresses*.xlsx         ← Store Name → Ship-to / Cust No
  │   └── mt_select_seq.json          ← {"HG": int, "TT": int}
  ├── Logs/                           ← timestamped .log per run
  └── <csv_folder>/
      ├── YourPO.csv
      └── output_mt/
          └── MT_Select_HG_TT_DD-MM-YYYY_HHMMSS.xlsx

================================================================================
MASTER FILE FORMATS
================================================================================
HG Master (SKU → EAN):
  - Sheet: 'HG SKU MASTER' (preferred) or first sheet
  - Header auto-detected (scans first 6 rows for 'sku_code')
  - Columns: sku_code, ENN code  (duplicate ENN code → first used)
  - Blank EAN rows skipped with WARNING

Items Master (EAN → D365 Item):
  - Sheet: 'Item Master' (preferred) or first sheet
  - Columns: GTIN, No., Description, Mrp, GST Group Code

Address Master (Store → Ship-to):
  - Sheet: 'Ship-To B2B' (preferred) or 'Ship-to Address List'
  - Columns: Del Location (or Name), Ship to (or Code), Cust No (optional)

CSV INPUT FORMAT (case-sensitive column names):
  PO_NO, STORE_NAME, SKU_CODE, QUANTITY, MRP
  Duplicate (PO_NO, SKU_CODE) → quantities are summed.

================================================================================
OUTPUT — 9 SHEETS
================================================================================
  Sheet 1 — Headers (SO)  : all SOs (regular + tester) — D365 import ready
  Sheet 2 — Lines (SO)    : all lines; line numbers reset per SO
  Sheet 3 — Summary       : one row per PO showing both regular + tester SOs
  Sheet 4 — SKU Pivot     : totals per SKU across regular + tester orders
  Sheet 5 — Validation    : line-level detail — EAN, MRP, qty, type, status
  Sheet 6 — Warnings      : all mapping issues
  Sheet 7 — Raw Data      : generated-line audit trail with source/master detail
  Sheet 8 — Control Check : non-blocking reconciliation checks with PASS/WARN
  Sheet 9 — Input Audit   : every source CSV row with its processing disposition

================================================================================
INTERNAL COMMERCIAL REFERENCE  (review whenever features materially change)
================================================================================
Current as-is handover position for this standalone workflow:
  - One-company internal-use licence, .py + executable, no future support:
    quote INR 1,25,000; negotiation floor normally INR 75,000 to INR 1,00,000.
  - Exclusive ownership / resale rights: do not quote below INR 2,50,000 to
    INR 3,50,000 without re-evaluating scope, documentation, and liability.

This is an internal commercial estimate, not a guarantee of sale value. Reassess
it after meaningful upgrades, new marketplace/D365 integrations, packaging,
formal automated tests, support obligations, or proven production results.

================================================================================
DEPENDENCIES:  pip install pandas openpyxl
EXPIRY:        30-06-2026
VERSION:       4.0  |  2026-05-24
================================================================================
"""

# ==============================================================================
# IMPORTS
# ==============================================================================

import os           # file operations, os.startfile (Windows open)
import sys          # sys.exit on expiry
import json         # sequence file read/write
import time         # elapsed timing, watcher sleep
import shutil       # copy master files into data_mt/
import logging      # structured logging: file + console + GUI
import threading    # background watcher daemon thread
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass, field

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ==============================================================================
# SECTION 1 — FOLDER PATHS
# ==============================================================================

SCRIPT_DIR  = Path(__file__).parent
# All paths are relative to the script file so the tool is fully portable.

DATA_MT_DIR = SCRIPT_DIR / "data_mt"
# Holds all three master Excel files + sequence JSON.
# The watcher monitors this folder for file changes.

OUTPUT_MT   = SCRIPT_DIR / "output_mt"
# Fallback output folder. In normal use output goes to <csv_folder>/output_mt/.

LOG_DIR = SCRIPT_DIR / "Logs"
# One timestamped .log file per run. Never auto-deleted.

# Create all required folders at import time.
DATA_MT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_MT.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)


# ==============================================================================
# SECTION 2 — LOGGING
# ==============================================================================

LOG_FILE = LOG_DIR / f"mt_select_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
# Timestamped log — each run gets its own file.

file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(
    logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
)

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.DEBUG)
console_handler.setFormatter(
    logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
)

log = logging.getLogger("mt_select")
log.setLevel(logging.DEBUG)
log.addHandler(file_handler)
log.addHandler(console_handler)
# GuiLogHandler attached later after the Tkinter Text widget is built.


# ==============================================================================
# SECTION 3 — CONFIGURATION
# ==============================================================================

EXPIRY_DATE = "30-06-2026"
# Hard expiry (DD-MM-YYYY). Update before deploying a new build.

WAREHOUSES: Dict[str, str] = {
    "AHD": "PICK",        # Ahmedabad → D365 Location Code
    "BLR": "DS_BL_OFF1",  # Bangalore → D365 Location Code
}
# To add a new warehouse: add key/value here only. No other changes needed.

DEFAULT_WAREHOUSE = "AHD"

TESTER_CP: float = 0.54
# Fixed Cost Price for ALL tester order lines. Business rule — never changes
# based on product MRP.

MASTER_WATCH_INTERVAL_SEC: int = 30
# Background watcher poll interval. Increase to 60+ on network shares.

SEQ_FILE = DATA_MT_DIR / "mt_select_seq.json"
# Persists the shared sequence and the date on which it was last used.
# "TT" mirrors "HG" for compatibility with older builds.
# Format: {"date": "2026-05-24", "HG": 24530, "TT": 24530}


# ==============================================================================
# SECTION 4 — SEQUENCE MANAGEMENT
# ==============================================================================
#
# HOW SEQUENCES WORK (v4.0):
# ---------------------------
# The 5-digit sequence number in the SO is based on today's date in DDMYY format.
#
#   DDMYY = day(zero-padded 2 digits) + month(1-2 digits, no padding) + year(last 2)
#   Example: 24-May-2026  → DD=24, M=5, YY=26  → base = 24526
#
# Rule 1 — First SO of the day:
#   If the stored sequence < (today's base - 1), reset to (base - 1).
#   This ensures the first generated SO of the day = base (after +1 increment).
#
# Rule 2 — Subsequent SOs same day:
#   Each new PO increments the sequence by 1.
#   So:  24526, 24527, 24528, ...
#
# Rule 3 — Stored value already higher:
#   If a previous run today already reached 24530, the next run continues
#   from 24530 (not reset to base). This prevents duplicate SO numbers.
#
# Rule 4 — Regular and tester SOs share the sequence:
#   For a PO assigned regular SO/HG/05/24526, its tester SO is
#   SO/HG/TT/24526. The next PO receives 24527 for both forms.
#
# Rule 5 — Tester runs alongside Regular:
#   When the Tester checkbox is ticked, BOTH regular and tester SOs are
#   generated for every PO. 5 POs → 5 regular SOs + 5 tester SOs = 10 total.

def _todays_base_sequence() -> int:
    """
    Calculate today's DDMYY base sequence number.

    Formula: int(f"{day:02d}{month}{year%100:02d}")
    Examples:
        24-May-2026  → int("24526")  = 24526
        01-Jan-2026  → int("01126")  = 1126
        31-Dec-2026  → int("311226") = 311226  (6-digit for Dec 31+)

    Returns:
        Integer base sequence for today.
    """
    now = datetime.now()
    # day: zero-padded 2 digits | month: no padding | year: last 2 digits
    return int(f"{now.day:02d}{now.month}{now.year % 100:02d}")


def load_sequences() -> Dict[str, object]:
    """
    Load the shared regular/tester sequence counter from SEQ_FILE.

    On each call:
    1. Read stored values (or use defaults if file missing/corrupt).
    2. Calculate today's DDMYY benchmark and ISO calendar date.
    3. If a dated saved sequence belongs to another calendar day, reset it to
       benchmark - 1 so the first PO today receives the benchmark.
    4. If an older undated file is found, preserve a higher stored HG value for
       one migration run to avoid issuing duplicate SO numbers.
    5. Return the counter ready for the caller to increment before assignment.

    Returns:
        dict containing ``date``, ``HG``, and mirrored ``TT`` values.
    """
    today_base = _todays_base_sequence()
    today_date = datetime.now().date().isoformat()
    defaults   = {"date": today_date, "HG": today_base - 1, "TT": today_base - 1}

    if not SEQ_FILE.exists():
        log.info(
            f"[Sequence] No sequence file found. "
            f"Starting fresh from today's base: shared={today_base - 1} "
            f"(first SO will be {today_base})"
        )
        return defaults

    try:
        with open(SEQ_FILE, 'r') as f:
            data = json.load(f)

        # Migrate old v1 format {"last_sequence": N} if present.
        if isinstance(data, dict) and "HG" not in data:
            old = int(data.get("last_sequence", today_base - 1))
            data = {"HG": old}
            log.info(f"[Sequence] Reading legacy last_sequence value: HG={old}")

        hg = int(data.get("HG", today_base - 1))
        stored_date = data.get("date")

        if stored_date and stored_date != today_date:
            log.info(
                f"[Sequence] New calendar date: stored={stored_date}, today={today_date}. "
                f"Resetting shared counter to {today_base-1} "
                f"(first SO pair will use {today_base})"
            )
            hg = today_base - 1
        elif hg < (today_base - 1):
            log.info(
                f"[Sequence] Stored shared={hg} < base-1={today_base-1}. "
                f"Resetting to {today_base-1} (first SO pair will use {today_base})"
            )
            hg = today_base - 1
        elif not stored_date:
            log.info(
                "[Sequence] Legacy undated sequence loaded; preserving the stored "
                "value for duplicate protection and stamping today's date on save."
            )

        result = {"date": today_date, "HG": hg, "TT": hg}
        log.debug(
            f"[Sequence] Loaded: shared={hg} | date={today_date} | "
            f"Today base={today_base} | "
            f"Next HG SO=SO/HG/{datetime.now().strftime('%m')}/{hg+1} | "
            f"Next Tester SO=SO/HG/TT/{hg+1}"
        )
        return result

    except Exception as e:
        log.warning(f"[Sequence] Failed to read {SEQ_FILE}: {e}. Using today's base.")
        return defaults


def save_sequences(seqs: Dict[str, object]) -> None:
    """
    Persist the shared sequence counter to SEQ_FILE immediately after generation.

    Written before the popup so a force-kill cannot cause duplicate SOs.
    The legacy ``TT`` key is written as a mirror of ``HG`` so older builds can
    still read the file without creating a second active sequence.

    Args:
        seqs: Mapping containing the latest ``HG`` shared counter value.
    """
    shared = int(seqs["HG"])
    payload = {
        "date": datetime.now().date().isoformat(),
        "HG": shared,
        "TT": shared,
    }
    with open(SEQ_FILE, 'w') as f:
        json.dump(payload, f, indent=2)
    log.debug(f"[Sequence] Saved: date={payload['date']}, shared={shared}")


def generate_so_number(seq: int, is_tester: bool = False) -> str:
    """
    Build a formatted SO number string.

    Args:
        seq:        Integer sequence (e.g. 24526).
        is_tester:  True → format "SO/HG/TT/seq"; False → format "SO/HG/MM/seq".

    Returns:
        Regular:  e.g. "SO/HG/05/24526"
        Tester:   e.g. "SO/HG/TT/24526"  (same seq, different prefix)
    """
    if is_tester:
        return f"SO/HG/TT/{seq}"
    else:
        month = datetime.now().strftime("%m")   # zero-padded current month
        return f"SO/HG/{month}/{seq}"


# ==============================================================================
# SECTION 5 — DATA MODELS
# ==============================================================================

@dataclass
class SORow:
    """
    One processed Sales Order line (one SKU within one PO).

    One SORow → one row in Lines (SO), Validation, and Raw Data sheets.
    Multiple SORows share so_number when they belong to the same PO.

    Fields
    ------
    po_number   : Original PO number from CSV (e.g. "6449964").
    so_number   : Generated SO (e.g. "SO/HG/05/24526").
    row_type    : "REGULAR" or "TESTER" — distinguishes which set this row
                  belongs to. Critical for inspection in Summary/Validation.
    sku_code    : Original SKU from CSV (kept for Raw Data audit).
    item_no     : D365 Item Number from Items Master.
                  Blank if SKU/EAN mapping is missing; see Warnings sheet.
    qty         : Final quantity (always 1 for tester; actual for regular).
    store_name  : CSV STORE_NAME value.
    ship_to     : Ship-to Code from Address Master.
    cust_no     : Customer Number (= ship_to for H&G).
    ean         : EAN resolved via HG Master. Empty if lookup failed.
    description : Product description from Items Master.
    mrp         : MRP from Items Master (0.0 if missing).
    input_mrp   : MRP supplied in the input CSV for source/master comparison.
    source_files: Source CSV filename(s) contributing to this consolidated line.
    unit_price  : 0.54 for tester; 0.0 (written blank) for regular.
    line_no     : 1000, 2000,... resets per SO (not per PO, so regular and
                  tester versions of the same PO each have their own line sequence).
    is_tester   : True when row was generated in tester mode.
    status      : "OK" if all lookups succeeded; "WARN" if any failed.
    """
    po_number:   str
    so_number:   str
    row_type:    str          # "REGULAR" or "TESTER"
    sku_code:    str
    item_no:     str
    qty:         int
    store_name:  str
    ship_to:     str
    cust_no:     str
    ean:         str
    description: str
    mrp:         float = 0.0
    input_mrp:   float = 0.0
    source_files: str  = ""
    unit_price:  float = 0.0
    line_no:     int   = 0
    is_tester:   bool  = False
    status:      str   = "OK"


@dataclass
class ProcessingResult:
    """
    Everything produced by one process_csv_files() call.

    Stateless container — created once per Generate click, passed to writer.

    Fields
    ------
    rows            : All SORow objects (regular + tester interleaved by PO).
    warnings        : (po_number, sku_code, message) tuples.
    input_files     : CSV paths processed (for Raw Data source column).
    warehouse_code  : D365 Location Code (e.g. "PICK").
    warehouse_display: User label (e.g. "AHD").
    generate_tester : Whether tester SOs were also generated this run.
    hg_sequence     : HG counter AFTER this run (used for both regular and tester).
    regular_so_map  : {po_number → regular_so_number}
    tester_so_map   : {po_number → tester_so_number}  (empty if no tester)
    po_summary      : {po_number → summary_dict} for Summary sheet.
                      summary_dict keys: regular_so, tester_so, store,
                      ship_to, cust_no, items, regular_qty, tester_qty,
                      status, warnings_count.
    input_po_count  : Unique valid PO count after input validation.
    input_line_count: Unique valid (PO, SKU) count after duplicate consolidation.
    input_qty_total : Total valid regular quantity after duplicate consolidation.
    skipped_rows    : Count of invalid input rows skipped with warnings.
    input_rows_read : Count of physical data rows read from CSV files.
    input_file_events: File-level audit entries where rows could not be audited.
    control_checks  : Non-blocking reconciliation results written to Excel.
    input_audit     : One trace record per source CSV row, plus file-level
                      errors where a CSV cannot be read.
    """
    rows:             List[SORow]               = field(default_factory=list)
    warnings:         List[Tuple[str, str, str]] = field(default_factory=list)
    input_files:      List[str]                 = field(default_factory=list)
    warehouse_code:   str                       = "PICK"
    warehouse_display: str                      = "AHD"
    generate_tester:  bool                      = False
    hg_sequence:      int                       = 0
    regular_so_map:   Dict[str, str]            = field(default_factory=dict)
    tester_so_map:    Dict[str, str]            = field(default_factory=dict)
    po_summary:       Dict[str, dict]           = field(default_factory=dict)
    input_po_count:   int                       = 0
    input_line_count: int                       = 0
    input_qty_total:  int                       = 0
    skipped_rows:     int                       = 0
    input_rows_read:  int                       = 0
    input_file_events: int                      = 0
    control_checks:   List[Tuple[str, str, str, str, str]] = field(default_factory=list)
    input_audit:      List[Dict[str, object]]    = field(default_factory=list)


# ==============================================================================
# SECTION 6 — MASTER FILE LOADERS
# ==============================================================================

class HGMasterLoader:
    """
    Loads SKU → EAN mapping from the HG Master Excel file.

    Sheet priority:
      1. Sheet named 'HG SKU MASTER' (case-insensitive).
      2. Fallback: first sheet with header auto-detected (scans 6 rows).

    Column detection:
      SKU: any of  sku_code | sku code | sku  (case-insensitive).
      EAN: any of  enn code | ean | gtin | enn  (first match wins —
           handles your duplicate 'ENN code' columns by always taking col C).

    Blank EAN rows (e.g. rows 141-147 in your master) are skipped with
    a WARNING log entry — they do not crash the load.

    Attributes:
        sku_to_ean  : {str(sku) → str(ean)}
        source_path : Last successfully loaded file.
        last_mtime  : os.path.getmtime at load time (used by watcher).
    """

    def __init__(self):
        self.sku_to_ean:  Dict[str, str] = {}
        self.source_path: Optional[Path] = None
        self.last_mtime:  float          = 0.0

    def load(self, path: Path) -> int:
        """
        Load SKU→EAN mappings. Returns count of valid pairs loaded.
        Raises FileNotFoundError or ValueError on failure.
        """
        log.info(f"[HG Master] Loading: {path.name}")
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        self.source_path = path
        self.last_mtime  = os.path.getmtime(path)

        # Inspect available sheets.
        try:
            xls = pd.ExcelFile(path)
            sheets = [s.strip() for s in xls.sheet_names]
            log.debug(f"[HG Master] Sheets found: {sheets}")
        except Exception as e:
            raise ValueError(f"Cannot open Excel: {e}")

        # Choose sheet: prefer 'HG SKU MASTER', fall back to auto-detect.
        target = next(
            (s for s in sheets if s.lower() in ('hg sku master', 'hg_sku_master')), None
        )

        if target:
            try:
                df = pd.read_excel(path, sheet_name=target, header=0)
                log.info(f"[HG Master] Using sheet '{target}'")
            except Exception as e:
                raise ValueError(f"Cannot read sheet '{target}': {e}")
        else:
            # Auto-detect header row by scanning first 6 rows.
            try:
                df_scan = pd.read_excel(path, header=None, nrows=6)
            except Exception as e:
                raise ValueError(f"Cannot scan Excel: {e}")

            header_row = None
            for i, row in df_scan.iterrows():
                vals = [str(v).strip().lower() for v in row.values]
                if any(v in ('sku_code', 'sku code', 'sku') for v in vals):
                    header_row = i
                    log.info(f"[HG Master] Header auto-detected at row {i} (Excel row {i+1})")
                    break

            if header_row is None:
                log.warning("[HG Master] Header not found — defaulting to row 0")
                header_row = 0

            try:
                df = pd.read_excel(path, header=header_row)
            except Exception as e:
                raise ValueError(f"Cannot read with header={header_row}: {e}")

        df.columns = [str(c).strip() for c in df.columns]
        log.debug(f"[HG Master] Columns: {list(df.columns)} | Rows: {len(df)}")

        # Locate SKU column.
        sku_col = next(
            (c for c in df.columns if c.lower() in ('sku_code', 'sku code', 'sku')), None
        )
        if not sku_col:
            raise ValueError(
                f"No SKU column found.\nAvailable: {list(df.columns)}\n"
                f"Expected: sku_code | sku code | sku"
            )
        log.info(f"[HG Master] SKU column: '{sku_col}'")

        # Locate EAN column — FIRST match only (handles duplicate 'ENN code').
        # Stripping pandas suffix '.1'/'.2' before comparing.
        ean_col = next(
            (c for c in df.columns
             if c.lower().split('.')[0].strip() in ('enn code', 'ean', 'gtin', 'enn', 'ean code')),
            None
        )
        if not ean_col:
            raise ValueError(
                f"No EAN column found.\nAvailable: {list(df.columns)}\n"
                f"Expected: ENN code | EAN | GTIN | ENN"
            )
        log.info(f"[HG Master] EAN column: '{ean_col}' (first match — duplicates ignored)")

        # Build SKU → EAN dict.
        self.sku_to_ean.clear()
        loaded = 0
        skipped_blank_ean = 0
        skipped_blank_sku = 0

        for idx, row in df.iterrows():
            raw_sku = row[sku_col]
            raw_ean = row[ean_col]

            if pd.isna(raw_sku):
                skipped_blank_sku += 1
                continue
            sku = str(raw_sku).strip()
            if not sku or sku.lower() == 'nan':
                skipped_blank_sku += 1
                continue

            # Blank EAN — skip gracefully with a WARNING (not a crash).
            if pd.isna(raw_ean) or str(raw_ean).strip().lower() in ('', 'nan'):
                skipped_blank_ean += 1
                log.warning(f"[HG Master] Row {idx+2}: SKU '{sku}' has blank EAN — skipped")
                continue

            # Normalise EAN type (float, int, str).
            if isinstance(raw_ean, float):
                if raw_ean != raw_ean:   # NaN check (NaN != NaN is the only such float)
                    skipped_blank_ean += 1
                    continue
                ean = str(int(raw_ean)) if raw_ean == int(raw_ean) else str(raw_ean)
            elif isinstance(raw_ean, int):
                ean = str(raw_ean)
            else:
                ean = str(raw_ean).strip()
                if ean.endswith('.0'):
                    ean = ean[:-2]

            if not ean or ean.lower() == 'nan':
                skipped_blank_ean += 1
                continue

            self.sku_to_ean[sku] = ean
            loaded += 1

        log.info(
            f"[HG Master] Done: loaded={loaded} | "
            f"blank_EAN={skipped_blank_ean} | blank_SKU={skipped_blank_sku}"
        )
        if loaded == 0:
            raise ValueError(
                f"No valid SKU→EAN mappings found.\n"
                f"SKU col='{sku_col}', EAN col='{ean_col}'\n"
                f"All cols: {list(df.columns)}"
            )
        log.debug(f"[HG Master] Sample (first 5): {list(self.sku_to_ean.items())[:5]}")
        return loaded


class ItemsMasterLoader:
    """
    Loads EAN → D365 Item details from the Items Master Excel file.

    Sheet priority:
      1. Sheet named 'Item Master' (case-insensitive).
      2. Fallback: first sheet.

    Required columns: GTIN, No., Description, Mrp, GST Group Code

    Attributes:
        ean_to_item : {str(EAN) → {item_no, description, mrp, gst_code}}
        source_path : Last loaded file.
        last_mtime  : Mtime at last load.
    """

    def __init__(self):
        self.ean_to_item:  Dict[str, Dict] = {}
        self.source_path:  Optional[Path]  = None
        self.last_mtime:   float           = 0.0

    def load(self, path: Path) -> int:
        """Load EAN→Item mappings. Returns count loaded."""
        log.info(f"[Items Master] Loading: {path.name}")
        if not path.exists():
            raise FileNotFoundError(f"Not found: {path}")

        self.source_path = path
        self.last_mtime  = os.path.getmtime(path)

        try:
            xls    = pd.ExcelFile(path)
            sheets = [s.strip() for s in xls.sheet_names]
        except Exception as e:
            raise ValueError(f"Cannot open Excel: {e}")

        target = next((s for s in sheets if s.lower() == 'item master'), None)
        try:
            df = pd.read_excel(path, sheet_name=target, header=0) if target \
                 else pd.read_excel(path, header=0)
        except Exception as e:
            raise ValueError(f"Cannot read Items Master: {e}")

        log.debug(f"[Items Master] Columns: {list(df.columns)} | Rows: {len(df)}")

        for col in ('GTIN', 'No.'):
            if col not in df.columns:
                raise ValueError(
                    f"Missing column '{col}'.\nAvailable: {list(df.columns)}"
                )

        df['GTIN_str'] = (
            df['GTIN'].astype(str).str.strip()
                      .str.replace(r'\.0$', '', regex=True)
        )

        self.ean_to_item.clear()
        for _, row in df.iterrows():
            ean     = row['GTIN_str']
            item_no = str(row['No.']).strip()
            desc    = str(row.get('Description', '')) if pd.notna(row.get('Description')) else ''
            mrp     = float(row['Mrp']) if pd.notna(row.get('Mrp')) else 0.0
            gst_raw = row.get('GST Group Code')
            gst     = str(gst_raw).strip() if pd.notna(gst_raw) else ''
            self.ean_to_item[ean] = {
                'item_no': item_no, 'description': desc,
                'mrp': mrp, 'gst_code': gst,
            }

        log.info(f"[Items Master] Loaded {len(self.ean_to_item)} EAN mappings")
        return len(self.ean_to_item)


class AddressMasterLoader:
    """
    Loads Store Name → Ship-to / Customer Number from H&G Addresses Excel.

    Sheet priority:
      1. 'Ship-To B2B'
      2. 'Ship-to Address List'

    Column detection (flexible):
      Store:   Del Location | Del_Location | Name | Store | Location Name
      Ship-to: Ship to | Ship-to | Ship_to | ShipTo
      Cust No: Cust No | Cust_No | Customer  (falls back to ship_to if absent)

    Attributes:
        store_to_ship : {str(store_name) → {ship_to, cust_no}}
        source_path   : Last loaded file.
        last_mtime    : Mtime at last load.
    """

    def __init__(self):
        self.store_to_ship: Dict[str, Dict] = {}
        self.source_path:   Optional[Path]  = None
        self.last_mtime:    float           = 0.0

    def load(self, path: Path) -> int:
        """Load Store→ShipTo mappings. Returns count loaded."""
        log.info(f"[Address Master] Loading: {path.name}")
        if not path.exists():
            raise FileNotFoundError(f"Not found: {path}")

        self.source_path = path
        self.last_mtime  = os.path.getmtime(path)

        try:
            xls    = pd.ExcelFile(path)
            sheets = [s.strip() for s in xls.sheet_names]
        except Exception as e:
            raise ValueError(f"Cannot open Excel: {e}")

        # Find the best sheet.
        target = None
        for name in ('ship-to b2b', 'ship-to address list', 'shipto', 'addresses'):
            target = next((s for s in sheets if s.strip().lower() == name), None)
            if target:
                break
        if not target:
            raise ValueError(
                f"Cannot find sheet. Available: {sheets}\n"
                f"Expected: 'Ship-To B2B' or 'Ship-to Address List'"
            )

        try:
            df = pd.read_excel(path, sheet_name=target, header=0)
            log.info(f"[Address Master] Using sheet '{target}'")
        except Exception as e:
            raise ValueError(f"Cannot read sheet '{target}': {e}")

        log.debug(f"[Address Master] Columns: {list(df.columns)} | Rows: {len(df)}")
        cols_lower = {str(c).strip().lower(): c for c in df.columns}

        # Flexible column detection.
        store_col = next(
            (cols_lower[k] for k in
             ('del location', 'del_location', 'del-location', 'name', 'store',
              'location name', 'location_name')
             if k in cols_lower), None
        )
        ship_col = next(
            (cols_lower[k] for k in
             ('ship to', 'ship-to', 'ship_to', 'shipto', 'ship to code', 'code')
             if k in cols_lower), None
        )
        cust_col = next(
            (cols_lower[k] for k in ('cust no', 'cust_no', 'custno', 'customer', 'cust')
             if k in cols_lower), None
        )

        if not store_col or not ship_col:
            raise ValueError(
                f"Missing store or ship-to column.\n"
                f"Available: {list(df.columns)}\n"
                f"Expected store: Del Location / Name\n"
                f"Expected ship-to: Ship to / Code"
            )
        log.info(
            f"[Address Master] store='{store_col}' | "
            f"ship_to='{ship_col}' | cust='" + str(cust_col or '(none, using ship_to)') + "'"
        )

        self.store_to_ship.clear()
        for _, row in df.iterrows():
            store   = str(row[store_col]).strip() if pd.notna(row.get(store_col)) else ''
            ship_to = str(row[ship_col]).strip()  if pd.notna(row.get(ship_col))  else ''
            cust_no = str(row[cust_col]).strip()  if cust_col and pd.notna(row.get(cust_col)) else ''

            if not cust_no:
                cust_no = ship_to   # H&G default: Customer No = Ship-to Code

            if not store or store.lower() == 'nan' or not ship_to:
                continue

            self.store_to_ship[store] = {'ship_to': ship_to, 'cust_no': cust_no}

        log.info(f"[Address Master] Loaded {len(self.store_to_ship)} store mappings")
        return len(self.store_to_ship)


# ==============================================================================
# SECTION 7 — BACKGROUND MASTER-FILE WATCHER
# ==============================================================================

class MasterWatcher(threading.Thread):
    """
    Daemon thread that silently reloads master files when they change on disk.

    Design:
    - daemon=True so it dies automatically when the main window closes.
    - Uses threading.Event for clean shutdown via stop().
    - All GUI updates scheduled via root.after(0, fn) — never touches Tkinter
      directly from the background thread (Tkinter is not thread-safe).
    - 2-second startup delay lets the GUI fully initialise first.
    - Polls DATA_MT_DIR every MASTER_WATCH_INTERVAL_SEC seconds.
    - Each loader is checked independently — a change in one file doesn't
      force reload of the others.
    """

    def __init__(self, app: 'MTSelectApp'):
        super().__init__(daemon=True)
        self.app   = app
        self._stop = threading.Event()

    def stop(self):
        """Signal the watcher to exit cleanly."""
        self._stop.set()

    def run(self):
        log.debug("[Watcher] Started")
        time.sleep(2)   # wait for GUI init
        while not self._stop.is_set():
            self._check_all()
            self._stop.wait(timeout=MASTER_WATCH_INTERVAL_SEC)
        log.debug("[Watcher] Stopped")

    def _check_all(self):
        """Check each master loader for file changes and reload if needed."""
        checks = [
            (self.app.hg_master,      "HG Master*.xlsx",      self.app.hg_path_var,      "HG Master",      lambda f, n: f"{f.name}  ({n} SKUs) ✓ [auto-reloaded]"),
            (self.app.items_master,   "Items*.xlsx",           self.app.items_path_var,   "Items Master",   lambda f, n: f"{f.name}  ({n} EANs) ✓ [auto-reloaded]"),
            (self.app.address_master, "H&G Addresses*.xlsx",   self.app.address_path_var, "Address Master", lambda f, n: f"{f.name}  ({n} stores) ✓ [auto-reloaded]"),
        ]
        for loader, pattern, label_var, tag, fmt in checks:
            files = sorted(DATA_MT_DIR.glob(pattern), reverse=True)
            if not files:
                continue
            f = files[0]
            try:
                mtime = os.path.getmtime(f)
            except OSError:
                continue
            if mtime <= loader.last_mtime:
                continue
            log.info(f"[Watcher] {tag} changed — reloading {f.name}")
            try:
                n     = loader.load(f)
                label = fmt(f, n)
                def _gui(lv=label_var, lbl=label):
                    lv.set(lbl)
                    self.app._update_master_status()
                self.app.root.after(0, _gui)
            except Exception as e:
                log.error(f"[Watcher] {tag} reload failed: {e}")


# ==============================================================================
# SECTION 8 — CSV PROCESSING ENGINE
# ==============================================================================

def _resolve_sku(sku: str, po: str,
                 hg_master: HGMasterLoader,
                 items_master: ItemsMasterLoader,
                 warnings: List[Tuple[str, str, str]]) -> Tuple[str, str, str, float, str]:
    """
    Resolve a single SKU to its D365 item details.

    Pipeline: SKU → HG Master → EAN → Items Master → item_no, description, mrp, gst

    Returns:
        (ean, item_no, description, mrp, status)
        status is "OK" or "WARN".
    """
    ean = hg_master.sku_to_ean.get(sku)
    if not ean:
        msg = f"SKU '{sku}' not in HG Master — no EAN mapping"
        log.warning(f"[Process] PO={po}: {msg}")
        warnings.append((po, sku, msg))
        return "", "", "", 0.0, "WARN"

    info = items_master.ean_to_item.get(ean)
    if not info:
        msg = f"EAN '{ean}' (SKU '{sku}') not in Items Master"
        log.warning(f"[Process] PO={po}: {msg}")
        warnings.append((po, sku, msg))
        return ean, "", "", 0.0, "WARN"

    return (
        ean,
        info['item_no'],
        info.get('description', ''),
        info.get('mrp', 0.0),
        "OK"
    )


def process_csv_files(
    file_paths:      List[str],
    hg_master:       HGMasterLoader,
    items_master:    ItemsMasterLoader,
    address_master:  AddressMasterLoader,
    warehouse_code:  str,
    generate_tester: bool,
    sequences:       Dict[str, object]
) -> ProcessingResult:
    """
    Core processing engine.

    For each unique PO:
      1. Assign a regular SO number from the shared daily sequence.
      2. If generate_tester=True, assign the paired tester SO using the same
         numeric sequence with the ``SO/HG/TT/`` prefix.
      3. For each SKU in the PO:
           - Resolve SKU → EAN → Item No + MRP + description.
           - Create a regular SORow (actual qty, blank price).
           - If tester: create a second SORow (qty=1, price=0.54).
      4. Build po_summary showing BOTH the regular and tester SOs.

    Rows are grouped by PO. When testers are enabled, each SKU contributes its
    regular row followed by its paired tester row.

    Args:
        file_paths:      Absolute paths to H&G CSV files.
        hg_master:       Loaded HGMasterLoader.
        items_master:    Loaded ItemsMasterLoader.
        address_master:  Loaded AddressMasterLoader.
        warehouse_code:  D365 Location Code (e.g. "PICK").
        generate_tester: True → also generate tester SOs for every PO.
        sequences:       Mapping containing the shared ``HG`` starting counter.
                         It is incremented once per PO and is not mutated here.

    Returns:
        Populated ProcessingResult.
    """
    result = ProcessingResult()
    result.input_files      = file_paths
    result.warehouse_code   = warehouse_code
    result.warehouse_display = next(
        (k for k, v in WAREHOUSES.items() if v == warehouse_code), "AHD"
    )
    result.generate_tester = generate_tester
    result.hg_sequence     = int(sequences["HG"])

    log.info(
        f"[Process] Files={len(file_paths)} | "
        f"Warehouse={warehouse_code} | "
        f"Tester={'YES — tester SOs use same sequence as regular' if generate_tester else 'NO — regular only'}"
    )
    log.info(f"[Process] Shared SO seq starting at {sequences['HG']}")

    # ------------------------------------------------------------------
    # Phase 1: Read all CSVs into a flat list
    # ------------------------------------------------------------------
    all_rows:    List[dict] = []
    seen_po_sku: set        = set()

    for fp in file_paths:
        source_name = os.path.basename(fp)
        log.info(f"[Process] Reading: {source_name}")
        try:
            df = pd.read_csv(fp)
        except Exception as e:
            msg = f"Cannot read {source_name}: {e}"
            log.error(f"[Process] {msg}")
            result.warnings.append(("", "", msg))
            result.input_audit.append({
                'source_file': source_name, 'source_row': "",
                'po': "", 'store': "", 'sku': "", 'input_qty': "",
                'input_mrp': "", 'disposition': "FILE ERROR",
                'reason': msg, 'regular_so': "", 'tester_so': "",
            })
            result.input_file_events += 1
            continue

        log.debug(f"[Process] Cols={list(df.columns)} | Rows={len(df)}")
        result.input_rows_read += len(df)
        required = ['PO_NO', 'STORE_NAME', 'SKU_CODE', 'QUANTITY', 'MRP']
        missing  = [c for c in required if c not in df.columns]
        if missing:
            msg = f"{source_name}: missing {missing}. Got: {list(df.columns)}"
            log.error(f"[Process] {msg}")
            result.warnings.append(("", "", msg))
            for row_idx, row in df.iterrows():
                result.input_audit.append({
                    'source_file': source_name, 'source_row': row_idx + 2,
                    'po': "" if pd.isna(row.get('PO_NO')) else str(row.get('PO_NO', '')).strip(),
                    'store': "" if pd.isna(row.get('STORE_NAME')) else str(row.get('STORE_NAME', '')).strip(),
                    'sku': "" if pd.isna(row.get('SKU_CODE')) else str(row.get('SKU_CODE', '')).strip(),
                    'input_qty': "" if pd.isna(row.get('QUANTITY')) else row.get('QUANTITY', ''),
                    'input_mrp': "" if pd.isna(row.get('MRP')) else row.get('MRP', ''),
                    'disposition': "SKIPPED", 'reason': msg,
                    'regular_so': "", 'tester_so': "",
                })
                result.skipped_rows += 1
            if df.empty:
                result.input_audit.append({
                    'source_file': source_name, 'source_row': "",
                    'po': "", 'store': "", 'sku': "", 'input_qty': "",
                    'input_mrp': "", 'disposition': "FILE ERROR",
                    'reason': msg, 'regular_so': "", 'tester_so': "",
                })
                result.input_file_events += 1
            continue

        for row_idx, row in df.iterrows():
            po    = str(row['PO_NO']).strip()
            store = str(row['STORE_NAME']).strip()
            sku   = str(row['SKU_CODE']).strip()
            audit = {
                'source_file': source_name, 'source_row': row_idx + 2,
                'po': "" if po == 'nan' else po,
                'store': "" if store == 'nan' else store,
                'sku': "" if sku == 'nan' else sku,
                'input_qty': "" if pd.isna(row['QUANTITY']) else row['QUANTITY'],
                'input_mrp': "" if pd.isna(row['MRP']) else row['MRP'],
                'disposition': "", 'reason': "",
                'regular_so': "", 'tester_so': "",
            }
            try:
                qty = int(float(row['QUANTITY'])) if pd.notna(row['QUANTITY']) else 0
            except (ValueError, TypeError):
                qty = 0

            if (not po or po == 'nan' or not store or store == 'nan'
                    or not sku or sku == 'nan' or qty <= 0):
                msg = (
                    f"{source_name} row {row_idx + 2}: invalid required input "
                    f"(PO={po}, Store={store}, SKU={sku}, Qty={qty}); row skipped"
                )
                log.warning(f"[Process] {msg}")
                result.warnings.append((po if po != 'nan' else "", sku if sku != 'nan' else "", msg))
                result.skipped_rows += 1
                audit['disposition'] = "SKIPPED"
                audit['reason'] = msg
                result.input_audit.append(audit)
                continue

            try:
                mrp = float(row['MRP']) if pd.notna(row.get('MRP')) else 0.0
            except (ValueError, TypeError):
                mrp = 0.0
                msg = f"{source_name} row {row_idx + 2}: invalid input MRP; processing continues using mapped item data"
                log.warning(f"[Process] {msg}")
                result.warnings.append((po, sku, msg))
                audit['reason'] = msg

            key = (po, sku)
            if key in seen_po_sku:
                msg = f"Duplicate (PO={po}, SKU={sku}) — qtys will be summed"
                log.warning(f"[Process] {msg}")
                result.warnings.append((po, sku, msg))
                audit['disposition'] = "CONSOLIDATED"
                audit['reason'] = f"{audit['reason']}; {msg}".strip("; ")
            else:
                audit['disposition'] = "PROCESSED"
            seen_po_sku.add(key)
            result.input_audit.append(audit)
            all_rows.append({
                'po': po, 'store': store, 'sku': sku, 'qty': qty, 'mrp': mrp,
                'source': source_name,
            })

    audit_counts: Dict[str, int] = {}
    for audit in result.input_audit:
        disposition = str(audit.get('disposition', 'UNLABELLED'))
        audit_counts[disposition] = audit_counts.get(disposition, 0) + 1
    log.info(f"[Process] Input Audit dispositions: {audit_counts}")
    log.info(f"[Process] Valid CSV rows: {len(all_rows)}")
    if not all_rows:
        result.warnings.append(("", "", "No valid rows found in any selected CSV"))
        return result

    # ------------------------------------------------------------------
    # Phase 2: Group by PO → SKU, summing quantities
    # ------------------------------------------------------------------
    po_groups: Dict[str, Dict[str, dict]] = {}
    po_stores: Dict[str, set] = {}
    for r in all_rows:
        po, sku = r['po'], r['sku']
        po_stores.setdefault(po, set()).add(r['store'])
        if po not in po_groups:
            po_groups[po] = {}
        if sku not in po_groups[po]:
            po_groups[po][sku] = {
                'qty': 0, 'store': r['store'], 'mrp': r['mrp'],
                'sources': set(), 'mrps': set(),
            }
        po_groups[po][sku]['qty'] += r['qty']
        po_groups[po][sku]['sources'].add(r['source'])
        po_groups[po][sku]['mrps'].add(r['mrp'])

    log.info(f"[Process] Unique POs: {len(po_groups)}")
    result.input_po_count = len(po_groups)
    result.input_line_count = sum(len(sku_dict) for sku_dict in po_groups.values())
    result.input_qty_total = sum(
        details['qty']
        for sku_dict in po_groups.values()
        for details in sku_dict.values()
    )
    log.info(
        f"[Process] Will generate: "
        f"{len(po_groups)} regular SOs"
        + (f" + {len(po_groups)} tester SOs = {len(po_groups)*2} total" if generate_tester else "")
    )

    # ------------------------------------------------------------------
    # Phase 3: Assign SOs and resolve all lookups
    # ------------------------------------------------------------------
    hg_seq = int(sequences["HG"])

    result.regular_so_map = {}
    result.tester_so_map  = {}
    result.po_summary     = {}

    for po, sku_dict in po_groups.items():
        po_data_warn_count = 0
        if len(po_stores.get(po, set())) > 1:
            msg = f"PO '{po}' contains multiple store names: {sorted(po_stores[po])}; first mapped store used"
            log.warning(f"[Process] {msg}")
            result.warnings.append((po, "", msg))
            po_data_warn_count += 1

        # Assign regular SO (always).
        hg_seq += 1
        regular_so = generate_so_number(hg_seq, is_tester=False)
        result.regular_so_map[po] = regular_so
        result.hg_sequence = hg_seq

        # Assign the paired tester SO, using the same number as the regular SO.
        tester_so = ""
        if generate_tester:
            tester_so = generate_so_number(hg_seq, is_tester=True)
            result.tester_so_map[po] = tester_so
        for audit in result.input_audit:
            if audit['disposition'] in ("PROCESSED", "CONSOLIDATED") and audit['po'] == po:
                audit['regular_so'] = regular_so
                audit['tester_so'] = tester_so

        log.debug(
            f"[Process] PO={po} -> Regular={regular_so}"
            + (f" | Tester={tester_so}" if generate_tester else "")
        )

        # Resolve store → ship-to.
        store_name = next(iter(sku_dict.values()))['store']
        addr_info  = address_master.store_to_ship.get(store_name)
        if not addr_info:
            msg = f"Store '{store_name}' not in Address Master — Ship-to/Cust blank"
            log.warning(f"[Process] PO={po}: {msg}")
            result.warnings.append((po, "", msg))
            ship_to = cust_no = ""
        else:
            ship_to = addr_info['ship_to']
            cust_no = addr_info['cust_no']

        po_has_warn   = (ship_to == "" or po_data_warn_count > 0)
        regular_qty   = 0
        tester_qty    = 0
        po_items      = 0
        po_warn_count = (1 if ship_to == "" else 0) + po_data_warn_count

        for sku, details in sku_dict.items():
            if len(details['mrps']) > 1:
                msg = f"PO '{po}', SKU '{sku}' contains multiple input MRP values: {sorted(details['mrps'])}; first value used for comparison"
                log.warning(f"[Process] {msg}")
                result.warnings.append((po, sku, msg))
                po_has_warn = True
                po_warn_count += 1
            ean, item_no, description, item_mrp, status = _resolve_sku(
                sku, po, hg_master, items_master, result.warnings
            )
            if (status == "OK" and details['mrp'] and item_mrp
                    and abs(details['mrp'] - item_mrp) > 0.001):
                msg = (
                    f"Input MRP {details['mrp']} does not match Items Master MRP "
                    f"{item_mrp} for SKU '{sku}'"
                )
                log.warning(f"[Process] PO={po}: {msg}")
                result.warnings.append((po, sku, msg))
                status = "WARN"
            if status == "WARN":
                po_has_warn = True
                po_warn_count += 1

            po_items    += 1
            regular_qty += details['qty']

            # Regular SORow.
            result.rows.append(SORow(
                po_number   = po,
                so_number   = regular_so,
                row_type    = "REGULAR",
                sku_code    = sku,
                item_no     = item_no,
                qty         = details['qty'],
                store_name  = store_name,
                ship_to     = ship_to,
                cust_no     = cust_no,
                ean         = ean,
                description = description,
                mrp         = item_mrp,
                input_mrp   = details['mrp'],
                source_files = ", ".join(sorted(details['sources'])),
                unit_price  = 0.0,
                is_tester   = False,
                status      = status,
            ))

            # Tester SORow (only when requested).
            if generate_tester:
                tester_qty += 1
                result.rows.append(SORow(
                    po_number   = po,
                    so_number   = tester_so,
                    row_type    = "TESTER",
                    sku_code    = sku,
                    item_no     = item_no,
                    qty         = 1,            # always 1 per SKU for testers
                    store_name  = store_name,
                    ship_to     = ship_to,
                    cust_no     = cust_no,
                    ean         = ean,
                    description = description,
                    mrp         = item_mrp,
                    input_mrp   = details['mrp'],
                    source_files = ", ".join(sorted(details['sources'])),
                    unit_price  = TESTER_CP,    # always 0.54 for testers
                    is_tester   = True,
                    status      = status,
                ))

        # Build Summary row for this PO.
        result.po_summary[po] = {
            'regular_so':    regular_so,
            'tester_so':     tester_so if generate_tester else "N/A",
            'store':         store_name,
            'ship_to':       ship_to,
            'cust_no':       cust_no,
            'items':         po_items,
            'regular_qty':   regular_qty,
            'tester_qty':    tester_qty if generate_tester else 0,
            'status':        "WARN" if po_has_warn else "OK",
            'warnings_count': po_warn_count,
        }

    log.info(
        f"[Process] Complete: "
        f"total_rows={len(result.rows)} | "
        f"regular_SOs={len(result.regular_so_map)} | "
        f"tester_SOs={len(result.tester_so_map)} | "
        f"warnings={len(result.warnings)}"
    )
    return result


# ==============================================================================
# SECTION 9 — EXCEL OUTPUT WRITER
# ==============================================================================

# Shared style objects — created once, reused across all sheets.
HDR_FILL     = PatternFill("solid", fgColor="1A237E")   # dark navy header
HDR_FONT     = Font(bold=True, color="FFFFFF", size=10)
WARN_FILL    = PatternFill("solid", fgColor="FFF3E0")   # amber for warnings
OK_FILL      = PatternFill("solid", fgColor="E8F5E9")   # green for OK
TESTER_FILL  = PatternFill("solid", fgColor="E3F2FD")   # light blue for tester rows
REGULAR_FILL = PatternFill("solid", fgColor="FFFFFF")   # white for regular rows
BOLD_FONT    = Font(bold=True)
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
LEFT_ALIGN   = Alignment(horizontal='left',   vertical='center')
THIN_BORDER  = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'),  bottom=Side(style='thin'),
)


def _hdr(cell, value: str):
    """Apply standard header styling (dark navy, white bold, centred, bordered)."""
    cell.value     = value
    cell.font      = HDR_FONT
    cell.fill      = HDR_FILL
    cell.alignment = CENTER_ALIGN
    cell.border    = THIN_BORDER


def _autofit(ws, max_w: int = 45):
    """Auto-fit column widths, capped at max_w characters."""
    for col in ws.columns:
        letter  = col[0].column_letter
        max_len = max((len(str(c.value)) for c in col if c.value), default=8)
        ws.column_dimensions[letter].width = min(max_len + 3, max_w)


def _so_sequence_value(so_number: str) -> Optional[int]:
    """Return the trailing numeric SO sequence, or None for an invalid format."""
    try:
        return int(so_number.rsplit("/", 1)[-1])
    except (AttributeError, TypeError, ValueError):
        return None


def _build_control_checks(result: ProcessingResult, header_sos: set) -> None:
    """
    Build non-blocking reconciliation checks for review before D365 upload.

    Failed checks are added to the Warnings sheet and log, but they never stop
    workbook creation or replace uncertain mapped values with invented data.
    """
    result.warnings = [
        warning for warning in result.warnings
        if warning[1] != "CONTROL CHECK"
    ]
    processing_warning_count = len(result.warnings)
    checks: List[Tuple[str, str, str, str, str]] = []

    def add_check(name: str, expected: str, actual: str,
                  passed: bool, note: str = "") -> None:
        status = "PASS" if passed else "WARN"
        checks.append((name, expected, actual, status, note))
        if not passed:
            msg = f"{name}: expected {expected}; actual {actual}"
            if note:
                msg += f". {note}"
            result.warnings.append(("", "CONTROL CHECK", msg))
            log.warning(f"[Control Check] {msg}")

    regular_rows = [row for row in result.rows if not row.is_tester]
    tester_rows = [row for row in result.rows if row.is_tester]
    regular_sos = set(result.regular_so_map.values())
    tester_sos = set(result.tester_so_map.values())
    line_sos = {row.so_number for row in result.rows}
    month = datetime.now().strftime("%m")

    add_check(
        "Invalid input rows skipped",
        "0",
        str(result.skipped_rows),
        result.skipped_rows == 0,
        "Skipped rows remain listed in Warnings and are not written as SO lines."
    )
    disposition_counts: Dict[str, int] = {}
    for audit in result.input_audit:
        disposition = str(audit.get('disposition', '')).strip()
        disposition_counts[disposition] = disposition_counts.get(disposition, 0) + 1
    blank_dispositions = disposition_counts.get("", 0)
    disposition_summary = ", ".join(
        f"{key}={value}" for key, value in sorted(disposition_counts.items()) if key
    ) or "No input rows captured"
    add_check(
        "Input audit disposition completeness",
        "Every captured input row/file error labelled",
        disposition_summary,
        blank_dispositions == 0,
        "Review Input Audit for the outcome of each original CSV row."
    )
    expected_audit_rows = result.input_rows_read + result.input_file_events
    add_check(
        "Input source rows vs audit rows",
        str(expected_audit_rows),
        str(len(result.input_audit)),
        expected_audit_rows == len(result.input_audit),
        "All readable CSV data rows must have one Input Audit record."
    )
    add_check(
        "Input unique POs vs regular SOs",
        str(result.input_po_count),
        str(len(regular_sos)),
        result.input_po_count == len(regular_sos)
    )
    add_check(
        "Input PO/SKU lines vs regular lines",
        str(result.input_line_count),
        str(len(regular_rows)),
        result.input_line_count == len(regular_rows),
        "Duplicate PO/SKU input rows are expected to consolidate to one sales line."
    )
    regular_qty = sum(row.qty for row in regular_rows)
    add_check(
        "Input quantity vs regular output quantity",
        str(result.input_qty_total),
        str(regular_qty),
        result.input_qty_total == regular_qty
    )
    add_check(
        "Regular SO numbers unique",
        str(len(result.regular_so_map)),
        str(len(regular_sos)),
        len(result.regular_so_map) == len(regular_sos)
    )
    add_check(
        "Every sales line has a header",
        "No orphan lines",
        "No orphan lines" if line_sos <= header_sos else str(sorted(line_sos - header_sos)),
        line_sos <= header_sos
    )
    add_check(
        "Every header has sales lines",
        "No empty headers",
        "No empty headers" if header_sos <= line_sos else str(sorted(header_sos - line_sos)),
        header_sos <= line_sos
    )

    regular_numbers = [_so_sequence_value(so) for so in result.regular_so_map.values()]
    format_ok = all(
        so.startswith(f"SO/HG/{month}/") and number is not None
        for so, number in zip(result.regular_so_map.values(), regular_numbers)
    )
    add_check(
        "Regular SO format",
        f"SO/HG/{month}/<sequence>",
        "Valid" if format_ok else "Invalid format found",
        format_ok
    )
    continuous = all(
        current == previous + 1
        for previous, current in zip(regular_numbers, regular_numbers[1:])
        if previous is not None and current is not None
    ) and all(number is not None for number in regular_numbers)
    add_check(
        "Regular SO sequence within run",
        "Increment by 1 per PO",
        ", ".join(str(number) for number in regular_numbers) if regular_numbers else "No SO rows",
        continuous,
        "The first number may continue from an earlier run on the same date."
    )

    if result.generate_tester:
        expected_pairs = {
            po: generate_so_number(_so_sequence_value(regular_so), is_tester=True)
            for po, regular_so in result.regular_so_map.items()
            if _so_sequence_value(regular_so) is not None
        }
        add_check(
            "Tester SO count vs regular SO count",
            str(len(regular_sos)),
            str(len(tester_sos)),
            len(regular_sos) == len(tester_sos)
        )
        add_check(
            "Tester SO paired sequence",
            "Same trailing number as regular SO",
            "Valid" if expected_pairs == result.tester_so_map else "Mismatch found",
            expected_pairs == result.tester_so_map
        )
        add_check(
            "Tester line count vs regular line count",
            str(len(regular_rows)),
            str(len(tester_rows)),
            len(regular_rows) == len(tester_rows)
        )
        add_check(
            "Tester quantity",
            "1 on every tester line",
            "Valid" if all(row.qty == 1 for row in tester_rows) else "Invalid quantity found",
            all(row.qty == 1 for row in tester_rows)
        )
        add_check(
            "Tester unit price",
            str(TESTER_CP),
            "Valid" if all(row.unit_price == TESTER_CP for row in tester_rows) else "Invalid price found",
            all(row.unit_price == TESTER_CP for row in tester_rows)
        )
    else:
        add_check(
            "Tester lines when tester mode is off",
            "0",
            str(len(tester_rows)),
            len(tester_rows) == 0
        )

    unmapped_lines = sum(
        1 for row in result.rows
        if not row.ean or not row.item_no or not row.ship_to or not row.cust_no
    )
    add_check(
        "Required mapping fields populated",
        "No blank EAN/Item/Ship-to/Cust fields",
        f"{unmapped_lines} incomplete output line(s)",
        unmapped_lines == 0,
        "Unresolved fields are intentionally left blank; review Warnings before upload."
    )
    mrp_mismatches = sum(
        1 for row in regular_rows
        if row.input_mrp and row.mrp and abs(row.input_mrp - row.mrp) > 0.001
    )
    add_check(
        "Input MRP vs Items Master MRP",
        "No mismatches",
        f"{mrp_mismatches} mismatch(es)",
        mrp_mismatches == 0,
        "Input and mapped master MRP are shown separately in Validation and Raw Data."
    )

    bad_line_sequences = []
    for so_number in sorted(line_sos):
        actual = [row.line_no for row in result.rows if row.so_number == so_number]
        expected = list(range(1000, (len(actual) + 1) * 1000, 1000))
        if actual != expected:
            bad_line_sequences.append(so_number)
    add_check(
        "Line numbering per SO",
        "1000, 2000, ... reset for each SO",
        "Valid" if not bad_line_sequences else ", ".join(bad_line_sequences),
        not bad_line_sequences
    )
    add_check(
        "Processing warnings present",
        "0 warnings",
        str(processing_warning_count),
        processing_warning_count == 0,
        "Warnings may be informational, but should be reviewed before upload."
    )

    result.control_checks = checks
    failed = sum(1 for check in checks if check[3] == "WARN")
    log.info(f"[Control Check] Completed: {len(checks) - failed} PASS | {failed} WARN")


def write_output_workbook(result: ProcessingResult, output_path: Path) -> None:
    """
    Write the full 9-sheet D365-ready workbook with non-blocking control checks.

    Sheet 1 — Headers (SO):
        One row per unique SO number (both regular and tester).
        Tester rows highlighted in light blue for easy visual distinction.

    Sheet 2 — Lines (SO):
        One row per SKU line. Line numbers reset per SO (not per PO), so
        the regular and tester versions of the same PO each have independent
        1000/2000/... numbering.
        Tester rows highlighted in light blue.

    Sheet 3 — Summary:
        One row per PO showing BOTH the regular SO and tester SO (or N/A).
        Columns: PO | Regular SO | Tester SO | Store | Ship-to | Cust No |
                 Items | Regular Qty | Tester Qty | Status | Warnings Count
        Full colour-coded status so problems are immediately visible.

    Sheet 4 — SKU Pivot:
        One row per unique SKU code (not per line). Groups quantities by type.
        Columns: SKU Code | EAN | Item No | Description | MRP |
                 Regular Qty | Tester Qty | Total Qty
        For verification that quantities sum correctly across all POs.

    Sheet 5 — Validation:
        Every SORow with all resolved fields.
        TYPE column clearly shows REGULAR vs TESTER.
        Shows input CSV MRP beside Items Master MRP for comparison.
        STATUS column: OK (green) or WARN (amber).
        Allows complete pre-upload inspection line by line.

    Sheet 6 — Warnings:
        All non-fatal issues: missing SKUs/EANs/stores, duplicates, bad files.
        Green "No warnings" row if everything mapped successfully.

    Sheet 7 — Raw Data:
        Generated-line audit trail with source file, original SKU, EAN, final
        quantity and type. Used for post-upload debugging.

    Sheet 8 — Control Check:
        Reconciles input totals, SO/header/line structures, mapping completeness,
        sequence pairing, and tester rules. WARN results never block output.

    Sheet 9 — Input Audit:
        Lists each source CSV row exactly once with its processing disposition
        and reason, including rows skipped before SO generation.

    Args:
        result:      Populated ProcessingResult.
        output_path: Where to save the .xlsx file.
    """
    log.info(f"[Writer] Writing: {output_path.name}")
    today = datetime.now().strftime("%d-%m-%Y")

    wb = Workbook()
    wb.remove(wb.active)   # remove default empty sheet

    # -----------------------------------------------------------------------
    # SHEET 1: Headers (SO)
    # -----------------------------------------------------------------------
    ws_hdr = wb.create_sheet("Headers (SO)")
    hcols = [
        "Document Type",            # always "Order"
        "Document No.",             # SO number
        "Sell-to Customer No.",     # from Address Master
        "Ship-to Code",             # from Address Master
        "Posting Date",             # today
        "Order Date",               # today
        "Document Date",            # today
        "Invoice From Date",        # today
        "Invoice To Date",          # today
        "External Document No.",    # original PO number (for traceability)
        "Location Code",            # warehouse D365 code
        "Dimension Set ID",         # blank (D365 required column)
        "Supply Type",              # always "B2B"
        "Voucher Narration",        # blank
        "Brand Code (Dimension)",   # blank
        "Channel Code (Dimension)", # blank
        "Catagory (Dimension)",     # blank (typo preserved from D365 template)
        "Geography Code (Dimension)", # blank
    ]
    for c, h in enumerate(hcols, 1):
        _hdr(ws_hdr.cell(1, c), h)

    seen_so: set = set()
    r = 2
    for sorow in result.rows:
        if sorow.so_number in seen_so:
            continue
        seen_so.add(sorow.so_number)

        # External Document No: original PO for regular, "TESTERS" for tester SOs
        ext_doc_no = "TESTERS" if sorow.is_tester else sorow.po_number

        vals = [
            "Order", sorow.so_number, sorow.cust_no, sorow.ship_to,
            today, today, today, today, today,
            ext_doc_no,               # External Doc No = PO for regular, "TESTERS" for tester
            result.warehouse_code, "",
            "B2B",
            "", "", "", "", "",
        ]
        row_fill = TESTER_FILL if sorow.is_tester else REGULAR_FILL
        for c, v in enumerate(vals, 1):
            cell = ws_hdr.cell(r, c, v)
            cell.fill = row_fill

        r += 1

    ws_hdr.freeze_panes = "A2"
    _autofit(ws_hdr)

    # -----------------------------------------------------------------------
    # SHEET 2: Lines (SO)
    # -----------------------------------------------------------------------
    ws_line = wb.create_sheet("Lines (SO)")
    lcols = [
        "Document Type",   # always "Order"
        "Document No.",    # SO number
        "Line No.",        # 1000, 2000,... resets per SO
        "Type",            # always "Item"
        "No.",             # D365 Item Number
        "Location Code",   # warehouse code
        "Quantity",        # final qty (1 for tester; actual for regular)
        "Unit Price",      # 0.54 for tester; blank for regular
    ]
    for c, h in enumerate(lcols, 1):
        _hdr(ws_line.cell(1, c), h)

    r             = 2
    line_no_by_so: Dict[str, int] = {}

    for sorow in result.rows:
        # Regular and tester rows can be interleaved, so track each SO
        # independently rather than resetting based on the previous row.
        line_no = line_no_by_so.get(sorow.so_number, 0) + 1000
        line_no_by_so[sorow.so_number] = line_no
        sorow.line_no = line_no

        row_fill = TESTER_FILL if sorow.is_tester else REGULAR_FILL
        vals = [
            "Order", sorow.so_number, line_no, "Item",
            sorow.item_no, result.warehouse_code, sorow.qty,
            sorow.unit_price if sorow.is_tester else "",
        ]
        for c, v in enumerate(vals, 1):
            cell = ws_line.cell(r, c, v)
            cell.fill = row_fill

        r += 1

    ws_line.freeze_panes = "A2"
    _autofit(ws_line)

    # Checks depend on written header membership and assigned line numbers.
    _build_control_checks(result, seen_so)

    # -----------------------------------------------------------------------
    # SHEET 3: Summary
    # -----------------------------------------------------------------------
    # One row per PO with BOTH regular and tester SOs visible side by side.
    # This is the primary inspection sheet for the user before D365 upload.

    ws_sum = wb.create_sheet("Summary")
    scols = [
        "PO Number",      # original PO
        "Regular SO",     # e.g. SO/HG/05/24526
        "Tester SO",      # e.g. SO/HG/TT/24526 (or N/A)
        "Store",          # raw store name from CSV
        "Ship-to Code",   # from Address Master
        "Cust No",        # from Address Master
        "Items (SKUs)",   # unique SKU count in this PO
        "Regular Qty",    # sum of all quantities (regular)
        "Tester Qty",     # number of tester lines (= items count, each qty=1)
        "Status",         # OK / WARN
        "Warnings",       # count of warnings for this PO
    ]
    for c, h in enumerate(scols, 1):
        _hdr(ws_sum.cell(1, c), h)

    r = 2
    for po, info in result.po_summary.items():
        ws_sum.cell(r, 1,  po)
        ws_sum.cell(r, 2,  info['regular_so'])
        ws_sum.cell(r, 3,  info['tester_so'])     # "N/A" if tester not requested
        ws_sum.cell(r, 4,  info['store'])
        ws_sum.cell(r, 5,  info['ship_to'])
        ws_sum.cell(r, 6,  info['cust_no'])
        ws_sum.cell(r, 7,  info['items'])
        ws_sum.cell(r, 8,  info['regular_qty'])
        ws_sum.cell(r, 9,  info['tester_qty'])

        status_cell      = ws_sum.cell(r, 10, info['status'])
        status_cell.fill = OK_FILL if info['status'] == "OK" else WARN_FILL
        status_cell.font = BOLD_FONT

        warn_cell      = ws_sum.cell(r, 11, info['warnings_count'])
        warn_cell.fill = OK_FILL if info['warnings_count'] == 0 else WARN_FILL

        r += 1

    ws_sum.freeze_panes = "A2"
    _autofit(ws_sum)

    # -----------------------------------------------------------------------
    # SHEET 4: SKU Pivot Summary
    # -----------------------------------------------------------------------
    # Group by SKU code and sum quantities by type (regular vs tester).
    # Shows EAN, Item No, Description, and totals for verification.

    ws_sku = wb.create_sheet("SKU Pivot")
    pcols = [
        "SKU Code",       # original SKU from CSV
        "EAN",            # resolved EAN code
        "Item No",        # D365 Item Number
        "Description",    # from Items Master
        "MRP",            # from Items Master
        "Regular Qty",    # sum of regular quantities
        "Tester Qty",     # sum of tester quantities
        "Total Qty",      # regular + tester
    ]
    for c, h in enumerate(pcols, 1):
        _hdr(ws_sku.cell(1, c), h)

    # Build SKU pivot: {sku_code → {ean, item_no, description, mrp, regular_qty, tester_qty}}
    sku_pivot: Dict[str, Dict] = {}
    for sorow in result.rows:
        if sorow.sku_code not in sku_pivot:
            sku_pivot[sorow.sku_code] = {
                'ean': sorow.ean,
                'item_no': sorow.item_no,
                'description': sorow.description,
                'mrp': sorow.mrp if sorow.mrp else "",
                'regular_qty': 0,
                'tester_qty': 0,
            }
        if sorow.is_tester:
            sku_pivot[sorow.sku_code]['tester_qty'] += sorow.qty
        else:
            sku_pivot[sorow.sku_code]['regular_qty'] += sorow.qty

    r = 2
    for sku_code in sorted(sku_pivot.keys()):
        data = sku_pivot[sku_code]
        regular_qty = data['regular_qty']
        tester_qty = data['tester_qty']
        total_qty = regular_qty + tester_qty

        ws_sku.cell(r, 1, sku_code)
        ws_sku.cell(r, 2, data['ean'])
        ws_sku.cell(r, 3, data['item_no'])
        ws_sku.cell(r, 4, data['description'])
        ws_sku.cell(r, 5, data['mrp'])
        ws_sku.cell(r, 6, regular_qty)
        ws_sku.cell(r, 7, tester_qty)
        ws_sku.cell(r, 8, total_qty)

        r += 1

    ws_sku.freeze_panes = "A2"
    _autofit(ws_sku)

    # -----------------------------------------------------------------------
    # SHEET 5: Validation
    # -----------------------------------------------------------------------
    # Complete line-level inspection — every SORow in full detail.
    # TYPE column clearly separates REGULAR from TESTER rows.
    # Tester rows highlighted in blue so they're instantly recognisable.

    ws_val = wb.create_sheet("Validation")
    vcols = [
        "PO",          # original PO number
        "SO Number",   # assigned SO
        "TYPE",        # REGULAR or TESTER — critical for inspection
        "Item No",     # D365 Item Number
        "SKU Code",    # original SKU from CSV
        "EAN",         # resolved EAN
        "Description", # from Items Master
        "Input MRP",   # supplied in source CSV
        "Master MRP",  # from Items Master
        "Qty",         # final quantity
        "Unit Price",  # 0.54 for tester; blank for regular
        "Store",       # raw store name
        "Ship-to",     # resolved from Address Master
        "Cust No",     # resolved from Address Master
        "Status",      # OK / WARN
    ]
    for c, h in enumerate(vcols, 1):
        _hdr(ws_val.cell(1, c), h)

    r = 2
    for sorow in result.rows:
        row_fill = TESTER_FILL if sorow.is_tester else REGULAR_FILL
        vals = [
            sorow.po_number, sorow.so_number, sorow.row_type,
            sorow.item_no, sorow.sku_code, sorow.ean, sorow.description,
            sorow.input_mrp if sorow.input_mrp else "",
            sorow.mrp if sorow.mrp else "",
            sorow.qty,
            sorow.unit_price if sorow.is_tester else "",
            sorow.store_name, sorow.ship_to, sorow.cust_no,
            sorow.status,
        ]
        for c, v in enumerate(vals, 1):
            cell = ws_val.cell(r, c, v)
            cell.fill = row_fill

        # Status cell overrides row fill.
        ws_val.cell(r, 15).fill = OK_FILL if sorow.status == "OK" else WARN_FILL
        ws_val.cell(r, 15).font = BOLD_FONT

        r += 1

    ws_val.freeze_panes = "A2"
    _autofit(ws_val)

    # -----------------------------------------------------------------------
    # SHEET 6: Warnings
    # -----------------------------------------------------------------------
    ws_warn = wb.create_sheet("Warnings")
    for c, h in enumerate(["PO", "SKU / Item", "Warning Message"], 1):
        _hdr(ws_warn.cell(1, c), h)

    if result.warnings:
        for r_idx, (po, sku, msg) in enumerate(result.warnings, 2):
            ws_warn.cell(r_idx, 1, po)
            ws_warn.cell(r_idx, 2, sku)
            ws_warn.cell(r_idx, 3, msg)
            for c in range(1, 4):
                ws_warn.cell(r_idx, c).fill = WARN_FILL
    else:
        ws_warn.cell(2, 3, "No warnings — all SKUs, EANs, and stores mapped successfully ✓")
        ws_warn.cell(2, 3).fill = OK_FILL

    _autofit(ws_warn)

    # -----------------------------------------------------------------------
    # SHEET 7: Raw Data
    # -----------------------------------------------------------------------
    ws_raw = wb.create_sheet("Raw Data")
    rcols = [
        "Source File(s)", "PO", "Store", "SKU Code", "EAN",
        "Description", "Final Qty", "Input MRP", "Master MRP", "Unit Price",
        "SO Number", "Item No", "Warehouse", "TYPE", "Tester", "Status",
    ]
    for c, h in enumerate(rcols, 1):
        _hdr(ws_raw.cell(1, c), h)

    r = 2
    for sorow in result.rows:
        row_fill = TESTER_FILL if sorow.is_tester else REGULAR_FILL
        vals = [
            sorow.source_files, sorow.po_number, sorow.store_name, sorow.sku_code, sorow.ean,
            sorow.description, sorow.qty,
            sorow.input_mrp if sorow.input_mrp else "",
            sorow.mrp if sorow.mrp else "",
            sorow.unit_price if sorow.is_tester else "",
            sorow.so_number, sorow.item_no, result.warehouse_code,
            sorow.row_type,
            "YES" if sorow.is_tester else "NO",
            sorow.status,
        ]
        for c, v in enumerate(vals, 1):
            cell = ws_raw.cell(r, c, v)
            cell.fill = row_fill
        ws_raw.cell(r, 16).fill = OK_FILL if sorow.status == "OK" else WARN_FILL
        r += 1

    ws_raw.freeze_panes = "A2"
    _autofit(ws_raw)

    # -----------------------------------------------------------------------
    # SHEET 8: Control Check
    # -----------------------------------------------------------------------
    ws_control = wb.create_sheet("Control Check")
    control_cols = ["Check", "Expected", "Actual", "Result", "Review Note"]
    for c, h in enumerate(control_cols, 1):
        _hdr(ws_control.cell(1, c), h)

    for r, (name, expected, actual, status, note) in enumerate(result.control_checks, 2):
        values = [name, expected, actual, status, note]
        for c, value in enumerate(values, 1):
            ws_control.cell(r, c, value)
        status_cell = ws_control.cell(r, 4)
        status_cell.fill = OK_FILL if status == "PASS" else WARN_FILL
        status_cell.font = BOLD_FONT

    ws_control.freeze_panes = "A2"
    _autofit(ws_control, max_w=75)

    # -----------------------------------------------------------------------
    # SHEET 9: Input Audit
    # -----------------------------------------------------------------------
    ws_input = wb.create_sheet("Input Audit")
    input_cols = [
        "Source File", "Source Row", "PO", "Store", "SKU Code",
        "Input Qty", "Input MRP", "Disposition", "Reason",
        "Regular SO", "Tester SO",
    ]
    for c, h in enumerate(input_cols, 1):
        _hdr(ws_input.cell(1, c), h)

    for r, audit in enumerate(result.input_audit, 2):
        vals = [
            audit.get('source_file', ''), audit.get('source_row', ''),
            audit.get('po', ''), audit.get('store', ''), audit.get('sku', ''),
            audit.get('input_qty', ''), audit.get('input_mrp', ''),
            audit.get('disposition', ''), audit.get('reason', ''),
            audit.get('regular_so', ''), audit.get('tester_so', ''),
        ]
        for c, value in enumerate(vals, 1):
            ws_input.cell(r, c, value)
        disposition = audit.get('disposition', '')
        ws_input.cell(r, 8).fill = OK_FILL if disposition in ("PROCESSED", "CONSOLIDATED") else WARN_FILL
        ws_input.cell(r, 8).font = BOLD_FONT

    ws_input.freeze_panes = "A2"
    _autofit(ws_input, max_w=80)

    wb.save(output_path)
    log.info(
        f"[Writer] Saved: {len(seen_so)} SOs | "
        f"{len(result.rows)} lines | 9 sheets | {output_path.name}"
    )


# ==============================================================================
# SECTION 10 — GUI LOG HANDLER
# ==============================================================================

class GuiLogHandler(logging.Handler):
    """
    Appends log records to a Tkinter Text widget in real time.

    Thread-safe: uses widget.after(0, fn) so GUI mutations always happen
    on the main Tkinter thread, even when called from the watcher thread.
    """

    def __init__(self, text_widget: tk.Text):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record: logging.LogRecord):
        msg = self.format(record)
        def _append():
            self.text_widget.config(state='normal')
            self.text_widget.insert('end', msg + '\n')
            self.text_widget.see('end')
            self.text_widget.config(state='disabled')
        try:
            self.text_widget.after(0, _append)
        except Exception:
            pass


# ==============================================================================
# SECTION 11 — MAIN GUI APPLICATION
# ==============================================================================

class MTSelectApp:
    """
    Main Tkinter GUI for MT Select H&G Processor v4.0.

    Layout (top to bottom):
      1.  Title + subtitle
      2.  Sequence info banner (live — shows next SO numbers)
      3.  Warehouse dropdown + Tester checkbox
      4.  Master Files section (auto-loaded; Browse for override)
      5.  CSV Files section
      6.  Buttons: Generate | Open Output | Open Logs | Reload Masters
      7.  Status label (colour-coded)
      8.  Processing Log (scrollable, read-only, mirrors log file)

    Key behaviours:
    - Masters auto-loaded from data_mt/ on startup.
    - Browse copies the selected file into data_mt/ so future restarts
      auto-load it without needing to Browse again.
    - Background MasterWatcher reloads changed files every 30 s.
    - Sequence banner updates live on Tester checkbox toggle and after each run.
    - Output goes to <csv_folder>/output_mt/ next to the source CSVs.
    """

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("MT Select (Health & Glow) Processor  v4.0")
        self.root.geometry("1000x1100")
        self.root.resizable(True, True)

        self.csv_paths:   List[str]            = []
        self.last_output: Optional[Path]       = None
        self.last_result: Optional[ProcessingResult] = None

        self.warehouse_var     = tk.StringVar(value=DEFAULT_WAREHOUSE)
        self.tester_var        = tk.BooleanVar(value=False)
        self.status_var        = tk.StringVar(value="Initialising…")
        self.hg_path_var       = tk.StringVar(value="Not selected")
        self.items_path_var    = tk.StringVar(value="Not selected")
        self.address_path_var  = tk.StringVar(value="Not selected")
        self.master_status_var = tk.StringVar(value="Masters: Loading…")
        self.seq_var           = tk.StringVar(value="Loading sequence…")
        self.csv_var           = tk.StringVar(value="No files selected")

        self.hg_master      = HGMasterLoader()
        self.items_master   = ItemsMasterLoader()
        self.address_master = AddressMasterLoader()

        self._build_ui()

        gui_handler = GuiLogHandler(self.log_text)
        gui_handler.setLevel(logging.DEBUG)
        gui_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)-7s] %(message)s", "%H:%M:%S")
        )
        log.addHandler(gui_handler)

        log.info("=" * 60)
        log.info("MT Select (Health & Glow) Processor  v4.0  started")
        log.info(f"Script  : {SCRIPT_DIR}")
        log.info(f"data_mt : {DATA_MT_DIR}")
        log.info(f"Log     : {LOG_FILE}")
        log.info(f"Today's base sequence: {_todays_base_sequence()}")
        log.info("=" * 60)

        self._auto_load_masters()

        self.watcher = MasterWatcher(self)
        self.watcher.start()
        log.info(f"[Watcher] Started (interval={MASTER_WATCH_INTERVAL_SEC}s)")

    # --------------------------------------------------------------------------
    # UI BUILDER
    # --------------------------------------------------------------------------

    def _build_ui(self):
        tk.Label(self.root, text="MT Select  (Health & Glow)",
                 font=("Arial", 15, "bold")).pack(pady=(12, 2))
        tk.Label(self.root, text="CSV  →  D365 Sales Order Import  |  v4.0",
                 font=("Arial", 9), fg="gray").pack(pady=(0, 6))

        # Sequence banner
        seq_frame = tk.Frame(self.root, bg="#E3F2FD", relief='groove', bd=1)
        seq_frame.pack(fill='x', padx=20, pady=(0, 6))
        tk.Label(seq_frame, textvariable=self.seq_var,
                 font=("Consolas", 8), bg="#E3F2FD", fg="#0D47A1").pack(pady=5, padx=8)

        # Warehouse + Tester
        top = tk.Frame(self.root)
        top.pack(fill='x', padx=20, pady=4)
        tk.Label(top, text="Warehouse:", font=("Arial", 10, "bold")).pack(side='left')
        wh = ttk.Combobox(top, textvariable=self.warehouse_var,
                          values=list(WAREHOUSES.keys()), state='readonly', width=8)
        wh.pack(side='left', padx=8)
        wh_lbl = tk.Label(top, text=f"→ {WAREHOUSES[self.warehouse_var.get()]}",
                          font=("Arial", 9), fg="gray")
        wh_lbl.pack(side='left', padx=4)
        wh.bind('<<ComboboxSelected>>',
                lambda e: wh_lbl.config(text=f"→ {WAREHOUSES[self.warehouse_var.get()]}"))
        tk.Checkbutton(
            top,
            text="Tester Orders  (generates BOTH regular + tester SOs per PO)",
            variable=self.tester_var, font=("Arial", 9),
            command=self._refresh_seq_display
        ).pack(side='right')

        # Master Files
        mf = tk.LabelFrame(self.root,
                            text="Master Files  (auto-loaded from data_mt/  ·  Browse copies file to data_mt/)",
                            font=("Arial", 10, "bold"), padx=10, pady=8)
        mf.pack(fill='x', padx=20, pady=6)
        self._master_row(mf, "HG Master (SKU→EAN):",
                         self.hg_path_var, self._select_hg_master)
        self._master_row(mf, "Items Master (EAN→Item):",
                         self.items_path_var, self._select_items_master)
        self._master_row(mf, "Address Master (Store→ShipTo):",
                         self.address_path_var, self._select_address_master)
        tk.Label(mf, textvariable=self.master_status_var,
                 font=("Arial", 9), fg="blue").pack(anchor='w', pady=(4, 0))

        # CSV Files
        cf = tk.LabelFrame(self.root, text="Input CSV Files",
                           font=("Arial", 10, "bold"), padx=10, pady=8)
        cf.pack(fill='x', padx=20, pady=6)
        crow = tk.Frame(cf)
        crow.pack(fill='x')
        tk.Label(crow, text="CSV Files:", font=("Arial", 9)).pack(side='left')
        tk.Label(crow, textvariable=self.csv_var, font=("Arial", 9),
                 fg="blue", width=46, anchor='w').pack(side='left', padx=8)
        tk.Button(crow, text="Browse", command=self._select_csv_files).pack(side='right')

        # Buttons row 1
        bf1 = tk.Frame(self.root)
        bf1.pack(pady=8)
        tk.Button(bf1, text="▶  Generate SO", width=26,
                  font=("Arial", 10, "bold"), bg="#00C853", fg="white",
                  command=self.generate).pack(side='left', padx=6)
        self.open_btn = tk.Button(bf1, text="📂  Open Last Output", width=26,
                                  state=tk.DISABLED, command=self.open_last)
        self.open_btn.pack(side='left', padx=6)

        # Buttons row 2
        bf2 = tk.Frame(self.root)
        bf2.pack(pady=2)
        tk.Button(bf2, text="📂  Open Log Folder", width=26,
                  command=self.open_log_folder).pack(side='left', padx=6)
        tk.Button(bf2, text="🔄  Reload Masters", width=26,
                  command=self._auto_load_masters).pack(side='left', padx=6)

        # Status
        self.status_label = tk.Label(self.root, textvariable=self.status_var,
                                     font=("Arial", 10), fg="gray", wraplength=720)
        self.status_label.pack(pady=4)

        # Log panel
        lf = tk.LabelFrame(self.root,
                            text="Processing Log  ·  also saved to Logs/ folder",
                            font=("Arial", 9))
        lf.pack(fill='both', expand=True, padx=20, pady=(0, 12))
        sc = ttk.Scrollbar(lf, orient='vertical')
        sc.pack(side='right', fill='y')
        self.log_text = tk.Text(lf, height=14, font=("Consolas", 8),
                                state='disabled', wrap='word', yscrollcommand=sc.set)
        self.log_text.pack(fill='both', expand=True)
        sc.config(command=self.log_text.yview)

    def _master_row(self, parent, label, path_var, cmd):
        f = tk.Frame(parent)
        f.pack(fill='x', pady=3)
        tk.Label(f, text=label, font=("Arial", 9), width=30, anchor='w').pack(side='left')
        tk.Label(f, textvariable=path_var, font=("Arial", 9), fg="blue",
                 width=30, anchor='w').pack(side='left', padx=4)
        tk.Button(f, text="Browse", command=cmd).pack(side='right')

    # --------------------------------------------------------------------------
    # SEQUENCE DISPLAY
    # --------------------------------------------------------------------------

    def _refresh_seq_display(self):
        """
        Update the banner with the shared next SO number and optional tester pair.
        Reads SEQ_FILE fresh each time to reflect external edits.
        Called on: startup, Tester toggle, after each Generate run.
        """
        seqs  = load_sequences()
        month = datetime.now().strftime("%m")
        base  = _todays_base_sequence()
        shared = int(seqs["HG"])

        next_hg = f"SO/HG/{month}/{shared+1}"
        next_tt = f"SO/HG/TT/{shared+1}"

        if self.tester_var.get():
            self.seq_var.set(
                f"REGULAR + TESTER MODE  ·  "
                f"Next Regular: {next_hg}  ·  "
                f"Next Tester: {next_tt}  ·  "
                f"Shared seq={shared}  ·  "
                f"Today base={base}"
            )
        else:
            self.seq_var.set(
                f"REGULAR ONLY MODE  ·  "
                f"Next SO: {next_hg} (shared seq={shared})  ·  "
                f"Today base={base}"
            )

    # --------------------------------------------------------------------------
    # MASTER STATUS
    # --------------------------------------------------------------------------

    def _update_master_status(self) -> bool:
        """
        Update the master status banner. Returns True if all 3 masters ready.
        Called after each load and by the watcher thread (via root.after).
        """
        if (self.hg_master.sku_to_ean
                and self.items_master.ean_to_item
                and self.address_master.store_to_ship):
            self.master_status_var.set(
                f"Masters: Ready ✓  ·  "
                f"HG={len(self.hg_master.sku_to_ean)} SKUs  ·  "
                f"Items={len(self.items_master.ean_to_item)} EANs  ·  "
                f"Addresses={len(self.address_master.store_to_ship)} stores"
            )
            self.status_var.set("Ready — select CSV files and click ▶ Generate SO")
            self.status_label.config(fg="darkgreen")
            return True

        missing = []
        if not self.hg_master.sku_to_ean:         missing.append("HG Master")
        if not self.items_master.ean_to_item:     missing.append("Items Master")
        if not self.address_master.store_to_ship: missing.append("Address Master")
        self.master_status_var.set(f"Masters: Missing → {', '.join(missing)}")
        self.status_var.set("Place master files in data_mt/ or use Browse")
        self.status_label.config(fg="orange")
        return False

    def _fmt_mtime(self, path: Path) -> str:
        """Return formatted last-modified time string for a file."""
        try:
            t = datetime.fromtimestamp(os.path.getmtime(path))
            return f"  [updated {t.strftime('%Y-%m-%d %H:%M')}]"
        except Exception:
            return ""

    # --------------------------------------------------------------------------
    # AUTO-LOAD MASTERS
    # --------------------------------------------------------------------------

    def _auto_load_masters(self):
        """
        Scan DATA_MT_DIR for master files and load them.

        File matching (alphabetically last wins = most recent by name convention):
          HG Master:      HG Master*.xlsx
          Items Master:   Items*.xlsx
          Address Master: H&G Addresses*.xlsx

        Also bound to the Reload Masters button for manual refresh after
        dropping new files into data_mt/.
        """
        log.info(f"[Auto-load] Scanning: {DATA_MT_DIR}")
        if not DATA_MT_DIR.exists():
            log.warning(f"[Auto-load] data_mt/ not found at {DATA_MT_DIR}")
            self._update_master_status()
            self._refresh_seq_display()
            return

        for f in sorted(DATA_MT_DIR.glob("HG Master*.xlsx"), reverse=True):
            try:
                cnt = self.hg_master.load(f)
                self.hg_path_var.set(f"{f.name}  ({cnt} SKUs) ✓{self._fmt_mtime(f)}")
                log.info(f"[Auto-load] HG Master: {f.name} ({cnt} SKUs)")
                break
            except Exception as e:
                log.error(f"[Auto-load] HG Master failed ({f.name}): {e}")

        for f in sorted(DATA_MT_DIR.glob("Items*.xlsx"), reverse=True):
            try:
                cnt = self.items_master.load(f)
                self.items_path_var.set(f"{f.name}  ({cnt} EANs) ✓{self._fmt_mtime(f)}")
                log.info(f"[Auto-load] Items Master: {f.name} ({cnt} EANs)")
                break
            except Exception as e:
                log.error(f"[Auto-load] Items Master failed ({f.name}): {e}")

        for f in sorted(DATA_MT_DIR.glob("H&G Addresses*.xlsx"), reverse=True):
            try:
                cnt = self.address_master.load(f)
                self.address_path_var.set(f"{f.name}  ({cnt} stores) ✓{self._fmt_mtime(f)}")
                log.info(f"[Auto-load] Address Master: {f.name} ({cnt} stores)")
                break
            except Exception as e:
                log.error(f"[Auto-load] Address Master failed ({f.name}): {e}")

        self._update_master_status()
        self._refresh_seq_display()

    # --------------------------------------------------------------------------
    # BROWSE HANDLERS (copy to data_mt/ for persistence)
    # --------------------------------------------------------------------------

    def _copy_and_load(self, src_path: str, dest_prefix: str,
                       loader, path_var, label_suffix: str):
        """
        Copy a manually selected file into data_mt/ (so future restarts
        auto-load it), then load it immediately.

        Args:
            src_path:     Source file path chosen by user.
            dest_prefix:  Prefix for the filename stored in data_mt/.
            loader:       The master loader instance to use.
            path_var:     StringVar to update with result.
            label_suffix: e.g. "SKUs", "EANs", "stores".
        """
        src   = Path(src_path)
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        dest  = DATA_MT_DIR / f"{dest_prefix} {stamp}.xlsx"
        shutil.copy2(src, dest)
        log.info(f"[Browse] Copied {src.name} -> {dest.name}")
        cnt = loader.load(dest)
        path_var.set(f"{dest.name}  ({cnt} {label_suffix}) ✓{self._fmt_mtime(dest)}")
        self._update_master_status()
        messagebox.showinfo(
            "Loaded",
            f"Loaded {cnt} {label_suffix}.\nPersisted to: {dest.name}"
        )

    def _select_hg_master(self):
        p = filedialog.askopenfilename(
            title="Select HG Master (SKU→EAN)",
            initialdir=str(DATA_MT_DIR),
            filetypes=[("Excel files", "*.xlsx")]
        )
        if p:
            try:
                self._copy_and_load(p, "HG Master", self.hg_master,
                                    self.hg_path_var, "SKUs")
            except Exception as e:
                log.error(f"[Browse] HG Master: {e}")
                messagebox.showerror("Error", str(e))

    def _select_items_master(self):
        p = filedialog.askopenfilename(
            title="Select Items Master (EAN→Item)",
            initialdir=str(DATA_MT_DIR),
            filetypes=[("Excel files", "*.xlsx")]
        )
        if p:
            try:
                self._copy_and_load(p, "Items", self.items_master,
                                    self.items_path_var, "EANs")
            except Exception as e:
                log.error(f"[Browse] Items Master: {e}")
                messagebox.showerror("Error", str(e))

    def _select_address_master(self):
        p = filedialog.askopenfilename(
            title="Select Address Master (Store→ShipTo)",
            initialdir=str(DATA_MT_DIR),
            filetypes=[("Excel files", "*.xlsx")]
        )
        if p:
            try:
                self._copy_and_load(p, "H&G Addresses", self.address_master,
                                    self.address_path_var, "stores")
            except Exception as e:
                log.error(f"[Browse] Address Master: {e}")
                messagebox.showerror("Error", str(e))

    def _select_csv_files(self):
        paths = filedialog.askopenfilenames(
            title="Select H&G PO CSV files",
            filetypes=[("CSV files", "*.csv")]
        )
        if paths:
            self.csv_paths = list(paths)
            self.csv_var.set(f"{len(self.csv_paths)} file(s) selected")
            log.info(f"[CSV] Selected: {[os.path.basename(p) for p in self.csv_paths]}")

    # --------------------------------------------------------------------------
    # GENERATE SO — main action
    # --------------------------------------------------------------------------

    def generate(self):
        """
        Generate Sales Orders from selected CSV files.

        Steps:
        1. Validate masters and CSV selection.
        2. Load current sequences (with daily base reset applied).
        3. Call process_csv_files() — produces regular + optional tester rows.
        4. Save updated sequences immediately when SO rows were produced.
        5. Write 9-sheet Excel to <csv_folder>/output_mt/, including Control Check
           and row-by-row Input Audit.
        6. Update UI and show completion popup.

        If generate_tester is True (checkbox ticked):
          - 5 POs produce 5 regular/tester SO pairs (10 SOs total).
          - A tester SO uses the same number as its regular SO, with prefix
            ``SO/HG/TT/`` instead of ``SO/HG/MM/``.
          - Both appear in the output, with TESTER rows highlighted blue.
          - Failed control checks are warnings only; output is still generated.
        """
        if not self._update_master_status():
            messagebox.showerror(
                "Masters Not Ready",
                "All three master files must be loaded.\n"
                "Check data_mt/ folder or click 🔄 Reload Masters."
            )
            return
        if not self.csv_paths:
            messagebox.showwarning("No CSV", "Select at least one H&G PO CSV file.")
            return

        self.status_var.set("Processing — please wait…")
        self.status_label.config(fg="blue")
        self.root.update()
        t0 = time.time()

        seqs = load_sequences()
        log.info(
            f"[Generate] Mode={'REGULAR+TESTER' if self.tester_var.get() else 'REGULAR ONLY'} | "
            f"Shared sequence={seqs['HG']} | "
            f"Today base={_todays_base_sequence()}"
        )

        result = process_csv_files(
            file_paths      = self.csv_paths,
            hg_master       = self.hg_master,
            items_master    = self.items_master,
            address_master  = self.address_master,
            warehouse_code  = WAREHOUSES[self.warehouse_var.get()],
            generate_tester = self.tester_var.get(),
            sequences       = seqs
        )

        if result.rows:
            # Save sequences immediately after successful SO creation.
            seqs["HG"] = result.hg_sequence
            seqs["TT"] = result.hg_sequence
            save_sequences(seqs)
            self._refresh_seq_display()
        else:
            log.warning(
                "[Generate] No valid SO rows created. "
                "Writing warning/control workbook without consuming a sequence."
            )
        self.last_result = result

        # Output path: next to the first CSV file.
        out_dir = Path(self.csv_paths[0]).parent / "output_mt"
        out_dir.mkdir(parents=True, exist_ok=True)
        ts   = datetime.now().strftime("%d-%m-%Y_%H%M%S")
        mode = "HG_TT" if self.tester_var.get() else "HG"
        outf = out_dir / f"MT_Select_{mode}_{ts}.xlsx"

        write_output_workbook(result, outf)
        self.last_output = outf
        self.open_btn.config(state=tk.NORMAL)

        elapsed  = time.time() - t0
        reg_sos  = len(result.regular_so_map)
        test_sos = len(result.tester_so_map)
        summary  = (
            f"Done: {len(result.rows)} lines | "
            f"{reg_sos} regular SOs"
            + (f" + {test_sos} tester SOs = {reg_sos+test_sos} total" if test_sos else "")
            + f" | {len(result.warnings)} warnings | {elapsed:.2f}s"
        )
        self.status_var.set(summary)
        self.status_label.config(fg="darkgreen" if not result.warnings else "darkorange")
        log.info(f"[Generate] {summary}")
        log.info(f"[Generate] Output: {outf}")

        messagebox.showinfo(
            "Complete",
            f"{summary}\n\n"
            f"Output: {outf}\n"
            f"Log:    {LOG_FILE}"
        )

    # --------------------------------------------------------------------------
    # UTILITY
    # --------------------------------------------------------------------------

    def open_last(self):
        if self.last_output and self.last_output.exists():
            try:
                os.startfile(str(self.last_output))
            except Exception as e:
                messagebox.showerror("Error", f"Cannot open:\n{self.last_output}\n{e}")
        else:
            messagebox.showwarning("Not Found", "No output yet. Run Generate first.")

    def open_log_folder(self):
        try:
            os.startfile(str(LOG_DIR))
        except Exception:
            messagebox.showinfo("Log Folder", str(LOG_DIR))

    def run(self):
        """Start Tkinter main loop (blocking until window closed)."""
        self.root.mainloop()


# ==============================================================================
# SECTION 12 — ENTRY POINT
# ==============================================================================

def main():
    """
    Application entry point.
    1. Check expiry date.
    2. Launch MTSelectApp.
    """
    expiry = datetime.strptime(EXPIRY_DATE, "%d-%m-%Y").date()
    if datetime.now().date() > expiry:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Tool Expired",
            f"Expired on {EXPIRY_DATE}.\nContact Order Management Automation Team."
        )
        sys.exit(0)

    app = MTSelectApp()
    app.run()
    log.info("Application closed.")


if __name__ == "__main__":
    main()
