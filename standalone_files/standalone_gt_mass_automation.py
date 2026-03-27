from __future__ import annotations

import os
import sys
import platform
import time
import logging
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple
from datetime import datetime


# ---------------------------
# Logging Configuration
# ---------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)


# ---------------------------
# Expiry Check
# ---------------------------

EXPIRY_DATE = "31-03-2026"

def check_expiry():
    expiry = datetime.strptime(EXPIRY_DATE, "%d-%m-%Y").date()
    today = datetime.now().date()

    if today > expiry:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Application Expired",
            f"This application expired on {EXPIRY_DATE}.\n\n"
            f"Please contact the administrator for an updated version."
        )
        root.destroy()
        sys.exit(0)

    days_remaining = (expiry - today).days
    if days_remaining <= 7:
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning(
            "Expiration Warning",
            f"⚠️ This application will expire in {days_remaining} day(s).\n\n"
            f"Expiry Date: {EXPIRY_DATE}\n\n"
            f"Please contact the administrator for renewal."
        )
        root.destroy()


# ---------------------------
# Known State / Zone Values
# These should never appear as a Distributor name.
# If they do, it means rows were swapped in the source file.
# ---------------------------

STATE_LIKE_VALUES = {
    "up", "mp", "ap", "hp", "uk", "jk", "wb", "tn", "kl", "ka",
    "gj", "rj", "hr", "pb", "br", "od", "as", "mh", "cg", "jh",
    "north", "south", "east", "west", "central",
    "uttar pradesh", "madhya pradesh", "rajasthan", "punjab",
    "maharashtra", "gujarat", "karnataka", "tamil nadu",
    "haryana", "delhi", "u.p", "u.p.", "m.p", "m.p."
}


# ---------------------------
# Data Model
# ---------------------------

@dataclass
class OrderRow:
    so_number: str
    item_no: str
    qty: int
    distributor: str
    city: str
    state: str


@dataclass
class ProcessingResult:
    """Holds all rows + any issues found during processing."""
    rows: List[OrderRow] = field(default_factory=list)
    failed_files: List[Tuple[str, str]] = field(default_factory=list)   # (filename, reason)
    warned_files: List[Tuple[str, str]] = field(default_factory=list)   # (filename, warning)
    output_path: Optional[Path] = None


# ---------------------------
# SO Formatter
# ---------------------------

class SOFormatter:

    @staticmethod
    def from_filename(filepath: Path) -> Optional[str]:
        match = re.search(r"\d+", filepath.stem)
        if not match:
            logging.warning(f"SO number not found in {filepath}")
            return None
        return f"SO/GTM/{match.group()}"


# ---------------------------
# File Reader
#
# Reading strategy by extension:
#   .xlsx / .xlsm  →  openpyxl  (built-in, no extra install)
#   .xls           →  xlrd      (pip install xlrd)
#
# If the correct library is missing, a clear install instruction
# is shown rather than a cryptic error.
# ---------------------------

class FileReader:

    @staticmethod
    def read(file_path: Path) -> pd.DataFrame:
        """
        Returns a raw DataFrame (no header) for the first sheet.
        Raises RuntimeError with a clear message on failure.
        """
        ext = file_path.suffix.lower()

        # --- .xlsx / .xlsm → openpyxl ---
        if ext in (".xlsx", ".xlsm"):
            try:
                df = pd.read_excel(file_path, header=None, engine="openpyxl")
                logging.info(f"{file_path.name} — read via openpyxl")
                return df
            except Exception as e:
                raise RuntimeError(
                    f"Cannot read '{file_path.name}'.\n"
                    f"The file may be corrupt or password-protected.\n"
                    f"Error: {e}"
                )

        # --- .xls → xlrd ---
        if ext == ".xls":
            try:
                df = pd.read_excel(file_path, header=None, engine="xlrd")
                logging.info(f"{file_path.name} — read via xlrd")
                return df
            except ImportError:
                raise RuntimeError(
                    f"Cannot read '{file_path.name}' — xlrd is not installed.\n\n"
                    f"Fix: open your terminal / command prompt and run:\n"
                    f"    pip install xlrd\n\n"
                    f"Then restart this application and try again."
                )
            except Exception as e:
                raise RuntimeError(
                    f"Cannot read '{file_path.name}'.\n"
                    f"The file may be corrupt or password-protected.\n"
                    f"Error: {e}"
                )

        # --- Unsupported extension ---
        raise RuntimeError(
            f"Unsupported file format: '{ext}'.\n"
            f"Only .xlsx, .xlsm, and .xls files are supported."
        )


# ---------------------------
# Meta Extractor
# ---------------------------

class MetaExtractor:
    """
    Scans raw header rows (above the data table) to extract
    Distributor Name, City, and State by label matching.
    Row positions vary across files so we scan by label — never hardcode.
    Returns extracted values + any warnings found.
    """

    @staticmethod
    def extract(raw_df: pd.DataFrame, header_row: int) -> Tuple[dict, List[str]]:

        distributor = ""
        city = ""
        state_values = []
        warnings = []

        meta_df = raw_df.iloc[:header_row]

        for _, row in meta_df.iterrows():

            label = str(row.iloc[0]).strip().lower()
            value = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""

            if value.lower() in ("nan", ""):
                value = ""

            if label == "distributor name" and not distributor:
                distributor = value
                logging.info(f"Distributor found: '{distributor}'")

            elif label == "city" and not city:
                city = value
                logging.info(f"City found: '{city}'")

            elif label == "state":
                state_values.append(value)

        # Pick last non-blank state (bottom State row is always the proper state)
        state = ""
        for s in reversed(state_values):
            if s:
                state = s
                break

        logging.info(f"State found: '{state}'")

        # --- Validation ---

        if not distributor:
            warnings.append(
                "Distributor Name is blank — 'Distributor Name' label not found or value is empty."
            )

        if not city:
            warnings.append(
                "City is blank — 'City' label not found or value is empty."
            )

        if not state:
            warnings.append(
                "State is blank — both State rows are empty or missing."
            )

        if distributor and distributor.strip().lower() in STATE_LIKE_VALUES:
            warnings.append(
                f"Distributor value '{distributor}' looks like a state or zone name. "
                f"Rows may be swapped in the source file — please verify manually."
            )

        return {
            "distributor": distributor,
            "city": city,
            "state": state
        }, warnings


# ---------------------------
# Excel Parser
# ---------------------------

class ExcelParser:

    BC_COLUMN  = "bc code"
    QTY_COLUMN = "order qty"

    def parse(self, file_path: Path) -> Tuple[List[OrderRow], List[str]]:
        """
        Returns (rows, warnings).
        Raises RuntimeError if the file cannot be read or structure is broken.
        """

        logging.info(f"Reading {file_path.name}")
        warnings = []

        # --- Read raw file ---
        raw_df = FileReader.read(file_path)   # raises RuntimeError if unreadable

        # --- Find header row ---
        header_row = None
        for i, row in raw_df.iterrows():
            row_values = [str(v).lower() for v in row.values]
            if "bc code" in row_values and any("order qty" in v for v in row_values):
                header_row = i
                break

        if header_row is None:
            raise RuntimeError(
                "Header row not found — could not locate both 'BC Code' and 'Order Qty' "
                "columns. File format may have changed."
            )

        # --- Extract meta fields ---
        meta, meta_warnings = MetaExtractor.extract(raw_df, header_row)
        warnings.extend(meta_warnings)

        # --- Build data table from raw_df (avoids re-reading the file) ---
        df = raw_df.iloc[header_row + 1:].copy()
        df.columns = raw_df.iloc[header_row].values
        df = df.reset_index(drop=True)

        bc_col, qty_col = self._detect_columns(df)

        if bc_col is None:
            raise RuntimeError(
                "Column 'BC Code' not found in data table. "
                "Column name may have changed in this file."
            )

        if qty_col is None:
            raise RuntimeError(
                "Column 'Order Qty' not found in data table. "
                "Column name may have changed in this file."
            )

        # --- SO number ---
        so_number = SOFormatter.from_filename(file_path)
        if so_number is None:
            warnings.append(
                "Could not extract SO number from filename. "
                "Filename should contain digits e.g. SOGTM6290.xlsx"
            )
            so_number = "SO/GTM/UNKNOWN"

        # --- Extract ordered rows ---
        rows: List[OrderRow] = []

        for _, row in df.iterrows():

            bc_code = row[bc_col]

            if pd.isna(bc_code):
                continue

            try:
                bc_code = int(bc_code)
            except:
                continue

            qty = self._clean_qty(row[qty_col])

            if qty <= 0:
                continue

            rows.append(
                OrderRow(
                    so_number=so_number,
                    item_no=str(bc_code),
                    qty=qty,
                    distributor=meta["distributor"],
                    city=meta["city"],
                    state=meta["state"]
                )
            )

        if not rows:
            warnings.append(
                "No ordered rows found — all Order Qty values are 0 or blank. "
                "This file will not contribute any lines to the dump."
            )

        return rows, warnings

    def _detect_columns(self, df):
        bc_col  = None
        qty_col = None
        for col in df.columns:
            name = str(col).strip().lower()
            if name == self.BC_COLUMN:
                bc_col = col
            if self.QTY_COLUMN in name:
                qty_col = col
        return bc_col, qty_col

    @staticmethod
    def _clean_qty(value) -> int:
        if pd.isna(value):
            return 0
        value = str(value).strip()
        if value in ("", "-"):
            return 0
        value = value.replace(",", "")
        try:
            return int(float(value))
        except ValueError:
            return 0


# ---------------------------
# Dump Exporter
# ---------------------------

class DumpExporter:

    def export(self, result: ProcessingResult) -> Optional[Path]:
        """
        Writes the Excel output.
        Returns the output file path on success, or None if nothing to export.
        """

        # --- Show popup for failed files BEFORE export ---
        if result.failed_files:
            msg = "The following files could NOT be read and were skipped:\n\n"
            for fname, reason in result.failed_files:
                msg += f"  • {fname}\n    Reason: {reason}\n\n"
            msg += "Please fix these files and re-process them."
            messagebox.showerror("Files Failed to Read", msg)

        # --- If no rows at all, stop ---
        if not result.rows:
            messagebox.showwarning(
                "No Data",
                "No valid rows found across all selected files.\n"
                "Nothing to export."
            )
            return None

        # --- Sales Lines (SO Number | Item No | Qty) ---
        lines_df = pd.DataFrame(
            [
                {
                    "SO Number": r.so_number,
                    "Item No":   r.item_no,
                    "Qty":       r.qty
                }
                for r in result.rows
            ]
        )

        # --- Sales Header (SO Number | Qty | Distributor | City | State) ---
        full_df = pd.DataFrame(
            [
                {
                    "SO Number":   r.so_number,
                    "Qty":         r.qty,
                    "Distributor": r.distributor,
                    "City":        r.city,
                    "State":       r.state
                }
                for r in result.rows
            ]
        )

        header_df = (
            full_df.groupby("SO Number", sort=False)
            .agg(
                Qty         = ("Qty",         "sum"),
                Distributor = ("Distributor", "first"),
                City        = ("City",        "first"),
                State       = ("State",       "first")
            )
            .reset_index()
        )

        # --- Warnings sheet (only if warnings exist) ---
        warn_rows = [
            {"File": fname, "Warning": warning}
            for fname, warning in result.warned_files
        ]

        # --- Write output ---
        output_folder = Path("output")
        output_folder.mkdir(exist_ok=True)

        today     = datetime.now().strftime("%d-%m-%Y_%H%M%S")
        file_path = output_folder / f"gt_mass_dump_{today}.xlsx"

        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            lines_df.to_excel(writer,  sheet_name="Sales Lines",  index=False)
            header_df.to_excel(writer, sheet_name="Sales Header", index=False)
            if warn_rows:
                pd.DataFrame(warn_rows).to_excel(writer, sheet_name="Warnings", index=False)

        return file_path


# ---------------------------
# File Opener (cross-platform)
# ---------------------------

def open_file(file_path: Path):
    """Opens the file using the OS default application."""
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(str(file_path))
        elif system == "Darwin":
            import subprocess
            import subprocess as sp
            sp.Popen(["open", str(file_path)])
        else:
            import subprocess as sp
            sp.Popen(["xdg-open", str(file_path)])
    except Exception as e:
        messagebox.showerror("Open File Error", f"Could not open file:\n{e}")


# ---------------------------
# Main Automation Engine
# ---------------------------

class GTMassAutomation:

    def __init__(self):
        self.parser   = ExcelParser()
        self.exporter = DumpExporter()

    def process_files(self, files: List[Path]) -> ProcessingResult:

        result = ProcessingResult()

        for file in files:
            fname = file.name
            try:
                rows, warnings = self.parser.parse(file)
                result.rows.extend(rows)

                for w in warnings:
                    result.warned_files.append((fname, w))
                    logging.warning(f"{fname}: {w}")

            except RuntimeError as e:
                result.failed_files.append((fname, str(e)))
                logging.error(f"{fname} FAILED: {e}")

            except Exception as e:
                result.failed_files.append((fname, f"Unexpected error: {e}"))
                logging.error(f"{fname} UNEXPECTED ERROR: {e}")

        logging.info(
            f"Processing complete — "
            f"{len(result.rows)} rows | "
            f"{len(result.failed_files)} failed | "
            f"{len(result.warned_files)} warnings"
        )

        return result


# ---------------------------
# Tkinter UI
# ---------------------------

class AutomationUI:

    def __init__(self, automation: GTMassAutomation):

        self.automation       = automation
        self.files: List[Path] = []
        self.last_output_path: Optional[Path] = None

        self.root = tk.Tk()
        self.root.title("GT Mass Dump Generator")
        self.root.geometry("440x340")
        self.root.resizable(False, False)

        # --- Title ---
        tk.Label(
            self.root,
            text="GT Mass Dump Generator",
            font=("Arial", 14, "bold")
        ).pack(pady=10)

        # --- File count label ---
        self.label = tk.Label(
            self.root,
            text="Selected Files: 0",
            font=("Arial", 10)
        )
        self.label.pack(pady=4)

        # --- Select Files button ---
        tk.Button(
            self.root,
            text="Select Excel Files",
            width=22,
            command=self.select_files
        ).pack(pady=6)

        # --- Generate Dump button ---
        tk.Button(
            self.root,
            text="Generate Dump",
            width=22,
            command=self.generate_dump
        ).pack(pady=6)

        # --- Open Last Output File button (disabled until first dump) ---
        self.open_button = tk.Button(
            self.root,
            text="Open Last Output File",
            width=22,
            state=tk.DISABLED,
            command=self.open_last_file
        )
        self.open_button.pack(pady=6)

        # --- Status label ---
        self.status = tk.Label(
            self.root,
            text="Status: Waiting",
            font=("Arial", 10),
            fg="gray"
        )
        self.status.pack(pady=6)

        # --- Time taken label ---
        self.time_label = tk.Label(
            self.root,
            text="",
            font=("Arial", 9),
            fg="darkgreen"
        )
        self.time_label.pack(pady=2)

    def select_files(self):

        files = filedialog.askopenfilenames(
            title="Select Sales Order Files",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )

        self.files = [Path(f) for f in files]
        self.label.config(text=f"Selected Files: {len(self.files)}")
        self.time_label.config(text="")
        self.status.config(text="Status: Files selected", fg="gray")

    def generate_dump(self):

        if not self.files:
            messagebox.showwarning("Warning", "Please select files first.")
            return

        # Start timer
        start_time = time.time()

        self.status.config(text="Status: Processing files...", fg="blue")
        self.time_label.config(text="")
        self.root.update()

        # Process
        result      = self.automation.process_files(self.files)
        output_path = self.automation.exporter.export(result)

        # Stop timer
        elapsed     = time.time() - start_time
        elapsed_str = f"{elapsed:.2f} seconds"

        failed = len(result.failed_files)
        warned = len(result.warned_files)
        rows   = len(result.rows)
        sos    = len(set(r.so_number for r in result.rows)) if result.rows else 0

        if output_path:
            self.last_output_path = output_path
            self.open_button.config(state=tk.NORMAL)

            if failed > 0 or warned > 0:
                self.status.config(
                    text=f"Done — {rows} rows | {failed} failed | {warned} warning(s)",
                    fg="orange"
                )
            else:
                self.status.config(
                    text=f"Done — {rows} rows across {sos} SO(s)",
                    fg="darkgreen"
                )

            self.time_label.config(text=f"⏱  Time taken: {elapsed_str}")

            warn_note = (
                f"\n⚠️  {warned} warning(s) found — check 'Warnings' sheet."
                if warned else ""
            )
            fail_note = (
                f"\n❌  {failed} file(s) failed to read — see error popup."
                if failed else ""
            )

            answer = messagebox.askyesno(
                "Dump Generated",
                f"Dump generated successfully!\n\n"
                f"File   : {output_path.name}\n"
                f"Rows   : {rows}\n"
                f"SO(s)  : {sos}\n"
                f"Time   : {elapsed_str}"
                f"{warn_note}{fail_note}\n\n"
                f"Do you want to open the output file?"
            )

            if answer:
                open_file(output_path)

        else:
            self.status.config(text="Status: No data to export", fg="red")
            self.time_label.config(text=f"⏱  Time taken: {elapsed_str}")

    def open_last_file(self):
        if self.last_output_path and self.last_output_path.exists():
            open_file(self.last_output_path)
        else:
            messagebox.showwarning(
                "File Not Found",
                "The output file no longer exists.\n"
                "Please generate a new dump."
            )

    def run(self):
        self.root.mainloop()


# ---------------------------
# Entry Point
# ---------------------------

def main():
    check_expiry()
    automation = GTMassAutomation()
    ui = AutomationUI(automation)
    ui.run()


if __name__ == "__main__":
    main()