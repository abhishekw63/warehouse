from __future__ import annotations

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional
from datetime import datetime
import logging
import re
import sys


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
# Data Model
# ---------------------------

@dataclass
class OrderRow:
    so_number: str
    item_no: str
    qty: int


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
# Excel Parser
# ---------------------------

class ExcelParser:

    BC_COLUMN = "bc code"
    QTY_COLUMN = "order qty"

    def parse(self, file_path: Path) -> List[OrderRow]:

        logging.info(f"Reading {file_path.name}")

        try:
            raw_df = pd.read_excel(file_path, header=None)
        except Exception as e:
            logging.error(f"Failed reading {file_path}: {e}")
            return []

        header_row = None

        for i, row in raw_df.iterrows():
            row_values = [str(v).lower() for v in row.values]

            if "bc code" in row_values and any("order qty" in v for v in row_values):
                header_row = i
                break

        if header_row is None:
            logging.warning(f"Header row not found in {file_path}")
            return []

        df = pd.read_excel(file_path, header=header_row)

        bc_col, qty_col = self._detect_columns(df)

        if bc_col is None or qty_col is None:
            logging.warning(f"Required columns not found in {file_path}")
            return []

        so_number = SOFormatter.from_filename(file_path)

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
                    qty=qty
                )
            )

        return rows

    def _detect_columns(self, df):

        bc_col = None
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

    def export(self, rows: List[OrderRow]):

        if not rows:
            messagebox.showwarning("No Data", "No valid rows found.")
            return

        df = pd.DataFrame(
            [
                {
                    "SO Number": r.so_number,
                    "Item No": r.item_no,
                    "Qty": r.qty
                }
                for r in rows
            ]
        )

        # create output folder if not exists
        output_folder = Path("output")
        output_folder.mkdir(exist_ok=True)

        # generate date and time based filename
        today = datetime.now().strftime("%d-%m-%Y_%H%M%S") # i want this in ddmmyyy_hhmmss format 17-09-2024_153045

        file_path = output_folder / f"gt_mass_dump_{today}.xlsx"

        df.to_excel(file_path, index=False)

        messagebox.showinfo(
            "Success",
            f"Dump generated:\n{file_path}"
        )


# ---------------------------
# Main Automation Engine
# ---------------------------

class GTMassAutomation:

    def __init__(self):

        self.parser = ExcelParser()
        self.exporter = DumpExporter()

    def process_files(self, files: List[Path]) -> List[OrderRow]:

        all_rows: List[OrderRow] = []

        for file in files:

            rows = self.parser.parse(file)

            all_rows.extend(rows)

        logging.info(f"{len(all_rows)} rows extracted")

        return all_rows


# ---------------------------
# Tkinter UI
# ---------------------------

class AutomationUI:

    def __init__(self, automation: GTMassAutomation):

        self.automation = automation
        self.files: List[Path] = []

        self.root = tk.Tk()
        self.root.title("GT Mass Dump Generator")
        self.root.geometry("420x260")

        title = tk.Label(
            self.root,
            text="GT Mass Dump Generator",
            font=("Arial", 14, "bold")
        )
        title.pack(pady=10)

        self.label = tk.Label(self.root, text="Selected Files: 0")
        self.label.pack(pady=5)

        self.select_button = tk.Button(
            self.root,
            text="Select Excel Files",
            width=20,
            command=self.select_files
        )
        self.select_button.pack(pady=10)

        self.generate_button = tk.Button(
            self.root,
            text="Generate Dump",
            width=20,
            command=self.generate_dump
        )
        self.generate_button.pack(pady=10)

        self.status = tk.Label(self.root, text="Status: Waiting")
        self.status.pack(pady=20)

    def select_files(self):

        files = filedialog.askopenfilenames(
            title="Select Sales Order Files",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        self.files = [Path(f) for f in files]

        self.label.config(text=f"Selected Files: {len(self.files)}")

    def generate_dump(self):

        if not self.files:
            messagebox.showwarning("Warning", "Please select files first")
            return

        self.status.config(text="Processing files...")

        rows = self.automation.process_files(self.files)

        self.automation.exporter.export(rows)

        self.status.config(text="Dump generated")

    def run(self):

        self.root.mainloop()


# ---------------------------
# Entry Point
# ---------------------------

def main():

    # Run expiry check before launching UI
    check_expiry()

    automation = GTMassAutomation()

    ui = AutomationUI(automation)

    ui.run()


if __name__ == "__main__":
    main()