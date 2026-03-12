from __future__ import annotations

import pandas as pd
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Any
import logging
import re
from datetime import datetime
import io

# ---------------------------
# Logging Configuration
# ---------------------------

logger = logging.getLogger(__name__)

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
    def from_filename(filename: str) -> Optional[str]:
        # Strip extension if present to easily grab the numbers
        name_without_ext = Path(filename).stem
        match = re.search(r"\d+", name_without_ext)

        if not match:
            logger.warning(f"SO number not found in {filename}")
            return None

        return f"SO/GTM/{match.group()}"


# ---------------------------
# Excel Parser
# ---------------------------

class ExcelParser:

    BC_COLUMN = "bc code"
    QTY_COLUMN = "order qty"

    def parse(self, file_obj: Any, filename: str) -> List[OrderRow]:
        logger.info(f"Reading {filename}")

        try:
            raw_df = pd.read_excel(file_obj, header=None)
        except Exception as e:
            logger.error(f"Failed reading {filename}: {e}")
            return []

        header_row = None

        for i, row in raw_df.iterrows():
            row_values = [str(v).lower() for v in row.values]

            if "bc code" in row_values and any("order qty" in v for v in row_values):
                header_row = i
                break

        if header_row is None:
            logger.warning(f"Header row not found in {filename}")
            return []

        # Reset file pointer to read again with correct header
        file_obj.seek(0)
        df = pd.read_excel(file_obj, header=header_row)

        bc_col, qty_col = self._detect_columns(df)

        if bc_col is None or qty_col is None:
            logger.warning(f"Required columns not found in {filename}")
            return []

        so_number = SOFormatter.from_filename(filename)

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

    def export_to_memory(self, rows: List[OrderRow]) -> io.BytesIO:
        if not rows:
            return None

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

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        output.seek(0)
        return output

# ---------------------------
# Main Automation Engine
# ---------------------------

class GTMassAutomation:

    def __init__(self):
        self.parser = ExcelParser()
        self.exporter = DumpExporter()

    def process_files(self, file_objects: List[Any]) -> List[OrderRow]:
        all_rows: List[OrderRow] = []

        for file_obj in file_objects:
            # InMemoryUploadedFile has a .name attribute
            rows = self.parser.parse(file_obj, file_obj.name)
            all_rows.extend(rows)

        logger.info(f"{len(all_rows)} rows extracted")
        return all_rows
