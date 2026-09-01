import pandas as pd
from datetime import datetime
import xlwings as xw
import os
import re

# ===================================================
# READ XLS/XLSX VIA EXCEL
# ===================================================
def read_xls_with_xlwings(file_path):
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    data = None
    try:
        wb = app.books.open(file_path)
        sheet = wb.sheets[0]
        data = sheet.used_range.value
        wb.close()
    except:
        pass
    finally:
        app.quit()

    return data


# ===================================================
# PROCESSOR
# ===================================================
class FlipkartPOProcessor:

    def __init__(self):
        self.today = datetime.today().strftime("%d-%m-%Y")

    # Extract PO number from filename
    def extract_po_number(self, file_path):
        base = os.path.basename(file_path).lower()
        name = base.replace(".xls", "").replace(".xlsx", "")
        if name.startswith("purchase_order_"):
            return name.replace("purchase_order_", "").upper()
        return name.upper()

    # Find header row by scanning for fsn/ean/qty
    def find_header_row(self, raw_data):
        for i, row in enumerate(raw_data):
            if not row:
                continue

            row_str = " ".join(str(x).lower() for x in row if x)

            if ("fsn" in row_str or
                "fsn/isbn13" in row_str or
                "ean" in row_str or
                "quantity" in row_str):

                return i
        return None

    # Extract EXACT shipped-to address from the cell AFTER the label
    def extract_shipped_to_address(self, raw_data):
        for row in raw_data:
            if not row:
                continue

            for i, cell in enumerate(row):
                if cell and "shipped to address" in str(cell).lower():
                    for j in range(i + 1, len(row)):
                        next_cell = row[j]
                        if next_cell and str(next_cell).strip():
                            return str(next_cell).strip()

        return ""

    # Clean address (minimal cleaning — only keep Flipkart India → last pincode)
    def clean_address(self, address):
        if not address:
            return ""

        addr = str(address)

        # Always keep from "flipkart india"
        start = addr.lower().find("flipkart india")
        if start != -1:
            addr = addr[start:]

        # Keep only up to last pincode (6-digit)
        pincodes = list(re.finditer(r"\b\d{6}\b", addr))
        if pincodes:
            last_end = pincodes[-1].end()
            addr = addr[:last_end]

        return addr.strip()

    # Main processing loop
    def process_files(self, files, progress_callback=None):

        final_rows = []
        total_files = len(files)

        if progress_callback: progress_callback(f"Total PO files detected: {total_files}")

        for index, file_path in enumerate(files, start=1):

            if progress_callback: progress_callback(f"Processing {index}/{total_files}: {os.path.basename(file_path)}")

            raw_data = read_xls_with_xlwings(file_path)
            if not raw_data:
                if progress_callback: progress_callback("❌ Could not read file")
                continue

            header_row = self.find_header_row(raw_data)
            if header_row is None:
                if progress_callback: progress_callback("❌ Header row not found")
                continue

            # Load data into DataFrame
            df = pd.DataFrame(raw_data[header_row:])
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)

            # Normalize columns
            df.columns = [str(c).strip().lower() for c in df.columns]

            # Extract shipped-to address
            address = self.extract_shipped_to_address(raw_data)
            address = self.clean_address(address)

            po_number = self.extract_po_number(file_path)

            # Build rows
            for r in df.to_dict('records'):

                fsn = str(r.get("fsn/isbn13", r.get("fsn", ""))).strip()
                title = str(r.get("title", "")).strip()
                qty_raw = r.get("quantity", r.get("qty", ""))

                if not fsn or not title:
                    continue

                try:
                    qty = int(float(str(qty_raw)))
                except:
                    continue

                row = {
                    "Date": self.today,
                    "PO": po_number,
                    "FSN Code": fsn,
                    "EAN": r.get("ean", ""),
                    "Item": "",
                    "Qty": qty,
                    "COST PRICE": r.get("supplier price", ""),
                    "total_amount": r.get("total amount", ""),
                    "description": title,
                    "Address": address,
                    "Ship TO": ""
                }

                final_rows.append(row)

        df_final = pd.DataFrame(final_rows)
        if progress_callback: progress_callback("✔ Successfully extracted ALL PO data")

        return df_final

def process_flipkart_dump(files, progress_callback=None):
    """Main entrypoint for the worker thread to generate flipkart dump"""
    processor = FlipkartPOProcessor()
    df = processor.process_files(files, progress_callback)

    if df.empty:
        raise ValueError("No valid SKU rows found in PO files.")

    # Auto-save with timestamp
    base_dir = os.path.dirname(files[0])
    timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    filename = f"FL_DUMP_COMPILATION_{timestamp}.xlsx"
    save_path = os.path.join(base_dir, filename)

    df.to_excel(save_path, index=False)

    # Calculate summary
    total_files = len(files)
    total_rows = len(df)
    total_qty = int(df["Qty"].sum())

    # Handle numeric conversion for total_amount
    try:
        df["total_amount"] = pd.to_numeric(df["total_amount"], errors="coerce")
    except:
        pass

    total_amount = float(df["total_amount"].sum())

    return {
        'marketplace': 'Flipkart Dump',
        'output_file': save_path,
        'summary': {
            'total_files': total_files,
            'total_rows': total_rows,
            'total_qty': total_qty,
            'total_amount': total_amount
        }
    }