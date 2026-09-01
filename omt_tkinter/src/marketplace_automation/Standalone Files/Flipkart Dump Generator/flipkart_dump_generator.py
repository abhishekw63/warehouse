import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
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
    def process_files(self, files):

        final_rows = []
        total_files = len(files)

        print(f"\n====================================")
        print(f"Total PO files detected: {total_files}")
        print("====================================\n")

        for index, file_path in enumerate(files, start=1):

            print(f"Processing {index}/{total_files}: {os.path.basename(file_path)}")

            raw_data = read_xls_with_xlwings(file_path)
            if not raw_data:
                print("❌ Could not read file\n")
                continue

            header_row = self.find_header_row(raw_data)
            if header_row is None:
                print("❌ Header row not found\n")
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

            print("✔ Extracted\n")

        df_final = pd.DataFrame(final_rows)

        print("====================================")
        print("✔ Successfully extracted ALL PO data")
        print("====================================\n")

        return df_final


# ===================================================
# TKINTER UI
# ===================================================
class FlipkartPOApp:

    def __init__(self, root):
        self.root = root
        self.processor = FlipkartPOProcessor()

        root.title("Flipkart PO Dump Automation")
        root.geometry("600x260")
        root.resizable(False, False)

        tk.Label(
            root,
            text="Flipkart PO Dump Automation",
            font=("Arial", 14, "bold")
        ).pack(pady=20)

        tk.Button(
            root,
            text="Select PO Files",
            font=("Arial", 12),
            width=20,
            command=self.select_files
        ).pack(pady=20)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Flipkart PO files",
            filetypes=[("Excel Files", "*.xls *.xlsx")]
        )

        if not files:
            return

        df = self.processor.process_files(files)

        if df.empty:
            messagebox.showerror("Error", "No valid SKU rows found in PO.")
            return

        # Auto-save with timestamp
        base_dir = os.path.dirname(files[0])
        timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        filename = f"FL_DUMP_COMPILATION_{timestamp}.xlsx"
        save_path = os.path.join(base_dir, filename)

        df.to_excel(save_path, index=False)

        # SUMMARY POPUP
        # ============================
        total_files = len(files)
        total_rows = len(df)
        total_qty = df["Qty"].sum()

        # Handle numeric conversion for total_amount
        try:
            df["total_amount"] = pd.to_numeric(df["total_amount"], errors="coerce")
        except:
            pass

        total_amount = df["total_amount"].sum()

        summary_message = (
            f"Total PO files processed: {total_files}\n"
            f"Total SKUs extracted: {total_rows}\n"
            f"Total Quantity: {total_qty}\n"
            f"Total Amount: ₹{total_amount:,.2f}\n\n"
            f"File saved at:\n{save_path}\n\n"
            f"Do you want to open the file?"
        )

        # Ask user if they want to open file
        open_now = messagebox.askyesno("Summary", summary_message)

        if open_now:
            os.startfile(save_path)



# ===================================================
# MAIN
# ===================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = FlipkartPOApp(root)
    root.mainloop()
