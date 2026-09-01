import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os

# ===================== LICENSE CONTROL =====================
SOFTWARE_EXPIRY_DATE = "2026-06-28"  # YYYY-MM-DD
EXPIRY_WARNING_DAYS = 10

def check_license():
    expiry = datetime.strptime(SOFTWARE_EXPIRY_DATE, "%Y-%m-%d").date()
    today = datetime.today().date()
    if today > expiry:
        messagebox.showerror(
            "License Expired",
            "Software license has expired.\nPlease contact admin."
        )
        return False
    return True

def check_expiry_warning():
    expiry = datetime.strptime(SOFTWARE_EXPIRY_DATE, "%Y-%m-%d").date()
    today = datetime.today().date()
    days_left = (expiry - today).days
    if 0 < days_left <= EXPIRY_WARNING_DAYS:
        messagebox.showwarning(
            "License Expiring Soon",
            f"Your software license will expire in {days_left} day(s).\nPlease contact admin to renew."
        )

# ===================== MAIN APP =====================
class InventoryAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventory Bin Filter Automation")
        self.root.geometry("700x500")
        self.root.resizable(False, False)

        self.file_path = None

        self.build_ui()

        # Show expiry warning on startup
        self.root.after(500, check_expiry_warning)

    def build_ui(self):
        tk.Label(self.root, text="Upload Inventory File", font=("Arial", 12, "bold")).pack(pady=10)

        tk.Button(
            self.root,
            text="Upload Excel File",
            width=25,
            command=self.upload_file
        ).pack()

        self.file_label = tk.Label(self.root, text="No file selected", fg="gray")
        self.file_label.pack(pady=5)

        tk.Label(self.root, text="Enter Bin Codes (one per line)", font=("Arial", 12, "bold")).pack(pady=10)

        self.bin_text = tk.Text(self.root, height=10, width=60)
        self.bin_text.pack()

        tk.Button(
            self.root,
            text="Process",
            width=20,
            bg="#2c7be5",
            fg="white",
            command=self.process_file
        ).pack(pady=20)

    def upload_file(self):
        filetypes = [("Excel Files", "*.xlsx *.xls")]
        self.file_path = filedialog.askopenfilename(filetypes=filetypes)
        if self.file_path:
            self.file_label.config(text=os.path.basename(self.file_path), fg="green")

    def process_file(self):
        if not check_license():
            return

        if not self.file_path:
            messagebox.showwarning("Missing File", "Please upload inventory file.")
            return

        bin_codes = self.bin_text.get("1.0", tk.END).strip().splitlines()
        bin_codes = [b.strip() for b in bin_codes if b.strip()]

        if not bin_codes:
            messagebox.showwarning("Missing Bin Codes", "Please enter at least one Bin Code.")
            return

        try:
            df = pd.read_excel(self.file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Unable to read file:\n{e}")
            return

        required_columns = ["Bin Code", "Item No.", "GTIN", "ItemDescription", "Quantity"]
        missing_cols = [c for c in required_columns if c not in df.columns]

        if missing_cols:
            messagebox.showerror(
                "Missing Columns",
                f"These required columns are missing:\n{', '.join(missing_cols)}"
            )
            return

        filtered_df = df[df["Bin Code"].isin(bin_codes)]

        if filtered_df.empty:
            messagebox.showwarning(
                "No Data",
                "No matching data found for entered Bin Codes."
            )
            return

        # ===================== SHEET 1 =====================
        sku_summary = (
            filtered_df
            .groupby(["Item No.", "GTIN", "ItemDescription"], as_index=False)["Quantity"]
            .sum()
            .rename(columns={
                "Item No.": "SKU",
                "Quantity": "Total Qty"
            })
        )

        # ===================== SHEET 2 =====================
        bin_content = (
            filtered_df
            .groupby(
                ["Bin Code", "Item No.", "GTIN", "ItemDescription"],
                as_index=False
            )["Quantity"]
            .sum()
            .rename(columns={
                "Item No.": "SKU",
                "Quantity": "Qty"
            })
        )

        # ===================== AUTO SAVE =====================
        input_dir = os.path.dirname(self.file_path)
        timestamp = datetime.now().strftime("%d%m%Y_%H%M")
        output_file = f"Inventory_Output_{timestamp}.xlsx"
        output_path = os.path.join(input_dir, output_file)

        try:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                sku_summary.to_excel(writer, sheet_name="SKU_Summary", index=False)
                bin_content.to_excel(writer, sheet_name="Bin_Content", index=False)

            messagebox.showinfo(
                "Success",
                f"File saved successfully:\n{output_file}"
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")


# ===================== RUN =====================
if __name__ == "__main__":
    root = tk.Tk()
    if not check_license():
        root.destroy()
    else:
        app = InventoryAutomationApp(root)
        root.mainloop()