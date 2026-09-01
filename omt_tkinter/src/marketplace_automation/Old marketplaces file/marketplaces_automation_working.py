import pandas as pd
from pathlib import Path
from datetime import datetime
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class POReportApp:
    """
    A GUI application to generate Purchase Order (PO) reports for different marketplaces.
    Currently, Blinkit integration is functional.
    Flipkart and Swiggy placeholders added for future development.
    """

    def __init__(self, root):
        """Initialize the main window and configure the UI."""
        self.root = root
        self.root.title("PO Report Generator")
        self.root.geometry("460x320")
        self.root.resizable(False, False)

        self.marketplace_var = tk.StringVar(value="Blinkit")

        self._build_ui()

    def _build_ui(self):
        """Construct the Tkinter widgets for the application."""
        ttk.Label(self.root, text="PO Report Generator", font=("Segoe UI", 16, "bold")).pack(pady=(20, 10))

        ttk.Label(self.root, text="Select Marketplace:").pack(pady=(10, 5))
        self.marketplace_dropdown = ttk.Combobox(
            self.root,
            textvariable=self.marketplace_var,
            values=["Blinkit", "Flipkart", "Swiggy", "Zepto"],
            state="readonly",
            width=25
        )
        self.marketplace_dropdown.pack()

        ttk.Button(
            self.root,
            text="Select CSV/Xlsx and Generate Report",
            command=self.generate_report
        ).pack(pady=(30, 10))

        ttk.Label(
            self.root,
            text="Other marketplaces are coming soon!",
            font=("Segoe UI", 9, "italic"),
            foreground="green"
        ).pack()
        ttk.Label(
            self.root,
            text=f"Last Updated: {datetime.now().strftime('%d-%m-%Y')} | Owner: RENEE-723",
            font=("Segoe UI", 8),
            foreground="gray"
        ).pack(pady=(5, 0))

    @staticmethod
    def format_indian(number):
        s = str(int(number))
        if len(s) <= 3:
            return s
        last3 = s[-3:]
        remaining = s[:-3]
        parts = []
        while len(remaining) > 2:
            parts.insert(0, remaining[-2:])
            remaining = remaining[:-2]
        if remaining:
            parts.insert(0, remaining)
        return ','.join(parts) + ',' + last3

    def show_summary_popup(self, df, marketplace):
        """
        Show summary popup for the given dataframe.
        Works for Blinkit and Flipkart.
        """
        if marketplace == "Blinkit":
            total_pos = df['po_number'].nunique()
            total_units = df['units_ordered'].sum()
            total_value = df['total_amount'].sum()
            min_date = df['order_date'].min()
            max_date = df['expiry_date'].max()
        elif marketplace == "Flipkart":
            total_pos = df['PO'].nunique()
            total_units = df['PO Qty'].sum()
            # Remove ₹ sign if already present, convert to float
            df['PO Value'] = df['PO Value'].replace('[₹,]', '', regex=True).astype(float)
            total_value = df['PO Value'].sum()
            min_date = df['Order Date'].min()
            max_date = df['Expiry Date'].max()
        elif marketplace == "Swiggy":
            total_pos = df['PO'].nunique()
            total_units = df['PO Qty'].sum()
            # Remove ₹ sign if already present, convert to float
            df['PO Value'] = df['PO Value'].replace('[₹,]', '', regex=True).astype(float)
            total_value = df['PO Value'].sum()
            min_date = df['Order Date'].min()
            max_date = df['Expiry Date'].max()
        else:
            total_pos = df.shape[0]
            total_units = 0
            total_value = 0
            min_date = max_date = None

        summary_text = (
            f"📊 SUMMARY REPORT 📊\n\n"
            f"Total POs: {total_pos}\n"
            f"Order Date Range: {min_date} to {max_date}\n"
            f"Total Units Ordered: {self.format_indian(total_units)}\n"
            f"Total Order Value: ₹ {self.format_indian(total_value)}"
        )
        messagebox.showinfo("PO Summary", summary_text)


    def generate_report(self):
        """Main function to generate PO reports for different marketplaces."""
        marketplace = self.marketplace_var.get()

        # Ask user to select CSV file
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"),("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            # --- Blinkit Logic --- 
            if marketplace == "Blinkit":
                df = pd.read_csv(file_path)
                df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

                df['po_number'] = df['po_number'].astype(str).str.replace(r'\.0$', '', regex=True)
                df['order_date'] = pd.to_datetime(df['order_date'], errors='coerce').dt.date
                df['expiry_date'] = pd.to_datetime(df['expiry_date'], errors='coerce').dt.date

                tracker_summary = df.groupby(['po_number', 'facility_name'], as_index=False).agg({
                    'order_date': 'first',
                    'expiry_date': 'first',
                    'total_amount': 'sum',
                    'units_ordered': 'sum'
                })
                tracker_summary.insert(0, 'marketplace', marketplace)
                tracker_summary['total_amount'] = tracker_summary['total_amount'].apply(
                    lambda x: f"₹ {self.format_indian(x)}"
                )
                tracker_summary = tracker_summary.sort_values('facility_name', ascending=True)
                df['upc'] = df['upc'].apply(lambda x: str(int(float(x))))
                sku_summary = df.groupby(['upc', 'name'], as_index=False).agg({'units_ordered': 'sum'})
                sku_summary.rename(columns={'units_ordered': 'total_units'}, inplace=True)
                sku_summary = sku_summary.sort_values(by='total_units', ascending=False)

                timestamp = datetime.now().strftime("%d_%m_%Y__%H_%M_%S")
                output_file = Path(file_path).parent / f"{marketplace}_PO_Report_{timestamp}.xlsx"

                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    tracker_summary.to_excel(writer, sheet_name='PO Tracker', index=False) #what is used of index
                    sku_summary.to_excel(writer, sheet_name='SKU Summary', index=False)

                    from openpyxl.styles import Alignment
                    from openpyxl.utils import get_column_letter

                    workbook = writer.book
                    for sheet_name in ['PO Tracker', 'SKU Summary']:
                        ws = workbook[sheet_name]
                        for row in ws.iter_rows(): #what is use of iter_rows
                            for cell in row:
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                        for col in ws.columns:
                            max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                            col_letter = get_column_letter(col[0].column) 
                            ws.column_dimensions[col_letter].width = max_length + 2

                self.show_summary_popup(df, marketplace)
                messagebox.showinfo("Success", f"Report created successfully at:\n{output_file}")
                if messagebox.askyesno("Open File", "Do you want to open the generated Excel file?"):
                    self.root.destroy()
                    os.startfile(output_file)

            # --- Flipkart Logic ---
            elif marketplace == "Flipkart":
                try:
                    # Use already selected file_path, no need to ask again
                    # Detect CSV or Excel
                    if str(file_path).endswith('.csv'):
                        df = pd.read_csv(file_path)
                    else:
                        df = pd.read_excel(file_path)
                    # Keep required columns and rename
                    cols_map = {
                        "Purchase Order ID": "PO",
                        "Origin Warehouse": "Location",
                        "Order Date": "Order Date",
                        "Expiry Date": "Expiry Date",
                        "Total Amount": "PO Value",
                        "Total Ordered Quantity": "PO Qty"
                    }
                    df = df[list(cols_map.keys())].rename(columns=cols_map) #what is this use of here

                    # Convert dates
                    df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce").dt.date
                    df["Expiry Date"] = pd.to_datetime(df["Expiry Date"], errors="coerce").dt.date

                    # Flipkart Alpha locations
                    FLIPKART_ALPHA_LOCS = [
                        "ban_ven_wh_nl_01nl",
                        "frk_bts",
                        "gur_san_wh_nl_01nl",
                        "malur_bts",
                        "nad_har_wh_kl_nl_01nl"
                    ]

                    # Classify Marketplace
                    df["Marketplace"] = df["Location"].apply(lambda x: "Flipkart Alpha" if x in FLIPKART_ALPHA_LOCS else "Flipkart Hyperlocal")

                    # Prepare tracker_summary format
                    tracker_summary = df[["Marketplace", "PO", "Location", "Order Date", "Expiry Date", "PO Value", "PO Qty"]].copy()
                    tracker_summary = tracker_summary.sort_values("Location", ascending=True)

                    # Since Flipkart doesn't have SKUs in this file, create empty SKU summary
                    # sku_summary = pd.DataFrame(columns=["UPC", "Name", "Total Units"])

                    # Excel writing same as Blinkit
                    timestamp = datetime.now().strftime("%d_%m_%Y__%H_%M_%S")
                    output_file = Path(file_path).parent / f"{marketplace}_PO_Report_{timestamp}.xlsx"

                    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                        tracker_summary.to_excel(writer, sheet_name='PO Tracker', index=False)
                        # sku_summary.to_excel(writer, sheet_name='SKU Summary', index=False)

                        from openpyxl.styles import Alignment
                        from openpyxl.utils import get_column_letter

                        workbook = writer.book
                        for sheet_name in ['PO Tracker']:
                            ws = workbook[sheet_name]
                            for row in ws.iter_rows():
                                for cell in row:
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                            for col in ws.columns:
                                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                                col_letter = get_column_letter(col[0].column)
                                ws.column_dimensions[col_letter].width = max_length + 2

                    self.show_summary_popup(tracker_summary, marketplace)

                    messagebox.showinfo("Success", f"Report created successfully at:\n{output_file}")
                    if messagebox.askyesno("Open File", "Do you want to open the generated Excel file?"):
                        self.root.destroy()
                        os.startfile(output_file)

                except Exception as e:
                    messagebox.showerror("Error", f"Something went wrong with Flipkart processing:\n{e}")

                # --- Swiggy Logic Placeholder ---
            elif marketplace == "Swiggy":
                try:
                    if str(file_path).endswith('.csv'):
                        df = pd.read_csv(file_path)
                    else:
                        df = pd.read_excel(file_path)
                    
                    df.columns = df.columns.str.strip().str.replace(" ", "").str.upper()

                    # Filter only CONFIRMED POs
                    df = df[df['STATUS'] == 'CONFIRMED']

                    # Convert dates
                    df['POCREATEDAT'] = pd.to_datetime(df['POCREATEDAT'], errors='coerce').dt.date
                    df['POEXPIRYDATE'] = pd.to_datetime(df['POEXPIRYDATE'], errors='coerce').dt.date

                    # Aggregate PO Tracker
                    tracker_summary = df.groupby(['PONUMBER', 'CITY'], as_index=False).agg({
                        'POCREATEDAT': 'first',
                        'POEXPIRYDATE': 'first',
                        'POLINEVALUEWITHTAX': 'sum',
                        'ORDEREDQTY': 'sum'
                    })

                    tracker_summary.insert(0, 'Marketplace', 'Swiggy')
                    tracker_summary.rename(columns={
                        'PONUMBER': 'PO',
                        'CITY': 'Location',
                        'POCREATEDAT': 'Order Date',
                        'POEXPIRYDATE': 'Expiry Date',
                        'POLINEVALUEWITHTAX': 'PO Value',
                        'ORDEREDQTY': 'PO Qty'
                    }, inplace=True)

                    # Format PO Value
                    tracker_summary['PO Value'] = tracker_summary['PO Value'].apply(lambda x: f"₹ {self.format_indian(x)}")

                    # Create separate folder for Swiggy PO workbooks
                    timestamp = datetime.now().strftime("%d_%m_%Y__%H_%M_%S")
                    output_folder = Path(file_path).parent / f"Swiggy_PO_Workbooks_{timestamp}"
                    output_folder.mkdir(parents=True, exist_ok=True)

                    # Enhanced format function with proper EAN formatting
                    def format_sheet(ws, ean_columns=None):
                        from openpyxl.styles import Alignment
                        from openpyxl.utils import get_column_letter
                        
                        # Set all cells to center alignment
                        for row in ws.iter_rows():
                            for cell in row:
                                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        
                        # Format EAN columns as text to prevent scientific notation
                        if ean_columns:
                            for col_letter in ean_columns:
                                for cell in ws[col_letter]:
                                    cell.number_format = '@'  # Text format
                                    if cell.value and str(cell.value).replace('.', '').isdigit():
                                        cell.value = str(int(float(cell.value)))
                        
                        # Auto-adjust column widths
                        for col in ws.columns:
                            max_length = 0
                            col_letter = get_column_letter(col[0].column)
                            for cell in col:
                                try:
                                    if cell.value:
                                        max_length = max(max_length, len(str(cell.value)))
                                except:
                                    pass
                            adjusted_width = min(max_length + 4, 50)  # Cap at 50 to avoid extremely wide columns
                            ws.column_dimensions[col_letter].width = adjusted_width

                    # Create individual workbook per PO
                    from openpyxl import Workbook
                    from openpyxl.styles import Font

                    for po_number in tracker_summary['PO']:
                        po_data = df[df['PONUMBER'] == po_number]
                        city = po_data['CITY'].iloc[0]

                        wb = Workbook()

                        # --- PO Sheet ---
                        ws_po = wb.active
                        ws_po.title = "PO Sheet"
                        ws_po.append(['EAN', 'SKUDESCRIPTION', 'UNITBASEDCOST', 'ORDEREDQTY'])

                        for idx, row in po_data.iterrows():
                            ean_value = str(int(float(row['EAN']))) if pd.notna(row['EAN']) else ""
                            ws_po.append([
                                ean_value,
                                row['SKUDESCRIPTION'],
                                row['UNITBASEDCOST'],
                                row['ORDEREDQTY']
                            ])
                        
                        format_sheet(ws_po, ean_columns=['A'])

                        # --- Scan Sheet ---
                        ws_scan = wb.create_sheet("Scan Sheet")
                        ws_scan.append(["Scan", "EAN", "Title", "PO Qty", "Qty", "Box"])
                        
                        for r in range(2, len(po_data) + 20):
                            ws_scan[f'B{r}'] = f'=IF(A{r}="","",TEXT(INDEX(\'PO Sheet\'!A:A,MATCH(A{r},\'PO Sheet\'!A:A,0)),"0"))'
                            ws_scan[f'C{r}'] = f'=IF(A{r}="","",INDEX(\'PO Sheet\'!B:B,MATCH(A{r},\'PO Sheet\'!A:A,0)))'
                            ws_scan[f'D{r}'] = f'=IF(A{r}="","",INDEX(\'PO Sheet\'!D:D,MATCH(A{r},\'PO Sheet\'!A:A,0)))'
                        
                        format_sheet(ws_scan, ean_columns=['A', 'B'])

                        # --- Packing Slip Sheet ---
                        ws_packing = wb.create_sheet("Packing Slip")
                        
                        # Insert header row with PO + Marketplace + City
                        ws_packing['A1'] = f"{po_number} Swiggy {city}"
                        ws_packing['A1'].font = Font(bold=True, size=12)
                        ws_packing.merge_cells('A1:D1')
                        
                        # Column headers in row 2
                        ws_packing.append(["Scan", "Title", "Box", "Total"])
                        
                        # Add formulas starting from row 3
                        for r in range(3, len(po_data) + 3):
                            scan_row = r - 1  # corresponding row in Scan Sheet
                            ws_packing[f'A{r}'] = f'=IF(\'Scan Sheet\'!A{scan_row}="","",TEXT(\'Scan Sheet\'!A{scan_row},"0"))'
                            ws_packing[f'B{r}'] = f'=IF(\'Scan Sheet\'!A{scan_row}="","",\'Scan Sheet\'!C{scan_row})'
                            ws_packing[f'C{r}'] = f'=IF(\'Scan Sheet\'!A{scan_row}="","",\'Scan Sheet\'!F{scan_row})'
                            ws_packing[f'D{r}'] = f'=IF(\'Scan Sheet\'!A{scan_row}="","",\'Scan Sheet\'!D{scan_row})'
                        
                        # Grand total
                        total_row = len(po_data) + 3
                        ws_packing[f'D{total_row}'] = f"=SUBTOTAL(109,D3:D{total_row-1})"
                        ws_packing[f'D{total_row}'].font = Font(bold=True)
                        
                        format_sheet(ws_packing, ean_columns=['A'])

                        # Save workbook
                        wb_path = output_folder / f"{po_number}.xlsx"
                        wb.save(wb_path)

                    self.show_summary_popup(tracker_summary, marketplace)
                    messagebox.showinfo("Success", f"Swiggy PO workbooks created successfully in folder:\n{output_folder}")
                    if messagebox.askyesno("Open Folder", "Do you want to open the output folder?"):
                        os.startfile(output_folder)

                except Exception as e:
                    messagebox.showerror("Error", f"Something went wrong with Swiggy processing:\n{e}")

            # --- Zepto (Future Placeholder) ---
            elif marketplace == "Zepto":
                messagebox.showinfo("Info", "Zepto integration is coming soon!")

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n{e}")

def main():
    root = tk.Tk()
    app = POReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
