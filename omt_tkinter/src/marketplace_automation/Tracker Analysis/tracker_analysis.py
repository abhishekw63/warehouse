import pandas as pd
import numpy as np
from datetime import datetime
from tqdm import tqdm  # for progress bar

class POTracker:
    def __init__(self, file_name, sheet_name="OnlineB2B"):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.df = pd.read_excel(file_name, sheet_name=sheet_name)
        print(f"Loaded sheet '{sheet_name}' with shape {self.df.shape}")

    def column_optimization(self):
        """Provides unique data for selected columns and writes it to a new sheet."""
        cols_of_interest = [
            'Marketplace', 'Status', 'Courier Name', 'Mode', 
            'Ops Status', 'Logistics Status', 'Invoice Uploading'
        ]

        unique_dict = {}
        for col in tqdm(cols_of_interest, desc="Processing unique columns"):
            if col in self.df.columns:
                unique_dict[col] = self.df[col].dropna().unique().tolist()
            else:
                unique_dict[col] = ["Column not found"]

        max_len = max(len(v) for v in unique_dict.values())
        for k in unique_dict:
            unique_dict[k] += [''] * (max_len - len(unique_dict[k]))

        unique_df = pd.DataFrame(unique_dict)
        with pd.ExcelWriter(self.file_name, engine="openpyxl", mode="a") as writer:
            unique_df.to_excel(writer, sheet_name="Unique Data", index=False)
        print("✅ Unique data written to sheet 'Unique Data'")

    def dispatched_without_appointment(self):
        """Finds POs dispatched but appointment date not confirmed, excluding Delivered/RTO."""
        ops_status = self.df['Ops Status'].astype(str).str.strip().str.lower()
        app_date = self.df['App Date'].replace(r'^\s*$', np.nan, regex=True)
        logistics_status = self.df['Logistics Status'].astype(str).str.strip().str.lower()

        filtered_df = self.df[
            (ops_status == 'dispatched') &
            (app_date.isna()) &
            (~logistics_status.isin(['delivered', 'rto delivered']))
        ]

        print(f"Found {filtered_df.shape[0]} POs dispatched but no appointment date.")
        timestamp = datetime.now().strftime("%d_%m_%Y__%H_%M_%S")
        output_file = f"Dispatched_No_App_{timestamp}.xlsx"
        filtered_df.to_excel(output_file, index=False)
        print(f"✅ File '{output_file}' created successfully.")
        return filtered_df

    # 🧾 1️⃣ Invoiced on a given date
    def invoiced_on_date(self, invoice_date):
        """Filters POs invoiced on a given date."""
        date_str = pd.to_datetime(invoice_date).date()
        df_copy = self.df.copy()
        df_copy['Invoice Date'] = pd.to_datetime(df_copy['Invoice Date'], errors='coerce').dt.date

        filtered_df = df_copy[df_copy['Invoice Date'] == date_str]
        print(f"Found {filtered_df.shape[0]} POs invoiced on {date_str}")

        timestamp = datetime.now().strftime("%d_%m_%Y__%H_%M_%S")
        output_file = f"Invoiced_{date_str}_{timestamp}.xlsx"
        filtered_df.to_excel(output_file, index=False)
        print(f"✅ File '{output_file}' created successfully.")
        return filtered_df

    # 🚚 2️⃣ Dispatched on a given date
    def dispatched_on_date(self, dispatch_date):
        """Filters POs dispatched on a given date."""
        date_str = pd.to_datetime(dispatch_date).date()
        df_copy = self.df.copy()
        df_copy['Dispatch Date'] = pd.to_datetime(df_copy['Dispatch Date'], errors='coerce').dt.date

        filtered_df = df_copy[df_copy['Dispatch Date'] == date_str]
        print(f"Found {filtered_df.shape[0]} POs dispatched on {date_str}")

        timestamp = datetime.now().strftime("%d_%m_%Y__%H_%M_%S")
        output_file = f"Dispatched_{date_str}_{timestamp}.xlsx"
        filtered_df.to_excel(output_file, index=False)
        print(f"✅ File '{output_file}' created successfully.")
        return filtered_df

    # 📦 3️⃣ RTD (Return To Depot) on a given date
    def rtd_on_date(self, rtd_date):
        """Filters POs with RTD (Return To Depot) on a given date."""
        date_str = pd.to_datetime(rtd_date).date()
        df_copy = self.df.copy()
        df_copy['Dispatch Date'] = pd.to_datetime(df_copy['Dispatch Date'], errors='coerce').dt.date

        filtered_df = df_copy[df_copy['Dispatch Date'] == date_str]
        print(f"Found {filtered_df.shape[0]} POs RTD on {date_str}")

        timestamp = datetime.now().strftime("%d_%m_%Y__%H_%M_%S")
        output_file = f"RTD_{date_str}_{timestamp}.xlsx"
        filtered_df.to_excel(output_file, index=False)
        print(f"✅ File '{output_file}' created successfully.")
        return filtered_df

    # 🔍 4️⃣ Compare current vs reference sheet for tracking changes
    def compare_with_reference(self, reference_file):
        """Compares current file against a reference version and logs all cell-level changes."""
        print("\n🔎 Starting comparison with reference file...")

        # Load reference and current data
        ref_df = pd.read_excel(reference_file, sheet_name=self.sheet_name)
        ref_df["Composite_Key"] = (
            ref_df["PO"].astype(str).str.strip() + "_" +
            ref_df["PO Qty"].astype(str).str.strip()
        )

        curr_df = self.df.copy()
        curr_df["Composite_Key"] = (
            curr_df["PO"].astype(str).str.strip() + "_" +
            curr_df["PO Qty"].astype(str).str.strip()
        )

        # Merge reference and current based on composite key
        merged = pd.merge(ref_df, curr_df, on="Composite_Key", suffixes=("_old", "_new"), how="outer", indicator=True)
        changes = []

        # Convert to records to avoid iterrows overhead
        for row in tqdm(merged.to_dict('records'), desc="Comparing records"):
            if row["_merge"] == "left_only":
                changes.append([row["Composite_Key"], "REMOVED", "", "", ""])
            elif row["_merge"] == "right_only":
                changes.append([row["Composite_Key"], "ADDED", "", "", ""])
            else:
                for col in ref_df.columns:
                    if col not in ["Composite_Key"]:
                        old_val = row.get(f"{col}_old", np.nan)
                        new_val = row.get(f"{col}_new", np.nan)
                        # Compare both NaN and string differences
                        if pd.notna(old_val) or pd.notna(new_val):
                            if str(old_val).strip() != str(new_val).strip():
                                changes.append([
                                    row["Composite_Key"],
                                    "MODIFIED",
                                    col,
                                    old_val,
                                    new_val
                                ])

        # Create and export change log
        changes_df = pd.DataFrame(changes, columns=["Composite_Key", "Change_Type", "Column", "Old_Value", "New_Value"])

        timestamp = datetime.now().strftime("%d_%m_%Y__%H_%M_%S")
        output_file = f"Change_Log_{timestamp}.xlsx"
        changes_df.to_excel(output_file, index=False)

        print(f"✅ Change log written to '{output_file}' with {len(changes_df)} records.")
        return changes_df


# ✅ Usage Example
file_name = "Reference_PO_Format.xlsx"
tracker = POTracker(file_name)

# Example calls
tracker.compare_with_reference("New PO format.xlsx")
# tracker.invoiced_on_date("2025-10-15")
# tracker.dispatched_on_date("2025-10-15")
# tracker.rtd_on_date("2025-10-15")
