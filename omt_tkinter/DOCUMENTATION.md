# Marketplace Automation: Architecture & Scalability Guide

## Overview

This repository provides an automated desktop application for generating Purchase Order (PO) intelligence reports and email summaries across multiple e-commerce marketplaces (e.g., Blinkit, Flipkart, Swiggy).

The application utilizes **PyQt6** to deliver a premium, "iOS-like", frameless UI experience with highly responsive, background-threaded operations.

## Architecture & Technology Stack

- **UI Framework:** `PyQt6` is used to create smooth, scalable, and modern graphical interfaces. Native window controls are bypassed in favor of a `FramelessWindowHint` and a custom `DraggableTitleBar` to enforce branding and custom styling via QSS (Qt Style Sheets).
- **Zero to Hero Refactoring:** The main UI assembly logic has been refactored into logical sub-components (`_build_header`, `_build_content`, `_build_footer`) to enforce strict modularity and separation of concerns, ensuring high scalability and maintainability.
- **Asynchronous Execution:** Heavy I/O bound processing (like parsing huge Excel files using `pandas` and formatting them via `openpyxl`) is fully delegated to a `QThread` (`ReportWorker`). This decoupled design ensures the main UI thread never blocks, preventing the "Application Not Responding" operating system overlay.
- **Data Layer:** Standardized data manipulation relies on `pandas` to read, filter, clean, and aggregate data.
- **Notification Layer:** The `EmailService` class abstracts away SMTP server communication, transforming Pandas DataFrames into cleanly styled HTML tables before dispatch.
- **Configuration:** Controlled via `config.py` using `dotenv`, ensuring API keys, passwords, and sensitive emails are isolated from source code.

## How to Add a New Marketplace

The application is built to easily scale as the business onboards new vendors or marketplaces.

To add a new marketplace (e.g., *Zepto*):

1. **Create a Processing Module:**
   - Under `src/marketplace_automation/marketplaces/`, create a new file named `zepto.py`.
   - Implement a function `process_zepto(file_path: str) -> dict`.
   - This function should return a standardized dictionary containing:
     - `'marketplace'`: `"Zepto"`
     - `'output_file'`: Path to the generated Excel file.
     - `'has_sku_data'`: Boolean indicating if SKU insights were generated.
     - `'sku_count'`: Integer count of SKUs.
     - `'tracker_df'`: The cleaned `pandas.DataFrame` representing the core summary.
     - `'sku_df'`: The DataFrame of aggregated SKU data (or `None`).

2. **Register the Logic in the Worker Thread:**
   - In `src/marketplace_automation/marketplaces_automation.py`, locate the `ReportWorker.run()` method.
   - Import `process_zepto`.
   - Add an `elif self.marketplace == "Zepto": result = process_zepto(self.file_path)` block.

3. **Update the UI Dropdown:**
   - In `POReportApp._build_ui()`, locate the `self.marketplace_dropdown` widget initialization.
   - Simply add `"Zepto"` to the list of `addItems(...)`.
   - Remove any temporary `"Zepto"` mock/error traps inside `generate_report()`.

4. **Update the Summary Calculations:**
   - In `POReportApp.calculate_summary_data()`, add custom mapping logic (if required) for how the specific Zepto PO headers map to standardized values like `total_pos`, `total_units`, `min_date`, etc.

## UI Theming and Customization

The UI relies heavily on a centralized QSS string defined in `_apply_custom_styles`.
- **Shadows:** The main UI relies on an outer wrapper with padding (`self.wrapper`) and a `QGraphicsDropShadowEffect`. If extending the window size, ensure the layout constraints keep the padding intact, or shadows will appear clipped by the OS window manager.
- **Dialogs:** Never use standard `QMessageBox`. Always use `ui_smooth.SmoothDialog` methods (`show_info`, `show_error`, `ask_question`) to guarantee a seamless aesthetic.