# Online PO Management — Architecture & Flow

> **Living document.** This is the map of how the code actually works.
> **Thumb rule: every new development updates this file** — add/adjust the
> relevant flow, the component table, and a line in the Changelog at the
> bottom. Keep the diagrams honest; if a diagram and the code disagree, the
> code is right and the diagram must be fixed.

Last updated: 2026-06-13 · Covers: Manual + Auto modes, history DB (SQLite→MySQL), DB-sourced tracker.

---

## 1. What this tool does

Converts a marketplace's Purchase Order dump (Excel / CSV / PDF) into
**D365 Business Central importable Sales Orders (SO)** or **Transfer Orders
(TO)**, validates pricing against our master, and records every order in a
**history database** so the same PO is never uploaded twice.

Two ways to run it:

```
                    ┌─────────────────────────────────────────┐
                    │            OnlinePOApp (GUI)             │
                    └─────────────────────────────────────────┘
                         │                          │
            ┌────────────┘                          └─────────────┐
            ▼                                                      ▼
   ▶ Generate SO  (MANUAL mode)              ⚙ Auto Mode (AUTO mode)
   one file you pick, one marketplace        every Dump/Online/<mp> folder,
                                             unattended, all marketplaces
            │                                                      │
            └───────────────────┬──────────────────────────────────┘
                                ▼
                    Same engine + exporters + history
```

---

## 2. The core pipeline (shared by both modes)

Everything funnels through one engine call that turns a file into a
`ProcessingResult` (the in-memory representation of one marketplace's POs):

```
 input file ──► MarketplaceEngine.process()
                   │  (load excel/csv/pdf → DataFrame)
                   │  alias-resolve columns → validate → SO/TO branch
                   ▼
            ProcessingResult            ◀── the single source for all outputs
                   │   .rows  : list[SORow]  (one per item line)
                   │   .warnings, .marketplace, .warehouse_code, .raw_df …
                   ▼
   ┌──────────────┬──────────────┬───────────────┬───────────────┐
   ▼              ▼              ▼               ▼               ▼
 SO workbook   D365 export   Tracker rows    History (DB)    Email report
 (.xlsx)       (Headers/     (build_tracker  (order_headers) (optional)
               Lines)        _rows)
```

**Key idea:** the `ProcessingResult` is computed **once**; every output is
derived from it. `build_tracker_rows(result)` is the single function that
shapes per-PO tracker/ history rows — so the tracker and the DB never drift.

---

## 3. Package map (what lives where)

```
online_po_management/
├── main.py                    # launcher → online_po_processor.app.main()
└── online_po_processor/
    ├── app.py                 # bootstrap: logging, expiry check, launch GUI
    │
    ├── config/                # ── configuration (no logic) ──
    │   ├── marketplaces.py     #   MARKETPLACE_CONFIGS (the registry), warehouses
    │   ├── constants.py        #   expiry date, misc constants
    │   ├── paths.py            #   bundled master/mapping locations
    │   └── email_config.py     #   SMTP settings
    │
    ├── data/                  # ── pure data + loaders ──
    │   ├── models.py           #   SORow, ProcessingResult (dataclasses)
    │   ├── master_loader.py    #   Items Master (EAN→Item/MRP/GST), _clean_code
    │   └── mapping_loader.py   #   Ship-To B2B (location→customer/ship-to)
    │
    ├── engine/                # ── the brain ──
    │   ├── marketplace_engine.py  # process(), process_multi(),
    │   │                          #   process_consignments(), validation,
    │   │                          #   per-line margin, item/EAN resolution
    │   ├── avenue_pdf_parser.py   # Dmart/Avenue PDF → DataFrame
    │   ├── firstcry_pdf_parser.py # FirstCry PDF → DataFrame
    │   └── reliance_pdf_parser.py # Reliance PDF → DataFrame
    │                              #   (PDF parsers registered in PDF_PARSERS)
    │
    ├── exporter/              # ── result → files ──
    │   ├── so_exporter.py       #   writes the per-marketplace SO workbook
    │   ├── d365_exporter.py     #   writes the D365 import package
    │   ├── _styles.py           #   shared openpyxl styles/helpers
    │   └── sheets/              #   one module per sheet in the SO workbook:
    │       ├── headers_sheet.py #     Headers (SO/TO)  ← D365 import
    │       ├── lines_sheet.py   #     Lines (SO/TO)    ← D365 import
    │       ├── summary_sheet.py #     per-PO human summary (RK inc-GST col)
    │       ├── tracker_sheet.py #     build_tracker_rows() + per-mp Tracker
    │       ├── validation_sheet.py #  Vendor vs Our MRP/Landing/CP
    │       ├── warnings_sheet.py
    │       └── raw_data_sheet.py
    │
    ├── auto/                  # ── AUTO mode + history (v2.4.0) ──
    │   ├── auto_runner.py        #  headless batch over Dump/Online/<mp>/
    │   ├── consolidated_exporter.py # combined Headers/Lines/Summary/
    │   │                            #   Validation + export_tracker_from_db()
    │   └── history_db.py         #  HistoryStore interface + Sqlite/MySQL
    │                             #    impls + record_history/record_manual
    │
    ├── gui/                   # ── Tkinter UI ──
    │   ├── app_window.py        #   OnlinePOApp — MANUAL mode (the main window)
    │   ├── auto_window.py       #   AutoWindow  — AUTO mode (per-mp warehouse)
    │   ├── _file_row.py, _update_dialog.py
    │
    ├── emailer/               # HTML report builder + SMTP sender
    └── utils/                 # platform_open (open files cross-platform)
```

---

## 4. MANUAL mode (`OnlinePOApp.generate`)

The original flow — pick one file, one marketplace, generate one SO.
**Unchanged in behaviour**; history + tracker were *added* around it.

```
 user picks marketplace + warehouse + PO file → ▶ Generate SO
        │
        ▼
 load mapping (party) + master  →  MarketplaceEngine.process(file, config)
        │
        ▼
 ProcessingResult  ──► SOExporter.export()  → <input>/output/<mp>_so_<ts>.xlsx
        │                  (Headers, Lines, Summary, Validation, Warnings, Raw)
        │
        ├──► record_manual(result)      → history DB (mode=MANUAL)  [v2.4.0]
        └──► export_tracker_from_db(run) → Dump/Tracker/Online/...   [v2.4.0]
        │
        ▼
 success popup shows: items / PO / qty / History: ✓ recorded (N new, M dup)
 Buttons enabled: Open Last Output · Export D365 · Send Email
 Always available: 📜 View Order History
```

---

## 5. AUTO mode (`AutoWindow` → `AutoRunner`)

Drop each marketplace's dump in `Dump/Online/<Marketplace>/`, set each
marketplace's AHD/BLR warehouse, hit **Run** — unattended, no per-file clicks.

```
 Dump/Online/
   ├── RK/        POItemExport_*.xls
   ├── Zepto/     zepto.xlsx
   ├── Reliance/  *.pdf
   ├── Flipkart-TO/  Consignment_Details_*.csv   (consignment)
   └── …(12 marketplaces)…

 AutoWindow.Run ──► AutoRunner.run(Dump/Online)
        │  for each <mp> folder:
        │     pick engine path by config:
        │        consignment_mode  → process_consignments(all csvs)
        │        source_format=pdf → process_multi(all pdfs)
        │        else              → process(each excel/csv)
        │     stamp per-marketplace warehouse (AHD/BLR)
        │     SOExporter.export() → <mp>/output/<mp>_so_<ts>.xlsx
        ▼
   list[MarketplaceRun]  (each holds its ProcessingResult)
        │
        ▼  (worker, in order)
   1) record_history(runs)        → history DB (mode=AUTO), returns run_id
   2) export_tracker_from_db(run) → Dump/Tracker/Online/Online_Tracker_<ts>.xlsx
   3) export_consolidated(runs)   → Dump/Online/_Consolidated/consolidated_<ts>.xlsx
        │
        ▼
   Auto window log: per-mp summary + duplicates flagged
   Buttons: 📘 Open Consolidated · 📋 Open Tracker · 📜 View History
```

**Consolidated workbook** (one file for the whole run, for D365 + review):
`Overall Summary` · `Headers (SO)` · `Lines (SO)` · `Summary (SO)` ·
`Validation (SO)` · *(+ the TO variants only when TO marketplaces ran)*.
SO and TO are **never mixed** — separate sheets (different D365 imports).

---

## 6. History & deduplication (the DB layer)

> Goal: *"track which orders we are uploading"* and never double-upload.

```
                       db_config.json   (LOCAL: %LOCALAPPDATA%\OnlinePOProcessor\,
                          backend=mysql      NOT in the repo — holds the password)
                              │
                              ▼
   record_history / record_manual ──► get_history_store()  ◀── single swap point
                              │            │
                              │     backend == 'mysql' ?
                              │       │yes            │no / MySQL down
                              ▼       ▼               ▼
                       order rows  MySqlHistoryStore  SqliteHistoryStore
                                    (renee_orders)     (Dump/Tracker/history.db)
                                        │
                                        ▼
                              ┌──────────────────────┐
                              │  runs                 │ one row per run (batch header)
                              │  order_headers        │ one row per (marketplace, PO)
                              └──────────────────────┘
```

- **Interface:** `HistoryStore` (ABC). Implementations: `SqliteHistoryStore`,
  `MySqlHistoryStore`. **To add Postgres/SQL Server: write one new class +
  point `get_history_store` at it — no caller changes.**
- **Dedup-skip** (`apply_dedup`, gated by `constants.DEDUP_SKIP_ENABLED`,
  on for ALL marketplaces): before export, each PO is checked against
  `existing_pos()`. Already-uploaded POs are **removed from `result.rows`**
  (so they never reach Headers/Lines / D365) and summarised on
  `result.skipped_orders` → the **"Skipped POs" output sheet**. The DB is
  **not** told about duplicates — it holds **only new POs**, no duplicate
  tracking. Trust model: a PO counts as uploaded the moment it's generated.
- **Backend now = MySQL** (`renee_orders`), viewable live in MySQL Workbench.
  Falls back to SQLite automatically if MySQL is unreachable.

### `order_headers` columns (NEW POs only — no duplicate tracking)
`order_id`(PK) · `run_id` · `run_ts` · `mode`(ENUM AUTO/MANUAL) ·
`segment`(`OnlineB2B`; future offline/GT shares the DB) ·
`marketplace` · `marketplace_label` · `po`(VARCHAR) · `location` ·
`warehouse` · `po_date`(DATE) · `exp_date`(DATE) · `order_type`(ENUM SO/TO) ·
`items` · `qty` · `order_value`(DECIMAL, **GST-inclusive**) · `output_file` ·
`created_at`.
(`is_duplicate` / `first_seen_ts` were removed — duplicates aren't stored.)
Future child table `order_lines` (item-level) keys on `order_id`.

---

## 7. The Tracker (DB-sourced, single source of truth)

The Tracker is the paste-ready list for the warehouse's master "New PO
format" sheet. **It is generated FROM the history DB, not recomputed** — so
it can never disagree with what's recorded.

```
 record to DB (run_id) ─► export_tracker_from_db(run_id)
                              │  store.fetch_orders(run_id, only_new=True)
                              ▼
                  Dump/Tracker/Online/Online_Tracker_<ts>.xlsx
                  9 cols: Segment | Market Place | PO | Location | PO Date |
                          Exp Date | PO Aging | Order Value | Order Qty
```

- **`only_new=True`** → lists only this run's *new* POs (duplicates already
  pasted before are excluded — no re-pasting).
- Marketplace labels are mapped to the master tracker's names
  (`Blink→Blinkit`, `Firstcry→First Cry`, `Flipkart-TO→Flipkart Branch`,
  `Meesho-TO→Meesho-SB`, …) in `tracker_sheet._MARKETPLACE_DISPLAY`.
- Dates are written as **real Excel date values** with a `DD-MM-YYYY` number
  format (not text) — so they display day-first, stay reformattable/sortable,
  and paste into the master tracker as real dates. `tracker_sheet._coerce_date`
  accepts date objects (MySQL), day-first text, and Blink's ISO
  `order_date`/`expiry_date`.

---

## 8. Config-driven marketplaces

A marketplace is one entry in `config/marketplaces.py → MARKETPLACE_CONFIGS`.
Key fields drive everything (no per-marketplace code unless it's a new PDF):

| Field | Meaning |
|---|---|
| `party_name` | label used in the Ship-To mapping |
| `po_col`, `loc_col`, `qty_col`, `ean_col` | column names in the dump |
| `item_resolution` | `from_ean` (all today) vs `from_column` |
| `fob_col`, `compare_basis` | price column + `cost`/`landing` validation basis |
| `amount_col` | order value column; `amount_is_pre_gst` (RK) → grossed up |
| `source_format` | `excel` / `pdf`; `pdf_parser` names the parser |
| `consignment_mode` | TO bulk-CSV mode (Flipkart-TO, Meesho-TO) |
| `output_type` | `so` (default) / `to` |
| `margin_rules` | per-line margin by category (Nykaa) |

**To add a marketplace:** add a config entry; for a PDF, also write a parser
and register it in `engine/marketplace_engine.PDF_PARSERS`. Then add it to a
`Dump/Online/<Name>/` folder for Auto mode.

---

## 9. Data sources (bundled, auto-loaded)

- **Items Master** (`Calculation Data/Items March.xlsx`, sheet `Item Master`):
  EAN → Item No / MRP / GST / HSN / Description. Read by `MasterLoader`.
- **Ship-To B2B mapping**: marketplace location → D365 Customer No / Ship-To
  code. Read by `MappingLoader`, filtered by `party_name`.
- Both are bundled and auto-loaded on startup; updatable via the GUI.

---

## 10. Changelog (append one line per development)

- **2026-06-13** — **Dedup-skip**: already-uploaded POs are now **removed**
  from Headers/Lines (not re-sent to D365) and listed on a new **"Skipped
  POs"** output sheet; the DB stores **only new POs** (removed
  `is_duplicate` + `first_seen_ts` columns via migration, rows kept). New
  `apply_dedup` + `DEDUP_SKIP_ENABLED` + `existing_pos()`; `record` no
  longer returns duplicates.
- **2026-06-13** — Added a **`Segment`** column (`OnlineB2B`) to the tracker
  (col 1) and `order_headers` (migration + back-fill of existing rows) for
  future offline/GT orders in the same DB/tracker. Const in
  `config.constants.ORDER_SEGMENT`. Also fixed Flipkart PO date (added
  bare `Date` to the candidates) — all marketplaces now emit real dates.
- **2026-06-13** — Tracker dates now written as **real Excel `DD-MM-YYYY`
  dates** (were text, so Excel couldn't reformat them); added
  `tracker_sheet._coerce_date` / `_write_date_cell`.
- **2026-06-13** — MySQL backend (`MySqlHistoryStore`) behind `HistoryStore`;
  config in local AppData; Manual mode records to history + shows it
  (popup + 📜 View Order History button); Tracker made **DB-sourced**
  (`export_tracker_from_db`, only-new) and removed from the consolidated
  workbook; Blink date fix (underscore column match + ISO date parsing);
  GUI relaid out (2-column buttons, wider window); created this doc.
- **2026-06-13** — History/dedup added (SQLite): `runs` + `orders`, dedup
  flagging, `record_history`/`record_manual`, View History export.
- **2026-06-13** — Auto mode: `AutoRunner` (headless batch over
  `Dump/Online/<mp>`), per-marketplace warehouse, consolidated workbook,
  standalone tracker; `⚙ Auto Mode` button in the main window.
- **earlier (v2.3.1)** — Whole-number Item No/EAN fix; RK inc-GST Summary
  column + Tracker; Reliance/FirstCry Tracker sheets; Nykaa per-line margin;
  Reliance PDF rewrite; side-by-side Vendor/Our Validation.

<!-- When you add a feature: 1) update the relevant flow/diagram above,
     2) update the component table/column list if structures changed,
     3) add a dated line here. -->
