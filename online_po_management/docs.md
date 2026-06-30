# Online PO Management — Architecture & Flow

> **Living document.** This is the map of how the code actually works.
> **Thumb rule: every new development updates this file** — add/adjust the
> relevant flow, the component table, and a line in the Changelog at the
> bottom. Keep the diagrams honest; if a diagram and the code disagree, the
> code is right and the diagram must be fixed.

Last updated: 2026-06-27 · Covers: Manual + Auto modes, web-only order store (web-owned lock/dedup, no order_issue_lines), DB-sourced tracker, channel_sku_map, unified exceptions overlay, pricing-rule banner, 18 marketplaces.

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
    │   ├── reliance_pdf_parser.py # Reliance PDF → DataFrame
    │   └── myntra_pdf_parser.py   # Myntra PDF → DataFrame (DUAL-format:
    │                              #   .pdf here OR .xlsx via the Excel path)
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
   ├── Flipkart/  purchase_order_*.xlsx   (one per PO → compiled in-app)
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
| `accepted_extensions` | file-picker filter; enables **dual-format** (Myntra: `['.xlsx', '.pdf']`) |
| `consignment_mode` | TO bulk-CSV mode (Flipkart-TO, Meesho-TO); `filename_loc_from_shipto` + `po_date_today` + `exp_date_offset_days` for Meesho |
| `output_type` | `so` (default) / `to` |
| `margin_rules` | per-line margin by category (Nykaa) |
| `gst_margin_discount` | GST-dependent keep% = 1−d×(1+GST) (Reliance) |
| `file_parser` | custom non-PDF reader routed by config key, not extension (Big Basket Excel-preamble, Purplle tab-separated `.XLS`, Flipkart per-PO portal `.xlsx`) |
| `loc_match='address'` | resolve ship-to by full postal ADDRESS (pincode + survey-no/village body overlap) instead of the generic name tiers (Flipkart — new portal drops the `Flipkart India Pvt. Ltd.,` prefix the Del Locations carry) |
| `item_resolution='from_swiggy_sku'` | dump has only a SkuCode → recover EAN via the master's `Swiggy` sheet, then resolve like `from_ean` (`sku_col`) |
| `po_total_col` | per-PO grand-total column used VERBATIM as tracker Order Value (Nykaa `PO Amount` = portal exactly) instead of summing per-line |
| `gst_pct_col` | the file's own GST% per row → used to gross up `amount_is_pre_gst` to inc-GST (Reliance, Big Basket) |

**Dual-format (Myntra, v2.4.0/2.4.1):** a marketplace can accept *both* Excel
and PDF. Myntra keeps `source_format='excel'` and adds `pdf_parser='myntra'`
+ `accepted_extensions=['.pdf','.xlsx']` (PDF leads → file picker defaults to
PDF). At load time `process()` checks the *uploaded file's extension* — a
`.pdf` whose config has a `pdf_parser` routes to the PDF parser; anything else
uses the Excel path. One config drives both: the PDF parser emits the same
column names (`GTIN`, `Quantity`, `Landing Price`, `List price…`, `Mrp`,
`Location`) and additionally injects real PO dates (`__po_date__`/
`__exp_date__`) the Excel punch lacks.

* **Multi-file:** because a registered `pdf_parser` makes
  `_supports_multi_file` true, Myntra is multi-file in both modes — the manual
  picker takes many PDFs and Auto mode combines every PDF in
  `Dump/Online/Myntra/` into one SO batch (via `process_multi`, which runs
  each file through `process()`). A folder of `.xlsx` still works (one run per
  dump); a folder of `.pdf` is one combined batch.
* **Parser robustness:** the Myntra PDF has vertical column rules but no
  horizontal row borders (and none on page 2+), so the parser derives the
  column x-boundaries once and reconstructs rows by anchoring on the SKU
  column — rejoining EANs pdfplumber splits across line-wraps. It handles
  **both PO layouts** — intra-state (CGST+SGST, 17 cols) and inter-state
  (single IGST, 15 cols) — by gating on the *mapped header fields*, not a
  fixed column count, and both 2-letter (`Bangalore,KA`) and full-name
  (`…,Haryana,…`) ship-to address tails. Every parse self-checks its line
  count/qty against the PO's own `Total Quantity` footer.

**Flipkart-TO Auto visibility (v2.4.1):** drop the per-PO
`Consignment_Details_<PO>_*.csv` files **and** the
`Consignment_Visibility_Report_*.csv` in `Dump/Online/Flipkart-TO/`. Auto mode
detects the report by `visibility_filename_regex`, excludes it from the
consignment set, and threads it into `process_consignments` for Location
resolution — matching the manual GUI (which took the report as a separate
pick). Without it, Locations fall back to empty as before.

**To add a marketplace:** add a config entry; for a PDF, also write a parser
and register it in `engine/marketplace_engine.PDF_PARSERS`. Then add it to a
`Dump/Online/<Name>/` folder for Auto mode.

---

## 9. Data sources (bundled, auto-loaded)

- **Items Master** (`Calculation Data/Items March.xlsx`, sheet `Item Master`):
  EAN → Item No / MRP / GST / HSN / Description. Read by `MasterLoader`. The
  same workbook also carries the **`Swiggy`** sheet (SkuCode→EAN, 261 entries)
  and **`Swiggy Deal SKUs`** sheet (per-EAN negotiated cost) — auto-loaded by
  `MasterLoader._load_swiggy_sheets`.
- **Ship-To B2B mapping**: marketplace location → D365 Customer No / Ship-To
  code. Read by `MappingLoader`, filtered by `party_name`. v2.4.0 adds a
  reverse index `by_shipto` (look up by the Transfer-to CODE itself, e.g.
  `MS_BLR`) — used by Meesho-TO whose destination comes from the filename.
- **Master Exceptions** (`Calculation Data/Master Exceptions.xlsx`, OPTIONAL,
  auto-loaded if present beside the master): ONE central overlay file the
  operator edits — no per-config/per-master patching. Two kinds:
  * **item alias** (`Source Code`→`Maps To`): an EAN that isn't in the master
    verbatim but exists under a variant key (FirstCry `…885`→`…885_1`).
    Applied in `MasterLoader.lookup()` after a direct miss.
  * **price override** (`Override MRP`/`Override Margin %`, optional
    `Marketplace` scope): a deal price where MRP×default-margin would wrongly
    flag MISMATCH (Blinkit EPISENSE: 24% off a ₹899 deal MRP). Applied in
    `_validate_against_master` via `MasterLoader.price_override()`.
  Every applied exception is recorded on `result.exceptions_applied` and
  surfaced on the **Exceptions** output sheet (`exceptions_sheet.py`).
- All are bundled and auto-loaded on startup; updatable via the GUI.

### Pricing-rule banner (Summary sheet)
The Summary footer shows **`Pricing: <rule>`** — the ACTUAL rule, not a
misleading flat % (`summary_sheet.pricing_rule_str`): `70% straight`,
Nykaa's `category rule — Perfume/Fragrance 69% · Cosmetics 66%`, Reliance's
`GST-based — …`, and ` · N deal/override SKU(s) applied` when any fired. The
Validation `Our Landing (…)` header / info row use the same rule-aware label
(`validation_sheet._margin_label`: `per rule` / `GST-based` / straight %)
instead of a single number that lies for rule-based marketplaces.

---

## 10. Changelog (append one line per development)

- **2026-06-30** — **Email skeleton + Issues "Email" feature.** New reusable
  `online_b2b/services/mailer.py` — ONE SMTP layer (`send_html` + `EmailReport`
  base) that **reuses the frozen desktop app's `email_config.get_email_config`**
  verbatim (same Gmail sender/recipients, `Calculation Data/email_config.json`
  override). Every future email feature subclasses `EmailReport` (subject/html/
  to/cc → preview/send). First consumer: `issue_email.IssuesEmailReport` — emails
  the **currently-filtered** issue lines (PO/item/EAN/qty/CP/diff/status +
  **Action: Excluded/Override/Kept** + remark) to management. Issues page gets a
  **✉ Email** button → **preview modal** (subject, To/Cc, full HTML in an iframe)
  → **Send**. Views `issues_email_preview` (GET, render-only) + `issues_email_send`
  (POST); both read the same `_issue_filters` as the page/export. The email modal
  + JS are generic (point `data-preview-url`/`data-send-url` at any future report).
  See [[dry-skeleton-first]]. Additive; never touches the frozen engine.

- **2026-06-30** — **First Cry web integration + item-master NaT fix + reusable
  loading overlay + SOP tab.** (1) **First Cry** is now a live web pilot: added
  `'Firstcry'` to `PILOT_MARKETPLACES` (label "First Cry") and a
  `FirstcryProcessor(Processor)` whose `_source_dates_by_po()` backfills po_date/
  exp_date from the FirstCry PDF header (`parse_firstcry_pdf` → `PO Date` /
  `PO Expiry Date`), mirroring `DmartProcessor`; hub chip flipped soon→live.
  Verified E2E on a sample PDF (1 PO / 16 lines / ₹53,941.98, dates backfilled).
  (2) **Item-master MRP bug** — blank Start/End cells parse to `NaT` (which is
  *truthy*), so `_read_effective_mrp` fell into `NaT <= today` → "Cannot compare
  NaT with datetime.date" on Preview & rebuild. Fixed with `pd.notna()` guards
  everywhere a window date is tested/compared; NaT now stored as `None`.
  (3) **Reusable processing overlay** — `core/base.html` gains a `body_end` block;
  `base_b2b.html` injects `#b2b-load` (spinner + message) that any element opts
  into via `data-loading="…"` (forms) / `data-loading-click="…"`. Wired on the
  Item-Master Preview & rebuild form. (4) Operator **SOP** moved to its own tab in
  Rules & Exceptions (applies to both segments). [[tracker-dates-and-tat]]
  [[item-master-in-db]]
- **2026-06-28** — **Rules & Exceptions split into Online B2B / Offline segments;
  GT Mass file rules drafted.** Top-of-page segment tabs: **Online B2B** holds all
  existing content (validation, per-marketplace rules, formats, exceptions,
  decisions, Flipkart map); **Offline** has a full **GT Mass — file rules &
  regulations** card: hard rules (first-sheet-only, BC Code+Order Qty header, PO
  Number label+value) each paired with the **exact error thrown**
  (`Template violation: …`), soft rules (Location/Distributor/City/State blank,
  EAN-only rescue, EAN-not-in-master) with their warning text, the columns the
  engine reads, meta fields, and best practice. Mirrors the actual
  `offline/utils.py:TemplateValidator` logic (reads Sheet1 only).
- **2026-06-28** — **Analytics → SKU demand (qty + value, top-10).** New section on
  the Analytics page (`order_db.sku_analytics`) over uploaded POs: overall demanded
  **qty**, **value** (Σ qty × unit_price), distinct **SKUs / POs**, and **Top 10 SKUs
  by Qty** and **by Value**. Own filter bar — **marketplace** + **upload-date range**
  (`run_ts`), **defaulting to today's uploads** (`sku_from`/`sku_to`/`sku_mp` query
  params; "Today" / "All time" buttons). Note: alias `lines` is a MySQL reserved
  word — use `nlines`. Enhanced: polished cards (rank badges + proportion bars +
  hover), **click-to-sort** Qty/Value/POs (asc/desc, client-side via `data-v`),
  and a **Full view** page `/b2b/analytics/sku/` (`SkuDemandView`) listing every
  SKU with the same filters + **CSV export** (`b2b_sku_demand_export`).

- **2026-06-28** — **Big Basket enabled in the web app + Project Map "Keep in
  sync".** Big Basket was engine-ready (config + `bigbasket` parser) but not in
  the web layer; enabled by adding `Bigbasket` to `engine_bridge.PILOT_MARKETPLACES`
  (+ `PILOT_LABELS` → "Big Basket") and flipping its `ONLINE_CHANNELS` chip
  `soon → live`; added a Big Basket sample to `template_samples.json` so Rules →
  "See full template" works. Verified end-to-end (4 POs / 10 lines). (Reliance was
  already live — no change.) **New:** Project Map → **🔗 Keep in sync** tab — a
  curated change-impact / coupling map ("change here ⇒ also change there") for
  adding an online marketplace, an offline channel, a DB column, or a page, with
  the always-do rules (docs changelog, Rules/See-full-template, never touch the
  frozen engine) and which surfaces auto-update.

- **2026-06-28** — **Item-master update diff + 15-day refresh reminder.** Item
  Master upload **preview** now shows a "What will change" section
  (`iml.diff_against_current(rows)`): **new** items, **MRP changed** (old → new),
  **removed** (live non-manual item absent from the file), unchanged — with a
  clear "**Nothing to update**" when new=0 & MRP-changed=0, so the operator sees
  exactly what the update touches before confirming. Plus `iml.last_updated()`
  (`MAX(updated_at)`) drives a **Hub reminder** banner ("Items & MRP refresh due
  · last updated N days ago") that appears at **≥ 15 days**. Additive — reads/
  compares around the existing upload; the rebuild logic is unchanged.

- **2026-06-28** — **Project Map page (graphical, auto-generated).** Staff-only
  `/dev/map/` (`ProjectMapView`, linked from Dev · Health). `core.project_map`
  introspects the LIVE codebase each load — so it updates whenever code changes —
  to render: a collapsible **file tree** (apps → modules → templates, colour-coded
  by type, frozen engine flagged), the real **URL→view routes** (from Django's
  resolver, grouped by app), the **DB models/tables** (app registry; managed vs
  external), and a graphical **data-flow** diagram (Upload → Process → Review →
  Confirm → renee_orders → Dashboards; who-feeds-the-pipeline; architecture
  layers). Read-only; additive.

- **2026-06-28** — **Dev · Health page (observability + code audit).** Additive
  perf capture + a staff-only dashboard at `/dev/` (`DevDashboardView`,
  `user.is_staff`). `core.observability.PerfMiddleware` (registered last in
  MIDDLEWARE) times **every** request and appends a JSON line to `logs/perf.jsonl`
  (path, status, ms, SQL count + time via `execute_wrapper`, bytes, user) — never
  touches the business DB, never alters a response. The page shows perf KPIs,
  per-endpoint aggregates (avg/p95/max ms, queries, flags: slow / N+1? / large),
  recent requests, and an on-demand **all-angles code audit**
  (`core.code_audit`): ruff (standards), AST metrics (big files/functions),
  duplication scan, TODO/FIXME, and a high-signal **security** scan (eval/exec/
  shell=True/pickle/verify=False/mark_safe/SQL-string-building). `logs/`
  git-ignored. Sidebar shows a staff-only "Dev · Health" link; GT Mass sidebar
  link now points to the new `/offline/gt-mass-flow/`.

- **2026-06-28** — **Shared PO-flow scaffold; GT Mass migrated onto it.** New
  reusable `online_b2b/services/po_flow.py` (`FlowSpec` + token store + cached
  preview + decisions + confirm + discard/download) and shared templates
  `po_flow/upload.html` + `po_flow/review.html` give every segment ONE
  import→review→confirm flow. **Capability flags** (`warehouse/margin/marketplace/
  vendor_cols/override/ean_fix/exclude/d365`) + a **null `extra_partial` slot**
  keep one template working across mismatched channels (online has vendor-compare
  + Override + D365; GT Mass has none but has file-level exceptions). GT Mass now
  uses it (`offline/flows.py:GT_MASS_SPEC`, `offline/services/gt_mass_flow.py:
  GTMassProcessor` wrapping the frozen `GTMassRecorder` — additive, recorder
  untouched): import → review (Lines/Orders/Affected/Skipped + a GT-Mass
  **file-exceptions** panel for PO-missing/template-mismatch/rescued) → per-line
  **Exclude** → confirm records to renee_orders + downloadable dump. URLs under
  `/offline/gt-mass-flow/`; the old single-page recorder + dump generator stay as
  fallbacks. **Next:** migrate Online B2B onto the same scaffold (fast-follow), so
  there's a single flow new channels plug into via a Processor + a FlowSpec.

- **2026-06-28** — **Rules → "See full template" (visual format preview).** Each
  marketplace card on Rules §3 now has a **See full template** link →
  `/b2b/rules/template/<name>/` (`MarketplaceTemplateView`). The page shows the
  **full column list + a few real sample rows** of that channel's file, with the
  columns the engine actually reads **highlighted by role** (PO / Destination /
  Item / Qty / Vendor cost / MRP / …) and the rest **dulled**. Columns + sample
  rows are a frozen fixture (`online_b2b/services/template_samples.json`, captured
  from real files); the **used/role tagging is computed live from
  `MARKETPLACE_CONFIGS`** so it never drifts. Backend:
  `engine_bridge.marketplace_templates()` / `marketplace_template()`. Web-native,
  animated successor to the desktop "download template".

- **2026-06-28** — **Tracker dates for PDF marketplaces → TAT works for them.**
  PDF channels (DMart/Avenue) carry **PO Date** + **PO Validity** in the PDF
  *header*, not as a row column, so the engine left `po_date`/`exp_date` blank →
  those orders never showed on the TAT page. Added a web-side **date backfill**:
  `Processor._source_dates_by_po()` (no-op by default; `DmartProcessor` reads the
  dates via the avenue parser) → `lines_store.set_po_dates(run_id, …)` fills the
  blanks on `order_headers` (COALESCE — never overwrites engine-set dates). DMart
  now shows in TAT (verified: PO 4502194340 = 17-Jun → 8 working days over). The
  framework is reusable for the other PDF channels (Reliance/FirstCry) — Excel
  channels already date via `po_date_col`. Existing DMart rows backfilled one-off.
  **Thumb-rule:** whenever a marketplace's date capture changes, update this
  changelog + the Rules “file format” section. See [DB_STRUCTURE.md](DB_STRUCTURE.md).

- **2026-06-27** — **Meesho Branch (Meesho-TO) integrated (web).** Added to
  `PILOT_MARKETPLACES` + a "Meesho Branch" Online chip, with a new
  `MeeshoTOProcessor` (bulk-consignment-only: Meesho exports one
  `order-line-items-<PO>[_<city>].csv` per order → `engine.process_consignments`,
  no visibility report). PO from filename; **Location from the filename city
  token** (`MS_BLR`/`MS_GGN`/`MS_KOL` → Transfer-to Code via Ship-To B2B). It's a
  **Transfer Order with no amount in the source** (Meesho's `sellingPricePerUnit`
  is a selling price, deliberately ignored), so — like Flipkart Branch — the
  **inc-GST value is computed from OUR master** (Landing × qty), never zero:
  verified on the 23-06 batch (3 POs · 40 lines · ₹275,854.20; per-PO ₹34,201 /
  ₹166,221 / ₹75,432), label `Meesho-SB`, locations resolved, dedup + web-owned
  writes (no `order_issue_lines`). **Margin = 60% is a PLACEHOLDER** (per the
  engine config) — confirm the real Meesho TO margin; the value scales with it
  (overridable per run). 23 tests green.
- **2026-06-27** — **GT Mass fully integrated into the web app (dashboard
  recorder).** New `offline/services/gt_mass_bridge.py` + page `/offline/gt-mass/`
  (preview → confirm) records GT Mass into the shared `renee_orders` (segment
  `Offline`, marketplace `GT Mass`) with the **order_lines audit** the desktop
  never wrote, so GT Mass now shows on the dashboard with Orders **and** Line
  Items. **Value is read from the raw file itself** — the GT Mass Excel carries
  `Basic Price` (unit, ex-GST), `CLP` (line ex-GST), `GST` (flat 18%) and `TOTAL`
  (line inc-GST); `order_value` = Σ `TOTAL` (inc-GST, matches the online
  channels), line `unit_price` = `Basic Price` — no margin guesswork (the per-SKU
  margin is already baked into Basic Price via the file's Retailer/Scheme/Ullage/
  DB-margin stack). qty = Order Qty (ERP auto-adds testers); tester qty captured
  in the line `remark`. **Web-owned writes** (runs+headers direct, lines via
  `lines_store`) — the engine history store is never opened, so `order_issue_lines`
  isn't resurrected; **PO-level dedup** skips SOs already recorded (incl. the
  desktop's). **EAN-only fallback**: files missing the `BC Code` column (e.g. the
  Indian-Secrets "Pack of 3" format) are rescued by resolving Item No from the
  item master via EAN; EANs absent from the master become explicit warnings, never
  a silent drop. The **frozen Tkinter standalone + the existing "Generate Dump"
  page are untouched** and remain the fallback. Verified E2E on real 27.06 files
  (record + value + lines + dedup + cleanup); 23 tests green.
- **2026-06-27** — **Item Master de-duplication (2 more redundancies removed;
  no backend break).** **(1) Dropped `item_master.swiggy_sku_code`** — it
  duplicated `channel_sku_map`. `DBMasterLoader` now builds the engine's
  `swiggy_sku` map wholly from `channel_sku_map` (EAN resolved live via item_no);
  the Item Master add/edit form routes a typed Swiggy code into `channel_sku_map`
  (new `channel_map.upsert_code`) instead of the column; status/search/admin/
  templates updated. Parity byte-identical: Swiggy resolution 272 == 272 before/
  after the `ALTER … DROP COLUMN`. **(2) Folded `item_master_manual` into
  `item_master`** via the `batch_id='manual'` source-flag pattern (same as
  `ship_to_mapping`): `replace_item_master` clears only `batch_id<>'manual'` and
  upserts the ERP rows, so hand-added items survive a full rebuild and the ERP
  source wins once it carries the item — verified on a scratch table (manual
  survives when absent from the ERP set; batch flips when present). Table dropped.
  **Design note:** kept `channel_sku_map` as channels-as-**rows** (not a column
  per marketplace) — a new code-only channel (HG/Natural/…) is an INSERT, never an
  `ALTER TABLE`, and each mapping keeps its own source/ean/updated_at. 23 tests +
  dashboard smoke green. See [DB_STRUCTURE.md](DB_STRUCTURE.md) §7.
- **2026-06-27** — **DB restructuring — web-only, de-duplicated** (no backend
  break; full end-to-end checks). Two redundant tables removed: **(1)
  `item_swiggy_map` → `channel_sku_map`** — generalised, keyed by `channel`, so
  Swiggy (272) and future code-only channels (Health & Glow, which has no EAN)
  share one map; EAN is resolved LIVE from `item_master` via `item_no` (stored
  `ean` is the fallback), `item_master.swiggy_sku_code` kept as a fast derived
  copy. Loader (`item_master_loader`), `channel_map` service, admin, router all
  repointed; parity byte-identical (272==272 swiggy_sku; 42 Swiggy lines OK after
  the drop). **(2) `order_issue_lines` DROPPED** — the desktop-only issue table
  the web double-wrote via the engine's `record_manual`/`apply_dedup`. Both lock
  and dedup are now **web-owned** (`lines_store.record_run_headers` + `web_dedup`),
  so the engine's history store is never invoked. Parity-verified byte-identical
  (runs + order_headers) before the switch; E2E re-tested on Reliance (3 POs · 65
  lines): runs/headers/lines/validation all written, `order_issue_lines` NOT
  recreated, re-preview deduped all 3 POs. Order store is now 100% web-owned.
  Kept (deliberate, not redundant): denormalised `run_ts/mode/marketplace`
  columns (avoid dashboard joins), `runs.consolidated_path/tracker_path` (empty
  legacy cols, referenced by 3 INSERT sites — dropping buys nothing), and empty
  `item_master_manual` (durable manual-add overlay). 23 tests green; `DB_STRUCTURE.md`
  refreshed. See [DB_STRUCTURE.md](DB_STRUCTURE.md).
- **2026-06-26** — **SKU Summary (SKU-wise validation pivot)**: new page
  `/b2b/sku-summary/` (sidebar ▸ Data) — every recorded line rolled up per
  **(Item No + EAN)** across all POs: qty + line-count per status (OK / Mismatch /
  Not-in-master, with a Qty↔Lines toggle), Our vs Their MRP (⚠ when vendor MRP
  varies across a SKU's POs), #POs, worst diff, marketplaces; filters (marketplace
  / date / search / issues-only) and **click-to-expand drill-down** to the SKU's
  individual PO-lines. **READ-ONLY** aggregation over `order_lines_full`
  (`order_db.sku_summary` / `sku_lines`) — NO DB changes, nothing existing touched.
  Nothing hidden (all SKUs + 3 statuses; "showing N of M" if ever capped). Also
  **appends a per-run `SKU Summary` sheet to the SO Workbook** via web post-process
  (`Processor._append_sku_sheet`) — engine + its other sheets untouched.
- **2026-06-26** — **Reliance integrated (web)**: added to `PILOT_MARKETPLACES` +
  an Online "live" chip. PDF (`pdf_parser='reliance'`, multi-file via base
  `Processor.process_multi`), `from_ean`, cost basis, **GST-dependent margin**
  (`gst_margin_discount=0.31` → keep% = 1 − 0.31×(1+GST), in-engine). Order value
  grossed up by each line's PDF GST rate. Verified on the 26-06 batch: 5 POs · 71
  lines · ₹291,091 (69 OK / 2 MISMATCH), margins 63.42% (18% GST) + 67.45% (5% GST)
  applied correctly. Also added **`DB_STRUCTURE.md`** (graphical DB map for review).
- **2026-06-26** — **Unified exceptions table**: merged `item_swiggy_deals` INTO
  `item_exceptions` — ONE table for every per-code override, with a `kind` column
  (`exception` = EAN remap / CP override / vendor-CP; `swiggy_deal` = deal SKU).
  `build_overlay_workbook` splits by `kind` to regenerate the engine's two sheets,
  so parity stays byte-identical (re-verified: exceptions / price_overrides /
  vendor_cp / 5 deals all match). Dropped the separate table + its admin model;
  `ItemException` admin now filters by kind. Cleaner single source for exceptions.
- **2026-06-26** — **Ship-To Mapping dedicated page** (`/b2b/ship-to/`, sidebar under
  Data): status KPIs (rows/parties/last-updated), **upload→preview→replace** the
  Ship-To B2B Excel (manual rows preserved), one-click **re-seed from bundled**,
  party filter + as-you-type search, and inline **CRUD** (add / edit / delete;
  UI-added rows are durable `source='manual'`). Mirrors the Item Master page.
  Backed by `mapping_store`. Verified: page + search + add/edit/delete all green.
- **2026-06-25** — **Bundled data Excels fully retired from the web (single source
  of truth = DB)**: the last two Excel-sourced overlays moved to DB via new
  `overrides_store` — `item_exceptions` (Master Exceptions: Firstcry remap / Blink
  EPISENSE / Myntra Goddess) + `item_swiggy_deals` (Swiggy deal SKUs). Parity is
  guaranteed by storing each sheet's RAW cells and regenerating a tiny workbook fed
  to the engine's OWN parsers (`load_exceptions` / `_load_swiggy_sheets`) →
  byte-identical (verified: exceptions / price_overrides / vendor_cp / 5 deals all
  match). `DBMasterLoader.load_from_db` now sources overlays from the DB (no
  `overlay_master_path`); `engine_bridge._run` is **DB-ONLY** — no bundled Excel for
  item master, Ship-To mapping, OR overrides (empty table → clear "seed it" error,
  never a silent Excel fallback). Desktop app keeps its Excel; only the web is
  retired off it. **Admin**: all DB-master tables registered (`ItemMaster`,
  `ItemSwiggyMap`, `ShipToMapping`, `ItemException`, `ItemSwiggyDeal`) + router
  updated so they read the `orders` MySQL DB.
- **2026-06-25** — **Ship-To B2B mapping → DB (addresses retired off Excel)**:
  new web-owned table `ship_to_mapping` (party / del_location / cust_no / ship_to +
  address fields + `source`); `mapping_store` parses/replaces it from the bundled
  `Ship to B2B.xlsx` and `DBMappingLoader(MappingLoader)` overrides only `load()` to
  fill the SAME `self.mappings`/`self.by_shipto` from the table — every lookup tier
  inherited untouched. `engine_bridge._run` is DB-first with Excel fallback.
  Seeded 736 rows / 25 parties; **parity verified byte-identical vs the Excel for
  all 25 parties**. **CRUD** (add/edit/delete a single mapping from the UI) with a
  durable `source='manual'` overlay that survives an Excel re-upload (like
  `item_master_manual`). Engine untouched. (Mapping UI page still pending.)
- **2026-06-25** — **Myntra reverted to Excel-primary (off the PDF parser)**: the
  operator now manually compiles the PO PDFs into one accurate `dump.xlsx`, so
  `accepted_extensions` leads with `.xlsx` (PDF kept as a selectable fallback). The
  config already drove both formats (same column names); no parser/wiring change.
  Verified on the day's `dump.xlsx`: 5 POs · 222 lines · ₹17.9L · **0 malformed
  EANs** (vs the PDF's 11 NOT_IN_MASTER from merged-EAN page wraps). Web upload hint
  updated; web preview confirmed.
- **2026-06-25** — **Bug: `TooManyFieldsSent` on large POs (Lock/Discard/Generate)**:
  a big PO's review form carries 4 decision fields per flagged line
  (`aff_key`/`aff_action`/`aff_override_cp`/`aff_remark`); a few hundred affected
  lines exceeded Django's default 1000-field cap (`DATA_UPLOAD_MAX_NUMBER_FIELDS`,
  a DoS guard) → POST rejected before the view ran. Raised the cap to 100000 in
  settings (trusted internal LAN tool). File-size guard
  (`DATA_UPLOAD_MAX_MEMORY_SIZE`) unchanged. Also made **Discard** a standalone
  csrf-only form (`#discard-form`, button targets it via HTML5 `form=`) so it no
  longer carries the big confirm-form payload — instant + safe at any PO size.
- **2026-06-25** — **Import UX + speed: AJAX upload, cached preview, real "✓
  Imported" completion**: the upload page now submits via AJAX and runs the engine
  import server-side, showing a progress overlay with real elapsed time that snaps
  to a definitive **"✓ Imported: N PO(s) · M line(s) · K to review"** before
  navigating (operators previously had no clear "done" signal). The preview result
  is **cached per token** (`<token>/preview.json`, keyed on files + EAN-fixes via
  `_preview_sig`) so the review page — and every reload — renders from cache:
  measured **8.3s → 0.002s** (engine ran once). `review()` reads the cache;
  `_cached_preview` re-runs only when the signature changes (new files / EAN-fix
  re-validate). Non-JS clients fall back to the plain redirect (review runs the same
  cached preview). Engine untouched.
- **2026-06-25** — **Myntra integrated (web)**: added to `PILOT_MARKETPLACES` + an
  Online "live" chip. SO, `from_ean`, landing basis, margin 70, PDF (`pdf_parser=
  myntra` → base `Processor.process_multi`); Goddess exception in-engine; order
  value from the dump (`amount_col`). Verified on the day's 5 POs: 222 lines ·
  ₹17.71 L (163 OK / 48 MISMATCH / 11 NOT_IN_MASTER). **Robustness fix**: a PDF
  parser can emit a malformed over-long EAN (two EANs merged on a page wrap); the
  `ean` column is VARCHAR(20) so one bad row crashed the WHOLE lock (DataError 1406).
  `Processor._lines()` now caps the EAN to 20 + warns (deduped; never silent) — the
  line is NOT_IN_MASTER for operator correction. Lock now records all 222 lines.
- **2026-06-25** — **Myntra PDF parser: header-band widened 45→60pt**: some POs
  (PO-MYNJ-RNEE240626-2/-4) failed with "line-item grid not recognised · Mapped
  columns: []". Root cause: `_map_columns` searched for column headers only 45pt
  above the first SKU; when the first line-item's NAME wraps tall, its SKU (col-0
  anchor) sits ~50pt below the header → header fell outside the band → 0 columns
  mapped. Widened the band to 60pt (safe — `_map_columns` only assigns on a header
  needle match; emails/long-digits already filtered). Verified all 5 of the day's
  Myntra POs parse with complete data (was 3/5). Engine-internal heuristic — the
  desktop app benefits too.
- **2026-06-25** — **Flipkart location→marketplace map: +4 warehouses**: added
  `che_gsh_wh_nl_01nl`/`guw_gsh_wh_nl_01nl`/`jai_sh_wh_nl_01nl` → FK Hyperlocal and
  `coi_app_wh_g_01` → FK Grocery to `flipkart_tracker.LOCATION_MARKETPLACE` (were
  falling through to "FK (review)"). 21 codes mapped; all warehouses in the
  current tracker now classify. NB: this is the sub-marketplace classification map
  — separate from the Flipkart-TO `warehouse_aliases` (Transfer-to Code resolution)
  in `config/marketplaces.py`.
- **2026-06-25** — **Bug-fix follow-up: "Download SO Workbook" 404**: the earlier
  export redirect sent ALL workbooks to a shared `MEDIA_ROOT/b2b_exports/`, but
  `review_download` reads from the **per-token** `b2b_uploads/<token>/output/` →
  404 (and a shared dir risked one upload's workbook masking another's). Fixed:
  `Processor._export()` now redirects **only when the input is OUTSIDE
  `MEDIA_ROOT`** (direct/script runs against source); web uploads stay under their
  token dir so the engine writes `output/` there and the download finds it.
  Verified both paths.
- **2026-06-25** — **Bulk decisions on the review Affected tab**: per-row
  checkboxes + select-all + a bulk bar — tick a subset → set one Action
  (Include / Override(+CP) / Exclude) → "Apply to selected"; or tick rows → type
  one Correct EAN → "Fill into selected" → Apply & re-validate. Lets the operator
  split a batch (e.g. 10 lines → 5 Exclude, 5 fix-EAN) without touching every
  field. **NOT_IN_MASTER lines are now Exclude-able** pre-lock (needed for freebie
  POs whose placeholder EAN has no real item): every affected row now emits
  `aff_key/aff_action/aff_override_cp/aff_remark` (NIM rows use hidden
  override/remark to keep `confirm`'s index-zip aligned) + the Correct-EAN input.
  Verified: 19 freebie NIM lines bulk-Excluded → recorded `NOT_IN_MASTER`+`EXCLUDE`
  + dropped from the D365 dump. Template/CSS only — no restart needed.
- **2026-06-25** — **Nykaa integrated (web)**: added to `PILOT_MARKETPLACES` + an
  Online "live" chip. Standard SO, `from_ean`, cost basis; **per-line margin by
  category** (`margin_rules`: Perfume/Fragrance 69%, Cosmetics 66% default, `hair`
  excluded from perfume, HSN cross-check) is applied **inside the engine**, so no
  per-marketplace web code. Order value from `Unit Cost × Qty` (`PO Amount` per-PO
  total). Verified on the 19-06 PO: 19 POs · 980 lines · ₹1.40 Cr (864 OK, 116
  real MISMATCH); both margins 66/69 confirmed in the DB. (A freebie PO with
  placeholder EAN `RENEE00001301` @ ₹0.01 correctly flags NOT_IN_MASTER.)
- **2026-06-25** — **Bug: web exports no longer land next to the source files**:
  the engine's `SOExporter` writes the workbook to `input_file_path.parent/output`
  (correct for the desktop app, wrong for web — it dropped `…/output/*.xlsx` next
  to the operator's picked files). `input_file_path` is used ONLY to locate that
  folder (never embedded in the workbook), so `Processor._export()` now redirects
  it to a web-owned `MEDIA_ROOT/b2b_exports/` for the export call and restores it
  after — engine untouched. Verified: workbook lands in
  `media/b2b_exports/output/`, source folder gets zero new files.
- **2026-06-25** — **Swiggy integrated (web)**: added to `PILOT_MARKETPLACES`
  + an Online "live" chip. Flat `PO_<id>.csv`; `item_resolution='from_swiggy_sku'`
  (SkuCode→EAN) resolves via the **DB master's Swiggy map** (`item_swiggy_map`,
  loaded by `DBMasterLoader`), so no per-marketplace web code. Cost basis, STRAIGHT
  80%; order value from the dump's `PoLineValueWithTax` (inc-GST). Verified on the
  24-06 sample `PO_1782289205468.csv`: 5 POs · 42 lines · ₹180,238.40 (all OK,
  SkuCode→item resolution confirmed).
- **2026-06-25** — **Purplle integrated (web)**: added to `PILOT_MARKETPLACES`
  + an Online "live" chip. Standard SO marketplace — `file_parser='purplle'`
  (tab-separated `.XLS`), `from_ean`, cost basis, margin 70%; the base
  `Processor` already routes file_parser configs through `process_multi`, so no
  per-marketplace code. Order value comes from the dump itself
  (`amount_col: Price × Qty`, inc-GST), so no zero-amount handling. Verified on
  the 24-06 sample `EXECL_ATTACHED.XLS`: 13 POs · 145 lines · ₹414,722.34
  (142 OK, 3 real price MISMATCH).
- **2026-06-24** — **Flipkart Branch (Flipkart-TO) integrated (web)**: added to
  `PILOT_MARKETPLACES` with friendly label **"Flipkart Branch"** (`pilot_choices()`
  drives the upload dropdown). New **`FlipkartTOProcessor`** routes the per-PO
  `Consignment_Details_<PO>_<date>.csv` files to the engine's
  `process_consignments` (PO from filename, optional `Consignment_Visibility_Report`
  for destination Locations) via a new `Processor.run_engine` hook; falls back to a
  single consolidated dump. **Zero-amount fix**: a TO dump carries no price, so the
  inc-GST transfer value is **computed from our master pricing** (Landing × qty =
  calculated CP inc-GST) — filled on the headers in `_headers()` (preview/summary)
  and **locked** into `order_headers.order_value` + `runs.total_value` via
  `lines_store.set_order_value()` after insert (engine untouched). Never silent: a
  warning states the value is COMPUTED, not received. Verified on the 09-06 sample
  (2 POs, 39 lines, **₹60,556.20** total = 36,927.60 + 23,628.60).
- **2026-06-24** — **Review-page Line Items tab = final ready-to-go view**: each
  line in the Line Items tab now carries a **Decision** column reflecting the
  operator's action on the affected ones — **ready** (clean/OK), **✓ Included**,
  **✎ Override @CP**, **⊘ Excluded** (struck-through, dropped from D365 but kept
  in the SO Workbook), or **● needs decision** (affected, not yet actioned). The
  review view attaches `decisions[po|item|ean]` to `res['lines']`, so the
  pre-lock Line Items tab shows the same disposition the post-lock Issues page
  does — just earlier. Lock button is also AJAX-click-ONLY now (`type=button` +
  hardened submit guard) so Enter/Tab can never "suddenly" lock.
- **2026-06-24** — **Post-lock EAN resolution on the Issues page (lock-first ops
  model)**: a NOT_IN_MASTER line can now be resolved *after* lock, on
  `/b2b/issues/`, matching how ops actually work — lock first (record the
  problem), then work the resolution. Each pending NOT_IN_MASTER row carries an
  inline **"correct EAN → Fix & resolve"** box. `lines_store.apply_issue_ean_fix(
  line_id, correct_ean)` re-resolves the item against the DB master, **recomputes
  OUR pricing with the engine's own helpers** (`MasterLoader.calc_landing_price`
  / `calc_cost_price` — engine untouched), updates the facts (`ean`/`item_no`/
  `description`) and the validation row (`our_*`, `diff`, `status` decided on the
  marketplace's `status_basis`, `exception_label='EAN remap'`), and keeps the
  **wrong EAN as `received_ean`**. Result: the line flips to OK/MISMATCH, leaves
  **Pending**, lands in **Resolved** + the **"Wrong EANs received"** escalation
  audit. View `issues_fix_ean` (JSON) at `issues/fix-ean/`. Verified end-to-end
  on a DMart line (MRP 750 × 45% ÷ 1.18 = 286.02 CP → diff 0 → OK).
- **2026-06-24** — **`order_lines` split into facts + validation (scalable model) +
  EAN-fix audit**: web-owned `order_lines` now holds **immutable order FACTS only**;
  the comparison/decision layer moved to a new **`order_line_validation`** (1:1 by
  `line_id`, FK `ON DELETE CASCADE`, only for validated lines). Reads go through a
  join **VIEW `order_lines_full`** (so query sites only swapped the table name).
  Migrated 2446 existing lines with 100% parity (0 diffs), all pages 200. New
  **`received_ean`** column on the validation table: on the review page a
  NOT_IN_MASTER line gets a **"Correct EAN"** field → re-validates against the DB
  master → `order_lines` ships the **correct** EAN while the **wrong** one is kept
  as `received_ean` (audit). A repeat wrong EAN **auto-resolves** via a map derived
  from `received_ean` (no alias table) — flagged on review as a **temporary fix**,
  and counted on the Issues page (**"Wrong EANs received N×"**) for vendor
  escalation. Engine + Tkinter untouched.
- **2026-06-24** — **DMart + Zepto + GT Select integrated; Item Master moved to DB**
  (see [[item-master-in-db]]): Item master now built from two ERP exports into
  `item_master`/`item_swiggy_map`/`item_master_manual`; engine reads it via
  `DBMasterLoader`. GT Select = D365-finalised headers+lines import (offline).
- **2026-06-24** — **Decision-driven D365: Include / Override(CP) / Exclude → Lock →
  Generate**: each affected review line now has **Include** (as-is), **Override**
  (include with an operator CP, pre-filled with vendor's *Their CP*, editable), or
  **Exclude** (drop). Flow: **🔒 Lock & Record** (push to DB + freeze decisions on the
  token) → **⬇ Generate D365** (enabled only after lock). The D365 dump reflects the
  decisions — Excludes dropped, Overrides repriced via the engine's own
  `forced_unit_price` (read as the D365 Unit Price). `engine_bridge.generate_d365()` +
  `_apply_decisions()` build a *copy* of the result (originals never mutated;
  `_run(skip_dedup=True)` so the ERP file carries the full upload). **Full SO Workbook
  stays 100% intact.** `order_lines` gains **`override_cp`** + **`decided_at`**;
  `status`/`diff`/`exception_label` remain the permanent engine snapshot (MISMATCH stays
  MISMATCH forever). Engine frozen; verified end-to-end on a real Flipkart upload.

- **2026-06-24** — **Analytics: animated tree-branch + any-date filter**: the breakdown
  is now an **interactive collapsible tree** (segment → marketplace → child) with branch
  connectors, rotating carets, **value-share bars** (animated grow-in), and per-node
  **POs · qty · value**; nodes stagger-slide in on expand. Added a **date picker** (check
  any single day) alongside the 7/30/90 range toggle — `intake_hierarchy(days, date)` now
  scopes to a specific `created_at` day. MT children roll up under the MT parent (and
  Flipkart→FK Hyperlocal/Grocery, EKA→its children). Tree JS-free (native `<details>`),
  so no load cost. NOTE: minor stray indent in `flipkart_tracker.py:56` (harmless).

- **2026-06-24** — **Management Analytics page (daily intake + segment→mkt→child)**:
  new class-based `AnalyticsView` at `/b2b/analytics/` — a **daily stacked bar chart**
  (orders received per day by `created_at`, stacked by segment, Value/Orders/Items
  toggle) + a **date-range toggle (7d/30d/90d)** + period totals + a **breakdown table
  segment → parent marketplace → child** (so Flipkart shows FK Hyperlocal/Grocery, EKA
  shows its children, MT shows its channels). Data via `order_db.daily_intake()` +
  `intake_hierarchy()`. Linked from hub (📊 Analytics) + sidebar. Chart JS
  balance-verified. Animations: chart render, breakdown-card staggered fadeUp.

- **2026-06-24** — **Orders scoped by segment (Online / Offline / All) + micro-motion**:
  the Orders page now has a **Segment selector** (All / Online B2B / Offline) that scopes
  the orders AND the Marketplace dropdown to that segment (backend already segment-aware;
  re-surfaced the UI). Online shows only online marketplaces, Offline only offline, All =
  everything; title reflects the scope. Branch "View orders" links pre-scope to their
  segment. Added subtle **micro-interactions** (card hover-lift, button/chip transitions,
  table-row + focus-ring states, staggered KPI fade-in, `prefers-reduced-motion` guard).
  NEXT refinement: roll Offline's MT children up under an "MT" parent in the dropdown.

- **2026-06-24** — **Pro pass #2: collapsible sidebar, 12-KPI hub, Departments =
  OM+GRN, MT parent/child**: sidebar now **collapses to icons** (toggle, persisted in
  localStorage) with a **smooth cubic-bezier transition** (width/padding/labels
  animate). Hub KPI strip expanded to **12 cards, 4×3 equally allocated** (added Avg
  PO Value, Received·7d, Line Items, Resolved via `order_db.hub_extra_kpis`).
  **Departments** consolidated to **Order Management** (all online + offline inside it)
  + **GRN (coming soon)**. **Offline MT aligned to parent→child**: the old "Shoppers
  Stop" page is now **Modern Trade (MT)** with a child-channel selector (Shoppers Stop /
  Health & Glow / Naturals / Apollo / Lulu) — `mt_bridge.WEB_CHANNELS` = MT children,
  testers now config-driven (`tester_qty_divisor`); GT Mass stays a separate parent.
  (SS verified end-to-end; other MT children share the generic pipeline — test before
  production.)

- **2026-06-24** — **Pro dashboard pass #1: wider layout + left sidebar nav**:
  content widened 1180→**1560px** (fills the side gutters). Added a persistent
  **left sidebar** (Hub / Online / Offline / Orders / Line Items / Issues / Process)
  with active-route highlighting — `core/base.html` got harmless `{% block body_class %}`
  + `{% block sidebar %}` hooks; new `online_b2b/base_b2b.html` injects the rail and
  all 11 b2b pages now extend it; `_sidebar.html` partial. Sidebar is scoped to
  `.b2b-app` (other departments unaffected). Header made sticky on b2b pages; content
  + footer shift right of the rail; collapses on ≤980px. (Remaining pro-pass items:
  per-KPI sparklines, date-range control, full visual polish.)

- **2026-06-23** — **Chart syntax-error fix + hub recent-activity feed + serve.bat
  clean restart**: the overview charts were dead due to a **stray `}`** (89 `{` vs 90
  `}`) — a JS syntax error that discarded the whole chart `<script>`, leaving it stuck
  on "Loading…". Rewrote the chart init clean + **balance-verified** (66/66 braces);
  it renders on DOM-ready, per-chart try/catch, visible empty/error states. Hub: KPI
  grid balanced to 4×2 and a **Recent activity** feed (`order_db.recent_orders`) added
  to fill the page elegantly. `serve.bat` now **kills any stale process on port 8000**
  before starting (stale processes were serving old code, masking every fix). NOTE on
  offline taxonomy: orders ARE stored parent/child (`marketplace`=MT/EKA/GT Mass vs
  `marketplace_label`=Naturals/SS/Airport…) — the Offline view still groups by child
  label; rolling MT sub-channels up under "MT" is the next step.

- **2026-06-23** — **'Received Today' hub card + uncached templates on prod**:
  new `order_db.today_intake()` counts orders **received today** (filtered by
  `order_headers.created_at`, not PO date) — total POs + value, split by segment.
  The hub's KPI strip now has a **Received Today** card (replacing Updated·2d) whose
  **hover popup shows the Online vs Offline distribution** (e.g. Online 54 ·
  ₹1.0Cr / Offline 10 · ₹1.72L). Also switched Django to **explicit uncached
  template loaders** (removed `APP_DIRS`) so template/HTML/chart edits go live on the
  prod server (DEBUG=0) with just a browser refresh — no restart (Python/settings
  changes still need a waitress restart). Chart init also hardened to wait for real
  container width (ResizeObserver + load fallback) so charts never paint at 0-width.

- **2026-06-23** — **Offline branch = same rich dashboard + chart render fix +
  hub KPIs**: the **overview template is now shared** by both branches via a
  `branch` context ({kind, label}) — `/b2b/offline/` (`OfflineBranchView`) renders
  the SAME KPIs + charts + marketplace-mix as `/b2b/online/`, scoped to
  `segment=Offline`, with offline-appropriate header actions (Shoppers Stop / GT
  Mass). **Chart-empty bug fixed**: charts were rendering into `.reveal`
  animating containers and measuring 0 width → blank on prod; init now **defers a
  frame** (double `requestAnimationFrame`), wraps each chart in try/catch, and
  shows an empty-state instead of a blank box. Hub KPI strip enriched to **8 cards
  with trend deltas** (POs/Value ▲%, Channels, Updated·2d, Expiring, Needs
  attention, Issue Lines). RK confirmed live in the online Process-PO flow. All
  web/template layer — core logic untouched. ruff/check/18 tests green.

- **2026-06-23** — **Central Order-Mgmt hub + 3 trees (hub counts / RK / SS line
  audit)**: `/b2b/` is now a **central hub** (class-based `CentralHubView`) — compact
  overall KPIs + two group cards **Online B2B** (`/b2b/online/`, the existing rich
  dashboard scoped to `segment=OnlineB2B`) and **Offline** (`/b2b/offline/`,
  `OfflineBranchView`). Channels differ a lot, so the two worlds stay distinct
  branches; the hub stays uncongested. Then, tree by tree: **(hub)** per-group
  POs/value/qty via `order_db.segment_kpis()`; **(online)** **RK enabled** —
  `PILOT_MARKETPLACES += 'RK'` (margin 70, generic bridge handles it); **(offline)**
  **SS line-item audit** — `mt_bridge._record_lines()` maps each resolved SS POLine →
  web-owned `order_lines` (Our MRP, Our Landing = MRP×0.6106; Vendor blank since the SS
  file has no cost), so SS now gets the same **Line Items** view as online. All web
  layer — engine + frozen Tkinter + DB schema untouched. ruff/check/18 tests green.

- **2026-06-23** — **SS unified into the online dashboard + DB; segment switch
  removed**: Shoppers Stop now follows the online **preview → confirm** flow.
  `mt_bridge.preview()` parses/validates with NO side effects (no SO number burned,
  no workbook, no DB); `mt_bridge.confirm()` assigns SO numbers (once), writes the
  workbook, and records order headers into the shared **renee_orders** DB via the
  desktop's own `record_offline_batch` (segment Offline, `marketplace_label='Shoppers
  Stop'`). New `SSPreviewView`/`SSConfirmView` + `shoppers-stop/preview|confirm/`
  routes; the SS page is now 2-step (Preview POs → Confirm & Record to DB → Download).
  **Dashboard unified** — the Segment switch (Online B2B / Offline) is removed from
  `overview.html` + `orders.html`; `order_db.SEGMENT=''` (no segment filter) so ALL
  orders (online marketplaces + SS + CSD) show together, distinguished by Marketplace.
  Verified end-to-end on a real SS file (run #55: SO/SS/06/230629, 541 qty, ₹200,870.25
  → shows on `/b2b/`). NOTE: SS records HEADERS only (line-item audit/order_lines is a
  possible follow-up). Desktop MT tool untouched.

- **2026-06-23** — **Shoppers Stop (MT Select) integrated into the web app +
  Flipkart 77% default**: (1) **Offline SS** — new headless bridge
  `offline/services/mt_bridge.py` imports the FROZEN
  `standalone_mt_select_automation.py` as a library (Tkinter is lazy-imported, so
  module load is headless) and runs the EXACT desktop Generate sequence
  (`load_all_masters → read_channel_csv_batch → assign_so_numbers →
  write_so_workbook`) → identical 6-sheet `ss_so_*.xlsx`. Masters load from the
  SAME source the desktop uses (saved `master_path` in `mt_select_config.json` →
  OneDrive dump, snapshotted read-only) and SO numbers share `mt_select_seq.json`.
  New views `ShoppersStopView`/`SSProcessView`/`SSDownloadView` + `shoppers-stop/`
  routes + `offline/shoppers_stop.html` (upload → verification table of per-PO SO
  numbers + Warnings panel + download). Linked from the Offline dashboard. Desktop
  tool untouched. (2) **Flipkart 77% default** — the upload form hardcoded the
  margin to Blink's 70; now `engine_bridge.margin_defaults()` feeds a per-marketplace
  auto-fill (Flipkart 77 / Blink 70) and the field is optional → blank falls back to
  the marketplace's configured default landing rate server-side (Tkinter-like).

- **2026-06-23** — **LAN hosting + Flipkart sub-marketplace mapping + order→line
  cascade**: (1) **Hosting** — `renee_cosmetics` now serves on the office LAN via
  **waitress** + **WhiteNoise** (env-driven `DEBUG`/`SECRET_KEY`/`ALLOWED_HOSTS`,
  `STORAGES`, `serve.bat`, `requirements.txt`, **HOSTING.md** walkthrough incl.
  dev-server vs hosted-server). (2) **Flipkart location map refreshed** to latest
  operator mapping in `engine/flipkart_tracker.py`: `bhu_men_wh_g_01` → **FK
  Grocery** (was Hyperlocal), added missing `lud_gsh_wh_nl_01nl` → FK Hyperlocal.
  (3) **Dashboard sub-marketplace** — `FlipkartProcessor` now stamps each order's
  `marketplace_label` with the per-PO tracker class (FK Hyperlocal / FK Grocery)
  instead of the blanket 'Flipkart Alpha' — overrides preview `_headers()` + a
  post-confirm UPDATE on the web-owned label column (only when the header CSV is
  present). (4) **Relation order→lines** — `lines_store.ensure_cascade_trigger()`
  adds an `AFTER DELETE` trigger on `order_headers` that auto-deletes matching
  `order_lines` (web-owned cascade; engine inserts untouched). (5) Removed today's
  mistakenly-uploaded Flipkart runs (39/43/53) — backed up to
  `backups/flipkart_removed_20260623_135511.json` first (reversible).

- **2026-06-23** — **Comparison-basis columns (Our/Their MRP·Landing·CP) + donut chart
  fix**: every line table (review Line Items + Affected tabs, Issues, Line Items explorer,
  Run detail Line Items + Affected) now shows the full **Our MRP / Their MRP · Our Landing /
  Their Landing · Our CP / Their CP** pairs with the **validation basis highlighted**
  (`.basis-on`) + a `Basis` badge. `lines_store.build_lines` tags each line `basis` =
  `CP` (cost) or `Landing`; `order_db._tag_basis` re-derives it on read from which vendor
  rate is present; reads now SELECT `vendor_mrp/vendor_landing/our_landing` everywhere
  (`issues`, `line_items`, `line_items_page`, `run_detail`). This explains the **"empty
  vendor CP" on Flipkart** — Flipkart validates on **Landing** (77% rule), so vendor CP is
  legitimately blank and the Landing pair is highlighted instead; Blink highlights the CP
  pair. Overview **donut** center now formats the hovered-slice value (`value` formatter →
  `inrShort`) instead of showing a raw number; **Discard** button restyled (`.btn-discard`).
  ruff + 18 tests green; Flipkart Landing / Blink CP basis verified on live data.

- **2026-06-23** — **Flipkart integrated (class-based bridge) + Tkinter↔Django parity**:
  refactored `engine_bridge` from functions to **classes** — `Processor` (base: load
  masters → run engine single/**multi-file** → dedup → preview/confirm) + `FlipkartProcessor`
  (always multi-file `purchase_order_*.xlsx`; optional `purchase-orders-*.csv` header →
  `result.flipkart_tracker_rows`; FK Grocery / hyperlocal ship-to is master-driven, engine
  unchanged). `processor_for()` factory; module `preview`/`confirm` delegate (views
  unchanged). `PILOT_MARKETPLACES = ['Blink','Flipkart']`. **Verified full parity** on 24
  real PO files: Django output == engine (Tkinter) — 328 lines / 24 POs / 30,060 qty /
  ₹80.39 L, per-PO qty diffs NONE; Blink single-file regression intact. ruff + 18 tests green.

- **2026-06-23** — **Line Items explorer + clickable issue cards**: new **Line Items**
  page (`/b2b/lines/`, `order_db.line_items`/`line_items_page`) — browsable view of the
  full `order_lines` audit with marketplace/status/PO filters, search, KPIs and
  load-more pagination; each PO links to its run. Linked from the Overview + Orders
  pages ("Line Items →"). Also made the Issues count cards **clickable** (the RESOLVED
  card jumps to the Resolved view; Mismatch/Not-in-master/Shown set the status filter).

- **2026-06-23** — **Per-line Action + Remark on affected (mismatch) lines + Process
  progress overlay**: on the Review screen's **Affected** tab each flagged line now has
  an **Action** dropdown (Keep / Override / Exclude) + a **Remark** field; these post
  with Confirm (single form, `formaction` buttons — no nested forms) and are stored
  against the line in **`order_lines`** (new web-owned `action`/`remark` columns, added
  additively in `lines_store.ensure_table`). `build_lines(actions=…)` maps them by
  `po|item_no|ean`; run-detail shows the recorded Action/Remark. It's a recorded
  *decision* (does NOT mutate the engine workbook). **The Issues page (`/b2b/issues/`)
  also has the editor** — each flagged line has an Action dropdown + Remark that
  **auto-save** to the DB after upload (`lines_store.update_action` + `issues_save`
  AJAX endpoint, `order_db.issues` returns `line_id`/`action`/`remark`).
  **Resolution model**: any action set ⇒ the line is RESOLVED — it leaves the Issues
  page's default **Pending** view and stops counting in the dashboard's *Needs
  attention* / *Issue Lines* KPIs (`order_db.issues` gains a `resolution`
  pending/resolved/all filter; `_kpis` counts only `action IS NULL/''`). Added a
  **bulk-set** bar (one Action+Remark applied to all shown lines via
  `update_action_bulk` + `issues_save_bulk`), a Resolution filter, and refitted the
  Issues table (wider page, horizontal-scroll container, compact grid). Also: the
  Process-PO page shows a staged **progress overlay** (percent bar + stage labels +
  live elapsed/ETA) while the request runs — purely client-side; and the web download
  now always serves the **FULL** workbook (Summary/Validation/Raw Data) with a `_full_`
  name, never the headers-only `*_d365.xlsx` sibling (`_full_workbook`/`_full_name`).
  Engine untouched throughout.
- **2026-06-23** — **Bulk import of ERP "Sales Orders" + Segment switch**: new
  `online_b2b/services/erp_import.py` parses the Business Central "Sales Orders"
  header export and imports each row into `order_headers` as **Offline** orders so
  manually-created (non-package) orders reflect on the dashboard — channel derived
  from the SO No prefix (`SO/CSD/06/…` → `CSD`), `po`=SO No, plus a new nullable
  **`external_doc`** column (added web-side, additive; the engine's inserts omit it
  → NULL) carrying the customer PO. Dedup on SO No; `mode='MANUAL'` (the runs/headers
  `mode` ENUM only allows AUTO/MANUAL). UI: **Bulk Import** page (upload → review →
  confirm, `bulk_upload`/`bulk_review`) + a **Segment switch** (Online B2B / Offline /
  All) on the overview and Orders pages — `order_db` reads are now segment-parameterized
  (`_seg`/`_kpis`/`_trends`/`_mp_sparklines`/`_where`/`overview`/`dashboard`). Engine
  untouched (web owns the import + the new column); ruff + 18 tests green.
- **2026-06-23** — **Frontend "magnificent" pass (server-rendered, no React)**: split the
  single dashboard into **Overview** (`/b2b/`) and **Orders** (`/b2b/orders/`) pages —
  the landing shows KPIs + charts + marketplace summary; the full filter/sort/paginate
  table lives on Orders. Replaced the hand-drawn SVG chart with **ApexCharts** (vendored
  to `static/online_b2b/vendor/`, offline-safe): animated gradient **area** chart with a
  Value⇄POs toggle + a **donut** of marketplace mix (data via `order_db.overview()` →
  `charts` payload + `json_script`). Added **Inter** font, **count-up** KPI numbers, and
  staggered **entrance animations**. Stays 100% Django templates + small vanilla JS (no
  React/Node) — fully maintainable. New `overview()` service, `overview.html`/`orders.html`/
  `_orders_results.html`; dead `dashboard.html`/`_dashboard_results.html` removed. Engine
  untouched; ruff + 18 tests green.
- **2026-06-22** — **Freeze backup engine + dev-tooling foundation (web only)**: the
  Tkinter `online_po_management` (engine) and `offline_po_management` are the FROZEN
  backup — reverted the `order_lines` additions I'd briefly made to engine
  `history_db.py`; the full line-item audit is now **owned entirely by the Django
  side** in `online_b2b/services/lines_store.py` (creates/writes `order_lines` via the
  shared raw connection; engine unchanged, only its existing `record_manual` is reused
  for headers). Added web-only quality tooling at the repo root: `pyproject.toml`
  (ruff + mypy + pytest config, scoped to `renee_cosmetics`/`core`/`online_b2b`,
  excludes the engine/offline), `renee_cosmetics/settings_test.py` (sqlite-only), and a
  `tests/` pytest-django suite (18 tests green: money filters, order_db helpers,
  lines_store builder). Ruff clean.
- **2026-06-22** — **Admin CRUD on the order DB (Django admin)**: added a second
  Django DB connection `orders` → MySQL `renee_orders` (creds from the engine's
  `db_config.json`, pymysql shim `version_info=(2,2,8)` for Django 6) with
  managed=False ORM models (`Run`/`OrderHeader`/`OrderLine`) + an `OrdersRouter`
  that blocks ALL migrations on that connection (Django never creates/alters/drops
  the engine-owned tables — only INSERT/UPDATE/DELETE on rows an admin edits).
  Registered in Django admin (list/search/filter/inline-edit/delete), gated to
  staff; a "DB Admin" link shows on the dashboard for staff. Dashboards still read
  via raw pymysql; the ORM path is admin-only.
- **2026-06-22** — **Web Review → Confirm flow + 2-table line audit (Blink pilot)**:
  the web upload is now **3 steps** — Upload (stash files under a token) → **Review**
  (`engine_bridge.preview`, processes in memory, **no DB write**; shows summary, all
  line items, affected/mismatch, already-uploaded dedup, warnings + a downloadable
  preview workbook) → **Confirm** (`engine_bridge.confirm`, re-processes the same files
  and PUSHES). Persistence consolidated to **2 tables**: `order_headers` (per-PO) +
  NEW **`order_lines`** (every line, with full vendor-vs-our comparison cols + status) —
  the "affected" view is just `status IN ('MISMATCH','NOT_IN_MASTER')`, so the separate
  issue-lines table is no longer needed by the web flow. (`order_lines` is owned by the
  Django side in `online_b2b/services/lines_store.py` — the engine is NOT modified; see
  the later "freeze backup engine" entry.) Django: `engine_bridge.preview/confirm`,
  views `review/confirm/discard/review_download`, `review.html` (tabbed), run-detail
  now shows Line Items; dashboard Issues/KPIs read affected from `order_lines`. Tkinter
  untouched (it still uses `order_issue_lines`).
- **2026-06-22** — **Web UI restyle (clean & corporate design system)**: extracted a
  shared stylesheet `online_b2b/static/online_b2b/b2b.css` (one indigo accent, white
  surfaces, soft borders/shadows, airy) applied across the **dashboard, Issues page,
  run-detail, and Departments hub** — inline `<style>` blocks removed from all
  templates; added a `{% block extra_css %}` hook in `core/base.html` and Indian-money
  filters (`inr_short` → ₹4.58 Cr / ₹3.21 L, `compact` → 1.2k) in `b2b_extras`.
  **Elegance pass**: emoji KPI glyphs → crisp Feather-style SVG line-icons
  (`_icons.html`); sparse bar chart → smooth SVG **area chart** with gradient fill +
  per-day hover tooltips (paths pre-computed in `_trends`); balanced 4×2 KPI grid,
  tabular-aligned numbers, no-wrap currency, header divider.
  **Interaction pass**: chart **Value⇄POs toggle** (dual paths from `_trends`);
  **per-marketplace sparklines** in the rollup (`_mp_sparklines`); **skeleton-shimmer**
  placeholders while AJAX results load; **upload page** redesigned into the design
  system (dropzone with filename echo + submit-busy state).
- **2026-06-22** — **Online-B2B dashboard v2 (KPIs, trends, filters, issues, export)**:
  `online_b2b` dashboard upgraded — KPI cards now include Expiring-≤7d, Needs-attention
  (POs with a flagged line), and ▲/▼ deltas (last-7d vs prior-7d) on POs/Value; a
  30-day order-value trend chart; per-marketplace value bars; expanded AJAX filters
  (warehouse, SO/TO, custom PO-date range) + sortable columns + Load-more pagination;
  a dedicated **Issues page** (clickable Issue/Attention cards → flagged-SKU list,
  vendor vs our MRP/CP/diff); and **Excel export** of the filtered view. All reads
  stay read-only on MySQL `renee_orders` (`order_db.py` extended:
  `_kpis`/`_deltas`/`_trends`/`issues`/`orders_page`/`orders_for_export`). Engine
  unchanged.
- **2026-06-22** — **Django web frontend (Phase 0 — Blink pilot)**: new
  `online_b2b` app in the `renee_cosmetics` Django project reuses THIS engine as
  a **library** (Option A — `online_po_management/` added to `sys.path` in
  settings; the Tkinter app + engine source are untouched). Flow: upload PO →
  `online_b2b/services/engine_bridge.run_marketplace()` replicates the desktop
  Generate path (bundled master+mapping → `MarketplaceEngine.process` → `apply_dedup`
  → `SOExporter.export` → `record_manual` + `record_issue_lines_manual`) → result
  recorded into the SAME MySQL `renee_orders` history. A **dashboard** reads that DB
  read-only (`online_b2b/services/order_db.py`, no Django migration ever runs against
  MySQL) — scoped to `segment='OnlineB2B'` (Offline channel rows excluded). KPI cards
  (Total POs, Updated last-2d, Qty, Order Value, Marketplaces, Issue Lines) + a
  per-marketplace rollup + a filterable Orders table (marketplace / period / PO search),
  plus a per-run detail page with SO-workbook download. Django's own sqlite holds
  auth/session only — no order data is duplicated there. Engine code unchanged.
- **2026-06-22** — **Deal/exception price written to Lines + always highlighted**:
  pricing exceptions (Vendor-CP, **Swiggy deal**, price override) now write their
  ecom-AGREED cost into the D365 Lines Unit Price (`forced_unit_price = cost_price_ref`)
  so the ERP uses the deal price instead of the marketplace's flat margin (e.g.
  Swiggy flat 80% in D365 was overriding negotiated deal costs). Previously only
  Vendor-CP forced the price; Swiggy deals/price-overrides fell through to the
  flat margin. ALSO: the Validation sheet now amber-highlights EVERY exception
  row + comment — regardless of OK/MISMATCH — so deals are unmistakable
  (`exception_label` set for Swiggy deals too). Engine `_process_row` +
  `validation_sheet`.
- **2026-06-20** — **Issue-line audit DB + 'Push Issues to DB'**: new
  `order_issue_lines` table (SQLite + MySQL) records ONLY flagged lines
  (status MISMATCH / NOT_IN_MASTER) — the exact Validation data (vendor vs our
  MRP/landing/CP, diff, margin, status). New GUI button "Push Issues to DB"
  (separate from the header push; enabled only when there are flags).
  **Append with a value-aware guard**: identical re-push is skipped, a revised
  MRP/CP/status is recorded as a new dated snapshot — full per-SKU history of
  problem lines. `issue_lines_from_result` / `_insert_issue_lines` /
  `record_issue_lines_manual` in history_db. Also surfaced as an **'Issue
  Lines' tab in the history export**. Safety: only CREATE-IF-NOT-EXISTS +
  INSERT on the new table; `runs`/`order_headers` untouched.
- **2026-06-19** — **Nykaa: 'Hair Perfume' is Cosmetics, not Perfume**: the
  margin-rule matcher now supports an `excludes` list (wins over `contains`).
  Nykaa's perfume rule gets `excludes: ['hair']` so a HAIR PERFUME / HAIR
  FRAGRANCE (e.g. RENEE Caramel Crush Hair Perfume) is a hair product → regular
  Cosmetics rate (66%), not 69%. Engine: `_resolve_row_margin` checks excludes
  before contains.
- **2026-06-19** — **Flipkart Tracker from header file**: on a Flipkart run the
  GUI asks "upload the header file (portal 'purchase-orders-*.csv') for the
  Tracker?" — if yes, `engine/flipkart_tracker.py` builds one row per PO
  (PO/Location/PO Date/Exp Date/PO Aging/Order Value/Order Qty) and assigns
  **Market Place by Origin Warehouse via a LOCKED location→marketplace map**
  (FK Hyperlocal default; bhi_pad_wh_nl_04nl→FK; unknown→'FK (review)', never
  blank, amber-flagged). Rendered as a 'Tracker' sheet
  (`flipkart_tracker_sheet.py`, ₹ Indian grouping). Mirrors the old
  Marketplace_Automation `flipkart.py` approach. Verified: the 8 POs in the
  19-06 header file map exactly to the operator's expected tracker.
- **2026-06-19** — **FirstCry: exact-address-first ship-to resolution**: the
  PDF's delivery ADDRESS (`__loc_address__`, the 'Address:' line after
  'Delivered To:') is now EXACT-matched against Ship-To B2B BEFORE the name/
  fuzzy tiers — so one buyer name with several ship-tos (OM ENTERPRISES →
  20493_1 vs the Pune Survey-27/1B address → 20493_2) resolves right.
  `loc_addr_col` config + `MappingLoader.lookup(fuzzy=False)` +
  `_resolve_mapping(address=...)`.
- **2026-06-19** — **Dmart/Avenue: vendor MRP + CP surfaced, status on CP**:
  the Avenue PDF already carried MRP / Basic Price / Landed Price — Dmart now
  maps `mrp_col='MRP'` (Vendor MRP) and `ref_fob_col='Basic Price'` (Vendor CP
  = Landed ÷ (1+GST)), so the Validation sheet shows all three pairs
  (Vendor/Our MRP, LR, CP). New `status_basis='cost'` finalizes OK/MISMATCH on
  the CP pair (|Vendor CP − Our CP|) instead of landing; the landing diff is
  still shown and the MRP/LR/CP cells light-yellow on any diff (per-metric
  amber tint, as in FirstCry). Engine: `_validate_against_master` picks
  `ref_diffn` for status when `status_basis='cost'`.
- **2026-06-18** — **Never drop a qty-bearing line (missing EAN/Item)**: a row
  with a PO + qty but an EMPTY EAN (or empty Item No) is no longer skipped —
  `_resolve_item_no` now returns a blank placeholder so the line is KEPT in
  Lines with a BLANK Item No and flagged (per-PO warning + NOT_IN_MASTER on
  Validation), for manual fill before import. (Previously dropped-but-logged.)
  EAN-present-but-not-in-master already appeared with the EAN as placeholder.
  Golden rule: paste lines as-is, flag for manual audit, never silently drop.
- **2026-06-18** — **Swiggy PO-status review (flag, don't drop)**: Swiggy lines
  whose `Status` ≠ CONFIRMED (EXPIRED / COMPLETED / CANCELLED / PENDING) are
  KEPT in the output ("pasted as-is") and FLAGGED — one named warning per PO —
  so the operator manually audits/removes them (a status can be wrongly given).
  Config-driven (`status_col`, `status_keep`) via `MarketplaceEngine._flag_po_status`;
  nothing is auto-dropped. GOLDEN RULE: nothing skipped silently — prefer
  flag-and-keep over auto-drop.
- **2026-06-18** — **'Rules & Exceptions' sheet**: new sheet on every output
  (``rules_sheet.py``) listing ALL marketplaces one row each — the pricing
  RULE (from config: ``margin_rules`` / ``gst_margin_discount`` / straight
  ``default_margin`` × ``compare_basis``) + that marketplace's EXCEPTIONS
  (from ``exception_registry``), e.g. "Blinkit 70% — EPISENSE 24%", "Swiggy
  80% — 5 deal SKUs", "Flipkart 77%". Current marketplace highlighted. This is
  the SINGLE exceptions view — the earlier detailed per-EAN ``exceptions_sheet``
  is no longer written (one consolidated 'Rules & Exceptions' sheet only).
- **2026-06-18** — **Myntra dual landing+cost validation** (`also_check_cost`):
  a Myntra row is now OK only when BOTH the landing pair (vendor Landing vs
  MRP×m%) AND the cost pair (vendor CP = 'List price' ref vs our CP =
  MRP×m%÷GST) agree — either failing → MISMATCH (previously only landing drove
  status). Engine downgrades an otherwise-OK row when |ref_fob − cost_price_ref|
  > threshold and logs a 'Cost mismatch' warning; deal/vendor-CP exception rows
  (e.g. Goddess) are exempt via a `_pricing_exception` flag. Opt-in per config,
  so other marketplaces are unchanged.
- **2026-06-18** — **Auto D365 "Edit in Excel" package**: every SO run now also
  writes a sibling ``<output>_d365.xlsx`` populated into the bound connector
  template (``templates/d365_so_template.xlsx`` = the operator's
  'ABHISHEK WAGH - SO' package), so Headers/Lines no longer need hand-copying.
  New ``exporter/d365_package.py`` does binding-preserving ZIP surgery: keeps
  rows 1 (PackageCode/TableID) & 3 (bound headers), replaces data rows 4+ as
  inline strings, stretches the table ``ref`` — and copies xmlMaps /
  connections / tableSingleCells / styles / sharedStrings through BYTE-FOR-BYTE
  (verified identical), so the XML-map → D365 field binding survives. Header
  cols A–R (18) / Line cols A–H (8) mirror ``headers_sheet`` / ``lines_sheet``;
  data-cell style auto-detected from the template. Wired in
  ``SOExporter.export`` (best-effort; never breaks the main output). **TO
  (Transfer Order, Flipkart-TO) also wired** — ``templates/d365_to_template.xlsx``,
  Transfer Header A–N (14) / Transfer Line A–G (7: DocNo|LineNo|ItemNo|Qty|UoM|
  Bin|Transfer Price=``calc_price``); verified ref/binding preserved. The GUI
  **'Export D365 Package' button no longer prompts for a template** — it fills
  the bundled SO/TO template via the same surgery (no file dialog), superseding
  the old ``D365Exporter`` fragile sharedStrings-rebuild path (and its wrong
  12/9-col TO layout).
- **2026-06-17** — **Exceptions sheet = full cross-marketplace registry**: every
  output now lists ALL `Master Exceptions.xlsx` rows (all marketplaces),
  HIGHLIGHTING (green/bold) the ones that apply to the marketplace being
  processed (own + blank=ALL rows) — e.g. EPISENSE highlighted in Blinkit's
  file, Goddess in Myntra's. `MasterLoader.exception_registry` (full row list +
  derived `effect`/`kinds`) is stamped onto `result.exception_registry`
  (process / process_multi / consignments); `exceptions_sheet` rewritten with
  Marketplace/Type/Source/Maps To/MRP/Margin/Vendor CP/Effect/Note + an
  **Applied (this run)** column cross-referencing `exceptions_applied`. The
  registry also folds in the **Swiggy deal SKUs** (master 'Swiggy Deal SKUs'
  sheet, type `swiggy_deal`, Swiggy-scoped) so they appear too. NB: Reliance's
  'Reliance Deal SKUs' sheet is the GST-margin RULE table (0.31 →
  69/67.45/63.42%), not per-SKU exceptions — applied via `gst_margin_discount`,
  shown in the pricing-rule banner, so it's intentionally not listed here.
- **2026-06-17** — **Myntra cross-page wrap fix** (`myntra_pdf_parser.py`):
  rows are now reconstructed on one document-global y-axis (`_rows_all_pages`)
  so a line item whose cells wrap onto the next page's top rejoin — fixes the
  last-item-per-page EAN truncation (PO MYNJ-RNEE160626-3 body-mist
  `89061216`+`48782` → `8906121648782`). Verified all 4 POs: Σqty == footer
  qty, zero short EANs. Also **`process_multi` now aggregates
  `exceptions_applied`** across files, so the Exceptions sheet is populated for
  MULTI-FILE marketplaces (Myntra Goddess vendor-CP, Flipkart) — previously
  recorded per-file and lost at merge.
- **2026-06-17** — **Flipkart dump generation moved in-app**: the new vendor
  portal emits one `purchase_order_<PO>.xlsx` per PO (two-row hierarchical
  header). `flipkart_dump_parser.py` (`file_parser='flipkart'`) reads each →
  dump columns in memory; setting `file_parser` auto-enables MULTI-FILE, so
  the operator drops ALL the day's PO files → one SO batch (retires the
  standalone generator + the FL_DUMP intermediate; no xlwings). PO from the
  FILENAME; carries vendor MRP (`mrp_col='MRP'`, `Supplier MRP`); `total_amount`
  verbatim (`Total Amount`). New `loc_match='address'` →
  `MappingLoader.lookup_by_address` resolves ship-to by pincode + survey-no/
  village body overlap (the portal drops the `Flipkart India Pvt. Ltd.,`
  prefix the Del Locations carry; ambiguous pincodes disambiguated by body).
  Verified: 25 files → 344 rows, all ship-to resolved, 335 OK / 9 genuine
  MRP-mismatch flags.
- **2026-06-17** — **Exception marketplace-scoping fix**: the Swiggy deal-SKU
  CP override was leaking into other marketplaces (Blinkit Villain combo
  8906121643282 wrongly clamped to Swiggy's 355.42 deal CP → false MISMATCH).
  Gated on `marketplace=='swiggy'`. Also: Myntra Goddess `Use Vendor CP`
  exception (Myntra-only) accepts vendor CP + writes it into Lines Unit Price;
  Validation/Lines amber-highlight each exception row; Myntra added to tracker
  `_SUPPORTED`.
- **2026-06-16** — **Swiggy** marketplace: flat CSV, `item_resolution=
  'from_swiggy_sku'` (dump SkuCode → master `Swiggy` sheet → EAN → Item No),
  80% straight, order value = `PoLineValueWithTax` (inc-GST). **Swiggy deal
  SKUs**: master `Swiggy Deal SKUs` sheet overrides the expected CP to the
  sheet's `Cost after GST` (validation OK + logged). `MasterLoader` gained
  `swiggy_sku`/`swiggy_deals`/`_load_swiggy_sheets`/`resolve_swiggy_sku`.
- **2026-06-16** — **Purplle** marketplace: SAP-style tab-separated `.XLS`
  via `file_parser='purplle'` (`purplle_parser.py`, cleans the
  `…590'`-style zero-padded/quoted EAN), 70% straight, `Price`=post-GST cost.
- **2026-06-16** — **Pricing-rule banner**: Summary shows the real rule
  (`pricing_rule_str`) incl. deal-SKU count; Validation `Our Landing`
  header/info use a rule-aware label (`_margin_label`) — no more misleading
  flat `66%` for Nykaa/Reliance. Front-end only; no pricing logic changed.
- **2026-06-16** — **Nykaa tracker** added to `_SUPPORTED`; Order Value uses
  `po_total_col='PO Amount'` (per-PO grand total = portal exactly, GST-incl)
  + `po_date_col`/`exp_date_col` honoured by `_build_po_date_lookup`.
- **2026-06-16** — **Meesho-TO** tracker + ship-to from filename: city token
  (`…-blr.csv` → `MS_BLR`) resolved via new `MappingLoader.by_shipto` reverse
  index; PO Date = today, Exp Date = +7; order value = landing × qty.
- **2026-06-15** — **Central exceptions overlay** (`Master Exceptions.xlsx`):
  item-alias remaps + price overrides, auto-loaded beside the master,
  applied in `lookup()`/`_validate_against_master`, logged on the new
  **Exceptions** sheet (`exceptions_sheet.py`). EPISENSE (Blink) + FirstCry
  live in it. `result.exceptions_applied` added to `ProcessingResult`.
- **2026-06-15** — **Flipkart-TO tracker** completed: per-PO Order Value =
  landing × qty (= D365 Total Incl. GST, verified), PO/Exp dates threaded
  from the Consignment Visibility Report, vendor (portal) total logged for
  reference; `Bigbasket`/`Blink` added to tracker `_SUPPORTED`.
- **2026-06-15** — **Big Basket** marketplace: per-PO `<PO>.xlsx` with a
  multi-row preamble via `file_parser='bigbasket'` (`bigbasket_parser.py`);
  70% straight, inc-GST order value. **Validation Qty column** added (spot
  low-qty lines); Blink got `mrp_col` (Vendor MRP now populated).
- **2026-06-13** — **Dedup-skip**: already-uploaded POs are now **removed**
  from Headers/Lines (not re-sent to D365) and listed on a new **"Skipped
  POs"** output sheet; the DB stores **only new POs** (removed
  `is_duplicate` + `first_seen_ts` columns via migration, rows kept). New
  `apply_dedup` + `DEDUP_SKIP_ENABLED` + `existing_pos()`; `record` no
  longer returns duplicates.
- **2026-06-13** — Added a **`Segment`** column (`OnlineB2B`) to the tracker
  (col 1) and `order_headers` (migration + back-fill of existing rows) for
  future offline/GT orders in the same DB/tracker. Const in
  bare `Date` to the candidates) — all marketplaces now emit real dates.
  `config.constants.ORDER_SEGMENT`. Also fixed Flipkart PO date (added
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
