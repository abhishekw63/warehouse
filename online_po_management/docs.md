# Online PO Management — Architecture & Flow

> **Living document.** This is the map of how the code actually works.
> **Thumb rule: every new development updates this file** — add/adjust the
> relevant flow, the component table, and a line in the Changelog at the
> bottom. Keep the diagrams honest; if a diagram and the code disagree, the
> code is right and the diagram must be fixed.

Last updated: 2026-06-16 · Covers: Manual + Auto modes, history DB (SQLite→MySQL), DB-sourced tracker, central exceptions overlay, pricing-rule banner, 18 marketplaces.

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
