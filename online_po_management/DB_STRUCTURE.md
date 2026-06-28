# 🗄️ Database Structure — `renee_orders` (MySQL)

> **Single source of truth for the web app.** Every marketplace/channel PO the
> web processes, and all the master data it validates against, lives here. The
> bundled Excels have been **retired** from the web — the engine reads this DB.
> Connection: `db_config.json` (`backend=mysql`) → `renee_orders @ 127.0.0.1:3306`.
>
> Last reviewed: **2026-06-27** · 11 web marketplaces live + offline channels.
> **Web-only DB** — the desktop Tkinter store has been retired; the web app now
> owns every order write (no engine history store, no `order_issue_lines`).
> **De-duplicated master** — per-channel SKU codes live ONLY in `channel_sku_map`
> (no `item_master.swiggy_sku_code` column); hand-added items live IN `item_master`
> (`batch_id='manual'`), so `item_master_manual` is gone.

---

## 1. Big picture — two domains

```
                        ┌──────────────────────────────────────────────┐
                        │              renee_orders  (MySQL)            │
                        └──────────────────────────────────────────────┘
                                   │                        │
              ┌────────────────────┘                        └───────────────────┐
              ▼                                                                  ▼
   ╔══════════════════════╗                                       ╔══════════════════════════╗
   ║   MASTER DATA        ║   read on every PO run, validate       ║         ORDERS           ║
   ║   (reference)        ║   prices / resolve locations / EANs    ║   (what we processed)    ║
   ╚══════════════════════╝                                       ╚══════════════════════════╝
   • item_master          ────────────── EAN / Item No lookup ───►  • runs
     (manual rows: batch_id='manual')                              • order_headers
   • channel_sku_map      ──── code-only channels (Swiggy/HG) ──►   • order_lines (FACTS)
   • ship_to_mapping      ───────── location → ship-to code  ───►   • order_line_validation
   • item_exceptions      ───────── per-SKU price overrides  ───►   • order_lines_full (VIEW)
```

---

## 2. MASTER DATA  (the retired-Excel tables)

```
┌─ item_master ───────────────────────────────────────────── 1520 rows ─┐
│  PK item_no                                                            │
│  ean · description · gst_code · hsn · mrp · mrp_start · mrp_end        │
│  base_uom · brand · category · batch_id · updated_at                   │
│  ◦ Built from 2 ERP exports (Items + Item M.R.P.), effective-MRP-today │
│  ◦ Engine reads it via DBMasterLoader (no Excel)                       │
│  ◦ Hand-added items live HERE, flagged batch_id='manual' — durable: a   │
│    full ERP rebuild clears only batch_id<>'manual', so they survive     │
│    until the ERP export carries them (then the source row wins).        │
│  ◦ No swiggy_sku_code column — per-channel codes are in channel_sku_map.│
└────────────────────────────────────────────────────────────────────────┘
        ▲ rebuilt on upload (manual rows preserved)
        │                              ┌─ channel_sku_map ──── 272 rows ─────────┐
        │                              │  PK id  · INDEX channel, sku_code        │
        │   the per-channel codes      │  channel · sku_code · ean · item_no      │
        └── source of truth ──────────►│  source · updated_at                     │
                                       │  ◦ per-channel SKU-code → item/EAN map   │
                                       │  ◦ channel='Swiggy' (272). Health&Glow   │
                                       │    + future code-only channels slot in   │
                                       │    here as ROWS (HG has no EAN). Adding   │
                                       │    a channel = INSERT, never ALTER TABLE. │
                                       │  ◦ EAN resolved live via item_no join,   │
                                       │    stored ean is the fallback.           │
                                       │  ◦ Generalises the old item_swiggy_map.  │
                                       └──────────────────────────────────────────┘

┌─ ship_to_mapping ─────────────────────────────────────────── 737 rows ─┐
│  PK id   · INDEX party, ship_to                                        │
│  party · del_location · cust_no · ship_to                              │
│  name · address · address2 · postcode · city                          │
│  source = 'excel' (bulk upload, replaced) | 'manual' (UI add, durable) │
│  ◦ 25 parties (online marketplaces + offline channels)                 │
│  ◦ location → ERP Cust No + Ship-to code · DBMappingLoader (no Excel)  │
│  ◦ UI: /b2b/ship-to/  ·  Admin: ShipToMapping                          │
└────────────────────────────────────────────────────────────────────────┘

┌─ item_exceptions ─── ALL per-code overrides, one table ───────── 8 rows ─┐
│  PK id   · INDEX kind, source_code                                       │
│  kind = 'exception'   → EAN remap / CP override / use-vendor-CP          │
│         (source_code · maps_to · override_mrp · override_margin ·        │
│          use_vendor_cp · marketplace · note)                            │
│  kind = 'swiggy_deal' → Swiggy negotiated deal SKU                       │
│         (item_id · source_code=EAN · note=name · override_mrp=MRP ·      │
│          correct_gst · cost_with_gst · cost_after_gst)                  │
│  source = 'excel' | 'manual'                                            │
│  ◦ On load, split by kind → regenerates the engine's 2 sheets → its own  │
│    parsers interpret it (byte-identical parity, engine untouched)        │
│  ◦ Admin: ItemException (filter by kind)                                │
└──────────────────────────────────────────────────────────────────────────┘
```

---

## 3. ORDERS  (the processing record — now 100% web-owned)

```
┌─ runs ─────────────────────────── 62 rows ─┐
│  PK run_id                                  │   one row per lock/upload batch
│  mode · source · marketplaces               │
│  total_pos · total_items · total_qty        │
│  total_value · run_ts                       │   (consolidated_path/tracker_path:
└───────────────┬─────────────────────────────┘    legacy desktop cols, web writes '')
                │ 1 : N   (run_id)
                ▼
┌─ order_headers ───────────────── 530 rows ─┐
│  PK order_id  · run_id                      │   one row per PO
│  marketplace · marketplace_label · segment  │
│  po · location · warehouse · order_type     │
│  po_date · exp_date · items · qty           │
│  order_value · external_doc · output_file   │
└───────────────┬─────────────────────────────┘
                │ logical  (run_id + po)
                ▼
┌─ order_lines ───── FACTS ─────── 4000 rows ─┐        ┌─ order_line_validation ── 4000 rows ─┐
│  PK line_id · run_id                        │ 1 : 1  │  PK line_id  (FK → order_lines,      │
│  marketplace · po · location                │◄──────►│       ON DELETE CASCADE)             │
│  item_no · ean · description · qty           │        │  our/vendor: mrp · landing · cp      │
│  gst_code · unit_price · output_file         │        │  diff · margin_pct · status          │
│  ◦ immutable order facts                     │        │  exception_label · received_ean      │
└──────────────────────────────────────────────┘        │  action · override_cp · remark       │
                                                         │  ◦ comparison + operator decision    │
                ┌────────────────────────────┐           └──────────────────────────────────────┘
                │  order_lines_full  (VIEW)  │  = order_lines ⨝ order_line_validation
                │  4000 rows                 │    (COALESCE status→'OK')  ◄── ALL reads go here
                └────────────────────────────┘
```

**Why the order_lines / order_line_validation split?** Scalability + a clean
audit trail: `order_lines` holds immutable facts; the comparison/decision layer
(and the wrong-EAN `received_ean`) lives 1:1 in `order_line_validation`. All reads
use the `order_lines_full` VIEW so query sites only see one logical row.

**Where did `order_issue_lines` go?** It was the desktop Tkinter app's issue
table. The web never read it but used to double-write it via the engine's
`record_manual` / `apply_dedup`. Those calls are now replaced by web-owned
replicas (`lines_store.record_run_headers` + `web_dedup`), so the engine's
history store is never invoked and the table was **dropped** (2026-06-27).
Lock + dedup were parity-verified byte-identical before the switch.

---

## 4. How the data flows on a PO run

```
  Upload PO file(s)
        │
        ▼
  ENGINE (frozen) ── reads ──►  item_master · ship_to_mapping · item_exceptions   (DB only)
        │                       channel_sku_map (Swiggy/HG SkuCode→item)
        │  validate price (GST-aware) · resolve location · apply overrides
        ▼
  Review page  ── operator decides (Include / Override / Exclude / fix EAN)
        │
        ▼  Lock & Record   (web-owned writers — no engine history store)
  web_dedup  ── drop POs already in order_headers (replica of engine apply_dedup)
        │
        ▼
  WRITE ──►  runs (1)  →  order_headers (N)  →  order_lines (N)  +  order_line_validation (1:1)
```

---

## 5. Integrated marketplaces / channels

```
ONLINE B2B  (online_po_processor engine, DB-only)
  ✓ Blink   ✓ Flipkart   ✓ RK   ✓ DMart   ✓ Zepto   ✓ Flipkart Branch (TO)
  ✓ Purplle ✓ Swiggy     ✓ Nykaa ✓ Myntra  ✓ Reliance  ✓ Meesho Branch (TO)
                                                                    · BlinkMP (soon)

OFFLINE  (recorded to the SAME order tables)
  ✓ MT (Modern Trade — SS via headless bridge)   ✓ GT Select   ✓ EKA   ✓ CSD
  ✓ GT Mass — web recorder (offline/services/gt_mass_bridge.py): preview→confirm,
    value from the file's own TOTAL (inc-GST), order_lines audit, PO-dedup,
    EAN-only fallback; the Tkinter dump generator stays as the fallback.
```

Adding an online marketplace = its key in `PILOT_MARKETPLACES`
(`online_b2b/services/engine_bridge.py`) + a chip in `ONLINE_CHANNELS`
(`views.py`). Engine config already exists in `config/marketplaces.py`.

---

## 6. Ownership & rules

| Table                   | Owner | Web writes | Web reads | Notes                                  |
|-------------------------|-------|-----------|-----------|----------------------------------------|
| item_master            | web   | rebuild+CRUD| ✓ (run) | from ERP exports; manual rows batch_id='manual' (durable) |
| channel_sku_map        | web   | seed/CRUD | ✓         | per-channel code→item (Swiggy/HG/+), channels as rows |
| ship_to_mapping        | web   | upload/CRUD| ✓        | 25 parties · source flag                |
| item_exceptions        | web   | seed/CRUD | ✓ (run)   | unified overrides (kind)                |
| runs                   | web   | ✓ (lock)  | ✓         | one row per lock batch                   |
| order_headers          | web   | ✓ (lock)  | ✓         | one row per PO                           |
| order_lines            | web   | ✓ (lock)  | ✓         | facts                                   |
| order_line_validation  | web   | ✓ (lock)  | ✓         | validation + decisions                  |
| order_lines_full (VIEW)| web   | —         | ✓         | the read surface                        |

**Migrations:** Django never DDLs the orders DB (`OrdersRouter`). Tables are
created/owned by the web services (`managed = False` models exist only for admin
browsing). Admin: `/admin/online_b2b/…`.

---

## 7. Known redundancy (for cleanup review)

1. ~~**`order_issue_lines`**~~ — **RESOLVED 2026-06-27.** Web no longer writes via
   the engine; replaced by web-owned `record_run_headers` + `web_dedup`. Table
   dropped.
2. ~~**`item_swiggy_map`**~~ — **RESOLVED.** Generalised into `channel_sku_map`
   (keyed by `channel`, channels stored as ROWS not columns) so Health & Glow and
   other code-only channels share one table without a schema change. The live EAN
   is resolved from `item_master` via `item_no`.
3. ~~**`item_master.swiggy_sku_code`**~~ — **RESOLVED 2026-06-27.** Dropped; it
   duplicated `channel_sku_map`. `DBMasterLoader` builds the engine's
   `swiggy_sku` map wholly from `channel_sku_map`; the Item Master add/edit form
   routes a typed Swiggy code there. Parity byte-identical (272 before == after).
4. ~~**`item_master_manual`**~~ — **RESOLVED 2026-06-27.** Folded into
   `item_master` using the `batch_id='manual'` flag (same source-flag pattern as
   `ship_to_mapping`). Rebuild clears only `batch_id<>'manual'` and upserts, so
   manual rows survive and the ERP source wins once it carries the item. Table
   dropped.
5. **Denormalised columns** — `order_lines` repeats `marketplace/po/location/
   run_ts` from `order_headers`; `order_headers` repeats `run_ts/mode` from
   `runs`. **Deliberate** (avoid joins on dashboards) — kept on purpose.
6. **`runs.consolidated_path` / `runs.tracker_path`** — legacy desktop workbook
   paths; the web writes `''`. Harmless nullable cols, referenced by 3 INSERT
   sites; left in place (dropping them buys nothing and risks the lock path).
