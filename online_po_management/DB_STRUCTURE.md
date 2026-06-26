# 🗄️ Database Structure — `renee_orders` (MySQL)

> **Single source of truth for the web app.** Every marketplace/channel PO the
> web processes, and all the master data it validates against, lives here. The
> bundled Excels have been **retired** from the web — the engine reads this DB.
> Connection: `db_config.json` (`backend=mysql`) → `renee_orders @ 127.0.0.1:3306`.
>
> Last reviewed: **2026-06-26** · 11 web marketplaces live + offline channels.

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
   • item_master_manual                                             • order_headers
   • item_swiggy_map                                                • order_lines (FACTS)
   • ship_to_mapping      ───────── location → ship-to code  ───►   • order_line_validation
   • item_exceptions      ───────── per-SKU price overrides  ───►   • order_lines_full (VIEW)
                                                                    • order_issue_lines (legacy)
```

---

## 2. MASTER DATA  (the retired-Excel tables)

```
┌─ item_master ───────────────────────────────────────────── 1520 rows ─┐
│  PK item_no                                                            │
│  ean · description · gst_code · hsn · mrp · mrp_start · mrp_end        │
│  swiggy_sku_code* · base_uom · brand · category · batch_id · updated_at│
│  ◦ Built from 2 ERP exports (Items + Item M.R.P.), effective-MRP-today │
│  ◦ Engine reads it via DBMasterLoader (no Excel)                       │
└────────────────────────────────────────────────────────────────────────┘
        ▲ rebuilt fully on upload          ▲ *denormalised copy of the map
        │                                  │
┌─ item_master_manual ─── 0 rows ─┐   ┌─ item_swiggy_map ──── 272 rows ─┐
│  PK item_no                     │   │  PK item_no                     │
│  hand-added items, re-applied   │   │  swiggy_sku_code                │
│  after every rebuild (durable)  │   │  (durable SkuCode→item source)  │
└─────────────────────────────────┘   └─────────────────────────────────┘

┌─ ship_to_mapping ─────────────────────────────────────────── 736 rows ─┐
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

## 3. ORDERS  (the processing record)

```
┌─ runs ─────────────────────────── 57 rows ─┐
│  PK run_id                                  │   one row per lock/upload batch
│  mode · source · marketplaces               │
│  total_pos · total_items · total_qty        │
│  total_value · consolidated_path · run_ts   │
└───────────────┬─────────────────────────────┘
                │ 1 : N   (run_id)
                ▼
┌─ order_headers ───────────────── 505 rows ─┐
│  PK order_id  · run_id                      │   one row per PO
│  marketplace · marketplace_label · segment  │
│  po · location · warehouse · order_type     │
│  po_date · exp_date · items · qty           │
│  order_value · external_doc · output_file   │
└───────────────┬─────────────────────────────┘
                │ logical  (run_id + po)
                ▼
┌─ order_lines ───── FACTS ─────── 3368 rows ─┐        ┌─ order_line_validation ── 3368 rows ─┐
│  PK line_id · run_id                        │ 1 : 1  │  PK line_id  (FK → order_lines,      │
│  marketplace · po · location                │◄──────►│       ON DELETE CASCADE)             │
│  item_no · ean · description · qty           │        │  our/vendor: mrp · landing · cp      │
│  gst_code · unit_price · output_file         │        │  diff · margin_pct · status          │
│  ◦ immutable order facts                     │        │  exception_label · received_ean      │
└──────────────────────────────────────────────┘        │  action · override_cp · remark       │
                                                         │  ◦ comparison + operator decision    │
                ┌────────────────────────────┐           └──────────────────────────────────────┘
                │  order_lines_full  (VIEW)  │  = order_lines ⨝ order_line_validation
                │  3368 rows                 │    (COALESCE status→'OK')  ◄── ALL reads go here
                └────────────────────────────┘

┌─ order_issue_lines ── ⚠ LEGACY ── 16 rows ─┐
│  written by the engine's record_manual,    │   The WEB never reads it (its Issues page reads
│  duplicates line + price + status cols      │   order_line_validation + the view). Only the
│  ◦ used by the DESKTOP Tkinter app          │   desktop app uses it. → redundancy candidate.
└─────────────────────────────────────────────┘
```

**Why the order_lines / order_line_validation split?** Scalability + a clean
audit trail: `order_lines` holds immutable facts; the comparison/decision layer
(and the wrong-EAN `received_ean`) lives 1:1 in `order_line_validation`. All reads
use the `order_lines_full` VIEW so query sites only see one logical row.

---

## 4. How the data flows on a PO run

```
  Upload PO file(s)
        │
        ▼
  ENGINE (frozen) ── reads ──►  item_master · ship_to_mapping · item_exceptions   (DB only)
        │                       item_swiggy_map (Swiggy SkuCode→EAN)
        │  validate price (GST-aware) · resolve location · apply overrides
        ▼
  Review page  ── operator decides (Include / Override / Exclude / fix EAN)
        │
        ▼  Lock & Record
  WRITE ──►  runs (1)  →  order_headers (N)  →  order_lines (N)  +  order_line_validation (1:1)
        └──►  [engine also writes order_issue_lines — legacy, web ignores it]
```

---

## 5. Integrated marketplaces / channels

```
ONLINE B2B  (online_po_processor engine, DB-only)
  ✓ Blink   ✓ Flipkart   ✓ RK   ✓ DMart   ✓ Zepto   ✓ Flipkart Branch (TO)
  ✓ Purplle ✓ Swiggy     ✓ Nykaa ✓ Myntra  ✓ Reliance              · BlinkMP (soon)

OFFLINE  (recorded to the SAME order tables)
  ✓ MT (Modern Trade — SS via headless bridge)   ✓ GT Mass   ✓ GT Select
```

Adding an online marketplace = its key in `PILOT_MARKETPLACES`
(`online_b2b/services/engine_bridge.py`) + a chip in `ONLINE_CHANNELS`
(`views.py`). Engine config already exists in `config/marketplaces.py`.

---

## 6. Ownership & rules

| Table                   | Owner | Web writes | Web reads | Notes                                  |
|-------------------------|-------|-----------|-----------|----------------------------------------|
| item_master            | web   | rebuild   | ✓ (run)   | from ERP exports                        |
| item_master_manual     | web   | CRUD      | ✓         | durable manual items                    |
| item_swiggy_map        | web   | seed      | ✓         | durable SkuCode source                  |
| ship_to_mapping        | web   | upload/CRUD| ✓        | 25 parties · source flag                |
| item_exceptions        | web   | seed/CRUD | ✓ (run)   | unified overrides (kind)                |
| runs / order_headers   | engine| ✓ (lock)  | ✓         | history store                           |
| order_lines            | web   | ✓ (lock)  | ✓         | facts                                   |
| order_line_validation  | web   | ✓ (lock)  | ✓         | validation + decisions                  |
| order_lines_full (VIEW)| web   | —         | ✓         | the read surface                        |
| order_issue_lines      | engine| ✓ (engine)| ✗         | **legacy** — web ignores; desktop uses  |

**Migrations:** Django never DDLs the orders DB (`OrdersRouter`). Tables are
created/owned by the engine + the web services (`managed = False` models exist
only for admin browsing). Admin: `/admin/online_b2b/…`.

---

## 7. Known redundancy (for cleanup review)

1. **`order_issue_lines`** — the web double-writes it (via the engine) but never
   reads it; only the desktop app does. *Removable from the web path* if we stop
   populating it on web locks (won't affect the web; check desktop expectations).
2. **`item_master.swiggy_sku_code` ⟷ `item_swiggy_map`** — same code in two
   places (intentional: the map is the durable source, the column a fast copy).
3. **Denormalised columns** — `order_lines` repeats `marketplace/po/location/
   run_ts` from `order_headers`; `order_headers` repeats `run_ts/mode` from
   `runs`; file paths in 3 tables. Deliberate (avoid joins on dashboards).
4. **`item_master_manual`** (0 rows) — not redundant, just empty (durability).
