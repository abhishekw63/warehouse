# RENÉE Order Management — Engineering Handoff & Optimization Brief

> Paste this whole file into any AI coding agent as context. It contains: the system,
> the hard constraints, what was just changed, the FULL audit findings (done + pending),
> and the prioritized backlog. Obey the constraints in §2 exactly — this is a money-path app.

---

## 0. How to use this
You are picking up work on a live Django app that generates financial upload files.
Before changing anything: read §2 (constraints) and §7 (how to verify). Prefer the
smallest safe change; prove read-path rewrites are output-identical; never blind-ship
anything you can't verify.

## 1. System overview
- **App:** "RENÉE Order Management" — B2B/warehouse PO automation for RENÉE Cosmetics (Django 6, Python 3.13).
- **What it does:** ingests marketplace/vendor POs (PDF/Excel) → produces Microsoft **D365 sales-order upload workbooks**; tracks orders (Consolidated Tracker); inventory fill-rate cockpit; availability checker; analytics; record-verify/reconcile.
- **Hosting:** **Render free tier** (0.1 CPU, 512MB, spins down when idle; 1 gunicorn worker / 8 threads / 300s timeout — see `render.yaml`, `Procfile`) + remote **TiDB** (MySQL-compatible) in **Singapore** (ap-southeast-1). **Auto-deploys from `origin/main`.**
- **Data layer:** raw **pymysql** (not the ORM) in `online_b2b/services/order_db.py`: `_conn()`/`_conn_tx()` context managers yield `(cursor, dialect)`; `dialect['ph']='%s'`, `dialect['orders']='order_headers'`, `dialect['kind']='mysql'|'sqlite'`. `_CountingCursor` counts queries/time per request. In-process query cache `_stable(key, builder, ttl=60s)` + `_stable_bust(prefix)`. A per-thread warm pymysql pool + `CONN_MAX_AGE=300` handle connection reuse.
- **Engine boundary (frozen, don't edit casually):** `online_po_management/online_po_processor/` (online engine) and `offline_po_management/channels/mt_select/standalone_mt_select_automation.py` (offline MT engine, aliased `eng`). Web wraps them via `*_bridge.py`.
- **MT (Modern Trade) channels:** registered in `offline/services/mt_bridge.py` — each is a `ChannelConfig` + optional `_normalize_<x>_excel()` + a dispatch branch + membership in `WEB_CHANNELS`. Masters: `item_master` (DB) for online; the offline MT path reads `MT_Masters.xlsx` (`items_by_gtin`) materialized from the `offline_master_file` DB table. Ship-to via `ship_to_mapping` (keyed `(party, del_location)`).
- **Observability:** staff-only `/dev/` page; `core/observability.py` `PerfMiddleware` → `logs/perf.jsonl` (per-request ms/queries) + durable `audit_log` DB table. `logs/` is **ephemeral on Render**.
- **RBAC:** `WriteGuardMiddleware` (Editor vs Viewer, deny-writes-by-default with a read allowlist).

## 2. HARD CONSTRAINTS (non-negotiable)
1. **FREE + fully ISOLATED.** No paid APIs/services. **No Claude/Anthropic or any LLM API dependency in production** (owner will not have it long-term). No external SaaS/CDN. Google Fonts is the only allowed remote in artifacts (N/A to server code).
2. **Money-path safety.** Output feeds D365 (financial/inventory). **Never auto-change parse/pricing/mapping logic.** 100% parsing/mapping accuracy is a hard rule.
3. **MT SO uploads leave Unit Price BLANK** (D365 auto-prices). **Testers = ₹1.** **Store timestamps UTC, display IST.** **Never skip silently** — every dropped PO/line must be logged (named) on Warnings. **Tracker/analytics total = FULL order_value** (exclusions live only in the D365 upload + Verify page).
4. **Minimize tables/columns** (owner preference) but **never at cost of speed**. **No money-path DB ALTER without explicit approval.**
5. **Careful surgery.** Verify every change (tests + Django check + before/after equality for read rewrites). Commit locally; **only `git push` when the owner explicitly confirms** (push auto-deploys to prod).
6. **DRY / skeleton-first.** Reuse scaffolds; no per-feature duplication. **Page-asset separation:** no inline `<style>`/`<script>` in page templates — per-page files in `static/online_b2b/pages/`, wired via `extra_css`/`extra_js`; Django values via a JSON-config `<script>` block.

## 3. Empirical performance picture (measured — `logs/perf.jsonl`, 18,789 requests)
**KEY INSIGHT: the bottleneck is NOT the database.** On every slow route the query count is 2–4 and DB time is 1–5ms — prior index/query work already won that fight. The real cost is **CPU-bound Python processing + large HTML payloads + region round-trips**.

Everyday page loads (p50 / p90 ms · avg queries · payload):
- `/b2b/` dashboard — **5810 / 13552** · q3.0 · 44KB  ← worst everyday page, only 3 queries → pure compute (now cached, see §4)
- `/b2b/tracker/` — 3932 / 7266 · q2.7 · **678KB**
- `/b2b/inventory/` — 3215 / 9533 · q3.9 · **576KB**
- `/b2b/ship-to/` — 3452 / 6571 · q4.0 · **446KB**
- `/b2b/daily/` — 3204 / 8242 · q2.7
- `/b2b/analytics/` — 1047 / 2679 · q2.1 (already optimized — the target state)

Heavy synchronous actions (block the user; p50):
- review/mt/gt **confirm** — ~23–25s (workbook gen + DB record)
- **download** completed workbook — ~14s
- **record-verify** — ~14s (already cut from 363s)

**Implications:** the biggest remaining levers are (a) move the 14–25s confirm/download/verify work OFF the request path into a **background job with progress**, (b) a **faster workbook writer** (openpyxl is CPU-slow), (c) **shrink 400–700KB HTML payloads** + compression. Heavy new DB indexing has near-zero marginal value.

## 4. What was just shipped (8 commits, on `origin/main`, verified, 33/33 tests pass)
All free/isolated, no money-path behaviour change:
1. **`_READY` guards** on 8 more `ensure_table()`/`_ensure()` (lines_store — the worst at ~5 DDL round-trips/call, tat_store, daily_checklist, item_master_loader, channel_map, draft_store, offline_seq_store, offline_master_store). DDL runs once per process now.
2. **Hub read-bundle cached** — `overview()` split into a `_stable`-cached wrapper keyed by `(segment,window)` + `_overview_build()`; `hub_extra_kpis()` wrapped in `_stable('hub_extra')` (removes the 3× `order_lines_full` 350k-row scan). **`_deltas()` folded 2 queries → 1** (conditional aggregation, byte-identical). Lock&Record now `_stable_bust()`s all read-aggregates.
3. **Config:** added `Brotli` to requirements (WhiteNoise emits `.br`); aligned `Procfile` to `render.yaml` free-tier tuning; added `LOGGING`→stderr (durable 500s via Render's free log stream); added a DB-less **`/healthz`** liveness route.
4. **FE (~187KB lighter/page):** removed **Motion One** (−63KB/page; `revealBars` reimplemented with native IntersectionObserver + CSS transition); **lazy-load confetti** on first `B2B.celebrate()` (data-attr URL); **deleted 5 dead assets** (`alpine.min.js`, `htmx.min.js`, `daterange.js/.css`, `pages/orders.js` — verified unreferenced); **`content-visibility:auto`** on inventory (`.iv-stock`) + availability (`.av-lines`) tables.
5. **Availability `check_orders` N+1** → 2 chunked batch queries (was 2×K). **Proved byte-identical** on 26 real POs (deep-compared before/after).
6. **`_backend()` memoized** (was re-reading DB config on every `_conn()`); **PerfMiddleware buffered** (batch flush + atexit, trim off the hot path).
7. **GitHub Actions CI** (`.github/workflows/ci.yml`): ruff (advisory — see §5), pytest (settings_test/sqlite, no TiDB), `manage.py check`.
8. **Fixed stale test** `tests/test_gt_select_import.py` (asserted removed `location` key → now `ship_to_name`).

(Earlier same session, also on main: **AHLC** = Apollo HealthCo new MT child channel; **Channel SKU Map** folded into the Item Master page as a tab.)

## 5. FULL AUDIT — every finding (suggestion log), marked DONE / PENDING
Format: `[Impact][Effort]`. "PENDING (approval)" = needs owner OK before touching.

### Frontend
- `[H][M]` Dual chart libs — analytics page loads **both** ApexCharts + ECharts (~560KB gz). **PENDING (deferred):** per-page loading is already optimal (overview=Apex only, tracker=ECharts only); the only win is porting ~5 ApexCharts charts (area/donut/stacked-bar in `pages/overview.js` + `pages/analytics.js`) to ECharts, then deleting `apexcharts.min.js`. **Requires a browser to verify — do on a branch, not straight to prod.**
- `[H][S]` Motion One global load — **DONE.**
- `[H][S]` No Brotli — **DONE.**
- `[M][S]` `content-visibility` on big tables — **PARTIAL:** done for inventory + availability; **PENDING** for item_master (`.im-table`) and analytics tables.
- `[M][S]` confetti global — **DONE (lazy).**
- `[M][S]` Dead assets — **DONE.**
- `[M][M]` Page-asset-separation violations — inline `<style>`/`<script>` in `_mp_profile.html`, `po_flow/upload.html`/`review.html`/`_verification_modal.html`, `_sidebar.html` (~45-line inline JS), `base_b2b.html` (two big inline `<style>` + ~145 lines inline JS for navProgress/overlay); many inline `style=`/`onclick=`. **PENDING** (also blocks a strict CSP).
- `[L][M]` Own CSS/JS unminified — **PENDING** (needs an esbuild/lightningcss build step; there is currently NO Node toolchain — adding one is a decision).
- `[L][S]` A11y cheap wins (`:focus-visible`, `aria-pressed` on toggles) — **PENDING.**

### Backend (`online_b2b/`, `offline/`, `core/`)
- `[H][S]` Unguarded `ensure_table()` DDL — **DONE.**
- `[H][M]` `overview()` ~9 uncached round-trips — **DONE (cached).**
- `[H][M]` `channel_map.load_master_file` upserts row-by-row (2×N round-trips; duplicated in `offline/services/hg_recon.py`) — **PENDING (approval — money-adjacent):** preload existing ids in one SELECT, bucket, two `executemany`. Verify inserted/updated counts match on a real master.
- `[M][M]` `availability.check_orders` N+1 — **DONE.** (Same shape still at `availability.py:~422` `wh_scenarios` — **PENDING** follow-up.)
- `[M][S]` Every write = 2 audit round-trips (INSERT in `core/access.py:165` + UPDATE in `core/observability.py`) — **PENDING:** collapse to a single INSERT at request end (keep pre-write INSERT only for Lock&Record).
- `[M][S]` MT-confirm per-PO `UPDATE order_headers … external_doc` loop (`offline/services/mt_bridge.py:~2140`) — **PENDING (approval — MONEY PATH):** batch via `executemany`; validate stamped `external_doc` before/after on a real MT dump.
- `[M][S]` Dead code `order_db.dashboard()` + its 3-DISTINCT dropdown N+1 (zero callers) — **PENDING (safe cleanup):** delete (keep `_fetch_orders`/`_count_orders`, still used by `orders_page`).
- `[M][M]` Run-index JSON + `_d365.xlsx`/`_completed.xlsx` sidecars on **ephemeral Render disk** (`online_b2b/views.py:~1018`) vanish on redeploy/spin-down — **PENDING:** move index to a DB table + regenerate workbooks on demand (deterministic; reuse `export_decided_workbook`).
- `[L][S]` `load_db_config()` re-read per `_conn()` — **DONE (memoized).**
- `[L][S]` PerfMiddleware synchronous per-request write + inline 5MB trim — **DONE (buffered).**
- `[L][M]` DRY: 3 near-identical openpyxl export builders (`views.py:1491/1831/3713`) + duplicated SKU-upsert — **PENDING:** extract `services/xlsx.py` + `upsert_sku_map()`.
- `[L][M]` Large exports built fully in RAM (512MB tier) — **PENDING:** openpyxl `write_only=True` + `FileResponse` streaming.

### Database (TiDB Singapore; 45 tables; round-trip count is the lever, not indexes)
- `H1` Hub scans `order_lines_full` 3× per load — **DONE** (via §4.2 caching).
- `H2` Dashboard 3 DISTINCT dropdown scans — **MOOT/cleanup:** `dashboard()` is dead code (see backend item above); the live hub uses `overview()` which is now cached.
- `H3` `_deltas` 2 queries → 1 — **DONE.**
- `M1` Write-path audit INSERT+UPDATE → single — **PENDING** (= backend audit item).
- `M2` `check_orders` N+1 — **DONE.**
- `M3` Tracker calls `list_manual()` twice per render — **SKIPPED ON PURPOSE:** the first loop mutates the dicts in place (`mm['wh']`, `mm['order_value']`), so sharing one fetch risks aliasing bugs; not worth 1 round-trip on a tiny table. (If revisited, deep-copy before the first loop.)
- `M4` Tracker latest-run self-join runs 3× per render (`order_db.py:1246/1251/1328`) — **PENDING (optional, low priority):** already index-covered (`idx_mp_po_ts`); could fold the facility breakdown into the totals query with `GROUP BY warehouse WITH ROLLUP`.
- `L1` No index on `order_lines.ean` — **PENDING (approval, low value):** no hot path filters ean alone today; add only if an ean-keyed feature appears.
- `L2` Dead table `order_issue_lines` (0 rows) — **PENDING (approval):** `DROP TABLE` for table-minimization. (Keep `parked_draft`/`_file`, `ship_to_field`, `tracker_manual` — active-but-empty, not dead.)
- `L3` Cosmetic AUTO_INCREMENT churn (`ship_to_mapping`, `audit_log`) — no action.
- **Table sizes:** order_lines 178k (30MB), order_line_validation 172k (13MB), inventory_bin_line 135k, order_headers 7.6k (well-indexed, tiny).

### Architecture / hosting / tooling / security
- `[H][S]` **Leaked Gmail App Password** (`bomn ktfx jhct xexy`) in git history (HEAD already env-only) — **PENDING (OWNER ACTION):** revoke + reissue in Google Account → App Passwords; set via `EMAIL_PASSWORD` env. Rotation is the real fix (history rewrite optional).
- `[H][S]` No CI — **DONE.**
- `[H][S]` No spin-down mitigation — **DONE** (`/healthz` added); **PENDING (OWNER ACTION):** wire a FREE ping (cron-job.org or a GH Actions `schedule:`) every ~10 min during IST work hours.
- `[M/H][M]` Django cache framework unused — **PARTIAL:** hub is now `_stable`-cached; **PENDING/optional:** `@cache_page(60)` on other read-only dashboards (tracker/inventory) with 1 worker → LocMemCache is fine; DB-table cache if workers grow. Never cache money-path/live pages.
- `[M][S]` Brotli — **DONE.** · `[M][S]` stderr LOGGING — **DONE.** · `[M][S]` Procfile drift — **DONE.**
- `[M][S]` No pre-commit / uv / asset minify — **PARTIAL** (CI added); **PENDING:** `pre-commit` (ruff), `uv` for faster installs, minify step (decision — needs Node).
- `[L][M]` Heavy deps (pandas/numpy/pdfplumber/pillow/calamine) on 512MB — **PENDING:** confirm they're lazily imported in workbook/PDF paths, not at web-view module top-level.
- **Already good (don't redo):** WhiteNoise hashed+compressed static; gunicorn tuned for 0.1 CPU; region pinned Singapore; per-thread pymysql pool; parameterized SQL; `settings_test.py` (sqlite) + `tests/` exist.
- **Lint debt:** ~65 pre-existing ruff violations project-wide (ruff scoped to the web project in `pyproject.toml`, `select=E,F,I,UP,B`). CI runs ruff **advisory** until cleared — then flip to a hard gate.

## 6. Pending backlog (prioritized for the next agent)
**Tier 1 — biggest felt-speed win (the real remaining lever):**
- Move the 14–25s **confirm / download / record-verify** work off the request path → a **DB-backed background job** (no celery/redis available — use a small `job` table + a worker thread or a management command polled by the existing `/healthz` ping) **with a progress bar**. Regenerate D365/Completed workbooks on demand instead of ephemeral-disk sidecars (folds in the reliability finding).
- Faster workbook writer (openpyxl → xlsxwriter or `write_only=True` streaming).

**Tier 2 — safe, no approval needed:**
- Delete dead `order_db.dashboard()`. Batch `wh_scenarios` like `check_orders`. `content-visibility` on item_master/analytics tables. Extract `services/xlsx.py` (DRY exports) + streaming exports. Page-asset-separation cleanup (move inline style/script to per-page files). A11y `:focus-visible`.

**Tier 3 — needs OWNER approval (money-path or schema):**
- `executemany` batching: MT-confirm `external_doc` UPDATE; channel-master/hg_recon SKU upsert (verify counts equal before/after).
- DB ALTER: `DROP TABLE order_issue_lines`; optional `order_lines.ean` index.

**Tier 4 — needs a browser / build decision:**
- ApexCharts → ECharts port (on a branch, verify each chart) → delete ApexCharts + tree-shake ECharts (needs a Node build step — a toolchain decision).

**OWNER actions (not code):** rotate the leaked Gmail app password; wire the free `/healthz` keep-alive ping.

## 7. How to verify ANY change (required)
- `python -m pytest -q` (uses `renee_cosmetics.settings_test`, sqlite — no TiDB needed). Must stay green (33 tests; 1 was a stale test, now fixed).
- `python manage.py check` (and `DJANGO_SETTINGS_MODULE=renee_cosmetics.settings_test python manage.py check`).
- **Read-path rewrites:** capture the live output on real data BEFORE the change, deep-compare AFTER (see how `availability.check_orders` was proven byte-identical on 26 POs). 
- **Money-path:** never auto-change; batch-only optimizations must be proven output-identical and are owner-gated.
- DB access for evidence: `from online_b2b.services import order_db as odb; with odb._conn() as (cur,d): ...` (read-only; `d['ph']='%s'`).
- Commit locally; **do not `git push` without explicit owner confirmation** (auto-deploys to prod).

---
*Generated as a handoff brief. Delete or move this file as you like — it is not part of the app.*
