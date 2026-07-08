# Session Summary — 2026-07-05

Order-Management web app (Renée Cosmetics). Django wrapping frozen engines,
recording to MySQL `renee_orders`. This session covered five threads: the Issues
email MP-wise loss + uploaded-%, email-modal polish, offline/online overlay
parity, the **Reliance Smart Bazaar** MT channel integration, and a full
Excel↔PDF↔D365 triangular check. Plus a valuation discussion.

---

## 1. Issues email — MP-wise loss + uploaded %

**File:** `online_b2b/services/issue_email.py`, `online_b2b/services/order_db.py`

Added a **Loss by marketplace** table to the Issues email with columns:

| Marketplace | Lot Qty | Uploaded % | Value | Loss | Mismatch | Not-in-master | Excluded |
|---|---|---|---|---|---|---|---|

- **Lot Qty** = full uploaded lot per MP (all lines), via new `order_db.mp_lot_qty(marketplace, date_from, date_to)` — queries `order_lines_full` over the same date/marketplace scope as the email.
- **Uploaded %** = (lot − excluded qty) ÷ lot, shown to **2 decimals** (fixed a bug where 33 excl / 24,565 lot rounded up to a misleading "100%" — now shows the true 99.87%). Green only at exactly 100.00%, amber otherwise, with a `(n excl / n lot)` note.
- **Mismatch / Not-in-master / Excluded** split into three separate ₹ columns (was one squashed text cell).

Verified via live email preview (status 200, all columns present).

---

## 2. Email modal polish (Issues → ✉ Email)

**File:** `online_b2b/templates/online_b2b/issues.html` + `issue_email.py`

- **Table headers** in the email body switched from heavy navy `#1A237E` fill to a **light slate** header (`#eef2f7` bg, slate-700 text, subtle bottom rule). Zero blue fills remain.
- **Loss-by-marketplace table** made full-width (removed `max-width:640px`) so it lines up with the lines table below.
- **Modal chrome**: To/Cc/Subject rows now vertically centered; Note top-aligned; consistent 66px label column; soft indigo focus accent `#818cf8` (off the hard navy); inputs on light slate fill turning white on hover/focus; header bar tinted; close-button hover state.

---

## 3. Offline MT recording overlay → parity with Online B2B

**File:** `online_b2b/templates/po_flow/review.html`

The offline po_flow "Recording…" overlay had only a lock + progress bar; Online B2B had a **3-step animated progress list**. Ported the same 3-step markup (Recording POs → Generating D365 dump → Finalizing) and the identical stepped-progress JS driver (`startProgress`/`finishProgress`/`failProgress`), reusing the shared `.lo-*` styles in `b2b.css`. MT + GT Mass record modals are now visually identical to Online B2B's.

---

## 4. Reliance Smart Bazaar — new MT child channel

**Files:** `offline/services/mt_bridge.py`, `online_b2b/services/marketplaces.py`, `ship_to_mapping` (DB)

### What it is
**Reliance Smart Bazaar** = Reliance's **hypermarket** format, D365 sell-to **20615** — a *separate customer* from **Reliance Retail (Centro)** = channel `RL`, cust 20043. Identified from PO PDFs:
- Store emails `Smart_Bazaar_*@zmail.ril.com` (Kurla = `hyper_kurla…_phnx`)
- "SB" store-name prefix
- D365 sell-to "Reliance Retail Limited (Reliance Smart Bazaar)"

### Wiring
- Channel code **`RSB`**, cloned from `MET` (Metro) — identical Excel schema (`PurchaseOrders*.xlsx`, sheet `Purchase Orders`; DC_CODE / PURCH_ORDER_NUMBER / EAN_NO / TOTAL_QUANTITY / MRP + dates). Reuses `_normalize_metro_excel`.
- Store key = **DC_CODE** (exact match to `del_location`). Mapping-only, **no price check** (MT rule); supply-margin computed as a note.
- Added to `WEB_CHANNELS`, `CHANNEL_REQUIREMENTS`, routing block, and registry child `mt_rsb` (shows under MT Select → *Reliance Smart Bazaar*).
- SO prefix = **`SO/RSB/…`** (code-derived; still open for confirmation).

### Ship-to mapping (7 rows, party `Reliance Smart Bazaar`, source manual)

| DC_CODE | Ship-to | Store | Pin |
|---|---|---|---|
| FR73 | 20615_1 | COSMOS MALL (SB Siliguri) | 734001 |
| FRBS | 20615_2 | AVANI (SB Howrah) | 711102 |
| FRBW | 20615_3 | JADAVPUR-ORBIT MALL (SB Kolkata) | 700047 |
| FRCB | 20615_4 | SALT LAKE (SB Kolkata) | 700106 |
| FRCG | 20615_5 | WOOD SQUARE MALL (SB Kolkata) | 700103 |
| 6220 | 20615_6 | Phoenix Kurla | 400070 |
| FR49 | 20615_7 | SBZ Rajaji Nagar (SB Bengaluru) | 560010 |

DC_CODE is corroborated three ways: Excel `DC_CODE` = PDF `Site:` = PDF store-email suffix → foolproof key.

### Preview verification (end-to-end)
7 POs, **all READY, 0 errors, 0 warnings**. Every DC → correct ship-to. **qty 2,481 · 510 lines · value ₹5,94,226.38** = Excel = PDF basic. `manage.py check` clean.

---

## 5. Triangular check — Excel ↔ PDF ↔ D365

**Inputs:** `PurchaseOrders (10).xlsx`, 7 PO PDFs (from `.eml`), D365 `Sales Orders` + `Sales Lines` exports.
**Output:** `RELIANCE_TRIANGULAR_EXCEL_PDF_D365.xlsx` (3 sheets: Header Triangle, Line Triangle (510), Price Discrepancies).

### Per-PO facts

| PO | SO | Ship-to | Qty | Excel ₹ | D365 ₹ | Diff |
|---|---|---|---|---|---|---|
| 5115075619 | SO/RSB/07/050726 | 20615_3 | 667 | 146,097.46 | 144,821.43 | −1,276.03 |
| 5115075618 | SO/RSB/07/050727 | 20615_1 | 452 | 118,422.58 | 117,560.36 | −862.22 |
| 5115075617 | SO/RSB/07/050728 | 20615_4 | 374 | 98,517.94 | 97,505.92 | −1,012.02 |
| 5115074671 | SO/RSB/07/050729 | 20615_5 | 435 | 94,551.24 | 93,249.54 | −1,301.70 |
| 5115074670 | SO/RSB/07/050730 | 20615_2 | 389 | 91,408.55 | 90,711.87 | −696.68 |
| 5115074669 | SO/RSB/07/050731 | 20615_7 | 94 | 28,255.01 | 28,255.01 | 0 |
| 5115074668 | SO/RSB/07/050732 | 20615_6 | 70 | 16,973.60 | 16,973.60 | 0 |

### Verdict

| Dimension | Result |
|---|---|
| **Qty** | ✅ 100% — Excel = D365 header = D365 lines, 0 diffs across 510 lines |
| **Ship-to** | ✅ 100% — every DC → correct 20615_N, name + pin |
| **Value** | ⚠ 46 lines differ → net **D365 ₹5,148.65 lower** than Reliance |

### Root cause — 10 SKUs, two families (qty identical; pure rate disagreement)

| SKU family | Item Nos | MRP | Reliance unit | Our D365 unit | Direction |
|---|---|---|---|---|---|
| PRO BANANA POWDER ×3 | 200632/633/634 | 550 | ₹288.10 | ₹313.76 | we bill **more** |
| PRO POWER PUFF COMPACT ×7 | 200687–200693 | 699 | ₹366.14 | ₹324.76 | we bill **less** (drives the net −) |

Reliance's PO costs every SKU at a flat **~52.4% of MRP**; our item-master carries SKU-specific negotiated rates. Because RSB is mapping-only (no price check), SOs correctly went out at **our master price** — the triangle is what surfaces the gap. **Action:** raise with Reliance / verify these 10 SKUs in the item master. Not a processing error.

Supply rate overall sits ~**46–57% of MRP** depending on SKU (vs Reliance's flat 52.4%).

---

## 6. Decisions made

- **FR73 Siliguri pin**: PO PDF shows **734401**, D365 shows **734001**. Decision: **keep mapped to D365 (734001)**; change only if a resolution comes. Not a blocker.
- **Reliance Smart Bazaar is a separate channel** from Centro (different customer 20615 vs 20043).
- Email table headers must be **light**, not blue; tables aligned/full-width.
- Offline recording UI must be **on par** with Online B2B.
- Common "Sending…" email modal polish was **deferred to the end**, then completed.

---

## 7. Valuation discussion (GT Mass standalone Tkinter app)

Question: fair one-time internal price for a standalone GT Mass Tkinter app (no email, no support, no commercial markup).

- **Single no-frills channel (read Excel → SO dump)**: replacement cost ≈ **₹40k–75k** (2–3 weeks dev at Indian rates).
- **Production-hardened (real mapping, dedup, ship-to, clean D365 import)**: **₹1.25–2 lakh**.
- **₹1 lakh** is fair/defensible for a single working channel; below ₹75k undersells a working automation; above ₹3 lakh crosses into value-stream/support pricing (excluded).
- Bigger value is in the **whole system** (all MPs + dashboard + validation), not one channel. Internal software has no real market — "price" ≈ labor it replaces.

---

## Open items

- Confirm **SO prefix** `SO/RSB/…` for Reliance Smart Bazaar.
- Decide whether to standardise the side-by-side **margin** to a single ex-GST basis (~40.5%).
- Add the 10 price-diff SKUs to the **item-master review list** / raise with Reliance (optional).
- FR73 pin — settled to 734001; revisit only on resolution.

## Artifacts produced (in the Reliance Offline folder)

- `RELIANCE_OFFLINE_VERIFICATION.xlsx` — Address Triangle + Value & Margin.
- `RELIANCE_TRIANGULAR_EXCEL_PDF_D365.xlsx` — Header Triangle + Line Triangle (510) + Price Discrepancies.
