# OMT Offline

A small desktop copy of the RENÉE Order Management pricing flow. Runs on one
machine with **no database, no internet, no login**.

Give it a PO file and a margin; it gives you the sheet:

| EAN | Item No (Master) | MRP | Landing (70%) | GST Code | Cost Price | Diffn with Cost | Issue |
|---|---|---|---|---|---|---|---|

Channels: **Blinkit · RK · Swiggy · GT Mass**

---

## Run it

```bash
pip install -r requirements.txt
python main.py
```

That's it — two files do the work: `main.py` is the window, `core.py` is
everything else.

## Use it

1. **Item Master** → *Browse* — load it **once**.
2. **Ship-To Mapping** and **Swiggy SKU Map** → *Browse* — same, once.
3. Pick a **Channel**. The margin fills in with that channel's default
   (Blinkit 70 · RK 70 · Swiggy 80 · GT Mass 70) — change it freely.
4. **PO file** → *Browse*, then **Generate**.
5. **Download sheet** — writes the 7-sheet workbook.

| Shortcut | |
|---|---|
| `Ctrl+O` | open a PO file |
| `F5` | generate |
| `Ctrl+S` | download the workbook |
| `Ctrl+F` | jump to Find |
| `Ctrl+C` | copy the selected row |

Results can be filtered as you type, narrowed to **Only flagged**, and sorted by
clicking any column heading.

### The three reference files

Load them once — the app keeps its own copy and reloads them at every launch,
showing when each was **last updated**. Browse again only when you want to
refresh one. Because it holds its own copy, it keeps working if the original
moves out of OneDrive or you're offline.

| Box | Accepts | Needed for |
|---|---|---|
| Item Master | `/b2b/item-master/export/`, or `Items March.xlsx` | everything |
| Ship-To Mapping | `/b2b/ship-to/export/`, or `Ship to B2B.xlsx` | Headers (SO) cust/ship-to |
| Swiggy SKU Map | channel SKU export | Swiggy files with no EAN |

Column names are matched loosely (`Item No` / `item_no` / `ITEMNO` all work), so
a renamed header won't break it, and extra columns are ignored.

### Templates

**Templates** menu — saves a ready-made file with the right columns and a
*How to use* sheet:

- Item Master · Ship-To Mapping · Swiggy SKU Map
- A **standard PO file template for each channel** (Blinkit, RK, Swiggy,
  GT Mass) using that channel's real column names — including GT Mass's
  metadata rows above the header.

## What you get out

A 7-sheet workbook in the same shape and styling as the web app:

| # | Sheet | |
|---|---|---|
| 1 | Headers (SO) | one row per PO — customer + ship-to from your mapping |
| 2 | Lines (SO) | line no. steps 10000 per PO · **Unit Price left blank** |
| 3 | Summary | per-PO totals with a TOTAL row |
| 4 | Validation | the priced block you see on screen |
| 5 | SKU Summary | per item across the file |
| 6 | Raw Data | the punch as read, plus what we computed |
| 7 | Warnings | every flagged line, named |

> **Unit Price stays blank on Lines (SO)** — D365 prices from the vendor master.
> Our computed cost lives on Validation instead.

## How the price is worked out

```
Landing  = MRP × margin%
Cost Price = Landing ÷ (1 + GST)
Diffn with Cost = vendor's cost on the punch − our Cost Price
```

One visible margin per run. No hidden per-SKU rules — what's on screen is what
produced the number. GST comes from the item's code (`G-18-S` → 18%, `G-5` → 5%,
`0-G` → 0%).

A line that can't be priced is **kept and flagged** in the Issue column
(`Not in item master`, `MRP missing`, `Cost differs by +11.69`) — never dropped
quietly.

## Build the .exe

```bash
pip install pyinstaller
pyinstaller omt_offline.spec --noconfirm
```

Produces a single `dist/OMT Offline.exe` — no Python needed on the target
machine. Settings save next to the .exe (or `%LOCALAPPDATA%` if that folder is
read-only), so it works from a USB stick.

---

`src/` and the two `*.md` files are the older PyQt6 reporting app, kept for
reference. Nothing in this app imports them.
