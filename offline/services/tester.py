"""
offline.services.tester
========================

Store-grouped **tester Sales Order** generator for MT channels.

Some channels ship a separate *tester-requirement sheet* alongside the
regular replenishment PO(s). Each row asks for one tester unit of a SKU at a
given store. This module turns that sheet into a small set of tester SOs —
**one SO per store** — that are appended to the run's output (AFTER all the
regular orders) and recorded to the dashboard, so testers are clearly
separated and identifiable.

Design — skeleton first
-----------------------
This is a **channel-agnostic** producer: it takes a channel config + a run
date + a resolved counter and returns plain header/line dicts plus named
warnings. **Lifestyle (LS) is the first consumer** (``build_ls_testers``),
but another channel can reuse :func:`build_testers` by supplying its own
``(store, ean)`` rows and party. No engine code is touched — this reads the
DB (``ship_to_mapping`` / ``item_master``) and the frozen engine's
``ChannelConfig`` only.

SO numbering — ``SO/<CH>/TT/<counter>``
---------------------------------------
Testers share the SAME daily counter space as the regular run (the frozen
``mt_select_seq.json`` counter, base ``DDMMYY`` incrementing per SO), but use
the literal ``TT`` segment where a regular LS SO uses the month (``SO/LS/07``).
The regular run assigns its block first (burning the counter); tester numbers
continue from the persisted ``next_counter`` so they never collide.

Never silent
------------
An unresolved EAN or a store with no ship-to mapping is NOT dropped quietly —
the line/store is listed with a named warning so the operator sees exactly
what could not be placed.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

# Per-unit tester price for LS (task spec). A channel that wants a different
# price passes it explicitly to build_testers().
LS_TESTER_UNIT_PRICE = 0.54


@dataclass
class TesterResult:
    """Everything a run needs to append + record the tester SOs.

    ``headers`` — one dict per store SO (keys mirror the online Headers/Tracker
    shape: po=SO number, ship_to, external_doc, location, qty, order_value).
    ``lines``   — one dict per (store, item) tester line (qty 1, unit_price).
    ``warnings``— named, never-silent notes (unresolved EANs, missing ship-to,
    non-Approved remarks).
    """
    headers: list = field(default_factory=list)
    lines: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    stores: int = 0
    price: float = LS_TESTER_UNIT_PRICE

    @property
    def line_count(self) -> int:
        return len(self.lines)

    @property
    def value(self) -> float:
        return round(sum(float(ln.get('unit_price') or 0) * int(ln.get('qty') or 1)
                         for ln in self.lines), 2)


def _cid(v) -> str:
    """Coerce an id the way the engine does — str, stripped, no trailing '.0'."""
    s = str(v if v is not None else '').strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


# ── tester-sheet detection ───────────────────────────────────────────────
# The LS tester sheet is distinguished from the .xlsb replenishment (which
# has 'Order No' / 'Plant ID' / 'Final Order Qty') by carrying a STORE code
# column + an EAN column + a 'Tester Req' column. Detect by COLUMNS (robust);
# a filename containing 'test'/'teste' is only a hint.
_STORE_ALIASES = {'store code', 'store', 'store_code', 'code'}
_EAN_ALIASES = {'ean', 'ean/upc', 'ean code', 'ean no', 'ean no.', 'barcode'}
_FLAG_ALIASES = {'tester req', 'tester', 'tester_req', 'testerreq', 'tester req '}
_REMARK_ALIASES = {'commercial remarks', 'remarks', 'remark', 'commercial remark'}
_NAME_ALIASES = {'product name', 'description', 'name', 'item name'}


def is_ls_tester_file(path) -> bool:
    """True if a sheet carries a STORE column + an EAN column + a Tester flag
    column — the LS tester-requirement layout. Detected by columns (not the
    filename) so it is robust to naming."""
    import pandas as pd
    if not str(path).lower().endswith(('.xlsx', '.xls', '.xlsm')):
        return False
    try:
        xl = pd.ExcelFile(path)
    except Exception:  # noqa: BLE001
        return False
    for sh in xl.sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name=sh, header=None, nrows=8, dtype=str)
        except Exception:  # noqa: BLE001
            continue
        for _, row in raw.iterrows():
            cells = {str(c).strip().lower() for c in row if c is not None}
            if (cells & _STORE_ALIASES) and (cells & _EAN_ALIASES) and (cells & _FLAG_ALIASES):
                return True
    return False


def _read_tester_rows(path) -> tuple[list[dict], list[str]]:
    """Parse the LS tester sheet → ``[{store, ean, name, remark}]`` for rows
    with a truthy 'Tester Req'. Returns ``(rows, warnings)``; never raises."""
    import pandas as pd
    warnings: list[str] = []
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:  # noqa: BLE001
        return [], [f"Tester sheet: cannot open ({type(e).__name__}: {e})."]
    for sh in xl.sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name=sh, header=None, dtype=str)
        except Exception:  # noqa: BLE001
            continue
        hdr_idx = None
        for i in range(min(8, len(raw))):
            cells = {str(c).strip().lower() for c in raw.iloc[i] if c is not None}
            if (cells & _STORE_ALIASES) and (cells & _EAN_ALIASES) and (cells & _FLAG_ALIASES):
                hdr_idx = i
                break
        if hdr_idx is None:
            continue
        low = [str(c).strip().lower() for c in raw.iloc[hdr_idx]]

        def _col(aliases, low=low):
            return next((j for j, h in enumerate(low) if h in aliases), None)
        cs, ce, cf = _col(_STORE_ALIASES), _col(_EAN_ALIASES), _col(_FLAG_ALIASES)
        cr, cn = _col(_REMARK_ALIASES), _col(_NAME_ALIASES)
        rows: list[dict] = []
        for i in range(hdr_idx + 1, len(raw)):
            store = _cid(raw.iat[i, cs]) if cs is not None else ''
            ean = _cid(raw.iat[i, ce]) if ce is not None else ''
            flag = _cid(raw.iat[i, cf]).lower() if cf is not None else ''
            if not store or not ean or store.lower() == 'nan' or ean.lower() == 'nan':
                continue
            if flag in ('', '0', 'nan', 'none'):
                continue                          # Tester Req not truthy → skip
            rows.append({
                'store': store, 'ean': ean,
                'name': (_cid(raw.iat[i, cn]) if cn is not None else ''),
                'remark': (_cid(raw.iat[i, cr]) if cr is not None else ''),
            })
        return rows, warnings
    return [], ['Tester sheet: no (Store, EAN, Tester Req) header row found.']


# ── store → ship-to + EAN → item resolution (party-scoped) ────────────────
def _resolve_shipto(party: str, stores) -> tuple[dict, dict]:
    """``{store number: ship_to_code}`` for the party. Match on the DB ``name``
    column (the bare store number) first, else the trailing number of
    ``del_location``. Returns ``(store→ship_to, store→del_location)``."""
    to_ship: dict = {}
    to_loc: dict = {}
    try:
        from online_b2b.services.order_db import _conn
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(
                f"SELECT del_location, ship_to, name FROM ship_to_mapping "
                f"WHERE party={ph}", (party,))
            for dl, shp, nm in cur.fetchall():
                dl, shp, nm = str(dl or ''), str(shp or ''), _cid(nm)
                if nm:
                    to_ship.setdefault(nm, shp)
                    to_loc.setdefault(nm, dl)
                m = re.search(r'(\d{3,6})\s*$', dl)   # trailing store number
                if m:
                    to_ship.setdefault(m.group(1), shp)
                    to_loc.setdefault(m.group(1), dl)
                # Also index by the FULL del_location (upper-cased) — for channels
                # keyed by store NAME rather than a trailing number (e.g. BN:
                # del_location 'PRAGATHI NAGAR'). Additive; number keys still win.
                if dl:
                    to_ship.setdefault(dl.strip().upper(), shp)
                    to_loc.setdefault(dl.strip().upper(), dl)
    except Exception:  # noqa: BLE001 — no DB → empty maps, caller warns
        pass
    return to_ship, to_loc


def _resolve_items(eans) -> dict:
    """``{ean: item_no}`` from ``item_master`` for the given EANs."""
    out: dict = {}
    uniq = sorted({_cid(e) for e in eans if _cid(e)})
    if not uniq:
        return out
    try:
        from online_b2b.services.order_db import _conn
        with _conn() as (cur, d):
            ph = ','.join([d['ph']] * len(uniq))
            cur.execute(f"SELECT ean, item_no FROM item_master WHERE ean IN ({ph})",
                        uniq)
            for e, itm in cur.fetchall():
                out[_cid(e)] = str(itm)
    except Exception:  # noqa: BLE001
        pass
    return out


# ── the core producer (channel-agnostic) ─────────────────────────────────
def build_testers(rows: list[dict], *, party: str, sell_to: str, channel_code: str,
                  run_date, counter_start: int, location_code: str = 'PICK',
                  unit_price: float = LS_TESTER_UNIT_PRICE) -> tuple[TesterResult, int]:
    """Turn ``rows`` ([{store, ean, name?, remark?}]) into store-grouped tester
    SOs. Returns ``(TesterResult, next_counter)`` — ``next_counter`` is the
    counter value AFTER burning one number per emitted store SO, so the caller
    can persist it back to the shared seq state.

    One SO per store: header Sell-to ``sell_to``, Ship-to = the store's code,
    External Document No. ``TESTER-<store>``, dates = ``run_date``, Location
    ``location_code``; SO number ``SO/<CH>/TT/<counter>`` (6-digit, mirrors the
    regular block). Lines: Item, qty 1, Unit Price ``unit_price`` each.
    """
    import datetime as _dt
    res = TesterResult(price=round(float(unit_price), 2))
    if not rows:
        return res, counter_start

    stores = sorted({r['store'] for r in rows}, key=lambda s: (len(s), s))
    to_ship, to_loc = _resolve_shipto(party, stores)
    e2item = _resolve_items(r['ean'] for r in rows)

    if isinstance(run_date, (_dt.date, _dt.datetime)):
        date_str = run_date.strftime('%d-%m-%Y')
    else:
        date_str = str(run_date)

    # Never-silent: flag any non-Approved remark (still generated).
    non_appr = [r for r in rows
                if r.get('remark') and r['remark'].strip().lower() != 'approved']
    if non_appr:
        by_store: dict = {}
        for r in non_appr:
            by_store.setdefault(r['store'], []).append(r.get('remark') or '?')
        detail = '; '.join(f"store {s}: {len(v)} × '{v[0]}'" for s, v in by_store.items())
        res.warnings.append(
            f"Tester: {len(non_appr)} row(s) with a non-Approved remark "
            f"generated anyway — {detail}. Verify before dispatch.")

    counter = counter_start
    for store in stores:
        ship = to_ship.get(store)
        loc = to_loc.get(store) or store
        so = f"SO/{channel_code}/TT/{counter:06d}"
        srows = [r for r in rows if r['store'] == store]

        placed = 0
        placed_qty = 0
        line_no = 10000
        so_lines: list[dict] = []
        for r in srows:
            item = e2item.get(_cid(r['ean']))
            if not item:
                res.warnings.append(
                    f"Tester store {store}: EAN {r['ean']} "
                    f"({r.get('name') or 'unnamed'}) not in item_master — line "
                    f"skipped from the SO (listed here, not silent).")
                continue
            q = int(r.get('qty', 1) or 1)      # per-row qty (BN=1; SS GWP=order qty)
            so_lines.append({
                'po': so, 'line_no': line_no, 'item_no': item,
                'ean': _cid(r['ean']),
                'description': (r.get('name') or '')[:255],
                'qty': q, 'unit_price': res.price,
                'store': store, 'ship_to': ship or '',
                'location': loc, 'is_tester': True,
            })
            line_no += 10000
            placed += 1
            placed_qty += q

        if ship is None:
            res.warnings.append(
                f"Tester store {store}: no ship-to mapping (party '{party}') — "
                f"SO {so} written with a blank Ship-to; add the store (cust "
                f"{sell_to}) and re-run.")

        if not so_lines:
            # Store resolved to zero placeable lines → no empty SO. Already
            # warned per line above; note the store too.
            res.warnings.append(
                f"Tester store {store}: 0 lines placed — no tester SO created.")
            continue

        counter += 1
        res.headers.append({
            'po': so, 'sell_to': sell_to, 'ship_to': ship or '',
            'external_doc': f"TESTER-{store}", 'location': loc,
            'location_code': location_code, 'store': store,
            'po_date': date_str, 'exp_date': date_str,
            'qty': placed_qty, 'order_value': round(placed_qty * res.price, 2),
            'is_tester': True,
        })
        res.lines.extend(so_lines)

    res.stores = len(res.headers)
    return res, counter


# ── LS convenience wrapper (the first consumer) ───────────────────────────
def build_ls_testers(path, channel, *, run_date=None, counter_start: int,
                     location_code: str = 'PICK') -> tuple[TesterResult, int]:
    """Read the LS tester sheet at ``path`` and produce store-grouped tester
    SOs for the LS channel. ``counter_start`` is the shared seq counter to
    continue from (after the regular run burned its block)."""
    import datetime as _dt
    rows, warns = _read_tester_rows(path)
    party = getattr(channel, 'party', 'LS')
    sell_to = getattr(channel, 'sell_to', '20044')
    code = getattr(channel, 'code', 'LS')
    price = getattr(channel, 'tester_unit_price', None) or LS_TESTER_UNIT_PRICE
    res, nxt = build_testers(
        rows, party=party, sell_to=sell_to, channel_code=code,
        run_date=run_date or _dt.date.today(), counter_start=counter_start,
        location_code=location_code, unit_price=price)
    res.warnings = warns + res.warnings
    return res, nxt


# ── BN (Apollo) convenience wrapper ───────────────────────────────────────────
# The BN tester file ("B&N tester July.xlsx" → sheet 'Tester Sample format') is
# store-NAME keyed (STORE_NAME → ship_to_mapping del_location), NOT store-number
# keyed like LS. QMART is excluded per owner instruction.
BN_QMART_CODES = {'59000'}
BN_QMART_NAMES = {'QMART'}


def _read_bn_tester_rows(path):
    """Parse the BN tester sheet → ``[{store, ean, name, remark, code}]`` for rows
    with a truthy Tester Qty, EXCLUDING QMART. ``store`` = STORE_NAME (upper-cased,
    the ship_to_mapping del_location key). Returns ``(rows, warnings)``."""
    import pandas as pd
    warnings = []
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:  # noqa: BLE001
        return [], [f"BN tester sheet: cannot open ({type(e).__name__}: {e})."]
    for sh in xl.sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name=sh, header=None, dtype=str)
        except Exception:  # noqa: BLE001
            continue
        hdr_idx = None
        for i in range(min(8, len(raw))):
            cells = {str(c).strip().lower() for c in raw.iloc[i] if c is not None}
            if ('store_name' in cells or 'store name' in cells) and \
               ('ean' in cells) and ('tester qty' in cells or 'tester q' in cells):
                hdr_idx = i
                break
        if hdr_idx is None:
            continue
        low = [str(c).strip().lower() for c in raw.iloc[hdr_idx]]

        def _col(names, low=low):
            return next((j for j, h in enumerate(low) if h in names), None)
        c_name = _col({'store_name', 'store name'})
        c_code = _col({'store code', 'store_code', 'store', 'code'})
        c_ean = _col({'ean', 'ean code', 'gtin'})
        c_prod = _col({'product name', 'itemdescription', 'description', 'name'})
        c_tq = _col({'tester qty', 'tester q', 'tester'})
        rows, excluded = [], 0
        for i in range(hdr_idx + 1, len(raw)):
            sname = _cid(raw.iat[i, c_name]) if c_name is not None else ''
            code = _cid(raw.iat[i, c_code]) if c_code is not None else ''
            ean = _cid(raw.iat[i, c_ean]) if c_ean is not None else ''
            tq = _cid(raw.iat[i, c_tq]) if c_tq is not None else ''
            if not sname or not ean or sname.lower() == 'nan' or ean.lower() == 'nan':
                continue
            if tq in ('', '0', 'nan', 'none'):
                continue                                  # no tester asked
            if sname.strip().upper() in BN_QMART_NAMES or code.strip() in BN_QMART_CODES:
                excluded += 1
                continue                                  # QMART excluded (owner)
            rows.append({
                'store': sname.strip().upper(), 'ean': ean,
                'name': (_cid(raw.iat[i, c_prod]) if c_prod is not None else ''),
                'remark': '', 'code': code})
        if excluded:
            warnings.append(f"QMART excluded: {excluded} tester row(s) dropped "
                            f"(owner instruction).")
        return rows, warnings
    return [], ['BN tester sheet: no (STORE_NAME, EAN, Tester Qty) header found.']


def build_bn_testers(path, *, sell_to='20735', channel_code='BN', run_date=None,
                     counter_start, location_code='PICK',
                     unit_price=LS_TESTER_UNIT_PRICE):
    """Read the BN tester sheet and produce store-grouped tester SOs (QMART
    excluded). Resolution is by STORE_NAME → party 'BN' del_location. Returns
    ``(TesterResult, next_counter)``."""
    import datetime as _dt
    rows, warns = _read_bn_tester_rows(path)
    res, nxt = build_testers(
        rows, party='BN', sell_to=sell_to, channel_code=channel_code,
        run_date=run_date or _dt.date.today(), counter_start=counter_start,
        location_code=location_code, unit_price=unit_price)
    res.warnings = warns + res.warnings
    return res, nxt


# ── SS (Shoppers Stop) GWP wrapper ────────────────────────────────────────────
# The SS "tester" file is the GWP (Gift-With-Purchase) SAP PO — same layout as the
# regular SS PO (Purchasing Document · EAN Code · Plant · Order Quantity), NOT a
# 'Tester Sample format' sheet. Store key = Plant (site code) → resolved by the
# trailing-number rule (party 'SS'). Qty = the sheet's Order Quantity (NOT 1, per
# owner). Nominal tester price per unit (0.54) — flagged; change if regular price.
def _read_ss_gwp_rows(path):
    """Parse the SS GWP PO → ``[{store, ean, name, qty}]`` (store=Plant/site code,
    qty=Order Quantity). Returns ``(rows, warnings)``."""
    import pandas as pd
    warnings = []
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:  # noqa: BLE001
        return [], [f"SS GWP sheet: cannot open ({type(e).__name__}: {e})."]
    for sh in xl.sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name=sh, header=None, dtype=str)
        except Exception:  # noqa: BLE001
            continue
        hdr_idx = None
        for i in range(min(6, len(raw))):
            cells = {str(c).strip().lower() for c in raw.iloc[i] if c is not None}
            if ('plant' in cells) and ('ean code' in cells or 'ean' in cells) \
               and ('order quantity' in cells or 'order qty' in cells):
                hdr_idx = i
                break
        if hdr_idx is None:
            continue
        low = [str(c).strip().lower() for c in raw.iloc[hdr_idx]]

        def _col(names, low=low):
            return next((j for j, h in enumerate(low) if h in names), None)
        c_site = _col({'plant'})
        c_ean = _col({'ean code', 'ean', 'gtin'})
        c_qty = _col({'order quantity', 'order qty', 'quantity', 'po qty'})
        c_txt = _col({'short text', 'product name', 'description'})
        rows = []
        for i in range(hdr_idx + 1, len(raw)):
            site = _cid(raw.iat[i, c_site]) if c_site is not None else ''
            ean = _cid(raw.iat[i, c_ean]) if c_ean is not None else ''
            qty = _cid(raw.iat[i, c_qty]) if c_qty is not None else ''
            if not site or not ean or site.lower() == 'nan' or ean.lower() == 'nan':
                continue
            try:
                q = int(float(qty))
            except (TypeError, ValueError):
                q = 0
            if q <= 0:
                continue
            rows.append({
                'store': site, 'ean': ean, 'qty': q, 'remark': '',
                'name': (_cid(raw.iat[i, c_txt]) if c_txt is not None else '')})
        return rows, warnings
    return [], ['SS GWP sheet: no (Plant, EAN Code, Order Quantity) header found.']


def build_ss_testers(path, *, sell_to='20041', channel_code='SS', run_date=None,
                     counter_start, location_code='PICK',
                     unit_price=LS_TESTER_UNIT_PRICE):
    """Read the SS GWP PO and produce store-grouped tester SOs — qty = the sheet's
    Order Quantity, resolved by Plant/site code (party 'SS'). Returns
    ``(TesterResult, next_counter)``."""
    import datetime as _dt
    rows, warns = _read_ss_gwp_rows(path)
    res, nxt = build_testers(
        rows, party='SS', sell_to=sell_to, channel_code=channel_code,
        run_date=run_date or _dt.date.today(), counter_start=counter_start,
        location_code=location_code, unit_price=unit_price)
    res.warnings = warns + res.warnings
    return res, nxt

