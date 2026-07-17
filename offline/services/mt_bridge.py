"""
offline.services.mt_bridge
===========================

Headless library bridge over the FROZEN MT-Select desktop automation
(``offline_po_management/channels/mt_select/standalone_mt_select_automation.py``).

It runs the **exact** processing path the desktop "▶ Generate Sales Orders"
button runs — *load masters → read batch → assign SO numbers → write the
6-sheet workbook* — WITHOUT importing Tkinter, so the web app produces an
identical ``ss_so_*.xlsx``. The standalone file is **never modified** (it stays
the backup): we import it as a module and call its module-level functions.

Why the output matches the desktop exactly
-------------------------------------------
* Same functions, same call order, same args as ``App._do_process`` (the GUI
  worker) — see :meth:`MTProcessor.generate`.
* The standalone resolves masters / sequence-counter / output folder relative to
  its OWN directory (``Path(__file__).parent``), so importing it makes every
  path land in the ``mt_select`` folder — same files the desktop touches.
* Masters are loaded from the SAME source the desktop uses: the saved
  ``master_path`` in ``mt_select_config.json`` (falling back to the bundled
  ``MT_Masters.xlsx``) — :func:`_resolve_master_path`. ``load_all_masters`` only
  *reads* it (snapshotting to a private internal copy), never edits it.
* SO numbers come from the SAME ``mt_select_seq.json`` counter — so web and
  desktop never collide and numbering is continuous.
"""

from __future__ import annotations

import importlib.util
import io
import os
import sys
from contextlib import redirect_stdout
from pathlib import Path

from django.conf import settings

# ── Locate + import the frozen standalone (headless) ─────────────────────
_MT_DIR = (Path(settings.BASE_DIR) / 'offline_po_management' / 'channels'
           / 'mt_select')
_MT_SCRIPT = _MT_DIR / 'standalone_mt_select_automation.py'
_MODULE_NAME = 'mt_select_automation'   # alias (not the file stem) to avoid clashes

_engine_mod = None


def _engine():
    """Import the standalone as a library (cached). Tkinter is imported lazily
    inside its GUI methods, so importing the module is headless-safe."""
    global _engine_mod
    if _engine_mod is None:
        spec = importlib.util.spec_from_file_location(_MODULE_NAME, _MT_SCRIPT)
        if spec is None or spec.loader is None:
            raise ImportError(f"Cannot load MT-Select automation at {_MT_SCRIPT}")
        mod = importlib.util.module_from_spec(spec)
        sys.modules[_MODULE_NAME] = mod
        spec.loader.exec_module(mod)
        _register_web_channels(mod)      # inject web-only channels (RL) at runtime
        _engine_mod = mod
    return _engine_mod


def _register_web_channels(eng) -> None:
    """Register web-only MT channels into the frozen engine's CHANNELS registry
    at RUNTIME — so the frozen ``standalone_mt_select_automation.py`` is never
    edited (golden rule), yet the web exposes new channels. Idempotent."""
    # Reliance Retail (Centro) — cust 20043. Tabular Excel ('Renee.XLSX'), one
    # row per line; EAN lookup; store key = Site code (T8VY/T8WL/…) matched EXACT
    # to the DB ship-to Del Location. No CP check (mapping-only, like BN/LL). The
    # source carries only a per-unit 'Item Price' (pre-GST), so mt_bridge
    # pre-normalizes it and injects an inc-GST 'Value' line total (see
    # _normalize_reliance_excel) — hence csv_value_col='Value'.
    if 'RL' not in eng.CHANNELS:
        eng.CHANNELS['RL'] = eng.ChannelConfig(
            code='RL',
            display_name='Reliance Retail (Centro)',
            party='Reliance Retail',          # ship_to_mapping party (cust 20043)
            input_folder_name='Input_RL',
            output_folder_name='Output_RL',
            sell_to='20043',                  # Reliance Retail Limited
            csv_required_cols=['PO Number', 'Site', 'EAN No.', 'PO Qty'],
            csv_po_col='PO Number',
            csv_store_col='Site',             # Site code → Del Location (exact)
            csv_id_col='EAN No.',
            csv_qty_col='PO Qty',
            csv_mrp_col='MRP',
            csv_cost_col='Item Price',        # per-unit pre-GST (reference only)
            csv_value_col='Value',            # inc-GST line total (pre-normalized)
            csv_expdate_col='Deliv. Date',
            lookup_via='EAN',
            channel_master_sheet=None,
            store_match='exact',
            tester_unit_price=None,           # no testers
            expected_landing_ratio=None,      # value comes from 'Value'
        )
    # Metro Cash & Carry India — cust 20410. Tabular Excel ('PurchaseOrders*.xlsx',
    # sheet 'Purchase Orders'); EAN lookup; store key = DC_CODE (T0SM/T0SL/…)
    # matched EXACT to the DB ship-to Del Location. NO price check (mapping-only).
    # PO + expected dates are IN the file. mt_bridge pre-normalizes it (single
    # clean sheet + inc-GST 'Value' + effective-margin note) — see
    # _normalize_metro_excel.
    if 'MET' not in eng.CHANNELS:
        eng.CHANNELS['MET'] = eng.ChannelConfig(
            code='MET',
            display_name='Metro Cash & Carry',
            party='Metro',                    # ship_to_mapping party (cust 20410)
            input_folder_name='Input_MET',
            output_folder_name='Output_MET',
            sell_to='20410',                  # Metro Cash And Carry India Limited
            csv_required_cols=['PURCH_ORDER_NUMBER', 'DC_CODE', 'EAN_NO',
                               'TOTAL_QUANTITY'],
            csv_po_col='PURCH_ORDER_NUMBER',
            csv_store_col='DC_CODE',          # DC code → Del Location (exact)
            csv_id_col='EAN_NO',
            csv_qty_col='TOTAL_QUANTITY',
            csv_mrp_col='MRP_PER_UNIT',
            csv_cost_col='LANDING_COST_INCL_TAX_PER_UNIT',   # inc-GST (reference)
            csv_value_col='COST_PRICE_INCL_TAX_PER_PO_OU',   # inc-GST line total
            csv_date_col='PURCH_ORDER_DATE',
            csv_expdate_col='EXPECTED_DATE',
            lookup_via='EAN',
            channel_master_sheet=None,
            store_match='exact',
            tester_unit_price=None,           # no testers
            expected_landing_ratio=None,      # no price check (mapping-only)
        )
    # Lifestyle International — cust 20044. Excel BINARY workbook ('Renee Repl Po
    # *.xlsb', one 'Sheet1' holding ALL POs); EAN lookup; store key = numeric
    # 'Plant ID' (3107/1695/…) which mt_bridge maps to the party='LS' Del Location
    # so store_match='exact' resolves it. NO price check (mapping-only). PO +
    # expiry dates are IN the file as Excel serial ints. mt_bridge pre-normalizes
    # it (serial→date, Plant ID→Del Location, inc-GST 'Value' + margin note) — see
    # _normalize_lifestyle_excel.
    if 'LS' not in eng.CHANNELS:
        eng.CHANNELS['LS'] = eng.ChannelConfig(
            code='LS',
            display_name='Lifestyle',
            party='LS',                       # ship_to_mapping party (cust 20044)
            input_folder_name='Input_LS',
            output_folder_name='Output_LS',
            sell_to='20044',                  # Lifestyle International Pvt Ltd
            csv_required_cols=['Order No', 'Plant ID', 'EAN/UPC',
                               'Final Order Qty'],
            csv_po_col='Order No',
            csv_store_col='Plant ID',         # Plant ID → Del Location (normalized)
            csv_id_col='EAN/UPC',
            csv_qty_col='Final Order Qty',
            csv_mrp_col='Item MRP',
            csv_cost_col='Item Unit Value',   # per-unit pre-GST (reference)
            csv_value_col='Total Order value',  # inc-GST line total (in file)
            csv_date_col='Created On Date',
            csv_expdate_col='Not After Date',
            lookup_via='EAN',
            channel_master_sheet=None,
            store_match='exact',
            tester_unit_price=None,           # no testers
            expected_landing_ratio=None,      # no price check (mapping-only)
        )

    # Manash = Purplle OFFLINE. Same tab-separated '.XLS' as online Purplle
    # (normalised by _normalize_manash_excel); multi-store — the 'Address' column
    # is the ship-to lookup key (exact-matches del_location for party 'Manash',
    # cust 20328). Mapping-only like LS: NO price/CP check (MT rule).
    if 'PPL' not in eng.CHANNELS:
        eng.CHANNELS['PPL'] = eng.ChannelConfig(
            code='PPL',
            display_name='Manash (Purplle offline)',
            party='Manash',                   # ship_to_mapping party (cust 20328)
            input_folder_name='Input_MANASH',
            output_folder_name='Output_MANASH',
            sell_to='20328',
            csv_required_cols=['PO Document Number', 'Address', 'EAN Number', 'Qty'],
            csv_po_col='PO Document Number',
            csv_store_col='Address',          # source Address → del_location (exact)
            csv_id_col='EAN Number',
            csv_qty_col='Qty',
            csv_mrp_col='MRP',
            csv_cost_col='Price',             # reference only (no check below)
            csv_value_col='Total value',      # Price × Qty (computed in normalizer)
            csv_date_col='PO Date',
            csv_expdate_col='Expiry Date',
            lookup_via='EAN',
            channel_master_sheet=None,
            store_match='exact',
            tester_unit_price=None,           # no testers
            expected_landing_ratio=None,      # NO price check (mapping-only)
        )

    # Reliance Smart Bazaar (Reliance's HYPERMARKET format) — cust 20615, a
    # DIFFERENT customer from Reliance Retail/Centro (RL, cust 20043). Same
    # 'PurchaseOrders*.xlsx' schema as Metro (sheet 'Purchase Orders', DC_CODE +
    # PURCH_ORDER_NUMBER + EAN_NO + TOTAL_QUANTITY + MRP + dates), so it reuses
    # _normalize_metro_excel. Store key = DC_CODE (FR73/FRBS/6220/…) matched
    # EXACT to the party='Reliance Smart Bazaar' Del Location. NO price check
    # (mapping-only, MT rule); the effective supply margin is computed + noted.
    if 'RSB' not in eng.CHANNELS:
        eng.CHANNELS['RSB'] = eng.ChannelConfig(
            code='RSB',
            display_name='Reliance Smart Bazaar',
            party='Reliance Smart Bazaar',    # ship_to_mapping party (cust 20615)
            input_folder_name='Input_RSB',
            output_folder_name='Output_RSB',
            sell_to='20615',                  # Reliance Retail Limited (Smart Bazaar)
            csv_required_cols=['PURCH_ORDER_NUMBER', 'DC_CODE', 'EAN_NO',
                               'TOTAL_QUANTITY'],
            csv_po_col='PURCH_ORDER_NUMBER',
            csv_store_col='DC_CODE',          # DC code → Del Location (exact)
            csv_id_col='EAN_NO',
            csv_qty_col='TOTAL_QUANTITY',
            csv_mrp_col='MRP_PER_UNIT',
            csv_cost_col='LANDING_COST_INCL_TAX_PER_UNIT',   # inc-GST (reference)
            csv_value_col='COST_PRICE_INCL_TAX_PER_PO_OU',   # inc-GST line total
            csv_date_col='PURCH_ORDER_DATE',
            csv_expdate_col='EXPECTED_DATE',
            lookup_via='EAN',
            channel_master_sheet=None,
            store_match='exact',
            tester_unit_price=None,           # no testers
            expected_landing_ratio=None,      # no price check (mapping-only)
        )

    # H&B (Health & Beauty) — cust 20010. Excel BINARY workbook ('Renee Rep PO
    # Excel *.xlsb', one 'Sheet1' with ALL POs); EAN lookup; store key = numeric
    # 'Site' code matched EXACT to the party='h&b' Del Location. Mapping-only
    # (MT rule): NO price check. Normalised by _normalize_hb_excel (serial→date,
    # de-.0 ids). Site codes are added to Ship-To B2B separately (authoritative
    # list); until a Site is mapped its lines flag UNMAPPED (never silent).
    if 'HB' not in eng.CHANNELS:
        eng.CHANNELS['HB'] = eng.ChannelConfig(
            code='HB',
            display_name='Health & Beauty',
            party='h&b',                      # ship_to_mapping party (cust 20040)
            input_folder_name='Input_HB',
            output_folder_name='Output_HB',
            sell_to='20040',                  # H&B Stores Limited (ship-to 20040_n)
            csv_required_cols=['Purchasing Document', 'Site', 'EAN',
                               'Order Quantity'],
            csv_po_col='Purchasing Document',
            csv_store_col='Site',             # numeric Site code → Del Location
            csv_id_col='EAN',
            csv_qty_col='Order Quantity',
            csv_mrp_col='MRP',
            csv_cost_col='Net price',         # post-GST unit cost (reference only)
            csv_value_col='Net Order Value',  # inc-GST line total (in file)
            csv_date_col='Document Date',
            csv_expdate_col=None,             # no expected-delivery col in file
            lookup_via='EAN',
            channel_master_sheet=None,
            store_match='exact',
            tester_unit_price=None,           # no testers
            expected_landing_ratio=None,      # NO price check (mapping-only)
        )

    # Lulu (LL): Lulu now sends a clean tabular EXCEL instead of the PDF, so
    # flip LL's input PDF → Excel for 100% accuracy. REVERSIBLE + non-destructive:
    # the frozen PDF config + parse_lulu_pdf are left fully intact — we only clear
    # LL.pdf_parser at runtime (so the .xlsx is read as Excel) and normalise the
    # Excel into the columns LL already maps (see _normalize_lulu_excel). Set
    # LULU_EXCEL_MODE=False below to restore the PDF path with zero code churn.
    if LULU_EXCEL_MODE and 'LL' in eng.CHANNELS:
        eng.CHANNELS['LL'].pdf_parser = None
    # LL's ship-to rows live under party 'Lulu' in ship_to_mapping (the name-based
    # convention, same as Metro/Manash/Naturals) but the frozen config says
    # party='LL' — so ship-to never resolved. Align the party to the data.
    if 'LL' in eng.CHANNELS and eng.CHANNELS['LL'].party == 'LL':
        eng.CHANNELS['LL'].party = 'Lulu'


# Lulu input switch: True = read the new Excel (PDF path OFF, reversible);
# False = restore the frozen 'lulu' PDF parser. See _register_web_channels + the
# LL routing in MTProcessor._load.
LULU_EXCEL_MODE = True


# MT (Modern Trade) child channels exposed on the web. MT is the parent
# marketplace; these are its children (the operator picks one). Off Institutional
# (INST) is a SEPARATE parent and is not listed here. SS is verified end-to-end;
# the others share the same generic pipeline (test each before production use).
WEB_CHANNELS = ['SS', 'HG', 'NT', 'BN', 'LL', 'RL', 'MET', 'LS', 'PPL', 'RSB', 'HB']

# ── Per-channel input REQUIREMENTS ──────────────────────────────────────
# Shown on the upload page (so the operator knows what each channel demands)
# AND surfaced as a run note. 'required' = must-have file(s); 'optional' = extra
# file(s) that unlock more checks; 'if_absent' = what happens without the
# optional file (never blocked — golden rule: nothing silent).
CHANNEL_REQUIREMENTS: dict = {
    'HG': {'required': 'H&G PO CSV/Excel (SKU_CODE, QUANTITY, PURCHASE_COST, PO_VALUE).',
           'optional': 'Tester-requirement sheet → mints paired tester SOs (qty 1 @ 0.54).',
           'if_absent': 'No tester sheet → no testers minted (regular SOs only).'},
    'SS': {'required': 'Shoppers Stop SAP Excel — one .xlsx per PO (EAN Code, Plant, Order Quantity).',
           'optional': '', 'if_absent': ''},
    'NT': {'required': 'Naturals PO PDF(s).', 'optional': '', 'if_absent': ''},
    'BN': {'required': 'Apollo tabular Excel (PO Number, DC Name, EAN Code, PO Qty).',
           'optional': '', 'if_absent': ''},
    'LL': {'required': 'Lulu PO PDF(s).', 'optional': '', 'if_absent': ''},
    'RL': {'required': 'Reliance tabular Excel (e.g. Renee.XLSX) — one row per PO line '
                       '(PO Number, Site, EAN No., PO Qty, Item Price).',
           'optional': 'PO PDF(s) — auto cross-checks delivery ADDRESS vs D365 + PO totals, '
                       'and backfills PO Date.',
           'if_absent': 'No PDF → processed on the Excel alone (ship-to from DB, PO Date blank). '
                        'Not blocked.'},
    'MET': {'required': 'Metro tabular Excel (PurchaseOrders*.xlsx, sheet "Purchase Orders") — '
                        'DC_CODE, PURCH_ORDER_NUMBER, EAN_NO, TOTAL_QUANTITY, MRP + dates.',
            'optional': '', 'if_absent': 'No price check — records the inc-GST value; the effective '
            'supply margin (landing ex-GST ÷ MRP) is computed and noted.'},
    'LS': {'required': "Lifestyle replenishment workbook (.xlsb, 'Renee Repl Po *.xlsb', one "
                       "'Sheet1' with all POs) — Order No, Plant ID, EAN/UPC, Final Order Qty, "
                       'Item MRP, Item Unit Value, Total Order value + dates.',
           'optional': 'PO PDF(s) — per-store delivery detail (reference only). '
                       'Tester-requirement sheet (STORE CODE, EAN, Tester Req) → mints '
                       'one tester SO per store (SO/LS/TT/…, qty 1 @ 0.54), appended last.',
           'if_absent': 'No price check — records the inc-GST value; the effective supply margin '
                        '(Item Unit Value ÷ MRP) is computed and noted. No tester sheet → no '
                        'testers. Not blocked.'},
    'PPL': {'required': "Manash (Purplle offline) tab-separated '.XLS' — same format as online "
                           'Purplle (PO Document Number, EAN Number, Qty, MRP, Price, Plant, '
                           'Address + dates). The Address column is the ship-to lookup key.',
               'optional': '', 'if_absent': 'Mapping-only — NO price check. Records value = Price × Qty '
               '(MRP × 0.70 × Qty). Unknown Address → flagged, never silent.'},
    'HB': {'required': "H&B (Health & Beauty) replenishment workbook (.xlsb, 'Renee Rep "
                       "PO Excel *.xlsb', one 'Sheet1' with all POs) — Purchasing Document, "
                       'Site (store code), EAN, Order Quantity, MRP, Net price + Document Date.',
           'optional': '', 'if_absent': 'Mapping-only — NO price check. Records the inc-GST '
           'value; the effective supply ratio (Net price ÷ MRP) is computed and noted. '
           'Unknown Site code → flagged UNMAPPED, never silent (add it to Ship-To B2B, cust 20010).'},
    'RSB': {'required': 'Reliance Smart Bazaar tabular Excel (PurchaseOrders*.xlsx, sheet '
                        '"Purchase Orders") — DC_CODE, PURCH_ORDER_NUMBER, EAN_NO, '
                        'TOTAL_QUANTITY, MRP + dates. DC_CODE (FR73/FRBS/6220/…) is the '
                        'ship-to lookup key (exact) — cust 20615, separate from Centro.',
            'optional': 'PO PDF(s) — reference for the per-store delivery address cross-check.',
            'if_absent': 'No price check — records the inc-GST value; the effective supply '
            'margin (landing ex-GST ÷ MRP) is computed and noted. Unknown DC_CODE → flagged, '
            'never silent.'},
}


def channel_requirements(code: str) -> dict | None:
    """The input requirements descriptor for a channel (for the upload-page hint
    + run note). None if the channel has no descriptor."""
    return CHANNEL_REQUIREMENTS.get(code)

# Accepted upload extensions (SS ships .xlsx; NT/LL ship PDF; RL Excel + optional
# cross-check PDFs).
ACCEPTED_EXTENSIONS = ('.xlsx', '.xls', '.xlsb', '.csv', '.pdf')

# Feature flag — LS PDF address cross-check (the visible verification modal).
# OFF for now: parsing the ~860-page LS PDFs with pdfplumber costs ~3.7 min and
# runs on BOTH preview AND confirm. Re-enable once it's made lazy (on-demand,
# AJAX) + switched to PyMuPDF (~10-20x faster). The condensed multi-store
# warning stays ON regardless (it's cheap). Flip to True to turn it back on.
LS_PDF_VERIFICATION = False


def _normalize_reliance_excel(src_path) -> str:
    """Reliance 'Renee.XLSX' → the flat shape the RL ChannelConfig reads. The
    source has only a per-unit pre-GST 'Item Price', so we inject an inc-GST
    'Value' line total = Item Price × PO Qty × (1 + Tax Rate/100) (this matches
    the PDF's 'Total Order Value' to the paisa). Returns a temp .xlsx path;
    frozen engine reads it exactly like BN's clean Excel."""
    import tempfile

    import pandas as pd
    df = pd.read_excel(src_path)
    df = df[df['PO Number'].notna()].copy()
    mrp_col = next((c for c in df.columns
                    if str(c).strip().lower().startswith('mrp')), None)
    q = pd.to_numeric(df['PO Qty'], errors='coerce')
    ip = pd.to_numeric(df['Item Price'], errors='coerce')
    tr = pd.to_numeric(df.get('Tax Rate'), errors='coerce').fillna(0)
    out = pd.DataFrame({
        'PO Number': [_cid(v) for v in df['PO Number']],
        'Site': df['Site'].astype(str).str.strip(),
        'EAN No.': [_cid(v) for v in df['EAN No.']],
        'PO Qty': q,
        'MRP': (pd.to_numeric(df[mrp_col], errors='coerce') if mrp_col else ''),
        'Item Price': ip,
        'Value': (ip * q * (1 + tr / 100)).round(2),
        'Deliv. Date': df.get('Deliv. Date'),
    })
    fd, path = tempfile.mkstemp(suffix='_reliance_norm.xlsx')
    os.close(fd)
    out.to_excel(path, index=False)
    return path


def _normalize_lulu_excel(src_path):
    """Lulu '<PONumber>.XLSX' → the flat shape the LL ChannelConfig reads
    (PO No · Store · EAN · Qty · MRP · Gross Price · Amount · PO Date ·
    Delivery Date). Lulu now sends a clean tabular Excel — this replaces the PDF
    path (parse_lulu_pdf) 1:1. Drops the trailing total row, cleans EAN + PO,
    emits day-first date STRINGS (engine re-parses dayfirst). Store = the exact
    'Delivery to' (e.g. 'Lulu Hypermarket,Hyderabad') → exact ship-to match.
    Returns ``(temp_path, note)``."""
    import tempfile

    import pandas as pd
    df = pd.read_excel(src_path)
    df = df[df['PO Number'].notna()].copy()          # drop the total row
    # Store = the delivery CITY (last comma token of 'Delivery to', e.g.
    # 'Lulu Hypermarket,Hyderabad' → 'Hyderabad') — LL resolves ship-to via
    # store_match='city_in_name' (the city must appear in the Del Location name),
    # exactly like the PDF path fed the city.
    out = pd.DataFrame({
        'PO No':       [_cid(v) for v in df['PO Number']],
        'Store':       df['Delivery to'].astype(str).str.split(',').str[-1].str.strip(),
        'EAN':         [_cid(v) for v in df['EAN']],
        'Qty':         pd.to_numeric(df['Total Quantity'], errors='coerce'),
        'MRP':         pd.to_numeric(df['MRP'], errors='coerce'),
        'Gross Price': pd.to_numeric(df['Gross Price'], errors='coerce'),
        'Amount':      pd.to_numeric(df['Total Invoice Cost'], errors='coerce'),
    })
    for col in ('PO Date', 'Delivery Date'):
        out[col] = (pd.to_datetime(df[col], errors='coerce')
                    .dt.strftime('%d-%m-%Y').fillna(''))
    fd, path = tempfile.mkstemp(suffix='_lulu_norm.xlsx')
    os.close(fd)
    out.to_excel(path, index=False)
    note = (f"Lulu read from EXCEL ({len(out)} line(s)) — PDF path OFF (reversible). "
            f"Store = 'Delivery to' → exact ship-to match.")
    return path, note


def _normalize_metro_excel(src_path):
    """Metro 'PurchaseOrders*.xlsx' → the flat shape the MET ChannelConfig reads:
    read the 'Purchase Orders' sheet (not 'Sheet1'), drop the junk 'Unnamed'
    column + blank/total rows, clean EAN + PO, keep the columns the config maps.
    Also compute the **effective supply margin** per line — landing ex-GST ÷ MRP —
    so we can tell the operator at what margin we're supplying (no price *check*,
    just the figure). Returns ``(temp_path, margin_note)``."""
    import tempfile

    import pandas as pd
    df = pd.read_excel(src_path, sheet_name='Purchase Orders')
    df = df[df['PURCH_ORDER_NUMBER'].notna()].copy()
    df = df[[c for c in df.columns if not str(c).startswith('Unnamed')]]
    mrp = pd.to_numeric(df['MRP_PER_UNIT'], errors='coerce')
    land_inc = pd.to_numeric(df['LANDING_COST_INCL_TAX_PER_UNIT'], errors='coerce')
    tax = pd.to_numeric(df.get('TAX_PER'), errors='coerce').fillna(0)
    land_ex = land_inc / (1 + tax / 100)
    # effective margin (keep%) we supply at = landing ex-GST ÷ MRP
    marg = (land_ex / mrp * 100).where(mrp > 0)
    df['PURCH_ORDER_NUMBER'] = [_cid(v) for v in df['PURCH_ORDER_NUMBER']]
    df['EAN_NO'] = [_cid(v) for v in df['EAN_NO']]
    # Dates are 'dd.mm.yyyy' — parse DAY-FIRST, then re-emit as DAY-FIRST
    # 'dd-mm-YYYY' STRINGS. The frozen engine's _date_from_raw stringifies the
    # cell and re-parses it with dayfirst=True, so an ISO/date object would be
    # MISread (01.07 → Jan 7); a day-first string round-trips correctly.
    for dc in ('PURCH_ORDER_DATE', 'EXPECTED_DATE'):
        if dc in df.columns:
            df[dc] = (pd.to_datetime(df[dc], dayfirst=True, errors='coerce')
                      .dt.strftime('%d-%m-%Y').fillna(''))
    fd, path = tempfile.mkstemp(suffix='_metro_norm.xlsx')
    os.close(fd)
    df.to_excel(path, index=False)
    m = marg.dropna()
    note = ''
    if len(m):
        note = (f"Supply margin (landing ex-GST ÷ MRP): avg {m.mean():.1f}% "
                f"· range {m.min():.1f}–{m.max():.1f}% across {len(m)} line(s). "
                f"No price check — informational.")
    return path, note


def _lifestyle_store_map() -> dict:
    """``{store number: (Del Location, {all Del Locations for that number})}`` for
    party='LS' — the Plant ID in the .xlsb is a bare store number (3107), but the
    DB ship-to key is the full Del Location with that number embedded. The
    ship_to_mapping ``name`` column holds the store number, so we key on it."""
    out: dict = {}
    try:
        from online_b2b.services.order_db import _conn
        with _conn() as (cur, _d):
            cur.execute("SELECT name, del_location FROM ship_to_mapping WHERE party='LS'")
            for nm, dl in cur.fetchall():
                k = _cid(nm)
                if not k or not dl:
                    continue
                slot = out.setdefault(k, [dl, set()])
                slot[1].add(dl)
    except Exception:  # noqa: BLE001 — no DB → empty map, engine flags stores
        pass
    return out


def _normalize_lifestyle_excel(src_path):
    """Lifestyle 'Renee Repl Po *.xlsb' → the flat shape the LS ChannelConfig
    reads. The source is an Excel BINARY workbook (one 'Sheet1' holding ALL POs),
    with dates as Excel SERIAL integers and the store key as a bare numeric
    'Plant ID' that does NOT equal the DB ship-to Del Location. So we (1) read via
    pyxlsb, (2) convert the two date serials to real dates, (3) de-.0 Order No +
    EAN, (4) map Plant ID → the party='LS' Del Location so store_match='exact'
    resolves it, and (5) compute the effective supply margin (Item Unit Value ÷
    MRP — the source cost is already pre-GST). Returns ``(temp_path, margin_note,
    notes)``; the frozen engine reads the temp .xlsx exactly like Metro's."""
    import datetime as _dt
    import tempfile

    import pandas as pd
    df = pd.read_excel(src_path, sheet_name='Sheet1', engine='pyxlsb')
    df = df[df['Order No'].notna()].copy()
    smap = _lifestyle_store_map()

    def _plant(v):
        s = _cid(v)
        slot = smap.get(s)
        return slot[0] if slot else s      # → Del Location, else raw (engine flags)

    def _serial(v):
        # Excel serial int → DAY-FIRST 'dd-mm-YYYY' string. Day-first (not ISO)
        # because the frozen engine's _date_from_raw parses the cell with
        # dayfirst=True — an ISO 'YYYY-MM-DD' string would be MISread (e.g.
        # 2026-07-01 → 07-Jan). '' when unparseable.
        try:                               # 1900 epoch (Excel's 1899-12-30 base)
            d = _dt.date(1899, 12, 30) + _dt.timedelta(days=int(float(v)))
            return d.strftime('%d-%m-%Y')
        except (TypeError, ValueError):
            return ''

    mrp = pd.to_numeric(df['Item MRP'], errors='coerce')
    cost = pd.to_numeric(df['Item Unit Value'], errors='coerce')   # pre-GST unit
    marg = (cost / mrp * 100).where(mrp > 0)
    out = pd.DataFrame({
        'Order No': [_cid(v) for v in df['Order No']],
        'Plant ID': [_plant(v) for v in df['Plant ID']],
        'EAN/UPC': [_cid(v) for v in df['EAN/UPC']],
        'Final Order Qty': pd.to_numeric(df['Final Order Qty'], errors='coerce'),
        'Item MRP': mrp,
        'Item Unit Value': cost,
        'Total Order value': pd.to_numeric(df['Total Order value'],
                                           errors='coerce').round(2),
        'Created On Date': [_serial(v) for v in df['Created On Date']],
        'Not After Date': [_serial(v) for v in df['Not After Date']],
    })
    fd, path = tempfile.mkstemp(suffix='_lifestyle_norm.xlsx')
    os.close(fd)
    out.to_excel(path, index=False)

    notes: list[str] = []
    # Never silent: ambiguous store number (>1 ship-to code) + unmapped stores.
    for k, (used, dls) in smap.items():
        if len(dls) > 1:
            notes.append(
                f"Lifestyle store {k} maps to {len(dls)} ship-to codes "
                f"({', '.join(sorted(dls))}) — used '{used}'. Verify the correct one.")
    seen = {_cid(v) for v in df['Plant ID']}
    unmapped = sorted(s for s in seen if s and s not in smap)
    if unmapped:
        notes.append(
            f"Lifestyle: {len(unmapped)} store(s) NOT in ship-to mapping (party "
            f"'LS'): {', '.join(unmapped[:15])}{' …' if len(unmapped) > 15 else ''}. "
            f"Add them (cust 20044) — those lines won't resolve a ship-to.")
    m = marg.dropna()
    if len(m):
        notes.append(
            f"Lifestyle supply margin (Item Unit Value ÷ MRP): avg {m.mean():.1f}% "
            f"· range {m.min():.1f}–{m.max():.1f}% across {len(m)} line(s). "
            f"No price check — informational.")
    return path, notes


def _normalize_hb_excel(src_path):
    """H&B (Health & Beauty) 'Renee Rep PO Excel *.xlsb' → the flat .xlsx the HB
    ChannelConfig reads. Source is an Excel BINARY workbook ('Sheet1', ALL POs),
    'Document Date' as an Excel SERIAL int, and the store key = the numeric
    'Site' code (matched EXACT to the party='h&b' Del Location once the Site
    codes are loaded into Ship-To B2B — until then those lines flag UNMAPPED,
    never silent). We (1) read via pyxlsb, (2) serial→day-first date, (3) de-.0
    PO / Site / EAN, (4) note the effective supply ratio (Net price ÷ MRP).
    Returns ``(temp_path, notes)``. Mapping-only — NO price check (MT rule)."""
    import datetime as _dt
    import tempfile

    import pandas as pd
    df = pd.read_excel(src_path, sheet_name='Sheet1', engine='pyxlsb')
    df = df[df['Purchasing Document'].notna()].copy()

    def _serial(v):
        # Excel serial int → DAY-FIRST 'dd-mm-YYYY' (the frozen engine parses the
        # date cell with dayfirst=True; an ISO string would be misread).
        try:
            d = _dt.date(1899, 12, 30) + _dt.timedelta(days=int(float(v)))
            return d.strftime('%d-%m-%Y')
        except (TypeError, ValueError):
            return ''

    mrp = pd.to_numeric(df['MRP'], errors='coerce')
    net = pd.to_numeric(df['Net price'], errors='coerce')     # post-GST unit cost
    ratio = (net / mrp * 100).where(mrp > 0)
    out = pd.DataFrame({
        'Purchasing Document': [_cid(v) for v in df['Purchasing Document']],
        'Site': [_cid(v) for v in df['Site']],
        'Site Name': df['Site Name'].astype(str),
        'EAN': [_cid(v) for v in df['EAN']],
        'Order Quantity': pd.to_numeric(df['Order Quantity'], errors='coerce'),
        'MRP': mrp,
        'Net price': net,
        'Net Order Value': pd.to_numeric(df['Net Order Value'],
                                         errors='coerce').round(2),
        'Document Date': [_serial(v) for v in df['Document Date']],
    })
    fd, path = tempfile.mkstemp(suffix='_hb_norm.xlsx')
    os.close(fd)
    out.to_excel(path, index=False)

    notes: list[str] = []
    r = ratio.dropna()
    if len(r):
        notes.append(
            f"H&B supply ratio (Net price ÷ MRP): avg {r.mean():.1f}% · range "
            f"{r.min():.1f}–{r.max():.1f}% across {len(r)} line(s). No price "
            f"check (mapping-only) — informational.")
    return path, notes


def _normalize_manash_excel(src_path):
    """Manash (Purplle offline) tab-separated '.XLS' → the flat .xlsx the MANASH
    ChannelConfig reads. Same source shape as online Purplle. Cleans the
    zero-padded/quote-suffixed EAN ('000008904473100590'' → '8904473100590');
    the 'Address' column is the ship-to lookup key (exact-matches del_location
    for party 'Manash'). Value = Price × Qty (= MRP × 0.70 × Qty, inc-GST
    landing) — NO price check (mapping-only, MT rule). Returns (temp_path, notes)."""
    import tempfile

    import re as _re

    import pandas as pd
    df = pd.read_csv(src_path, sep='\t', dtype=str, engine='python').fillna('')
    df = df[df['PO Document Number'].astype(str).str.strip() != ''].copy()

    _ILLEGAL = _re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')   # openpyxl-illegal control chars

    def _cln(v):
        return _ILLEGAL.sub('', str(v)).strip()

    def _ean(v):
        s = _cln(v).strip("'")
        return s.lstrip('0') or s

    qty = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)
    price = pd.to_numeric(df['Price'], errors='coerce').fillna(0)
    out = pd.DataFrame({
        'PO Document Number': [_cln(v) for v in df['PO Document Number']],
        'Address': [_cln(v) for v in df['Address']],           # ship-to lookup key
        'EAN Number': [_ean(v) for v in df['EAN Number']],
        'Qty': qty,
        'MRP': pd.to_numeric(df['MRP'], errors='coerce').fillna(0),
        'Price': price,
        'Total value': (price * qty).round(2),                 # inc-GST landing value
        'PO Date': [_cln(v) for v in df['PO Date']],
        'Expiry Date': [_cln(v) for v in df['Expiry Date']],
    })
    fd, path = tempfile.mkstemp(suffix='_manash_norm.xlsx')
    os.close(fd)
    out.to_excel(path, index=False)

    # Ship-to resolution (source Address → del_location) + never-silent warnings
    # for genuinely-unmapped stores are handled by the engine itself (same as
    # every MT channel), so no redundant note here.
    return path, []


def _parse_reliance_pdf(path) -> dict:
    """Pull the header facts from a Reliance PO PDF for the optional cross-check:
    PO No, Site, PO Date, Total Order Value (inc GST), and the delivery pincode
    (from the 'DeliveryAddress' block)."""
    import re

    import pdfplumber
    with pdfplumber.open(path) as p:
        t = '\n'.join((pg.extract_text() or '') for pg in p.pages)
    def _find(pat):
        m = re.search(pat, t)
        return m.group(1).strip() if m else ''
    po = _find(r'PO NO\.?\s*:\s*(\S+)')
    site = _find(r'Site\s*:\s*(\S+)')
    pod = _find(r'PO Date\s*:\s*([\d.]+)')
    tov = re.search(r'Total Order Value\s*:\s*INR\s*([\d,]+\.\d{2})', t)
    val = float(tov.group(1).replace(',', '')) if tov else None
    da = re.search(r'DeliveryAddress:(.*)', t, re.S)
    pin = ''
    if da:
        blk = da.group(1)[:400]
        m = re.search(r'-\s*(\d{6})\b', blk) or re.search(r'(\d{6})', blk)
        pin = m.group(1) if m else ''
    return {'po': str(po), 'site': str(site), 'po_date': pod,
            'value': val, 'pin': pin}


def _parse_lifestyle_pdf(path) -> dict:
    """Parse a Lifestyle PO PDF into ``{store number: {'pincode', 'city'}}``.

    Each per-store delivery block starts ``DELIVERY <storeNo> - <address…>`` and
    runs until the next section marker (``GSTIN`` / ``BRAND`` / ``HSN`` / the next
    ``DELIVERY``). Within that window the delivery pincode is the FIRST whole
    6-digit token — matched with ``(?<!\\d)(\\d{6})(?!\\d)`` on the RAW text (NOT
    space-stripped: stripping would merge plot numbers like 'NO 94 300 304' into a
    fake 6-digit pincode). The city is the last word on the ``DELIVERY`` line
    (e.g. '… Noida'). Bounding at ``GSTIN`` keeps Lifestyle's own corporate pin
    ('Pin : 560037,Bangalore', which sits AFTER the delivery pin) out of the
    window."""
    import re

    import pdfplumber
    with pdfplumber.open(path) as p:
        t = '\n'.join((pg.extract_text() or '') for pg in p.pages)
    # Bound the per-store window at the FIRST section marker after the DELIVERY
    # line. 'PO No.' / 'Corporate' matter because the Lifestyle CORPORATE address
    # ('… Pin : 560037,Bangalore …') sits right after 'PO No.:' — without those
    # bounds a store whose block has NO delivery pincode would wrongly capture the
    # corporate 560037. The real delivery pincode (when present) is always on the
    # DELIVERY / LOCATION line, i.e. BEFORE 'PO No.:'.
    bound = re.compile(r'(?:PO No\.|Corporate|GSTIN|BRAND|HSN|DELIVERY\s+\d)')
    pin_re = re.compile(r'(?<!\d)(\d{6})(?!\d)')
    out: dict = {}
    for m in re.finditer(r'DELIVERY\s+(\d+)\s*-\s*(.*?)(?:\n|$)', t):
        store = m.group(1)
        head = m.group(2)                       # rest of the DELIVERY line (address)
        # City = last non-numeric word on the DELIVERY head line.
        cwords = re.findall(r'[A-Za-z][A-Za-z.]+', head)
        city = cwords[-1] if cwords else ''
        # Window: from just after the store number to the next section marker.
        win = t[m.end(1): m.end(1) + 400]
        b = bound.search(win)
        block = win[:b.start()] if b else win
        pm = pin_re.search(block)
        out.setdefault(store, {'pincode': pm.group(1) if pm else '',
                               'city': city})
    return out

# Channels where the store PO (External Doc No) is a SAFE dedup key — i.e. unique
# per store. Other MT channels can reuse the same PO number across different
# locations, so ext-doc dedup would wrongly skip valid orders there.
# Do NOT make this uniform — add a channel only after confirming its PO is unique.
#   HG — store PO unique per store.
#   HB — Health & Beauty (Dabur): External Doc = the store PO (e.g. 4600336371),
#        one PO per DC/store delivery, never reused → safe to dedup on it.
_EXTDOC_DEDUP_CHANNELS = {'HG', 'HB'}


def line_key(po: str, item_no: str, ean: str) -> str:
    """Stable per-line identity used by the PO-flow review (Exclude decisions).
    Uses the store PO number (SO numbers don't exist until confirm)."""
    return f"{po}|{item_no}|{ean}"


# ── Tester requirement file (SELECTIVE tester generation) ─────────────────
# The operator drops a tester-requirement sheet alongside the PO files. Any
# layout with a SKU column + a Store/Code column works (header row auto-found).
# Each (Store code, SKU code) present becomes tester-eligible → the engine mints
# a tester SO (SO/<ch>/TT/…, Ext Doc 'TESTERS', qty 1, unit price = channel
# tester price e.g. HG 0.54). See [[ss-web-bridge-pattern]].
_TESTER_SKU_ALIASES = {'sku_code', 'sku', 'skucode', 'sku code'}
_TESTER_STORE_ALIASES = {'store code', 'code', 'location_code', 'loc_code',
                         'store_code', 'location code'}
_TESTER_FLAG_ALIASES = {'tester req', 'tester', 'tester_req', 'testerreq',
                        'tester req '}


def _cid(v) -> str:
    """Coerce an id the way the engine does — str, stripped, no trailing '.0'."""
    s = str(v if v is not None else '').strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


def is_tester_file(path) -> bool:
    """True if any sheet has a header row carrying a SKU column + a tester
    column — used to auto-separate the tester sheet from the PO file(s)."""
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
            if (cells & _TESTER_SKU_ALIASES) and (cells & _TESTER_FLAG_ALIASES):
                return True
    return False


def build_tester_dump(path, eng):
    """Read the tester-requirement file (any sheet/layout with SKU + Store/Code
    columns) → the engine's ``TesterDump`` with ``eligible_keys`` =
    {(store_code, sku_code)}. Presence in the sheet = eligible (a 'Tester Req'
    column, if present, must be truthy). Never raises."""
    import pandas as pd
    dump = eng.TesterDump(source_path=path, source_name=os.path.basename(str(path)))
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:  # noqa: BLE001
        dump.add_finding('error', f"cannot open tester file: {e}")
        return dump
    for sh in xl.sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name=sh, header=None, dtype=str)
        except Exception:  # noqa: BLE001
            continue
        hdr_idx = None
        for i in range(min(8, len(raw))):
            cells = {str(c).strip().lower() for c in raw.iloc[i] if c is not None}
            if (cells & _TESTER_SKU_ALIASES) and (cells & _TESTER_STORE_ALIASES):
                hdr_idx = i
                break
        if hdr_idx is None:
            continue
        low = [str(c).strip().lower() for c in raw.iloc[hdr_idx]]

        def _col(aliases, low=low):
            return next((j for j, h in enumerate(low) if h in aliases), None)
        cs, ck, cf = (_col(_TESTER_STORE_ALIASES), _col(_TESTER_SKU_ALIASES),
                      _col(_TESTER_FLAG_ALIASES))
        if cs is None or ck is None:
            continue
        for i in range(hdr_idx + 1, len(raw)):
            store, sku = _cid(raw.iat[i, cs]), _cid(raw.iat[i, ck])
            if not store or not sku or store.lower() == 'nan' or sku.lower() == 'nan':
                continue
            if cf is not None and _cid(raw.iat[i, cf]).lower() in ('', '0', 'nan', 'none'):
                continue
            dump.eligible_keys.add((store, sku))
            dump.rows_loaded += 1
    if not dump.eligible_keys:
        dump.add_finding('warn', 'no (Store code, SKU code) tester keys found.')
    return dump


def channel_choices() -> list[tuple[str, str]]:
    """``[(code, display_name)]`` for the web-enabled MT channels."""
    eng = _engine()
    return [(c, eng.CHANNELS[c].display_name)
            for c in WEB_CHANNELS if c in eng.CHANNELS]


def warehouse_choices() -> list[str]:
    return list(_engine().WAREHOUSES.keys())


def default_warehouse() -> str:
    eng = _engine()
    return getattr(eng, 'DEFAULT_WAREHOUSE', 'AHD')


def _resolve_master_path():
    """Mirror the desktop's master resolution: saved ``master_path`` in
    ``mt_select_config.json`` if it exists, else the bundled ``MT_Masters.xlsx``."""
    eng = _engine()
    cfg = eng.load_config()
    saved = cfg.get('master_path')
    if saved and Path(saved).exists():
        return Path(saved)
    return eng.get_masters_path()


class MTProcessor:
    """Run one MT-Select channel headlessly and build a web summary, matching
    the desktop ``_do_process`` path exactly."""

    def __init__(self, channel_code: str, po_paths, warehouse: str | None = None,
                 tester_file=None):
        self.channel_code = channel_code
        self.po_paths = [Path(p) for p in (po_paths or [])]
        self.warehouse = warehouse or default_warehouse()
        self.tester_file = Path(tester_file) if tester_file else None
        self.report = ''
        self.output_path = None
        self.skipped_dups = []
        self.notes: list[str] = []          # never-silent info (requirements, PDF x-check)
        self._pdf_paths: list = []          # cross-check PDFs (Excel channels only)
        self._pdf_po_dates: dict = {}       # {store PO: iso po_date} from the PDF
        self._ls_tester = None              # TesterResult (store-grouped LS testers)

    def _load(self):
        """Masters + parsed batch. NO side effects — no SO numbers burned, no
        workbook, no DB. Safe to call for preview. Raises on a fatal problem."""
        eng = _engine()
        if self.channel_code not in eng.CHANNELS:
            raise ValueError(f"Unknown channel '{self.channel_code}'.")
        if not self.po_paths:
            raise ValueError('No PO file uploaded.')
        channel = eng.CHANNELS[self.channel_code]
        bundle = eng.load_all_masters(_resolve_master_path())
        master_errs = [m for lvl, m in getattr(bundle, 'findings', [])
                       if lvl == 'error']
        if master_errs:
            raise ValueError('Masters load failed: ' + '; '.join(master_errs))
        # SHIP-TO FROM DB — single source of truth. The engine built
        # bundle.ship_to_lookup from the Excel 'Ship-To B2B' sheet; swap in the
        # DB-backed mapping (ship_to_mapping) so MT resolves ship-to from the DB
        # the web keeps correct/verified — engine resolution logic untouched,
        # only its data source. Falls back to the Excel lookup if the DB is empty.
        db_note = self._apply_db_shipto(eng, bundle)
        db_note += self._apply_db_channel_master(eng, bundle, channel)

        # ── Requirements note (never-silent: what this channel demands) ──
        req = channel_requirements(channel.code)
        if req and not any('requires' in n for n in self.notes):
            self.notes.append(
                f"{channel.display_name} requires: {req['required']}"
                + (f" Optional: {req['optional']}" if req.get('optional') else ''))

        # ── Split cross-check PDFs from the engine inputs (Excel channels only;
        #    PDF-native channels like NT/LL feed PDFs straight to the engine). ──
        engine_paths = list(self.po_paths)
        if not getattr(channel, 'pdf_parser', None):
            self._pdf_paths = [p for p in self.po_paths
                               if str(p).lower().endswith('.pdf')]
            engine_paths = [p for p in self.po_paths
                            if not str(p).lower().endswith('.pdf')]

        # ── Reliance: pre-normalize the raw Excel (inject inc-GST 'Value'). ──
        if channel.code == 'RL':
            engine_paths = [
                Path(_normalize_reliance_excel(p))
                if str(p).lower().endswith(('.xlsx', '.xls')) else p
                for p in engine_paths]

        # ── Metro / Reliance Smart Bazaar: identical 'Purchase Orders' schema —
        #    read that sheet, clean, + supply-margin note. ──
        if channel.code in ('MET', 'RSB'):
            norm = []
            for p in engine_paths:
                if str(p).lower().endswith(('.xlsx', '.xls')):
                    np, mnote = _normalize_metro_excel(p)
                    norm.append(Path(np))
                    if mnote:
                        self.notes.append(mnote)
                else:
                    norm.append(p)
            engine_paths = norm

        # ── Lulu: read the new tabular Excel → LL columns (PDF path OFF, see
        #    LULU_EXCEL_MODE). Reversible; the frozen PDF parser stays intact. ──
        if channel.code == 'LL' and LULU_EXCEL_MODE:
            norm = []
            for p in engine_paths:
                if str(p).lower().endswith(('.xlsx', '.xls')):
                    np, lnote = _normalize_lulu_excel(p)
                    norm.append(Path(np))
                    if lnote:
                        self.notes.append(lnote)
                else:
                    norm.append(p)
            engine_paths = norm

        # ── Lifestyle: read the .xlsb, map Plant ID→ship-to, serial→date, note. ──
        if channel.code == 'LS':
            norm = []
            for p in engine_paths:
                if str(p).lower().endswith(('.xlsb', '.xlsx', '.xls')):
                    np, lnotes = _normalize_lifestyle_excel(p)
                    norm.append(Path(np))
                    self.notes.extend(lnotes)
                else:
                    norm.append(p)
            engine_paths = norm

        # ── H&B: read the .xlsb, serial→date; store key = numeric Site code
        #    (exact match to party='h&b' Del Location). Mapping-only. ──
        if channel.code == 'HB':
            norm = []
            for p in engine_paths:
                if str(p).lower().endswith(('.xlsb', '.xlsx', '.xls')):
                    np, hnotes = _normalize_hb_excel(p)
                    norm.append(Path(np))
                    self.notes.extend(hnotes)
                else:
                    norm.append(p)
            engine_paths = norm

        # ── Manash (Purplle offline): tab-separated '.XLS' → clean .xlsx; Address
        #    is the ship-to lookup key (exact). Mapping-only (no price check). ──
        if channel.code == 'PPL':
            norm = []
            for p in engine_paths:
                if str(p).lower().endswith(('.xls', '.xlsx', '.csv', '.txt')):
                    np, mnotes = _normalize_manash_excel(p)
                    norm.append(Path(np))
                    self.notes.extend(mnotes)
                else:
                    norm.append(p)
            engine_paths = norm

        buf = io.StringIO()
        with redirect_stdout(buf):
            batch = eng.read_channel_csv_batch(
                engine_paths, channel, bundle, store_override='')
        self.report = db_note + buf.getvalue()

        # ── Naturals: recover an UNRESOLVED ship-to from the PDF filename city. ──
        if channel.code == 'NT':
            self._fix_naturals_city(eng, batch, channel, bundle)

        # ── Optional PDF cross-check (address vs D365 + PO totals + PO date). ──
        if channel.code == 'RL':
            self._reliance_crosscheck(batch)

        # ── Lifestyle: condense the noisy "one PO across many stores" spam and
        #    (if a PDF was uploaded, and the feature is ON) run the visible
        #    address cross-check. ──
        if channel.code == 'LS':
            self._condense_ls_multistore_warnings(batch)
            if self._pdf_paths and LS_PDF_VERIFICATION:
                self._lifestyle_crosscheck(batch, channel)
        return eng, channel, batch

    def _fix_naturals_city(self, eng, batch, channel, bundle) -> None:
        """Naturals ship-to recovery. The frozen ``parse_naturals_pdf`` reads the
        delivery city as the word before the last 6-digit pincode — which grabs
        the STATE (e.g. 'Nadu' from 'Tamil Nadu') when a PO writes
        '<City>, Tamil Nadu <PIN>' (Tirupur does; Bengaluru/Erode don't). The PDF
        FILENAME ('Renee PO no NNN - <City>.pdf') is the reliable city, so when a
        PO's ship-to came out BLANK we re-derive the city from the filename and
        re-resolve. ADDITIVE — runs only on unresolved POs, never disturbs a good
        resolution; if the re-resolve still fails the original warning stands."""
        import re as _re
        for pf in getattr(batch, 'po_files', []):
            if getattr(pf, 'ship_to', ''):          # already resolved fine
                continue
            m = _re.search(r'-\s*([A-Za-z][A-Za-z .&]+?)\s*\.(?:pdf|xlsx?)$',
                           pf.source_name or '', _re.I)
            if not m:
                continue
            city = m.group(1).strip()
            if not city or city.lower() == (pf.store_name or '').lower():
                continue
            old = pf.store_name
            pf.store_name = city
            try:
                eng._resolve_ship_to(pf, channel, bundle)
            except Exception:  # noqa: BLE001 — never break the run over recovery
                pf.store_name = old
                continue
            if getattr(pf, 'ship_to', ''):
                self.notes.append(
                    f"Naturals: PO {pf.po_no or pf.source_name} ship-to recovered "
                    f"from the filename city '{city}' — the PDF put the state "
                    f"before the pincode, so the parser read '{old}'. → {pf.ship_to}.")

    def _condense_ls_multistore_warnings(self, batch) -> None:
        """TASK A — Lifestyle only. The frozen engine appends, per PO, a warning
        that the PO "appears across different stores: <30+ store names with
        tmpXXXX_lifestyle_norm.xlsx prefixes>". For a Lifestyle replenishment PO
        this is EXPECTED (one PO legitimately spans dozens of stores) and pure
        noise. Never-silent: we DROP those verbose cross_findings and REPLACE each
        with ONE concise line — ``"LS PO <PO> spans <N> stores — normal for a
        Lifestyle replenishment PO."`` (N = distinct stores for that PO). Mirrors
        how ``SwiggyProcessor._drop_non_confirmed`` filters engine warnings. Only
        LS calls this — other channels keep the (meaningful) warning."""
        import re
        findings = list(getattr(batch, 'cross_findings', []) or [])
        pat = re.compile(r'appears across\s+different stores:\s*(.*?)\.\s*Verify',
                         re.S)
        kept: list = []
        condensed: list = []
        for lvl, msg in findings:
            m = pat.search(str(msg))
            if not m:
                kept.append((lvl, msg))
                continue
            # PO number is the token right after the engine's csv_po_col label
            # ("Order No <PO> appears across …").
            pom = re.search(r'^\s*\S.*?\b(\S+)\s+appears across', str(msg))
            po = pom.group(1) if pom else '?'
            # N = distinct stores = distinct "(store)" groups in the detail list.
            stores = set(re.findall(r'\(([^()]+)\)', m.group(1)))
            n = len(stores) or (m.group(1).count(',') + 1)
            condensed.append(('info',
                f"LS PO {po} spans {n} stores — normal for a Lifestyle "
                f"replenishment PO."))
        # De-dup the condensed lines (one per PO), keep order; drop the spam.
        seen = set()
        uniq = [c for c in condensed
                if not (c[1] in seen or seen.add(c[1]))]
        batch.cross_findings = kept + uniq

    def _reliance_crosscheck(self, batch) -> None:
        """Cross-check each PO against its PDF (address vs D365 ship-to pincode,
        inc-GST total, PO date). Never silent: a note per PO OK, a warning per
        drift; if no PDF, a note that we ran on the Excel alone. Also stashes the
        PDF PO Date per store-PO for backfill after confirm."""
        req = channel_requirements('RL') or {}
        if not self._pdf_paths:
            self.notes.append("No PO PDF uploaded — " + req.get('if_absent', ''))
            return
        import datetime as _dt
        pdfs: dict = {}
        for p in self._pdf_paths:
            try:
                d = _parse_reliance_pdf(p)
                if d['po']:
                    pdfs[d['po']] = d
            except Exception as e:  # noqa: BLE001
                self.notes.append(f"PDF {Path(p).name}: parse failed ({type(e).__name__}).")
        self.notes.append(
            f"PDF cross-check: {len(pdfs)} PO PDF(s) read — verifying address, "
            f"totals & PO date against the Excel.")
        for pf in batch.po_files:
            d = pdfs.get(str(pf.po_no))
            if not d:
                self.notes.append(f"PO {pf.po_no}: no matching PDF — Excel-only for this PO.")
                continue
            # value (inc GST)
            exc = float(getattr(pf, 'input_po_value_total', 0) or 0)
            if d['value'] is not None and abs(exc - d['value']) > max(2.0, exc * 0.005):
                batch.cross_findings = list(getattr(batch, 'cross_findings', [])) + [(
                    'warn', f"PO {pf.po_no}: value Excel ₹{exc:.2f} vs PDF ₹{d['value']:.2f} "
                    f"— differs, verify.")]
            # delivery pincode (address confirmation vs the D365 ship-to)
            pcode = str(getattr(pf.ship_to_entry, 'postcode', '') or '') if pf.ship_to_entry else ''
            if d['pin'] and pcode and d['pin'] != pcode:
                batch.cross_findings = list(getattr(batch, 'cross_findings', [])) + [(
                    'warn', f"PO {pf.po_no} (Site {d['site']}): PDF delivery pincode "
                    f"{d['pin']} ≠ mapped ship-to {pf.ship_to} pincode {pcode} — "
                    f"WRONG ship-to? verify address.")]
            else:
                self.notes.append(
                    f"PO {pf.po_no}: Site {d['site']} → {pf.ship_to} "
                    f"(pin {pcode or '?'}) ✓ address+value match PDF.")
            # stash PO date (DD.MM.YYYY → iso) for backfill
            try:
                dd, mm, yy = d['po_date'].split('.')
                self._pdf_po_dates[str(pf.po_no)] = _dt.date(int(yy), int(mm), int(dd)).isoformat()
            except Exception:  # noqa: BLE001
                pass

    def _lifestyle_crosscheck(self, batch, channel) -> None:
        """TASK B — Lifestyle only, and only when a PO PDF is uploaded. Parse each
        LS PDF (per-store ``DELIVERY <storeNo> - <address…>`` blocks) and compare
        the PDF delivery pincode + city against our resolved DB ship-to (party
        'LS'). Builds a channel-agnostic ``batch.verification`` dict via the shared
        :mod:`online_b2b.services.verification` scaffold (the dedicated
        verification page renders it) plus a concise summary note (never-silent).
        Models :meth:`_reliance_crosscheck` but structured (not just warnings),
        because a single LS PO spans ~90 stores and per-store rows read far better
        in a table than in a warnings list. This method ONLY produces the generic
        dict — no route/UI code here; the next marketplace just does the same."""
        import re
        # 1) Parse every PDF → {store number: {'pincode', 'city'}}.
        pdf_by_store: dict = {}
        parsed_files = 0
        for p in self._pdf_paths:
            try:
                one = _parse_lifestyle_pdf(p)
            except Exception as e:  # noqa: BLE001 — never block; name the file
                self.notes.append(
                    f"LS PDF {Path(p).name}: parse failed ({type(e).__name__}).")
                continue
            parsed_files += 1
            for store, info in one.items():
                pdf_by_store.setdefault(store, info)   # first PDF wins per store

        # 2) Reverse map: del_location → store number (the .xlsb Plant ID). The
        #    batch resolves each pf to a Del Location; the PDF keys on the bare
        #    store number, so we bridge via the ship-to mapping.
        dl_to_store: dict = {}
        for store, (used, _dls) in _lifestyle_store_map().items():
            dl_to_store.setdefault(str(used), store)

        def _digits(v) -> str:
            return re.sub(r'\D', '', str(v or ''))

        findings: list = []
        ok = mism = nopdf = 0
        for pf in batch.po_files:
            if getattr(pf, 'has_hard_errors', False):
                continue
            entry = getattr(pf, 'ship_to_entry', None)
            del_loc = str(getattr(entry, 'del_location', '')
                          or getattr(pf, 'store_name', '') or '')
            store = dl_to_store.get(del_loc) or _digits(getattr(pf, 'store_name', ''))
            if not store:
                continue
            our_pin = _digits(getattr(entry, 'postcode', '')) if entry else ''
            our_city = str(getattr(entry, 'city', '') or '') if entry else ''
            ship_to = str(getattr(pf, 'ship_to', '') or '')
            pdf = pdf_by_store.get(store)
            if not pdf:
                nopdf += 1
                findings.append({
                    'store': store, 'ship_to': ship_to,
                    'our_pincode': our_pin, 'pdf_pincode': '',
                    'our_city': our_city, 'pdf_city': '',
                    'match': 'NO_PDF',
                    'detail': 'Store not found in the uploaded PDF(s).'})
                continue
            pdf_pin = _digits(pdf.get('pincode', ''))
            pdf_city = str(pdf.get('city', '') or '')
            # City: case-insensitive substring either way — ADVISORY only. The PDF
            # 'city' is the last word of the free-form delivery line and is often a
            # locality/floor fragment ('Sector', 'SF', 'West'), so it is NOT
            # reliable enough to fail a row on its own. MATCH is decided by the
            # pincode (the load-bearing signal); a city difference is surfaced in
            # the detail text for the operator to eyeball.
            oc, pc = our_city.strip().lower(), pdf_city.strip().lower()
            city_ok = (not oc or not pc) or (oc in pc) or (pc in oc)
            if not pdf_pin:
                # Block had no delivery pincode in the PDF — can't verify (never a
                # silent pass, but not a false mismatch either).
                nopdf += 1
                findings.append({
                    'store': store, 'ship_to': ship_to,
                    'our_pincode': our_pin, 'pdf_pincode': '',
                    'our_city': our_city, 'pdf_city': pdf_city,
                    'match': 'NO_PDF',
                    'detail': 'No delivery pincode in the PDF block — not verified.'})
                continue
            pin_ok = bool(our_pin) and our_pin == pdf_pin
            match = 'OK' if pin_ok else 'MISMATCH'
            if match == 'OK':
                ok += 1
                detail = ('Pincode matches the PDF.' if city_ok else
                          f"Pincode matches; city differs (ours '{our_city}' vs "
                          f"PDF '{pdf_city}') — advisory, verify if unsure.")
            else:
                mism += 1
                detail = (f"pincode our {our_pin or '?'} ≠ PDF {pdf_pin} — "
                          f"WRONG ship-to? verify address.")
            findings.append({
                'store': store, 'ship_to': ship_to,
                'our_pincode': our_pin, 'pdf_pincode': pdf_pin,
                'our_city': our_city, 'pdf_city': pdf_city,
                'match': match, 'detail': detail})

        # Hand the findings to the shared, channel-agnostic verification scaffold
        # (online_b2b.services.verification). LS only produces the generic
        # `verification` dict — the page/route/template that renders it are shared,
        # so the NEXT marketplace just builds the same dict (zero new UI code).
        from online_b2b.services import verification as vfy
        batch.verification = vfy.build(
            title='PDF Address Verification',
            subtitle=('Delivery pincode on each store PO PDF vs our D365 ship-to '
                      '(party LS). City is advisory; the pincode decides the match.'),
            columns=[
                {'key': 'store', 'label': 'Store', 'mono': True},
                {'key': 'ship_to', 'label': 'Ship-to', 'mono': True},
                {'key': 'our_pincode', 'label': 'Our Pincode', 'align': 'r', 'mono': True},
                {'key': 'pdf_pincode', 'label': 'PDF Pincode', 'align': 'r', 'mono': True},
                {'key': 'city', 'label': 'City (ours / PDF)'},
                {'key': 'match', 'label': 'Match', 'kind': 'match'},
            ],
            rows=[{
                'store': f['store'], 'ship_to': f['ship_to'] or '—',
                'our_pincode': f['our_pincode'] or '—',
                'pdf_pincode': f['pdf_pincode'] or '—',
                'city': f"{f['our_city'] or '—'} / {f['pdf_city'] or '—'}",
                'match': f['match'], 'detail': f['detail'],
            } for f in findings],
            match_key='match',
            source=f"{parsed_files} PDF(s) read for {channel.display_name}",
        )
        # Never-silent summary note.
        self.notes.append(
            f"PDF address verification: {parsed_files} PDF(s) read · "
            f"{len(findings)} store(s) checked — {ok} OK, {mism} mismatch(es)"
            + (f", {nopdf} not in PDF" if nopdf else '') + '.')

    @staticmethod
    def _apply_db_channel_master(eng, bundle, channel) -> str:
        """Merge the DB ``channel_sku_map`` (SKU→EAN) into the engine's
        ``bundle.channel_masters[code]`` so lookup_via='SKU' channels (HG) resolve
        SKUs to EANs from the DB — the web keeps it current (seeded from the Dec-25
        HG Master + bin-content matches). DB wins over the Excel HG Master; entries
        not in the DB keep the Excel value. No-op for EAN-lookup channels or when
        the DB is empty/unreachable — resolution is never blocked."""
        if getattr(channel, 'lookup_via', 'SKU') != 'SKU':
            return ''
        code = channel.code
        try:
            from online_b2b.services.order_db import _conn
            cm = bundle.channel_masters.setdefault(code, {})
            added = filled = 0
            with _conn() as (cur, d):
                ph = d['ph']
                cur.execute(
                    f"SELECT sku_code, ean, item_no FROM channel_sku_map "
                    f"WHERE channel={ph} AND ean IS NOT NULL AND ean <> ''", (code,))
                for sku, ean, _item in cur.fetchall():
                    sku, ean = str(sku), str(ean)
                    ex = cm.get(sku)
                    if ex is None:
                        cm[sku] = eng.ChannelMasterEntry(
                            sku_code=sku, sku_name='', enn_code=ean, status='Active')
                        added += 1
                    elif not getattr(ex, 'enn_code', None):
                        cm[sku] = eng.ChannelMasterEntry(
                            sku_code=sku, sku_name=getattr(ex, 'sku_name', ''),
                            enn_code=ean, status=getattr(ex, 'status', 'Active') or 'Active')
                        filled += 1
            if added or filled:
                return f"[hg-master] DB SKU->EAN: filled {filled}, added {added}\n"
            return ''
        except Exception as e:  # noqa: BLE001 — never block SO gen
            return f"[hg-master] DB skip ({type(e).__name__}: {e})\n"

    @staticmethod
    def _apply_db_shipto(eng, bundle) -> str:
        """Replace ``bundle.ship_to_lookup`` (built from Excel) with the
        DB-backed ``ship_to_mapping`` so ALL MT channels resolve ship-to from the
        DB. Keyed exactly like the engine — ``(party, del_location)`` → the
        engine's ``ShipToEntry`` (incl. full address for the Summary sheet). No-op
        (keeps the Excel lookup) if the DB is empty or unreachable — SO
        generation is never blocked."""
        try:
            from online_b2b.services.order_db import _conn
            lookup = {}
            with _conn() as (cur, _d):
                cur.execute(
                    "SELECT party, del_location, cust_no, ship_to, name, "
                    "address, address2, postcode, city FROM ship_to_mapping")
                for (party, dl, cust, st, name, addr,
                     addr2, pc, city) in cur.fetchall():
                    party = str(party or '').strip()
                    dl = str(dl or '').strip()
                    if not party or not dl:
                        continue
                    key = (party, dl)
                    if key in lookup:          # first row wins (mirror the engine)
                        continue
                    lookup[key] = eng.ShipToEntry(
                        party=party, del_location=dl,
                        cust_no=str(cust or ''), ship_to=str(st or ''),
                        name=str(name or ''), address=str(addr or ''),
                        address_2=str(addr2 or ''), postcode=str(pc or ''),
                        city=str(city or ''))
            if lookup:
                bundle.ship_to_lookup = lookup
                return f"[ship-to] resolved from DB ({len(lookup)} entries)\n"
            return "[ship-to] DB empty — using Excel Ship-To B2B\n"
        except Exception as e:  # noqa: BLE001 — never block SO gen
            return (f"[ship-to] DB unavailable ({type(e).__name__}: {e}) — "
                    f"using Excel Ship-To B2B\n")

    # ── phase 1: preview (parse + validate, NO writes) ──────────────────
    def preview(self) -> dict:
        """Parse + resolve + validate only. No SO numbers assigned, no workbook,
        no DB — mirrors the online ``preview`` so the operator can verify first."""
        try:
            eng, channel, batch = self._load()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e)}
        return self._summary(batch, channel, recorded=None, phase='preview')

    def _recorded_ext_docs(self) -> set:
        """Store POs (External Doc No) already recorded — so a re-uploaded dump is
        detected and never minted twice. **HG only, and scoped to the HG channel**:
        HG's store PO is unique per store, so it's a safe dedup key. Other MT
        channels can legitimately reuse the same PO number at different locations,
        so ext-doc dedup must NOT apply to them (return empty). Extend
        ``_EXTDOC_DEDUP_CHANNELS`` deliberately — never make this uniform."""
        if self.channel_code not in _EXTDOC_DEDUP_CHANNELS:
            return set()
        try:
            label = _engine().CHANNELS[self.channel_code].display_name
        except Exception:  # noqa: BLE001
            label = 'Health & Glow'
        try:
            from online_b2b.services.order_db import _conn
            with _conn() as (cur, d):
                ph = d['ph']
                cur.execute(
                    f"SELECT external_doc FROM order_headers WHERE marketplace='MT' "
                    f"AND marketplace_label={ph} AND external_doc IS NOT NULL "
                    f"AND external_doc <> ''", (label,))
                return {str(r[0]) for r in cur.fetchall()}
        except Exception:  # noqa: BLE001
            return set()

    def _is_ls_tester(self) -> bool:
        """True when the supplied tester sheet is the LS store-grouped layout
        (STORE + EAN + Tester Req) rather than the HG (Store, SKU) dump."""
        from . import tester as tester_svc
        return bool(self.tester_file) and tester_svc.is_ls_tester_file(self.tester_file)

    def ls_tester_preview(self) -> dict | None:
        """Count the store-grouped LS tester SOs/lines that WOULD be generated
        (no writes, no counter burn — counter_start is irrelevant to the count).
        None when the supplied sheet is not the LS tester layout."""
        if not self._is_ls_tester():
            return None
        from . import tester as tester_svc
        try:
            eng, channel, _batch = self._load()
        except Exception:  # noqa: BLE001
            return None
        res, _nxt = tester_svc.build_ls_testers(
            self.tester_file, channel, counter_start=0,
            location_code=eng.WAREHOUSES.get(self.warehouse, 'PICK'))
        return {'sos': res.stores, 'lines': res.line_count, 'value': res.value,
                'price': res.price, 'warnings': res.warnings}

    def _generate_ls_testers(self, eng, channel, warehouse_code) -> None:
        """Mint store-grouped LS tester SOs (SO/<CH>/TT/<counter>) from the SAME
        daily counter the regular run just used. Called INSIDE ``confirm`` right
        after ``assign_so_numbers`` has burned the regular block and persisted
        ``next_counter`` to ``mt_select_seq.json``. Continues from that counter,
        then persists the advanced value so tester numbers never collide with
        the regular run or a later batch. Stores the TesterResult on
        ``self._ls_tester`` for the workbook + DB record; every unresolved line /
        missing ship-to is a named warning (never silent)."""
        import datetime as _dt

        from . import tester as tester_svc
        state = eng.load_seq_state()
        ch_state = state.get(channel.code, {})
        today_iso = _dt.date.today().isoformat()
        # assign_so_numbers just ran, so ch_state is today's with next_counter
        # past the regular block. Defensive: if somehow missing, start at today's
        # DDMMYY base (matches the engine's reset rule).
        if ch_state.get('date') != today_iso or 'next_counter' not in ch_state:
            ch_state = {'date': today_iso,
                        'next_counter': int(_dt.date.today().strftime('%d%m%y'))}
        counter_start = ch_state['next_counter']
        res, nxt = tester_svc.build_ls_testers(
            self.tester_file, channel, run_date=_dt.date.today(),
            counter_start=counter_start, location_code=warehouse_code)
        self._ls_tester = res
        for w in res.warnings:
            self.notes.append(w)
        if res.headers:
            ch_state['next_counter'] = nxt
            state[channel.code] = ch_state
            eng.save_seq_state(state)
            self.report += (
                f"[tester] {res.line_count} tester line(s) @ {res.price} across "
                f"{res.stores} store(s) — SO/{channel.code}/TT/"
                f"{counter_start:06d}..{nxt - 1:06d}.\n")

    def tester_preview(self) -> dict | None:
        """Count tester SOs/lines that WOULD be generated (no writes) so the
        review page can show them. None when no tester file is supplied."""
        if not self.tester_file:
            return None
        if self._is_ls_tester():
            return self.ls_tester_preview()
        try:
            eng, channel, batch = self._load()
        except Exception:  # noqa: BLE001
            return None
        price = getattr(channel, 'tester_unit_price', None)
        if price is None:
            return {'sos': 0, 'lines': 0, 'value': 0.0,
                    'error': f"{channel.display_name} has no tester price configured."}
        dump = build_tester_dump(self.tester_file, eng)
        if dump.has_hard_errors:
            return {'sos': 0, 'lines': 0, 'value': 0.0,
                    'error': '; '.join(m for lvl, m in dump.findings if lvl == 'error')}
        sos = lines = 0
        for pf in batch.po_files:
            if getattr(pf, 'has_hard_errors', False):
                continue
            elig = [ln for ln in pf.lines
                    if getattr(ln, 'item_no', None) and getattr(ln, 'status', '') != 'SKIP'
                    and dump.is_eligible(pf.location_code, getattr(ln, 'sku_code', ''))]
            if elig:
                sos += 1
                lines += len(elig)
        return {'sos': sos, 'lines': lines, 'value': round(lines * price, 2),
                'price': price, 'keys': len(dump.eligible_keys)}

    # ── phase 2: confirm (assign + write + record to renee_orders) ──────
    def confirm(self, exclude_keys=None) -> dict:
        """Assign SO numbers (burns the ``mt_select_seq.json`` counter ONCE),
        write the 6-sheet workbook, and record order headers into the shared
        ``renee_orders`` DB (segment Offline) via the desktop's own
        ``record_offline_batch`` — so SS appears on the online dashboard.

        ``exclude_keys`` (a set of :func:`line_key` values from the PO-flow
        review's Exclude decisions) drops those lines BEFORE SO numbers are
        assigned, so an operator-excluded line never reaches the workbook or the
        DB. A PO left with no lines is dropped entirely (no empty SO)."""
        try:
            eng, channel, batch = self._load()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e)}
        if exclude_keys:
            for pf in batch.po_files:
                pf.lines = [
                    ln for ln in pf.lines
                    if line_key(pf.po_no, str(getattr(ln, 'item_no', '') or ''),
                                str(getattr(ln, 'ean', '') or '')) not in exclude_keys]
            batch.po_files = [pf for pf in batch.po_files if pf.lines]
        # De-dup by External Doc No (store PO): drop POs already recorded for MT
        # (a re-uploaded dump) so we never mint duplicate SOs. Reported as skipped.
        self.skipped_dups = []
        already = self._recorded_ext_docs()
        if already:
            dups = [pf for pf in batch.po_files if str(pf.po_no) in already]
            if dups:
                self.skipped_dups = [str(pf.po_no) for pf in dups]
                batch.po_files = [pf for pf in batch.po_files
                                  if str(pf.po_no) not in already]
                self.report += (f"[dedup] {len(dups)} PO(s) already uploaded "
                                f"(External Doc): {', '.join(self.skipped_dups)}\n")
        if not batch.po_files:
            return {'ok': False,
                    'error': ('All PO(s) in this dump are already uploaded '
                              f"(External Doc): {', '.join(self.skipped_dups)}."),
                    'skipped_dups': self.skipped_dups, 'report': self.report}
        warehouse_code = eng.WAREHOUSES.get(self.warehouse, 'PICK')
        buf = io.StringIO()
        # Testers: SELECTIVE mode when the operator supplied a tester-requirement
        # file (only its (Store, SKU) get a tester SO — Ext Doc 'TESTERS', qty 1,
        # unit price = channel.tester_unit_price e.g. HG 0.54). Channels with a
        # tester-qty divisor (e.g. Off-Institutional) always pair testers.
        # The LS store-grouped tester sheet (STORE + EAN + Tester Req) is a
        # SEPARATE path from the HG per-PO (Store, SKU) dump: LS testers are
        # grouped by store into their own SO/LS/TT/<counter> SOs generated AFTER
        # the regular block (see _generate_ls_testers). So route it away from the
        # HG dump here.
        ls_tester = self._is_ls_tester()
        tester_dump = None
        if (not ls_tester and self.tester_file
                and getattr(channel, 'tester_unit_price', None) is not None):
            tester_dump = build_tester_dump(self.tester_file, eng)
            for lvl, msg in tester_dump.findings:
                self.report += f"[tester dump] {lvl}: {msg}\n"
        gen_testers = ((tester_dump is not None and not tester_dump.has_hard_errors)
                       or getattr(channel, 'tester_qty_divisor', None) is not None)
        try:
            with redirect_stdout(buf):
                eng.assign_so_numbers(batch, channel,
                                      generate_testers=gen_testers,
                                      tester_dump=tester_dump)
                # LS: mint store-grouped tester SOs from the SAME daily counter,
                # continuing AFTER the regular block just burned (SO/LS/TT/...).
                if ls_tester:
                    self._generate_ls_testers(eng, channel, warehouse_code)
                eng.print_batch_report(batch)
                if any(pf.so_number for pf in batch.po_files):
                    self.output_path = eng.write_so_workbook(
                        batch, channel, warehouse_code,
                        output_path=None, add_non_stock=False)
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': f"{type(e).__name__}: {e}",
                    'report': self.report + buf.getvalue()}
        self.report += buf.getvalue()

        # ── Uniform workbook: re-render the download so its sheet structure
        #    MATCHES the Online B2B workbook (Headers/Lines/Summary/Tracker/
        #    Validation/Rules & Exceptions/Warnings/Raw Data/SKU Summary). The
        #    frozen engine already wrote its 6-sheet file at ``self.output_path``;
        #    we overwrite it in place via the ONLINE SOExporter path (Approach A).
        #    Best-effort — a failure keeps the frozen 6-sheet workbook, never
        #    losing the download (golden rule: never silent — logged to report). ──
        if self.output_path:
            try:
                from . import mt_workbook
                mt_workbook.write_unified_workbook(
                    batch, channel, self.warehouse, warehouse_code,
                    channel.display_name, self.output_path,
                    notes=self.notes, warnings=self._workbook_warnings(batch),
                    testers=self._ls_tester)
            except Exception as e:  # noqa: BLE001 — fall back to the 6-sheet file
                self.report += (f"\n[unified workbook skipped] "
                                f"{type(e).__name__}: {e} — served the frozen "
                                f"6-sheet workbook instead.")

        # Record into the same DB the online PO tool uses (soft-fails, never
        # blocks SO generation — the workbook is already written).
        recorded = None
        if self.output_path:
            recorded = eng.record_offline_batch(
                batch, channel, self.warehouse, str(self.output_path))
            # Also record the web-owned line-item audit (order_lines) so SS gets
            # the same Line Items / Issues view as the online marketplaces.
            if recorded and recorded.get('recorded') and recorded.get('run_id'):
                try:
                    self._record_lines(channel, batch, recorded['run_id'],
                                       self.output_path.name)
                except Exception as e:  # noqa: BLE001 — never block on the audit
                    self.report += f"\n[line audit skipped] {type(e).__name__}: {e}"
                # Testers: record the store-grouped tester SOs under the SAME
                # run_id (order_headers + order_lines), tagged by the SO/LS/TT/
                # number + external_doc 'TESTER-<store>' (no is_tester column).
                if self._ls_tester and self._ls_tester.headers:
                    try:
                        self._record_testers(channel, recorded['run_id'],
                                             self.output_path.name)
                    except Exception as e:  # noqa: BLE001 — never block
                        self.report += f"\n[tester record skipped] {type(e).__name__}: {e}"
                # Stamp External Doc No (store PO) onto each recorded header so a
                # re-uploaded dump is de-duplicated next time. (order_headers.po
                # holds the SO number; external_doc holds the store PO.)
                try:
                    from online_b2b.services.order_db import _conn
                    with _conn() as (cur, d):
                        ph = d['ph']
                        for pf in batch.po_files:
                            if pf.so_number and pf.po_no:
                                cur.execute(
                                    f"UPDATE order_headers SET external_doc={ph} "
                                    f"WHERE run_id={ph} AND po={ph}",
                                    (str(pf.po_no), recorded['run_id'], pf.so_number))
                        cur.connection.commit()
                except Exception as e:  # noqa: BLE001
                    self.report += f"\n[external_doc stamp skipped] {e}"
                # PO Date backfill from the PDF cross-check (Reliance Excel has
                # none). Map store PO → our SO number → set order_headers.po_date.
                if self._pdf_po_dates:
                    try:
                        from online_b2b.services import lines_store
                        by_so = {pf.so_number: {'po_date': self._pdf_po_dates[str(pf.po_no)]}
                                 for pf in batch.po_files
                                 if pf.so_number and str(pf.po_no) in self._pdf_po_dates}
                        if by_so:
                            n = lines_store.set_po_dates(recorded['run_id'], by_so,
                                                         force=True).get('updated', 0)
                            self.notes.append(f"PO Date backfilled from PDF on {n} SO(s).")
                    except Exception as e:  # noqa: BLE001
                        self.report += f"\n[po_date backfill skipped] {e}"
        return self._summary(batch, channel, recorded=recorded, phase='confirm')

    def _record_lines(self, channel, batch, run_id, output_file) -> int:
        """Map each resolved SS line → the web-owned ``order_lines`` audit (the
        SAME columns the desktop's Validation sheet shows). Reads POLine fields
        only — the engine is untouched. Vendor MRP/Landing stay blank when the SS
        file carries no cost (that's the true picture, not a silent drop)."""
        import datetime as _dt

        from online_b2b.services import lines_store
        ratio = getattr(channel, 'expected_landing_ratio', None)
        run_ts = _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        def _f(x):
            try:
                return round(float(x), 2) if x not in (None, '') else None
            except (TypeError, ValueError):
                return None

        rows = []
        for pf in batch.po_files:
            if pf.has_hard_errors or not pf.so_number:
                continue
            loc = pf.ship_to_entry.del_location if pf.ship_to_entry else ''
            for ln in pf.lines:
                item_no = str(ln.item_no or '')
                ean = str(getattr(ln, 'ean', '') or '')
                # Record real product lines even when UNRESOLVED (NOT_IN_MASTER),
                # so dropped qty is audited + surfaces on the Issues page — parity
                # with Online B2B (never-skip-silently). Only genuinely empty rows
                # (no item AND no EAN) are skipped. Unresolved → NOT_IN_MASTER +
                # EXCLUDE (they never reach the D365 dump); resolved → OK.
                if not item_no and not ean:
                    continue
                resolved = bool(item_no)
                our_mrp = _f(getattr(ln, 'items_master_mrp', None))
                # SS file carries no cost → 0 means "absent": show blank, not 0.00.
                vendor_landing = _f(getattr(ln, 'purchase_cost', None)) or None
                our_landing = (round(our_mrp * ratio, 2)
                               if ratio and our_mrp else None)
                diff = (round(vendor_landing - our_landing, 2)
                        if vendor_landing is not None and our_landing is not None
                        else None)
                rows.append({
                    'run_id': run_id, 'run_ts': run_ts,
                    'marketplace': channel.display_name,    # 'Shoppers Stop'
                    'po': pf.so_number,
                    'location': str(loc or ''),
                    'item_no': item_no,
                    'ean': ean,
                    'description': (str(getattr(ln, 'items_master_desc', None)
                                        or getattr(ln, 'sku_name', '') or ''))[:255],
                    'qty': int(getattr(ln, 'quantity', 0) or 0),
                    'order_type': 'SO',
                    'gst_code': str(getattr(ln, 'gst_code', '') or ''),
                    'unit_price': our_landing,
                    'vendor_mrp': _f(getattr(ln, 'mrp', None)) or None,
                    'our_mrp': our_mrp,
                    'vendor_landing': vendor_landing,
                    'our_landing': our_landing,
                    'vendor_cp': None, 'our_cp': None,      # SS validates on landing
                    'diff': diff,
                    'margin_pct': round(ratio * 100, 2) if ratio else None,
                    'status': 'OK' if resolved else 'NOT_IN_MASTER',
                    'exception_label': '',
                    'output_file': output_file or '',
                    'action': '' if resolved else 'EXCLUDE', 'remark': '',
                })
        if rows:
            lines_store.insert_lines(run_id, rows)
        return len(rows)

    def _record_testers(self, channel, run_id, output_file) -> int:
        """Record the store-grouped tester SOs under the SAME run_id — one
        ``order_headers`` row per tester SO (po=SO/<CH>/TT/..., external_doc=
        'TESTER-<store>', segment 'Offline') + one ``order_lines`` row per tester
        line (qty 1, unit_price = tester price). No ``is_tester`` column exists,
        so the SO/<CH>/TT/ number + external_doc are the tester marker. Reads
        ``self._ls_tester`` (built in ``_generate_ls_testers``)."""
        import datetime as _dt

        from online_b2b.services import lines_store
        from online_b2b.services.order_db import _conn
        res = self._ls_tester
        if not res or not res.headers:
            return 0
        run_ts = _dt.datetime.now()
        run_ts_s = run_ts.strftime('%Y-%m-%d %H:%M:%S')
        wh = self.warehouse or ''
        label = channel.display_name
        # Headers — direct insert (mirror gt_mass_bridge._record columns).
        with _conn() as (cur, d):
            ph = d['ph']
            hcols = ('run_id, run_ts, mode, segment, marketplace, '
                     'marketplace_label, po, location, warehouse, po_date, '
                     'exp_date, order_type, items, qty, order_value, '
                     'output_file, external_doc')
            marks = ', '.join([ph] * 17)
            for h in res.headers:
                pod = _dt.date.today()
                cur.execute(
                    f"INSERT INTO order_headers ({hcols}) VALUES ({marks})",
                    (run_id, run_ts, 'MANUAL', 'Offline', 'MT', label,
                     h['po'], h.get('location') or '', wh, pod, pod, 'SO',
                     h.get('qty') or 0, h.get('qty') or 0,
                     h.get('order_value') or 0.0, output_file or '',
                     h.get('external_doc') or f"TESTER-{h.get('store')}"))
            cur.connection.commit()
        # Lines — via the shared audit inserter (order_lines + validation).
        rows = []
        for ln in res.lines:
            rows.append({
                'run_id': run_id, 'run_ts': run_ts_s,
                'marketplace': label, 'po': ln['po'],
                'location': ln.get('location') or '',
                'item_no': str(ln.get('item_no') or ''),
                'ean': str(ln.get('ean') or ''),
                'description': (ln.get('description') or '')[:255],
                'qty': int(ln.get('qty') or 1), 'order_type': 'SO',
                'gst_code': '', 'unit_price': ln.get('unit_price'),
                'our_mrp': None, 'vendor_mrp': None,
                'our_landing': ln.get('unit_price'), 'vendor_landing': None,
                'our_cp': None, 'vendor_cp': None, 'diff': None,
                'margin_pct': None, 'status': 'OK', 'exception_label': '',
                'received_ean': None, 'action': '', 'remark': 'TESTER',
                'output_file': output_file or '',
            })
        if rows:
            lines_store.insert_lines(run_id, rows)
        self.report += (f"[tester record] {len(res.headers)} tester SO(s) + "
                        f"{len(rows)} line(s) recorded under run {run_id}.\n")
        return len(rows)

    def _workbook_warnings(self, batch) -> list:
        """Every non-fatal issue to surface on the workbook's Warnings sheet —
        cross-findings, per-file findings, and any unresolved/SKIP line (named,
        never silent) that was dropped from Lines. Same intent as the frozen
        engine's own Warnings sheet, expressed for the unified exporter."""
        out: list = []
        for lvl, msg in getattr(batch, 'cross_findings', []) or []:
            out.append(f"[{lvl}] {msg}")
        for pf in batch.po_files:
            for lvl, msg in getattr(pf, 'findings', []) or []:
                if lvl in ('error', 'warn', 'warning'):
                    out.append(f"{pf.source_name}: [{lvl}] {msg}")
            if pf.has_hard_errors or not pf.so_number:
                continue
            for ln in pf.lines:
                if not ln.item_no or ln.status == 'SKIP':
                    sku = str(getattr(ln, 'sku_code', '') or
                              getattr(ln, 'ean', '') or '?')
                    out.append(
                        f"PO {pf.po_no}: line (SKU/EAN {sku}) not written to "
                        f"Lines — {'unresolved' if not ln.item_no else 'skipped'}.")
        return list(dict.fromkeys(out))       # de-dup, keep order

    # ── web summary payload ─────────────────────────────────────────────
    def _summary(self, batch, channel, recorded=None, phase='confirm') -> dict:
        pos: list[dict] = []
        warnings: list[str] = []
        for lvl, msg in getattr(batch, 'cross_findings', []):
            warnings.append(f"[{lvl}] {msg}")

        total_qty = 0
        total_val = 0.0
        for pf in batch.po_files:
            for lvl, msg in getattr(pf, 'findings', []):
                if lvl in ('error', 'warn', 'warning'):
                    warnings.append(f"{pf.source_name}: [{lvl}] {msg}")
            qty = int(getattr(pf, 'input_qty_total', 0) or 0)
            val = float(getattr(pf, 'input_po_value_total', 0) or 0)
            total_qty += qty
            total_val += val
            pos.append({
                'file': pf.source_name,
                'po': pf.po_no,
                'store': pf.store_name,
                'so_number': pf.so_number,
                'ship_to': pf.ship_to or '',
                'cust_no': pf.cust_no or '',
                'lines': len(pf.lines),
                'qty': qty,
                'value': round(val, 2),
                'status': ('ERROR' if pf.has_hard_errors
                           else ('SO' if pf.so_number
                                 else ('READY' if phase == 'preview' else 'SKIP'))),
            })

        # "Eligible" = parsed cleanly with resolvable lines (preview) / has SO
        # number (confirm). Mirrors the desktop's any_eligible gate.
        eligible = [p for p in pos
                    if p['status'] in ('SO', 'READY') and p['lines'] > 0]
        ok = (bool(eligible) if phase == 'preview' else bool(self.output_path))
        rec_info = recorded or {}
        return {
            'ok': ok,
            'phase': phase,
            'channel': channel.display_name,
            'channel_code': channel.code,
            'warehouse': self.warehouse,
            'output_path': str(self.output_path) if self.output_path else None,
            'output_name': self.output_path.name if self.output_path else None,
            'pos': pos,
            'summary': {
                'files': len(batch.po_files),
                'sos': len(eligible),
                'errors': sum(1 for p in pos if p['status'] == 'ERROR'),
                'lines': sum(p['lines'] for p in pos),
                'qty': total_qty,
                'value': round(total_val, 2),
            },
            # DB recording outcome (confirm only).
            'recorded': bool(rec_info.get('recorded')),
            'run_id': rec_info.get('run_id'),
            'recorded_pos': rec_info.get('recorded_pos') or rec_info.get('new_orders'),
            'recorded_skipped': rec_info.get('skipped'),
            'recorded_reason': rec_info.get('reason'),
            'warnings': warnings,
            'notes': self.notes,
            'requirements': channel_requirements(channel.code),
            'report': self.report,
            'error': (None if ok else
                      ('No resolvable POs in the file(s).' if phase == 'preview'
                       else 'No files cleanly parsed — no workbook written.')),
        }


def preview(channel_code: str, po_paths, warehouse: str | None = None) -> dict:
    """Module entry point — phase 1 (no writes). Views call this."""
    return MTProcessor(channel_code, po_paths, warehouse).preview()


def confirm(channel_code: str, po_paths, warehouse: str | None = None) -> dict:
    """Module entry point — phase 2 (write workbook + record to DB)."""
    return MTProcessor(channel_code, po_paths, warehouse).confirm()
