"""
online_b2b.services.gt_select_import
====================================

GT Select is an **Offline parent marketplace** whose Sales Orders are finalised
**directly in D365** — there's no engine/validation step. The operator simply
exports two files from D365 and we **load them into the DB** for the dashboard:

  * **Sales Orders** (headers)  → ``order_headers`` (segment='Offline',
    marketplace='GT Select', mode='IMPORT').
  * **Sales Lines** (lines)     → ``order_lines`` (status='OK'; pricing/validation
    columns stay empty — we're not re-pricing a D365-finalised order).

Header ↔ line join: the line's **``Document No.``** equals the header's **``No.``**.

**Dedup key = External Document No.** (NOT the SO No): D365 mints a fresh SO No on
every export, but the External Document No. is the stable order reference — so we
skip any header whose external doc was already imported (and its lines).

Web-owned; the engine is the frozen backup and is untouched.
"""

from __future__ import annotations

import datetime as _dt
import os

import pandas as pd

from .order_db import _conn

MARKETPLACE = 'GT Select'
SEGMENT = 'Offline'
# runs.mode / order_headers.mode are an engine-owned ENUM('AUTO','MANUAL') — we
# can't add 'IMPORT' without altering the frozen shared schema. So we use the
# valid 'MANUAL' and convey "this was a GT Select import" via the run ``source``
# ("GT Select import: …") + marketplace='GT Select'.
MODE = 'MANUAL'


def _clean(v) -> str:
    """Stringify an id without a trailing '.0' (pandas float coercion)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v).strip()
    head, dot, tail = s.partition('.')
    if dot and head.lstrip('-').isdigit() and tail.strip('0') == '':
        return head
    return s


def _num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        return float(str(v).replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def _date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == '':
        return None
    dt = pd.to_datetime(v, errors='coerce')
    return None if pd.isna(dt) else dt.date()


def _colmap(df):
    return {''.join(str(c).split()).lower(): c for c in df.columns}


def _find(cmap, *needles, exact_only=False):
    """Resolve a column by normalised name — exact match first, then substring
    (unless ``exact_only``, used for ambiguous keys like 'no.')."""
    for n in needles:
        if n in cmap:
            return cmap[n]
    if exact_only:
        return None
    for n in needles:
        for k, orig in cmap.items():
            if n in k:
                return orig
    return None


# ── Parse ────────────────────────────────────────────────────────────────

def parse_headers(path: str) -> dict:
    """D365 'Sales Orders' export → mapped header rows."""
    try:
        df = pd.read_excel(path, sheet_name=0)
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"Headers file unreadable: {e}"}
    c = _colmap(df)
    so = _find(c, 'no.', exact_only=True)
    if not so:
        return {'ok': False, 'error':
                "Couldn't find the 'No.' (Sales Order) column — is this the "
                "D365 'Sales Orders' export?"}
    ext = _find(c, 'externaldocumentno.', 'externaldocument')
    loc = _find(c, 'ship-toname', 'sell-tocustomername', 'customername')
    wh = _find(c, 'locationcode')
    pod = _find(c, 'documentdate')
    exd = _find(c, 'invoicetodate')
    qty = _find(c, 'totalquantity')
    val = _find(c, 'totalamountincl.gst', 'totalamountincl', 'amountincludingvat')
    sts = _find(c, 'status')
    rows = []
    for _, r in df.iterrows():
        son = _clean(r.get(so))
        if not son:
            continue
        rows.append({
            'so_no': son,
            'external_doc': _clean(r.get(ext)) if ext else '',
            'location': str(r.get(loc)).strip() if loc and pd.notna(r.get(loc)) else '',
            'warehouse': str(r.get(wh)).strip() if wh and pd.notna(r.get(wh)) else '',
            'po_date': _date(r.get(pod)) if pod else None,
            'exp_date': _date(r.get(exd)) if exd else None,
            'qty': int(_num(r.get(qty)) or 0) if qty else 0,
            'order_value': _num(r.get(val)) or 0.0 if val else 0.0,
            'status': str(r.get(sts)).strip() if sts and pd.notna(r.get(sts)) else '',
            'order_type': 'TO' if son.upper().startswith('TO') else 'SO',
        })
    return {'ok': True, 'rows': rows}


def parse_lines(path: str) -> dict:
    """D365 'Sales Lines' export → mapped line rows (keyed to a header by
    ``Document No.``)."""
    try:
        df = pd.read_excel(path, sheet_name=0)
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"Lines file unreadable: {e}"}
    c = _colmap(df)
    doc = _find(c, 'documentno.', 'documentno')
    item = _find(c, 'no.', exact_only=True)
    if not doc or not item:
        return {'ok': False, 'error':
                "Couldn't find 'Document No.' / 'No.' columns — is this the "
                "D365 'Sales Lines' export?"}
    gtin = _find(c, 'gtin', 'ean', 'barcode')
    desc = _find(c, 'description')
    qty = _find(c, 'quantity')
    amt = _find(c, 'lineamountexcl.vat', 'lineamountexcl', 'lineamount')
    loc = _find(c, 'locationcode')
    typ = _find(c, 'type')
    rows = []
    for _, r in df.iterrows():
        son = _clean(r.get(doc))
        ino = _clean(r.get(item))
        if not son or not ino:
            continue
        # skip non-item lines (comments / charges) when a Type column exists
        if typ and pd.notna(r.get(typ)) and str(r.get(typ)).strip().lower() not in ('item', ''):
            continue
        q = int(_num(r.get(qty)) or 0) if qty else 0
        line_amt = _num(r.get(amt)) if amt else None
        unit = round(line_amt / q, 2) if (line_amt is not None and q) else None
        rows.append({
            'so_no': son,
            'item_no': ino,
            'ean': _clean(r.get(gtin)) if gtin else '',
            'description': str(r.get(desc)).strip() if desc and pd.notna(r.get(desc)) else '',
            'qty': q,
            'unit_price': unit,
            'location': str(r.get(loc)).strip() if loc and pd.notna(r.get(loc)) else '',
        })
    return {'ok': True, 'rows': rows}


# ── Dedup (on External Document No.) ──────────────────────────────────────

def existing_external_docs(docs: list[str]) -> set:
    """Subset of ``docs`` already imported for GT Select (the stable order key)."""
    docs = [d for d in docs if d]
    if not docs:
        return set()
    with _conn() as (cur, d):
        ph = d['ph']
        marks = ','.join([ph] * len(docs))
        cur.execute(
            f"SELECT external_doc FROM order_headers WHERE marketplace={ph} "
            f"AND external_doc IN ({marks})", (MARKETPLACE, *docs))
        return {r[0] for r in cur.fetchall() if r[0]}


# ── Preview + import ──────────────────────────────────────────────────────

def preview(headers_path: str, lines_path: str) -> dict:
    """Parse both files, join, dedup on external_doc, build the review payload.
    No DB write."""
    ph = parse_headers(headers_path)
    if not ph['ok']:
        return ph
    pl = parse_lines(lines_path)
    if not pl['ok']:
        return pl
    headers, lines = ph['rows'], pl['rows']

    # lines grouped by their SO (Document No.)
    by_so: dict = {}
    for ln in lines:
        by_so.setdefault(ln['so_no'], []).append(ln)

    warnings = []
    blank_ext = [h['so_no'] for h in headers if not h['external_doc']]
    if blank_ext:
        warnings.append(
            f"{len(blank_ext)} order(s) have a blank External Document No. — "
            f"they can't be deduped and will import every time "
            f"(e.g. {', '.join(blank_ext[:6])}).")

    seen = existing_external_docs([h['external_doc'] for h in headers])
    for h in headers:
        # New unless its external doc was already imported. A blank external doc
        # can't be deduped → treated as new (and warned above).
        h['is_new'] = not (h['external_doc'] and h['external_doc'] in seen)
        h['line_count'] = len(by_so.get(h['so_no'], []))
        if h['line_count'] == 0:
            warnings.append(f"Order {h['so_no']} has no matching lines in the "
                            f"Sales Lines file.")

    # lines whose SO isn't in the headers file
    hdr_sos = {h['so_no'] for h in headers}
    orphan = sorted({ln['so_no'] for ln in lines if ln['so_no'] not in hdr_sos})
    if orphan:
        warnings.append(f"{len(orphan)} line-SO(s) have no matching header — "
                        f"those lines are skipped (e.g. {', '.join(orphan[:6])}).")

    new_headers = [h for h in headers if h['is_new']]
    new_lines = sum(len(by_so.get(h['so_no'], [])) for h in new_headers)
    return {
        'ok': True, 'headers': headers, 'by_so': by_so,
        'sample_lines': lines[:12],
        'summary': {
            'total': len(headers), 'new': len(new_headers),
            'dup': len(headers) - len(new_headers),
            'lines': len(lines), 'new_lines': new_lines,
            'qty': sum(h['qty'] for h in new_headers),
            'value': sum(h['order_value'] for h in new_headers),
        },
        'warnings': warnings,
    }


def do_import(headers_path: str, lines_path: str) -> dict:
    """Insert the NEW GT Select headers + their lines under one IMPORT run.
    Returns {ok, run_id, imported, skipped, lines, error}."""
    pv = preview(headers_path, lines_path)
    if not pv['ok']:
        return pv
    new_headers = [h for h in pv['headers'] if h['is_new']]
    if not new_headers:
        return {'ok': True, 'run_id': None, 'imported': 0,
                'skipped': len(pv['headers']), 'lines': 0}

    by_so = pv['by_so']
    run_ts = _dt.datetime.now()
    src = f"GT Select import: {os.path.basename(headers_path)}"
    total_qty = sum(h['qty'] for h in new_headers)
    total_val = sum(h['order_value'] for h in new_headers)
    total_lines = sum(len(by_so.get(h['so_no'], [])) for h in new_headers)

    try:
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(
                f"INSERT INTO runs (run_ts, mode, source, marketplaces, "
                f"total_pos, total_items, total_qty, total_value, "
                f"consolidated_path, tracker_path) VALUES "
                f"({ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph})",
                (run_ts, MODE, src, 1, len(new_headers), total_lines,
                 total_qty, total_val, '', ''))
            run_id = cur.lastrowid

            hdr_cols = ['run_id', 'run_ts', 'mode', 'segment', 'marketplace',
                        'marketplace_label', 'po', 'location', 'warehouse',
                        'po_date', 'exp_date', 'order_type', 'items', 'qty',
                        'order_value', 'output_file', 'external_doc']
            hph = ', '.join([ph] * len(hdr_cols))
            hpayload = [(
                run_id, run_ts, MODE, SEGMENT, MARKETPLACE, MARKETPLACE,
                h['so_no'], h['location'], h['warehouse'], h['po_date'],
                h['exp_date'], h['order_type'], len(by_so.get(h['so_no'], [])),
                h['qty'], h['order_value'], src, h['external_doc'],
            ) for h in new_headers]
            cur.executemany(
                f"INSERT INTO order_headers ({', '.join(hdr_cols)}) "
                f"VALUES ({hph})", hpayload)

            # GT Select is D365-finalised → facts only, NO validation row
            # (the view reads such lines as status='OK' via COALESCE).
            ln_cols = ['run_id', 'run_ts', 'marketplace', 'po', 'location',
                       'item_no', 'ean', 'description', 'qty', 'order_type',
                       'unit_price', 'output_file']
            lph = ', '.join([ph] * len(ln_cols))
            lpayload = []
            for h in new_headers:
                for ln in by_so.get(h['so_no'], []):
                    lpayload.append((
                        run_id, run_ts, MARKETPLACE, ln['so_no'],
                        ln['location'] or h['location'], ln['item_no'],
                        ln['ean'], ln['description'], ln['qty'],
                        h['order_type'], ln['unit_price'], src,
                    ))
            if lpayload:
                cur.executemany(
                    f"INSERT INTO order_lines ({', '.join(ln_cols)}) "
                    f"VALUES ({lph})", lpayload)
            cur.connection.commit()
        return {'ok': True, 'run_id': run_id, 'imported': len(new_headers),
                'skipped': len(pv['headers']) - len(new_headers),
                'lines': len(lpayload)}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"{type(e).__name__}: {e}"}
