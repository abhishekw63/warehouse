"""
online_b2b.services.gt_select_import
====================================

**D365 catch-all import** — the reliable way to capture orders that are
finalised/auto-punched **directly in D365** and therefore never pass through the
web app's own upload flows. GT Select is the canonical case (staged in D365, DMS
auto-punches the SO), so the web tracker was blind to it and "orders received
today" was understated. This importer closes that gap.

The operator exports two files from D365 and uploads both:

  * **Sales Orders** (headers) — every order across every channel.
  * **Sales Lines**  (lines)   — every line, joined to a header by
    ``Document No.`` == the header's ``No.``.

What we do with them
--------------------
1. **Filter to real orders** — D365 leaves empty SO shells (a ``No.`` + zero
   totals, no customer, no lines). Those are dropped.
2. **Classify each order by ``Gen. Bus. Posting Group``** → its true segment +
   marketplace (``OFF -GT-SELECT`` → Offline/GT Select, ``OFF-GT MASS`` → GT
   Mass, ``OFF-MT SELECT`` → MT, ``ON-B2B`` → Online B2B). Never blanket-labelled.
3. **Dedup against ALL recorded orders** (``order_headers`` by ``po`` OR
   ``external_doc``, any marketplace) — so orders already captured by their own
   native flow are skipped, and only the genuinely NEW ones are imported.
4. **Import the new ones** under their correct segment/marketplace with every DB
   field we can populate + all their lines (**testers included** — they carry
   real qty/value that the normal flow drops, another source of the undercount).

Incremental: re-uploading the full dump each day only ever adds the new orders.
Web-owned; the desktop engine is the frozen backup and is untouched.
"""

from __future__ import annotations

import datetime as _dt
import os

import pandas as pd

from .order_db import _conn, _conn_tx

SEGMENT_DEFAULT = 'Offline'
MARKETPLACE_DEFAULT = 'D365 Import'      # unknown posting group → visible, not silent
MODE = 'MANUAL'                          # engine ENUM has no 'IMPORT'; run.source says it

# ── Gen. Bus. Posting Group → (segment, marketplace) ────────────────────────
# Normalised key = spaces stripped, upper-cased, '_'→'-' (handles the export's
# quirky 'OFF -GT-SELECT'/'OFF-GT MASS' spacing).
_PG_MAP = {
    'OFF-GTMASS':    ('Offline',   'GT Mass'),
    'OFF-GT-MASS':   ('Offline',   'GT Mass'),
    'OFF-GTSELECT':  ('Offline',   'GT Select'),
    'OFF-GT-SELECT': ('Offline',   'GT Select'),
    'OFF-MTSELECT':  ('Offline',   'MT'),
    'OFF-MT-SELECT': ('Offline',   'MT'),
    'ON-B2B':        ('OnlineB2B', 'Online B2B'),
    'ONB2B':         ('OnlineB2B', 'Online B2B'),
}


def _norm_pg(v) -> str:
    return ''.join(str(v or '').split()).upper().replace('_', '-')


def classify(posting_group, learned=None):
    """(segment, marketplace, marketplace_label) for a D365 posting group. Checks
    the operator-TAUGHT map first (durable, learned from earlier classifications),
    then the built-in map, then a visible default — never a silent bucket."""
    key = _norm_pg(posting_group)
    if learned and key in learned:
        m = learned[key]
        return m['segment'], m['marketplace'], m.get('marketplace_label') or m['marketplace']
    seg, mp = _PG_MAP.get(key, (SEGMENT_DEFAULT, MARKETPLACE_DEFAULT))
    return seg, mp, mp


# ── Learned posting-group → channel map (durable; taught in the UI) ──────────
_PG_TABLE = 'd365_posting_group_map'


def _ensure_pg_table(cur):
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS {_PG_TABLE} ("
        "pg_key VARCHAR(120) PRIMARY KEY, posting_group VARCHAR(120), "
        "segment VARCHAR(20), marketplace VARCHAR(50), "
        "marketplace_label VARCHAR(60), created_at DATETIME, created_by VARCHAR(150))")


def load_pg_map() -> dict:
    """``{norm_pg: {segment, marketplace, marketplace_label}}`` taught by operators
    — merged ahead of the built-in map so a once-classified group is never
    'unmapped' again. ``{}`` on any error (degrades to the built-in map)."""
    out: dict = {}
    try:
        with _conn() as (cur, d):
            _ensure_pg_table(cur)
            cur.execute(f"SELECT pg_key, segment, marketplace, marketplace_label FROM {_PG_TABLE}")
            for k, seg, mp, lbl in cur.fetchall():
                out[k] = {'segment': seg, 'marketplace': mp, 'marketplace_label': lbl or mp}
    except Exception:  # noqa: BLE001
        pass
    return out


def save_pg_map(overrides: dict, user: str = '') -> int:
    """Persist operator classifications so THIS batch + every future one auto-map.
    ``overrides = {norm_pg: {posting_group, segment, marketplace, marketplace_label}}``."""
    rows = {k: m for k, m in (overrides or {}).items() if m.get('marketplace')}
    if not rows:
        return 0
    with _conn() as (cur, d):
        ph = d['ph']
        _ensure_pg_table(cur)
        now = _dt.datetime.now()
        for key, m in rows.items():
            cur.execute(
                f"INSERT INTO {_PG_TABLE} (pg_key, posting_group, segment, marketplace, "
                f"marketplace_label, created_at, created_by) VALUES "
                f"({ph},{ph},{ph},{ph},{ph},{ph},{ph}) "
                f"ON DUPLICATE KEY UPDATE segment=VALUES(segment), "
                f"marketplace=VALUES(marketplace), marketplace_label=VALUES(marketplace_label)",
                (key, m.get('posting_group', ''), m.get('segment'), m['marketplace'],
                 m.get('marketplace_label') or m['marketplace'], now, user))
        cur.connection.commit()
    return len(rows)


# ── Cell helpers ────────────────────────────────────────────────────────────

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
    """Coerce any D365 date cell (datetime object, NaT, or messy string) to a
    ``date`` — ``None`` when blank/NaT/unparseable, so a nullable DATE column
    gets a real NULL, never the string 'NaT' (which the DB rejects)."""
    if v is None:
        return None
    try:
        if pd.isna(v):            # NaN / NaT — must be checked BEFORE isinstance
            return None           # (pandas NaT can pass isinstance(datetime)).
    except (TypeError, ValueError):
        pass
    if isinstance(v, _dt.datetime):
        return v.date()
    if isinstance(v, _dt.date):
        return v
    s = str(v).strip()
    if not s or s.lower() in ('nat', 'nan', 'none'):
        return None
    dt = pd.to_datetime(v, errors='coerce', dayfirst=False)
    if not pd.isna(dt):
        return dt.date()
    # last resort: pull a leading calendar date out of a messy timestamp string
    import re as _re
    m = _re.match(r'\s*(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        y, mo, d = m.groups()
        try:
            return _dt.date(int(y), int(mo), int(d))
        except ValueError:
            return None
    return None


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

def parse_headers(path: str, learned: dict | None = None) -> dict:
    """D365 'Sales Orders' export → mapped, classified header rows. ``learned`` is
    the operator-taught posting-group map (see :func:`load_pg_map`)."""
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
    pg = _find(c, 'gen.bus.postinggroup', 'businesspostinggroup', 'postinggroup')
    shipname = _find(c, 'ship-toname')
    custname = _find(c, 'sell-tocustomername', 'customername', 'bill-toname')
    custno = _find(c, 'sell-tocustomerno.', 'sell-tocustomerno', 'bill-tocustomerno.')
    shipcode = _find(c, 'ship-tocode')
    postcode = _find(c, 'ship-topostcode', 'sell-topostcode', 'postcode')
    wh = _find(c, 'locationcode')
    pod = _find(c, 'documentdate')
    exd = _find(c, 'invoicetodate')
    qty = _find(c, 'totalquantity')
    val = _find(c, 'totalamountincl.gst', 'totalamountincl', 'amountincludingvat')
    sts = _find(c, 'status')

    def _s(r, col):
        return str(r.get(col)).strip() if col and pd.notna(r.get(col)) else ''

    rows = []
    for _, r in df.iterrows():
        son = _clean(r.get(so))
        if not son:
            continue
        seg, mp, lbl = classify(r.get(pg), learned) if pg \
            else (SEGMENT_DEFAULT, MARKETPLACE_DEFAULT, MARKETPLACE_DEFAULT)
        rows.append({
            'so_no': son,
            'external_doc': _clean(r.get(ext)) if ext else '',
            'posting_group': _s(r, pg),
            'segment': seg, 'marketplace': mp, 'marketplace_label': lbl,
            'ship_name': _s(r, shipname) or _s(r, custname),
            'customer_name': _s(r, custname),
            'ship_to_name': _s(r, shipname),
            'cust_no': _clean(r.get(custno)) if custno else '',
            'ship_code': _s(r, shipcode),
            'postcode': _clean(r.get(postcode)) if postcode else '',
            'warehouse': _s(r, wh),
            'po_date': _date(r.get(pod)) if pod else None,
            'exp_date': _date(r.get(exd)) if exd else None,
            'qty': int(_num(r.get(qty)) or 0) if qty else 0,
            'order_value': _num(r.get(val)) or 0.0 if val else 0.0,
            'status': _s(r, sts),
            'order_type': 'TO' if son.upper().startswith('TO') else 'SO',
        })
    return {'ok': True, 'rows': rows}


def parse_lines(path: str) -> dict:
    """D365 'Sales Lines' export → mapped line rows (keyed by ``Document No.``).
    Keeps EVERY item line — testers included."""
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
    pin = _find(c, 'pincode', 'pin')
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
            'pincode': _clean(r.get(pin)) if pin else '',
        })
    return {'ok': True, 'rows': rows}


# ── Dedup — against ALL recorded orders (any marketplace) ────────────────────

def existing_recorded(keys: list[str]) -> set:
    """Subset of ``keys`` already present in ``order_headers`` as a ``po`` OR an
    ``external_doc`` (across every marketplace) — an order captured by ANY flow
    counts as already-recorded, so we never double-count."""
    keys = list({k for k in keys if k})
    if not keys:
        return set()
    found: set = set()
    with _conn() as (cur, d):
        ph = d['ph']
        for i in range(0, len(keys), 500):
            chunk = keys[i:i + 500]
            marks = ','.join([ph] * len(chunk))
            cur.execute(f"SELECT po FROM order_headers WHERE po IN ({marks})", tuple(chunk))
            found |= {str(r[0]) for r in cur.fetchall() if r[0] is not None}
            cur.execute(f"SELECT external_doc FROM order_headers WHERE external_doc IN ({marks})", tuple(chunk))
            found |= {str(r[0]) for r in cur.fetchall() if r[0]}
    return found


def _order_po(h) -> str:
    """The order reference we store as ``po`` — the External Document No. (the
    stable channel PO), falling back to the D365 ``No.`` when it's blank."""
    return h['external_doc'] or h['so_no']


# ── Preview + import ──────────────────────────────────────────────────────

def preview(headers_path: str, lines_path: str) -> dict:
    """Parse both files, classify, join, dedup vs ALL recorded. No DB write.
    Returns per-channel New/Already breakdown for the review page."""
    ph = parse_headers(headers_path, load_pg_map())
    if not ph['ok']:
        return ph
    pl = parse_lines(lines_path)
    if not pl['ok']:
        return pl
    headers, lines = ph['rows'], pl['rows']

    by_so: dict = {}
    for ln in lines:
        by_so.setdefault(ln['so_no'], []).append(ln)

    # Real orders only — an order with lines. D365's empty SO shells (no lines,
    # zero totals) are dropped (reported, not silent).
    real = [h for h in headers if by_so.get(h['so_no'])]
    empty = len(headers) - len(real)

    warnings = []
    if empty:
        warnings.append(f"{empty} empty D365 SO shell(s) with no lines — skipped "
                        f"(staging placeholders).")
    blank_key = [h['so_no'] for h in real if not _order_po(h)]
    if blank_key:
        warnings.append(f"{len(blank_key)} order(s) have no External Doc / No. — "
                        f"can't be deduped (e.g. {', '.join(blank_key[:6])}).")
    orphan = sorted({ln['so_no'] for ln in lines} - {h['so_no'] for h in headers})
    if orphan:
        warnings.append(f"{len(orphan)} line-SO(s) have no matching header — those "
                        f"lines are skipped (e.g. {', '.join(orphan[:6])}).")

    seen = existing_recorded([k for h in real for k in (h['external_doc'], h['so_no'])])
    for h in real:
        h['line_count'] = len(by_so.get(h['so_no'], []))
        h['is_new'] = not ((h['external_doc'] and h['external_doc'] in seen)
                           or (h['so_no'] in seen))

    # per-channel breakdown (segment · marketplace)
    chan: dict = {}
    for h in real:
        k = (h['segment'], h['marketplace'])
        c = chan.setdefault(k, {'segment': h['segment'], 'marketplace': h['marketplace'],
                                'new': 0, 'dup': 0, 'new_qty': 0, 'new_value': 0.0})
        if h['is_new']:
            c['new'] += 1
            c['new_qty'] += h['qty']
            c['new_value'] += h['order_value']
        else:
            c['dup'] += 1
    channels = sorted(chan.values(), key=lambda x: (-x['new'], x['marketplace']))

    # Unknown posting groups among the NEW orders → the operator must place each
    # (segment → marketplace → MT child) before importing; never silently bucketed.
    unmapped: dict = {}
    for h in real:
        if h['is_new'] and h['marketplace'] == MARKETPLACE_DEFAULT:
            key = _norm_pg(h['posting_group'])
            u = unmapped.setdefault(key, {
                'key': key, 'posting_group': h['posting_group'] or '(blank)',
                'count': 0, 'qty': 0, 'value': 0.0})
            u['count'] += 1
            u['qty'] += h['qty']
            u['value'] += h['order_value']

    new_headers = [h for h in real if h['is_new']]
    new_lines = sum(len(by_so.get(h['so_no'], [])) for h in new_headers)
    return {
        'ok': True, 'headers': real, 'by_so': by_so, 'channels': channels,
        'needs_class': sorted(unmapped.values(), key=lambda x: -x['count']),
        'sample_new': new_headers[:40], 'sample_lines': lines[:12],
        'summary': {
            'total': len(real), 'empty': empty,
            'new': len(new_headers), 'dup': len(real) - len(new_headers),
            'lines': len(lines), 'new_lines': new_lines,
            'qty': sum(h['qty'] for h in new_headers),
            'value': sum(h['order_value'] for h in new_headers),
        },
        'warnings': warnings,
    }


def do_import(headers_path: str, lines_path: str, overrides: dict | None = None,
              user: str = '', only_pos=None) -> dict:
    """Insert the NEW orders + their lines under one IMPORT run, each under its
    OWN segment/marketplace. ``overrides`` places unknown posting groups the
    operator classified in the UI: ``{norm_pg: {'posting_group','segment',
    'marketplace','marketplace_label'}}`` — PERSISTED so this batch and every
    future one auto-map. ``only_pos`` (optional) = the External-Doc / SO-No values
    the operator TICKED to push; None = push every new order. Returns {ok, run_id,
    imported, skipped, lines}."""
    overrides = overrides or {}
    if overrides:
        try:
            save_pg_map(overrides, user)      # teach it — now + next batch
        except Exception:  # noqa: BLE001
            pass
    pv = preview(headers_path, lines_path)    # re-classifies via the just-saved map
    if not pv['ok']:
        return pv
    new_headers = [h for h in pv['headers'] if h['is_new']]
    if only_pos is not None:                  # push ONLY the ticked orders
        want = {str(p).strip() for p in only_pos if str(p).strip()}
        new_headers = [h for h in new_headers
                       if (h['external_doc'] and h['external_doc'] in want) or h['so_no'] in want]
    if not new_headers:
        return {'ok': True, 'run_id': None, 'imported': 0,
                'skipped': len(pv['headers']), 'lines': 0}

    # Apply the operator's classification to any unknown posting group.
    for h in new_headers:
        if h['marketplace'] == MARKETPLACE_DEFAULT:
            ov = overrides.get(_norm_pg(h['posting_group']))
            if ov and ov.get('marketplace'):
                h['segment'] = ov.get('segment') or h['segment']
                h['marketplace'] = ov['marketplace']
                h['marketplace_label'] = ov.get('marketplace_label') or ov['marketplace']

    by_so = pv['by_so']
    run_ts = _dt.datetime.now()
    src = f"D365 import: {os.path.basename(headers_path)}"
    total_qty = sum(h['qty'] for h in new_headers)
    total_val = sum(h['order_value'] for h in new_headers)
    total_lines = sum(len(by_so.get(h['so_no'], [])) for h in new_headers)
    n_channels = len({h['marketplace'] for h in new_headers})

    def _location(h) -> str:
        """Ship-to name with the postcode appended so the tracker derives
        State/Zone from it (same shape as the address-keyed channels)."""
        name = h['ship_name'] or h['cust_no'] or ''
        pc = h['postcode']
        return f"{name},{pc}" if (name and pc) else (name or pc)

    try:
        # ATOMIC: run row + headers + lines commit together (no orphan header
        # without its lines — else dedup would skip it forever on retry).
        with _conn_tx() as (cur, d):
            ph = d['ph']
            cur.execute(
                f"INSERT INTO runs (run_ts, mode, source, marketplaces, "
                f"total_pos, total_items, total_qty, total_value, "
                f"consolidated_path, tracker_path) VALUES "
                f"({ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph})",
                (run_ts, MODE, src, n_channels, len(new_headers), total_lines,
                 total_qty, total_val, '', ''))
            run_id = cur.lastrowid

            hdr_cols = ['run_id', 'run_ts', 'mode', 'segment', 'marketplace',
                        'marketplace_label', 'po', 'location', 'warehouse',
                        'po_date', 'exp_date', 'order_type', 'items', 'qty',
                        'order_value', 'output_file', 'external_doc',
                        'so_no', 'customer_no', 'customer_name',
                        'ship_to_code', 'ship_to_name']
            hph = ', '.join([ph] * len(hdr_cols))
            # po = External Doc No (the distributor/marketplace PO); external_doc keeps
            # the same ref; so_no captures the D365 SO Number separately — both kept.
            hpayload = [(
                run_id, run_ts, MODE, h['segment'], h['marketplace'],
                h.get('marketplace_label') or h['marketplace'],
                _order_po(h), _location(h), h['warehouse'], h['po_date'],
                h['exp_date'], h['order_type'], len(by_so.get(h['so_no'], [])),
                h['qty'], h['order_value'], src, h['external_doc'],
                h['so_no'], h['cust_no'], h['customer_name'],
                h['ship_code'], h['ship_to_name'],
            ) for h in new_headers]
            cur.executemany(
                f"INSERT INTO order_headers ({', '.join(hdr_cols)}) "
                f"VALUES ({hph})", hpayload)

            # D365-finalised → facts only, NO validation row (the view reads such
            # lines as status='OK' via COALESCE).
            ln_cols = ['run_id', 'run_ts', 'marketplace', 'po', 'location',
                       'item_no', 'ean', 'description', 'qty', 'order_type',
                       'unit_price', 'output_file']
            lph = ', '.join([ph] * len(ln_cols))
            lpayload = []
            for h in new_headers:
                po = _order_po(h)
                loc = _location(h)
                for ln in by_so.get(h['so_no'], []):
                    lpayload.append((
                        run_id, run_ts, h['marketplace'], po,
                        ln['location'] or loc, ln['item_no'], ln['ean'],
                        ln['description'], ln['qty'], h['order_type'],
                        ln['unit_price'], src,
                    ))
            if lpayload:
                cur.executemany(
                    f"INSERT INTO order_lines ({', '.join(ln_cols)}) "
                    f"VALUES ({lph})", lpayload)
            # _conn_tx commits on clean exit / rolls back on any exception above.
        return {'ok': True, 'run_id': run_id, 'imported': len(new_headers),
                'skipped': len(pv['headers']) - len(new_headers),
                'lines': len(lpayload)}
    except Exception as e:  # noqa: BLE001
        return {'ok': False, 'error': f"{type(e).__name__}: {e}"}
