"""Headless EKA bridge — drive the shared EKA engine from the web app.

Swaps only two inputs for the web (no external-file dependencies):
  1. **Item Master** ← DB ``item_master`` (current MRP — fixes stale-dump pricing)
  2. **Store registry** ← DB ``eka_data`` (Active rows)

Mode = "Standalone PO files": each uploaded file is one store; the engine splits
into a regular doc (finished-goods PO at calculated cost) and a tester doc
(``/TT/`` — testers + PWP + GWP + non-stock at ₹0.54) and assigns SO/TO numbers.
EKA has **no CP check**. Records to the shared ``order_headers`` (single source of
truth) via the same history_db recorder every other channel uses.
"""
from __future__ import annotations

import os
import re
import tempfile
from typing import Dict, List, Optional, Tuple

# Tkinter-free shared engine — the web has no dependency on the desktop file.
from offline.services import eka_engine


# ── DB-sourced inputs ─────────────────────────────────────────────────────────
def _db_master() -> Dict:
    """POEngine.master from DB ``item_master`` — keyed by EAN and item_no."""
    from online_b2b.services.order_db import _conn
    master: Dict = {}
    with _conn() as (cur, _d):
        cur.execute("SELECT item_no, ean, description, gst_code, mrp "
                    "FROM item_master")
        for item_no, ean, desc, gst, mrp in cur.fetchall():
            rec = {'item_no': item_no, 'mrp': mrp,
                   'gst_code': str(gst or ''), 'description': str(desc or '')}
            e = str(ean or '').strip()
            if e and e.lower() != 'none':
                master[e] = rec
            ic = str(item_no).strip()
            if ic and ic not in master:
                master[ic] = rec
    return master


def _db_locations() -> List[Dict]:
    """eka_locations from DB ``eka_data`` (Active) — same shape as load_eka_data."""
    from online_b2b.services.order_db import _conn
    locs: List[Dict] = []
    with _conn() as (cur, _d):
        cur.execute(
            "SELECT short_name, prefix, short_code, transfer_code, "
            "location_code, kind, posting_group, bill_to, ship_to, status "
            "FROM eka_data")
        for sn, pf, sc, tc, loc, kind, pg, bt, st, status in cur.fetchall():
            if str(status or 'Active').strip().lower() == 'inactive':
                continue
            if not str(sc or '').strip():
                continue
            clean = lambda v: '' if (v is None or str(v).strip() == '-') else str(v).strip()
            locs.append({
                'short_name': str(sn or '').strip(),
                'prefix': str(pf or 'TO').strip(),
                'short_code': str(sc or '').strip(),
                'transfer_code': clean(tc),
                'location': str(loc or '').strip(),
                'type': str(kind or '').strip(),
                'posting_group': clean(pg),
                'bill_to': clean(bt),
                'ship_to': clean(st),
            })
    return locs


def _db_type_map() -> Dict[str, str]:
    """{store code (upper) → segment label (Airport/EBO/Kiosk)} from DB eka_data
    ``kind`` — keyed by location_code, transfer_code AND short_code so an order
    resolves however it's matched. The web equivalent of _load_eka_type_map()."""
    from online_b2b.services.order_db import _conn
    out: Dict[str, str] = {}
    with _conn() as (cur, _d):
        cur.execute("SELECT location_code, transfer_code, short_code, kind "
                    "FROM eka_data")
        for loc, tc, sc, kind in cur.fetchall():
            k = str(kind or '').strip()
            if not k:
                continue
            for v in (loc, tc, sc):
                v = str(v or '').strip()
                if v and v != '-':
                    out[v.upper()] = k
    return out


def _lookup_location(fname: str, locations: List[Dict]) -> Tuple[Optional[Dict], int]:
    """Mirror the desktop ``_lookup_location_from_filename`` (exact→suffix→prefix)."""
    loc_code = re.sub(r'\.(xlsx|xlsm|xls)$', '', os.path.basename(fname), flags=re.I)
    for loc in locations:
        if loc.get('location', '') == loc_code:
            return (loc, 0)
    m = re.match(r'^(.+)_(\d+)$', loc_code)
    if m:
        base, idx = m.group(1), int(m.group(2))
        for loc in locations:
            if loc.get('location', '') == base:
                return (loc, idx)
    for loc in locations:
        lv = loc.get('location', '')
        if lv and loc_code.startswith(lv):
            return (loc, 0)
    return (None, 0)


# ── core processing (shared by preview / confirm / workbook) ──────────────────
def process(paths: List[str], warehouse: str = 'AHD') -> Dict:
    """Run the engine over the uploaded store files (DB master + DB registry).
    Returns processed ``results`` (LocationResults, SO/TO numbers assigned) plus a
    processing log, per-file issues and warnings. NO DB write, NO workbook write."""
    POEngine = eka_engine.POEngine
    SpecialOrderEngine = eka_engine.SpecialOrderEngine

    engine = POEngine()
    engine.master = _db_master()
    locations = _db_locations()

    results = []
    processing_log: List[Dict] = []
    file_issues: List[Dict] = []
    warnings: List[str] = []
    counter = SpecialOrderEngine.get_today_date_code()

    for path in paths:
        fname = os.path.basename(path)
        loc = re.sub(r'\.(xlsx|xlsm|xls)$', '', fname, flags=re.I)
        log_entry = {'filename': fname, 'location': loc, 'status': 'OK',
                     'issues': [], 'actions': [], 'to_number': '', 'tt_number': ''}
        processing_log.append(log_entry)
        try:
            vlogs = engine.validate_file(path)
            for level, msg in vlogs:
                if level == 'alert':
                    log_entry['actions'].append(msg)
                    if log_entry['status'] == 'OK':
                        log_entry['status'] = 'AUTO_FIXED'
                elif level == 'warn':
                    log_entry['issues'].append(msg)
                    if log_entry['status'] == 'OK':
                        log_entry['status'] = 'WARNING'
                elif level == 'error':
                    log_entry['issues'].append(msg)
                    log_entry['status'] = 'FAILED'
            if any(lv == 'error' for lv, _ in vlogs):
                file_issues.append({'file': fname, 'problem': 'Validation failed',
                                    'detail': '; '.join(log_entry['issues']),
                                    'kind': 'error'})
                continue

            res = engine.process_file(path)
            eka_loc, suffix_idx = _lookup_location(path, locations)
            if eka_loc:
                has_regular = bool(res.regular_orders)
                has_tester = bool(res.tester_orders or res.pwp_orders
                                  or res.gwp_orders or res.nonstock_orders)
                short_code = eka_loc['short_code']
                if suffix_idx > 0:
                    short_code = f"{short_code}_{suffix_idx + 1}"
                to_regular = to_tester = ''
                if has_regular:
                    to_regular = SpecialOrderEngine.generate_to_number(
                        eka_loc['prefix'], short_code, is_tester=False,
                        date_code=counter)
                    counter += 1
                if has_tester:
                    to_tester = SpecialOrderEngine.generate_to_number(
                        eka_loc['prefix'], short_code, is_tester=True,
                        date_code=counter)
                    counter += 1
                tc, pg = eka_loc['transfer_code'], eka_loc['posting_group']
                res._so_bill_to = eka_loc.get('bill_to', '')
                res._so_ship_to = eka_loc.get('ship_to', '')
                for item in res.regular_orders:
                    item.to, item.transfer_to, item.posting_group = to_regular, tc, pg
                for item in (res.pwp_orders + res.tester_orders
                             + res.gwp_orders + res.nonstock_orders):
                    item.to, item.transfer_to, item.posting_group = to_tester, tc, pg
                log_entry['to_number'] = to_regular
                log_entry['tt_number'] = to_tester
            else:
                log_entry['issues'].append(
                    f"Location '{loc}' not in eka_data — TO/Transfer/Posting empty")
                if log_entry['status'] == 'OK':
                    log_entry['status'] = 'WARNING'
                file_issues.append({'file': fname, 'problem': 'Store not mapped',
                                    'detail': f"'{loc}' not in eka_data",
                                    'kind': 'warn'})

            results.append(res)
            if res.unmatched:
                warnings.append(f"{loc}: {len(res.unmatched)} EAN(s) not in master")
                file_issues.append({'file': fname, 'problem': 'Unmatched EANs',
                                    'detail': f"{len(res.unmatched)} EAN(s) not in "
                                              f"item master", 'kind': 'warn'})
        except Exception as e:  # noqa: BLE001
            log_entry['status'] = 'FAILED'
            log_entry['issues'].append(f"Processing error: {e}")
            file_issues.append({'file': fname, 'problem': 'Processing error',
                                'detail': str(e), 'kind': 'error'})

    return {'results': results, 'processing_log': processing_log,
            'file_issues': file_issues, 'warnings': warnings}


def write_review(results, processing_log, output_path: Optional[str] = None) -> str:
    """Write the EKA review workbook (the desktop 9-sheet PO_Output)."""
    if output_path is None:
        fd, output_path = tempfile.mkstemp(suffix='_EKA_review.xlsx')
        os.close(fd)
    eka_engine.ExcelWriter.write(results, output_path,
                                 processing_log=processing_log)
    return output_path


# ── DB recording (shared order_headers, single source of truth) ───────────────
def _history_db():
    """Import the shared history_db (backs order_headers). None if unavailable."""
    try:
        import online_po_processor.auto.history_db as H
        return H
    except Exception:  # noqa: BLE001 — walk up to online_po_management
        import sys
        from pathlib import Path
        here = Path(__file__).resolve()
        for base in [here, *here.parents]:
            cand = base / 'online_po_management'
            if (cand / 'online_po_processor' / 'auto' / 'history_db.py').exists():
                if str(cand) not in sys.path:
                    sys.path.insert(0, str(cand))
                try:
                    import online_po_processor.auto.history_db as H
                    return H
                except Exception:  # noqa: BLE001
                    return None
    return None


def record(results, output_file: str = '', warehouse: str = 'AHD') -> Dict:
    """Record the processed EKA batch into the shared ``order_headers`` (segment
    'Offline' / marketplace 'EKA'). One row per SO/TO number; new only (dedup by
    marketplace+po). Uses the DB segment labels. Soft-fails, never raises."""
    H = _history_db()
    if H is None:
        return {'recorded': False, 'reason': 'history_db module not found'}
    from datetime import date, datetime
    rows = eka_engine.build_eka_order_rows(
        results, output_file, type_map=_db_type_map(),
        po_date=date.today().isoformat(), warehouse=warehouse)
    if not rows:
        return {'recorded': False, 'reason': 'no orders to record'}
    try:
        db_path = H.default_history_db_path()
        store = H.get_history_store(db_path)
        try:
            existing = store.existing_pos()
        finally:
            store.close()
        new_rows = [r for r in rows
                    if (r['marketplace'], r['po']) not in existing]
        skipped = len(rows) - len(new_rows)
        if not new_rows:
            return {'recorded': False, 'reason': 'all orders already recorded',
                    'skipped': skipped}
        run_meta = {
            'run_ts': datetime.now().isoformat(timespec='seconds'),
            'mode': 'WEB',
            'online_root': (f'OFFLINE EKA (web): {output_file}'
                            if output_file else 'OFFLINE EKA (web)'),
            'marketplaces': 1,
            'total_pos': len(new_rows),
            'total_items': sum(r['items'] for r in new_rows),
            'total_qty': sum(r['qty'] for r in new_rows),
            'total_value': sum(r['order_value'] for r in new_rows),
            'consolidated_path': '', 'tracker_path': '',
        }
        res = H._record(new_rows, run_meta, db_path, skipped=skipped)
        return {'recorded': True, 'skipped': skipped, **res}
    except Exception as e:  # noqa: BLE001
        return {'recorded': False, 'reason': f'DB error: {e}'}


def existing_eka_po_stats() -> Dict:
    """{po: {'sku': items, 'qty': qty}} already recorded for EKA — for the review
    Skipped tab (dedup). Read-only."""
    try:
        from online_b2b.services.order_db import _conn
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(
                f"SELECT po, COALESCE(SUM(items),0), COALESCE(SUM(qty),0) "
                f"FROM order_headers WHERE marketplace={ph} GROUP BY po",
                (eka_engine.EKA_MARKETPLACE,))
            return {str(r[0]): {'sku': int(r[1] or 0), 'qty': int(r[2] or 0)}
                    for r in cur.fetchall()}
    except Exception:  # noqa: BLE001
        return {}
