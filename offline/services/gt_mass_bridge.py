"""
offline.services.gt_mass_bridge
===============================

Web-owned **GT Mass → ``renee_orders`` recorder** (preview / confirm), mirroring
the Shoppers-Stop :mod:`offline.services.mt_bridge` pattern so GT Mass shows on
the shared dashboard with the same Orders + Line-Items view as every other
channel.

Design / guarantees
-------------------
* **Dump generator stays intact.** The frozen Tkinter standalone and the web
  "Generate Dump" page (:class:`offline.views.ProcessFilesView` →
  :class:`offline.utils.GTMassAutomation`) are NOT touched — they remain the
  fallback. This module only *adds* a recording path.
* **Parse parity.** We reuse ``GTMassAutomation.process_files`` for the rows /
  SO grouping / warnings, so the recorded set matches the dump exactly.
* **Real value, no margin guesswork.** The raw GT Mass Excel already carries the
  pricing the dump throws away — ``Basic Price`` (unit, ex-GST), ``CLP``
  (= Basic Price × Order Qty, ex-GST), ``GST`` (flat 18 %), and ``TOTAL``
  (= CLP × 1.18, inc-GST). We read those columns: ``order_value`` = Σ ``TOTAL``
  (inc-GST, consistent with the online channels); line ``unit_price`` =
  ``Basic Price``.
* **Web-owned writes.** ``runs`` + ``order_headers`` are written DIRECTLY here
  and ``order_lines`` via :mod:`online_b2b.services.lines_store` — the engine
  history store is never opened, so the retired desktop ``order_issue_lines``
  table is never recreated. PO-level dedup skips SOs already recorded (including
  the ones the desktop app wrote).
"""

from __future__ import annotations

import datetime as _dt
import io
from pathlib import Path

import pandas as pd
from django.conf import settings

MARKETPLACE = 'GT Mass'
SEGMENT = 'Offline'
WAREHOUSE_CODE_MAP = {'AHD': 'PICK', 'BLR': 'DS_BL_OFF1'}
DEFAULT_WAREHOUSE = 'AHD'
GST_RATE = 0.18  # GT Mass is a flat 18% (the file's TOTAL = CLP × 1.18)


# ── headless wrapper so the frozen GTMassAutomation can read saved files ──────
class _DiskUpload:
    """Minimal stand-in for a Django UploadedFile backed by a file on disk, so
    the frozen ``GTMassAutomation`` (which calls ``.seek()``/``.read()``/``.name``)
    can process the token-saved confirm-phase files unchanged."""

    def __init__(self, path):
        self._path = Path(path)
        self.name = self._path.name
        self._bytes = None

    def _load(self):
        if self._bytes is None:
            self._bytes = self._path.read_bytes()
        return self._bytes

    def read(self, *_a, **_k):
        return self._load()

    def seek(self, *_a, **_k):
        return 0


def _automation():
    """Import the frozen web GT Mass engine lazily (kept untouched)."""
    from ..utils import GTMassAutomation
    return GTMassAutomation()


def warehouse_choices() -> list[str]:
    return list(WAREHOUSE_CODE_MAP.keys())


def default_warehouse() -> str:
    return DEFAULT_WAREHOUSE


# ── raw-file price extraction (the columns the dump parser ignores) ──────────
def _num(x):
    try:
        v = float(str(x).replace(',', '').strip())
        return None if v != v else v
    except (TypeError, ValueError):
        return None


def _read_priced_file(path) -> dict:
    """Read ONE raw GT Mass file → its pricing, keyed by BC Code (Item No).

    Returns ``{'so_value': inc_gst_total, 'by_item': {item_no: {basic_price,
    mrp, order_qty, line_total}}}`` — ``so_value`` is Σ TOTAL over ordered rows
    (the file's own inc-GST line totals). Empty/silent-safe: a file we can't
    price contributes 0 (the line audit still records, value just 0)."""
    out = {'so_value': 0.0, 'by_item': {}}
    try:
        df = pd.read_excel(path, header=None)
    except Exception:  # noqa: BLE001
        return out
    # locate the 'BC Code' + 'Order Qty' header row (same rule as the parser)
    hr = None
    for i, rv in enumerate(df.values):
        vals = [str(v).strip().lower() for v in rv]
        if 'bc code' in vals and any('order qty' in v for v in vals):
            hr = i
            break
    if hr is None:
        return out
    cols = [str(c).strip() for c in df.iloc[hr].tolist()]
    idx = {c.lower(): i for i, c in enumerate(cols)}

    def col(*names):
        for n in names:
            if n in idx:
                return idx[n]
        return None

    i_bc = col('bc code')
    i_oq = col('order qty')
    i_mrp = col('mrp')
    i_bp = col('basic price')
    i_clp = col('clp')
    i_tot = col('total')
    if i_bc is None or i_oq is None:
        return out
    total_sum = 0.0
    for rv in df.iloc[hr + 1:].values:
        bc = rv[i_bc]
        if pd.isna(bc):
            continue
        try:
            item_no = str(int(bc))
        except (ValueError, TypeError):
            continue
        oq = _num(rv[i_oq]) or 0
        if oq <= 0:
            continue                       # only ordered lines carry value
        bp = _num(rv[i_bp]) if i_bp is not None else None
        mrp = _num(rv[i_mrp]) if i_mrp is not None else None
        clp = _num(rv[i_clp]) if i_clp is not None else None
        tot = _num(rv[i_tot]) if i_tot is not None else None
        # Prefer the file's own TOTAL (inc-GST); else derive from Basic Price.
        if tot is None:
            line_ex = (clp if clp is not None
                       else (bp * oq if bp is not None else None))
            tot = round(line_ex * (1 + GST_RATE), 2) if line_ex is not None else 0.0
        total_sum += tot or 0.0
        agg = out['by_item'].setdefault(
            item_no, {'basic_price': bp, 'mrp': mrp, 'order_qty': 0,
                      'line_total': 0.0})
        agg['order_qty'] += int(oq)
        agg['line_total'] += tot or 0.0
        if agg['basic_price'] is None and bp is not None:
            agg['basic_price'] = bp
        if agg['mrp'] is None and mrp is not None:
            agg['mrp'] = mrp
    out['so_value'] = round(total_sum, 2)
    return out


def _price_index(paths) -> dict:
    """``{file_basename: _read_priced_file(...)}`` for every uploaded file."""
    out = {}
    for p in paths:
        out[Path(p).name] = _read_priced_file(p)
    return out


# ── core: process + (optionally) record ──────────────────────────────────────
class GTMassRecorder:
    def __init__(self, po_paths, warehouse: str | None = None):
        self.po_paths = [str(p) for p in (po_paths or [])]
        self.warehouse = warehouse or DEFAULT_WAREHOUSE
        self.result = None
        self.prices = {}

    def _process(self):
        """Run the frozen dump engine over the files (no DB) + read prices."""
        if not self.po_paths:
            raise ValueError('No PO file uploaded.')
        uploads = [_DiskUpload(p) for p in self.po_paths]
        self.result = _automation().process_files(uploads)
        self.prices = _price_index(self.po_paths)
        self._ean_fallback()
        return self.result

    # ── EAN-only fallback (files with no BC Code column) ────────────────
    def _ean_fallback(self):
        """Rescue files the standard parser rejected for a missing BC Code column
        but which DO carry EAN + Order Qty (e.g. the Indian-Secrets pack format):
        resolve Item No from the item master by EAN. Resolved lines join the
        recorded set; EANs not in the master become explicit warnings (never a
        silent drop). Files that are genuinely broken stay failed. The frozen
        dump generator is NOT affected — this only enriches the recorder."""
        res = self.result
        failed = list(getattr(res, 'failed_files', None) or [])
        if not failed:
            return
        try:
            from online_b2b.services import item_master_loader as iml
        except Exception:  # noqa: BLE001
            return
        rescued = []
        for fname, reason in failed:
            if 'header' not in str(reason).lower():
                continue                       # not a missing-header rejection
            path = next((p for p in self.po_paths if Path(p).name == fname), None)
            if not path:
                continue
            parsed = self._parse_ean_only(path, fname, iml)
            if parsed is None:
                continue                       # not an EAN-only file → leave failed
            rows, price, warns = parsed
            res.rows.extend(rows)
            self.prices[fname] = price
            for w in warns:
                res.warned_files.append((fname, w))
            rescued.append((fname, reason))
        if rescued:
            res.failed_files = [t for t in failed if t not in rescued]

    def _parse_ean_only(self, path, fname, iml):
        """Parse a BC-Code-less GT Mass file (EAN + Order Qty header). Returns
        ``(order_rows, price_dict, warnings)`` or ``None`` if the file isn't an
        EAN-only GT Mass sheet. item_no is resolved from the master via EAN."""
        from ..utils import (
            LOCATION_CODE_MAP,
            OrderRow,
            SONumberFormatter,
        )
        try:
            df = pd.read_excel(path, header=None)
        except Exception:  # noqa: BLE001
            return None
        hr = None
        for i, rv in enumerate(df.values):
            low = [str(v).strip().lower() for v in rv]
            if 'ean' in low and any('order qty' in v for v in low) \
                    and 'bc code' not in low:
                hr = i
                break
        if hr is None:
            return None                        # genuinely not EAN-only GT Mass
        cols = [str(c).strip().lower() for c in df.iloc[hr].tolist()]
        idx = {c: j for j, c in enumerate(cols)}

        def col(*names):
            for n in names:
                if n in idx:
                    return idx[n]
            return None

        i_ean = col('ean'); i_oq = col('order qty')
        i_tq = col('tester qty'); i_cat = col('category')
        i_desc = col('article description', 'description')
        i_mrp = col('mrp'); i_bp = col('basic price')
        i_tot = col('total')
        # meta: scan rows above the header for PO Number / Location / Distributor
        so = location = distributor = city = state = ''
        for rv in df.iloc[:hr].values:
            for j in range(min(len(rv) - 1, 10)):
                lab = str(rv[j]).strip().lower()
                if not lab or lab == 'nan':
                    continue
                nxt = ''
                for k in range(j + 1, min(j + 3, len(rv))):
                    if pd.notna(rv[k]) and str(rv[k]).strip().lower() not in ('', 'nan'):
                        nxt = str(rv[k]).strip(); break
                if lab == 'po number' and not so:
                    so = nxt
                elif lab == 'location' and not location:
                    location = nxt
                elif lab == 'distributor name' and not distributor:
                    distributor = nxt
                elif lab == 'city' and not city:
                    city = nxt
                elif lab == 'state':
                    state = nxt or state
        if not so:
            so = SONumberFormatter.from_filename(fname) or 'SO/GTM/UNKNOWN'
        loc_code = LOCATION_CODE_MAP.get(location.upper().strip(), location) if location else ''

        rows, warns = [], []
        price = {'so_value': 0.0, 'by_item': {}}
        total_sum = 0.0
        unresolved = 0
        for rv in df.iloc[hr + 1:].values:
            if i_ean is None:
                break
            ean = rv[i_ean]
            if pd.isna(ean):
                continue
            ean = str(ean).split('.')[0].strip()
            oq = int(_num(rv[i_oq]) or 0) if i_oq is not None else 0
            tq = int(_num(rv[i_tq]) or 0) if i_tq is not None else 0
            if oq <= 0 and tq <= 0:
                continue
            hit = iml.resolve_in_master(ean)
            if not hit or not hit.get('item_no'):
                unresolved += 1
                warns.append(f"EAN {ean} not in item master — line skipped "
                             f"(add this SKU to the master to record it).")
                continue
            item_no = str(hit['item_no'])
            bp = _num(rv[i_bp]) if i_bp is not None else None
            mrp = _num(rv[i_mrp]) if i_mrp is not None else (
                float(hit['mrp']) if hit.get('mrp') is not None else None)
            tot = _num(rv[i_tot]) if i_tot is not None else None
            if tot is None and bp is not None:
                tot = round(bp * oq * (1 + GST_RATE), 2)
            total_sum += tot or 0.0
            agg = price['by_item'].setdefault(
                item_no, {'basic_price': bp, 'mrp': mrp, 'order_qty': 0,
                          'line_total': 0.0})
            agg['order_qty'] += oq
            agg['line_total'] += tot or 0.0
            rows.append(OrderRow(
                so_number=so, item_no=item_no, ean=ean,
                category=(str(rv[i_cat]).strip() if i_cat is not None and pd.notna(rv[i_cat]) else ''),
                description=(hit.get('description') or
                             (str(rv[i_desc]).strip() if i_desc is not None and pd.notna(rv[i_desc]) else ''))[:255],
                qty=oq, tester_qty=tq, distributor=distributor, city=city,
                state=state, location=location, location_code=loc_code,
                source_file=fname))
        price['so_value'] = round(total_sum, 2)
        if rows:
            warns.insert(0, f"EAN-only format (no BC Code): resolved "
                         f"{len(rows)} line(s) from the master by EAN"
                         + (f", {unresolved} unresolved." if unresolved else "."))
        return rows, price, warns

    # ── per-SO aggregation shared by preview + confirm ──────────────────
    def _orders(self) -> dict:
        """``{so_number: {...header..., _lines:[...]}}`` from the parsed rows +
        the raw-file price index. One header per SO (GT Mass = one PO/file)."""
        from collections import OrderedDict
        orders: OrderedDict[str, dict] = OrderedDict()
        for r in (getattr(self.result, 'rows', None) or []):
            so = (r.so_number or '').strip()
            oqty = int(r.qty or 0)
            tqty = int(r.tester_qty or 0)
            if not so or (oqty <= 0 and tqty <= 0):
                continue
            o = orders.get(so)
            if o is None:
                o = {
                    'po': so, 'location': '', 'items': 0, 'qty': 0,
                    'order_value': 0.0, 'source_files': set(), '_lines': [],
                }
                orders[so] = o
            if not o['location'] and (r.distributor or '').strip():
                o['location'] = r.distributor.strip()
            o['source_files'].add(r.source_file)
            o['items'] += 1
            o['qty'] += oqty                       # ERP order qty (testers auto-added)
            pinfo = self.prices.get(r.source_file, {}).get('by_item', {}).get(
                str(r.item_no), {})
            unit = pinfo.get('basic_price')
            o['_lines'].append({
                'item_no': str(r.item_no or ''), 'ean': str(r.ean or ''),
                'description': (r.description or '')[:255],
                'order_qty': oqty, 'tester_qty': tqty,
                'unit_price': round(unit, 2) if unit is not None else None,
                'mrp': pinfo.get('mrp'),
                'category': r.category or '',
            })
        # order_value = Σ the SO's file TOTAL(s) (inc-GST, the file's own figure)
        for o in orders.values():
            o['order_value'] = round(
                sum(self.prices.get(f, {}).get('so_value', 0.0)
                    for f in o['source_files']), 2)
        return orders

    # ── phase 1: preview (no DB writes) ─────────────────────────────────
    def preview(self) -> dict:
        try:
            self._process()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e), 'phase': 'preview'}
        return self._summary(self._orders(), recorded=None, phase='preview')

    # ── phase 2: confirm (dedup + write runs/headers/lines + dump) ──────
    def confirm(self) -> dict:
        try:
            self._process()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e), 'phase': 'confirm'}
        orders = self._orders()
        recorded = self._record(orders)
        out_path = self._write_dump()
        summ = self._summary(orders, recorded=recorded, phase='confirm')
        summ['output_path'] = str(out_path) if out_path else None
        summ['output_name'] = out_path.name if out_path else None
        return summ

    def _record(self, orders: dict) -> dict:
        """Web-owned write of ``runs`` + ``order_headers`` + ``order_lines`` with
        PO-level dedup. Soft-fails (never blocks the dump)."""
        from online_b2b.services import lines_store
        from online_b2b.services.order_db import _conn
        if not orders:
            return {'recorded': False, 'reason': 'no clean orders to record'}
        try:
            with _conn() as (cur, d):
                ph = d['ph']
                cur.execute(f"SELECT DISTINCT po FROM order_headers WHERE "
                            f"marketplace={ph}", (MARKETPLACE,))
                existing = {str(r[0]) for r in cur.fetchall()}
        except Exception as e:  # noqa: BLE001
            return {'recorded': False, 'reason': f'DB read failed: {e}'}

        new = {so: o for so, o in orders.items() if so not in existing}
        skipped = len(orders) - len(new)
        if not new:
            return {'recorded': False, 'reason': 'all POs already recorded',
                    'skipped': skipped}

        run_ts = _dt.datetime.now()
        out_name = self._dump_name()
        wh = self.warehouse or ''
        po_date = _dt.date.today()
        total_pos = len(new)
        total_items = sum(o['items'] for o in new.values())
        total_qty = sum(o['qty'] for o in new.values())
        total_value = round(sum(o['order_value'] for o in new.values()), 2)
        try:
            with _conn() as (cur, d):
                ph = d['ph']
                cur.execute(
                    f"INSERT INTO runs (run_ts, mode, source, marketplaces, "
                    f"total_pos, total_items, total_qty, total_value, "
                    f"consolidated_path, tracker_path) VALUES "
                    f"({ph},'MANUAL',{ph},1,{ph},{ph},{ph},{ph},'','')",
                    (run_ts, f'OFFLINE GT MASS (web): {out_name}', total_pos,
                     total_items, total_qty, total_value))
                run_id = cur.lastrowid
                hcols = ('run_id, run_ts, mode, segment, marketplace, '
                         'marketplace_label, po, location, warehouse, po_date, '
                         'exp_date, order_type, items, qty, order_value, '
                         'output_file')
                marks = ', '.join([ph] * 16)
                for so, o in new.items():
                    cur.execute(
                        f"INSERT INTO order_headers ({hcols}) VALUES ({marks})",
                        (run_id, run_ts, 'MANUAL', SEGMENT, MARKETPLACE,
                         MARKETPLACE, so, o['location'], wh, po_date, None,
                         'SO', o['items'], o['qty'], o['order_value'], out_name))
                cur.connection.commit()
        except Exception as e:  # noqa: BLE001
            return {'recorded': False, 'reason': f'DB write failed: {e}'}

        # order_lines audit (the gap the desktop never filled).
        run_ts_s = run_ts.strftime('%Y-%m-%d %H:%M:%S')
        rows = []
        for so, o in new.items():
            for ln in o['_lines']:
                unit = ln['unit_price']
                tnote = (f"Tester: {ln['tester_qty']}" if ln['tester_qty'] else '')
                rows.append({
                    'run_id': run_id, 'run_ts': run_ts_s,
                    'marketplace': MARKETPLACE, 'po': so,
                    'location': o['location'], 'item_no': ln['item_no'],
                    'ean': ln['ean'], 'description': ln['description'],
                    'qty': ln['order_qty'], 'order_type': 'SO', 'gst_code': '18%',
                    'unit_price': unit, 'output_file': out_name,
                    'our_mrp': ln['mrp'], 'vendor_mrp': None,
                    'our_landing': unit, 'vendor_landing': None,
                    'our_cp': None, 'vendor_cp': None, 'diff': None,
                    'margin_pct': None, 'status': 'OK', 'exception_label': '',
                    'received_ean': None, 'action': '', 'remark': tnote,
                })
        try:
            lines_store.insert_lines(run_id, rows)
        except Exception as e:  # noqa: BLE001 — never block on the audit
            return {'recorded': True, 'run_id': run_id, 'recorded_pos': len(new),
                    'skipped': skipped, 'lines': 0,
                    'reason': f'lines audit skipped: {e}'}
        return {'recorded': True, 'run_id': run_id, 'recorded_pos': len(new),
                'skipped': skipped, 'lines': len(rows)}

    # ── dump (re-uses the frozen exporter; output identical to the page) ─
    def _dump_name(self) -> str:
        return f"gt_mass_dump_{_dt.datetime.now().strftime('%d%m%Y')}.xlsx"

    def _write_dump(self):
        """Write the SAME 7-sheet dump the existing page produces, to MEDIA so it
        can be downloaded post-confirm. Uses the frozen exporter unchanged."""
        try:
            buf = _automation().exporter.export_to_memory(self.result)
            if buf is None:
                return None
            out_dir = Path(settings.MEDIA_ROOT) / 'gt_mass_out'
            out_dir.mkdir(parents=True, exist_ok=True)
            path = out_dir / f"gt_mass_dump_{_dt.datetime.now():%d%m%Y_%H%M%S}.xlsx"
            path.write_bytes(buf.getvalue() if isinstance(buf, io.BytesIO)
                             else buf.read())
            return path
        except Exception:  # noqa: BLE001 — dump is a convenience, never blocks
            return None

    # ── web summary payload (shape matches the SS template renderer) ────
    def _summary(self, orders: dict, recorded=None, phase='confirm') -> dict:
        res = self.result
        warnings = []
        for fname, w in (getattr(res, 'warned_files', None) or []):
            warnings.append(f"{fname}: {w}")
        for fname, r in (getattr(res, 'failed_files', None) or []):
            warnings.append(f"[FAILED] {fname}: {r}")

        pos = []
        total_qty = total_val = 0
        for so, o in orders.items():
            total_qty += o['qty']
            total_val += o['order_value']
            pos.append({
                'file': ', '.join(sorted(o['source_files'])),
                'po': so, 'store': o['location'], 'so_number': so,
                'ship_to': '', 'cust_no': '',
                'lines': o['items'], 'qty': o['qty'],
                'value': round(o['order_value'], 2),
                'status': 'SO' if phase == 'confirm' else 'READY',
            })
        ok = bool(orders) if phase == 'preview' else bool((recorded or {}).get('recorded'))
        rec_info = recorded or {}
        return {
            'ok': ok if phase == 'preview' else True,
            'phase': phase, 'channel': MARKETPLACE, 'channel_code': 'GTM',
            'warehouse': self.warehouse,
            'pos': pos,
            'summary': {
                'files': len(getattr(res, 'attempted_files', []) or []),
                'sos': len(orders),
                'errors': len(getattr(res, 'failed_files', []) or []),
                'lines': sum(p['lines'] for p in pos),
                'qty': total_qty, 'value': round(total_val, 2),
            },
            'recorded': bool(rec_info.get('recorded')),
            'run_id': rec_info.get('run_id'),
            'recorded_pos': rec_info.get('recorded_pos'),
            'recorded_skipped': rec_info.get('skipped'),
            'recorded_reason': rec_info.get('reason'),
            'warnings': warnings,
            'error': (None if (ok or phase != 'preview')
                      else 'No resolvable POs in the file(s).'),
        }


def preview(po_paths, warehouse: str | None = None) -> dict:
    return GTMassRecorder(po_paths, warehouse).preview()


def confirm(po_paths, warehouse: str | None = None) -> dict:
    return GTMassRecorder(po_paths, warehouse).confirm()
