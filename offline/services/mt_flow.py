"""
offline.services.mt_flow
========================

Modern-Trade (MT) adapter for the shared :mod:`online_b2b.services.po_flow`
scaffold — so every MT child channel (Shoppers Stop, Health & Glow, Naturals…)
gets the SAME upload → review → **lock** → record-affected experience as the
online marketplaces, instead of the old single-page SS generator.

ADDITIVE — this does NOT modify the frozen MT-Select desktop automation nor the
:class:`mt_bridge.MTProcessor`; it *orchestrates* the bridge (``_load`` for a
no-write preview, ``confirm`` for the assign+write+record) and reshapes the
output into the flow's unified review payload.

Per-line decisions: **Exclude** only (SS-family files carry no vendor cost, so
there is no vendor-price compare / Override). Excluded lines are dropped before
SO numbers are burned. The channel is chosen at upload (the flow's
``marketplace`` capability → ``meta['marketplace']`` = the MT channel code).

NOTE — no dedup/skipped tab for MT: the DB records the assigned SO number, not
the store PO, so there is nothing to dedup a fresh upload against (same as the
old SS flow). Re-uploading the same file mints new SOs.
"""
from __future__ import annotations

import os

from . import mt_bridge
from .mt_bridge import MTProcessor, line_key


def _f(x):
    try:
        return round(float(x), 2) if x not in (None, '') else None
    except (TypeError, ValueError):
        return None


def _first_error(pf) -> str:
    for lvl, msg in getattr(pf, 'findings', []) or []:
        if lvl == 'error':
            return str(msg)
    return ''


class MTFlowProcessor:
    """Flow processor for MT. ``meta`` carries ``files`` + ``warehouse`` +
    ``marketplace`` (the MT channel code, e.g. 'SS' / 'HG')."""

    def __init__(self, meta: dict):
        self.channel_code = meta.get('marketplace') or 'SS'
        allf = list(meta.get('files', []))
        # Auto-separate a tester-requirement sheet from the PO file(s) — testers
        # generate alongside regulars. Two layouts are recognised (by COLUMNS,
        # robust to filenames): the HG (Store code + SKU + Tester) dump AND the
        # LS store-grouped sheet (STORE code + EAN + Tester Req).
        from .tester import is_ls_tester_file
        self.tester_file = next(
            (p for p in allf
             if mt_bridge.is_tester_file(p) or is_ls_tester_file(p)), None)
        self.paths = [p for p in allf if p != self.tester_file]
        self.warehouse = meta.get('warehouse') or None
        self._mt = MTProcessor(self.channel_code, self.paths, self.warehouse,
                               tester_file=self.tester_file)

    # ── unified payload from a parsed (no-write) batch ───────────────────
    def _payload(self, channel, batch, phase='preview') -> dict:
        ratio = getattr(channel, 'expected_landing_ratio', None)
        headers, lines, affected, file_issues, warnings = [], [], [], [], []
        skipped = []
        npos = total_qty = 0
        total_val = 0.0
        # Already-uploaded POs (by External Doc No = store PO) — a re-uploaded
        # dump. These are skipped (won't be re-recorded / re-minted).
        recorded_ext = self._mt._recorded_ext_docs()
        for lvl, msg in getattr(batch, 'cross_findings', []) or []:
            warnings.append(f"[{lvl}] {msg}")
        for pf in batch.po_files:
            for lvl, msg in getattr(pf, 'findings', []) or []:
                if lvl in ('error', 'warn', 'warning'):
                    warnings.append(f"{pf.source_name}: [{lvl}] {msg}")
            if getattr(pf, 'has_hard_errors', False):
                file_issues.append({
                    'file': pf.source_name, 'problem': 'Could not process',
                    'detail': _first_error(pf) or 'File has hard errors.',
                    'kind': 'error'})
                continue
            po = str(pf.po_no or pf.source_name)
            qty = int(getattr(pf, 'input_qty_total', 0) or 0)
            val = float(getattr(pf, 'input_po_value_total', 0) or 0)
            # Prefer the readable store name (e.g. 'HG-VIVACITY-MUM') over the
            # bare customer number for the Orders 'Location' column.
            raw_loc = str(getattr(pf, 'store_name', '') or '')
            ship_to = str(getattr(pf, 'ship_to', '') or '')
            loc = raw_loc or ship_to
            if po in recorded_ext:
                skipped.append({'po': po, 'location': loc, 'qty': qty,
                                'order_value': round(val, 2),
                                'marketplace_label': channel.display_name})
                continue
            npos += 1
            total_qty += qty
            total_val += val
            # Expose the ship-to resolution so the review can flag UNMAPPED stores
            # (parity with Online B2B's Mapping tab). mapped=False → no ship-to
            # resolved → the SO can't reach D365; the review banner warns on it.
            headers.append({'po': po, 'location': loc, 'order_type': 'SO',
                            'items': len(pf.lines), 'qty': qty,
                            'order_value': round(val, 2),
                            'raw_location': raw_loc or loc,
                            'ship_to': ship_to, 'mapped': bool(ship_to)})
            for ln in pf.lines:
                item_no = str(getattr(ln, 'item_no', '') or '')
                ean = str(getattr(ln, 'ean', '') or '')
                desc = str(getattr(ln, 'items_master_desc', None)
                           or getattr(ln, 'sku_name', '') or '')[:255]
                our_mrp = _f(getattr(ln, 'items_master_mrp', None))
                our_landing = (round(our_mrp * ratio, 2)
                               if ratio and our_mrp else None)
                resolved = bool(item_no) and getattr(ln, 'status', '') != 'SKIP'
                row = {
                    'po': po, 'item_no': item_no, 'ean': ean,
                    'description': desc,
                    'qty': int(getattr(ln, 'quantity', 0) or 0),
                    'unit_price': our_landing, 'our_mrp': our_mrp,
                    'status': 'OK' if resolved else 'NOT_IN_MASTER',
                    'exception_label': '',
                    'key': line_key(po, item_no, ean),
                }
                lines.append(row)
                if not resolved:
                    affected.append(row)
        # H&B is DC-routed: many franchisee stores of one DC share a single
        # delivery address (the DC), so the Site code is the ship-to key and the
        # PO PDF exists to VERIFY that address per PO — spell this out for the
        # operator so the shared ship-to never looks like a mistake.
        notes = list(getattr(self._mt, 'notes', []) or [])
        if self.channel_code == 'HB':
            notes.insert(0, (
                "H&B is DC-routed — several franchisee stores of a DC are "
                "delivered to ONE address (the DC itself), so the file's Site "
                "code (e.g. D009 = Sahibabad DC → 20040_73) is the ship-to key "
                "and many franchisees (Sri Ganga Nagar, Arera Colony…) correctly "
                "roll up to the SAME ship-to. The PO PDF is uploaded alongside "
                "the .xlsb only to CONFIRM that delivery address: each PO's PDF "
                "delivery pincode is cross-checked against the mapped ship-to, so "
                "a wrong address is caught (never trusted blindly). A never-seen "
                "Site code still flags UNMAPPED. Always include the PO PDFs."))
        ok = bool(headers) if phase == 'preview' else True
        if skipped:
            warnings.insert(0, f"{len(skipped)} PO(s) already uploaded "
                            f"(External Doc) — will be skipped: "
                            f"{', '.join(s['po'] for s in skipped[:10])}"
                            f"{'…' if len(skipped) > 10 else ''}.")
        return {
            'ok': ok if (headers or not skipped) else False,
            'summary': {'pos': npos, 'lines': len(lines), 'qty': total_qty,
                        'value': round(total_val, 2),
                        'affected': len(affected) + len(file_issues),
                        'skipped': len(skipped)},
            'headers': headers, 'lines': lines, 'affected': affected,
            'file_issues': file_issues, 'skipped': skipped, 'warnings': warnings,
            # Never-silent info: what this channel demands + PDF cross-check result
            # (+ the H&B DC-routing explainer prepended above, HB only).
            'notes': notes,
            'requirements': mt_bridge.channel_requirements(self.channel_code),
            # Channel-agnostic "additional verification" (online_b2b.services.
            # verification). Present only when a channel produced it (LS is the
            # first consumer). The review page shows just a link + summary; the
            # dedicated verification page renders the full table.
            'verification': getattr(batch, 'verification', None),
            'output_path': None,
            'error': (None if headers else
                      ('All PO(s) already uploaded (External Doc).' if skipped
                       else ('No resolvable POs in the uploaded file(s).'
                             if phase == 'preview' else None))),
        }

    # ── 'download' cap: full 8/9-sheet SO Workbook during REVIEW (pre-lock) ──
    def workbook(self):
        """Build the unified 8/9-sheet SO Workbook (same as the post-lock download
        and Online B2B) from the preview batch, so it's downloadable DURING review.
        No SO numbers burned, no DB write — SO No. is blank until Confirm."""
        return self._mt.preview_workbook()

    # ── flow protocol ────────────────────────────────────────────────────
    def preview(self) -> dict:
        try:
            _eng, channel, batch = self._mt._load()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e), 'summary': {}, 'headers': [],
                    'lines': [], 'affected': [], 'file_issues': [],
                    'skipped': [], 'warnings': []}
        payload = self._payload(channel, batch, phase='preview')
        # Tester summary (SELECTIVE) — shown so the operator sees what will be
        # appended before confirming. No SO numbers assigned here.
        if self.tester_file:
            tp = self._mt.tester_preview()
            if tp:
                payload['testers'] = tp
                if tp.get('error'):
                    payload.setdefault('warnings', []).append(
                        f"Testers: {tp['error']}")
                else:
                    payload.setdefault('warnings', []).insert(
                        0, f"Testers: +{tp['lines']} tester line(s) @ ₹{tp['price']} "
                        f"across {tp['sos']} store(s) = ₹{tp['value']} "
                        f"(from {os.path.basename(str(self.tester_file))}).")
                    # Surface each never-silent tester warning (unresolved EAN,
                    # missing ship-to, non-Approved remark) on the review page.
                    for w in tp.get('warnings', []) or []:
                        payload.setdefault('warnings', []).append(f"Testers: {w}")
        return payload

    def confirm(self, actions: dict | None = None) -> dict:
        excluded = {k for k, v in (actions or {}).items()
                    if (v or {}).get('action') == 'EXCLUDE'}
        res = self._mt.confirm(exclude_keys=excluded)
        run_id = res.get('run_id')
        summ = res.get('summary', {}) or {}
        return {
            'ok': bool(run_id),
            'run_id': run_id,
            'pos': res.get('recorded_pos') or 0,
            'lines': summ.get('lines', 0),
            'output_path': res.get('output_path'),
            'output_name': res.get('output_name'),
            'warnings': res.get('warnings', []),
            'error': (None if run_id else
                      (res.get('error') or res.get('recorded_reason')
                       or 'Workbook written but nothing recorded to the DB.')),
        }


class RelianceTrendsFlowProcessor:
    """MT-flow processor for **Reliance Trends** (BAP Excel). Same flow protocol as
    :class:`MTFlowProcessor` (preview / workbook / confirm) so it slots into the MT
    upload page's channel dropdown — but routes to the standalone
    :mod:`offline.services.reliance_trends_bridge` because the BAP SAP-export format
    is NOT one the frozen MT engine parses. cust 20418; BAP → Bhiwandi 20418_2;
    Unit Price left blank (D365 auto-prices)."""

    def __init__(self, meta: dict):
        self.channel_code = 'RT'
        self.paths = list(meta.get('files', []))
        self.warehouse = meta.get('warehouse') or 'AHD'
        self._path = self.paths[0] if self.paths else None
        self._parsed = None

    def _parse(self):
        if self._parsed is None:
            from .reliance_trends_bridge import parse
            self._parsed = parse(self._path) if self._path else {
                'ok': False, 'error': 'No file uploaded.', 'pos': {}}
        return self._parsed

    def _recorded_pos(self):
        from .reliance_trends_bridge import MARKETPLACE
        try:
            from online_b2b.services.order_db import _conn
            with _conn() as (cur, d):
                ph = d['ph']
                cur.execute(f"SELECT DISTINCT po FROM order_headers WHERE "
                            f"marketplace={ph}", (MARKETPLACE,))
                return {str(r[0]) for r in cur.fetchall()}
        except Exception:  # noqa: BLE001
            return set()

    def preview(self) -> dict:
        from .reliance_trends_bridge import MARKETPLACE
        p = self._parse()
        if not p.get('ok'):
            return {'ok': False, 'error': p.get('error'), 'summary': {},
                    'headers': [], 'lines': [], 'affected': [], 'file_issues': [],
                    'skipped': [], 'warnings': []}
        recorded = self._recorded_pos()
        headers, lines, affected, skipped = [], [], [], []
        npos = tq = 0
        tv = 0.0
        for po, pd in p['pos'].items():
            loc = f"{pd['city']} ({pd['ship_to']})"
            if po in recorded:
                skipped.append({'po': po, 'location': loc, 'qty': pd['qty'],
                                'order_value': pd['value'],
                                'marketplace_label': MARKETPLACE})
                continue
            npos += 1
            tq += pd['qty']
            tv += pd['value']
            headers.append({'po': po, 'location': loc, 'order_type': 'SO',
                            'items': len(pd['lines']), 'qty': pd['qty'],
                            'order_value': pd['value'], 'raw_location': pd['city'],
                            'ship_to': pd['ship_to'], 'mapped': bool(pd['ship_to'])})
            for ln in pd['lines']:
                resolved = bool(ln['item_no'])
                row = {'po': po, 'item_no': ln['item_no'], 'ean': ln['ean'],
                       'description': ln['description'], 'qty': ln['qty'],
                       'unit_price': None, 'our_mrp': None,
                       'status': 'OK' if resolved else 'NOT_IN_MASTER',
                       'exception_label': '',
                       'key': line_key(po, ln['item_no'], ln['ean'])}
                lines.append(row)
                if not resolved:
                    affected.append(row)
        warnings = list(p.get('warnings', []))
        if skipped:
            warnings.insert(0, f"{len(skipped)} PO(s) already recorded — skipped: "
                            f"{', '.join(s['po'] for s in skipped[:10])}.")
        return {
            'ok': bool(headers),
            'summary': {'pos': npos, 'lines': len(lines), 'qty': tq,
                        'value': round(tv, 2), 'affected': len(affected),
                        'skipped': len(skipped)},
            'headers': headers, 'lines': lines, 'affected': affected,
            'file_issues': [], 'skipped': skipped, 'warnings': warnings,
            'notes': ['Reliance Trends (cust 20418) — BAP replenishment PO → '
                      'Bhiwandi (ship-to 20418_2). Value inc-GST; Unit Price left '
                      'blank in the SO (D365 auto-prices).'],
            'requirements': {'required': 'Reliance Trends BAP Excel (Purchasing '
                             'document, EAN, PO Qty, Net Value / Total CP).'},
            'verification': None, 'output_path': None,
            'error': (None if headers else
                      ('All PO(s) already recorded.' if skipped
                       else 'No POs found in the uploaded file.')),
        }

    def workbook(self):
        if not self._path:
            return None
        from .reliance_trends_bridge import build_workbook
        out, _err = build_workbook(self._path, warehouse=self.warehouse)
        return str(out) if out else None

    def confirm(self, actions: dict | None = None) -> dict:
        if not self._path:
            return {'ok': False, 'error': 'No file uploaded.', 'run_id': None,
                    'pos': 0, 'lines': 0}
        from .reliance_trends_bridge import record, build_workbook
        res = record(self._path, warehouse=self.warehouse,
                     source_file=os.path.basename(self._path))
        out, _err = build_workbook(self._path, warehouse=self.warehouse,
                                   so_map=res.get('so_map'))
        recorded_ok = bool(res.get('ok') and (res.get('recorded') or out))
        return {
            'ok': recorded_ok,
            'run_id': res.get('run_id'),
            'pos': res.get('recorded_pos') or 0,
            'lines': res.get('lines') or 0,
            'output_path': str(out) if out else None,
            'output_name': os.path.basename(str(out)) if out else None,
            'warnings': ([] if res.get('recorded') else
                         [res.get('reason')] if res.get('reason') else []),
            'error': (res.get('error') if not res.get('ok') else None),
        }
