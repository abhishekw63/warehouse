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
            # Never-silent info: what this channel demands + PDF cross-check result.
            'notes': list(getattr(self._mt, 'notes', []) or []),
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
