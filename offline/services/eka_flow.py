"""
offline.services.eka_flow
=========================

EKA adapter for the shared :mod:`online_b2b.services.po_flow` scaffold — the
third offline channel alongside GT Mass and MT.

Each uploaded file is one store; the engine splits it into a regular doc (SO/TO,
finished-goods PO at calculated cost) and a tester doc (``/TT/`` — testers + PWP
+ GWP + non-stock at ₹0.54), assigning SO/TO numbers. Reuses the DB-driven
:mod:`offline.services.eka_bridge` (item master + store registry from the DB, and
records to the shared ``order_headers``).

**No CP check** — EKA has no vendor-price compare, so there are no per-line
decisions (no Exclude, no Override). Every resolvable line is recorded.
"""
from __future__ import annotations

import os
from collections import OrderedDict

from offline.services import eka_bridge, eka_engine


def _doc_lines(results) -> "OrderedDict[str, list]":
    """Group every processed row by its SO/TO number → {to: [line dicts]}, in the
    unified review-payload line shape."""
    docs: "OrderedDict[str, list]" = OrderedDict()
    for res in results:
        for row in (res.regular_orders + res.tester_orders + res.pwp_orders
                    + res.gwp_orders + res.nonstock_orders):
            to = (row.to or '').strip()
            if not to:
                continue
            docs.setdefault(to, []).append({
                'po': to,
                'item_no': row.item_no,
                'ean': row.ean,
                'description': row.product_name or '',
                'qty': int(row.qty or 0),
                'unit_price': round(float(row.unit_price or 0), 2),
                'our_mrp': '',
                'status': row.lookup_status or 'OK',
                'exception_label': ('' if row.source == 'PO' else row.source),
                'key': f"{to}|{row.item_no}|{row.ean}",
            })
    return docs


class EKAFlowProcessor:
    """Flow processor for EKA. ``meta`` carries ``files`` (+ optional warehouse)."""

    def __init__(self, meta: dict):
        self.paths = list(meta.get('files', []))
        self.warehouse = meta.get('warehouse') or 'AHD'

    # ── helpers ──────────────────────────────────────────────────────────
    def _build(self):
        return eka_bridge.process(self.paths, self.warehouse)

    def _payload(self, proc, phase='preview') -> dict:
        results = proc['results']
        rows = eka_engine.build_eka_order_rows(
            results, type_map=eka_bridge._db_type_map(), warehouse=self.warehouse)
        docs = _doc_lines(results)
        existing = eka_bridge.existing_eka_po_stats() if phase == 'preview' else {}

        headers, lines, skipped = [], [], []
        n_pos = n_lines = n_qty = 0
        n_value = 0.0
        n_revised = 0
        for r in rows:
            po = r['po']
            if po in existing and phase == 'preview':
                prev = existing[po]
                identical = (int(r['items']) == prev['sku']
                             and int(r['qty']) == prev['qty'])
                if not identical:
                    n_revised += 1
                skipped.append({
                    'po': po, 'location': r['location'], 'qty': r['qty'],
                    'order_value': r['order_value'],
                    'marketplace_label': r['marketplace_label'],
                    'dup_kind': 'identical' if identical else 'revised',
                    'recorded_sku': prev['sku'], 'recorded_qty': prev['qty'],
                    'incoming_sku': int(r['items']), 'incoming_qty': int(r['qty'])})
                continue
            n_pos += 1
            n_qty += r['qty']
            n_value += r['order_value']
            headers.append({
                'po': po, 'location': r['location'],
                'order_type': r['order_type'],
                'marketplace_label': r['marketplace_label'],
                'items': r['items'], 'qty': r['qty'],
                'order_value': r['order_value']})
            for ln in docs.get(po, []):
                n_lines += 1
                lines.append(ln)

        return {
            'ok': bool(headers) if phase == 'preview' else True,
            'summary': {'pos': n_pos, 'lines': n_lines, 'qty': n_qty,
                        'value': round(n_value, 2),
                        'affected': len(proc['file_issues']),
                        'skipped': len(skipped), 'revised': n_revised},
            'headers': headers, 'lines': lines, 'affected': [],
            'file_issues': proc['file_issues'], 'skipped': skipped,
            'warnings': proc['warnings'], 'output_path': None,
            'error': (None if headers or phase != 'preview'
                      else 'No resolvable orders in the uploaded file(s).'),
        }

    def workbook(self):
        """Build the EKA review workbook on demand (pre-confirm Download). No DB."""
        proc = self._build()
        return eka_bridge.write_review(proc['results'], proc['processing_log'])

    # ── flow protocol ────────────────────────────────────────────────────
    def preview(self) -> dict:
        try:
            proc = self._build()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e), 'summary': {}, 'headers': [],
                    'lines': [], 'affected': [], 'file_issues': [],
                    'skipped': [], 'warnings': []}
        return self._payload(proc, 'preview')

    def confirm(self, actions: dict | None = None) -> dict:
        # EKA has no CP check → `actions` is ignored (nothing to exclude/override).
        try:
            proc = self._build()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e)}
        results = proc['results']
        out = eka_bridge.write_review(proc['results'], proc['processing_log'])
        out_name = os.path.basename(out) if out else ''
        rec = eka_bridge.record(results, output_file=out_name,
                                warehouse=self.warehouse)
        recorded = bool(rec.get('recorded'))
        return {
            'ok': recorded,
            'run_id': rec.get('run_id'),
            'pos': rec.get('new_orders', rec.get('recorded_pos', 0)),
            'lines': sum(len(v) for v in _doc_lines(results).values()),
            'skipped': rec.get('skipped', 0),
            'reason': rec.get('reason', ''),
            'output_path': str(out) if out else None,
            'output_name': out_name or None,
            'error': (None if recorded else rec.get('reason', 'Nothing recorded.')),
        }
