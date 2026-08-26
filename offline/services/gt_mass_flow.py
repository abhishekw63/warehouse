"""
offline.services.gt_mass_flow
=============================

GT Mass adapter for the shared :mod:`online_b2b.services.po_flow` scaffold.

ADDITIVE — this does NOT modify :class:`gt_mass_bridge.GTMassRecorder`; it
*orchestrates* the recorder's existing methods (``_process`` / ``_orders`` /
``_record`` / ``_write_dump``) and reshapes the output into the flow's unified
review payload, plus GT-Mass-specific **file-level exceptions** (PO-number
missing, template mismatch, EAN-only rescue) that other channels don't have.

Per-line decisions supported: **Exclude** only (GT Mass has no vendor-price
compare, so no Override). Excluded lines are dropped before recording.
"""
from __future__ import annotations

from .gt_mass_bridge import MARKETPLACE, GTMassRecorder


def _line_key(po: str, item_no: str, ean: str) -> str:
    return f"{po}|{item_no}|{ean}"


def _existing_pos() -> set:
    """SOs already in the DB for GT Mass (read-only — for the Skipped tab)."""
    return set(_existing_po_stats().keys())


def _existing_po_stats() -> dict:
    """``{po: {'sku': items, 'qty': qty}}`` already recorded for GT Mass. Lets the
    review tell a **true duplicate** (same PO number AND same SKU count + total
    qty → safe to skip, no double-count) from a **revision** (same PO number but
    changed content → must be surfaced, never silently dropped). Read-only."""
    try:
        from online_b2b.services.order_db import _conn
        with _conn() as (cur, d):
            ph = d['ph']
            cur.execute(
                f"SELECT po, COALESCE(SUM(items),0), COALESCE(SUM(qty),0) "
                f"FROM order_headers WHERE marketplace={ph} GROUP BY po",
                (MARKETPLACE,))
            return {str(r[0]): {'sku': int(r[1] or 0), 'qty': int(r[2] or 0)}
                    for r in cur.fetchall()}
    except Exception:  # noqa: BLE001
        return {}


def _classify(reason: str) -> str:
    """Map a parser failure reason to an operator-facing problem label."""
    r = (reason or '').lower()
    if any(k in r for k in ('bc code', 'header', 'column', 'template', 'format')):
        return 'Template mismatch'
    if any(k in r for k in ('po', 'so number', 'so_number', 'order no')):
        return 'PO number missing'
    return 'Parse error'


def _file_issues(result) -> list:
    """GT-Mass file-level exceptions for the channel slot panel."""
    issues = []
    for fname, reason in (getattr(result, 'failed_files', None) or []):
        issues.append({'file': fname, 'problem': _classify(reason),
                       'detail': str(reason), 'kind': 'error'})
    for fname, warn in (getattr(result, 'warned_files', None) or []):
        issues.append({'file': fname, 'problem': 'Rescued / warning',
                       'detail': str(warn), 'kind': 'warn'})
    return issues


class GTMassProcessor:
    """Flow processor for GT Mass. ``meta`` carries ``files`` + ``warehouse``."""

    def __init__(self, meta: dict):
        self.paths = list(meta.get('files', []))
        self.warehouse = meta.get('warehouse') or None

    # ── helpers ──────────────────────────────────────────────────────────
    def _build(self):
        rec = GTMassRecorder(self.paths, self.warehouse)
        rec._process()
        return rec, rec._orders()

    def _payload(self, rec, orders, recorded=None, phase='preview',
                 output_path=None):
        existing = _existing_po_stats()
        headers, lines, skipped, blocked = [], [], [], []
        new_pos = new_lines = new_qty = 0
        new_value = 0.0
        n_revised = 0
        for so, o in orders.items():
            if o.get('blocked'):        # HELD BACK — never offered for recording
                blocked.append({'po': so, 'location': o.get('location'),
                                'qty': o['qty'], 'order_value': o['order_value'],
                                'marketplace_label': MARKETPLACE,
                                'reasons': o.get('block_reasons') or []})
                continue
            if so in existing and phase == 'preview':
                prev = existing[so]
                in_sku, in_qty = int(o['items']), int(o['qty'])
                identical = (in_sku == prev['sku'] and in_qty == prev['qty'])
                if not identical:
                    n_revised += 1
                skipped.append({'po': so, 'location': o['location'],
                                'qty': o['qty'], 'order_value': o['order_value'],
                                'marketplace_label': MARKETPLACE,
                                # identical → safe duplicate; revised → same PO no.
                                # but changed SKU count / qty → needs a decision.
                                'dup_kind': 'identical' if identical else 'revised',
                                'recorded_sku': prev['sku'], 'recorded_qty': prev['qty'],
                                'incoming_sku': in_sku, 'incoming_qty': in_qty})
                continue
            new_pos += 1
            new_qty += o['qty']
            new_value += o['order_value']
            headers.append({'po': so, 'location': o['location'],
                            'warehouse': o.get('warehouse', ''),
                            'order_type': 'SO', 'items': o['items'],
                            'qty': o['qty'], 'order_value': o['order_value']})
            for ln in o['_lines']:
                new_lines += 1
                lines.append({
                    'po': so, 'item_no': ln['item_no'], 'ean': ln['ean'],
                    'description': ln['description'], 'qty': ln['order_qty'],
                    'unit_price': ln['unit_price'], 'our_mrp': ln['mrp'],
                    'status': 'OK', 'exception_label': (
                        f"tester +{ln['tester_qty']}" if ln['tester_qty'] else ''),
                    'key': _line_key(so, ln['item_no'], ln['ean']),
                })
        warnings = []
        # per-PO validation guards (format · location · duplicate/collision) —
        # surfaced for EVERY order (new + already-recorded) so nothing slips through.
        for so, o in orders.items():
            for iss in (o.get('issues') or []):
                warnings.append(f"{so} — {iss}")
        for fname, w in (getattr(rec.result, 'warned_files', None) or []):
            warnings.append(f"{fname}: {w}")
        for fname, rsn in (getattr(rec.result, 'failed_files', None) or []):
            warnings.append(f"[FAILED] {fname}: {rsn}")
        fi = _file_issues(rec.result)
        return {
            'ok': bool(headers) if phase == 'preview' else True,
            'summary': {'pos': new_pos, 'lines': new_lines, 'qty': new_qty,
                        'value': round(new_value, 2), 'affected': len(fi),
                        'skipped': len(skipped), 'revised': n_revised,
                        'blocked': len(blocked)},
            'headers': headers, 'lines': lines, 'affected': [],
            'file_issues': fi, 'skipped': skipped, 'blocked': blocked,
            'warnings': warnings, 'output_path': output_path,
            'error': (None if headers or phase != 'preview'
                      else 'No resolvable POs in the uploaded file(s).'),
        }

    def workbook(self):
        """Generate the 7-sheet dump (SO Workbook) on demand — for the review
        'Download' link before confirm. Reuses the frozen exporter; no DB write."""
        rec = GTMassRecorder(self.paths, self.warehouse)
        rec._process()
        return rec._write_dump()

    # ── flow protocol ────────────────────────────────────────────────────
    def preview(self) -> dict:
        try:
            rec, orders = self._build()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e), 'summary': {}, 'headers': [],
                    'lines': [], 'affected': [], 'file_issues': [],
                    'skipped': [], 'warnings': []}
        return self._payload(rec, orders, phase='preview')

    def confirm(self, actions: dict | None = None) -> dict:
        try:
            rec, orders = self._build()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e)}
        # Apply per-line Excludes before recording.
        excluded = {k for k, v in (actions or {}).items()
                    if (v or {}).get('action') == 'EXCLUDE'}
        if excluded:
            for o in orders.values():
                kept = [ln for ln in o['_lines']
                        if _line_key(o['po'], ln['item_no'], ln['ean'])
                        not in excluded]
                o['_lines'] = kept
                o['items'] = len(kept)
                o['qty'] = sum(int(ln['order_qty'] or 0) for ln in kept)
            orders = {so: o for so, o in orders.items() if o['_lines']}
        recorded = rec._record(orders)
        out = rec._write_dump()
        return {
            'ok': bool(recorded.get('recorded')),
            'run_id': recorded.get('run_id'),
            'pos': recorded.get('recorded_pos', 0),
            'lines': recorded.get('lines', 0),
            'skipped': recorded.get('skipped', 0),
            'reason': recorded.get('reason', ''),
            'output_path': str(out) if out else None,
            'output_name': out.name if out else None,
            'error': (None if recorded.get('recorded')
                      else recorded.get('reason', 'Nothing recorded.')),
        }
