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
        _engine_mod = mod
    return _engine_mod


# MT (Modern Trade) child channels exposed on the web. MT is the parent
# marketplace; these are its children (the operator picks one). Off Institutional
# (INST) is a SEPARATE parent and is not listed here. SS is verified end-to-end;
# the others share the same generic pipeline (test each before production use).
WEB_CHANNELS = ['SS', 'HG', 'NT', 'BN', 'LL']

# Accepted upload extensions (SS ships .xlsx; other MT channels may use .csv).
ACCEPTED_EXTENSIONS = ('.xlsx', '.csv')


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

    def __init__(self, channel_code: str, po_paths, warehouse: str | None = None):
        self.channel_code = channel_code
        self.po_paths = [Path(p) for p in (po_paths or [])]
        self.warehouse = warehouse or default_warehouse()
        self.report = ''
        self.output_path = None

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
        buf = io.StringIO()
        with redirect_stdout(buf):
            batch = eng.read_channel_csv_batch(
                self.po_paths, channel, bundle, store_override='')
        self.report = buf.getvalue()
        return eng, channel, batch

    # ── phase 1: preview (parse + validate, NO writes) ──────────────────
    def preview(self) -> dict:
        """Parse + resolve + validate only. No SO numbers assigned, no workbook,
        no DB — mirrors the online ``preview`` so the operator can verify first."""
        try:
            eng, channel, batch = self._load()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e)}
        return self._summary(batch, channel, recorded=None, phase='preview')

    # ── phase 2: confirm (assign + write + record to renee_orders) ──────
    def confirm(self) -> dict:
        """Assign SO numbers (burns the ``mt_select_seq.json`` counter ONCE),
        write the 6-sheet workbook, and record order headers into the shared
        ``renee_orders`` DB (segment Offline) via the desktop's own
        ``record_offline_batch`` — so SS appears on the online dashboard."""
        try:
            eng, channel, batch = self._load()
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': str(e)}
        warehouse_code = eng.WAREHOUSES.get(self.warehouse, 'PICK')
        buf = io.StringIO()
        # Channels with a tester-qty rule (e.g. Off Institutional) always pair a
        # tester SO — mirror the desktop's `gen_testers` derivation.
        gen_testers = getattr(channel, 'tester_qty_divisor', None) is not None
        try:
            with redirect_stdout(buf):
                eng.assign_so_numbers(batch, channel,
                                      generate_testers=gen_testers, tester_dump=None)
                eng.print_batch_report(batch)
                if any(pf.so_number for pf in batch.po_files):
                    self.output_path = eng.write_so_workbook(
                        batch, channel, warehouse_code,
                        output_path=None, add_non_stock=False)
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': f"{type(e).__name__}: {e}",
                    'report': self.report + buf.getvalue()}
        self.report += buf.getvalue()

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
                if ln.status == 'SKIP' or not ln.item_no:
                    continue                       # unresolved → surfaced as a warning
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
                    'item_no': str(ln.item_no or ''),
                    'ean': str(getattr(ln, 'ean', '') or ''),
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
                    'status': 'OK',
                    'exception_label': '',
                    'output_file': output_file or '',
                    'action': '', 'remark': '',
                })
        if rows:
            lines_store.insert_lines(run_id, rows)
        return len(rows)

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
