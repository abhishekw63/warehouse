"""
online_b2b.services.engine_bridge
=================================

Class-based bridge that runs the EXISTING ``online_po_processor`` engine as a
library (Option A — sibling on ``sys.path``, wired in settings). The engine is
the frozen backup and is NEVER modified — these classes only orchestrate it.

  * :class:`Processor`        — base: loads masters, runs the engine (single or
    multi-file), applies dedup, and builds the web preview/confirm payloads.
  * :class:`FlipkartProcessor`— Flipkart: always multi-file (one
    ``purchase_order_*.xlsx`` per PO), plus the optional ``purchase-orders-*.csv``
    header → Tracker sheet (``result.flipkart_tracker_rows``). FK Grocery /
    hyperlocal ship-to resolution is master-driven (engine handles it).

Module-level :func:`preview` / :func:`confirm` keep their old signatures (views
call them) but delegate to the right Processor subclass via :func:`processor_for`.
"""

from __future__ import annotations

import re
import time
from pathlib import Path

PILOT_MARKETPLACES = ['Blink', 'Flipkart', 'RK', 'Dmart', 'Zepto', 'Flipkart-TO',
                      'Purplle', 'Swiggy', 'Nykaa', 'Myntra', 'Reliance',
                      'Meesho-TO']
# Friendly labels for the upload dropdown where the engine key isn't operator-
# facing. The engine key (value) is unchanged — only the shown text differs.
PILOT_LABELS = {'Flipkart-TO': 'Flipkart Branch', 'Meesho-TO': 'Meesho Branch'}
_ISSUE_STATUSES = {'MISMATCH', 'NOT_IN_MASTER'}


def pilot_choices() -> list:
    """``[(engine_key, label)]`` for the upload form — value stays the engine
    marketplace key, label is the operator-friendly name where one is defined."""
    return [(m, PILOT_LABELS.get(m, m)) for m in PILOT_MARKETPLACES]


def _engine_imports():
    """Import the engine (frozen backup) lazily and return its handles in a
    dict. Explicit dict (not locals()) so linters never strip the imports."""
    from online_po_processor.config.marketplaces import (
        DEFAULT_WAREHOUSE,
        MARKETPLACE_CONFIGS,
        WAREHOUSE_CODES,
        WAREHOUSE_DISPLAY_NAMES,
    )
    from online_po_processor.config.paths import (
        get_bundled_mapping_path,
        get_bundled_master_path,
    )
    from online_po_processor.data.mapping_loader import MappingLoader
    from online_po_processor.data.master_loader import MasterLoader
    from online_po_processor.engine.marketplace_engine import MarketplaceEngine
    from online_po_processor.exporter.so_exporter import SOExporter
    return {
        'MARKETPLACE_CONFIGS': MARKETPLACE_CONFIGS,
        'WAREHOUSE_CODES': WAREHOUSE_CODES,
        'DEFAULT_WAREHOUSE': DEFAULT_WAREHOUSE,
        'WAREHOUSE_DISPLAY_NAMES': WAREHOUSE_DISPLAY_NAMES,
        'get_bundled_master_path': get_bundled_master_path,
        'get_bundled_mapping_path': get_bundled_mapping_path,
        'MappingLoader': MappingLoader,
        'MasterLoader': MasterLoader,
        'MarketplaceEngine': MarketplaceEngine,
        'SOExporter': SOExporter,
    }


def warehouse_choices() -> list[str]:
    try:
        return list(_engine_imports()['WAREHOUSE_DISPLAY_NAMES'])
    except Exception:
        return ['AHD']


def default_warehouse() -> str:
    try:
        return _engine_imports()['DEFAULT_WAREHOUSE']
    except Exception:
        return 'AHD'


def default_margin_pct(marketplace: str) -> int:
    try:
        cfg = _engine_imports()['MARKETPLACE_CONFIGS'][marketplace]
        return int(cfg.get('default_margin', 70))
    except Exception:
        return 70


def marketplace_rules() -> list:
    """Per-marketplace rule summary straight from the engine config — drives the
    Rules & Exceptions reference page (so it can never drift from the engine).
    Read-only; never raises."""
    out = []
    try:
        cfgs = _engine_imports()['MARKETPLACE_CONFIGS']
        for name, c in cfgs.items():
            basis = c.get('compare_basis', 'cost')
            parser = (c.get('file_parser') or c.get('pdf_parser')
                      or c.get('source_format') or '')
            gmd = c.get('gst_margin_discount')
            # GST-dependent margin (Reliance): keep% = 1 − discount × (1+GST).
            # Computed from the config so the table can never drift.
            gst_margin = None
            if gmd is not None:
                gst_margin = [{'gst': g, 'pct': round((1 - gmd * (1 + g / 100.0)) * 100, 2)}
                              for g in (18, 12, 5, 0)]
            # Honest margin label: a flat "70%" for normal channels, but a RANGE
            # ("63.42–69% · by GST") for GST-dependent ones (Reliance) so the
            # table never reads as a single flat rate.
            if gst_margin:
                _p = [g['pct'] for g in gst_margin]
                margin_label = f"{min(_p):g}–{max(_p):g}%"
            else:
                margin_label = f"{c.get('default_margin', 70)}%"
            out.append({
                'name': name,
                'party': c.get('party_name', ''),
                'margin': c.get('default_margin', 70),
                'margin_label': margin_label,
                'basis': ('Landing — MRP × margin% (pre-GST)' if basis == 'landing'
                          else 'Cost — MRP × margin% ÷ GST (post-GST)'),
                'compare_label': c.get('compare_label', ''),
                'item': c.get('item_resolution', '') or 'mapping',
                'gst_discount': gmd,
                'gst_margin': gst_margin,
                'parser': parser,
                'pilot': name in PILOT_MARKETPLACES,
            })
        out.sort(key=lambda x: (not x['pilot'], x['name']))
    except Exception:  # noqa: BLE001
        pass
    return out


# Curated, operator-facing note on each marketplace's FILE SHAPE (what the team
# actually uploads). The columns themselves come from the engine config below, so
# only the shape (multi-file / extra report / how PO+location are found) is here.
_FMT_NOTES = {
    'Blink': 'One consolidated punch Excel — one row per ordered line.',
    'Flipkart': 'Many purchase_order_*.xlsx (one per PO) + an optional header .csv '
                'that classifies each PO as FK Hyperlocal / FK Grocery.',
    'Flipkart-TO': 'Per-PO Consignment_Details_<PO>.csv files + an optional '
                   'Consignment Visibility Report .csv for the destinations.',
    'Meesho-TO': 'One order-line-items-<PO>[_<city>].csv per order; the destination '
                 'comes from a city token in the filename (MS_BLR → Bengaluru…).',
    'Dmart': 'One or more Avenue/DMart PO PDFs; PO & store read from the PDF.',
    'Reliance': 'One or more Reliance PO PDFs; GST rate is read per line.',
    'Firstcry': 'A FirstCry PO PDF (bordered line-item table).',
    'Swiggy': 'A Swiggy consignment .csv; items arrive as a Swiggy SKU code that '
              'is resolved to an EAN via the channel map.',
    'Myntra': 'A compiled dump .xlsx (a Myntra PO PDF is a fallback).',
    'Purplle': 'A tab-separated .xls / .csv export.',
    'Bigbasket': 'A custom-layout BigBasket Excel.',
    'Nykaa': 'A Nykaa PO Excel (.xlsx).',
    'RK': 'An RK PO Excel (.xlsx).',
    'Zepto': 'A Zepto PO Excel (.xlsx).',
}


def marketplace_formats() -> list:
    """Per-marketplace **file format** reference for the Rules page — the file
    type, how the PO / location / item are found, and the key columns read.
    Columns come straight from the engine config (so they never drift); the
    file-shape note is curated (``_FMT_NOTES``). Read-only; never raises."""
    out = []
    try:
        cfgs = _engine_imports()['MARKETPLACE_CONFIGS']

        def _col(v):
            if isinstance(v, list):
                return ' / '.join(str(x) for x in v)
            if isinstance(v, dict):
                m = v.get('multiply')
                return ('computed (' + ' × '.join(m) + ')') if m else 'computed'
            if v and str(v).startswith('__'):
                return '(from the file)'
            return v or '—'

        for name, c in cfgs.items():
            ir = c.get('item_resolution', '')
            if ir == 'from_ean':
                item_by = f"EAN → master · col “{c.get('ean_col', '')}”"
            elif ir == 'from_swiggy_sku':
                item_by = f"Swiggy SKU → EAN → master · col “{c.get('sku_col') or c.get('ean_col') or 'SkuCode'}”"
            elif ir == 'from_column':
                item_by = f"Item No · col “{c.get('item_col', '')}”"
            else:
                item_by = _col(c.get('ean_col') or c.get('item_col'))
            exts = c.get('accepted_extensions')
            if not exts:                      # some (Flipkart-TO/Meesho) nest it
                for v in c.values():
                    if isinstance(v, dict) and v.get('accepted_extensions'):
                        exts = v['accepted_extensions']
                        break
            if c.get('source_format') == 'pdf':
                ftype = 'PDF'
            elif exts:
                ftype = ' / '.join(exts)
            else:
                ftype = '.xlsx'
            out.append({
                'name': name,
                'file_type': ftype,
                'note': _FMT_NOTES.get(name, ''),
                'po_by': _col(c.get('po_col')),
                'loc_by': _col(c.get('loc_col')),
                'item_by': item_by,
                'qty_by': _col(c.get('qty_col')),
                'cost_by': _col(c.get('fob_col')),
                'mrp_by': _col(c.get('mrp_col')) if c.get('mrp_col') else '—',
                'pilot': name in PILOT_MARKETPLACES,
            })
        out.sort(key=lambda x: (not x['pilot'], x['name']))
    except Exception:  # noqa: BLE001
        pass
    return out


_TEMPLATES_PATH = Path(__file__).with_name('template_samples.json')
# config key → (role label shown in the UI, CSS role class for colour)
_ROLE_KEYS = [
    ('po_col', 'PO', 'po'), ('loc_col', 'Destination', 'loc'),
    ('ean_col', 'Item · EAN', 'item'), ('item_col', 'Item', 'item'),
    ('sku_col', 'Item · SKU', 'item'), ('qty_col', 'Quantity', 'qty'),
    ('fob_col', 'Vendor cost', 'cost'), ('mrp_col', 'MRP', 'mrp'),
    ('po_date_col', 'PO date', 'date'), ('exp_date_col', 'Expiry', 'exp'),
    ('hsn_col', 'HSN', 'hsn'), ('amount_col', 'Amount', 'amt'),
]


def _role_names(v) -> list:
    """Real column name(s) from a config value (str / list); skip synthetic
    ``__…__`` keys and computed dicts."""
    if isinstance(v, str) and not v.startswith('__'):
        return [v]
    if isinstance(v, list):
        return [x for x in v if isinstance(x, str) and not x.startswith('__')]
    return []


def _role_map(cfg: dict) -> dict:
    """``{lower-cased header → (role label, role class)}`` for every column the
    engine actually reads — including columns nested in sub-configs
    (Flipkart-TO / Meesho-TO)."""
    out: dict = {}
    dicts = [cfg] + [v for v in cfg.values() if isinstance(v, dict)]
    for d in dicts:
        for key, role, cls in _ROLE_KEYS:
            for nm in _role_names(d.get(key)):
                out.setdefault(nm.strip().lower(), (role, cls))
    return out


def marketplace_templates() -> dict[str, dict]:
    """Per-marketplace **full file template** for the Rules “See full template”
    page: every column of a real sample file plus a few sample rows, each column
    tagged with the role the engine reads it as (``role`` / ``role_class``) or
    left blank (unused → dulled in the UI). Columns and sample rows are a frozen
    fixture (``template_samples.json``, captured from real files); the used/role
    tagging is computed **live** from the engine config so highlighting never
    drifts from what the parser actually reads. Read-only; never raises."""
    import json
    out: dict = {}
    try:
        samples = json.loads(_TEMPLATES_PATH.read_text(encoding='utf-8'))
    except Exception:  # noqa: BLE001
        return out
    try:
        cfgs = _engine_imports()['MARKETPLACE_CONFIGS']
    except Exception:  # noqa: BLE001
        cfgs = {}
    for name, s in samples.items():
        roles = _role_map(cfgs.get(name, {}))
        cols = []
        for col in s.get('columns', []):
            role, cls = roles.get(str(col).strip().lower(), ('', ''))
            cols.append({'name': col, 'role': role, 'role_class': cls,
                         'used': bool(role)})
        # Pre-align sample rows to the column order as cell dicts, so the template
        # can render + style each cell without dynamic dict-key lookup.
        grid = [
            [{'value': r.get(c['name'], ''), 'used': c['used'],
              'role_class': c['role_class']} for c in cols]
            for r in s.get('rows', [])
        ]
        legend, seen = [], set()
        for c in cols:
            if c['used'] and c['role'] not in seen:
                seen.add(c['role'])
                legend.append({'role': c['role'], 'role_class': c['role_class']})
        out[name] = {
            'name': name,
            'file_type': s.get('file_type', ''),
            'sample_file': s.get('sample_file', ''),
            'columns': cols,
            'grid': grid,
            'legend': legend,
            'used': sum(1 for c in cols if c['used']),
            'total': len(cols),
            'pilot': name in PILOT_MARKETPLACES,
        }
    return out


def marketplace_template(name: str):
    """Single marketplace template by engine key, or ``None`` if unknown."""
    return marketplace_templates().get(name)


def location_rules() -> list:
    """Flipkart Origin-Warehouse → sub-marketplace (FK Hyperlocal / FK Grocery)
    locked map — for the Rules page. Read-only."""
    try:
        from online_po_processor.engine.flipkart_tracker import LOCATION_MARKETPLACE
        return sorted(({'loc': k, 'mkt': v} for k, v in LOCATION_MARKETPLACE.items()),
                      key=lambda r: (r['mkt'], r['loc']))
    except Exception:  # noqa: BLE001
        return []


def margin_defaults() -> dict:
    """``{marketplace: default_margin%}`` for every pilot marketplace — so the
    upload form can auto-fill the right margin when the marketplace changes
    (e.g. Flipkart → 77, Blink → 70). Mirrors the desktop GUI's per-marketplace
    pre-fill."""
    return {m: default_margin_pct(m) for m in PILOT_MARKETPLACES}


# ── Processor (base) ────────────────────────────────────────────────────

class Processor:
    """Run the engine for one marketplace and build web payloads. Subclasses
    override the small hooks (``engine_files`` / ``use_multi`` / ``post_process``)."""

    def __init__(self, marketplace, po_paths, warehouse=None, margin_pct=None,
                 ean_fixes=None):
        self.marketplace = marketplace
        self.po_paths = [str(p) for p in (po_paths or [])]
        self.warehouse = warehouse
        self.margin_pct = margin_pct
        # operator's pending EAN corrections for THIS upload ({wrong → correct})
        self.ean_fixes = dict(ean_fixes or {})
        self.warnings: list[str] = []
        self.skipped: list = []
        self.env = None
        self.config = None
        self.result = None

    # ── overridable hooks ───────────────────────────────────────────────
    def engine_files(self) -> list[str]:
        """Files actually fed to the engine."""
        return self.po_paths

    def use_multi(self, config) -> bool:
        """Whether to call ``process_multi`` (mirrors the engine's
        ``_supports_multi_file``: pdf / pdf_parser / file_parser)."""
        return bool(config.get('source_format') == 'pdf'
                    or config.get('pdf_parser') or config.get('file_parser'))

    def post_process(self, result, env) -> None:
        """Hook after a successful engine run (e.g. Flipkart tracker)."""

    def run_engine(self, engine, files, config):
        """Call the engine for this marketplace. Default = standard single-file
        ``process`` (or ``process_multi`` for pdf/multi-file). Subclasses with a
        bespoke ingest (e.g. Flipkart-TO consignments) override this."""
        if self.use_multi(config):
            return engine.process_multi(files, config, margin_pct=self.margin_pct)
        if len(files) > 1:
            self.warnings.append(
                f"{self.marketplace} expects a single file; processing "
                f"'{Path(files[0]).name}', ignored {len(files) - 1} extra.")
        return engine.process(files[0], config, margin_pct=self.margin_pct)

    # ── core run ────────────────────────────────────────────────────────
    def _run(self, skip_dedup=False):
        """Process → ``self.result`` (dedup applied unless ``skip_dedup``).
        Returns an error dict on failure, or ``None`` on success. ``skip_dedup``
        is used when re-running purely to EXPORT (e.g. the D365 package) — the
        ERP file must carry the full upload, not a DB-deduped subset."""
        env = self.env = _engine_imports()
        configs = env['MARKETPLACE_CONFIGS']
        if self.marketplace not in configs:
            return {'ok': False, 'error': f"Unknown marketplace '{self.marketplace}'."}
        files = self.engine_files()
        if not files:
            return {'ok': False, 'error': "No PO file uploaded."}

        config = self.config = configs[self.marketplace]
        self.warehouse = self.warehouse or env['DEFAULT_WAREHOUSE']
        if self.margin_pct is None:
            self.margin_pct = config.get('default_margin', 70) / 100.0

        # ── Single source of truth: the DB. The web no longer reads ANY bundled
        # Excel — item master, Ship-To mapping, AND the pricing-override overlays
        # (exceptions + Swiggy deal SKUs) all come from MySQL. (The desktop app
        # keeps its Excel; only the web is retired off it.) Empty table → a clear
        # "seed it first" error, never a silent Excel fallback.
        party = config['party_name']
        from . import item_master_loader as iml
        from . import mapping_store as mstore

        if mstore.table_count() == 0:
            return {'ok': False, 'error': "Ship-To mapping DB is empty — seed it "
                    "on the Ship-To Mapping page first."}
        mapping = mstore.DBMappingLoader()
        loc_count = mapping.load(None, party, [])
        if loc_count == 0:
            return {'ok': False, 'error': f"No mapping locations for "
                    f"'{self.marketplace}' (party '{party}'). Add it on the "
                    f"Ship-To Mapping page."}

        if iml.table_count() == 0:
            return {'ok': False, 'error': "Item master DB is empty — upload it on "
                    "the Item Master page first."}
        try:
            master = iml.DBMasterLoader().load_from_db()
        except Exception as e:  # noqa: BLE001
            return {'ok': False,
                    'error': f"Item master load failed: {type(e).__name__}: {e}"}

        # Apply this upload's pending EAN corrections so the wrong EANs resolve
        # on re-validation (DB-backed master exposes add_session_aliases;
        # historical fixes are already merged at load).
        if self.ean_fixes and hasattr(master, 'add_session_aliases'):
            master.add_session_aliases(self.ean_fixes)

        engine = env['MarketplaceEngine'](mapping, master=master)
        result = self.run_engine(engine, files, config)

        result.margin_pct = self.margin_pct
        result.warehouse_display = self.warehouse
        result.warehouse_code = env['WAREHOUSE_CODES'].get(self.warehouse, 'PICK')
        for _po, _loc, msg in result.warnings:
            self.warnings.append(msg)

        if not result.rows:
            return {'ok': False, 'error': "No valid rows extracted from the PO file(s).",
                    'warnings': self.warnings}

        if not skip_dedup:
            try:
                from . import lines_store
                self.skipped = lines_store.web_dedup(result, self.marketplace) or []
            except Exception as e:  # noqa: BLE001
                self.warnings.append(f"Dedup check skipped ({type(e).__name__}: {e}).")

        self.post_process(result, env)
        self.result = result
        return None

    # ── D365 package from the operator's locked decisions ───────────────
    def _apply_decisions(self, actions):
        """Copy of ``self.result`` with operator decisions applied: EXCLUDE rows
        dropped, OVERRIDE rows repriced via the engine's own ``forced_unit_price``
        (which the D365 package reads as Unit Price). Originals are NOT mutated;
        the engine is untouched. ``actions`` is keyed ``po|item_no|ean``."""
        import copy
        actions = actions or {}
        rows = []
        for so in self.result.rows:
            key = f"{so.po_number}|{so.item_no or ''}|{so.ean or ''}"
            dec = actions.get(key) or {}
            act = str(dec.get('action') or '').upper()
            if act == 'EXCLUDE':
                continue
            if act == 'OVERRIDE':
                try:
                    ocp = float(dec.get('override_cp'))
                except (TypeError, ValueError):
                    ocp = None
                if ocp is not None:
                    so = copy.copy(so)
                    so.forced_unit_price = round(ocp, 2)
            rows.append(so)
        r2 = copy.copy(self.result)
        r2.rows = rows
        return r2

    def generate_d365(self, out_path, actions=None) -> dict:
        """Build the ERP-uploadable D365 package at ``out_path`` from LOCKED
        decisions. Reuses the engine's ``export_d365_package`` — engine + the full
        SO Workbook stay untouched. No DB write, no dedup."""
        err = self._run(skip_dedup=True)
        if err:
            return err
        try:
            from online_po_processor.exporter.d365_package import (
                export_d365_package,
            )
            from online_po_processor.exporter.so_exporter import SOExporter
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': f"D365 exporter unavailable: {e}"}
        filtered = self._apply_decisions(actions)
        if not filtered.rows:
            return {'ok': False, 'error': "No lines left after Exclude decisions."}
        is_to = getattr(filtered, 'output_type', 'so') == 'to'
        template = (SOExporter._D365_TO_TEMPLATE if is_to
                    else SOExporter._D365_SO_TEMPLATE)
        if not Path(template).exists():
            return {'ok': False, 'error': "D365 template not bundled on this host."}
        try:
            export_d365_package(filtered, Path(template), Path(out_path))
        except Exception as e:  # noqa: BLE001
            return {'ok': False, 'error': f"D365 export failed: "
                    f"{type(e).__name__}: {e}"}
        excluded = len(self.result.rows) - len(filtered.rows)
        return {'ok': True, 'd365_path': str(out_path),
                'lines': len(filtered.rows), 'excluded': excluded}

    # ── payload builders ────────────────────────────────────────────────
    def _amountless_to(self) -> bool:
        """True for a Transfer-Order marketplace whose dump carries NO vendor
        amount (``amount_col`` omitted) — e.g. Flipkart Branch (Flipkart-TO).
        Its file value is zero, so we derive the inc-GST transfer value from OUR
        master pricing instead."""
        cfg = self.config or {}
        return (getattr(self.result, 'output_type', 'so') == 'to'
                and not cfg.get('amount_col'))

    def _to_value_by_po(self, lines) -> dict:
        """``{po: Σ (our Landing × qty)}`` — the inc-GST transfer value computed
        from our pricing (Landing = MRP × margin% = Cost × (1+GST), so Landing ×
        qty is the GST-inclusive amount). Used to fill order_value for
        amount-less TOs where the dump total is zero."""
        vals: dict = {}
        for ln in lines:
            land = ln.get('our_landing')
            if land is None:
                continue
            po = str(ln.get('po', ''))
            vals[po] = vals.get(po, 0.0) + float(land) * int(ln.get('qty') or 0)
        return {p: round(v, 2) for p, v in vals.items()}

    def _headers(self):
        from online_po_processor.auto.history_db import order_rows_from_result
        headers = order_rows_from_result(
            self.result, self.marketplace, self.warehouse, '')
        # Amount-less TO (Flipkart Branch): the dump carries no price, so the
        # transfer value is OUR calculated pricing (Landing × qty = calculated
        # CP inc-GST), not the vendor's. Fill it so preview, the summary, and the
        # locked DB show the real value — and log it (never silent: the operator
        # must know the amount is computed, not received).
        if self._amountless_to():
            vals = self._to_value_by_po(self._lines())
            filled = 0
            for h in headers:
                v = vals.get(str(h.get('po', '')))
                if v is not None:
                    h['order_value'] = v
                    filled += 1
            if filled:
                total = round(sum(vals.values()), 2)
                label = PILOT_LABELS.get(self.marketplace, self.marketplace)
                self.warnings.append(
                    f"{label}: dump carries no price — order value is COMPUTED "
                    f"from our master pricing (calculated CP inc-GST = Landing × "
                    f"qty) for {filled} PO(s), total ₹{total:,.2f}.")
        return headers

    _EAN_MAX = 20   # matches the order_lines.ean VARCHAR(20) column

    def _lines(self, run_id=None, output_file='', actions=None):
        from . import lines_store
        # Combine historical + this-session EAN fixes so build_lines can swap
        # the wrong EAN → correct on the line and stamp received_ean (audit).
        combined = dict(lines_store.ean_alias_map())
        combined.update(self.ean_fixes)
        rows = lines_store.build_lines(self.result, run_id=run_id,
                                       output_file=output_file, actions=actions,
                                       ean_fixes=combined)
        # Defensive: a PDF parser (e.g. Myntra) can occasionally emit a
        # malformed over-long EAN — two EANs concatenated on a page wrap. The DB
        # ean column is VARCHAR(20), so one bad row would crash the WHOLE lock.
        # Cap it (DB-safe) + warn (never silent) — such a line is NOT_IN_MASTER
        # anyway, so the operator corrects it on the Affected/Issues tab.
        for r in rows:
            e = r.get('ean')
            if e and len(str(e)) > self._EAN_MAX:
                msg = (
                    f"Malformed EAN on PO {r.get('po')} item {r.get('item_no')}: "
                    f"'{e}' ({len(str(e))} chars > {self._EAN_MAX}) — looks like two "
                    f"EANs merged on a page wrap. Stored truncated; line is "
                    f"NOT_IN_MASTER — set the correct EAN on the Affected tab.")
                if msg not in self.warnings:   # _lines runs >once per flow
                    self.warnings.append(msg)
                r['ean'] = str(e)[:self._EAN_MAX]
        return rows

    def _summary(self, lines, headers):
        mism = sum(1 for l in lines if l['status'] == 'MISMATCH')
        nim = sum(1 for l in lines if l['status'] == 'NOT_IN_MASTER')
        return {
            'marketplace': self.marketplace, 'warehouse': self.warehouse,
            'margin_pct': round(self.margin_pct * 100, 2),
            'pos': len({l['po'] for l in lines}),
            'lines': len(lines),
            'qty': sum(int(l['qty'] or 0) for l in lines),
            'value': sum(float(h.get('order_value') or 0) for h in headers),
            'mismatch': mism, 'not_in_master': nim, 'affected': mism + nim,
            'skipped': len(self.skipped),
        }

    def _export(self):
        # The engine writes the workbook to ``input_file_path.parent/output``.
        # Web uploads already live UNDER MEDIA (b2b_uploads/<token>/), so the
        # engine writes output/ there — web-owned, per-token, and where
        # review_download reads it. ONLY when the input is OUTSIDE media (a
        # direct/script run against the operator's source folder) do we redirect
        # to a web-owned exports dir, so the workbook never lands next to source.
        # ``input_file_path`` is used solely to locate that folder (never embedded
        # in the workbook), so the swap is safe; we restore it afterwards.
        orig = getattr(self.result, 'input_file_path', '') or ''
        redirected = False
        try:
            from django.conf import settings
            media = Path(settings.MEDIA_ROOT).resolve()
            op = Path(orig).resolve() if orig else None
            under_media = bool(op and (op == media or media in op.parents))
            if orig and not under_media:
                exports = media / 'b2b_exports'
                exports.mkdir(parents=True, exist_ok=True)
                self.result.input_file_path = str(exports / Path(orig).name)
                redirected = True
        except Exception:  # noqa: BLE001 — fall back to engine default on any issue
            pass
        path = None
        try:
            path = self.env['SOExporter']().export(self.result, start_time=time.time())
        except Exception as e:  # noqa: BLE001
            self.warnings.append(f"Workbook export issue ({type(e).__name__}: {e}).")
        finally:
            if redirected:
                self.result.input_file_path = orig
        # Append a per-run 'SKU Summary' sheet (web post-process; the engine's
        # workbook + all its sheets are untouched). Best-effort — never fails the
        # export.
        if path:
            try:
                self._append_sku_sheet(str(path))
            except Exception as e:  # noqa: BLE001
                self.warnings.append(f"SKU Summary sheet skipped ({type(e).__name__}).")
        return path

    def _sku_pivot(self, lines):
        """Per-run SKU rollup grouped by (item_no, ean): qty per status, MRP
        comparison (+ varies flag), POs, worst diff. Returns sorted rows."""
        agg: dict = {}
        for ln in lines:
            key = (ln.get('item_no') or '', ln.get('ean') or '')
            a = agg.get(key)
            if a is None:
                a = agg[key] = {'desc': ln.get('description') or '',
                                'our_mrp': ln.get('our_mrp'), 'vmrps': set(),
                                'tot': 0, 'ok': 0, 'mis': 0, 'nim': 0,
                                'pos': set(), 'diffs': []}
            q = int(ln.get('qty') or 0)
            a['tot'] += q
            st = ln.get('status') or 'OK'
            a['ok' if st == 'OK' else 'mis' if st == 'MISMATCH'
              else 'nim' if st == 'NOT_IN_MASTER' else 'ok'] += q
            a['pos'].add(ln.get('po'))
            if ln.get('vendor_mrp') is not None:
                a['vmrps'].add(round(float(ln['vendor_mrp']), 2))
            if ln.get('diff') is not None:
                a['diffs'].append(float(ln['diff']))
        rows = []
        for (item_no, ean), a in agg.items():
            rows.append([
                item_no, ean, a['desc'], a['our_mrp'],
                (max(a['vmrps']) if a['vmrps'] else None),
                'YES' if len(a['vmrps']) > 1 else '', a['tot'], a['ok'],
                a['mis'], a['nim'], len(a['pos']),
                (min(a['diffs']) if a['diffs'] else None)])
        rows.sort(key=lambda r: (-r[8], -r[9], -r[6]))   # mismatch, nim, tot qty
        return rows

    def _append_sku_sheet(self, path):
        from openpyxl import load_workbook
        from openpyxl.styles import Font
        rows = self._sku_pivot(self._lines())
        wb = load_workbook(path)
        if 'SKU Summary' in wb.sheetnames:
            del wb['SKU Summary']
        ws = wb.create_sheet('SKU Summary')
        hdr = ['Item No', 'EAN', 'Description', 'Our MRP', 'Their MRP',
               'MRP varies', 'Tot Qty', 'OK Qty', 'Mismatch Qty',
               'Not-in-Master Qty', '# POs', 'Worst Diff']
        ws.append(hdr)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for r in rows:
            ws.append(r)
        wb.save(path)

    # ── phase 1: preview (no DB write) ──────────────────────────────────
    def preview(self) -> dict:
        err = self._run()
        if err:
            return err
        output_path = self._export()
        lines = self._lines()
        headers = self._headers()
        affected = [l for l in lines if l['status'] in _ISSUE_STATUSES]
        return {
            'ok': True, 'summary': self._summary(lines, headers),
            'headers': headers, 'lines': lines, 'affected': affected,
            'skipped': self.skipped, 'warnings': self.warnings,
            'output_path': str(output_path) if output_path else None,
        }

    # ── phase 2: confirm (push headers + lines) ─────────────────────────
    def confirm(self, actions=None) -> dict:
        err = self._run()
        if err:
            return err
        output_path = self._export()
        if output_path is None:
            return {'ok': False, 'error': "Workbook export failed.",
                    'warnings': self.warnings}
        lines = self._lines()
        out: dict = {
            'ok': True, 'output_path': str(output_path), 'run_id': None,
            'warnings': self.warnings,
            'summary': self._summary(lines, self._headers()),
        }
        try:
            from . import lines_store
            # Web-owned write of runs + order_headers (replaces the engine's
            # record_manual → no engine history store, so order_issue_lines is
            # never recreated). Byte-identical to the old path (parity-verified).
            rec = lines_store.record_run_headers(
                self.result, self.marketplace, self.warehouse, str(output_path))
            out['run_id'] = rec.get('run_id')
            out['new_orders'] = rec.get('new_orders', 0)
            # via _lines() so EAN fixes are applied (correct ean on the line +
            # received_ean stamped on the validation row).
            rows = self._lines(run_id=out['run_id'],
                               output_file=str(output_path), actions=actions)
            out['lines_recorded'] = lines_store.insert_lines(
                out['run_id'], rows).get('recorded', 0)
            # Amount-less TO (Flipkart Branch): the engine recorded order_value=0
            # (dump has no amount). Lock the inc-GST value we derived from our
            # table onto the headers we just wrote.
            if self._amountless_to():
                upd = lines_store.set_order_value(
                    out['run_id'], self._to_value_by_po(rows))
                out['value_backfilled'] = upd.get('updated', 0)
            # Tracker dates: PDF marketplaces (e.g. DMart) carry PO Date / validity
            # in the PDF header but not as a row column, so the engine leaves
            # po_date/exp_date blank → they'd never show on the TAT page. Backfill
            # the real dates from the source (fills blanks only). See
            # ``_source_dates_by_po`` (no-op for marketplaces the engine already
            # dates via po_date_col).
            dts = self._source_dates_by_po()
            if dts and out.get('run_id'):
                out['dates_backfilled'] = lines_store.set_po_dates(
                    out['run_id'], dts).get('updated', 0)
        except Exception as e:  # noqa: BLE001
            out['ok'] = False
            out['error'] = f"DB push failed: {type(e).__name__}: {e}"
        return out

    def _source_dates_by_po(self) -> dict:
        """``{po: {'po_date': date, 'exp_date': date}}`` from the source file(s),
        for marketplaces whose parser carries the dates in the header (not a row
        column). Base = none; PDF processors override. Used to backfill the
        tracker (po_date/exp_date) so TAT works for them."""
        return {}


# ── Flipkart ────────────────────────────────────────────────────────────

class FlipkartProcessor(Processor):
    """Flipkart: many ``purchase_order_*.xlsx`` (multi-file, PO from filename,
    address→ship-to) + an optional ``purchase-orders-*.csv`` header that drives
    the Tracker sheet. FK Grocery / hyperlocal resolution is master-driven."""

    def engine_files(self) -> list[str]:
        # The engine gets the PO xlsx; a uploaded .csv is the tracker header.
        return [p for p in self.po_paths if p.lower().endswith('.xlsx')]

    def use_multi(self, config) -> bool:
        return True

    # ── Tracker-driven sub-marketplace labels (FK Hyperlocal / FK Grocery) ──
    def _tracker_labels(self) -> dict:
        """``{PO: 'FK Hyperlocal'|'FK Grocery'|…}`` from the tracker rows. Empty
        when no header CSV was uploaded (then labels stay the engine default)."""
        tracker = getattr(self.result, 'flipkart_tracker_rows', None) or []
        return {str(t.get('PO', '')).strip(): t.get('Market Place', '')
                for t in tracker if t.get('Market Place')}

    def _headers(self):
        """Override the blanket 'Flipkart Alpha' label with the per-PO tracker
        classification so the dashboard/preview show FK Hyperlocal / FK Grocery."""
        rows = super()._headers()
        by_po = self._tracker_labels()
        for h in rows:
            label = by_po.get(str(h.get('po', '')).strip())
            if label:
                h['marketplace_label'] = label
        return rows

    def confirm(self, actions=None) -> dict:
        out = super().confirm(actions)
        # The engine's record_manual wrote 'Flipkart Alpha'; re-stamp the
        # web-owned display column per PO from the (latest) tracker mapping.
        if out.get('ok') and out.get('run_id'):
            by_po = self._tracker_labels()
            if by_po:
                try:
                    from .order_db import _conn
                    with _conn() as (cur, d):
                        ph = d['ph']
                        for po, label in by_po.items():
                            cur.execute(
                                f"UPDATE order_headers SET marketplace_label={ph} "
                                f"WHERE run_id={ph} AND po={ph} AND marketplace={ph}",
                                (label, out['run_id'], po, 'Flipkart'))
                        cur.connection.commit()
                except Exception as e:  # noqa: BLE001
                    out.setdefault('warnings', []).append(
                        f"Tracker re-label skipped ({type(e).__name__}: {e}).")
        return out

    def post_process(self, result, env) -> None:
        csv = next((p for p in self.po_paths if p.lower().endswith('.csv')), None)
        if not csv:
            self.warnings.append(
                "Flipkart Tracker: no 'purchase-orders-*.csv' header uploaded — "
                "Tracker sheet skipped (SO/values/locations are unaffected).")
            return
        try:
            from online_po_processor.engine.flipkart_tracker import (
                build_flipkart_tracker,
            )
            result.flipkart_tracker_rows = build_flipkart_tracker(csv)
        except Exception as e:  # noqa: BLE001
            self.warnings.append(f"Flipkart Tracker skipped ({type(e).__name__}: {e}).")


# ── Flipkart Branch (Flipkart-TO) ───────────────────────────────────────

class FlipkartTOProcessor(Processor):
    """Flipkart Branch (Transfer Orders). The operator hands us Flipkart's raw
    per-PO exports ``Consignment_Details_<PO>_<date>.csv`` (one per PO) plus an
    optional ``Consignment_Visibility_Report*.csv`` for destination Locations.
    The engine's ``process_consignments`` parses the PO from each filename,
    joins the visibility report, and runs the standard TO pipeline.

    The dump carries no order amount, so the inc-GST transfer value is derived
    from our master pricing (Landing × qty) in ``_headers`` / ``confirm`` — see
    ``Processor._amountless_to``. Falls back to the single consolidated dump
    when one .xlsx is uploaded instead."""

    _CONSIGNMENT_RE = re.compile(r'Consignment_Details_', re.I)
    _VISIBILITY_RE = re.compile(r'Consignment_Visibility_Report', re.I)

    def use_multi(self, config) -> bool:
        return True

    def _consignment_files(self) -> list:
        return [p for p in self.po_paths
                if self._CONSIGNMENT_RE.search(Path(p).name)]

    def _visibility_file(self):
        return next((p for p in self.po_paths
                     if self._VISIBILITY_RE.search(Path(p).name)), None)

    def run_engine(self, engine, files, config):
        cons = self._consignment_files()
        if cons:
            vis = self._visibility_file()
            if not vis:
                self.warnings.append(
                    "Flipkart Branch: no Consignment Visibility Report uploaded "
                    "— destination Locations may be blank (Transfer-to Code "
                    "unresolved). Add it to fill them.")
            return engine.process_consignments(
                cons, config, margin_pct=self.margin_pct,
                visibility_report_path=vis)
        # Fallback: a single pre-consolidated 7-column dump (.xlsx).
        return engine.process(files[0], config, margin_pct=self.margin_pct)


# ── Meesho Branch (Meesho-TO) ───────────────────────────────────────────

class MeeshoTOProcessor(Processor):
    """Meesho Branch (Transfer Orders). Bulk-consignment-ONLY: Meesho exports one
    CSV per order, ``order-line-items-<PO>[_<city>].csv`` (no consolidated dump,
    no visibility report). PO comes from the filename; Location from a city token
    in the filename (config ``filename_loc_from_shipto``: ``MS_BLR`` → 'blr', …),
    resolved to the Transfer-to Code via Ship-To B2B. The CSV carries no order
    amount (``sellingPricePerUnit`` is a selling price, deliberately ignored), so
    the inc-GST transfer value is derived from OUR master pricing (Landing × qty)
    in ``_headers`` / ``confirm`` — see ``Processor._amountless_to``."""

    _ITEMS_RE = re.compile(r'order-line-items', re.I)

    def use_multi(self, config) -> bool:
        return True

    def run_engine(self, engine, files, config):
        csvs = [p for p in self.po_paths if p.lower().endswith('.csv')]
        if not csvs:                      # tolerate any uploaded csv naming
            csvs = [p for p in files if str(p).lower().endswith('.csv')] or files
        # Location is filename-token driven for Meesho → no visibility report.
        return engine.process_consignments(
            csvs, config, margin_pct=self.margin_pct,
            visibility_report_path=None)


# ── DMart (Avenue) — PDF dates for the tracker ──────────────────────────

def _parse_ddmmyyyy(s):
    """'17.06.2026' / '17-06-2026' / '17/06/2026' → date, else None."""
    import datetime as _d
    s = (s or '').strip()
    for fmt in ('%d.%m.%Y', '%d-%m-%Y', '%d/%m/%Y'):
        try:
            return _d.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


class DmartProcessor(Processor):
    """DMart (Avenue PO PDFs). Processing is the base flow — this subclass only
    adds the tracker dates: the avenue parser reads ``Purchase Order Date`` and
    ``PO Validity`` from each PDF header (not a row column), so we backfill
    po_date/exp_date per PO after recording (so DMart shows on the TAT page)."""

    def _source_dates_by_po(self) -> dict:
        try:
            from online_po_processor.engine.avenue_pdf_parser import (
                parse_avenue_pdf,
            )
        except Exception:  # noqa: BLE001
            return {}
        out: dict = {}
        for p in self.po_paths:
            if not str(p).lower().endswith('.pdf'):
                continue
            try:
                po = parse_avenue_pdf(p)
                h = po.header
                if h.po_number:
                    out[str(h.po_number)] = {
                        'po_date': _parse_ddmmyyyy(h.po_date),
                        'exp_date': _parse_ddmmyyyy(h.validity_to),
                    }
            except Exception:  # noqa: BLE001 — never block on date extraction
                continue
        return out


# ── Factory + module entry points (views call these) ────────────────────

_PROCESSORS = {'Flipkart': FlipkartProcessor, 'Flipkart-TO': FlipkartTOProcessor,
               'Meesho-TO': MeeshoTOProcessor, 'Dmart': DmartProcessor}


def processor_for(marketplace, po_paths, warehouse=None, margin_pct=None,
                  ean_fixes=None) -> Processor:
    cls = _PROCESSORS.get(marketplace, Processor)
    return cls(marketplace, po_paths, warehouse, margin_pct, ean_fixes)


def preview(marketplace: str, po_paths, warehouse=None, margin_pct=None,
            ean_fixes=None) -> dict:
    return processor_for(marketplace, po_paths, warehouse, margin_pct,
                         ean_fixes).preview()


def confirm(marketplace: str, po_paths, warehouse=None, margin_pct=None,
            actions=None, ean_fixes=None) -> dict:
    return processor_for(marketplace, po_paths, warehouse, margin_pct,
                         ean_fixes).confirm(actions)


def generate_d365(marketplace: str, po_paths, out_path, warehouse=None,
                  margin_pct=None, actions=None, ean_fixes=None) -> dict:
    """Build the ERP D365 package reflecting the operator's locked Include/
    Override/Exclude decisions. Engine + full SO Workbook untouched."""
    return processor_for(marketplace, po_paths, warehouse, margin_pct,
                         ean_fixes).generate_d365(out_path, actions)
