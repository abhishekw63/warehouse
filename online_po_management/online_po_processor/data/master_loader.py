"""
data.master_loader
==================

Loads ``Items_March.xlsx`` (the canonical product master) into an
in-memory lookup table indexed by both GTIN/EAN and Item No.

The master is the source of truth for:

* ``MRP`` — used to compute Landing Cost (``MRP × margin%``) and post-GST
  Cost Price (``... ÷ GST divisor``).
* ``GST Group Code`` — drives which divisor we apply.
* ``Description`` — surfaced in the Validation sheet next to the EAN so
  the user can read what each item actually is at a glance.
* ``No.`` (Item No) — the canonical ERP code resolved when the
  marketplace only provides an EAN.
* ``HSN/SAC Code`` — Optional (v1.6.0). Used for HSN cross-checking
  when the marketplace punch file also carries an HSN column and the
  marketplace's config has ``hsn_col`` set. Missing column is
  tolerated — those rows get reported as NOT_IN_MASTER for the HSN
  check.

The static helpers ``calc_cost_price`` and ``calc_landing_price`` are
exposed as classmethods so the engine can call them without holding a
loader instance.
"""

from __future__ import annotations
from typing import Any, Dict, Optional

import pandas as pd


class MasterLoader:
    """
    In-memory lookup over the Items Master file.

    Indexes each row by BOTH the stringified GTIN and the stringified Item
    No (``No.``). This means callers can look up by either an EAN
    (Myntra-style) or an Item No (RK-style) using the same ``lookup()``
    method.
    """

    def __init__(self) -> None:
        # Map from key (EAN or Item No, both as str) → entry dict
        # Entry shape: {item_no, mrp, gst_code, description, hsn, gtin}
        self.master: Dict[str, Dict] = {}

        # v2.4.0 (Swiggy): secondary index keyed by the marketplace's own
        # SKU code. Swiggy's PO dump carries only a ``SkuCode`` (e.g.
        # '138856') — no EAN, no Item No. The Items Master gets an optional
        # 'Swiggy SKU code' column mapping each SKU to its master row, so
        # ``item_resolution='from_swiggy_sku'`` looks the dump's SkuCode up
        # here to recover Item No / MRP / GST / EAN. Empty when the master
        # has no such column yet (the feature is then simply inactive).
        self.swiggy_sku: Dict[str, Dict] = {}

        # v2.4.0 (Swiggy): per-EAN deal-price overrides. Swiggy negotiates
        # special "deal" prices on selected SKUs that don't follow the
        # default MRP×80% landing math. Loaded from an optional
        # 'Swiggy Deal SKUs' sheet in the master workbook, keyed by EAN →
        # {cp, cost_with_gst, mrp, gst}. When a resolved line's EAN is in
        # here, the Validation sheet expects the sheet's 'Cost after GST'
        # as our cost price instead of the computed MRP×margin value.
        self.swiggy_deals: Dict[str, Dict] = {}

        # v2.4.0: central EAN/code exception map — {source_code → master_key}.
        # Marketplaces sometimes send an identifier that isn't in the master
        # verbatim but exists under a variant key (e.g. FirstCry sends EAN
        # '8906121640885' while the master has '8906121640885_1'). Rather than
        # patch each marketplace or edit the master, all such overrides live in
        # ONE sibling file ('Master Exceptions.xlsx') next to the Items Master;
        # lookup() consults this map after a direct miss. Empty when the file
        # is absent (feature simply inactive).
        self.exceptions: Dict[str, str] = {}
        self.exceptions_source: str = ''

        # v2.4.0: central PRICE overrides — {source_code → {mrp, margin_pct,
        # marketplace}}. Same exceptions file, different columns: when a
        # marketplace's agreed price doesn't match what MRP×default-margin
        # computes (e.g. Blinkit's EPISENSE is a 24% discount on a 899 deal
        # MRP, not the master's 1099 at 70%), the operator records the deal's
        # MRP + landing% here. Validation then expects that figure for the
        # row instead of flagging a mismatch. Optional ``marketplace`` scopes
        # the override to one channel (blank = all).
        self.price_overrides: Dict[str, Dict] = {}

        # v2.4.0: 'Use Vendor CP' exceptions — {source_code → marketplace}.
        # For these EANs the vendor's stated cost is authoritative: validation
        # accepts it (no MISMATCH) and the D365 Lines Unit Price is written
        # from the vendor cost instead of left blank. e.g. Myntra's RENEE
        # Goddess perfume. Optional marketplace scopes it (blank = all).
        self.vendor_cp_overrides: Dict[str, str] = {}

        # v2.4.3: the FULL exception registry — one entry per row of
        # 'Master Exceptions.xlsx', carrying every column verbatim plus a
        # derived ``effect`` summary and ``kinds`` list (alias / price /
        # vendor_cp). Unlike the lookup dicts above (which are keyed and
        # split by type for the engine), this preserves the whole list so
        # EVERY marketplace's output can render the complete cross-marketplace
        # exception list on its Exceptions sheet (highlighting the rows that
        # belong to the marketplace being processed). Empty when the file is
        # absent.
        self.exception_registry: List[Dict] = []

    # ── Loading ────────────────────────────────────────────────────────

    def load(self, filepath: str) -> int:
        """
        Read the master file and rebuild the lookup table.

        Args:
            filepath: Path to ``Items_March.xlsx``.

        Returns:
            Number of rows loaded (also the row count of the input).

        Required columns: ``No.``, ``GTIN``, ``Description``,
        ``GST Group Code``, ``Mrp`` — all on the **'Item Master'** sheet.

        v2.3.1: the master workbook now bundles many sheets (Offline
        Tracker, Zepto, Flipkart, …) and the first sheet is no longer the
        item master. The EAN→item mapping lives on the 'Item Master'
        sheet, so it's selected by name (case/space-insensitive). Older
        single-sheet masters that have no such sheet fall back to the
        first sheet, so existing setups are unaffected.
        """
        xl = pd.ExcelFile(filepath)
        sheet = next((s for s in xl.sheet_names
                      if str(s).strip().lower() == 'item master'), 0)
        df = pd.read_excel(xl, sheet_name=sheet, header=0)

        # Pre-stringify GTIN for use as a dict key. v2.3.1: route through
        # ``_clean_code`` so a GTIN column read as float64 (which happens
        # when any cell in the column is blank) keys as '8904473104659'
        # rather than '8904473104659.0' — otherwise EAN lookups (which
        # use the clean key) would silently miss.
        df['GTIN_str'] = df['GTIN'].map(self._clean_code)

        self.master = {}

        for _, r in df.iterrows():
            desc = (str(r.get('Description', ''))
                    if pd.notna(r.get('Description')) else '')
            gst = (str(r['GST Group Code'])
                   if pd.notna(r.get('GST Group Code')) else '')
            mrp = r.get('Mrp')
            # v2.3.1: clean the Item No so a float64 'No.' column (any
            # blank cell coerces the whole column to float) yields
            # '200570' rather than '200570.0'. This is the canonical ERP
            # Item No that flows to every output sheet (Lines, Validation,
            # Raw Data) and the D365 export — it must be a whole number.
            item_no = self._clean_code(r['No.'])

            # v1.6.0: HSN/SAC Code is optional in the master file —
            # only Reliance (so far) does HSN cross-checking, and the
            # master may not have the column at all if the customer
            # hasn't added it yet. Missing column or blank cell →
            # empty string, which the engine treats as "no master
            # HSN known". The HSN cross-check then reports
            # NOT_IN_MASTER for those rows so the user knows to
            # update the master.
            hsn_raw = r.get('HSN/SAC Code')
            hsn = ''
            if hsn_raw is not None and pd.notna(hsn_raw):
                # Master HSNs sometimes come in as floats (e.g.
                # 33049990.0) via Excel's number formatting. Strip
                # any trailing .0 so comparisons against the punch
                # file's string HSN don't mis-match.
                try:
                    hsn = str(int(float(hsn_raw)))
                except (ValueError, TypeError):
                    hsn = str(hsn_raw).strip()

            entry = {
                'item_no': item_no,
                'mrp': mrp,
                'gst_code': gst,
                'description': desc,
                'hsn': hsn,
            }

            # Index by GTIN. The GTIN is the marketplace-facing identifier,
            # so EAN lookups go through this key.
            self.master[r['GTIN_str']] = entry

            # Also index by item code so a punch file with pre-resolved
            # Item No (RK-style, when ``item_resolution='from_column'``)
            # can find the entry too. Don't overwrite an existing GTIN
            # match — GTIN is more specific.
            if item_no not in self.master:
                self.master[item_no] = entry

        # v2.4.0: auto-load the central exceptions file if it sits next to
        # the master. One file, edited by the operator, covers every
        # marketplace — no per-config or per-master patching.
        self._auto_load_exceptions(filepath)

        # v2.4.0 (Swiggy): load the 'Swiggy' sheet (SkuCode → EAN) and the
        # 'Swiggy Deal SKUs' sheet (per-EAN deal price), both optional.
        self._load_swiggy_sheets(xl)

        return len(df)

    # ── Swiggy sheets (SkuCode map + deal prices) ───────────────────────

    def _load_swiggy_sheets(self, xl) -> None:
        """Build the Swiggy SkuCode→EAN index and per-EAN deal overrides from
        the master workbook's 'Swiggy' and 'Swiggy Deal SKUs' sheets. Silent
        no-op when a sheet is absent — Swiggy support is then inactive."""
        names = {str(s).strip().lower(): s for s in xl.sheet_names}

        # 'Swiggy' sheet: SkuCode → EAN. Swiggy's PO dump carries only a
        # SkuCode (no EAN/Item No); this recovers the EAN so the standard
        # master lookup resolves it.
        sw = names.get('swiggy')
        if sw is not None:
            try:
                df = pd.read_excel(xl, sheet_name=sw, header=0, dtype=str)
                cols = {''.join(str(c).split()).lower(): c for c in df.columns}
                sc = cols.get('skucode') or cols.get('sku')
                ea = cols.get('ean')
                if sc and ea:
                    for _, r in df.iterrows():
                        skc = self._clean_code(r.get(sc))
                        ean = self._clean_code(r.get(ea))
                        if (skc and ean and skc.lower() != 'nan'
                                and ean.lower() != 'nan'):
                            self.swiggy_sku.setdefault(skc, ean)
            except Exception:  # noqa: BLE001 — overlay must never break load
                pass

        # 'Swiggy Deal SKUs' sheet: EAN → deal price (explicit cost, not the
        # default MRP×80%). Cost after GST = our CP; Cost With GST = inc-GST.
        deal = names.get('swiggy deal skus')
        if deal is not None:
            try:
                df = pd.read_excel(xl, sheet_name=deal, header=0, dtype=str)
                cols = {''.join(str(c).split()).lower(): c for c in df.columns}
                ea = cols.get('ean')
                if ea:
                    for _, r in df.iterrows():
                        ean = self._clean_code(r.get(ea))
                        if not ean or ean.lower() == 'nan':
                            continue
                        cag = self._to_float(r.get(cols.get('costaftergst')))
                        d_mrp = self._to_float(r.get(cols.get('correctmrp')))
                        self.swiggy_deals[ean] = {
                            'mrp': d_mrp,
                            'gst_pct': self._to_float(r.get(cols.get('correctgst'))),
                            'cost_after_gst': cag,
                            'cost_with_gst': self._to_float(
                                r.get(cols.get('costwithgst'))),
                        }
                        # v2.4.3: surface Swiggy deal SKUs on the Exceptions
                        # sheet too — they're per-SKU price exceptions, just
                        # stored in this master sheet rather than in
                        # 'Master Exceptions.xlsx'. Scoped to Swiggy.
                        name = ''
                        ncol = cols.get('name')
                        if ncol and pd.notna(r.get(ncol)):
                            name = str(r.get(ncol)).strip()
                        self.exception_registry.append({
                            'source_code': ean,
                            'maps_to': '',
                            'override_mrp': d_mrp,
                            'override_margin_pct': None,
                            'use_vendor_cp': False,
                            'marketplace': 'Swiggy',
                            'note': name,
                            'kinds': ['swiggy_deal'],
                            'effect': ('deal CP after GST '
                                       + (f'{cag:g}' if cag is not None else '—')
                                       + (f' (MRP {d_mrp:g})' if d_mrp is not None
                                          else '')),
                        })
            except Exception:  # noqa: BLE001
                pass

    def resolve_swiggy_sku(self, sku) -> Optional[Dict]:
        """SkuCode → master entry, via the 'Swiggy' sheet's SkuCode→EAN map.
        None when the SkuCode is unknown or its EAN isn't in the master."""
        ean = self.swiggy_sku.get(self._clean_code(sku))
        return self.lookup(ean) if ean else None

    # ── Exceptions (central EAN/code override file) ─────────────────────

    # Candidate filenames for the sibling exceptions workbook (first hit
    # wins), matched case-insensitively in the master's folder.
    _EXCEPTION_FILENAMES = (
        'Master Exceptions.xlsx', 'Master Exceptions.xls',
        'EAN Exceptions.xlsx', 'Exceptions.xlsx',
    )

    def _auto_load_exceptions(self, master_path: str) -> None:
        """Look for the exceptions workbook beside the Items Master and load
        it if present. Silent no-op when absent or unreadable — the master is
        the source of truth; exceptions are an optional overlay."""
        import os
        try:
            folder = os.path.dirname(os.path.abspath(master_path))
            existing = {f.lower(): f for f in os.listdir(folder)}
        except OSError:
            return
        for cand in self._EXCEPTION_FILENAMES:
            actual = existing.get(cand.lower())
            if actual:
                self.load_exceptions(os.path.join(folder, actual))
                return

    def load_exceptions(self, path: str) -> int:
        """
        Load the central exceptions file → ``{source_code → master_key}``.

        The file's first sheet maps an identifier as a marketplace sends it
        (``Source Code`` / ``Source EAN``) to the key that DOES exist in the
        master (``Maps To`` / ``Master Code`` — an EAN or Item No). Optional
        ``Marketplace`` / ``Note`` columns are for the operator's reference
        only. Column names are matched case/space-insensitively.

        Returns the number of overrides loaded (0 on any read/format error —
        the feature then stays inactive rather than breaking the run).
        """
        self.exceptions = {}
        try:
            df = pd.read_excel(path, sheet_name=0, header=0, dtype=str)
        except Exception:  # noqa: BLE001 — never let an overlay break loading
            return 0

        def _find(cands):
            for col in df.columns:
                if ''.join(str(col).split()).lower() in cands:
                    return col
            return None

        src_col = _find({'sourcecode', 'sourceean', 'source', 'from',
                         'dumpean', 'marketplaceean', 'sourcegtin'})
        dst_col = _find({'mapsto', 'mastercode', 'masterean', 'to', 'master',
                         'correctean', 'mastergtin', 'masterkey', 'itemno'})
        # Pricing-override columns (all optional).
        mrp_col = _find({'overridemrp', 'mrp', 'dealmrp', 'correctmrp'})
        margin_col = _find({'overridemargin%', 'overridemargin', 'margin%',
                            'margin', 'landing%', 'landingpct'})
        mp_col = _find({'marketplace', 'channel', 'mp'})
        # 'Use Vendor CP' flag column (optional).
        vcp_col = _find({'usevendorcp', 'vendorcp', 'usevendorcost',
                         'takevendorcp'})
        # 'Override Unit Price' column (optional) — a typed ₹ value forced straight
        # into the D365 Lines Unit Price for that SKU (operator's explicit override).
        oup_col = _find({'overrideunitprice', 'unitprice', 'overrideprice',
                         'forceunitprice'})
        # Free-text note column (optional, display-only).
        note_col = _find({'note', 'notes', 'remark', 'remarks', 'comment',
                          'comments', 'reason'})
        if not src_col:
            return 0

        self.price_overrides = {}
        self.vendor_cp_overrides = {}
        self.override_unit_prices = {}
        self.exception_registry = []
        for _, r in df.iterrows():
            src = self._clean_code(r.get(src_col))
            if not src or src.lower() == 'nan':
                continue
            mp = ''
            if mp_col and pd.notna(r.get(mp_col)):
                mp = str(r.get(mp_col)).strip()
            note = ''
            if note_col and pd.notna(r.get(note_col)):
                note = str(r.get(note_col)).strip()

            kinds: List[str] = []
            effects: List[str] = []

            # 'Use Vendor CP' exception (Y/yes/true/1).
            use_vcp = False
            if vcp_col and pd.notna(r.get(vcp_col)):
                if str(r.get(vcp_col)).strip().lower() in (
                        'y', 'yes', 'true', '1'):
                    self.vendor_cp_overrides[src] = mp
                    use_vcp = True
                    kinds.append('vendor_cp')
                    effects.append('accept vendor CP → Lines Unit Price')
            # 'Override Unit Price' exception — a typed ₹ value forced straight into
            # the D365 Lines Unit Price (highest precedence, see _process_row).
            oup_v = self._to_float(r.get(oup_col)) if oup_col else None
            if oup_v is not None:
                self.override_unit_prices[src] = {'value': oup_v, 'marketplace': mp}
                kinds.append('override_unit_price')
                effects.append(f'unit price → {oup_v:g}')
            # Item-alias override (Source → Master key).
            dst = ''
            if dst_col:
                d = self._clean_code(r.get(dst_col))
                if d and d.lower() != 'nan':
                    self.exceptions[src] = d
                    dst = d
                    kinds.append('item_alias')
                    effects.append(f'EAN remap → {d}')
            # Price override (deal MRP + landing%).
            mrp_v = self._to_float(r.get(mrp_col)) if mrp_col else None
            margin_v = self._to_float(r.get(margin_col)) if margin_col else None
            margin_pct = None
            if mrp_v is not None or margin_v is not None:
                # Margin given as a percent (76) → decimal (0.76).
                margin_pct = (margin_v / 100.0
                              if margin_v is not None and margin_v > 1.5
                              else margin_v)
                self.price_overrides[src] = {
                    'mrp': mrp_v,
                    'margin_pct': margin_pct,
                    'marketplace': mp,
                }
                kinds.append('price_override')
                effects.append(
                    'deal '
                    + (f'MRP {mrp_v:g}' if mrp_v is not None else 'MRP —')
                    + (f' @ {margin_pct*100:g}%' if margin_pct is not None
                       else ''))

            # One registry entry per row — the full cross-marketplace list
            # rendered on every output's Exceptions sheet.
            self.exception_registry.append({
                'source_code': src,
                'maps_to': dst,
                'override_mrp': mrp_v,
                'override_margin_pct': margin_pct,
                'use_vendor_cp': use_vcp,
                'override_unit_price': oup_v,
                'marketplace': mp,           # '' = applies to all channels
                'note': note,
                'kinds': kinds,
                'effect': '; '.join(effects) if effects else '(no effect)',
            })

        self.exceptions_source = path
        return (len(self.exceptions) + len(self.price_overrides)
                + len(self.vendor_cp_overrides))

    def use_vendor_cp(self, *keys, marketplace: str = '') -> bool:
        """True when a 'Use Vendor CP' exception applies to any of ``keys``
        (EAN / Item No) for ``marketplace`` — the vendor's stated cost is then
        accepted as-is (no MISMATCH) and written into the Lines Unit Price.
        A blank override marketplace applies everywhere."""
        if not self.vendor_cp_overrides:
            return False

        def _norm(s):
            return ''.join(str(s).split()).lower()
        want = _norm(marketplace)
        for k in keys:
            scope = self.vendor_cp_overrides.get(self._clean_code(k))
            if scope is not None:
                s = _norm(scope)
                if not s or s == want:
                    return True
        return False

    def override_unit_price(self, *keys, marketplace: str = ''):
        """The OPERATOR-typed unit price (float) to force into the D365 Lines Unit
        Price when an 'Override Unit Price' exception applies to any of ``keys``
        (EAN / Item No) for ``marketplace`` — else ``None``. Blank override
        marketplace applies everywhere. Mirrors :meth:`use_vendor_cp`."""
        pool = getattr(self, 'override_unit_prices', None)
        if not pool:
            return None

        def _norm(s):
            return ''.join(str(s).split()).lower()
        want = _norm(marketplace)
        for k in keys:
            hit = pool.get(self._clean_code(k))
            if hit is not None:
                s = _norm(hit.get('marketplace', ''))
                if not s or s == want:
                    return hit.get('value')
        return None

    @staticmethod
    def _to_float(val):
        """Parse a possibly-string numeric cell to float; None on blank/NaN."""
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return None
        try:
            s = str(val).replace(',', '').strip()
            return float(s) if s and s.lower() != 'nan' else None
        except (ValueError, TypeError):
            return None

    def price_override(self, *keys, marketplace: str = '') -> Optional[Dict]:
        """Return the price override for any of ``keys`` (EAN / Item No),
        honouring marketplace scope. A blank override marketplace applies to
        every channel; otherwise it must match ``marketplace`` (space/case-
        insensitive). None when no applicable override exists."""
        if not self.price_overrides:
            return None

        def _norm(s):
            return ''.join(str(s).split()).lower()
        want = _norm(marketplace)
        for k in keys:
            ov = self.price_overrides.get(self._clean_code(k))
            if not ov:
                continue
            scope = _norm(ov.get('marketplace', ''))
            if not scope or scope == want:
                return ov
        return None

    # ── Identifier cleaning ────────────────────────────────────────────

    @staticmethod
    def _clean_code(val) -> str:
        """
        Stringify an identifier (Item No / GTIN) WITHOUT the spurious
        trailing ``.0`` that pandas introduces when a numeric ID column
        is read as float64 (which happens whenever any cell in the
        column is blank). Examples::

            300069.0          → '300069'
            8904473104659.0   → '8904473104659'
            '200570.0'        → '200570'
            'ABC-12'          → 'ABC-12'   (non-numeric — passed through)
            NaN / None        → ''

        Only whole-number values are de-decimalised; a genuine
        fractional value (unexpected for an ID) is left as-is so nothing
        is silently corrupted.
        """
        if val is None:
            return ''
        if isinstance(val, float):
            if pd.isna(val):
                return ''
            return str(int(val)) if val.is_integer() else str(val).strip()
        if isinstance(val, int):
            return str(val)
        s = str(val).strip()
        # Float-looking string such as '200570.0' / '200570.00'.
        head, dot, tail = s.partition('.')
        if dot and head.lstrip('-').isdigit() and tail.strip('0') == '':
            return head
        return s

    # ── Lookup ─────────────────────────────────────────────────────────

    def lookup(self, key: str) -> Optional[Dict]:
        """
        Find an entry by EAN or Item No.

        Tries the cleaned key first, then falls back to leading-zero-
        stripped form (EANs sometimes have a leading zero in source data
        but not in the master).

        Args:
            key: Stringified EAN (e.g. ``'8906121642599'``) or Item No
                 (e.g. ``'200074'``).

        Returns:
            ``{item_no, mrp, gst_code, description}`` dict on hit, ``None``
            on miss.
        """
        key_clean = str(key).strip()
        if key_clean in self.master:
            return self.master[key_clean]

        # Some sources include a leading zero on EANs that the master
        # file omits — try the trimmed form before giving up.
        stripped = key_clean.lstrip('0')
        if stripped in self.master:
            return self.master[stripped]

        # v2.4.0: central exceptions overlay — the source code isn't in the
        # master verbatim, but the operator mapped it to a key that is (e.g.
        # FirstCry's '8906121640885' → master's '8906121640885_1'). Resolve
        # the alias, then look the mapped key up the normal way.
        if self.exceptions:
            mapped = self.exceptions.get(key_clean) or self.exceptions.get(stripped)
            if mapped:
                if mapped in self.master:
                    return self.master[mapped]
                mapped_stripped = mapped.lstrip('0')
                if mapped_stripped in self.master:
                    return self.master[mapped_stripped]

        return None

    # ── Pricing helpers (static) ───────────────────────────────────────
    # These are called on every row by the engine. Kept as static methods
    # so the engine doesn't need a loader instance to compute them, and
    # so they're trivially unit-testable.

    @staticmethod
    def calc_cost_price(mrp, gst_code: str,
                        margin_pct: float) -> Optional[float]:
        """
        Post-GST Cost Price: ``MRP × margin% ÷ GST divisor``.

        The GST divisor depends on the master's ``GST Group Code``:

        =========  =====  =========
        Code       GST    Divisor
        =========  =====  =========
        0-G        0%     ÷ 1.00
        G-3        3%     ÷ 1.03
        G-5(-S)    5%     ÷ 1.05
        G-12(-S)   12%    ÷ 1.12
        G-18(-S)   18%    ÷ 1.18
        Unknown    -      ÷ 1.18 (defaults to 18% with engine warning)
        =========  =====  =========

        Args:
            mrp: Maximum Retail Price (may be ``None`` or NaN).
            gst_code: Tax code from ``Items_March['GST Group Code']``.
            margin_pct: Margin as decimal (e.g. ``0.70`` for 70%).

        Returns:
            Calculated cost price, or ``None`` if MRP is missing.
        """
        if mrp is None or pd.isna(mrp):
            return None

        landing = float(mrp) * margin_pct
        return landing / MasterLoader.gst_divisor(gst_code)

    @staticmethod
    def row_gst_divisor(so_row: Any) -> float:
        """
        GST divisor ``(1 + rate)`` for an SORow's GST-inclusive amount.

        Prefers the per-line rate the PUNCH/PDF carried
        (``so_row.gst_rate_pct``, e.g. Reliance's IGST% from the PO) so the
        inc-GST order value matches the document's own total. Falls back to
        the master ``gst_code`` mapping when the punch carried no rate.
        """
        rate = getattr(so_row, 'gst_rate_pct', None)
        if rate is not None:
            try:
                return 1.0 + float(rate) / 100.0
            except (TypeError, ValueError):
                pass
        return MasterLoader.gst_divisor(getattr(so_row, 'gst_code', ''))

    @staticmethod
    def gst_divisor(gst_code: str) -> float:
        """
        GST divisor ``(1 + rate)`` for a master ``GST Group Code``.

        =========  =====  =========
        Code       GST    Divisor
        =========  =====  =========
        0-G / G-0  0%     1.00
        G-3(-S)    3%     1.03
        G-5(-S)    5%     1.05
        G-12(-S)   12%    1.12
        G-18(-S)   18%    1.18
        Unknown    -      1.18 (defaults to 18%)
        =========  =====  =========

        v2.3.1: extracted so callers that need the rate itself (e.g.
        Reliance's GST-dependent margin = 1 − discount × divisor) share
        the SAME mapping ``calc_cost_price`` uses — no drift.
        """
        gst = str(gst_code).strip().upper()
        # 0% GST — code variants seen in the wild
        if gst in ('0-G', 'G-0', 'G-0-S', '0', '') or gst == 'NAN':
            return 1.00
        # 3% GST
        if gst in ('G-3', 'G-3-S'):
            return 1.03
        # 5% GST — accept "5" in code as long as it's not 12 or 18
        if '5' in gst and '18' not in gst and '12' not in gst:
            return 1.05
        # 12% GST
        if '12' in gst:
            return 1.12
        # 18% GST
        if '18' in gst:
            return 1.18
        # Unknown code — default to 18% (engine emits a warning separately)
        return 1.18

    @staticmethod
    def calc_landing_price(mrp,
                           margin_pct: float) -> Optional[float]:
        """
        Pre-GST Landing Rate: ``MRP × margin%``. No GST divisor.

        Used by marketplaces whose price column is itself pre-GST (e.g.
        Myntra's "Landing Price"). Avoiding GST division means the diff
        comes out cleanly to zero on a correctly-priced punch — no
        floating-point rounding noise from ``÷ 1.18``.

        Args:
            mrp: Maximum Retail Price (may be ``None`` or NaN).
            margin_pct: Margin as decimal.

        Returns:
            ``MRP × margin%``, or ``None`` if MRP is missing.
        """
        if mrp is None or pd.isna(mrp):
            return None
        return float(mrp) * margin_pct