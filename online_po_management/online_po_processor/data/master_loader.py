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
from typing import Dict, Optional

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
        # Entry shape: {item_no, mrp, gst_code, description}
        self.master: Dict[str, Dict] = {}

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

        return len(df)

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