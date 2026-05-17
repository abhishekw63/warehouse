"""
MT Select Constructor — v1.0.0
================================

Generates D365-ready Sales Header + Sales Line workbooks from
customer Purchase Order files (currently: Naturals Salon).

Pipeline per run:

1. Operator picks customer from dropdown (auto-discovered from
   ``customers/<Name>/`` subfolders, each holding a
   ``<Name>_Master.xlsx`` with Code → EAN → MRP → GST).
2. Operator queues 1..N PO files via "Add PO Files".
3. Operator types or accepts the suggested starting SO number
   (``SO/NS/<MM>/<DDMYY>`` for today). Each PO in the batch
   increments by 1.
4. Engine, per file:
   - Resolves ship-to: filename code (``20673_3``) → direct lookup,
     else city token (``Bangalore``) → alias map (``Bengaluru``) →
     City lookup, else UNMATCHED.
   - Reads PO header block (PO No., embedded address) and data
     rows (Item Code, Qty).
   - Per line: Naturals Code → Naturals_Master → EAN →
     Items_March → RENEE Item No. Unresolved items keep the
     Naturals code in the Item No column for manual override.
5. Exporter writes a single workbook with three sheets:
   - ``Sales Header`` — 18 cols, one row per PO
   - ``Sales Line`` — 8 cols, all PO lines combined, Line No.
     stepping 10000/20000/… continuously across the workbook
   - ``Warnings`` — per-file section with ship-to resolution
     log + unmapped items

Unit Price is left blank in this version (D365 fills from item
master). The engine reads MRP/GST onto each row anyway so a
future ``margin_pct`` flag can turn on price calculation
without re-plumbing.

Design parallels the standalone EKA constructor and online
po_management's marketplace engine — same data-flow shape,
same warning surface, same "loud-not-silent" failure mode.
"""

from __future__ import annotations

import os
import re
import sys
import logging
import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Tkinter is imported lazily inside App.__init__ so the engine can
# run headlessly (CI / scripts / pytest) on systems where tkinter
# isn't installed. The frozen .exe bundle always has it.

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


__version__ = '1.0.0'

# ════════════════════════════════════════════════════════════════════
# Logging
# ════════════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
)
log = logging.getLogger('mt_select')


# ════════════════════════════════════════════════════════════════════
# Paths
# ════════════════════════════════════════════════════════════════════

def _app_dir() -> Path:
    """Return the folder containing this script (or the frozen exe)."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


APP_DIR = _app_dir()
CUSTOMERS_DIR = APP_DIR / 'customers'
CALC_DATA_DIR = APP_DIR / 'Calculation Data'
OUTPUT_DIR = APP_DIR / 'output'

# Shared masters (same files used by online_po_management).
ITEMS_MARCH_PATH = CALC_DATA_DIR / 'Items March.xlsx'
SHIP_TO_PATH = CALC_DATA_DIR / 'B2B_Ship_to_Addresses.xlsx'


# ════════════════════════════════════════════════════════════════════
# City alias map — used as fallback after direct city lookup misses.
# Maintained here in code; adding a city = one-line edit + reinstall.
# Format: {filename_spelling: master_spelling} (case-insensitive on lookup)
# ════════════════════════════════════════════════════════════════════

CITY_ALIASES: Dict[str, str] = {
    'Bangalore': 'Bengaluru',
    'Tirupur': 'Tripur',
    # Future entries go here. Example formats observed in the wild:
    #   'Gurgaon': 'Gurugram',
    #   'Bombay': 'Mumbai',
    #   'Calcutta': 'Kolkata',
}


# ════════════════════════════════════════════════════════════════════
# Data models
# ════════════════════════════════════════════════════════════════════

@dataclass
class ShipToResolution:
    """Result of looking up a ship-to from a filename token.

    `raw` is what we extracted from the filename (e.g. ``Bangalore``
    or ``20673_3``). `resolved_city` is what the B2B dump has on
    file (e.g. ``Bengaluru`` — may equal raw). `code` is the
    ship-to code (e.g. ``20673_3``) we'll write to the output, or
    empty string when status is UNMATCHED. `method` records HOW we
    resolved so the operator can audit on the Warnings sheet.
    """
    raw: str
    resolved_city: str = ''
    code: str = ''
    name: str = ''
    address_master: str = ''     # joined from B2B dump's Address fields
    method: str = 'UNMATCHED'    # Direct Code | Direct City | Alias (X→Y) | UNMATCHED


@dataclass
class POLine:
    """One row of a PO file after Naturals→EAN→Item-No resolution.

    `item_no_resolved` carries the RENEE Item No on success, or
    the Naturals code as-is when the chain breaks (so the operator
    sees something concrete to overwrite in D365). `status`
    distinguishes the failure mode for the Warnings sheet.
    """
    naturals_code: int
    product_name: str
    qty: int
    ean: str = ''
    item_no_resolved: Any = None     # int (RENEE Item No) or fallback string
    mrp: Optional[float] = None
    gst_decimal: Optional[float] = None
    status: str = 'OK'               # OK | NATURALS_CODE_UNMAPPED | EAN_NOT_IN_ITEMS_MASTER
    # Future: unit_price field, populated by _compute_unit_price().


@dataclass
class POResult:
    """All output for one input PO file.

    ``source_file`` is the bare filename (no path) — appears in
    output as-is. ``source_path`` is the full path on disk, kept
    for the Raw Data exporter's need to re-open the file for a
    full column echo (we don't cache the cell-level data on this
    object to keep the memory profile flat).
    """
    source_file: str
    source_path: Optional[Path] = None
    po_number: str = ''
    so_number: str = ''
    customer_address_from_po: str = ''
    ship_to: ShipToResolution = field(
        default_factory=lambda: ShipToResolution(raw=''))
    lines: List[POLine] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# ════════════════════════════════════════════════════════════════════
# Data loaders
# ════════════════════════════════════════════════════════════════════

class CustomerMasterLoader:
    """Loads <Customer>_Master.xlsx with Code/EAN/Name/MRP/GST.

    Headers are case-insensitive and stripped, so trailing-space
    drift in customer-supplied files doesn't break the loader.
    Code is normalized to int when possible; EAN to bare string
    (the file stores it as int64, which would render as
    '1234567890123.0' on str() coercion).
    """

    def __init__(self) -> None:
        self.by_code: Dict[int, Dict[str, Any]] = {}
        self.path: Optional[Path] = None
        self.customer_name: str = ''

    def load(self, path: Path, customer_name: str,
              logs: List[str]) -> int:
        self.path = path
        self.customer_name = customer_name
        self.by_code.clear()

        wb = load_workbook(str(path), read_only=False, data_only=True)
        ws = wb.active
        # Map header label → column index (1-based), case-insensitive.
        header_map = {
            str(ws.cell(1, c).value).strip().lower(): c
            for c in range(1, ws.max_column + 1)
            if ws.cell(1, c).value is not None
        }
        required = {'code', 'ean', 'full name', 'mrp', 'gst'}
        missing = required - set(header_map)
        if missing:
            raise ValueError(
                f"{customer_name} master is missing required column(s): "
                f"{missing}. Found: {list(header_map)}")

        for r in range(2, ws.max_row + 1):
            code_v = ws.cell(r, header_map['code']).value
            if code_v is None or str(code_v).strip() == '':
                continue
            try:
                code = int(str(code_v).strip())
            except (ValueError, TypeError):
                logs.append(
                    f"{customer_name} master row {r}: "
                    f"non-numeric Code {code_v!r} skipped")
                continue
            ean_v = ws.cell(r, header_map['ean']).value
            ean = ''
            if ean_v is not None:
                try:
                    ean = str(int(float(ean_v)))
                except (ValueError, TypeError):
                    ean = str(ean_v).strip()
            mrp_v = ws.cell(r, header_map['mrp']).value
            try:
                mrp = float(mrp_v) if mrp_v is not None else None
            except (ValueError, TypeError):
                mrp = None
            gst_v = ws.cell(r, header_map['gst']).value
            try:
                gst_decimal = float(gst_v) if gst_v is not None else None
            except (ValueError, TypeError):
                gst_decimal = None
            self.by_code[code] = {
                'ean': ean,
                'full_name': str(ws.cell(r, header_map['full name']).value
                                 or '').strip(),
                'mrp': mrp,
                'gst_decimal': gst_decimal,
            }
        log.info("Loaded %s master: %d codes", customer_name,
                 len(self.by_code))
        return len(self.by_code)

    def lookup(self, code: int) -> Optional[Dict[str, Any]]:
        return self.by_code.get(int(code))


class ItemsMarchLoader:
    """Loads the shared Items March master, keyed by EAN.

    Looser than online_po_management's version — for THIS tool we
    only need EAN → Item No. MRP/GST/Description are pulled too
    for the warnings sheet but aren't required.
    """

    def __init__(self) -> None:
        self.by_ean: Dict[str, Dict[str, Any]] = {}
        self.path: Optional[Path] = None

    def load(self, path: Path, logs: List[str]) -> int:
        self.path = path
        self.by_ean.clear()
        wb = load_workbook(str(path), read_only=False, data_only=True)
        ws = wb.active
        header_map = {
            str(ws.cell(1, c).value).strip().lower(): c
            for c in range(1, ws.max_column + 1)
            if ws.cell(1, c).value is not None
        }
        # Tolerant column lookup: production Items March uses 'No.'
        # for item number; our test stub uses 'No.' or 'No' — accept
        # either.
        no_col = header_map.get('no.') or header_map.get('no')
        ean_col = header_map.get('ean')
        if no_col is None or ean_col is None:
            raise ValueError(
                f"Items March missing 'No.' or 'EAN' column. "
                f"Found: {list(header_map)}")
        for r in range(2, ws.max_row + 1):
            ean_v = ws.cell(r, ean_col).value
            no_v = ws.cell(r, no_col).value
            if ean_v is None or no_v is None:
                continue
            try:
                ean = str(int(float(ean_v)))
            except (ValueError, TypeError):
                ean = str(ean_v).strip()
            try:
                item_no = int(no_v)
            except (ValueError, TypeError):
                item_no = str(no_v).strip()
            desc_col = header_map.get('description')
            mrp_col = header_map.get('mrp')
            self.by_ean[ean] = {
                'item_no': item_no,
                'description': (str(ws.cell(r, desc_col).value or '').strip()
                                if desc_col else ''),
                'mrp': ws.cell(r, mrp_col).value if mrp_col else None,
            }
        log.info("Loaded Items March: %d EANs", len(self.by_ean))
        return len(self.by_ean)

    def lookup(self, ean: str) -> Optional[Dict[str, Any]]:
        return self.by_ean.get(str(ean).strip())


class ShipToLoader:
    """Loads B2B Ship-to dump, indexed both by Code and by City.

    City index is case-insensitive (lowercased keys). Both indexes
    are populated from the same file in one pass; the hybrid
    resolver below picks which to consult.
    """

    def __init__(self) -> None:
        self.by_code: Dict[str, Dict[str, Any]] = {}
        self.by_city_lower: Dict[str, Dict[str, Any]] = {}
        self.path: Optional[Path] = None

    def load(self, path: Path, logs: List[str]) -> int:
        self.path = path
        self.by_code.clear()
        self.by_city_lower.clear()
        wb = load_workbook(str(path), read_only=False, data_only=True)
        ws = wb.active
        header_map = {
            str(ws.cell(1, c).value).strip().lower(): c
            for c in range(1, ws.max_column + 1)
            if ws.cell(1, c).value is not None
        }
        code_col = header_map.get('code')
        name_col = header_map.get('name')
        addr_col = header_map.get('address')
        addr2_col = header_map.get('address 2')
        city_col = header_map.get('city')
        postcode_col = header_map.get('postcode')
        if code_col is None or city_col is None:
            raise ValueError(
                f"B2B ship-to file missing 'Code' or 'City' column. "
                f"Found: {list(header_map)}")
        for r in range(2, ws.max_row + 1):
            code_v = ws.cell(r, code_col).value
            if code_v is None or str(code_v).strip() == '':
                continue
            code = str(code_v).strip()
            city = str(ws.cell(r, city_col).value or '').strip()
            entry = {
                'code': code,
                'name': str(ws.cell(r, name_col).value or '').strip()
                        if name_col else '',
                'address': str(ws.cell(r, addr_col).value or '').strip()
                           if addr_col else '',
                'address2': str(ws.cell(r, addr2_col).value or '').strip()
                            if addr2_col else '',
                'postcode': str(ws.cell(r, postcode_col).value or '').strip()
                            if postcode_col else '',
                'city': city,
            }
            self.by_code[code] = entry
            if city:
                # Multiple branches in one city would collide here.
                # First one wins; later ones are ignored with a log
                # note. Caller can disambiguate by passing the code
                # explicitly in the filename.
                key = city.lower()
                if key in self.by_city_lower:
                    logs.append(
                        f"Ship-to dump: duplicate city '{city}' "
                        f"({self.by_city_lower[key]['code']} vs {code}). "
                        f"First-seen wins. Use code-in-filename to "
                        f"disambiguate.")
                else:
                    self.by_city_lower[key] = entry
        log.info("Loaded Ship-to: %d codes, %d unique cities",
                 len(self.by_code), len(self.by_city_lower))
        return len(self.by_code)

    def lookup_by_code(self, code: str) -> Optional[Dict[str, Any]]:
        return self.by_code.get(str(code).strip())

    def lookup_by_city(self, city: str) -> Optional[Dict[str, Any]]:
        return self.by_city_lower.get(str(city).strip().lower())


# ════════════════════════════════════════════════════════════════════
# Ship-to resolver — the hybrid logic
# ════════════════════════════════════════════════════════════════════

# Match a ship-to code like '20673_3', '21645_HO', '12345_12'.
# Greedy on digits, allows alphanumeric suffix after the underscore.
# Bounded by (start | non-word) on the left and (end | non-word but
# NOT underscore-digit) on the right. We can't use \b alone — `_` is
# a word char in Python re, so `\b(\d+_\d+)\b` fails to match
# `20673_3` because there's no boundary between `_` and `3`. Solution:
# anchor on lookbehind/ahead for delimiters seen in filenames (-, _,
# start, end, dot, space).
_SHIPTO_CODE_RE = re.compile(
    r'(?:(?<=[\-_])|^)(\d{4,8}_[A-Za-z0-9]+?)(?=[\-_.\s]|$)'
)


def _extract_city_token(filename_stem: str) -> str:
    """Extract the city portion after the last '_-_' separator.

    ``Renee_PO_no_333_-_Bangalore`` → ``Bangalore``.
    Falls back to the trailing token after the last underscore if
    ``_-_`` isn't present.
    """
    if '_-_' in filename_stem:
        return filename_stem.rsplit('_-_', 1)[1].strip()
    # Fallback: try splitting on ' - ' (with spaces, in case
    # filename has spaces rather than underscores).
    if ' - ' in filename_stem:
        return filename_stem.rsplit(' - ', 1)[1].strip()
    # Final fallback: last underscore-separated token. May be wrong
    # but the resolver will catch it as UNMATCHED.
    return filename_stem.rsplit('_', 1)[-1].strip()


def resolve_ship_to(filename: str,
                     ship_to_loader: ShipToLoader,
                     aliases: Dict[str, str]) -> ShipToResolution:
    """Hybrid resolver: try code-in-filename, then city + alias map.

    Order matters: code in filename is unambiguous (operator
    explicitly set it), so it wins over any city extraction. City
    fallback runs only when no code pattern is present.
    """
    stem = Path(filename).stem

    # ── Route 1: explicit ship-to code in filename ──────────────────
    m = _SHIPTO_CODE_RE.search(stem)
    if m:
        code = m.group(1)
        entry = ship_to_loader.lookup_by_code(code)
        if entry:
            return ShipToResolution(
                raw=code,
                resolved_city=entry['city'],
                code=entry['code'],
                name=entry['name'],
                address_master=_join_address(entry),
                method='Direct Code',
            )
        # Code in filename but not in master — surface as UNMATCHED
        # rather than silently falling through to city extraction
        # (which could happen to match wrong branch).
        return ShipToResolution(
            raw=code,
            method=f'UNMATCHED (code {code} not in ship-to dump)',
        )

    # ── Route 2: city token + alias map ─────────────────────────────
    city_token = _extract_city_token(stem)
    if not city_token:
        return ShipToResolution(raw='', method='UNMATCHED (no token in filename)')

    # 2a. Direct city lookup
    entry = ship_to_loader.lookup_by_city(city_token)
    if entry:
        return ShipToResolution(
            raw=city_token,
            resolved_city=entry['city'],
            code=entry['code'],
            name=entry['name'],
            address_master=_join_address(entry),
            method='Direct City',
        )

    # 2b. Alias map lookup (case-insensitive on the key)
    aliased = None
    aliased_to = ''
    for k, v in aliases.items():
        if k.lower() == city_token.lower():
            aliased_to = v
            aliased = ship_to_loader.lookup_by_city(v)
            break
    if aliased:
        return ShipToResolution(
            raw=city_token,
            resolved_city=aliased['city'],
            code=aliased['code'],
            name=aliased['name'],
            address_master=_join_address(aliased),
            method=f'Alias ({city_token}→{aliased_to})',
        )

    return ShipToResolution(
        raw=city_token,
        method=f'UNMATCHED (no city/alias for {city_token!r})',
    )


def _join_address(entry: Dict[str, Any]) -> str:
    """Build a single-line address from a ship-to entry."""
    parts = [entry.get('name', ''), entry.get('address', ''),
             entry.get('address2', ''), entry.get('city', ''),
             entry.get('postcode', '')]
    return ' | '.join(p for p in parts if p)


# ════════════════════════════════════════════════════════════════════
# PO file reader + engine
# ════════════════════════════════════════════════════════════════════

# These cells in the PO header block hold known values.
# Tested against the 3 sample files; all use the same template.
_PO_HEADER_CELLS = {
    'company_name': (2, 6),
    'addr1': (3, 6),
    'addr2': (4, 6),
    'addr3': (5, 6),
    'gst': (6, 6),
    'po_no': (10, 12),    # text: 'PO No. : HO/26/PO-333'
    'date': (11, 12),     # text: 'Date:15-05-2026'
}
_DATA_HEADER_ROW = 15
_DATA_START_ROW = 16


def read_po_file(filepath: Path,
                  customer_master: CustomerMasterLoader,
                  items_master: ItemsMarchLoader) -> POResult:
    """Parse one PO file end-to-end.

    Returns a POResult with all lines (resolved or not) and any
    per-file warnings. Ship-to resolution is NOT done here — the
    GUI/runner handles that and stamps it onto the result.
    """
    res = POResult(source_file=filepath.name, source_path=filepath)
    wb = load_workbook(str(filepath), data_only=True)
    ws = wb.active

    # ── Header block ────────────────────────────────────────────────
    def _hv(key: str) -> str:
        r, c = _PO_HEADER_CELLS[key]
        v = ws.cell(r, c).value
        return str(v).strip() if v is not None else ''

    # PO No. — strip 'PO No. : ' prefix if present, defensive about
    # whitespace and case variations.
    po_no_raw = _hv('po_no')
    m = re.match(r'^\s*PO\s*No\.?\s*:\s*(.+)$', po_no_raw, re.IGNORECASE)
    res.po_number = m.group(1).strip() if m else po_no_raw
    if not res.po_number:
        res.warnings.append(
            f"PO No. cell ({_PO_HEADER_CELLS['po_no']}) is blank in "
            f"'{filepath.name}'. External Document No. will be empty.")

    res.customer_address_from_po = ' | '.join(
        s for s in (_hv('company_name'), _hv('addr1'), _hv('addr2'),
                    _hv('addr3'), _hv('gst'))
        if s
    )

    # ── Sanity-check the data-table header row ──────────────────────
    expected_headers = {'item code', 'product name', 'qty'}
    actual_headers = {
        str(ws.cell(_DATA_HEADER_ROW, c).value or '').strip().lower()
        for c in range(1, ws.max_column + 1)
    }
    missing = expected_headers - actual_headers
    if missing:
        res.warnings.append(
            f"PO file '{filepath.name}' is missing required header(s) "
            f"on row {_DATA_HEADER_ROW}: {missing}. "
            f"Found: {sorted(s for s in actual_headers if s)}. "
            f"No lines will be emitted from this file.")
        return res

    # Find the actual column indices for the data table — defensive
    # against extra cols (POs 334/335 had manual EAN/ITEM columns
    # appended; we ignore those entirely and resolve fresh).
    header_to_col: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(_DATA_HEADER_ROW, c).value
        if v is not None:
            header_to_col[str(v).strip().lower()] = c
    col_code = header_to_col['item code']
    col_name = header_to_col['product name']
    col_qty = header_to_col['qty']

    # ── Data rows ───────────────────────────────────────────────────
    unmapped_codes_seen: set = set()
    missing_eans_seen: set = set()

    for r in range(_DATA_START_ROW, ws.max_row + 1):
        code_v = ws.cell(r, col_code).value
        if code_v is None or str(code_v).strip() == '':
            continue
        try:
            code = int(str(code_v).strip())
        except (ValueError, TypeError):
            # Non-numeric Item Code = footer/totals row, stop reading.
            break

        qty_v = ws.cell(r, col_qty).value
        try:
            qty = int(float(qty_v)) if qty_v is not None else 0
        except (ValueError, TypeError):
            qty = 0
        if qty <= 0:
            res.warnings.append(
                f"Row {r}: code {code} has qty={qty_v!r}. "
                f"Skipped (qty must be > 0).")
            continue

        product_name = str(ws.cell(r, col_name).value or '').strip()

        line = POLine(naturals_code=code, product_name=product_name,
                       qty=qty)

        # Naturals → EAN
        cm = customer_master.lookup(code)
        if cm is None:
            line.item_no_resolved = code  # write Naturals code as-is
            line.status = 'NATURALS_CODE_UNMAPPED'
            if code not in unmapped_codes_seen:
                unmapped_codes_seen.add(code)
                res.warnings.append(
                    f"Customer code {code} ('{product_name[:60]}') not "
                    f"found in {customer_master.customer_name} master. "
                    f"Item No column will carry the customer's code — "
                    f"overwrite manually in D365 before posting.")
            res.lines.append(line)
            continue

        line.ean = cm['ean']
        line.mrp = cm['mrp']
        line.gst_decimal = cm['gst_decimal']

        # EAN → RENEE Item No
        im = items_master.lookup(line.ean)
        if im is None:
            line.item_no_resolved = f'?EAN:{line.ean}'
            line.status = 'EAN_NOT_IN_ITEMS_MASTER'
            if line.ean not in missing_eans_seen:
                missing_eans_seen.add(line.ean)
                res.warnings.append(
                    f"EAN {line.ean} (customer code {code}, "
                    f"'{product_name[:50]}') resolved from "
                    f"{customer_master.customer_name} master but not "
                    f"found in Items March. Item No column will carry "
                    f"placeholder '?EAN:{line.ean}'.")
        else:
            line.item_no_resolved = im['item_no']

        res.lines.append(line)

    return res


# ════════════════════════════════════════════════════════════════════
# SO number generator
# ════════════════════════════════════════════════════════════════════

def suggest_starting_so_number(today: Optional[dt.date] = None) -> str:
    """Build today's suggested starting SO number per Vishal's convention.

    Format: ``SO/NS/<MM>/<DDMYY>`` where:
      - MM = month zero-padded ('05')
      - DD = day zero-padded ('16')
      - M  = month NOT zero-padded ('5' for May)
      - YY = 2-digit year ('26')
    """
    today = today or dt.date.today()
    mm = f'{today.month:02d}'
    dd = f'{today.day:02d}'
    m = str(today.month)
    yy = f'{today.year % 100:02d}'
    return f'SO/NS/{mm}/{dd}{m}{yy}'


def increment_so_number(so_number: str) -> str:
    """Increment the trailing numeric segment of an SO number by 1.

    Works on any format ending in digits: ``SO/NS/05/16526`` →
    ``16527``. If no trailing digits found, raises ValueError.
    """
    m = re.match(r'^(.*?)(\d+)$', so_number)
    if not m:
        raise ValueError(
            f"SO number {so_number!r} doesn't end in digits — "
            f"can't auto-increment.")
    prefix, num_str = m.group(1), m.group(2)
    next_num = int(num_str) + 1
    # Preserve leading zeros if any (unlikely for our format but
    # defensive — '00005' should stay 5-digit on increment).
    return f'{prefix}{next_num:0{len(num_str)}d}'


def derive_sell_to(ship_to_code: str) -> str:
    """Extract root customer number from ship-to code.

    ``20673_3`` → ``20673``, ``20673_HO`` → ``20673``,
    ``20673`` → ``20673`` (already root).
    """
    if not ship_to_code:
        return ''
    return ship_to_code.rsplit('_', 1)[0]


# ════════════════════════════════════════════════════════════════════
# Exporter
# ════════════════════════════════════════════════════════════════════
#
# Output workbook matches the existing online_po_management house
# style (see e.g. ``blink_so_16-05-2026_111954.xlsx`` for the canonical
# Blink output). Five sheets, all with headers on row 1 and data
# starting row 2. No title-block rows at the top of any sheet — that
# was the manual SO's convention; the engine-generated output uses
# the leaner Blink shape that D365 imports without complaint.
#
# Sheet inventory:
#   1. Headers (SO) — 18 cols, one row per PO file
#   2. Lines (SO)   — 8 cols, all PO lines combined, Line No.
#                     stepping by 10000 continuously across the
#                     workbook (not per-document)
#   3. Summary      — one row per PO + TOTAL row + metadata footer
#   4. Validation   — stub today (Unit Price calculation off); will
#                     populate per-line price comparison when
#                     ``margin_pct`` is enabled. Also displays
#                     unmapped items as red-fill rows so operators
#                     spot them at a glance
#   5. Raw Data     — combined echo of every PO file's source data
#                     table + appended resolved columns (EAN,
#                     Item No, Status). Cross-PO audit trail
# ════════════════════════════════════════════════════════════════════

# Column schema for Headers (SO). Order matters — these go into D365
# verbatim. Don't reorder without verifying against the D365 import
# spec; column position is significant.
_SALES_HEADER_COLS = [
    'Document Type', 'No.', 'Sell-to Customer No.', 'Ship-to Code',
    'Posting Date', 'Order Date', 'Document Date',
    'Invoice From Date', 'Invoice To Date',
    'External Document No.', 'Location Code', 'Dimension Set ID',
    'Supply Type', 'Voucher Narration',
    'Brand Code (Dimension)', 'Channel Code (Dimension)',
    'Catagory (Dimension)', 'Geography Code (Dimension)',
]

# Column schema for Lines (SO). Same order-sensitivity applies.
_SALES_LINE_COLS = [
    'Document Type', 'Document No.', 'Line No.', 'Type', 'No.',
    'Location Code', 'Quantity', 'Unit Price',
]

# Summary sheet columns. ``Location (Raw)`` is what the filename
# carried; ``Location (Mapped)`` is what the B2B dump resolved to.
# Logging both lets the operator verify the resolution at a glance
# without having to open the B2B dump.
_SUMMARY_COLS = [
    'PO', 'Location (Raw)', 'Location (Mapped)', 'Cust No',
    'Ship-to', 'Items', 'Total Qty', 'Total Amount', 'Status',
]

# Validation sheet columns. Used in two modes:
#   1. Stub (margin_pct is None): single explanatory row plus any
#      unmapped items in red.
#   2. Active (margin_pct set): per-line MRP × margin ÷ (1+GST) calc
#      shown alongside the customer's stated price for comparison.
#      Not implemented yet — wired but commented as TODO.
_VALIDATION_COLS = [
    'PO', 'Item No', 'EAN', 'Description', 'MRP',
    'Landing', 'GST Code', 'Our Cost Price',
    'Customer Cost', 'Difference', 'Status',
]

# ────────────────────────────────────────────────────────────────────
# Styles. Colours match the online_po_management palette so outputs
# look consistent across tools — same blue header band, same amber
# warning fill, same red error fill.
# ────────────────────────────────────────────────────────────────────

_HEADER_FILL = PatternFill(start_color='1F4E78', end_color='1F4E78',
                            fill_type='solid')
_HEADER_FONT = Font(bold=True, color='FFFFFF', size=11)
_WARN_FILL = PatternFill(start_color='FFF3CD', end_color='FFF3CD',
                          fill_type='solid')
_ERR_FILL = PatternFill(start_color='F8D7DA', end_color='F8D7DA',
                         fill_type='solid')
_SUBHEADER_FILL = PatternFill(start_color='D9E1F2', end_color='D9E1F2',
                               fill_type='solid')
_TOTAL_FILL = PatternFill(start_color='E2EFDA', end_color='E2EFDA',
                           fill_type='solid')
_TOTAL_BORDER = Border(top=Side(style='medium', color='548235'))
_BORDER_THIN = Border(
    left=Side(style='thin', color='999999'),
    right=Side(style='thin', color='999999'),
    top=Side(style='thin', color='999999'),
    bottom=Side(style='thin', color='999999'),
)
_INFO_ITALIC = Font(italic=True, color='666666')


def _today_str() -> str:
    """Date in DD-MM-YYYY format, matching the Blink output sample.

    Used for all five date columns on Headers (SO) and for the
    metadata footer on Summary. Centralised so we never get drift
    between sheets — same instant, same format.
    """
    return dt.date.today().strftime('%d-%m-%Y')


def _now_str() -> str:
    """Timestamp for metadata footer: DD-MM-YYYY HH:MM."""
    return dt.datetime.now().strftime('%d-%m-%Y %H:%M')


def _compute_unit_price(line: 'POLine',
                         margin_pct: Optional[float]) -> Optional[float]:
    """Compute the engine's Unit Price for one PO line.

    Returns ``None`` today because every customer ships with
    ``margin_pct=None`` — D365 fills Unit Price from its own item
    master at import time. Flip this behaviour on per-customer by
    passing a non-None margin to ``run_batch`` / ``export_workbook``;
    the formula is ``MRP × margin% ÷ (1 + GST_decimal)`` exactly as
    online_po_management's marketplace engine uses.

    When the margin is on but the line is missing MRP or GST (i.e.
    Naturals master had no value for those fields), returns None
    rather than guessing — the cell stays blank and the operator
    sees the gap on the Validation sheet.
    """
    if margin_pct is None:
        return None
    if line.mrp is None or line.gst_decimal is None:
        return None
    return round(line.mrp * margin_pct / (1 + line.gst_decimal), 2)


def export_workbook(results: List[POResult],
                     output_path: Path,
                     customer_name: str,
                     margin_pct: Optional[float] = None) -> Path:
    """Write the final five-sheet output workbook.

    Args:
        results: One :class:`POResult` per input PO file, in the
            order the operator queued them. Each carries its own
            resolved ship-to, lines, warnings, and SO number
            (assigned by the orchestrator before export).
        output_path: Where to write the .xlsx. Parent dir is
            created if missing.
        customer_name: Used on the Summary metadata footer
            (``Customer: Naturals``). Not used for routing logic.
        margin_pct: Currently always None (Unit Price stays blank).
            See :func:`_compute_unit_price` for the future-on
            behaviour.

    Returns:
        ``output_path`` (for caller convenience).
    """
    wb = Workbook()
    # Remove the default sheet; we create sheets explicitly in the
    # order they should appear in the tab strip.
    wb.remove(wb.active)

    _write_headers_so_sheet(wb, results)
    _write_lines_so_sheet(wb, results, margin_pct)
    _write_summary_sheet(wb, results, customer_name, margin_pct)
    _write_validation_sheet(wb, results, margin_pct)
    _write_raw_data_sheet(wb, results)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))
    log.info("Wrote output: %s", output_path)
    return output_path


# ────────────────────────────────────────────────────────────────────
# Per-sheet writers — each is a self-contained function so the
# top-level export_workbook reads as a clean orchestration step.
# Internal-only; not part of the public API.
# ────────────────────────────────────────────────────────────────────

def _write_headers_so_sheet(wb: Workbook,
                              results: List[POResult]) -> None:
    """Sheet 1: Headers (SO). One row per PO file.

    Every PO becomes a Sales Order header in D365. Fields:
      - Document Type    → always 'Order'
      - No.              → engine-generated SO number
                           (e.g. ``SO/NS/05/16526``)
      - Sell-to Cust No. → derived from ship-to code
                           (``20673_3`` → ``20673``)
      - Ship-to Code     → from B2B dump resolution
                           (blank if UNMATCHED)
      - 5× Date cols     → today's date in DD-MM-YYYY
      - Ext Doc No.      → PO No. from PO file header (e.g.
                           ``HO/26/PO-333``)
      - Location Code    → always 'PICK' (warehouse code)
      - Supply Type      → always 'B2B'
      - Dimension/Voucher columns left blank — set in D365
    """
    ws = wb.create_sheet('Headers (SO)')
    _write_header_row(ws, 1, _SALES_HEADER_COLS)

    today = _today_str()
    for i, r in enumerate(results, start=0):
        row = 2 + i
        sell_to = derive_sell_to(r.ship_to.code)
        ws.cell(row, 1, 'Order')
        ws.cell(row, 2, r.so_number)
        # Sell-to numeric when possible (D365 expects int for the
        # root customer code); string fallback covers customers with
        # alpha codes if that ever happens.
        ws.cell(row, 3, int(sell_to) if sell_to.isdigit() else sell_to)
        ws.cell(row, 4, r.ship_to.code)
        for c in (5, 6, 7, 8, 9):     # Posting / Order / Doc /
            ws.cell(row, c, today)      # Invoice From / Invoice To
        ws.cell(row, 10, r.po_number)
        ws.cell(row, 11, 'PICK')
        # col 12 (Dimension Set ID) intentionally blank
        ws.cell(row, 13, 'B2B')
        # cols 14-18 intentionally blank
        # Visual cue: rows with UNMATCHED ship-to get amber tint on
        # col 4 so they stand out before D365 import.
        if not r.ship_to.code:
            ws.cell(row, 4).fill = _WARN_FILL

    _auto_size(ws)
    ws.freeze_panes = 'A2'


def _write_lines_so_sheet(wb: Workbook, results: List[POResult],
                            margin_pct: Optional[float]) -> None:
    """Sheet 2: Lines (SO). All PO lines combined, in queue order.

    Line No. steps by 10000 continuously across the entire workbook
    (not per-document — this matches the manual reference SO; PO 333
    occupies 10000-290000, PO 334 takes 300000-540000, etc.).

    Item No. (col 5) carries the resolved RENEE Item No on success.
    On failure it carries either the customer's own code (Naturals
    code unmappable) or a ``?EAN:xxxx`` placeholder (EAN found in
    customer master but not in Items March) — both highlighted in
    red/amber so the operator can spot them and overwrite in the
    D365 import preview.

    Unit Price is computed via ``_compute_unit_price`` — None today
    leaves the cell blank, ready for D365 to fill from item master.
    """
    ws = wb.create_sheet('Lines (SO)')
    _write_header_row(ws, 1, _SALES_LINE_COLS)

    row_cursor = 2
    # Line No. is continuous across the workbook, stepping by 10000
    # every emitted line regardless of which PO it belongs to. Don't
    # reset per-document — D365 import handles this just fine and
    # the manual reference SO does the same.
    line_no = 10000
    for r in results:
        for line in r.lines:
            ws.cell(row_cursor, 1, 'Order')
            ws.cell(row_cursor, 2, r.so_number)
            ws.cell(row_cursor, 3, line_no)   # numeric int — matches
                                                # online_po_management
                                                # output, not stringified
            ws.cell(row_cursor, 4, 'Item')
            ws.cell(row_cursor, 5, line.item_no_resolved)
            ws.cell(row_cursor, 6, 'PICK')
            ws.cell(row_cursor, 7, line.qty)
            ws.cell(row_cursor, 8, _compute_unit_price(line, margin_pct))
            # Highlight item-resolution failures so they're impossible
            # to miss when reviewing the output before D365 import.
            if line.status == 'NATURALS_CODE_UNMAPPED':
                ws.cell(row_cursor, 5).fill = _ERR_FILL
            elif line.status == 'EAN_NOT_IN_ITEMS_MASTER':
                ws.cell(row_cursor, 5).fill = _WARN_FILL
            row_cursor += 1
            line_no += 10000

    _auto_size(ws)
    ws.freeze_panes = 'A2'


def _write_summary_sheet(wb: Workbook, results: List[POResult],
                           customer_name: str,
                           margin_pct: Optional[float]) -> None:
    """Sheet 3: Summary. One row per PO + TOTAL + metadata footer.

    The Summary sheet is the operator's first stop after generation
    — it shows whether every PO resolved cleanly without forcing
    them to scroll through the bulk Lines sheet. Status is OK when
    ship-to resolved and zero unmapped items; otherwise WARN or
    FAIL.

    Metadata footer (one row below the TOTAL row) carries audit
    info: customer name, margin setting, warehouse, file count,
    timestamp. Mirrors the Blink output's metadata convention.
    """
    ws = wb.create_sheet('Summary')
    _write_header_row(ws, 1, _SUMMARY_COLS)

    row = 2
    total_items = 0
    total_qty = 0
    total_amount = 0.0

    for r in results:
        items = len(r.lines)
        qty = sum(l.qty for l in r.lines)
        # Total Amount uses the engine's calc (None → 0 in the sum)
        # but only when margin_pct is on. With margin off, the
        # column is meaningless; show blanks instead of zeros.
        if margin_pct is not None:
            amount = sum(
                ((_compute_unit_price(l, margin_pct) or 0) * l.qty)
                for l in r.lines
            )
        else:
            amount = None

        # Status decision: ship-to UNMATCHED is the loudest signal
        # (header will lack a Ship-to Code). Item failures are
        # softer — the lines still emit, just with placeholders.
        if r.ship_to.method.startswith('UNMATCHED'):
            status = 'FAIL (no ship-to)'
            row_fill = _ERR_FILL
        elif any(l.status != 'OK' for l in r.lines):
            n_bad = sum(1 for l in r.lines if l.status != 'OK')
            status = f'WARN ({n_bad} unmapped item(s))'
            row_fill = _WARN_FILL
        else:
            status = 'OK'
            row_fill = None

        sell_to = derive_sell_to(r.ship_to.code)
        ws.cell(row, 1, r.po_number)
        ws.cell(row, 2, r.ship_to.raw)
        ws.cell(row, 3, r.ship_to.resolved_city)
        ws.cell(row, 4, sell_to)
        ws.cell(row, 5, r.ship_to.code)
        ws.cell(row, 6, items)
        ws.cell(row, 7, qty)
        ws.cell(row, 8, amount)
        ws.cell(row, 9, status)
        if row_fill is not None:
            for c in range(1, len(_SUMMARY_COLS) + 1):
                ws.cell(row, c).fill = row_fill

        total_items += items
        total_qty += qty
        if amount is not None:
            total_amount += amount
        row += 1

    # TOTAL row — green tint and a medium top border so it visually
    # separates from data even when row colours vary.
    ws.cell(row, 1, 'TOTAL').font = Font(bold=True)
    ws.cell(row, 6, total_items).font = Font(bold=True)
    ws.cell(row, 7, total_qty).font = Font(bold=True)
    if margin_pct is not None:
        ws.cell(row, 8, total_amount).font = Font(bold=True)
    for c in range(1, len(_SUMMARY_COLS) + 1):
        ws.cell(row, c).fill = _TOTAL_FILL
        ws.cell(row, c).border = _TOTAL_BORDER

    # Metadata footer — one blank row below TOTAL, then the audit
    # line. Mirrors Blink output's convention.
    row += 2
    margin_str = (f'{margin_pct * 100:.0f}%'
                  if margin_pct is not None else '—')
    file_list = ', '.join(r.source_file for r in results)
    if len(file_list) > 120:
        file_list = f'{len(results)} file(s)'
    metadata = (
        f"Customer: {customer_name}  |  "
        f"Margin: {margin_str}  |  "
        f"Warehouse: PICK  |  "
        f"Files: {file_list}  |  "
        f"Generated: {_now_str()}"
    )
    cell = ws.cell(row, 1, metadata)
    cell.font = _INFO_ITALIC
    ws.merge_cells(start_row=row, start_column=1,
                    end_row=row, end_column=len(_SUMMARY_COLS))

    _auto_size(ws)
    ws.freeze_panes = 'A2'


def _write_validation_sheet(wb: Workbook, results: List[POResult],
                              margin_pct: Optional[float]) -> None:
    """Sheet 4: Validation.

    Two modes:

    1. **Stub mode** (current — ``margin_pct=None``): displays a
       single explanatory row plus any unmapped items in red. This
       gives operators visibility into resolution failures without
       littering the Lines sheet with status columns.

    2. **Active mode** (future — ``margin_pct=0.60`` etc.): one row
       per emitted line, showing MRP, Landing rate, GST Code,
       engine-calculated Cost Price, customer's stated price, diff,
       Status. Same layout as the Blink Validation sheet so muscle
       memory transfers between tools.

    The empty-but-present-with-explanation approach is deliberate.
    Hiding the sheet entirely when off would surprise the operator
    later when it appears; leaving a note tells them "this feature
    exists, here's how to turn it on".
    """
    ws = wb.create_sheet('Validation')
    _write_header_row(ws, 1, _VALIDATION_COLS)

    if margin_pct is None:
        # Stub mode — explanatory row + any unmapped items.
        row = 2
        explanation = (
            'Unit Price calculation not enabled for this customer. '
            'Validation sheet will populate per-line cost-price '
            'comparison when margin_pct is set in run_batch().'
        )
        cell = ws.cell(row, 1, explanation)
        cell.font = _INFO_ITALIC
        cell.alignment = Alignment(wrap_text=True, vertical='center')
        ws.merge_cells(start_row=row, start_column=1,
                        end_row=row, end_column=len(_VALIDATION_COLS))
        ws.row_dimensions[row].height = 32
        row += 2

        # Unmapped items — red fill on the status cell so operators
        # see the count + which lines need manual override.
        unmapped_present = False
        for r in results:
            for line in r.lines:
                if line.status == 'OK':
                    continue
                if not unmapped_present:
                    # Section heading — only emit if there's content
                    ws.cell(row, 1, 'Items requiring manual review').font = (
                        Font(bold=True, color='C00000'))
                    row += 1
                    unmapped_present = True

                ws.cell(row, 1, r.po_number)
                ws.cell(row, 2, line.naturals_code)   # customer code, not RENEE
                ws.cell(row, 3, line.ean)
                ws.cell(row, 4, line.product_name)
                ws.cell(row, 5, line.mrp)
                # cols 6-9 blank (no Landing/GST/cost without margin)
                ws.cell(row, 11, line.status)
                # Whole-row fill — red for hard fails, amber for soft.
                fill = (_ERR_FILL
                        if line.status == 'NATURALS_CODE_UNMAPPED'
                        else _WARN_FILL)
                for c in range(1, len(_VALIDATION_COLS) + 1):
                    ws.cell(row, c).fill = fill
                row += 1

    else:
        # Active mode — per-line comparison. TODO: implement when
        # the first customer turns on margin_pct. Layout planned:
        # iterate r.lines, compute landing = mrp * margin, cost =
        # _compute_unit_price(line, margin_pct), populate the 11
        # columns. Status = OK when |diff| ≤ tolerance, MISMATCH
        # otherwise. Keep the unmapped-items section as a footer.
        row = 2
        ws.cell(row, 1,
                f'Active validation mode (margin {margin_pct * 100:.0f}%) '
                f'not yet implemented. Falling back to stub display.').font = (
            _INFO_ITALIC)
        ws.merge_cells(start_row=row, start_column=1,
                        end_row=row, end_column=len(_VALIDATION_COLS))

    _auto_size(ws)
    ws.freeze_panes = 'A2'


def _write_raw_data_sheet(wb: Workbook, results: List[POResult]) -> None:
    """Sheet 5: Raw Data. Full audit echo of every PO source file.

    Columns are: ``Source File, PO No.`` (added by the engine)
    followed by every column the PO file's data table carried
    (``Item Code, Product Name, MRP, Rate, Qty, UOM, Disc %,
    Amount, CGST %, SGST %, IGST %, Gross Amount``), followed by
    resolved columns added by the engine (``EAN, Item No (Master),
    Status``).

    Multiple PO files are concatenated into a single sheet — easier
    to filter and search than per-file tabs, and ``Source File`` /
    ``PO No.`` columns identify the origin of every row.

    The "manual annotation" columns (``EAN``, ``ITEM``) that Vishal
    used to add to PO 334/335 by hand are EXCLUDED from this echo
    because they're not part of the standard PO format — the
    engine resolves these fresh from the masters and writes the
    authoritative values into the appended ``EAN`` and
    ``Item No (Master)`` columns at the right.
    """
    ws = wb.create_sheet('Raw Data')

    # Standard column set from the PO data table (rows 15+ in input
    # files). Hard-coded rather than auto-detected from the first
    # file because we want consistent column ordering even when
    # some inputs have appended manual columns.
    native_cols = [
        'Item Code', 'Product Name', 'MRP', 'Rate', 'Qty', 'UOM',
        'Disc %', 'Amount', 'CGST %', 'SGST %', 'IGST %',
        'Gross Amount',
    ]
    engine_cols = ['EAN', 'Item No (Master)', 'Status']
    all_cols = ['Source File', 'PO No.'] + native_cols + engine_cols
    _write_header_row(ws, 1, all_cols)

    row = 2
    for r in results:
        # Re-open the source file. Slower than caching parsed data
        # on the result object but keeps the engine memory profile
        # flat — for raw-data echo we need cell-level access to the
        # original columns in their original positions, including
        # ones we don't normally read (Rate, UOM, Disc %, etc.).
        try:
            # Prefer the path stored on the result (set at read
            # time). Fall back to filename search for legacy results
            # created without source_path.
            src_path = r.source_path or _find_uploaded_file(r.source_file)
            if not src_path or not Path(src_path).exists():
                ws.cell(row, 1, r.source_file)
                ws.cell(row, 2, r.po_number)
                ws.cell(row, 3, '(source file not accessible for raw echo)')
                row += 1
                continue
            wb_src = load_workbook(str(src_path), data_only=True)
            ws_src = wb_src.active
        except Exception as e:
            ws.cell(row, 1, r.source_file)
            ws.cell(row, 2, r.po_number)
            ws.cell(row, 3, f'(failed to re-read for raw echo: {e})')
            row += 1
            continue

        # Map header label → src column index. PO files have data
        # headers on row 15.
        src_header_map: Dict[str, int] = {}
        for c in range(1, ws_src.max_column + 1):
            v = ws_src.cell(_DATA_HEADER_ROW, c).value
            if v is not None:
                src_header_map[str(v).strip().lower()] = c

        # Build a per-line lookup of engine resolutions keyed by
        # naturals_code so we can paint the resolved cols onto every
        # data row without re-running the lookups.
        resolved_by_code: Dict[int, 'POLine'] = {
            l.naturals_code: l for l in r.lines
        }

        for src_r in range(_DATA_START_ROW, ws_src.max_row + 1):
            code_v = ws_src.cell(src_r, src_header_map.get('item code', 1)).value
            if code_v is None or str(code_v).strip() == '':
                continue
            try:
                int(str(code_v).strip())
            except (ValueError, TypeError):
                # End of data table (footer/totals row).
                break

            # Source file + PO number columns (always present)
            ws.cell(row, 1, r.source_file)
            ws.cell(row, 2, r.po_number)
            # Native columns (echoed from the PO file as-is)
            for i, col_name in enumerate(native_cols, start=3):
                src_col = src_header_map.get(col_name.lower())
                if src_col is not None:
                    ws.cell(row, i, ws_src.cell(src_r, src_col).value)
            # Engine-appended columns (resolution outputs)
            try:
                code_int = int(str(code_v).strip())
            except (ValueError, TypeError):
                code_int = None
            line = resolved_by_code.get(code_int) if code_int else None
            if line is not None:
                ws.cell(row, len(all_cols) - 2, line.ean)
                ws.cell(row, len(all_cols) - 1, line.item_no_resolved)
                ws.cell(row, len(all_cols), line.status)
                # Highlight rows that didn't resolve cleanly so
                # they're easy to spot in the bulk audit dump.
                if line.status != 'OK':
                    fill = (_ERR_FILL
                            if line.status == 'NATURALS_CODE_UNMAPPED'
                            else _WARN_FILL)
                    for c in range(1, len(all_cols) + 1):
                        ws.cell(row, c).fill = fill
            row += 1

    _auto_size(ws, max_width=50)
    ws.freeze_panes = 'C2'  # Pin Source File + PO No. while scrolling
                              # horizontally — they're the row identity.


def _find_uploaded_file(filename: str) -> Optional[Path]:
    """Best-effort resolution of a bare filename to its path.

    The PO file paths in :class:`POResult` are stored as
    ``Path.name`` only (no directory) so the result is portable.
    For Raw Data echo we need to re-open the file, so we search
    candidate locations in order: cwd, APP_DIR, and any path
    matching the bare name.
    """
    bare = Path(filename).name
    for parent in (Path.cwd(), APP_DIR, APP_DIR / 'input'):
        candidate = parent / bare
        if candidate.exists():
            return candidate
    return None


def _write_header_row(ws, row: int, labels: List[str]) -> None:
    """Standard blue-band header row used on every sheet.

    Centralised so changing the header style is a one-spot edit
    rather than five identical fill/font/alignment blocks across
    the writers above.
    """
    for c, label in enumerate(labels, start=1):
        cell = ws.cell(row, c, value=label)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal='center',
                                    vertical='center',
                                    wrap_text=True)
    ws.row_dimensions[row].height = 22


def _auto_size(ws, max_width: int = 60) -> None:
    """Set column widths based on observed content lengths.

    Caps at ``max_width`` chars to prevent runaway column widths
    from long description fields. Uses approximate
    ``len(str(value))`` rather than real font-metric measurement —
    good enough for an audit workbook, fast enough to run on every
    export.
    """
    from openpyxl.utils import get_column_letter
    for c in range(1, ws.max_column + 1):
        max_len = 8
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            ln = len(str(v))
            if ln > max_len:
                max_len = ln
        ws.column_dimensions[get_column_letter(c)].width = min(
            max_len + 2, max_width)


# ════════════════════════════════════════════════════════════════════
# Orchestrator
# ════════════════════════════════════════════════════════════════════

def discover_customers() -> List[str]:
    """Scan customers/ for subfolders containing a *_Master.xlsx."""
    if not CUSTOMERS_DIR.is_dir():
        return []
    out = []
    for child in sorted(CUSTOMERS_DIR.iterdir()):
        if not child.is_dir():
            continue
        # Master file must match <FolderName>_Master.xlsx
        master = child / f'{child.name}_Master.xlsx'
        if master.is_file():
            out.append(child.name)
    return out


def run_batch(customer: str,
              po_paths: List[Path],
              starting_so: str,
              output_path: Path,
              margin_pct: Optional[float] = None,
              status_cb=None) -> Tuple[Path, List[POResult]]:
    """End-to-end runner.

    `status_cb(msg)` is called periodically with progress strings —
    GUI wires this to the status bar.
    """
    def _say(msg: str) -> None:
        log.info(msg)
        if status_cb:
            status_cb(msg)

    # Load all masters
    _say('Loading customer master…')
    load_logs: List[str] = []
    cm = CustomerMasterLoader()
    cm.load(CUSTOMERS_DIR / customer / f'{customer}_Master.xlsx',
             customer, load_logs)
    _say('Loading Items March…')
    im = ItemsMarchLoader()
    im.load(ITEMS_MARCH_PATH, load_logs)
    _say('Loading ship-to dump…')
    st = ShipToLoader()
    st.load(SHIP_TO_PATH, load_logs)

    # Process each PO file
    results: List[POResult] = []
    so_number = starting_so
    for i, path in enumerate(po_paths, start=1):
        _say(f'Processing [{i}/{len(po_paths)}]: {path.name}')
        res = read_po_file(path, cm, im)
        res.ship_to = resolve_ship_to(path.name, st, CITY_ALIASES)
        res.so_number = so_number
        results.append(res)
        so_number = increment_so_number(so_number)

    # Surface loader warnings on the first result for visibility
    if load_logs and results:
        results[0].warnings = load_logs + results[0].warnings

    _say('Writing output workbook…')
    export_workbook(results, output_path, customer, margin_pct)
    return output_path, results


# ════════════════════════════════════════════════════════════════════
# Tkinter GUI
# ════════════════════════════════════════════════════════════════════

# Conditional import so the engine above runs headlessly in
# environments without tkinter (some CI sandboxes). The frozen
# .exe always has it.
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    _TK_AVAILABLE = True
except ImportError:
    _TK_AVAILABLE = False
    tk = None  # type: ignore[assignment]


class App(getattr(tk, 'Tk', object) if _TK_AVAILABLE else object):
    """Main window. Single-screen workflow."""

    def __init__(self) -> None:
        super().__init__()
        self.title(f'MT Select Constructor v{__version__}')
        self.geometry('780x620')

        self.po_paths: List[Path] = []
        self._build_ui()
        self._refresh_customers()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill=tk.BOTH, expand=True)

        # ── Customer dropdown ──────────────────────────────────────
        row = ttk.Frame(outer)
        row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(row, text='Customer:', width=18).pack(side=tk.LEFT)
        self.customer_var = tk.StringVar()
        self.customer_combo = ttk.Combobox(
            row, textvariable=self.customer_var,
            state='readonly', width=30,
        )
        self.customer_combo.pack(side=tk.LEFT)
        ttk.Button(row, text='Refresh',
                    command=self._refresh_customers).pack(side=tk.LEFT, padx=8)

        # ── PO file picker ─────────────────────────────────────────
        row = ttk.Frame(outer)
        row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(row, text='PO Files:', width=18).pack(side=tk.LEFT)
        ttk.Button(row, text='Add Files…',
                    command=self._add_files).pack(side=tk.LEFT)
        ttk.Button(row, text='Remove Selected',
                    command=self._remove_selected).pack(side=tk.LEFT, padx=4)
        ttk.Button(row, text='Clear All',
                    command=self._clear_files).pack(side=tk.LEFT)

        list_frame = ttk.Frame(outer)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 8))
        self.file_list = tk.Listbox(list_frame, selectmode=tk.EXTENDED)
        self.file_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(list_frame, command=self.file_list.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_list.configure(yscrollcommand=sb.set)

        # ── Starting SO ────────────────────────────────────────────
        row = ttk.Frame(outer)
        row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(row, text='Starting SO Number:',
                   width=18).pack(side=tk.LEFT)
        self.start_so_var = tk.StringVar(value=suggest_starting_so_number())
        ttk.Entry(row, textvariable=self.start_so_var,
                   width=30).pack(side=tk.LEFT)
        ttk.Label(row,
                   text='Format: SO/NS/MM/DDMYY. Trailing number increments per PO.',
                   foreground='#666').pack(side=tk.LEFT, padx=8)

        # ── Generate ───────────────────────────────────────────────
        row = ttk.Frame(outer)
        row.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(row, text='Generate Sales Order Workbook',
                    command=self._generate).pack(side=tk.LEFT)
        ttk.Button(row, text='Open Output Folder',
                    command=self._open_output).pack(side=tk.LEFT, padx=8)

        # ── Status log ─────────────────────────────────────────────
        ttk.Label(outer, text='Status:').pack(anchor='w', pady=(8, 2))
        self.status = scrolledtext.ScrolledText(outer, height=10,
                                                 state=tk.DISABLED,
                                                 font=('Consolas', 9))
        self.status.pack(fill=tk.BOTH, expand=True)

    # ── Customer dropdown ──────────────────────────────────────────
    def _refresh_customers(self) -> None:
        customers = discover_customers()
        self.customer_combo['values'] = customers
        if customers and not self.customer_var.get():
            self.customer_var.set(customers[0])
        if not customers:
            self._say(f"No customers found under {CUSTOMERS_DIR}. "
                       f"Add a folder like 'customers/Naturals/Naturals_Master.xlsx'.")

    # ── File list ──────────────────────────────────────────────────
    def _add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title='Select PO files',
            filetypes=[('Excel files', '*.xlsx *.xlsm'),
                       ('All files', '*.*')],
        )
        for p in paths:
            pp = Path(p)
            if pp not in self.po_paths:
                self.po_paths.append(pp)
                self.file_list.insert(tk.END, pp.name)

    def _remove_selected(self) -> None:
        for i in reversed(self.file_list.curselection()):
            del self.po_paths[i]
            self.file_list.delete(i)

    def _clear_files(self) -> None:
        self.po_paths.clear()
        self.file_list.delete(0, tk.END)

    # ── Status output ──────────────────────────────────────────────
    def _say(self, msg: str) -> None:
        self.status.configure(state=tk.NORMAL)
        ts = dt.datetime.now().strftime('%H:%M:%S')
        self.status.insert(tk.END, f'[{ts}] {msg}\n')
        self.status.see(tk.END)
        self.status.configure(state=tk.DISABLED)
        self.update_idletasks()

    # ── Generate ───────────────────────────────────────────────────
    def _generate(self) -> None:
        customer = self.customer_var.get()
        if not customer:
            messagebox.showerror('Missing customer',
                                  'Select a customer first.')
            return
        if not self.po_paths:
            messagebox.showerror('No files',
                                  'Add at least one PO file.')
            return
        starting_so = self.start_so_var.get().strip()
        if not starting_so:
            messagebox.showerror('Missing SO number',
                                  'Enter a starting SO number.')
            return
        try:
            increment_so_number(starting_so)
        except ValueError as e:
            messagebox.showerror('Bad SO number', str(e))
            return

        ts = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
        out_path = OUTPUT_DIR / f'MT_Select_Output_{customer}_{ts}.xlsx'

        try:
            self._say(f'Starting batch: customer={customer}, '
                       f'{len(self.po_paths)} file(s), SO start={starting_so}')
            out_path, results = run_batch(
                customer=customer,
                po_paths=self.po_paths,
                starting_so=starting_so,
                output_path=out_path,
                margin_pct=None,  # Unit Price stays blank
                status_cb=self._say,
            )
            # Final summary
            total_lines = sum(len(r.lines) for r in results)
            unmatched_st = sum(1 for r in results
                                if r.ship_to.method.startswith('UNMATCHED'))
            unmapped_items = sum(
                1 for r in results for l in r.lines
                if l.status != 'OK'
            )
            self._say(
                f'Done. {len(results)} PO(s), {total_lines} line(s), '
                f'{unmatched_st} unmatched ship-to(s), '
                f'{unmapped_items} unmapped item(s).')
            self._say(f'Output: {out_path}')
            messagebox.showinfo(
                'Done',
                f'Generated {out_path.name}\n\n'
                f'{len(results)} PO(s), {total_lines} line(s)\n'
                f'Ship-to unmatched: {unmatched_st}\n'
                f'Items unmapped: {unmapped_items}\n\n'
                f'Check Warnings sheet for details.')
        except Exception as e:
            log.exception('Batch failed')
            self._say(f'ERROR: {e}')
            messagebox.showerror('Batch failed', str(e))

    def _open_output(self) -> None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith('win'):
                os.startfile(str(OUTPUT_DIR))  # type: ignore[attr-defined]
            elif sys.platform == 'darwin':
                os.system(f'open "{OUTPUT_DIR}"')
            else:
                os.system(f'xdg-open "{OUTPUT_DIR}"')
        except Exception as e:
            messagebox.showerror('Could not open folder', str(e))


# ════════════════════════════════════════════════════════════════════
# Entrypoint
# ════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    if not _TK_AVAILABLE:
        print('ERROR: tkinter is not installed in this Python environment. '
              'On Windows the bundled .exe includes it; on Linux you may '
              'need: apt install python3-tk')
        sys.exit(1)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    App().mainloop()