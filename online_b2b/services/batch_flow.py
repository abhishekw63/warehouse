"""online_b2b.services.batch_flow — Batch Run **Phase 0: READ-ONLY detector**.

Given a set of uploaded PO files, guess which marketplace each belongs to by
matching its header columns against the engine's *own* ``MARKETPLACE_CONFIGS``
column map (``po_col``/``loc_col``/``qty_col``/``ean_col``/``fob_col``/…) plus the
declared ``accepted_extensions``. It **reads headers only — records nothing,
processes nothing, touches no DB**. Its sole job is to let the operator *confirm*
file→MP before any batch processing is ever built on top (the make-or-break gate).

Purely additive + removable per the Batch Run safety contract: deleting this file
changes nothing else in the app. Scope = the online MPs in ``MARKETPLACE_CONFIGS``
(the daily-grind channels); offline MT/GT bridges are out of scope for Phase 0.
See [[project-backlog]] (Batch Run) and [[dry-skeleton-first]].
"""
from __future__ import annotations

import csv
import re
from pathlib import Path

# Config keys whose values name real header columns. Values may be a str, a list
# of alternatives, or a {'multiply': [...]} spec — all flattened to column names.
_COL_KEYS = ('po_col', 'loc_col', 'qty_col', 'ean_col', 'price_col',
             'fob_col', 'ref_fob_col', 'amount_col')


def _norm(s) -> str:
    """Fold a header/label to a comparable token: lowercase, alnum only."""
    return re.sub(r'[^a-z0-9]', '', str(s).lower())


def _flatten_cols(val) -> list:
    """Column names from a config value (str | list | {'multiply': [...]}).
    Skips engine placeholders like ``__po__`` (PDF-parser-supplied, not headers)."""
    out: list = []
    if val is None:
        return out
    if isinstance(val, str):
        if not val.startswith('__'):
            out.append(val)
    elif isinstance(val, (list, tuple)):
        for v in val:
            out += _flatten_cols(v)
    elif isinstance(val, dict):
        for v in val.values():
            out += _flatten_cols(v)
    return out


def signatures() -> dict:
    """``{mp: {'cols': set(normalized expected headers), 'exts': set, 'pdf': bool}}``
    built live from the engine's MARKETPLACE_CONFIGS — the single source of truth."""
    from .engine_bridge import _engine_imports
    cfgs = _engine_imports()['MARKETPLACE_CONFIGS']
    sigs: dict = {}
    for name, c in cfgs.items():
        cols = set()
        for k in _COL_KEYS:
            for col in _flatten_cols(c.get(k)):
                cols.add(_norm(col))
        cols.discard('')
        exts = {str(e).lower() for e in (c.get('accepted_extensions') or [])}
        sigs[name] = {'cols': cols, 'exts': exts, 'pdf': bool(c.get('pdf_parser'))}
    return sigs


# Header rows can sit deep: Flipkart has an address/payment block above its product
# table, Myntra a title row. We UNION the string tokens across the header region of
# every sheet so a deep/offset header is always captured — noise data values rarely
# collide with another MP's specific header names, so union is safe for matching.
_HEADER_SCAN_ROWS = 25


def _read_headers(path) -> tuple:
    """(ext, set(normalized header tokens)) — best-effort, READ-ONLY across
    .xlsx/.xlsm (openpyxl), .xls (xlrd), .xlsb (pyxlsb) and .csv. Never raises: on
    any read problem returns an empty token set so detection falls back to the
    extension/parser hint (→ low confidence, operator confirms)."""
    p = Path(path)
    ext = p.suffix.lower()
    toks: set = set()
    try:
        if ext in ('.xlsx', '.xlsm'):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            for ws in wb.worksheets:
                for row in ws.iter_rows(min_row=1, max_row=_HEADER_SCAN_ROWS,
                                        values_only=True):
                    toks |= {_norm(v) for v in row
                             if isinstance(v, str) and str(v).strip()}
            wb.close()
        elif ext == '.xls':
            import xlrd
            book = xlrd.open_workbook(path)
            for sh in book.sheets():
                for ri in range(min(_HEADER_SCAN_ROWS, sh.nrows)):
                    for ci in range(sh.ncols):
                        v = sh.cell_value(ri, ci)
                        if isinstance(v, str) and v.strip():
                            toks.add(_norm(v))
        elif ext == '.xlsb':
            from pyxlsb import open_workbook
            with open_workbook(path) as wb:
                for name in wb.sheets:
                    with wb.get_sheet(name) as sh:
                        for ri, row in enumerate(sh.rows()):
                            if ri >= _HEADER_SCAN_ROWS:
                                break
                            toks |= {_norm(c.v) for c in row
                                     if isinstance(c.v, str) and str(c.v).strip()}
        elif ext == '.csv':
            with open(path, newline='', encoding='utf-8', errors='ignore') as f:
                for i, row in enumerate(csv.reader(f)):
                    if i >= _HEADER_SCAN_ROWS:
                        break
                    toks |= {_norm(v) for v in row if str(v).strip()}
    except Exception:  # noqa: BLE001 — detection must never break on a bad file
        pass
    toks.discard('')
    return ext, toks


# Distinctive filename patterns (a second, independent signal). Portal exports keep
# stable prefixes, so these disambiguate MPs whose COLUMNS overlap — notably
# Swiggy (PO_<digits>) vs Zepto (PO_<hex>), which share an identical CSV schema.
# Matched against the lowercased basename; regex order = priority.
_FILENAME_HINTS = [
    ('RK', r'poitemexport'),
    ('Nykaa', r'purchase\s*orders?[_\s]*allpo|allpo'),
    ('Flipkart', r'purchase_order_fls'),
    ('Flipkart-TO', r'consignment'),
    ('Myntra', r'mynj'),
    ('Purplle', r'execl_attached'),
    ('Blink', r'bulk_po_csv'),
    ('Swiggy', r'^po_\d+\.(csv|xlsx?)$'),
    ('Zepto', r'^po_[0-9a-f]*[a-f][0-9a-f]*\.(csv|xlsx?)$'),
]


def _filename_hint(name: str):
    low = name.lower()
    for mp, pat in _FILENAME_HINTS:
        if re.search(pat, low):
            return mp
    return None


def detect_one(path) -> dict:
    """Detect the marketplace for ONE file from its header columns (engine
    signatures) AND its filename pattern. Returns a JSON-ready dict:
    ``{file, ext, guess, confidence, matched, expected, evidence, filename_hint,
    alternatives}``. ``confidence`` ∈ high | medium | low | unknown. Records
    nothing — the operator confirms every row before any processing exists."""
    name = Path(path).name
    ext, toks = _read_headers(path)
    sigs = signatures()
    fn = _filename_hint(name)

    scored = []
    for mp, sig in sigs.items():
        ext_ok = (not sig['exts']) or (not ext) or (ext in sig['exts'])
        matched = sorted(sig['cols'] & toks)
        scored.append({
            'mp': mp, 'cols': len(matched), 'ext_ok': ext_ok,
            'matched': matched, 'expected': len(sig['cols']),
            'fn': (mp == fn),
        })
    # Combined rank score: column matches + a filename bonus. Extension-compatible
    # candidates always outrank incompatible ones.
    for s in scored:
        s['score'] = s['cols'] + (3 if s['fn'] else 0)
    scored.sort(key=lambda s: (s['ext_ok'], s['score']), reverse=True)
    top = scored[0]
    second = scored[1] if len(scored) > 1 else {'score': 0}
    margin = top['score'] - second['score']
    ratio = top['cols'] / top['expected'] if top['expected'] else 0

    # Confidence — highest when filename AND columns agree; filename-only or
    # strong-columns-only is medium; anything weak/ambiguous → low/unknown.
    if top['score'] == 0:
        conf, guess = 'unknown', None
    elif not top['ext_ok']:
        conf, guess = 'low', top['mp']
    elif top['fn'] and top['cols'] >= 2:
        conf, guess = 'high', top['mp']
    elif top['cols'] >= 3 and margin >= 2 and ratio >= 0.5:
        conf, guess = 'high', top['mp']
    elif top['fn'] or (top['cols'] >= 2 and margin >= 1):
        conf, guess = 'medium', top['mp']
    else:
        conf, guess = 'low', top['mp']

    alts = [{'mp': s['mp'], 'cols': s['cols'], 'fn': s['fn']}
            for s in scored[1:4] if s['score'] > 0]
    return {
        'file': name, 'ext': ext, 'guess': guess, 'confidence': conf,
        'matched': top['cols'], 'expected': top['expected'],
        'evidence': top['matched'], 'filename_hint': fn, 'alternatives': alts,
    }


def detect(paths) -> list:
    """Detect the MP for each file. Pure/read-only — the operator confirms every
    row before anything is ever processed. Returns one dict per file."""
    return [detect_one(p) for p in paths]
