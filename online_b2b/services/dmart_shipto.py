"""
online_b2b.services.dmart_shipto
================================

Robust DMart (Avenue) ship-to (FC) resolution — in OUR layer, never touching the
frozen ``avenue_pdf_parser`` (which hardcodes ``'Bhiwandi'`` as the ship-to and
silently defaults EVERY other FC to it — see ``avenue_po_to_dataframe`` line
``'__loc__': po.header.ship_to_text or 'Bhiwandi'``).

We read the PDF's *Ship To* block ourselves, match the FC against the DMart
``ship_to_mapping`` rows by exact (normalised) del-location name, and cross-check
the pincode when the mapping carries a ``postcode``. The confirmed FC name is
handed back so the frozen engine maps the correct ``ship_to`` code for BOTH the
D365 output and the recorded row.

never-skip-silently: an FC that is not found, is ambiguous, or fails the pincode
cross-check is REPORTED (the caller blocks the run) — it is never defaulted.
"""
from __future__ import annotations

import re


def _norm(s: str) -> str:
    """Lower-case, strip everything but a-z0-9 — tolerant to pdfplumber's
    variable spacing (e.g. 'IFC- Kukse Bhiwandi' / 'IFC-KukseBhiwandi')."""
    return re.sub(r'[^a-z0-9]', '', (s or '').lower())


# Generic words that DON'T identify a specific FC — stripped before matching so
# the distinctive place token remains (e.g. 'Wagholi FC'→'wagholi', 'Pedda
# Amberpet FC'→'amberpet'). DMart PDFs prefix 'IFC-' and vary spellings/spacing.
_GENERIC = {'fc', 'ifc', 'pune', 'mumbai', 'chennai', 'the', 'and',
            'west', 'east', 'new', 'nagar', 'marg'}


def _fc_tokens(name: str) -> list:
    """Distinctive place tokens for an FC name (≥5 chars, non-generic). Used to
    match DMart's messy PDF ship-to text against the clean mapping name."""
    toks = re.findall(r'[a-z]+', (name or '').lower())
    core = [t for t in toks if t not in _GENERIC and len(t) >= 5]
    return core or [t for t in toks if t not in _GENERIC] or toks


def dmart_fc_directory() -> dict:
    """``{normalised del_location: {'fc', 'ship_to', 'postcode'}}`` for party
    'Dmart'. The engine keys DMart on ``del_location`` (e.g. 'Hennur FC')."""
    from .order_db import _conn
    out: dict = {}
    with _conn() as (cur, _d):
        cur.execute("SELECT del_location, ship_to, postcode FROM "
                    "ship_to_mapping WHERE party='Dmart'")
        for dl, st, pc in cur.fetchall():
            k = _norm(dl)
            if k:
                out[k] = {'fc': dl, 'ship_to': st, 'postcode': (pc or '').strip(),
                          'tokens': _fc_tokens(dl)}
    return out


def _pdf_header_text(path) -> str:
    """First-page text — the DMart header (PO no, Ship To, pincode) lives here."""
    import pdfplumber
    with pdfplumber.open(path) as p:
        pages = p.pages[:1] or p.pages
        return '\n'.join((pg.extract_text() or '') for pg in pages)


def _shipto_pincode(text: str) -> str:
    """The Ship-To pincode. DMart's delivery block always ends with the FC's
    ``EmailID:...@dmartindia.com`` — the ship-to pincode is the LAST 6-digit
    before that anchor (dodging the Corp-Off & vendor pincodes that sit earlier
    in the column-glued header). Falls back to '' if the anchor is absent."""
    low = text.lower()
    anchor = low.find('dmartindia.com')
    scope = text[:anchor] if anchor > 0 else ''
    pins = re.findall(r'(?<!\d)(\d{6})(?!\d)', scope)
    return pins[-1] if pins else ''


def resolve_pdf(path) -> dict:
    """Resolve one DMart PDF → ``{ok, po, fc, ship_to, pincode, reason}``.
    ``ok=False`` (with a human reason) whenever the FC can't be CONFIRMED."""
    blank = {'po': '', 'fc': '', 'ship_to': '', 'pincode': ''}
    try:
        text = _pdf_header_text(path)
    except Exception as e:  # noqa: BLE001
        return {**blank, 'ok': False, 'reason': f"PDF unreadable: {type(e).__name__}"}

    po_m = (re.search(r'Purchase\s*Order\s+(\d{6,})', text, re.I)
            or re.search(r'\b(45\d{8,})\b', text))
    po = po_m.group(1) if po_m else ''

    ntext = _norm(text)
    pincode = _shipto_pincode(text)
    # short human-readable ship-to hint for the block message
    snippet = (text[text.lower().find('ship'):][:110] if 'ship' in text.lower()
               else text[:110]).replace('\n', ' ').strip()

    directory = dmart_fc_directory()
    # Match on the FC's distinctive place token(s) — tolerant of DMart's PDF
    # variants ('IFC-Wagholi Pune'↔'Wagholi FC', 'IFC Pedaa Amberpet'↔'Pedda
    # Amberpet FC'). DMart's Bill-To and Ship-To are the same FC in practice; if
    # two DIFFERENT FCs' tokens both appear we BLOCK for review rather than guess.
    hits = {}
    for v in directory.values():
        if any(tok in ntext for tok in v['tokens']):
            hits[v['ship_to']] = v
    hits = list(hits.values())

    if not hits:
        return {**blank, 'po': po, 'pincode': pincode, 'ok': False,
                'reason': (f"FC not found in DMart ship-to mapping. PDF ship-to "
                           f"reads: '{snippet[:90]}…'. Add the FC on the "
                           f"Ship-To Mapping page.")}
    if len(hits) > 1:
        names = ', '.join(h['fc'] for h in hits)
        return {**blank, 'po': po, 'pincode': pincode, 'ok': False,
                'reason': f"Ambiguous ship-to — multiple FCs match ({names}); "
                          f"confirm manually."}

    fc = hits[0]
    if fc['postcode'] and pincode and fc['postcode'] != pincode:
        return {'po': po, 'fc': fc['fc'], 'ship_to': fc['ship_to'],
                'pincode': pincode, 'ok': False,
                'reason': (f"Pincode mismatch for {fc['fc']}: PDF {pincode} vs "
                           f"mapping {fc['postcode']}. Confirm the ship-to.")}
    return {'ok': True, 'po': po, 'fc': fc['fc'], 'ship_to': fc['ship_to'],
            'pincode': pincode, 'reason': ''}


def resolve_paths(paths) -> dict:
    """``{po: resolve_pdf(...)}`` across all .pdf paths (non-PDFs ignored)."""
    out: dict = {}
    for p in paths or []:
        if str(p).lower().endswith('.pdf'):
            r = resolve_pdf(p)
            out[r.get('po') or str(p)] = r
    return out
