"""
online_b2b.services.reliance_shipto
===================================

Reliance (**online**, cust 20015) ship-to correction + guard — in OUR layer, the
frozen ``reliance_pdf_parser`` untouched.

The parser reduces the PDF's Delivery Address to just the **city**
(``__loc__ = delivery_city``), so two Gurgaon-district DCs collide: the mapping's
substring tier matches ``'GURGAON'`` inside ``'Reliance Retail Limited-Gurgaon'``
(20015_5, MG Road, pin 122001) and never reaches ``'FARUKHNAGAR (Reliance)'``
(20015_6, pin 122506). That silent wrong-DC default is the exact class of the
DMart→Bhiwandi bug ([[dmart-shipto-fix]]).

Two parts, both SCOPED to Reliance online and both driven off :data:`PIN_DC`:

1. **Route (override)** — when the PDF's delivery **pincode** identifies a DC the
   city alone can't, override ``__loc__`` with the correct del-location token so
   the engine maps the right ``ship_to`` for the D365 output, the recorded row AND
   the tracker Location. 122001 (Gurgaon MG Road) already resolves correctly to
   20015_5, so its ``loc`` is ``None`` (no override needed) — it stays in the map
   only so :func:`confirm` recognises it as a known DC.

2. **Guard (block-on-ambiguous)** — a PO in a **multi-DC city** whose delivery
   pincode is NOT a known DC would fall back to the city-only guess. :func:`confirm`
   flags it so :class:`RelianceProcessor` BLOCKS the run (never routes to a default
   DC) — [[never-skip-silently]]. Single-DC cities and every mapped pincode pass
   untouched, so no other Reliance PO (and no other channel) changes behaviour.
"""
from __future__ import annotations


def _pin(pin) -> str:
    """Normalise a pincode to a bare 6-digit string ('122506' / 122506 / ' 122506 ')."""
    s = str(pin if pin is not None else '').strip()
    if s.endswith('.0'):          # numeric cells arrive as '122506.0'
        s = s[:-2]
    return s


# Authoritative Reliance DC directory for cities that host >1 DC (city alone can't
# decide → the delivery PINCODE must). Add a row here for every confirmed
# pincode→DC split. ``loc`` = token to force ``__loc__`` (None keeps the engine's
# city resolution when that already lands on the right ship_to).
PIN_DC: dict[str, dict] = {
    '122001': {'loc': None,          'ship_to': '20015_5', 'name': 'Gurgaon MG Road (Centro)'},
    '122506': {'loc': 'FARUKHNAGAR', 'ship_to': '20015_6', 'name': 'Farukhnagar Beauty DC'},
}

# Cities known to host >1 Reliance DC → a PO here MUST carry a pincode that is in
# PIN_DC, else its ship-to is only a city-guess and is BLOCKED for confirmation.
# Add a city here ONLY once ALL its DC pincodes are in PIN_DC — otherwise a
# legitimate PO to that city would block. (Known next candidate: Nagpur has two
# DCs — 20015_3 / 20015_4 — add it here with both pincodes when confirmed.)
MULTI_DC_CITIES: set[str] = {'gurgaon', 'gurugram'}

# Backward-compatible view: only the pincodes that actively override ``__loc__``.
PIN_LOC_OVERRIDE: dict[str, str] = {
    p: d['loc'] for p, d in PIN_DC.items() if d['loc']
}


def loc_override_for_pin(pin) -> str | None:
    """The ``__loc__`` override for a delivery pincode, or ``None`` to keep the
    engine's default city resolution. Only known, deliberate splits act — never a
    blanket override."""
    return PIN_LOC_OVERRIDE.get(_pin(pin))


def confirm(delivery_city, delivery_pin, site: str = '') -> tuple[bool, str]:
    """``(ok, reason)`` for one Reliance PO's ship-to.

    ``ok=False`` (with a human reason) ONLY for the silent wrong-DC class: a PO in
    a **multi-DC city** whose delivery pincode is not a known DC, so the engine
    would fall back to a city-only guess. Everything else — single-DC cities, and
    every pincode already in :data:`PIN_DC` — returns ``ok=True`` unchanged, so no
    other Reliance PO is affected."""
    city = str(delivery_city or '').strip().lower()
    pin = _pin(delivery_pin)
    if any(c in city for c in MULTI_DC_CITIES) and pin not in PIN_DC:
        return False, (
            f"Reliance ‘{delivery_city}’ hosts multiple DCs and the delivery "
            f"pincode {pin or '(missing)'} (site {site or '?'}) is not a known DC "
            f"— the ship-to would be a city-only guess (risk: wrong DC). Confirm "
            f"the DC / add its pincode→ship-to in reliance_shipto.PIN_DC before "
            f"upload."
        )
    return True, ''


# ── Reliance-family identifier reference ─────────────────────────────────────
# (full analysis: "Reliance Identifier Map & Verification" xlsx, 2026-07-23)
#
# Every Reliance PO carries a **Site code** — the unique per-DC identifier, like
# Blink's M10 / P2 / A1. Its NAMESPACE tells you online-vs-offline AND the customer:
#
#   Site pattern        Entity                  Cust    Channel   Input
#   S###  (SAE7,S4ZD)   Reliance                20015   ONLINE    PO PDF
#   T8##  (T8VY,T8WL)   Reliance Retail         20043   offline   PurchaseOrders xlsx (Site col)
#   FR## / 6220         Reliance Smart Bazaar   20615   offline   PurchaseOrders xlsx (DC_CODE)
#   TH## / TB##         Hamleys / RBL           20325   offline   PO PDF (Reliance Brands Ltd)
#   BAP / HO            Reliance Trends         20418   offline   BAP xlsx
#
# RULE: online == the Site code starts with 'S'. Anything else (T8/FR/TH/6/BAP) is
# OFFLINE and must NOT go through the online (cust 20015) uploader —
# offline_site_check() enforces this, data-driven from ship_to_mapping (not the
# letter) so it stays correct as codes are added.
#
# KNOWN GAP: the online (20015) ship_to rows store the CITY, not the Site code, so
# same-city DCs collide — 2 Gurgaon (20015_5 Centro/122001 vs 20015_6 Farukhnagar/
# 122506) and 2 Nagpur (20015_3 / 20015_4). confirm() blocks the ambiguous ones;
# capturing the Site code on the online rows would let online resolve by Site code
# like offline does, and retire the city-guess entirely.
# DATA ISSUE: FR49 is shared by Reliance Retail (20043_5) AND Smart Bazaar
# (20615_7), same Bangalore pincode — ambiguous; needs disambiguation.
# ─────────────────────────────────────────────────────────────────────────────
def _offline_site_owners() -> dict:
    """``{SITE_CODE: (party, cust_no)}`` for OFFLINE Reliance-family DCs whose
    ``del_location`` is a short site/DC code (T8VY, FR73, THK0, BAP …). Read live
    from ``ship_to_mapping`` so it stays current as DCs are added — no hardcoding.
    Empty on any DB hiccup (→ the wrong-channel guard simply doesn't act)."""
    import re
    try:
        from .order_db import _conn
    except Exception:  # noqa: BLE001
        return {}
    out: dict = {}
    try:
        with _conn() as (cur, _d):
            cur.execute(
                "SELECT party, cust_no, del_location FROM ship_to_mapping "
                "WHERE party IN ('Reliance Retail','Reliance Smart Bazaar',"
                "'Reliance Trends','Hamleys')")
            for party, cust, dl in cur.fetchall():
                code = str(dl or '').strip().upper()
                if re.fullmatch(r'[A-Z0-9]{2,6}', code):
                    out[code] = (party, cust)
    except Exception:  # noqa: BLE001
        return {}
    return out


def offline_site_check(site, owners: dict | None = None) -> tuple[bool, str]:
    """``(ok, reason)``. Blocks an ONLINE-channel Reliance PO whose **Site code**
    belongs to an OFFLINE Reliance customer — i.e. an offline PO punched into the
    online (cust 20015) uploader. Online DCs use ``S###`` site codes; the offline
    ones (T8··/FR··/TH··/BAP) are a different namespace, so a match here is a
    genuine wrong-channel upload, not an online DC."""
    s = str(site or '').strip().upper()
    if not s:
        return True, ''
    owners = _offline_site_owners() if owners is None else owners
    if s in owners:
        party, cust = owners[s]
        return False, (
            f"Site ‘{s}’ belongs to {party} (cust {cust}) — this is an OFFLINE "
            f"Reliance PO punched into the ONLINE (20015) channel. Process it via "
            f"the {party} offline channel, not here."
        )
    return True, ''


def confirm_paths(paths) -> dict:
    """``{po_number_or_path: (ok, reason)}`` across all Reliance PO **PDFs**.

    Two guards per PO, both block-don't-guess: (1) **wrong channel** — the Site
    code belongs to an offline Reliance customer; (2) **ambiguous DC** — a
    multi-DC city with an unmapped pincode. Read-only: parses each PDF with the
    frozen ``parse_reliance_pdf`` purely to read the header. Non-PDFs are ignored;
    a PDF that can't be parsed is left to the normal flow (``ok=True``) — the
    guards only ever ADD a block for these two specific cases."""
    out: dict = {}
    try:
        from online_po_processor.engine.reliance_pdf_parser import parse_reliance_pdf
    except Exception:  # noqa: BLE001 — no parser → nothing to guard
        return out
    owners = _offline_site_owners()
    for p in paths or []:
        if not str(p).lower().endswith('.pdf'):
            continue
        try:
            h = parse_reliance_pdf(p).header
        except Exception:  # noqa: BLE001 — defer to the normal flow, never block on this
            continue
        key = h.po_number or str(p)
        ok, reason = offline_site_check(h.site, owners)          # (1) wrong channel
        if ok:
            ok, reason = confirm(h.delivery_city, h.delivery_pin, h.site)  # (2) DC
        out[key] = (ok, reason)
    return out
