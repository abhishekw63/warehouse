"""
online_b2b.services.reliance_shipto
===================================

Reliance (**online**, cust 20015) ship-to correction — in OUR layer, the frozen
``reliance_pdf_parser`` untouched.

The parser reduces the PDF's Delivery Address to just the **city**
(``__loc__ = delivery_city``), so two Gurgaon-district DCs collide: the mapping's
substring tier matches ``'GURGAON'`` inside ``'Reliance Retail Limited-Gurgaon'``
(20015_5, MG Road, pin 122001) and never reaches ``'FARUKHNAGAR (Reliance)'``
(20015_6, pin 122506).

Fix: when the PDF's delivery **pincode** identifies a DC the city alone can't,
override ``__loc__`` with the correct del-location token so the engine maps the
right ``ship_to`` for the D365 output, the recorded row AND the tracker Location.

SCOPED to Reliance online only, and only to the deliberate pincode splits below —
every other Reliance PO (and every other channel) keeps the engine's default
city resolution untouched. 122001 (Gurgaon MG Road) already resolves correctly to
20015_5, so it is intentionally NOT overridden.
"""
from __future__ import annotations

# Delivery pincode → the ``__loc__`` token that resolves to the intended ship-to.
#   '122506' → 'FARUKHNAGAR'  (substring-matches del_location
#              'FARUKHNAGAR (Reliance)' = 20015_6)
# Add a row here only for a real, confirmed pincode→DC split.
PIN_LOC_OVERRIDE: dict[str, str] = {
    '122506': 'FARUKHNAGAR',
}


def loc_override_for_pin(pin) -> str | None:
    """The ``__loc__`` override for a delivery pincode, or ``None`` to keep the
    engine's default city resolution. Only known, deliberate splits act — never
    a blanket override."""
    return PIN_LOC_OVERRIDE.get(str(pin or '').strip())
