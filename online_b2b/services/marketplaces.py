"""
online_b2b.services.marketplaces
================================

**Marketplace / channel registry** — the single source of truth for the list of
channels the team works, across Online B2B and Offline. First consumer is the
Daily Activity Checklist; the hub chips, upload dropdown and See-full-template
will migrate to read from here too (skeleton/DRY refactor).

Each :class:`Channel` carries:
  * ``key``     — stable slug (used in the DB / URLs, never shown)
  * ``display`` — operator-facing label (matches their manual "Channels" sheet)
  * ``segment`` — 'Online' | 'Offline'
  * ``db_key``  — the ``order_headers.marketplace`` value used to AUTO-detect
                  "Uploaded (web)" for the day ('' = not web-integrated → manual)
  * ``live``    — web-integrated today (vs a coming-soon / manual-only channel)

API-ready: the ``*_dicts`` helpers return plain JSON-serializable structures.
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Channel:
    key: str
    display: str
    segment: str
    db_key: str = ''
    live: bool = True


# Order here = display order on the checklist. Labels mirror the operator's
# "Channels" sheet; db_key mirrors engine_bridge.PILOT_MARKETPLACES values.
CHANNELS: list[Channel] = [
    # ── Online ──
    Channel('blink', 'Blinkit', 'Online', 'Blink'),
    Channel('flipkart', 'Flipkart Alpha & Hyperlocal', 'Online', 'Flipkart'),
    Channel('flipkart_to', 'Flipkart Branch', 'Online', 'Flipkart-TO'),
    Channel('rk', 'RK', 'Online', 'RK'),
    Channel('zepto', 'Zepto', 'Online', 'Zepto'),
    Channel('swiggy', 'Swiggy', 'Online', 'Swiggy'),
    Channel('nykaa', 'Nykaa', 'Online', 'Nykaa'),
    Channel('myntra', 'Myntra', 'Online', 'Myntra'),
    Channel('purplle', 'Purplle', 'Online', 'Purplle'),
    Channel('reliance', 'Reliance', 'Online', 'Reliance'),
    Channel('bigbasket', 'Big Basket', 'Online', 'Bigbasket'),
    Channel('firstcry', 'First Cry', 'Online', 'Firstcry'),
    Channel('dmart', 'D Mart', 'Online', 'Dmart'),
    Channel('meesho', 'Meesho-SB', 'Online', 'Meesho-TO'),
    Channel('blinkmp', 'BlinkMP', 'Online', '', False),
    Channel('smytten', 'Smytten', 'Online', '', False),
    # ── Offline ──
    Channel('gt_mass', 'GT Mass', 'Offline', 'GT Mass'),
    Channel('gt_select', 'GT Select', 'Offline', 'GT Select'),
    Channel('mt_select', 'MT Select', 'Offline', ''),
    Channel('ebo_kiosk', 'EBO/Kiosk', 'Offline', '', False),
    Channel('airport', 'Airport', 'Offline', '', False),
    Channel('off_inst', 'OFF-INSTITUTIONAL', 'Offline', '', False),
    Channel('eka', 'EKA', 'Offline', '', False),
    Channel('csd', 'CSD', 'Offline', '', False),
]

_BY_KEY = {c.key: c for c in CHANNELS}
_SEGMENTS = ('Online', 'Offline')


def channels() -> list[Channel]:
    return CHANNELS


def get(key: str) -> Channel | None:
    return _BY_KEY.get(key)


def db_key_to_channel() -> dict:
    """``{order_headers.marketplace value: channel.key}`` for web auto-detect."""
    return {c.db_key: c.key for c in CHANNELS if c.db_key}


def as_dicts() -> list[dict]:
    """Flat JSON-safe list."""
    return [c.__dict__.copy() for c in CHANNELS]


def grouped() -> list[dict]:
    """``[{segment, channels:[dict,...]}, ...]`` in display order — API-ready."""
    return [{'segment': seg,
             'channels': [c.__dict__.copy() for c in CHANNELS if c.segment == seg]}
            for seg in _SEGMENTS]
