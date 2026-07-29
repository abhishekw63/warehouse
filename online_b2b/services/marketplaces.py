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
    steps: tuple = ()   # custom ((step_key, label), ...) for the Daily Tasks
    #                     checklist; empty = use the standard 5-step flow.
    parent: str = ''    # key of the parent channel this one nests under (e.g. an
    #                     MT-Select child under 'mt_select'); '' = top-level.
    db_label: str = ''  # ``order_headers.marketplace_label`` to match for web
    #                     auto-detect when ``db_key`` (= marketplace) is ambiguous.
    #                     All MT children share marketplace='MT', so they're told
    #                     apart by their recorded label (e.g. 'Health & Glow').


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
    # db_key='Blink RO': BlinkMP line items are recorded under marketplace
    # 'Blink RO' (the header shows 'BlinkMP'); without this the board can't fold
    # their qty/value → the row showed 0. Display stays 'BlinkMP'.
    Channel('blinkmp', 'BlinkMP', 'Online', 'Blink RO', False),
    Channel('smytten', 'Smytten', 'Online', '', False),
    # ── Offline ──
    Channel('gt_mass', 'GT Mass', 'Offline', 'GT Mass'),
    Channel('gt_select', 'GT Select', 'Offline', 'GT Select',
            steps=(('staging_am', 'Staging · morning'),
                   ('staging_pm', 'Staging · evening'),
                   ('sheet', 'Update in sheet'))),
    Channel('mt_select', 'MT Select', 'Offline', ''),
    # ── MT-Select children (nested under mt_select) — the web MT channels
    #    (offline.services.mt_bridge.WEB_CHANNELS). Each is a normal trackable
    #    channel with its own DB rows + the standard 5 steps; the mt_select parent
    #    is just an expandable container whose progress rolls these up. ──
    # db_label = the exact ``marketplace_label`` each records under (from
    # mt_bridge.channel_choices()), so web auto-detect can tell them apart even
    # though they all share marketplace='MT'.
    Channel('mt_ss', 'Shoppers Stop', 'Offline', '', True, parent='mt_select', db_label='Shoppers Stop'),
    Channel('mt_hg', 'H&G', 'Offline', '', True, parent='mt_select', db_label='Health & Glow'),
    Channel('mt_hb', 'Health & Beauty', 'Offline', '', True, parent='mt_select', db_label='Health & Beauty'),
    Channel('mt_nt', 'Naturals', 'Offline', '', True, parent='mt_select', db_label='Naturals'),
    # BN = Apollo's "Beauté Netrue" beauty retail arm (Cust 20735, Apollo
    # Pharmaproducts Pvt Ltd) — Apollo, BN and Beauté Netrue are the same party.
    Channel('mt_bn', 'Apollo (Beauté Netrue)', 'Offline', '', True, parent='mt_select', db_label='Apollo (Beaute Netrue)'),
    Channel('mt_ll', 'Lulu', 'Offline', '', True, parent='mt_select', db_label='Lulu Hypermarket'),
    Channel('mt_rl', 'Reliance Retail (Centro)', 'Offline', '', True, parent='mt_select', db_label='Reliance Retail (Centro)'),
    Channel('mt_met', 'Metro Cash & Carry', 'Offline', '', True, parent='mt_select', db_label='Metro Cash & Carry'),
    Channel('mt_ls', 'Lifestyle', 'Offline', '', True, parent='mt_select', db_label='Lifestyle'),
    # Manash = Purplle offline (daily-task tracking only for now).
    Channel('mt_manash', 'Manash', 'Offline', '', True, parent='mt_select', db_label='Manash (Purplle offline)'),
    # Reliance Smart Bazaar = Reliance hypermarket (cust 20615), separate from Centro.
    Channel('mt_rsb', 'Reliance Smart Bazaar', 'Offline', '', True, parent='mt_select', db_label='Reliance Smart Bazaar'),
    # Hamleys = Reliance BRANDS Limited (cust 20325), a separate entity from
    # Centro/Smart Bazaar; same PurchaseOrders*.xlsx / DC_CODE format as RSB.
    Channel('mt_rbl', 'Reliance Brands Limited (Hamleys)', 'Offline', '', True, parent='mt_select', db_label='Reliance Brands Limited (Hamleys)'),
    Channel('off_inst', 'OFF-INSTITUTIONAL', 'Offline', '', False),
    # EKA — own top-level offline channel (peer of GT Mass / MT / OFF-INST).
    # Relocated from the desktop app: its store registry now lives in the DB
    # (``eka_data`` table, see services.eka_data). db_key='EKA' → Daily Tasks
    # auto-detects the runs recorded under marketplace='EKA'. Still desktop-
    # processed for now (live=False, no web upload flow yet — migrated later).
    Channel('eka', 'EKA', 'Offline', 'EKA', False),
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


def db_label_to_channel() -> dict:
    """``{order_headers.marketplace_label: channel.key}`` — the fine-grained
    auto-detect key for channels that share a coarse ``marketplace`` (all MT
    children are marketplace='MT', told apart by their label)."""
    return {c.db_label: c.key for c in CHANNELS if c.db_label}


def children_of(key: str) -> list[Channel]:
    """The child channels nested under ``key`` (in display order); [] if none."""
    return [c for c in CHANNELS if c.parent == key]


def leaves() -> list[Channel]:
    """Trackable leaf channels — every non-parent channel. A channel is a parent
    (container only, no work steps of its own) when it has children; leaves are
    all the rest (children + standalone channels). Counts run over leaves."""
    parents = {c.parent for c in CHANNELS if c.parent}
    return [c for c in CHANNELS if c.key not in parents]


def as_dicts() -> list[dict]:
    """Flat JSON-safe list."""
    return [c.__dict__.copy() for c in CHANNELS]


def grouped() -> list[dict]:
    """``[{segment, channels:[dict,...]}, ...]`` in display order — API-ready.

    Nested: children are attached to their parent's ``children`` list and NOT
    listed at the top level. Every top-level channel dict carries a ``children``
    key (empty list when it has none) so consumers can render uniformly."""
    parents = {c.parent for c in CHANNELS if c.parent}
    out = []
    for seg in _SEGMENTS:
        chans = []
        for c in CHANNELS:
            if c.segment != seg or c.parent:
                continue                       # children handled under their parent
            d = c.__dict__.copy()
            d['children'] = [k.__dict__.copy() for k in children_of(c.key)]
            d['is_parent'] = c.key in parents
            chans.append(d)
        out.append({'segment': seg, 'channels': chans})
    return out
