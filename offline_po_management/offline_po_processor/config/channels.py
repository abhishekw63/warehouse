"""
offline_po_processor.config.channels
====================================

The **channel registry** — the single source of truth for which offline
order-processing tools exist and where their scripts live. Adding a new
channel is one entry here (no launcher code changes), which is what makes
the offline side scalable.

Each channel is, for now, a self-contained standalone script (its core
logic unchanged from the old ``standalone_files/`` version) living under
``offline_po_management/channels/<key>/``. The launcher runs the selected
channel as an independent subprocess, so the channels stay decoupled and a
future refactor (or the eventual OMT merge with the Online side) can swap a
script for a proper package without touching the registry contract.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Tuple


@dataclass(frozen=True)
class Channel:
    """One offline channel.

    Attributes:
        key:         stable identifier (folder name under ``channels/``).
        name:        display label shown on the launcher button.
        description: one-line summary shown beside the button.
        script:      path to the channel's entry script, **relative to the
                     ``channels/`` root** (so the registry stays
                     location-independent).
        enabled:     show it on the launcher (set False to stage a
                     not-yet-ready channel without removing its entry).
    """

    key: str
    name: str
    description: str
    script: str
    enabled: bool = True


# ── The registry ────────────────────────────────────────────────────────
# Order here = order on the launcher. Append new channels as they migrate.
CHANNELS: Tuple[Channel, ...] = (
    Channel(
        key='eka',
        name='EKA',
        description='Transfer / Sales Order constructor (EKA branches)',
        script='eka/standalone_EKA_constructor.py',
    ),
    Channel(
        key='gt_mass',
        name='GT Mass',
        description='GT Mass dump generator (PO → D365 import)',
        script='gt_mass/standalone_gt_mass_automation.py',
    ),
    Channel(
        key='mt_select',
        name='MT Select',
        description='Modern-trade multi-channel PO processor (e.g. H&G)',
        script='mt_select/standalone_mt_select_automation.py',
    ),
    Channel(
        key='off_inst',
        name='Off Institutional',
        description='Institutional PO processor (INCS) — regular + tester SOs',
        script='off_inst/standalone_off_inst_automation.py',
    ),
)


def channels_root() -> Path:
    """
    Absolute path to the ``channels/`` directory.

    Resolved relative to this file:
    ``offline_po_processor/config/channels.py`` → up to
    ``offline_po_management/`` → ``/channels``. Keeps the registry working
    regardless of the current working directory.
    """
    return Path(__file__).resolve().parents[2] / 'channels'


def channel_script(ch: Channel) -> Path:
    """Absolute path to a channel's entry script."""
    return channels_root() / ch.script


def channel_workdir(ch: Channel) -> Path:
    """
    Working directory to launch the channel from — its own folder. The
    standalone scripts resolve bundled data either script-relative
    (``get_script_dir()``) or CWD-relative (e.g. GT Mass's ``output/``), so
    launching with this as CWD makes both resolve inside the channel
    folder.
    """
    return channel_script(ch).parent


def enabled_channels() -> Tuple[Channel, ...]:
    """The channels to show on the launcher."""
    return tuple(c for c in CHANNELS if c.enabled)
