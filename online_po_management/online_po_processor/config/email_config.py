"""
config.email_config
===================

SMTP credentials and recipient list for the Online PO email reports.

Two-layer config
----------------
**Layer 1 — defaults baked into code** (this module):
    Sensible production defaults that work out of the box for the RENEE
    online-PO team. Same SMTP account as GT Mass Dump, TO matches the
    primary recipient, CC is trimmed to just the online-B2B channel.

**Layer 2 — optional JSON override at**
    ``Calculation Data/email_config.json``

    Lets the user/admin override any subset of the defaults without
    editing source code (handy for a new hire, an address change, or a
    temporary re-route). The JSON can contain any of these keys::

        {
            "EMAIL_SENDER":      "foo@gmail.com",
            "EMAIL_PASSWORD":    "abcd efgh ijkl mnop",
            "SMTP_SERVER":       "smtp.gmail.com",
            "SMTP_PORT":         587,
            "DEFAULT_RECIPIENT": "someone@reneecosmetics.in",
            "CC_RECIPIENTS":     ["a@...", "b@..."]
        }

    Unknown keys are ignored. Missing keys inherit from the defaults.

Public surface
--------------
``get_email_config()``
    Returns the merged, ready-to-use config dict.

The config dict shape mirrors GT Mass Dump's ``EMAIL_CONFIG`` so code
ported from there can use it verbatim.
"""

from __future__ import annotations
import json
import logging
import os
from typing import Any, Dict, List

from online_po_processor.config.paths import get_bundled_data_folder


# Filename for the optional user-override JSON (lives alongside the
# bundled master/mapping files in ``Calculation Data/``).
_EMAIL_CONFIG_FILENAME = "email_config.json"


# ── Defaults ───────────────────────────────────────────────────────────
# These are the production defaults for RENEE's online-PO automation.
# They match GT Mass Dump's SMTP account/sender but diverge on the CC
# list — online PO notifications go only to the offline-B2B inbox
# (the other 3 CCs from GT Mass are for dispatch/sales teams not
# involved in online-PO flow).

_DEFAULT_EMAIL_CONFIG: Dict[str, Any] = {
    # ── Sender credentials (Gmail App Password) ──
    # NEVER hardcoded. Sourced from the environment (.env on a host) or the
    # gitignored Calculation Data/email_config.json (get_email_config overlays
    # it over these defaults). Blank when unset → email soft-fails, never sends
    # from a committed secret.
    'EMAIL_SENDER':   os.environ.get('EMAIL_SENDER', ''),
    'EMAIL_PASSWORD': os.environ.get('EMAIL_PASSWORD', ''),

    # ── SMTP server ────────────────────────────────────────────────────
    'SMTP_SERVER': 'smtp.gmail.com',
    'SMTP_PORT':   587,

    # ── Primary TO recipient ───────────────────────────────────────────
    # Same as GT Mass — Abhishek owns order management and wants both
    # online and offline reports in one inbox.
    'DEFAULT_RECIPIENT': 'abhishek.wagh@reneecosmetics.in',

    # ── CC recipients ──────────────────────────────────────────────────
    # Only the online-B2B channel. GT Mass Dump CCs 3 additional people
    # for offline/dispatch visibility; those do not apply here.
    'CC_RECIPIENTS': [
        'onlineb2b@reneecosmetics.in',
        # 'kuldeep.joshi@reneecosmetics.in',
        'pintu.sharma@reneecosmetics.in',
        'aritra.barmanray@reneecosmetics.in',
        'jitendra.r@reneecosmetics.in'
    ],
}


def get_email_config() -> Dict[str, Any]:
    """
    Return the effective email config (defaults + any JSON overrides).

    The returned dict is a shallow copy — callers may mutate it without
    affecting the module-level defaults.

    Returns:
        Dict with keys matching the shape documented at the top of this
        module.
    """
    # Start from a fresh copy of the defaults so a caller mutating the
    # returned dict can't leak changes into later calls.
    config = dict(_DEFAULT_EMAIL_CONFIG)
    # Deep-copy the list so callers that append() don't affect defaults.
    config['CC_RECIPIENTS'] = list(_DEFAULT_EMAIL_CONFIG['CC_RECIPIENTS'])

    overrides = _load_overrides()

    if overrides:
        _apply_overrides(config, overrides)

    return config


# ── Internal helpers ───────────────────────────────────────────────────

def _load_overrides() -> Dict[str, Any]:
    """
    Try to read the optional JSON override file.

    Missing file → empty dict (normal case; defaults apply).
    Invalid JSON → warning logged, empty dict returned (app must not
    crash just because someone typo'd a comma).

    Returns:
        Parsed JSON dict, or empty dict if the file is absent/unreadable.
    """
    folder = get_bundled_data_folder(create=False)

    # Folder may not exist on a fresh install.
    if folder is None or not folder.exists():
        return {}

    override_path = folder / _EMAIL_CONFIG_FILENAME

    if not override_path.exists():
        return {}

    try:
        with override_path.open('r', encoding='utf-8') as fh:
            data = json.load(fh)
    except (OSError, json.JSONDecodeError) as e:
        logging.warning(
            "Could not read %s — using built-in defaults: %s",
            override_path, e,
        )
        return {}

    # Must be a dict at the top level; silently ignore arrays/scalars.
    if not isinstance(data, dict):
        logging.warning(
            "%s is not a JSON object — using built-in defaults",
            override_path,
        )
        return {}

    return data


def _apply_overrides(config: Dict[str, Any],
                      overrides: Dict[str, Any]) -> None:
    """
    Mutate ``config`` in place with values from ``overrides``.

    Validates types for each known key and silently drops invalid
    values (again, we never want to crash on a malformed config).

    Args:
        config:    The defaults dict to mutate.
        overrides: User-provided overrides parsed from JSON.
    """
    # Simple string keys
    for key in (
        'EMAIL_SENDER', 'EMAIL_PASSWORD',
        'SMTP_SERVER', 'DEFAULT_RECIPIENT',
    ):
        if key in overrides and isinstance(overrides[key], str):
            config[key] = overrides[key]

    # SMTP_PORT — must be int (ints and int-coercible strings both OK)
    if 'SMTP_PORT' in overrides:
        try:
            config['SMTP_PORT'] = int(overrides['SMTP_PORT'])
        except (TypeError, ValueError):
            logging.warning(
                "SMTP_PORT override is not an integer — keeping default "
                "(%s)", config['SMTP_PORT'],
            )

    # CC_RECIPIENTS — must be a list of strings; empty list allowed
    if 'CC_RECIPIENTS' in overrides:
        cc_val = overrides['CC_RECIPIENTS']
        if isinstance(cc_val, list) and all(
            isinstance(x, str) for x in cc_val
        ):
            config['CC_RECIPIENTS'] = list(cc_val)
        else:
            logging.warning(
                "CC_RECIPIENTS override is not a list of strings — "
                "keeping default list",
            )