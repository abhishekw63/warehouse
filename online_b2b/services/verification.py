"""
online_b2b.services.verification
================================

Reusable, **channel-agnostic verification scaffold** — the "additional
verification" analogue of :mod:`online_b2b.services.po_flow`. One skeleton, many
channels plug in: a channel's processor produces a single JSON-safe
``verification`` dict (via :func:`build`) and stashes it on the preview payload;
the shared review-page modal ``po_flow/_verification_modal.html`` renders *any*
such dict in a big popup (title + columns + rows + summary). Adding verification
for the NEXT marketplace is just building the dict — zero new UI/route code.

The contract (JSON-safe, API-ready)::

    verification = {
        'title':    'PDF Address Verification',   # page heading
        'subtitle': '…',                          # one-line description (optional)
        'source':   '3 PDF(s) read for Lifestyle',# provenance (optional)
        'columns': [                              # ordered column descriptors
            {'key': 'store', 'label': 'Store', 'mono': True},
            {'key': 'match', 'label': 'Match', 'kind': 'match'},  # OK/MISMATCH/NO_PDF pill
            ...
        ],
        'rows': [ {col_key: value, ..., 'detail': '…', 'match': 'OK'}, ... ],
        'match_key': 'match',                     # which row key holds OK/MISMATCH
        'summary': {'checked': N, 'mismatch': M, 'ok': K, 'other': J},
    }

The review page needs only ``summary`` (+ a link); it never touches internals.
"""
from __future__ import annotations

# Row states that count as a hard mismatch (drives the summary + red styling).
_MISMATCH_STATES = {'MISMATCH', 'FAIL', 'BAD', 'ERROR'}
_OK_STATES = {'OK', 'PASS', 'MATCH'}


def build(*, title: str, columns: list[dict], rows: list[dict],
          match_key: str = 'match', subtitle: str = '', source: str = '') -> dict:
    """Assemble a channel-agnostic ``verification`` dict from structured rows.

    Computes the ``summary`` (checked / ok / mismatch / other) from ``match_key``
    so every channel gets consistent counts and styling for free. Returns a plain
    JSON-safe dict (no objects) — API-ready."""
    checked = len(rows)
    mismatch = sum(1 for r in rows
                   if str(r.get(match_key, '')).upper() in _MISMATCH_STATES)
    ok = sum(1 for r in rows
             if str(r.get(match_key, '')).upper() in _OK_STATES)
    return {
        'title': title,
        'subtitle': subtitle,
        'source': source,
        'columns': columns,
        'rows': rows,
        'match_key': match_key,
        'summary': {
            'checked': checked,
            'ok': ok,
            'mismatch': mismatch,
            'other': checked - ok - mismatch,
        },
    }
