"""online_b2b.services.common — small shared web helpers.

Behavior-preserving consolidation of tiny utilities that were copy-pasted across
the views (path-traversal guard, upload save, AJAX detection, POST-dict). Pure and
side-effect-free except where noted; import and delegate to these instead of
re-implementing. See [[page-asset-separation]]/[[dry-skeleton-first]] spirit for the
Python side.
"""
from __future__ import annotations

from pathlib import Path

from django.http import Http404


def token_dir(root: Path, token: str) -> Path:
    """Resolve ``root/token`` and refuse anything that escapes ``root`` (path
    traversal). Returns the resolved Path or raises Http404 — identical to the
    per-feature ``_tok_dir``/``_token_dir`` guards it replaces."""
    root = Path(root)
    base = root.resolve()
    d = (root / token).resolve()
    if d != base and base not in d.parents:
        raise Http404()
    return d


def save_upload(fileobj, dest) -> Path:
    """Stream an uploaded file to ``dest`` in chunks (never loads it all into
    memory). Returns the destination Path. Caller ensures the parent dir exists."""
    dest = Path(dest)
    with open(dest, 'wb') as out:
        for chunk in fileobj.chunks():
            out.write(chunk)
    return dest


def is_ajax(request) -> bool:
    """True for an XMLHttpRequest fetch (the app's AJAX convention). Matches the
    existing ``_is_ajax`` exactly (header ``X-Requested-With: XMLHttpRequest``)."""
    return request.headers.get('x-requested-with') == 'XMLHttpRequest'


def post_dict(request) -> dict:
    """``request.POST`` as a plain dict minus the CSRF token — the exact pattern
    used by the ship-to / eka CRUD endpoints."""
    return {k: v for k, v in request.POST.items() if k != 'csrfmiddlewaretoken'}
