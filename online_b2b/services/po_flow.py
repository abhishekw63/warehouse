"""
online_b2b.services.po_flow
===========================

Reusable, channel-agnostic **PO flow scaffold**: upload → review → confirm,
shared by every segment so new channels/marketplaces plug in instead of
duplicating the page.

ADDITIVE — the existing online_b2b and offline backends are NOT modified. A
channel wires in through a :class:`FlowSpec` (a processor factory + capability
flags + URL names + base template + optional channel-specific slots). The flow
stores the upload under ``MEDIA/<dirname>/<token>/`` (files + ``meta.json`` +
a signature-keyed ``preview.json`` cache), exactly mirroring the proven online
token model.

**Capabilities & slots keep ONE template working across mismatched channels.**
Not every channel is fully compatible (online has vendor-price compare +
Override + D365; GT Mass has none of those but has its own file-level
exceptions like *PO-number-missing* / *template-mismatch*). So:

* every optional block renders only when its capability flag is on
  (``caps`` = ``{'warehouse','margin','marketplace','vendor_cols','override',
  'ean_fix','exclude','d365'}``), and
* ``extra_partial`` is a **null-by-default** include where a channel injects its
  own panel — dark for every other channel.

A processor (the ``FlowSpec.processor`` factory result) must expose:

    preview() -> payload                 # no DB writes
    confirm(actions: dict) -> result     # writes, returns {'ok', 'run_id', ...}

``payload`` (and ``result`` where noted) is the unified review shape::

    {ok, error?, summary:{pos,lines,qty,value,affected,skipped},
     headers:[{po,location,order_type,items,qty,order_value}],
     lines:[{po,item_no,ean,description,qty,unit_price,our_mrp,
             status,exception_label, key}],
     affected:[ <subset of lines, status != OK> ],
     file_issues:[{file,problem,detail,kind}],   # channel-specific exceptions
     skipped:[{po,location,qty,order_value,marketplace_label}],
     warnings:[...], output_path?}
"""
from __future__ import annotations

import hashlib
import json
import shutil
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from django.conf import settings


@dataclass(frozen=True)
class FlowSpec:
    """Per-channel configuration for the shared flow. One instance per channel."""
    key: str                       # 'gt_mass' — stable slug
    title: str                     # 'GT Mass' — shown in the UI
    segment: str                   # 'Offline' / 'Online B2B'
    base_template: str             # e.g. 'core/base.html' / 'online_b2b/base_b2b.html'
    upload_dirname: str            # MEDIA subdir for this channel's uploads
    processor: Callable[[dict], object]   # (meta) -> processor with preview()/confirm()
    urls: dict                     # name map: upload/review/confirm/decision/discard/download/back/dashboard
    caps: frozenset = frozenset()  # capability flags (see module docstring)
    warehouses: tuple = ()         # ((code, label), ...) — only if 'warehouse' cap
    marketplaces: tuple = ()       # ((value, label), ...) — only if 'marketplace' cap
    default_margin: float | None = None
    extra_partial: str | None = None       # back-compat alias for slots['after_kpis']
    # Named channel-specific slots — partial templates injected at fixed anchors so
    # each channel can add its own bits where its inconsistency needs them. Every
    # slot is null by default (dark for channels that don't set it). Anchors:
    #   upload:  'fields'        (extra form fields)
    #   review:  'top'           (panel above the KPI cards)
    #            'after_kpis'    (panel under the KPI cards — e.g. file exceptions)
    #            'tabs'          (extra tab button)  + 'panes' (its tab pane)
    #            'actions'       (extra confirm-bar button)
    slots: dict = field(default_factory=dict)
    intro: str = ''                # one-line page subtitle
    accept: str = '.xlsx,.xls,.xlsm,.csv,.pdf'

    def slot_map(self) -> dict:
        """Effective slots (``extra_partial`` folds into ``after_kpis``)."""
        s = dict(self.slots)
        if self.extra_partial and 'after_kpis' not in s:
            s['after_kpis'] = self.extra_partial
        return s

    def caps_map(self) -> dict:
        """``{cap: True}`` for template ``{% if caps.exclude %}`` checks."""
        return {c: True for c in self.caps}


# ── token / upload store ─────────────────────────────────────────────────
def _root(spec: FlowSpec) -> Path:
    return Path(settings.MEDIA_ROOT) / spec.upload_dirname


def _dir(spec: FlowSpec, token: str) -> Path:
    return _root(spec) / token


def new_token() -> str:
    return uuid.uuid4().hex[:12]


def save_upload(spec: FlowSpec, files, extra: dict | None = None) -> str:
    """Persist uploaded file(s) under a fresh token + write ``meta.json``."""
    token = new_token()
    d = _dir(spec, token)
    d.mkdir(parents=True, exist_ok=True)
    paths = []
    for f in files:
        name = Path(getattr(f, 'name', 'upload')).name
        dest = d / name
        with open(dest, 'wb') as out:
            for chunk in f.chunks():
                out.write(chunk)
        paths.append(str(dest))
    meta = {'files': paths, 'decisions': {}, 'locked': False, 'run_id': None}
    meta.update(extra or {})
    _write_meta(spec, token, meta)
    return token


def _meta_path(spec: FlowSpec, token: str) -> Path:
    return _dir(spec, token) / 'meta.json'


def load_meta(spec: FlowSpec, token: str) -> dict | None:
    p = _meta_path(spec, token)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except Exception:  # noqa: BLE001
        return None


def _write_meta(spec: FlowSpec, token: str, meta: dict) -> None:
    _meta_path(spec, token).write_text(
        json.dumps(meta, ensure_ascii=False, indent=1), encoding='utf-8')


def _sig(meta: dict) -> str:
    """Signature that invalidates the preview cache when inputs change."""
    basis = {'files': meta.get('files', []),
             'warehouse': meta.get('warehouse', ''),
             'margin': meta.get('margin_pct', ''),
             'ean_fixes': meta.get('ean_fixes', {})}
    return hashlib.md5(
        json.dumps(basis, sort_keys=True).encode('utf-8')).hexdigest()


def preview(spec: FlowSpec, token: str, meta: dict) -> dict:
    """Cached preview — runs the processor once per (inputs) signature."""
    cache = _dir(spec, token) / 'preview.json'
    sig = _sig(meta)
    if cache.exists():
        try:
            blob = json.loads(cache.read_text(encoding='utf-8'))
            if blob.get('sig') == sig:
                return blob['payload']
        except Exception:  # noqa: BLE001
            pass
    payload = spec.processor(meta).preview()
    try:
        cache.write_text(json.dumps({'sig': sig, 'payload': payload},
                                    ensure_ascii=False), encoding='utf-8')
    except Exception:  # noqa: BLE001
        pass
    return payload


def _invalidate(spec: FlowSpec, token: str) -> None:
    c = _dir(spec, token) / 'preview.json'
    if c.exists():
        try:
            c.unlink()
        except Exception:  # noqa: BLE001
            pass


def set_decision(spec: FlowSpec, token: str, key: str, action: str,
                 override_cp: str = '', remark: str = '') -> int:
    """Persist one per-line decision onto ``meta.json``. Returns total decided."""
    meta = load_meta(spec, token) or {}
    dec = meta.setdefault('decisions', {})
    if action:
        dec[key] = {'action': action,
                    'override_cp': override_cp or None,
                    'remark': remark or ''}
    else:
        dec.pop(key, None)
    _write_meta(spec, token, meta)
    return len(dec)


def discard(spec: FlowSpec, token: str) -> None:
    d = _dir(spec, token)
    if d.exists():
        shutil.rmtree(d, ignore_errors=True)


def _overlay_decisions(payload: dict, meta: dict) -> None:
    """Attach saved per-line decisions to lines/affected so the review page
    shows prior operator choices (and so confirm can replay them)."""
    dec = meta.get('decisions', {})
    for bucket in ('lines', 'affected'):
        for ln in payload.get(bucket, []) or []:
            ln['decision'] = dec.get(ln.get('key', ''), {})


def review_context(spec: FlowSpec, token: str, meta: dict) -> dict:
    """Build the full template context for the shared ``review.html``."""
    payload = preview(spec, token, meta)
    _overlay_decisions(payload, meta)
    return {
        'spec': spec,
        'title': spec.title,
        'segment': spec.segment,
        'base_template': spec.base_template,
        'intro': spec.intro,
        'caps': spec.caps_map(),
        'slots': spec.slot_map(),
        'token': token,
        'meta': meta,
        'r': payload,
        's': payload.get('summary', {}),
        'locked': bool(meta.get('locked')),
        'run_id': meta.get('run_id'),
        'has_download': bool(meta.get('output_path')),
        'warehouses': spec.warehouses,
        'marketplaces': spec.marketplaces,
        # URL names (used as `{% url u_confirm token %}` — variable name form)
        'u_upload': spec.urls['upload'], 'u_review': spec.urls['review'],
        'u_confirm': spec.urls['confirm'], 'u_decision': spec.urls['decision'],
        'u_discard': spec.urls['discard'], 'u_download': spec.urls['download'],
        'u_back': spec.urls['back'], 'u_dashboard': spec.urls['dashboard'],
    }


def confirm(spec: FlowSpec, token: str, meta: dict, actions: dict | None = None
            ) -> dict:
    """Run the processor's confirm (DB write) and lock the token on success."""
    result = spec.processor(meta).confirm(actions or meta.get('decisions', {}))
    if result.get('ok') and result.get('run_id'):
        meta['locked'] = True
        meta['run_id'] = result['run_id']
        if result.get('output_path'):
            meta['output_path'] = result['output_path']
        _write_meta(spec, token, meta)
    return result


def download_path(spec: FlowSpec, token: str) -> Path | None:
    """The workbook to serve on the review 'download' link. Prefers the workbook
    already produced at confirm; otherwise, if the channel's processor exposes a
    ``workbook()`` method (``'download'`` cap), generates it on demand so the
    operator can download during review — before confirm."""
    meta = load_meta(spec, token) or {}
    p = meta.get('output_path')
    if p and Path(p).exists():
        return Path(p)
    gen = getattr(spec.processor(meta), 'workbook', None)
    if callable(gen):
        try:
            out = gen()
            return Path(out) if out else None
        except Exception:  # noqa: BLE001
            return None
    return None
