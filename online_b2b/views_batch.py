"""Batch Run — **Phase 1 front half**: upload many PO files → READ-ONLY marketplace
detection → operator **confirm/override grid**.

RECORDS NOTHING and runs no engine. It only detects each file's marketplace
(``services.batch_flow``) and persists the operator-confirmed file→MP mapping to the
run token, so a *later, separate* processing step (Phase 1 back half) can pick it up.
The confirm gate enforces 100%% coverage — every file must have a marketplace chosen
before it's accepted; nothing is auto-trusted.

Standalone + removable: this file + services/batch_flow.py + the batch_detect
templates + the URL block + one nav link. Own media dir (``b2b_batch``). Deleting it
leaves the app exactly as before. See [[project-backlog]] (Batch Run).
"""
from __future__ import annotations

import json
import uuid
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin
from django.http import Http404, JsonResponse
from django.shortcuts import redirect
from django.urls import reverse
from django.views import View
from django.views.generic import TemplateView

from .services import batch_flow as bf
from .services import common

_UP = Path(settings.MEDIA_ROOT) / 'b2b_batch'


def _tok_dir(token: str) -> Path:
    return common.token_dir(_UP, token)


class BatchDetectView(LoginRequiredMixin, TemplateView):
    """Upload page + (``?token=``) the detection/confirm grid for a run."""
    template_name = 'online_b2b/batch_detect.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['marketplaces'] = bf.detectable_marketplaces()
        from .services import engine_bridge as eb
        ctx['warehouses'] = eb.warehouse_choices()
        ctx['default_warehouse'] = eb.default_warehouse()
        token = self.request.GET.get('token', '')
        ctx['token'] = token
        if token:
            d = _tok_dir(token)
            rp = d / 'detect.json'
            if rp.exists():
                res = json.loads(rp.read_text(encoding='utf-8'))
                ctx['rows'] = res.get('rows', [])
                ctx['confirmed'] = res.get('confirmed', False)
                ctx['plan'] = res.get('plan')
            pp = d / 'preview.json'
            if pp.exists():
                ctx['preview'] = json.loads(pp.read_text(encoding='utf-8'))
        return ctx


class BatchDetectRunView(LoginRequiredMixin, View):
    """POST many files → save to the token dir → detect (read-only) → redirect with
    token. Writes nothing to the DB and runs no engine."""

    def post(self, request):
        files = list(request.FILES.getlist('batch_files'))
        if not files:
            messages.error(request, 'Drop the marketplace PO files to detect.')
            return redirect('b2b_batch')
        token = uuid.uuid4().hex[:12]
        fdir = _UP / token / 'files'
        fdir.mkdir(parents=True, exist_ok=True)
        saved = []
        for i, f in enumerate(files):
            p = fdir / f"{i}_{Path(f.name).name}"
            common.save_upload(f, p)
            saved.append(str(p))

        rows = bf.detect(saved)                    # pure/read-only detection
        for i, r in enumerate(rows):
            r['idx'] = i                           # stable handle for the confirm POST
        (_UP / token / 'detect.json').write_text(
            json.dumps({'rows': rows, 'confirmed': False}, default=str),
            encoding='utf-8')

        n_hi = sum(1 for r in rows if r['confidence'] == 'high')
        n_chk = len(rows) - n_hi
        messages.info(request, f"Detected {len(rows)} file(s): {n_hi} high-confidence, "
                      f"{n_chk} to confirm. Nothing is processed until you confirm below.")
        return redirect(f"{reverse('b2b_batch')}?token={token}")


class BatchConfirmView(LoginRequiredMixin, View):
    """Persist the operator-confirmed file→MP mapping. RECORDS NOTHING and runs no
    engine — Phase 1 front half stops here. Enforces that EVERY file has a chosen
    marketplace (100%% coverage) before the mapping is accepted."""

    def post(self, request, token):
        rp = _tok_dir(token) / 'detect.json'
        if not rp.exists():
            raise Http404('Detection not found or expired.')
        res = json.loads(rp.read_text(encoding='utf-8'))
        rows = res.get('rows', [])

        plan, missing = [], []
        valid = set(bf.detectable_marketplaces())
        for r in rows:
            chosen = (request.POST.get(f"mp_{r['idx']}") or '').strip()
            if not chosen or chosen not in valid:
                missing.append(r['file'])
            plan.append({'file': r['file'], 'marketplace': chosen})
        if missing:
            msg = ('Pick a marketplace for every file first — missing/invalid: '
                   + ', '.join(missing[:6]))
            if common.is_ajax(request):
                return JsonResponse({'ok': False, 'error': msg}, status=400)
            messages.error(request, msg)
            return redirect(f"{reverse('b2b_batch')}?token={token}")

        res['confirmed'] = True
        res['plan'] = plan
        rp.write_text(json.dumps(res, default=str), encoding='utf-8')
        msg = (f"✓ Confirmed {len(plan)} file→marketplace mapping(s). "
               "Nothing was processed — detection/confirm only (Phase 1).")
        if common.is_ajax(request):
            return JsonResponse({'ok': True, 'plan': plan, 'message': msg})
        messages.success(request, msg)
        return redirect(f"{reverse('b2b_batch')}?token={token}")


class BatchPreviewView(LoginRequiredMixin, View):
    """Run each CONFIRMED file-group through its EXISTING per-MP processor to build
    ONE combined READ-ONLY preview (per-MP KPIs + master totals). Reuses
    ``engine_bridge.preview`` verbatim — it parses/prices but writes NOTHING to the
    business DB (that's ``confirm``, a separate later gate). Files are grouped by
    the confirmed marketplace so multi-file MPs are processed together, exactly like
    the single-MP flow."""

    _AGG = ('pos', 'lines', 'qty', 'value', 'affected', 'skipped')

    def post(self, request, token):
        from collections import defaultdict
        from .services import engine_bridge as eb

        d = _tok_dir(token)
        rp = d / 'detect.json'
        if not rp.exists():
            raise Http404('Detection not found or expired.')
        res = json.loads(rp.read_text(encoding='utf-8'))
        if not res.get('confirmed') or not res.get('plan'):
            messages.error(request, 'Confirm the file→marketplace mapping first.')
            return redirect(f"{reverse('b2b_batch')}?token={token}")

        warehouse = (request.POST.get('warehouse') or '').strip() or eb.default_warehouse()
        if warehouse not in set(eb.warehouse_choices()):
            warehouse = eb.default_warehouse()

        fdir = d / 'files'
        groups: dict = defaultdict(list)
        for item in res['plan']:
            p = fdir / item['file']
            if p.exists():
                groups[item['marketplace']].append(str(p))

        per_mp, totals = [], {k: 0 for k in self._AGG}
        for mp, paths in groups.items():
            try:
                pv = eb.preview(mp, paths, warehouse=warehouse,
                                margin_pct=eb.default_margin_pct(mp) / 100.0)
            except Exception as e:  # noqa: BLE001 — one MP's failure never 500s the batch
                per_mp.append({'mp': mp, 'files': len(paths), 'ok': False,
                               'error': f'{type(e).__name__}: {e}'})
                continue
            if not pv.get('ok'):
                per_mp.append({'mp': mp, 'files': len(paths), 'ok': False,
                               'error': pv.get('error', 'preview failed')})
                continue
            s = pv.get('summary') or {}
            row = {'mp': mp, 'files': len(paths), 'ok': True,
                   'mismatch': s.get('mismatch', 0)}
            for k in self._AGG:
                row[k] = s.get(k, 0) or 0
                totals[k] += row[k]
            per_mp.append(row)

        preview = {'warehouse': warehouse, 'per_mp': per_mp, 'totals': totals,
                   'n_ok': sum(1 for r in per_mp if r['ok']),
                   'n_fail': sum(1 for r in per_mp if not r['ok'])}
        (d / 'preview.json').write_text(json.dumps(preview, default=str), encoding='utf-8')
        messages.info(request, f"Previewed {len(per_mp)} marketplace group(s) — "
                      f"{totals['pos']} PO(s), {totals['lines']} line(s), "
                      f"{totals['affected']} affected. Nothing was recorded.")
        return redirect(f"{reverse('b2b_batch')}?token={token}")
