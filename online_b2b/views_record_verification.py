"""Record Verification — STANDALONE, class-based views.

Upload D365 Headers + Lines → reconcile against our recorded data (recorded +
excluded) per PO → flag mismatches (qty / value / pincode) → persist a checked-PO
log with per-PO deltas. Thin views; all logic lives in
``services.record_verification`` (which reuses ``full_validation``).

Self-contained + removable: this file + services/record_verification.py +
templates/online_b2b/record_verification.html + the URL block + one nav link.
Own media dir (``b2b_recordcheck``).
"""
from __future__ import annotations

import datetime as _dt
import json
import uuid
from pathlib import Path

import pandas as pd
from django.conf import settings
from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin
from django.http import FileResponse, Http404, JsonResponse
from django.shortcuts import redirect
from django.urls import reverse
from django.views import View
from django.views.generic import TemplateView

from .services import common
from .services import record_verification as rv

_UP = Path(settings.MEDIA_ROOT) / 'b2b_recordcheck'


def _tok_dir(token: str) -> Path:
    return common.token_dir(_UP, token)


def _result_path(token: str) -> Path:
    return _tok_dir(token) / 'result.json'


def _load_result(token: str):
    """(path, result-dict) for a run token, or Http404 if the run is gone/expired."""
    rp = _result_path(token)
    if not rp.exists():
        raise Http404('Review not found or expired.')
    return rp, json.loads(rp.read_text(encoding='utf-8'))


def _save_result(rp: Path, res: dict) -> None:
    rp.write_text(json.dumps(res, default=str), encoding='utf-8')


def _token_url(token: str) -> str:
    return f"{reverse('b2b_record_verify')}?token={token}"


def _actor(request) -> str:
    return getattr(request.user, 'username', '') or 'system'


def _done(request, token: str, msg: str, **extra):
    """Uniform success response: AJAX → JSON, otherwise flash + redirect to the run."""
    if common.is_ajax(request):
        return JsonResponse({'ok': True, 'message': msg, **extra})
    messages.success(request, msg)
    return redirect(_token_url(token))


def _looks_like(path: str, needle: str) -> bool:
    """True if the first row of the workbook has a column matching ``needle``."""
    try:
        cols = pd.read_excel(path, nrows=0).columns
        return any(needle.lower() == str(c).lower() for c in cols)
    except Exception:  # noqa: BLE001
        return False


_LINE_PROBLEM = {'MISSING_IN_D365', 'EXTRA_IN_D365', 'QTY_MISMATCH'}


def _augment(data: dict) -> dict:
    """Add the review-page-style tab data on top of the reconcile result:
    per-PO line drill-down, the Orders/Externals split, the affected-line list, and
    KPI totals. Pure/derived — no new reconciliation, just reshaping for display."""
    headers = data.get('headers', [])
    lines = data.get('lines', [])
    by_po: dict = {}
    for ln in lines:
        by_po.setdefault(ln.get('po'), []).append(ln)
    for h in headers:
        h['lines'] = by_po.get(h.get('po'), [])
        h['bad_lines'] = sum(1 for ln in h['lines'] if ln.get('status') in _LINE_PROBLEM)
    data['orders'] = [h for h in headers if h.get('status') != 'EXTERNAL']
    data['externals'] = [h for h in headers if h.get('status') == 'EXTERNAL']
    data['affected_lines'] = [ln for ln in lines if ln.get('status') in _LINE_PROBLEM]
    data['n_lines'] = len(lines)
    data['n_affected_lines'] = len(data['affected_lines'])
    data['tot_qty'] = int(sum(ln.get('d365_qty') or 0 for ln in lines))
    data['tot_val'] = round(sum(h.get('d365_val') or 0 for h in headers))
    return data


def _saved_runs(limit: int = 12) -> list:
    """Verification runs parked as 'review later' drafts (not yet recorded), newest
    first, for the resume list. Each run already persists under its token; this just
    surfaces the ones flagged as saved."""
    out: list = []
    for rp in sorted(_UP.glob('*/result.json'),
                     key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            res = json.loads(rp.read_text(encoding='utf-8'))
        except Exception:  # noqa: BLE001
            continue
        if res.get('draft') and not res.get('confirmed'):
            vs = (res.get('data') or {}).get('verify_summary', {})
            out.append({'token': rp.parent.name, 'saved_at': res.get('saved_at'),
                        'saved_by': res.get('saved_by'), 'note': res.get('saved_note'),
                        'checked': vs.get('checked'), 'ok': vs.get('ok'),
                        'mismatch': vs.get('mismatch'), 'external': vs.get('external')})
            if len(out) >= limit:
                break
    return out


class RecordVerificationView(LoginRequiredMixin, TemplateView):
    """Upload page + coverage + checked-PO log + the last run's result (``?token=``)."""
    template_name = 'online_b2b/record_verification.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['coverage'] = rv.coverage()
        ctx['log'] = rv.checked_log(limit=300)
        ctx['saved_runs'] = _saved_runs()
        token = self.request.GET.get('token', '')
        ctx['token'] = token
        rp = _result_path(token) if token else None
        if rp and rp.exists():
            res = json.loads(rp.read_text(encoding='utf-8'))
            if res.get('ok'):
                ctx['data'] = _augment(res['data'])
                ctx['confirmed'] = res.get('confirmed', False)
        return ctx


class RecordVerificationRunView(LoginRequiredMixin, View):
    """POST Headers + Lines → verify → stash result + Excel → redirect with token.
    Accepts the two named inputs, or a 2-file batch it auto-sorts by columns."""

    def post(self, request):
        token = uuid.uuid4().hex[:12]
        d = _UP / token
        fdir = d / 'files'
        fdir.mkdir(parents=True, exist_ok=True)

        # gather files: named inputs first, else a dropped batch
        named = {k: request.FILES.get(k) for k in ('headers_file', 'lines_file')}
        batch = [f for f in named.values() if f] or list(request.FILES.getlist('rv_files'))
        if not batch:
            messages.error(request, 'Upload the D365 Sales Orders (Headers) and Sales Lines files.')
            return redirect('b2b_record_verify')

        saved = []
        for i, f in enumerate(batch):
            p = fdir / f"{i}_{Path(f.name).name}"
            common.save_upload(f, p)
            saved.append(str(p))

        # resolve which file is Headers vs Lines (by their signature columns)
        headers = next((p for p in saved if _looks_like(p, 'External Document No.')), None)
        lines = next((p for p in saved if _looks_like(p, 'Document No.') and p != headers), None)
        if not headers or not lines:
            messages.error(request, "Couldn't identify both files — need the D365 Sales "
                            "Orders (has 'External Document No.') and Sales Lines (has 'Document No.').")
            return redirect('b2b_record_verify')

        # PHASE 1: preview only — nothing is written until Confirm.
        res = rv.preview(headers, lines)
        if res.get('ok'):
            res['confirmed'] = False
            try:
                rv.build_workbook(res['data'], str(d / 'record_verification.xlsx'))
            except Exception as e:  # noqa: BLE001 — Excel is a convenience
                res['data']['excel_error'] = f'{type(e).__name__}: {e}'
        _save_result(d / 'result.json', res)
        if not res.get('ok'):
            messages.error(request, res.get('error', 'Verification failed.'))
            return redirect('b2b_record_verify')
        vs = res['data'].get('verify_summary', {})
        messages.info(request, f"Reviewed {vs.get('checked', 0)} PO(s) — "
                      f"{vs.get('ok', 0)} clean, {vs.get('mismatch', 0)} mismatch, "
                      f"{vs.get('external', 0)} external. Review, then Confirm to record.")
        return redirect(_token_url(token))


class RecordVerificationConfirmView(LoginRequiredMixin, View):
    """PHASE 2 — persist the reviewed verification to the checked-PO log."""

    def post(self, request, token):
        rp, res = _load_result(token)
        if not res.get('ok'):
            messages.error(request, 'Nothing to confirm.')
            return redirect('b2b_record_verify')
        # review-page style: operator ticks which POs to record (EXTERNAL POs are
        # "pushed" without a cross-check). No ticks posted → record all (back-compat).
        picked = request.POST.getlist('push_po')
        out = rv.confirm(res['data'].get('headers', []),
                         checked_by=_actor(request), only_pos=picked or None)
        res['confirmed'] = True
        _save_result(rp, res)
        n = out.get('confirmed', 0)
        return _done(request, token, f"✓ Verification recorded — {n} PO(s) logged.",
                     confirmed=n)


class RecordVerificationSaveLaterView(LoginRequiredMixin, View):
    """Park this verification as a **review-later draft** — nothing is recorded. The
    run already persists under its token; this flags it + timestamps it so it shows
    in the 'Resume a saved check' list and can be reopened via its token URL."""

    def post(self, request, token):
        rp, res = _load_result(token)
        res['draft'] = True
        res['saved_at'] = _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        res['saved_by'] = _actor(request)
        res['saved_note'] = (request.POST.get('note') or '').strip()[:200]
        _save_result(rp, res)
        return _done(request, token,
                     '🕒 Saved for review later — resume it anytime from "Resume a saved check".')


class RecordVerificationDownloadView(LoginRequiredMixin, View):
    """Serve the reconciliation Excel for a run."""

    def get(self, request, token):
        xp = _tok_dir(token) / 'record_verification.xlsx'
        if not xp.exists():
            raise Http404('File not found or expired.')
        return FileResponse(open(xp, 'rb'), as_attachment=True,
                            filename='Record_Verification.xlsx')
