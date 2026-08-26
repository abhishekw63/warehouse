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
import shutil
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
from .services import gt_select_import as gts   # external-order capture (classify + dedup)
from .services import marketplaces as mkt
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
    # A line is "affected" if its qty/SKU is off (_LINE_PROBLEM) OR its VALUE is off
    # (val_ok == 'NO') — a value-only mismatch (qty matches but the amount differs
    # beyond tolerance) is what trips a header 'value' flag, so it must be visible on
    # the Affected tab, not hidden behind a green qty 'OK'.
    def _affected(ln):
        return ln.get('status') in _LINE_PROBLEM or ln.get('val_ok') == 'NO'
    ext_pos = {h.get('po') for h in headers if h.get('status') == 'EXTERNAL'}
    for h in headers:
        h['lines'] = by_po.get(h.get('po'), [])
        # EXTERNAL POs (GT Select etc.) are beyond our cross-check → their lines are
        # never "affected"; don't show a red line-count badge on them.
        h['bad_lines'] = 0 if h.get('status') == 'EXTERNAL' else sum(1 for ln in h['lines'] if _affected(ln))
        # An empty D365 SO shell — ZERO qty AND value (blank customer/pincode too).
        # Nothing to reconcile or capture → struck through so it reads as skippable.
        # NB: test qty+value, NOT lines — a GT Select external legitimately has qty/
        # value but NO associated lines in the reconcile (its header keys on External
        # Doc while its lines key on the SO No), so it must NOT be treated as empty.
        h['is_empty'] = not (h.get('d365_qty') or 0) and not (h.get('d365_val') or 0)
    data['orders'] = [h for h in headers if h.get('status') != 'EXTERNAL']
    data['externals'] = [h for h in headers if h.get('status') == 'EXTERNAL']
    # EXTERNAL POs are beyond comparison — keep their lines OFF the affected tab.
    data['affected_lines'] = [ln for ln in lines if _affected(ln) and ln.get('po') not in ext_pos]
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
        ctx['saved_runs'] = _saved_runs()   # the checked-PO log now lives on its own page
        token = self.request.GET.get('token', '')
        ctx['token'] = token
        rp = _result_path(token) if token else None
        if rp and rp.exists():
            res = json.loads(rp.read_text(encoding='utf-8'))
            if res.get('ok'):
                ctx['data'] = _augment(res['data'])
                ctx['confirmed'] = res.get('confirmed', False)
                ctx['capture'] = res.get('capture')          # external-order capture panel
                ctx['captured'] = res.get('captured', False)
                # per-tab discard states (top-bar Discard deletes the whole check)
                ctx['import_discarded'] = res.get('import_discarded', False)
                ctx['verify_discarded'] = res.get('verify_discarded', False)
                ctx['taxonomy'] = mkt.classification_options()
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
            # ── Capture pass: classify the EXTERNAL (in-D365, not-recorded) orders
            #    by Gen. Bus. Posting Group + dedup vs everything recorded, so the
            #    operator can RECORD them into the tracker via the separate Capture
            #    action. Files persist under the token → the capture view re-reads.
            res['headers_path'] = headers
            res['lines_path'] = lines
            try:
                cap = gts.preview(headers, lines)
                if cap.get('ok'):
                    new_orders = [{
                        'so_no': h['so_no'], 'external_doc': h['external_doc'],
                        'marketplace': h['marketplace'], 'segment': h['segment'],
                        'posting_group': h['posting_group'],
                        # normalized posting-group key (same _norm_pg the classify rows
                        # use) so the UI can gate ONLY on groups with a selected order
                        'class_key': gts._norm_pg(h['posting_group']),
                        'ship_name': h['ship_name'],
                        'customer_name': h['customer_name'], 'cust_no': h['cust_no'],
                        'ship_code': h['ship_code'], 'postcode': h['postcode'],
                        'warehouse': h['warehouse'], 'line_count': h['line_count'],
                        'qty': h['qty'], 'order_value': h['order_value'],
                        'po_date': h['po_date'].isoformat() if h['po_date'] else '',
                    } for h in cap['headers'] if h['is_new']]
                    # NB: do NOT cap new_orders — every new order needs a push
                    # checkbox, else Select-all silently under-captures the tail
                    # (only_pos would miss them and they'd be mislabelled 'already
                    # present'). The list is bounded by one D365 dump's new orders.
                    res['capture'] = {'summary': cap['summary'], 'channels': cap['channels'],
                                      'needs_class': cap['needs_class'], 'new_orders': new_orders}
                    # Externals show their TRUE channel (from Gen. Bus. Posting
                    # Group) instead of 'UNKNOWN' — the D365 file classifies them
                    # even when they aren't in our records yet.
                    cls = {}
                    for hh in cap['headers']:
                        for k in (hh['external_doc'], hh['so_no']):
                            if k:
                                cls[str(k).upper()] = hh['marketplace']
                    for row in res['data'].get('headers', []):
                        if row.get('status') == 'EXTERNAL':
                            m = cls.get(str(row.get('po') or '').upper())
                            if m:
                                row['mp'] = m
            except Exception as e:  # noqa: BLE001 — capture is additive; never blocks verify
                res['capture_error'] = f'{type(e).__name__}: {e}'
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


class RecordVerificationCaptureView(LoginRequiredMixin, View):
    """Record the EXTERNAL (in-D365, not-recorded) orders into the TRACKER
    (order_headers / order_lines) via the classify + dedup importer — the separate
    'Capture' action. Distinct from the verification-log confirm. AJAX → JSON;
    ``overrides`` places unknown posting groups (segment / marketplace / child)."""

    def post(self, request, token):
        rp, res = _load_result(token)
        if not res.get('ok') or not res.get('headers_path'):
            return JsonResponse({'ok': False, 'error': 'Nothing to capture for this check.'}, status=400)
        overrides, only_pos = {}, None
        try:
            body = json.loads(request.body or '{}')
            if isinstance(body, dict):
                overrides = body.get('overrides') or {}
                only_pos = body.get('only_pos')      # None = push every new order
        except (ValueError, TypeError):
            pass
        out = gts.do_import(res['headers_path'], res['lines_path'],
                            overrides=overrides, user=_actor(request), only_pos=only_pos)
        if not out.get('ok'):
            return JsonResponse({'ok': False, 'error': out.get('error', 'Capture failed.')}, status=400)
        res['captured'] = True
        res['capture_result'] = {'imported': out['imported'], 'lines': out['lines'],
                                 'skipped': out.get('skipped', 0)}
        _save_result(rp, res)
        return JsonResponse({'ok': True, 'imported': out['imported'], 'lines': out['lines'],
                             'skipped': out.get('skipped', 0), 'redirect': reverse('b2b_tracker')})


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


class RecordVerificationLogView(LoginRequiredMixin, TemplateView):
    """Separate 'Checked POs' history page — the persisted verification log, split
    into All / Verified / Mismatch / External tabs. Read-only; the main page keeps
    just upload + the current run's result."""
    template_name = 'online_b2b/record_verification_log.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        log = rv.checked_log(limit=3000)
        ctx['log'] = log
        ctx['coverage'] = rv.coverage()
        ctx['ok'] = [r for r in log if r.get('status') == 'OK']
        ctx['mismatch'] = [r for r in log if r.get('status') not in ('OK', 'EXTERNAL')]
        ctx['external'] = [r for r in log if r.get('status') == 'EXTERNAL']
        return ctx


class RecordVerificationClearLogView(LoginRequiredMixin, View):
    """Clear the whole checked-PO log (start fresh). Recorded orders are untouched."""

    def post(self, request):
        out = rv.clear_log()
        msg = f"Cleared the verification log - {out.get('deleted', 0)} entry(ies) removed."
        if common.is_ajax(request):
            return JsonResponse({'ok': True, 'message': msg,
                                 'redirect': reverse('b2b_record_verify_log')})
        messages.info(request, msg)
        return redirect('b2b_record_verify_log')


class RecordVerificationDiscardView(LoginRequiredMixin, View):
    """Discard a check — delete its token dir (result + uploaded files). Nothing was
    recorded (only Confirm records), so this is pure cleanup; returns to a clean page."""

    def post(self, request, token):
        d = _tok_dir(token)
        # Discard is a run-level action (top of the page): drop the WHOLE check —
        # both the Import and Verify tabs — and return to a clean upload. The UI
        # confirm warns when imports are still pending, so this is never a surprise.
        if d.exists() and d.resolve() != _UP.resolve():
            shutil.rmtree(d, ignore_errors=True)
        msg = 'Discarded — the check was removed. Nothing was recorded.'
        if common.is_ajax(request):
            return JsonResponse({'ok': True, 'message': msg,
                                 'redirect': reverse('b2b_record_verify')})
        messages.info(request, msg)
        return redirect('b2b_record_verify')


class RecordVerificationDiscardPartView(LoginRequiredMixin, View):
    """Discard ONE tab's work — the import OR the verification — leaving the other
    tab intact. The top-bar Discard drops the whole check; this is the per-tab one.
    POST ``part=import|verify``. Keeps the token (nothing is deleted, just flagged)."""

    def post(self, request, token):
        rp, res = _load_result(token)
        part = (request.POST.get('part') or '').strip().lower()
        if part == 'import':
            res['import_discarded'] = True
            msg = 'Import discarded — no new orders captured. Verification is untouched.'
        elif part == 'verify':
            res['verify_discarded'] = True
            msg = 'Verification discarded — nothing recorded. Your import is untouched.'
        else:
            return JsonResponse({'ok': False, 'error': 'Unknown part to discard.'}, status=400)
        _save_result(rp, res)
        return _done(request, token, msg, redirect=_token_url(token))


class RecordVerificationDownloadView(LoginRequiredMixin, View):
    """Serve the reconciliation Excel for a run."""

    def get(self, request, token):
        xp = _tok_dir(token) / 'record_verification.xlsx'
        if not xp.exists():
            raise Http404('File not found or expired.')
        return FileResponse(open(xp, 'rb'), as_attachment=True,
                            filename='Record_Verification.xlsx')
