"""
Django views for the Offline / GT Mass Dump Generator.

Routes:
    GET  /offline/                    → OfflineDashboardView
    GET  /offline/gt-mass-dump/       → IndexView (upload form)
    POST /offline/process/            → ProcessFilesView (generate dump)
    POST /offline/export-d365/        → ExportD365View (fill D365 template)
    POST /offline/send-email/         → SendEmailView (send HTML report)
    GET  /offline/download-template/  → DownloadTemplateView (blank PO template)
"""

import json
import os
import uuid
from datetime import datetime
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin
from django.http import FileResponse, Http404, HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.generic import TemplateView, View

from online_b2b.services import po_flow

from .flows import GT_MASS_SPEC
from .services import gt_mass_bridge
from .utils import (
    EMAIL_CONFIG,
    D365Exporter,
    EmailSender,
    GTMassAutomation,
    TemplateGenerator,
    result_from_session,
    result_to_session,
)


class OfflineDashboardView(LoginRequiredMixin, TemplateView):
    """Department landing page."""
    template_name = 'offline/dashboard.html'


class IndexView(LoginRequiredMixin, TemplateView):
    """Upload form for GT Mass Dump Generator."""
    template_name = 'offline/index.html'


class ProcessFilesView(LoginRequiredMixin, View):
    """
    POST /offline/process/
    Process uploaded files → generate 7-sheet dump → store result in session.

    On success: returns the Excel file as attachment + stats in custom headers.
    On failure: returns JSON with error details.
    """

    def post(self, request, *args, **kwargs):
        files = request.FILES.getlist('files')

        if not files:
            return JsonResponse({"error": "No files selected"}, status=400)

        # Process
        automation = GTMassAutomation()
        result = automation.process_files(files)

        # Store in session for D365 export / email (Option A)
        request.session['gt_mass_result'] = result_to_session(result)
        request.session['gt_mass_elapsed'] = '(web)'

        # Export to memory
        output = automation.exporter.export_to_memory(result)

        if output is None:
            return JsonResponse({
                "error": "No valid data found in selected files",
                "details": {
                    "attempted": len(result.attempted_files),
                    "failed": len(result.failed_files),
                    "failures": [
                        {"file": f, "reason": r}
                        for f, r in result.failed_files
                    ],
                },
            }, status=400)

        # Build response
        today = datetime.now().strftime("%d%m%Y")
        filename = f"gt_mass_dump_{today}.xlsx"

        unique_sos = len({r.so_number for r in result.rows})
        missing_loc = len({r.so_number for r in result.rows if not r.location_code})

        response = HttpResponse(
            output.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'

        # Stats headers for frontend
        response['X-GT-Attempted'] = str(len(result.attempted_files))
        response['X-GT-Rows'] = str(len(result.rows))
        response['X-GT-SOs'] = str(unique_sos)
        response['X-GT-Failed'] = str(len(result.failed_files))
        response['X-GT-Warnings'] = str(len(result.warned_files))
        response['X-GT-MissingLocation'] = str(missing_loc)
        response['Access-Control-Expose-Headers'] = (
            'X-GT-Attempted, X-GT-Rows, X-GT-SOs, '
            'X-GT-Failed, X-GT-Warnings, X-GT-MissingLocation'
        )

        return response


class ExportD365View(LoginRequiredMixin, View):
    """
    POST /offline/export-d365/
    Upload D365 template → fill with stored result → return filled file.

    Requires:
        - 'template' file in request.FILES
        - 'gt_mass_result' in session (from prior ProcessFilesView call)
    """

    def post(self, request, *args, **kwargs):
        # Check session for stored result
        session_data = request.session.get('gt_mass_result')

        if not session_data:
            return JsonResponse(
                {"error": "No processed data found. Generate the dump first."},
                status=400,
            )

        template_file = request.FILES.get('template')

        if not template_file:
            return JsonResponse(
                {"error": "No D365 template file uploaded."},
                status=400,
            )

        # Restore result from session
        result = result_from_session(session_data)

        if not result.rows:
            return JsonResponse(
                {"error": "Stored result has no data rows."},
                status=400,
            )

        try:
            output = D365Exporter.export(result, template_file)
        except (ValueError, RuntimeError) as e:
            return JsonResponse({"error": str(e)}, status=400)
        except Exception as e:
            return JsonResponse({"error": f"D365 export failed: {e}"}, status=500)

        today = datetime.now().strftime("%d%m%Y")
        filename = f"d365_import_{today}.xlsx"

        response = HttpResponse(
            output.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'

        # Stats for frontend
        response['X-D365-SOs'] = str(len({r.so_number for r in result.rows}))
        response['X-D365-Items'] = str(len(result.rows))
        response['Access-Control-Expose-Headers'] = 'X-D365-SOs, X-D365-Items'

        return response


class SendEmailView(LoginRequiredMixin, View):
    """
    POST /offline/send-email/
    Send the HTML email report using stored result from session.
    """

    def post(self, request, *args, **kwargs):
        session_data = request.session.get('gt_mass_result')

        if not session_data:
            return JsonResponse(
                {"error": "No processed data found. Generate the dump first."},
                status=400,
            )

        result = result_from_session(session_data)

        if not result.rows:
            return JsonResponse(
                {"error": "Stored result has no data rows."},
                status=400,
            )

        elapsed = request.session.get('gt_mass_elapsed', '')
        success, error_msg = EmailSender.send_report(result, elapsed)

        if success:
            cc_list = ', '.join(EMAIL_CONFIG['CC_RECIPIENTS']) or 'none'
            return JsonResponse({
                "success": True,
                "to": EMAIL_CONFIG['DEFAULT_RECIPIENT'],
                "cc": cc_list,
            })
        else:
            return JsonResponse(
                {"error": f"Email failed: {error_msg}"},
                status=500,
            )


class DownloadTemplateView(LoginRequiredMixin, View):
    """
    GET /offline/download-template/
    Generate and return a blank GT-Mass PO template.
    """

    def get(self, request, *args, **kwargs):
        output = TemplateGenerator.generate()

        response = HttpResponse(
            output.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        response['Content-Disposition'] = 'attachment; filename="GT-Mass_PO_Template.xlsx"'

        return response


# ─────────────────────────────────────────────────────────────────────────
#  MT Select — Shoppers Stop (and future MT channels)
#
#  Wraps the FROZEN desktop automation via offline.services.mt_bridge, which
#  runs the EXACT desktop pipeline headlessly → same ss_so_*.xlsx workbook.
# ─────────────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────────────
#  GT Mass — Dashboard recorder (preview → confirm → record to renee_orders)
#
#  ADDITIVE: the existing "GT Mass Dump Generator" page (IndexView /
#  ProcessFilesView) and the frozen Tkinter standalone are untouched and remain
#  the fallback. This flow records GT Mass into the shared dashboard (Orders +
#  Line Items) with real value read from each file's own TOTAL column, via
#  offline.services.gt_mass_bridge.
# ─────────────────────────────────────────────────────────────────────────

class GTMassRecorderView(LoginRequiredMixin, TemplateView):
    """Upload + preview/confirm page that records GT Mass to the dashboard."""
    template_name = 'offline/gt_mass_recorder.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['warehouses'] = gt_mass_bridge.warehouse_choices()
        ctx['default_warehouse'] = gt_mass_bridge.default_warehouse()
        return ctx


class GTMPreviewView(LoginRequiredMixin, View):
    """Phase 1: save uploaded file(s) under a token, run a NO-WRITE preview
    (parse + price + per-SO summary). Nothing recorded."""

    def post(self, request, *args, **kwargs):
        files = request.FILES.getlist('files')
        if not files:
            return JsonResponse({'ok': False, 'error': 'No files selected'},
                                status=400)
        warehouse = (request.POST.get('warehouse', '')
                     or gt_mass_bridge.default_warehouse())
        token = uuid.uuid4().hex[:12]
        up_dir = Path(settings.MEDIA_ROOT) / 'gt_mass_uploads' / token
        up_dir.mkdir(parents=True, exist_ok=True)
        paths = []
        for f in files:
            dest = up_dir / Path(f.name).name
            with open(dest, 'wb') as out:
                for chunk in f.chunks():
                    out.write(chunk)
            paths.append(str(dest))
        (up_dir / 'meta.json').write_text(
            json.dumps({'warehouse': warehouse, 'files': paths}),
            encoding='utf-8')
        result = gt_mass_bridge.preview(paths, warehouse)
        result['token'] = token
        return JsonResponse(result, status=200 if result.get('ok') else 400)


class GTMConfirmView(LoginRequiredMixin, View):
    """Phase 2: record runs + order_headers + order_lines into renee_orders
    (dedup) and produce the 7-sheet dump for download."""

    def post(self, request, *args, **kwargs):
        token = request.POST.get('token', '')
        up_dir = Path(settings.MEDIA_ROOT) / 'gt_mass_uploads' / token
        meta_p = up_dir / 'meta.json'
        if not token or not meta_p.exists():
            return JsonResponse(
                {'ok': False, 'error': 'Upload expired — please re-upload.'},
                status=400)
        meta = json.loads(meta_p.read_text(encoding='utf-8'))
        result = gt_mass_bridge.confirm(meta['files'], meta['warehouse'])
        if result.get('ok') and result.get('output_path'):
            request.session[f'gtm_out_{token}'] = result['output_path']
            result['download_url'] = reverse('gtm_download', args=[token])
        return JsonResponse(result, status=200 if result.get('ok') else 400)


class GTMDownloadView(LoginRequiredMixin, View):
    """GET: serve the generated dump workbook for a processed token."""

    def get(self, request, token, *args, **kwargs):
        path = request.session.get(f'gtm_out_{token}')
        if not path or not os.path.exists(path):
            messages.info(request, "That dump has expired — re-run GT Mass to regenerate it.")
            return redirect('gt_mass_recorder')
        return FileResponse(open(path, 'rb'), as_attachment=True,
                            filename=os.path.basename(path))


# ── Reliance Trends (BAP Excel) recorder — upload → preview → confirm ─────────
#   Records to renee_orders like GT Mass; frozen engine untouched. NEW channel
#   (cust 20418); BAP replen PO ships to Bhiwandi 20418_2 (S0HZ ambiguity noted).


# ── Shared PO-flow scaffold (upload → review → confirm → lock) ───────────────
# Generic, spec-driven CBVs: all logic lives in online_b2b.services.po_flow.
# A new offline channel = a processor adapter + a FlowSpec + a 6-line block of
# subclasses that just set ``spec`` (see GT Mass and MT below). The Tkinter-style
# single-page recorders above stay untouched as fallbacks.


def _channel_reqs(spec) -> dict:
    """{channel code: requirements descriptor} for the upload-page hint, so the
    operator sees what each channel demands (and the if-absent behaviour) BEFORE
    uploading. MT channels only; empty for other segments."""
    try:
        from .services import mt_bridge
    except Exception:  # noqa: BLE001
        return {}
    out = {}
    for code, _label in (spec.marketplaces or ()):
        req = mt_bridge.channel_requirements(code)
        if req:
            out[code] = req
    return out


def _flow_upload_ctx(spec):
    return {
        'spec': spec, 'title': spec.title, 'segment': spec.segment,
        'base_template': spec.base_template, 'intro': spec.intro,
        'caps': spec.caps_map(), 'slots': spec.slot_map(),
        'warehouses': spec.warehouses,
        'marketplaces': spec.marketplaces, 'default_margin': spec.default_margin,
        'accept': spec.accept,
        'channel_reqs': _channel_reqs(spec),
        'u_upload': spec.urls['upload'], 'u_back': spec.urls['back'],
        'u_dashboard': spec.urls['dashboard'],
        # 'Review Later' Drafts list (only if the spec wired it) + a live count.
        'u_drafts': spec.urls.get('drafts'),
        'n_drafts': (len(po_flow.collect_drafts(spec))
                     if spec.urls.get('drafts') else 0),
    }


def _is_ajax(request):
    return request.headers.get('x-requested-with') == 'XMLHttpRequest'


class _FlowUploadView(LoginRequiredMixin, View):
    """Generic upload view — subclasses set ``spec``."""
    spec = None

    def get(self, request):
        return render(request, 'po_flow/upload.html', _flow_upload_ctx(self.spec))

    def post(self, request):
        spec = self.spec
        files = request.FILES.getlist('po_files')
        if not files:
            return JsonResponse({'ok': False, 'error': 'Choose at least one file.'})
        extra = {}
        if 'warehouse' in spec.caps:
            extra['warehouse'] = (request.POST.get('warehouse')
                                  or (spec.warehouses[0][0] if spec.warehouses else ''))
        if 'margin' in spec.caps:
            extra['margin_pct'] = request.POST.get('margin_pct') or spec.default_margin
        if 'marketplace' in spec.caps:
            extra['marketplace'] = request.POST.get('marketplace')
        token = po_flow.save_upload(spec, files, extra)
        meta = po_flow.load_meta(spec, token)
        payload = po_flow.preview(spec, token, meta)
        s = payload.get('summary', {})
        return JsonResponse({
            'ok': True, 'review_url': reverse(spec.urls['review'], args=[token]),
            'pos': s.get('pos', 0), 'lines': s.get('lines', 0),
            'affected': s.get('affected', 0),
            'issues': len(payload.get('file_issues', [])),
            'warnings': len(payload.get('warnings', [])),
        })


class _FlowReviewView(LoginRequiredMixin, View):
    spec = None

    def get(self, request, token):
        meta = po_flow.load_meta(self.spec, token)
        if meta is None:
            messages.info(request, "That upload was already recorded or has expired — nothing to review.")
            return redirect(reverse(self.spec.urls['upload']))
        # Reopen-from-Drafts (?revalidate=1): drop the cached preview so the run
        # is re-validated live against the CURRENT masters (picks up any fix made
        # while it was parked). No-op once locked.
        if request.GET.get('revalidate') and not meta.get('locked'):
            po_flow.invalidate(self.spec, token)
        return render(request, 'po_flow/review.html',
                      po_flow.review_context(self.spec, token, meta))


class _FlowConfirmView(LoginRequiredMixin, View):
    spec = None

    def post(self, request, token):
        spec = self.spec
        meta = po_flow.load_meta(spec, token)
        if meta is None:
            return JsonResponse({'ok': False, 'error': 'Upload expired.'}, status=404)
        review_url = reverse(spec.urls['review'], args=[token])
        if meta.get('locked'):
            return JsonResponse({'ok': True, 'run_id': meta.get('run_id'),
                                 'review_url': review_url, 'already': True})
        # Record ONLY on an explicit Confirm click (the AJAX button sends
        # confirm=1). A stray / native / implicit form submit must never record.
        if request.POST.get('confirm') != '1':
            if _is_ajax(request):
                return JsonResponse(
                    {'ok': False, 'error': 'Confirm intent missing — click '
                     'Confirm & Record.'}, status=400)
            return redirect(review_url)
        result = po_flow.confirm(spec, token, meta)
        result['review_url'] = review_url
        if not _is_ajax(request):
            return redirect(review_url)
        return JsonResponse(result)


class _FlowDecisionView(LoginRequiredMixin, View):
    spec = None

    def post(self, request, token):
        n = po_flow.set_decision(self.spec, token, request.POST.get('key', ''),
                                 request.POST.get('action', ''),
                                 request.POST.get('override_cp', ''),
                                 request.POST.get('remark', ''))
        return JsonResponse({'ok': True, 'saved': n})


class _FlowDiscardView(LoginRequiredMixin, View):
    spec = None

    def post(self, request, token):
        po_flow.discard(self.spec, token)
        return redirect(reverse(self.spec.urls['dashboard']))


class _FlowDownloadView(LoginRequiredMixin, View):
    spec = None

    def get(self, request, token):
        p = po_flow.download_path(self.spec, token)
        if not p:
            messages.info(request, "That workbook isn't available yet or the upload has expired.")
            return redirect(reverse(self.spec.urls['review'], args=[token]))
        # Uniform, self-describing name on par with Online B2B (Mp_Npo_ts_kind) —
        # never the raw tmpXXXX temp name.
        meta = po_flow.load_meta(self.spec, token) or {}
        return FileResponse(open(p, 'rb'), as_attachment=True,
                            filename=po_flow.download_name(self.spec, meta, p))


class _FlowExportView(LoginRequiredMixin, View):
    """Download the review data (Orders + Line items) as a plain Excel — NO SO
    numbers — for eyeballing before Confirm. Available pre- and post-lock."""
    spec = None

    def get(self, request, token):
        meta = po_flow.load_meta(self.spec, token)
        if meta is None:
            messages.info(request, "That upload has expired — the export is no longer available.")
            return redirect(reverse(self.spec.urls['upload']))
        p = po_flow.export_review_xlsx(self.spec, token, meta)
        if not p:
            raise Http404('Could not build the review export.')
        return FileResponse(open(p, 'rb'), as_attachment=True, filename=p.name)


class _FlowSaveLaterView(LoginRequiredMixin, View):
    """Park the WHOLE run as a 'Review Later' draft (kept intact, NOT recorded).
    AJAX → ``{ok, redirect}``; native submit → redirect to Drafts."""
    spec = None

    def post(self, request, token):
        spec = self.spec
        ok = po_flow.save_draft(spec, token, request.POST.get('note', ''))
        drafts_url = reverse(spec.urls['drafts'])
        if not ok:
            if _is_ajax(request):
                return JsonResponse(
                    {'ok': False,
                     'error': 'Already recorded or expired — nothing to defer.'},
                    status=400)
            return redirect(reverse(spec.urls['review'], args=[token]))
        if _is_ajax(request):
            return JsonResponse({'ok': True, 'redirect': drafts_url})
        return redirect(drafts_url)


class _FlowDraftsView(LoginRequiredMixin, View):
    """List all parked 'Review Later' runs for this flow — reopen (→ re-validate)
    or discard from here."""
    spec = None

    def get(self, request):
        spec = self.spec
        return render(request, 'po_flow/drafts.html', {
            'spec': spec,
            'title': spec.title,
            'segment': spec.segment,
            'base_template': spec.base_template,
            'drafts': po_flow.collect_drafts(spec),
            'u_review': spec.urls['review'],
            'u_discard': spec.urls['discard'],
            'u_upload': spec.urls['upload'],
            'u_dashboard': spec.urls['dashboard'],
        })


# ── GT Mass (unchanged behaviour — now on the generic base) ──────────────
class GTMFlowUploadView(_FlowUploadView):
    spec = GT_MASS_SPEC


class GTMFlowReviewView(_FlowReviewView):
    spec = GT_MASS_SPEC


class GTMFlowConfirmView(_FlowConfirmView):
    spec = GT_MASS_SPEC


class GTMFlowDecisionView(_FlowDecisionView):
    spec = GT_MASS_SPEC


class GTMFlowDiscardView(_FlowDiscardView):
    spec = GT_MASS_SPEC


class GTMFlowDownloadView(_FlowDownloadView):
    spec = GT_MASS_SPEC


class GTMFlowExportView(_FlowExportView):
    spec = GT_MASS_SPEC


class GTMFlowSaveLaterView(_FlowSaveLaterView):
    spec = GT_MASS_SPEC


class GTMFlowDraftsView(_FlowDraftsView):
    spec = GT_MASS_SPEC


# ── Modern Trade (MT) — on par with the online marketplaces ──────────────


# ── EKA (EBO / Kiosk / Airport → SO/TO) — third offline channel on po_flow ───

