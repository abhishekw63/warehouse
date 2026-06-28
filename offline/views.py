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
from django.contrib.auth.mixins import LoginRequiredMixin
from django.http import FileResponse, Http404, HttpResponse, JsonResponse
from django.urls import reverse
from django.views.generic import TemplateView, View

from .services import gt_mass_bridge, mt_bridge
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

class ShoppersStopView(LoginRequiredMixin, TemplateView):
    """Upload + generate page for the MT-Select Shoppers Stop channel."""
    template_name = 'offline/shoppers_stop.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['channels'] = mt_bridge.channel_choices()
        ctx['warehouses'] = mt_bridge.warehouse_choices()
        ctx['default_warehouse'] = mt_bridge.default_warehouse()
        return ctx


class SSPreviewView(LoginRequiredMixin, View):
    """Phase 1: save uploaded PO file(s) under a token, run a NO-WRITE preview
    (parse + resolve + validate). Nothing is recorded and no SO number is burned —
    same as the online review step."""

    def post(self, request, *args, **kwargs):
        files = request.FILES.getlist('files')
        if not files:
            return JsonResponse({'ok': False, 'error': 'No files selected'},
                                status=400)
        channel = request.POST.get('channel', 'SS')
        warehouse = (request.POST.get('warehouse', '')
                     or mt_bridge.default_warehouse())

        token = uuid.uuid4().hex[:12]
        up_dir = Path(settings.MEDIA_ROOT) / 'mt_uploads' / token
        up_dir.mkdir(parents=True, exist_ok=True)
        paths = []
        for f in files:
            dest = up_dir / Path(f.name).name
            with open(dest, 'wb') as out:
                for chunk in f.chunks():
                    out.write(chunk)
            paths.append(str(dest))
        (up_dir / 'meta.json').write_text(
            json.dumps({'channel': channel, 'warehouse': warehouse,
                        'files': paths}), encoding='utf-8')

        result = mt_bridge.preview(channel, paths, warehouse)
        result['token'] = token
        return JsonResponse(result, status=200 if result.get('ok') else 400)


class SSConfirmView(LoginRequiredMixin, View):
    """Phase 2: assign SO numbers, write the workbook, and record into the shared
    renee_orders DB — so SS appears on the online dashboard."""

    def post(self, request, *args, **kwargs):
        token = request.POST.get('token', '')
        up_dir = Path(settings.MEDIA_ROOT) / 'mt_uploads' / token
        meta_p = up_dir / 'meta.json'
        if not token or not meta_p.exists():
            return JsonResponse(
                {'ok': False, 'error': 'Upload expired — please re-upload.'},
                status=400)
        meta = json.loads(meta_p.read_text(encoding='utf-8'))
        result = mt_bridge.confirm(meta['channel'], meta['files'],
                                   meta['warehouse'])
        if result.get('ok') and result.get('output_path'):
            request.session[f'ss_out_{token}'] = result['output_path']
            result['download_url'] = reverse('ss_download', args=[token])
        return JsonResponse(result, status=200 if result.get('ok') else 400)


class SSDownloadView(LoginRequiredMixin, View):
    """GET: serve the generated workbook for a processed token."""

    def get(self, request, token, *args, **kwargs):
        path = request.session.get(f'ss_out_{token}')
        if not path or not os.path.exists(path):
            raise Http404('Generated workbook not found or expired.')
        return FileResponse(open(path, 'rb'), as_attachment=True,
                            filename=os.path.basename(path))


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
            raise Http404('Generated dump not found or expired.')
        return FileResponse(open(path, 'rb'), as_attachment=True,
                            filename=os.path.basename(path))