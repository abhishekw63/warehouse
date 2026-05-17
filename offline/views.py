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

from django.views.generic import TemplateView, View
from django.contrib.auth.mixins import LoginRequiredMixin
from django.http import HttpResponse, JsonResponse
from datetime import datetime

from .utils import (
    GTMassAutomation,
    D365Exporter,
    EmailSender,
    EMAIL_CONFIG,
    TemplateGenerator,
    result_to_session,
    result_from_session,
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