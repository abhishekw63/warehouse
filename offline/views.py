from django.shortcuts import render
from django.views.generic import TemplateView, View

class IndexView(TemplateView):
    template_name = 'offline/index.html'


from django.http import HttpResponse, JsonResponse
from .utils import GTMassAutomation
from datetime import datetime

class ProcessFilesView(View):
    def post(self, request, *args, **kwargs):
        files = request.FILES.getlist('files')

        if not files:
            return JsonResponse({"error": "No files selected"}, status=400)

        automation = GTMassAutomation()
        rows = automation.process_files(files)

        output_buffer = automation.exporter.export_to_memory(rows)

        if output_buffer is None:
             return JsonResponse({"error": "No valid data found in selected files"}, status=400)

        today = datetime.now().strftime("%d%m%Y")
        filename = f"gt_mass_dump_{today}.xlsx"

        response = HttpResponse(
            output_buffer.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
