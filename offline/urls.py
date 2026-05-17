from django.urls import path
from . import views

urlpatterns = [
    path('', views.OfflineDashboardView.as_view(), name='offline_dashboard'),
    path('gt-mass-dump/', views.IndexView.as_view(), name='index'),
    path('process/', views.ProcessFilesView.as_view(), name='process_files'),
    path('export-d365/', views.ExportD365View.as_view(), name='export_d365'),
    path('send-email/', views.SendEmailView.as_view(), name='send_email'),
    path('download-template/', views.DownloadTemplateView.as_view(), name='download_template'),
]