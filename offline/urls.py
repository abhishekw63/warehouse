from django.urls import path
from . import views

urlpatterns = [
    path('', views.OfflineDashboardView.as_view(), name='offline_dashboard'),
    path('gt-mass-dump/', views.IndexView.as_view(), name='index'),
    path('process/', views.ProcessFilesView.as_view(), name='process_files'),
    path('export-d365/', views.ExportD365View.as_view(), name='export_d365'),
    path('send-email/', views.SendEmailView.as_view(), name='send_email'),
    path('download-template/', views.DownloadTemplateView.as_view(), name='download_template'),
    # MT Select — Shoppers Stop (preview → confirm → records to renee_orders)
    path('shoppers-stop/', views.ShoppersStopView.as_view(), name='shoppers_stop'),
    path('shoppers-stop/preview/', views.SSPreviewView.as_view(), name='ss_preview'),
    path('shoppers-stop/confirm/', views.SSConfirmView.as_view(), name='ss_confirm'),
    path('shoppers-stop/download/<str:token>/', views.SSDownloadView.as_view(), name='ss_download'),
]