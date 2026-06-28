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
    # GT Mass — Dashboard recorder (preview → confirm → records to renee_orders).
    # The dump generator above (gt-mass-dump/) stays as the untouched fallback.
    path('gt-mass/', views.GTMassRecorderView.as_view(), name='gt_mass_recorder'),
    path('gt-mass/preview/', views.GTMPreviewView.as_view(), name='gtm_preview'),
    path('gt-mass/confirm/', views.GTMConfirmView.as_view(), name='gtm_confirm'),
    path('gt-mass/download/<str:token>/', views.GTMDownloadView.as_view(), name='gtm_download'),
    # GT Mass on the shared PO-flow scaffold (upload → review → confirm), the new
    # standard reused across segments. The single-page recorder above stays as a
    # fallback.
    path('gt-mass-flow/', views.GTMFlowUploadView.as_view(), name='gtm_flow_upload'),
    path('gt-mass-flow/<str:token>/', views.GTMFlowReviewView.as_view(), name='gtm_flow_review'),
    path('gt-mass-flow/<str:token>/confirm/', views.GTMFlowConfirmView.as_view(), name='gtm_flow_confirm'),
    path('gt-mass-flow/<str:token>/decision/', views.GTMFlowDecisionView.as_view(), name='gtm_flow_decision'),
    path('gt-mass-flow/<str:token>/discard/', views.GTMFlowDiscardView.as_view(), name='gtm_flow_discard'),
    path('gt-mass-flow/<str:token>/download/', views.GTMFlowDownloadView.as_view(), name='gtm_flow_download'),
]