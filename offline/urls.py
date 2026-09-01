from django.urls import path

from . import views

urlpatterns = [
    path('', views.OfflineDashboardView.as_view(), name='offline_dashboard'),
    path('gt-mass-dump/', views.IndexView.as_view(), name='index'),
    path('process/', views.ProcessFilesView.as_view(), name='process_files'),
    path('export-d365/', views.ExportD365View.as_view(), name='export_d365'),
    path('send-email/', views.SendEmailView.as_view(), name='send_email'),
    path('download-template/', views.DownloadTemplateView.as_view(), name='download_template'),
    path('gt-mass/', views.GTMassRecorderView.as_view(), name='gt_mass_recorder'),
    path('gt-mass/preview/', views.GTMPreviewView.as_view(), name='gtm_preview'),
    path('gt-mass/confirm/', views.GTMConfirmView.as_view(), name='gtm_confirm'),
    path('gt-mass/download/<str:token>/', views.GTMDownloadView.as_view(), name='gtm_download'),
    # GT Mass on the shared PO-flow scaffold (upload → review → confirm), the new
    # standard reused across segments. The single-page recorder above stays as a
    # fallback.
    path('gt-mass-flow/', views.GTMFlowUploadView.as_view(), name='gtm_flow_upload'),
    # 'drafts' list — MUST precede the <str:token> review route (else captured as a token).
    path('gt-mass-flow/drafts/', views.GTMFlowDraftsView.as_view(), name='gtm_flow_drafts'),
    path('gt-mass-flow/<str:token>/', views.GTMFlowReviewView.as_view(), name='gtm_flow_review'),
    path('gt-mass-flow/<str:token>/confirm/', views.GTMFlowConfirmView.as_view(), name='gtm_flow_confirm'),
    path('gt-mass-flow/<str:token>/decision/', views.GTMFlowDecisionView.as_view(), name='gtm_flow_decision'),
    path('gt-mass-flow/<str:token>/discard/', views.GTMFlowDiscardView.as_view(), name='gtm_flow_discard'),
    path('gt-mass-flow/<str:token>/download/', views.GTMFlowDownloadView.as_view(), name='gtm_flow_download'),
    path('gt-mass-flow/<str:token>/export/', views.GTMFlowExportView.as_view(), name='gtm_flow_export'),
    path('gt-mass-flow/<str:token>/save-later/', views.GTMFlowSaveLaterView.as_view(), name='gtm_flow_save_later'),
]
