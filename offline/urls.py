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
    # Reliance Trends — BAP Excel recorder (preview → confirm → records to renee_orders).
    path('reliance-trends/', views.RelianceTrendsView.as_view(), name='reliance_trends'),
    path('reliance-trends/preview/', views.RTPreviewView.as_view(), name='rt_preview'),
    path('reliance-trends/confirm/', views.RTConfirmView.as_view(), name='rt_confirm'),
    path('reliance-trends/download/<str:token>/', views.RTDownloadView.as_view(), name='rt_download'),
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
    # 'drafts' list — MUST precede the <str:token> review route (else captured as a token).
    path('gt-mass-flow/drafts/', views.GTMFlowDraftsView.as_view(), name='gtm_flow_drafts'),
    path('gt-mass-flow/<str:token>/', views.GTMFlowReviewView.as_view(), name='gtm_flow_review'),
    path('gt-mass-flow/<str:token>/confirm/', views.GTMFlowConfirmView.as_view(), name='gtm_flow_confirm'),
    path('gt-mass-flow/<str:token>/decision/', views.GTMFlowDecisionView.as_view(), name='gtm_flow_decision'),
    path('gt-mass-flow/<str:token>/discard/', views.GTMFlowDiscardView.as_view(), name='gtm_flow_discard'),
    path('gt-mass-flow/<str:token>/download/', views.GTMFlowDownloadView.as_view(), name='gtm_flow_download'),
    path('gt-mass-flow/<str:token>/export/', views.GTMFlowExportView.as_view(), name='gtm_flow_export'),
    path('gt-mass-flow/<str:token>/save-later/', views.GTMFlowSaveLaterView.as_view(), name='gtm_flow_save_later'),
    # Modern Trade (MT) on the shared PO-flow scaffold (upload → review → confirm
    # → lock), on par with the online marketplaces. Channel picked at upload. The
    # old single-page shoppers-stop generator above stays as a fallback.
    path('mt-flow/', views.MTFlowUploadView.as_view(), name='mt_flow_upload'),
    # 'drafts' list — MUST precede the <str:token> review route (else captured as a token).
    path('mt-flow/drafts/', views.MTFlowDraftsView.as_view(), name='mt_flow_drafts'),
    path('mt-flow/<str:token>/', views.MTFlowReviewView.as_view(), name='mt_flow_review'),
    path('mt-flow/<str:token>/confirm/', views.MTFlowConfirmView.as_view(), name='mt_flow_confirm'),
    path('mt-flow/<str:token>/decision/', views.MTFlowDecisionView.as_view(), name='mt_flow_decision'),
    path('mt-flow/<str:token>/discard/', views.MTFlowDiscardView.as_view(), name='mt_flow_discard'),
    path('mt-flow/<str:token>/download/', views.MTFlowDownloadView.as_view(), name='mt_flow_download'),
    path('mt-flow/<str:token>/export/', views.MTFlowExportView.as_view(), name='mt_flow_export'),
    path('mt-flow/<str:token>/save-later/', views.MTFlowSaveLaterView.as_view(), name='mt_flow_save_later'),
    # EKA (EBO / Kiosk / Airport → SO/TO) — third offline channel on the shared
    # po-flow scaffold (upload single/bulk → review → confirm → record). No CP check.
    path('eka-flow/', views.EKAFlowUploadView.as_view(), name='eka_flow_upload'),
    path('eka-flow/drafts/', views.EKAFlowDraftsView.as_view(), name='eka_flow_drafts'),
    path('eka-flow/<str:token>/', views.EKAFlowReviewView.as_view(), name='eka_flow_review'),
    path('eka-flow/<str:token>/confirm/', views.EKAFlowConfirmView.as_view(), name='eka_flow_confirm'),
    path('eka-flow/<str:token>/decision/', views.EKAFlowDecisionView.as_view(), name='eka_flow_decision'),
    path('eka-flow/<str:token>/discard/', views.EKAFlowDiscardView.as_view(), name='eka_flow_discard'),
    path('eka-flow/<str:token>/download/', views.EKAFlowDownloadView.as_view(), name='eka_flow_download'),
    path('eka-flow/<str:token>/export/', views.EKAFlowExportView.as_view(), name='eka_flow_export'),
    path('eka-flow/<str:token>/save-later/', views.EKAFlowSaveLaterView.as_view(), name='eka_flow_save_later'),
]
