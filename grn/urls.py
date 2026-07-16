from django.urls import path

from . import views

urlpatterns = [
    path('', views.index, name='grn_index'),
    path('upload/', views.upload, name='grn_upload'),
    path('<str:token>/', views.result, name='grn_result'),
    path('<str:token>/export/', views.export, name='grn_export'),
]
