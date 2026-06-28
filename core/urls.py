from django.urls import path

from . import views

urlpatterns = [
    path('', views.HomeView.as_view(), name='home'),
    path('login/', views.HomeView.as_view(), name='login'),
    path('signup/', views.SignUpView.as_view(), name='signup'),
    path('logout/', views.CustomLogoutView.as_view(), name='logout'),
    path('departments/', views.DepartmentsView.as_view(), name='departments'),
    path('profile/', views.ProfileView.as_view(), name='profile'),
    path('password-change/', views.CustomPasswordChangeView.as_view(), name='password_change'),
    # Dev · Health (staff-only) — request perf + code audit.
    path('dev/', views.DevDashboardView.as_view(), name='dev_dashboard'),
    path('dev/audit/', views.DevAuditView.as_view(), name='dev_audit'),
    path('dev/map/', views.ProjectMapView.as_view(), name='dev_map'),
]