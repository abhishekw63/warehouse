from django.urls import path
from django.views.generic import RedirectView

from . import views

urlpatterns = [
    path('', views.HomeView.as_view(), name='home'),
    path('login/', views.HomeView.as_view(), name='login'),
    path('signup/', views.SignUpView.as_view(), name='signup'),
    path('logout/', views.CustomLogoutView.as_view(), name='logout'),
    # Departments picker removed — everything lands straight on Order Management.
    # The name is KEPT so the ~50 breadcrumb {% url 'departments' %} refs resolve.
    path('departments/', RedirectView.as_view(pattern_name='b2b_dashboard',
                                              permanent=False), name='departments'),
    path('profile/', views.ProfileView.as_view(), name='profile'),
    path('password-change/', views.CustomPasswordChangeView.as_view(), name='password_change'),
    # Dev · Health (staff-only) — request perf + code audit.
    path('dev/', views.DevDashboardView.as_view(), name='dev_dashboard'),
    path('dev/audit/', views.DevAuditView.as_view(), name='dev_audit'),
    path('dev/map/', views.ProjectMapView.as_view(), name='dev_map'),
    # Users & Roles (admin/superuser-only) — assign Editor/Viewer, manage logins.
    path('users/', views.UsersRolesView.as_view(), name='users_roles'),
    path('users/create/', views.UserCreateView.as_view(), name='user_create'),
    path('users/<int:user_id>/role/', views.UserSetRoleView.as_view(), name='user_set_role'),
    path('users/<int:user_id>/password/', views.UserSetPasswordView.as_view(), name='user_set_password'),
    path('users/<int:user_id>/toggle-active/', views.UserToggleActiveView.as_view(), name='user_toggle_active'),
    # Audit trail (staff-only) — who did every write, when.
    path('audit/', views.AuditLogView.as_view(), name='audit_log'),
    # client Navigation-Timing beacon (frontend load telemetry → NAV audit rows)
    path('perf/nav/', views.perf_nav, name='perf_nav'),
]