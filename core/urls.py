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
    # Users & Roles (admin/superuser-only) — assign Editor/Viewer, manage logins.
    path('users/', views.UsersRolesView.as_view(), name='users_roles'),
    path('users/create/', views.UserCreateView.as_view(), name='user_create'),
    path('users/<int:user_id>/role/', views.UserSetRoleView.as_view(), name='user_set_role'),
    path('users/<int:user_id>/password/', views.UserSetPasswordView.as_view(), name='user_set_password'),
    path('users/<int:user_id>/toggle-active/', views.UserToggleActiveView.as_view(), name='user_toggle_active'),
]