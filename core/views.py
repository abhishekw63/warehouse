from django.contrib import messages
from django.contrib.auth import login
from django.contrib.auth.forms import UserCreationForm
from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.contrib.auth.models import User
from django.contrib.auth.views import LoginView, LogoutView, PasswordChangeView
from django.shortcuts import redirect
from django.urls import reverse_lazy
from django.views.generic import CreateView, TemplateView, UpdateView, View


class CustomLogoutView(LogoutView):
    next_page = reverse_lazy('home')

    # NB: no "Logout successful" message here. The login page (landing.html) is a
    # standalone screen with no message/toast layer, so a logout message can't be
    # shown there — it would defer and stack onto the NEXT authenticated page,
    # surfacing the *previous* user's logout toast alongside the new user's
    # "Login successful". Landing back on the login screen is feedback enough.

def _home_stats() -> dict:
    """Lightweight LIVE figures for the home/login showcase — total POs, value,
    marketplaces and line items recorded in the shared order DB. Fully fail-safe
    (returns Nones if the DB is unavailable) so the login page always renders."""
    s = {'pos': None, 'value': None, 'marketplaces': None, 'lines': None}
    try:
        from online_b2b.services.order_db import _conn
        with _conn() as (cur, _d):
            cur.execute("SELECT COUNT(DISTINCT CONCAT(marketplace,'|',po)), "
                        "COALESCE(SUM(order_value),0), COUNT(DISTINCT marketplace) "
                        "FROM order_headers")
            r = cur.fetchone()
            if r:
                s['pos'] = int(r[0] or 0)
                s['value'] = float(r[1] or 0)
                s['marketplaces'] = int(r[2] or 0)
            cur.execute("SELECT COUNT(*) FROM order_lines")
            s['lines'] = int((cur.fetchone() or [0])[0] or 0)
    except Exception:  # noqa: BLE001 — never block the login page on the DB
        pass
    # friendly INR display (Cr / L) for the live value chip
    v = s.get('value')
    if v:
        s['value_disp'] = (f"₹{v/1e7:.2f} Cr" if v >= 1e7
                           else f"₹{v/1e5:.1f} L" if v >= 1e5
                           else f"₹{v:,.0f}")
    else:
        s['value_disp'] = None
    return s


class HomeView(LoginView):
    template_name = 'core/landing.html'
    redirect_authenticated_user = False

    def get_success_url(self):
        return reverse_lazy('departments')

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['stats'] = _home_stats()
        return ctx

    def form_valid(self, form):
        # Remember me → persist 2 weeks; else expire when the browser closes.
        self.request.session.set_expiry(
            60 * 60 * 24 * 14 if self.request.POST.get('remember') else 0)
        messages.success(self.request, 'Login successful')
        return super().form_valid(form)

class DepartmentsView(LoginRequiredMixin, TemplateView):
    template_name = 'core/departments.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['stats'] = _home_stats()                 # live at-a-glance snapshot
        try:
            from online_b2b.services.order_db import hub_extra_kpis, recent_orders
            ctx['hub'] = hub_extra_kpis()
            ctx['recent'] = recent_orders(6)         # recent-activity feed
        except Exception:  # noqa: BLE001 — never block the hub on the DB
            ctx['hub'] = {}
            ctx['recent'] = []
        return ctx

class SignUpView(CreateView):
    form_class = UserCreationForm
    template_name = 'core/signup.html'
    success_url = reverse_lazy('departments')

    def form_valid(self, form):
        response = super().form_valid(form)
        login(self.request, self.object)
        messages.success(self.request, 'Account created successfully. Welcome!')
        return response

    def dispatch(self, request, *args, **kwargs):
        if request.user.is_authenticated:
            return redirect('departments')
        return super().dispatch(request, *args, **kwargs)

class ProfileView(LoginRequiredMixin, UpdateView):
    model = User
    fields = ['first_name', 'last_name', 'email']
    template_name = 'core/profile.html'
    success_url = reverse_lazy('profile')

    def get_object(self):
        return self.request.user

    def form_valid(self, form):
        messages.success(self.request, 'Your profile details were updated successfully.')
        return super().form_valid(form)

class CustomPasswordChangeView(LoginRequiredMixin, PasswordChangeView):
    template_name = 'core/password_change.html'
    success_url = reverse_lazy('profile')

    def form_valid(self, form):
        messages.success(self.request, 'Your password was successfully updated.')
        return super().form_valid(form)


class _StaffOnly(LoginRequiredMixin, UserPassesTestMixin):
    """Gate dev tooling to staff users."""
    def test_func(self):
        return bool(getattr(self.request.user, 'is_staff', False))


class DevDashboardView(_StaffOnly, TemplateView):
    """Dev · Health — live request perf (from the timing middleware) + an
    on-demand all-angles code audit. Staff-only; read-only; never touches the
    business backend."""
    template_name = 'core/dev_dashboard.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        from . import code_audit
        from . import observability as obs
        rows = obs.recent(800)
        ctx['kpi'] = obs.kpis(rows)
        ctx['agg'] = obs.aggregate(rows)
        ctx['recent'] = list(reversed(rows))[:60]
        ctx['audit'] = code_audit.last_audit()
        return ctx


class DevAuditView(_StaffOnly, View):
    """Run the code audit now (POST), then back to the dashboard."""
    def post(self, request, *args, **kwargs):
        from . import code_audit
        try:
            code_audit.run_audit()
            messages.success(request, 'Code audit complete.')
        except Exception as e:  # noqa: BLE001
            messages.error(request, f'Audit failed: {e}')
        return redirect('dev_dashboard')


class ProjectMapView(_StaffOnly, TemplateView):
    """Project Map — a graphical, always-current map of the whole system: the
    file tree (apps → modules → templates), the real URL→view routes, the DB
    models/tables, and the upload→review→confirm→record data flow. Auto-generated
    from the live codebase, so it updates whenever the code changes. Staff-only."""
    template_name = 'core/project_map.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        from . import project_map as pm
        ctx['tree'] = pm.app_tree()
        ctx['routes'] = pm.routes()
        ctx['models'] = pm.models()
        ctx['summary'] = pm.summary()
        ctx['updated'] = pm.last_updated()
        return ctx


# ── Users & Roles (admin-only) ───────────────────────────────────────────────
# Manage who can write. "Editor" = member of the Editors group (or superuser) →
# may perform every write; everyone else is a view-only "Viewer". The write-guard
# middleware (core.access) is what actually enforces it; this page just assigns
# the role. Superuser-only — account management is sensitive.
from django.contrib.auth.models import Group  # noqa: E402
from django.http import JsonResponse  # noqa: E402
from django.shortcuts import get_object_or_404  # noqa: E402

from .access import AdminRequiredMixin, EDITORS_GROUP, is_editor  # noqa: E402


def _editors_group():
    grp, _ = Group.objects.get_or_create(name=EDITORS_GROUP)
    return grp


class UsersRolesView(AdminRequiredMixin, TemplateView):
    """List every user with their role (Editor/Viewer) + active state, plus a
    form to add a new user. The admin toggles roles inline."""
    template_name = 'core/users_roles.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        users = User.objects.all().order_by('-is_superuser', 'username')
        rows = [{
            'obj': u, 'id': u.id, 'username': u.username,
            'full_name': u.get_full_name(), 'email': u.email,
            'is_editor': is_editor(u), 'is_superuser': u.is_superuser,
            'is_active': u.is_active,
            'last_login': u.last_login,
        } for u in users]
        ctx['rows'] = rows
        ctx['n_editors'] = sum(1 for r in rows if r['is_editor'])
        ctx['n_viewers'] = sum(1 for r in rows if not r['is_editor'])
        return ctx


class _AdminAction(AdminRequiredMixin, View):
    """Base for the JSON POST actions — returns {ok,...} for the inline UI, or
    redirects back to the page for a plain form post."""
    def _resp(self, request, payload, msg=None, level='success'):
        if request.headers.get('x-requested-with') == 'XMLHttpRequest':
            return JsonResponse(payload)
        if msg:
            getattr(messages, level)(request, msg)
        return redirect('users_roles')


class UserCreateView(_AdminAction):
    """Create a new user with a chosen role (editor/viewer)."""
    def post(self, request, *args, **kwargs):
        username = (request.POST.get('username') or '').strip()
        password = request.POST.get('password') or ''
        role = (request.POST.get('role') or 'viewer').strip().lower()
        if not username or not password:
            return self._resp(request, {'ok': False, 'error': 'Username and password are required.'},
                              'Username and password are required.', 'error')
        if User.objects.filter(username=username).exists():
            return self._resp(request, {'ok': False, 'error': 'That username already exists.'},
                              'That username already exists.', 'error')
        u = User.objects.create_user(username=username, password=password)
        if role == 'editor':
            u.groups.add(_editors_group())
        return self._resp(request, {'ok': True, 'id': u.id, 'username': u.username, 'role': role},
                          f"User '{username}' created as {role.title()}.")


class UserSetRoleView(_AdminAction):
    """Promote/demote a user between Editor and Viewer."""
    def post(self, request, user_id, *args, **kwargs):
        u = get_object_or_404(User, pk=user_id)
        role = (request.POST.get('role') or '').strip().lower()
        if u.is_superuser:
            return self._resp(request, {'ok': False, 'error': 'Superusers are always Editors.'},
                              'Superusers are always Editors.', 'error')
        if role == 'editor':
            u.groups.add(_editors_group())
        elif role == 'viewer':
            u.groups.remove(_editors_group())
        else:
            return self._resp(request, {'ok': False, 'error': 'Unknown role.'}, 'Unknown role.', 'error')
        return self._resp(request, {'ok': True, 'id': u.id, 'role': role},
                          f"{u.username} is now a {role.title()}.")


class UserSetPasswordView(_AdminAction):
    """Reset a user's password."""
    def post(self, request, user_id, *args, **kwargs):
        u = get_object_or_404(User, pk=user_id)
        pw = request.POST.get('password') or ''
        if len(pw) < 1:
            return self._resp(request, {'ok': False, 'error': 'Password cannot be blank.'},
                              'Password cannot be blank.', 'error')
        u.set_password(pw)
        u.save(update_fields=['password'])
        return self._resp(request, {'ok': True, 'id': u.id}, f"Password reset for {u.username}.")


class UserToggleActiveView(_AdminAction):
    """Activate / deactivate a login (deactivated users can't sign in)."""
    def post(self, request, user_id, *args, **kwargs):
        u = get_object_or_404(User, pk=user_id)
        if u == request.user:
            return self._resp(request, {'ok': False, 'error': "You can't deactivate yourself."},
                              "You can't deactivate yourself.", 'error')
        if u.is_superuser and u.is_active:
            return self._resp(request, {'ok': False, 'error': "Can't deactivate a superuser."},
                              "Can't deactivate a superuser.", 'error')
        u.is_active = not u.is_active
        u.save(update_fields=['is_active'])
        state = 'active' if u.is_active else 'deactivated'
        return self._resp(request, {'ok': True, 'id': u.id, 'is_active': u.is_active},
                          f"{u.username} is now {state}.")


class AuditLogView(_StaffOnly, TemplateView):
    """Staff-only 'who did what, when' trail — every write is logged by the RBAC
    middleware (core.audit). Read-only; filter by user / text."""
    template_name = 'core/audit_log.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        from . import audit
        u = (self.request.GET.get('user') or '').strip()
        q = (self.request.GET.get('q') or '').strip()
        ctx['rows'] = audit.recent(400, user=u, q=q)
        ctx['f_user'] = u
        ctx['f_q'] = q
        return ctx
