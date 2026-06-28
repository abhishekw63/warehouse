from django.contrib import messages
from django.contrib.auth import login
from django.contrib.auth.forms import UserCreationForm
from django.contrib.auth.mixins import LoginRequiredMixin
from django.contrib.auth.models import User
from django.contrib.auth.views import LoginView, LogoutView, PasswordChangeView
from django.shortcuts import redirect
from django.urls import reverse_lazy
from django.views.generic import CreateView, TemplateView, UpdateView


class CustomLogoutView(LogoutView):
    next_page = reverse_lazy('home')

    def dispatch(self, request, *args, **kwargs):
        messages.success(request, 'Logout successful.')
        return super().dispatch(request, *args, **kwargs)

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
        messages.success(self.request, 'Login successful')
        return super().form_valid(form)

class DepartmentsView(LoginRequiredMixin, TemplateView):
    template_name = 'core/departments.html'

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
