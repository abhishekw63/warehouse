"""
core.access
===========

Role-based write control (RBAC), fully additive — it never touches the business
DB and never changes what a page *shows*, only whether a **write** is allowed.

Two roles:
  • **Editor** — superuser OR a member of the ``Editors`` group. Can do every
    write: upload POs, review decisions, Lock & Record, mapping/inventory/item-
    master CRUD, deletes, send emails, etc.
  • **Viewer** — every other logged-in user. Can open every page and
    **download / export / search / preview**, but cannot write.

Enforcement is **deny-writes-by-default** at the middleware layer
(:class:`WriteGuardMiddleware`): any unsafe HTTP method (POST/PUT/PATCH/DELETE)
from a non-Editor is blocked *unless* its URL is a genuinely read-only endpoint
(export / download / search / pagination / preview / the reconciliation tools).
So a brand-new write view added later is protected automatically — there is no
per-view decorator to remember.

The template flag ``can_write`` (context processor) drives the UI: write buttons
render disabled with a "View-only access" tooltip for Viewers. That is cosmetic;
the middleware is the real security boundary.
"""
from __future__ import annotations

from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin
from django.http import JsonResponse
from django.shortcuts import render

EDITORS_GROUP = 'Editors'

# Unsafe methods that mutate — anything else (GET/HEAD/OPTIONS) is always allowed.
_UNSAFE = {'POST', 'PUT', 'PATCH', 'DELETE'}

# Read-only POST endpoints a Viewer may still call. Two rules, OR'd:
#   1) the URL name ends in one of these suffixes (exports, searches, previews,
#      pagination, data fetches — none of which write), or
#   2) the URL name is in the explicit set below (reads that don't fit a suffix).
# A write endpoint NEVER ends in these suffixes, so the suffix rule is safe.
_SAFE_SUFFIXES = ('_export', '_download', '_search', '_more', '_preview', '_data')

_SAFE_POST_NAMES = frozenset({
    # auth / self-service account
    'login', 'logout', 'signup', 'password_change', 'profile',
    # reconciliation tools — upload D365 files but write NOTHING to the business
    # DB (they only produce an Excel report). Allowed for Viewers by decision.
    'b2b_full_validation_run', 'b2b_triangular_run',
    # availability checker — read-only reports (the best-WH comparison too;
    # 'b2b_availability_shift'/'..._run_delete' are deliberately NOT here — they
    # write, Editors only). 'record' is an append-only snapshot → Viewers may keep
    # an audit of the FR they saw, so it's allowed.
    'b2b_availability_check', 'b2b_availability_bins', 'b2b_availability_scenarios',
    'b2b_availability_record',
    # read-only data fetches that don't match a suffix
    'b2b_cockpit_po_skus',
    # generating the ERP D365 dump is an export of an already-locked run (no DB
    # write) — a Viewer who can see a locked run may pull its dump.
    'b2b_generate_d365',
})

# Endpoints that manage users/roles — allowed to pass the write-guard ONLY for an
# admin (superuser). The views themselves also gate on superuser, so this is
# belt-and-suspenders, never a hole for a Viewer.
_ADMIN_POST_NAMES = frozenset({
    'user_create', 'user_set_role', 'user_set_password', 'user_toggle_active',
})


def is_editor(user) -> bool:
    """True if the user may perform writes (superuser or in the Editors group)."""
    return bool(
        user and user.is_authenticated
        and (user.is_superuser or user.groups.filter(name=EDITORS_GROUP).exists())
    )


def is_role_admin(user) -> bool:
    """True if the user may manage other users / assign roles (the admin).
    Kept to superuser — account management is sensitive."""
    return bool(user and user.is_authenticated and user.is_superuser)


def _post_is_read_only(name: str | None) -> bool:
    if not name:
        return False
    if name in _SAFE_POST_NAMES:
        return True
    return name.endswith(_SAFE_SUFFIXES)


class WriteGuardMiddleware:
    """Deny writes by default for non-Editors. Uses ``process_view`` so the URL
    is already resolved (``request.resolver_match`` is set) and we can allow the
    read-only endpoints by name. Editors and safe reads pass straight through."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        return self.get_response(request)

    def process_view(self, request, view_func, view_args, view_kwargs):
        if request.method not in _UNSAFE:
            return None
        user = getattr(request, 'user', None)
        match = getattr(request, 'resolver_match', None)
        name = match.url_name if match else None
        if is_editor(user):
            self._audit(request, name, view_kwargs)   # record the editor's write
            return None
        # Django admin has its own permission system (staff/superuser) — leave it.
        if request.path.startswith('/admin/'):
            return None
        if _post_is_read_only(name):
            return None
        # role-management endpoints: allow only for the admin (superuser)
        if name in _ADMIN_POST_NAMES and is_role_admin(user):
            self._audit(request, name, view_kwargs)
            return None
        return self._deny(request)

    @staticmethod
    def _audit(request, name, view_kwargs):
        """Log a real write to the audit trail (skip read-only POSTs like exports/
        search). Best-effort — never affects the request. Captures who + what +
        a compact target (token / row id / PO) + the action."""
        if _post_is_read_only(name) or request.path.startswith('/admin/'):
            return
        try:
            from core import audit
            u = getattr(request.user, 'username', '') or 'system'
            vk = view_kwargs or {}
            tk = (vk.get('token') or vk.get('row_id') or vk.get('user_id')
                  or vk.get('run_id') or '')
            po = (request.POST.get('po') or request.POST.get('aff_key') or '')
            target = ' · '.join(str(x) for x in (tk, po) if x)[:300]
            detail = (request.POST.get('aff_action') or request.POST.get('action')
                      or request.POST.get('role') or '')[:500]
            audit.log(u, request.method, name or '', request.path, target, detail)
        except Exception:  # noqa: BLE001 — audit must never break a request
            pass

    @staticmethod
    def _deny(request):
        msg = ("View-only access — you don't have permission to make changes. "
               "Ask an admin for Editor access.")
        wants_json = (
            request.headers.get('x-requested-with') == 'XMLHttpRequest'
            or 'application/json' in request.headers.get('accept', '')
        )
        if wants_json:
            return JsonResponse({'ok': False, 'error': msg}, status=403)
        return render(request, 'core/403_write.html', {'message': msg}, status=403)


# ── CBV mixins ───────────────────────────────────────────────────────────────
class EditorRequiredMixin(LoginRequiredMixin, UserPassesTestMixin):
    """Gate a whole class-based view to Editors (extra guard on top of the
    middleware — use on views that should not even render for a Viewer)."""
    def test_func(self):
        return is_editor(self.request.user)


class AdminRequiredMixin(LoginRequiredMixin, UserPassesTestMixin):
    """Gate to the admin (superuser) — for the Users & Roles management page."""
    def test_func(self):
        return is_role_admin(self.request.user)


# ── template context ─────────────────────────────────────────────────────────
def roles(request):
    """Expose ``can_write`` (Editor?) and ``is_role_admin`` to every template so
    the UI can disable write buttons and show the admin link."""
    user = getattr(request, 'user', None)
    return {'can_write': is_editor(user), 'is_role_admin': is_role_admin(user)}
