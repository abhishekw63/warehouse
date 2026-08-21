"""
online_b2b.views_tables
=======================

Class-based views for the dashboard **Tables** tab — a no-code master-tables
manager (create tables + full CRUD), backed by
:mod:`online_b2b.services.custom_tables`.

Self-contained & removable: its own views module + urls + template. All writes
are AJAX (JSON in / JSON out) so the page never refreshes. API-ready shape:
every endpoint returns ``{ok, ...}`` / ``{ok:false, error}``.
"""
from __future__ import annotations

import json

from django.contrib.auth.mixins import LoginRequiredMixin
from django.http import JsonResponse
from django.shortcuts import render
from django.views import View

from .services import custom_tables as ct


class _JsonSafe:
    """Mixin for AJAX endpoints: any unhandled exception in the view becomes a
    ``{ok:false, error}`` JSON response (logged) instead of a raw Django 500 HTML
    page — the client does ``r.json()`` and would otherwise see only a vague
    'Network error'. Apply to JSON CBVs only (not the HTML shell)."""

    def dispatch(self, request, *args, **kwargs):
        try:
            return super().dispatch(request, *args, **kwargs)
        except Exception as e:  # noqa: BLE001
            import logging
            logging.getLogger(__name__).exception('Tables endpoint failed')
            return JsonResponse(
                {'ok': False, 'error': f'{type(e).__name__}: {e}'}, status=200)


def _body(request) -> dict:
    try:
        return json.loads(request.body or '{}')
    except Exception:  # noqa: BLE001
        return {}


class TablesHomeView(LoginRequiredMixin, View):
    """The Tables tab shell — table list + the active table's grid."""

    def get(self, request):
        ct.ensure_schema()
        tables = ct.list_tables()
        active = request.GET.get('t') or (tables[0]['slug'] if tables else '')
        current = ct.get_table(active) if active else None
        rows = ct.list_rows(current['id']) if current else []
        return render(request, 'online_b2b/tables.html', {
            'tables': tables, 'current': current, 'rows': rows,
        })


class TableDataView(_JsonSafe, LoginRequiredMixin, View):
    """JSON for one table (columns + colour rules + rows) — used on tab switch."""

    def get(self, request, slug):
        t = ct.get_table(slug)
        if not t:
            return JsonResponse({'ok': False, 'error': 'Table not found.'}, status=404)
        return JsonResponse({'ok': True, 'table': t, 'rows': ct.list_rows(t['id'])})


class RowAddView(_JsonSafe, LoginRequiredMixin, View):
    def post(self, request):
        d = _body(request)
        if not d.get('table_id'):
            return JsonResponse({'ok': False, 'error': 'table_id required'}, status=400)
        rid = ct.add_row(int(d['table_id']), d.get('data') or {})
        return JsonResponse({'ok': True, 'id': rid})


class RowUpdateView(_JsonSafe, LoginRequiredMixin, View):
    def post(self, request, row_id):
        ct.update_row(int(row_id), _body(request).get('data') or {})
        return JsonResponse({'ok': True})


class RowDeleteView(_JsonSafe, LoginRequiredMixin, View):
    def post(self, request, row_id):
        ct.delete_row(int(row_id))
        return JsonResponse({'ok': True})


class TableCreateView(_JsonSafe, LoginRequiredMixin, View):
    def post(self, request):
        d = _body(request)
        name = (d.get('name') or '').strip()
        columns = d.get('columns') or []
        if not name or not columns:
            return JsonResponse({'ok': False, 'error': 'Name and at least one column are required.'}, status=400)
        tid = ct.create_table(name, columns, d.get('color_rules') or {})
        return JsonResponse({'ok': True, 'id': tid, 'slug': ct.get_table(tid)['slug']})


class TableRenameView(_JsonSafe, LoginRequiredMixin, View):
    def post(self, request, table_id):
        name = (_body(request).get('name') or '').strip()
        if name:
            ct.update_table(int(table_id), name=name)
        return JsonResponse({'ok': True})


class TableDeleteView(_JsonSafe, LoginRequiredMixin, View):
    def post(self, request, table_id):
        ct.delete_table(int(table_id))
        return JsonResponse({'ok': True})
