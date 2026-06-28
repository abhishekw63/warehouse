"""
online_b2b.views — Phase 0 (Blink pilot)

Dashboard reads the order history straight from MySQL ``renee_orders``.
Upload runs the existing engine as a library and records the result back into
the same DB, so the dashboard reflects web runs immediately.
"""

import json
import os
import shutil
import uuid
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.contrib.auth.mixins import LoginRequiredMixin
from django.http import FileResponse, Http404, HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.views.decorators.http import require_POST
from django.views.generic import TemplateView

from .forms import UploadForm
from .services import engine_bridge, erp_import, order_db

# ── Central hub + branch dashboards (class-based) ───────────────────────────
# /b2b/ is the central Order-Management hub: compact overall KPIs + two group
# cards (Online B2B / Offline). Each group drills into its own branch dashboard.
# Channels differ a lot between the two worlds, so they stay distinct branches.

# Channels shown under each group card on the hub.
ONLINE_CHANNELS = [
    {'name': 'Blink', 'tag': 'live'},
    {'name': 'Flipkart', 'tag': 'live'},
    {'name': 'RK', 'tag': 'live'},
    {'name': 'DMart', 'tag': 'live'},
    {'name': 'Zepto', 'tag': 'live'},
    {'name': 'Flipkart Branch', 'tag': 'live'},
    {'name': 'Purplle', 'tag': 'live'},
    {'name': 'Swiggy', 'tag': 'live'},
    {'name': 'Nykaa', 'tag': 'live'},
    {'name': 'Myntra', 'tag': 'live'},
    {'name': 'Reliance', 'tag': 'live'},
    {'name': 'Meesho Branch', 'tag': 'live'},
    # ── TO-DO · pending web integration (shown as "coming soon") ──
    {'name': 'BlinkMP', 'tag': 'soon'},
    {'name': 'Big Basket', 'tag': 'soon'},
    {'name': 'First Cry', 'tag': 'soon'},
    {'name': 'Smytten', 'tag': 'soon'},
]
OFFLINE_CHANNELS = [
    {'name': 'MT (Modern Trade)', 'tag': 'live', 'url': '/offline/shoppers-stop/'},
    {'name': 'GT Mass', 'tag': 'live', 'url': '/offline/gt-mass-dump/'},
    {'name': 'GT Select', 'tag': 'live', 'url': '/b2b/gt-select/'},
    # ── TO-DO · pending web integration (shown as "coming soon") ──
    {'name': 'EKA', 'tag': 'soon'},
    {'name': 'CSD', 'tag': 'soon'},
    {'name': 'Off-Institutional', 'tag': 'soon'},
    {'name': 'Airport', 'tag': 'soon'},
    {'name': 'EBO / Kiosk', 'tag': 'soon'},
]


class CentralHubView(LoginRequiredMixin, TemplateView):
    """`/b2b/` — central hub: compact overall KPIs + Online B2B / Offline groups."""
    template_name = 'online_b2b/central.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        data = order_db.overview(segment='')          # all channels
        ctx['ok'] = data.get('ok', False)
        ctx['error'] = data.get('error')
        ctx['backend'] = data.get('backend', '')
        ctx['kpis'] = data.get('kpis', {})
        ctx['issue_count'] = data.get('issue_count', 0)
        ctx['online_channels'] = ONLINE_CHANNELS
        ctx['offline_channels'] = OFFLINE_CHANNELS
        # Collapse the pending ones into a single "+N coming soon" chip (tooltip
        # lists them) so the hub card stays compact.
        ctx['online_soon'] = [c['name'] for c in ONLINE_CHANNELS if c.get('tag') == 'soon']
        ctx['offline_soon'] = [c['name'] for c in OFFLINE_CHANNELS if c.get('tag') == 'soon']
        ctx['online_kpis'] = order_db.segment_kpis('OnlineB2B')
        ctx['offline_kpis'] = order_db.segment_kpis('Offline')
        ctx['today'] = order_db.today_intake()
        ctx['recent'] = order_db.recent_orders(8)
        ctx['recent_runs'] = order_db.recent_runs(8)
        from .services import tat_store
        tat_total = tat_store.breach_count()          # all TAT breaches
        total_pos = (data.get('kpis') or {}).get('pos') or 0
        ctx['tat_total'] = tat_total
        ctx['tat_rate'] = round(tat_total / total_pos * 100, 1) if total_pos else 0
        ctx['extra'] = order_db.hub_extra_kpis()
        # Item-master staleness — 15-day refresh reminder.
        from .services import item_master_loader as iml
        ctx['im_status'] = iml.last_updated()
        return ctx


# Branch descriptors — drive the shared overview template's header/actions.
ONLINE_BRANCH = {'kind': 'online', 'label': 'Online B2B'}
OFFLINE_BRANCH = {'kind': 'offline', 'label': 'Offline'}


class RulesView(LoginRequiredMixin, TemplateView):
    """Marketplace Rules & Exceptions reference — margins, compare basis, item
    resolution, the engine's exception types + operator decisions, and Flipkart's
    location map. Reads the engine config so it never drifts."""
    template_name = 'online_b2b/rules.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        rules = engine_bridge.marketplace_rules()
        ctx['rules'] = rules
        # Marketplaces whose margin is GST-dependent (Reliance) — for the
        # elaborated exceptions section.
        ctx['gst_margin_rules'] = [r for r in rules if r.get('gst_margin')]
        ctx['formats'] = engine_bridge.marketplace_formats()
        # Marketplaces that have a full "See full template" preview available.
        ctx['template_names'] = list(engine_bridge.marketplace_templates().keys())
        ctx['locations'] = engine_bridge.location_rules()
        # Actual Swiggy deal SKUs (name + agreed prices) for accuracy on the card.
        try:
            from .services import overrides_store
            ctx['swiggy_deals'] = [r for r in overrides_store.list_all()
                                   if r.get('kind') == 'swiggy_deal']
        except Exception:  # noqa: BLE001
            ctx['swiggy_deals'] = []
        return ctx


class MarketplaceTemplateView(LoginRequiredMixin, TemplateView):
    """Rules → "See full template": the full column list + a few real sample rows
    for one marketplace, with the columns the engine actually reads highlighted
    (by role) and the rest dulled — a visual, drift-proof successor to the desktop
    "download template"."""
    template_name = 'online_b2b/template.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        name = self.kwargs.get('slug', '')
        tpl = engine_bridge.marketplace_template(name)
        if tpl is None:
            raise Http404(f'No template captured for “{name}”.')
        ctx['tpl'] = tpl
        return ctx


class AnalyticsView(LoginRequiredMixin, TemplateView):
    """Management daily-intake analytics: daily stacked chart by segment +
    segment→marketplace→child breakdown. Date range via ?days= (7/30/90)."""
    template_name = 'online_b2b/analytics.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        try:
            days = int(self.request.GET.get('days') or 30)
        except (TypeError, ValueError):
            days = 30
        days = days if days in (7, 30, 90) else 30
        ctx['days'] = days
        # Single-date scope for the breakdown tree (YYYY-MM-DD). Defaults to
        # TODAY when the page first opens (no 'date' key); an explicit empty
        # 'date=' (the Clear button) shows the whole range instead.
        import datetime as _dt
        raw = self.request.GET.get('date')
        if raw is None:
            date = _dt.date.today().isoformat()
        else:
            date = raw.strip()
        if date:
            try:
                _dt.date.fromisoformat(date)
            except ValueError:
                date = ''
        ctx['date'] = date
        ctx['hier'] = order_db.intake_hierarchy(days, date=date)
        daily = order_db.daily_intake(days)          # raw dict → json_script encodes
        if date:
            try:
                daily['focus'] = _dt.date.fromisoformat(date).strftime('%d %b')
            except ValueError:
                pass
        ctx['daily'] = daily
        return ctx


class OfflineBranchView(LoginRequiredMixin, TemplateView):
    """`/b2b/offline/` — the Offline branch: SAME rich dashboard as Online B2B
    (KPIs + charts + marketplace mix), scoped to the Offline segment."""
    template_name = 'online_b2b/overview.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['d'] = order_db.overview(segment='Offline')
        ctx['branch'] = OFFLINE_BRANCH
        return ctx

_MEDIA = Path(settings.MEDIA_ROOT)
_RUNS_INDEX = _MEDIA / 'b2b_runs'          # run_id -> {output_path,...} sidecars


def _save_run_index(run_id, payload: dict) -> None:
    if run_id is None:
        return
    _RUNS_INDEX.mkdir(parents=True, exist_ok=True)
    (_RUNS_INDEX / f"{run_id}.json").write_text(
        json.dumps(payload), encoding='utf-8')


def _load_run_index(run_id) -> dict:
    p = _RUNS_INDEX / f"{run_id}.json"
    if p.exists():
        try:
            return json.loads(p.read_text(encoding='utf-8'))
        except Exception:
            return {}
    return {}


def _persist_d365(run_id, marketplace, paths, warehouse, margin_pct, actions,
                  ean_fixes=None):
    """Build the decided D365 dump once at lock time and store it as a web-owned
    sidecar next to the run index. Returns the saved path, or None on any
    failure — a D365 hiccup (e.g. every line Excluded) must never block the lock.
    Pure file output: no DB connection is opened here."""
    try:
        _RUNS_INDEX.mkdir(parents=True, exist_ok=True)
        out = _RUNS_INDEX / f"{run_id}_d365.xlsx"
        res = engine_bridge.generate_d365(
            marketplace, paths, str(out),
            warehouse=warehouse, margin_pct=margin_pct, actions=actions,
            ean_fixes=ean_fixes)
        if res.get('ok') and os.path.exists(res.get('d365_path', '')):
            return str(out)
    except Exception:
        pass
    return None


def _int(request, name):
    try:
        return int(request.GET.get(name) or 0)
    except (TypeError, ValueError):
        return 0


def _filters(request):
    g = request.GET
    return {
        'segment': g.get('segment', '').strip(),
        'marketplace': g.get('marketplace', '').strip(),
        'days': _int(request, 'days'),
        'q': g.get('q', '').strip(),
        'warehouse': g.get('warehouse', '').strip(),
        'order_type': g.get('order_type', '').strip(),
        'date_from': g.get('date_from', '').strip(),
        'date_to': g.get('date_to', '').strip(),
        'sort': g.get('sort', 'date').strip() or 'date',
        'direction': g.get('dir', 'desc').strip() or 'desc',
    }


def _is_ajax(request):
    return bool(request.GET.get('partial')) or request.headers.get(
        'x-requested-with') == 'XMLHttpRequest'


@login_required
def dashboard(request):
    """Online B2B branch dashboard — KPIs + charts + marketplace summary for the
    online marketplaces (Blink/Flipkart/RK/DMart/BlinkMP). Reached from the central hub."""
    data = order_db.overview(segment='OnlineB2B')
    return render(request, 'online_b2b/overview.html',
                  {'d': data, 'branch': ONLINE_BRANCH})


@login_required
def orders(request):
    """Full order list — filter bar + table + load-more + export."""
    f = _filters(request)
    data = order_db.dashboard(offset=_int(request, 'offset'), **f)
    if _is_ajax(request):
        return render(request, 'online_b2b/_orders_results.html', {'d': data})
    return render(request, 'online_b2b/orders.html', {'d': data})


@login_required
def orders_more(request):
    """'Load more' → next page of order rows only."""
    f = _filters(request)
    page = order_db.orders_page(offset=_int(request, 'offset'), **f)
    return render(request, 'online_b2b/_orders_rows.html',
                  {'orders': page['orders'], 'page': page})


def _line_filters(request):
    g = request.GET
    return {
        'marketplace': g.get('marketplace', '').strip(),
        'status': g.get('status', '').strip(),
        'po': g.get('po', '').strip(),
        'q': g.get('q', '').strip(),
    }


@login_required
def lines(request):
    """Browsable line-items explorer (order_lines, all lines)."""
    data = order_db.line_items(offset=_int(request, 'offset'),
                               **_line_filters(request))
    if _is_ajax(request):
        return render(request, 'online_b2b/_lines_results.html', {'d': data})
    return render(request, 'online_b2b/lines.html', {'d': data})


@login_required
def lines_more(request):
    """'Load more' → next page of line rows only."""
    page = order_db.line_items_page(offset=_int(request, 'offset'),
                                    **_line_filters(request))
    return render(request, 'online_b2b/_lines_rows.html',
                  {'rows': page['rows']})


def _issue_filters(request) -> dict:
    """Shared Issues filters (used by the page + the export)."""
    return {
        'marketplace': request.GET.get('marketplace', '').strip(),
        'q': request.GET.get('q', '').strip(),
        'status': request.GET.get('status', '').strip(),
        'resolution': request.GET.get('resolution', 'pending').strip() or 'pending',
        'date_from': request.GET.get('date_from', '').strip(),
        'date_to': request.GET.get('date_to', '').strip(),
    }


@login_required
def issues(request):
    data = order_db.issues(**_issue_filters(request))
    if _is_ajax(request):
        return render(request, 'online_b2b/_issues_table.html', {'d': data})
    return render(request, 'online_b2b/issues.html',
                  {'d': data, 'eanfix': order_db.ean_corrections()})


@login_required
def issues_export(request):
    """Download the currently-filtered issue lines as .xlsx (respects every
    filter, including the upload-date window). No row cap on the export."""
    import datetime as _dt
    import io as _io

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    data = order_db.issues(limit=0, **_issue_filters(request))
    rows = data.get('rows', []) if data.get('ok') else []
    cols = [('run_ts', 'Upload Date'), ('marketplace', 'Marketplace'),
            ('po', 'PO'), ('item_no', 'Item No'), ('ean', 'EAN'),
            ('received_ean', 'Received EAN'), ('description', 'Description'),
            ('qty', 'Qty'), ('vendor_mrp', 'Vendor MRP'), ('our_mrp', 'Our MRP'),
            ('vendor_cp', 'Vendor CP'), ('our_cp', 'Our CP'),
            ('vendor_landing', 'Vendor Landing'), ('our_landing', 'Our Landing'),
            ('diff', 'Diff'), ('status', 'Status'),
            ('exception_label', 'Exception'), ('action', 'Action'),
            ('remark', 'Remark')]
    wb = Workbook(); ws = wb.active; ws.title = 'Issues'
    hf = Font(bold=True, color='FFFFFF'); navy = PatternFill('solid', fgColor='1A237E')
    for c, (_k, h) in enumerate(cols, 1):
        cell = ws.cell(1, c, h)
        cell.font = hf; cell.fill = navy
        cell.alignment = Alignment(horizontal='center')
    for r, row in enumerate(rows, 2):
        for c, (k, _h) in enumerate(cols, 1):
            v = row.get(k)
            ws.cell(r, c, str(v) if k == 'run_ts' and v is not None else v)
    for col in ws.columns:
        L = col[0].column_letter
        w = max((len(str(c.value or '')) for c in col), default=8)
        ws.column_dimensions[L].width = min(w + 2, 48)
    buf = _io.BytesIO(); wb.save(buf); buf.seek(0)
    f = _issue_filters(request)
    scope = f['date_from'] or 'all'
    if f['date_to'] and f['date_to'] != f['date_from']:
        scope = f"{f['date_from'] or 'start'}_to_{f['date_to']}"
    stamp = _dt.datetime.now().strftime('%Y%m%d_%H%M%S')
    fname = f"issues_{f['resolution']}_{scope}_{stamp}.xlsx"
    resp = HttpResponse(
        buf.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


@login_required
@require_POST
def issues_save(request):
    """Save the operator's Action + Remark on one flagged line (Issues page)."""
    from .services import lines_store
    res = lines_store.update_action(
        request.POST.get('line_id'),
        request.POST.get('action', ''),
        request.POST.get('remark', ''))
    return JsonResponse(res)


@login_required
@require_POST
def issues_save_bulk(request):
    """Apply one Action + Remark to many flagged lines at once (Issues page)."""
    from .services import lines_store
    res = lines_store.update_action_bulk(
        request.POST.getlist('line_ids'),
        request.POST.get('action', ''),
        request.POST.get('remark', ''))
    return JsonResponse(res)


@login_required
@require_POST
def issues_fix_ean(request):
    """Post-lock EAN correction on a NOT_IN_MASTER line (Issues page).

    The lock-first model: the wrong EAN is already recorded. Operator types the
    correct EAN here → the line re-resolves against the item master, OUR pricing
    is recomputed with the engine's own helpers, the line flips to OK/MISMATCH,
    and the wrong EAN is kept as ``received_ean`` for the vendor-escalation audit.
    """
    from .services import lines_store
    res = lines_store.apply_issue_ean_fix(
        request.POST.get('line_id'),
        (request.POST.get('correct_ean', '') or '').strip())
    return JsonResponse(res, status=200 if res.get('ok') else 400)


# ── TAT (24h SLA) — breaches + reasons ──────────────────────────────────────

def _tat_filters(request) -> dict:
    return {
        'marketplace': request.GET.get('marketplace', '').strip(),
        'segment': request.GET.get('segment', '').strip(),
        'q': request.GET.get('q', '').strip(),
        'status': request.GET.get('status', 'pending').strip() or 'pending',
        'date_from': request.GET.get('date_from', '').strip(),
        'date_to': request.GET.get('date_to', '').strip(),
        'run': request.GET.get('run', '').strip(),
    }


@login_required
def tat(request):
    """TAT / SLA page — orders uploaded later than 1 working day after the PO
    date. Operator records a breach reason here (reactive). Filterable run-wise."""
    from .services import tat_store
    data = tat_store.breaches(**_tat_filters(request))
    data['reasons'] = tat_store.REASONS
    if _is_ajax(request):
        return render(request, 'online_b2b/_tat_rows.html', {'d': data})
    data['runs'] = order_db.recent_runs(50)
    return render(request, 'online_b2b/tat.html', {'d': data})


@login_required
@require_POST
def tat_save(request):
    """Set / clear the TAT breach reason for one order."""
    from .services import tat_store
    res = tat_store.set_reason(
        request.POST.get('order_id'),
        request.POST.get('reason_code', ''),
        request.POST.get('note', ''),
        by=getattr(request.user, 'username', ''))
    return JsonResponse(res, status=200 if res.get('ok') else 400)


@login_required
def tat_export(request):
    """Download the filtered TAT breaches as .xlsx (respects all filters)."""
    import datetime as _dt
    import io as _io

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    from .services import tat_store
    f = _tat_filters(request)
    data = tat_store.breaches(limit=100000, **f)
    rows = data.get('rows', []) if data.get('ok') else []
    cols = [('po_date', 'PO Date'), ('run_ts', 'Uploaded'),
            ('segment', 'Segment'), ('marketplace', 'Channel'), ('po', 'PO / SO'),
            ('location', 'Location'), ('qty', 'Qty'), ('order_value', 'Value'),
            ('wd_late', 'Working days taken'), ('days_over', 'Days over TAT'),
            ('reason_code', 'Reason'), ('note', 'Note'),
            ('reason_by', 'By'), ('reason_at', 'Reason at')]
    wb = Workbook(); ws = wb.active; ws.title = 'TAT breaches'
    hf = Font(bold=True, color='FFFFFF'); navy = PatternFill('solid', fgColor='1A237E')
    for c, (_k, h) in enumerate(cols, 1):
        cell = ws.cell(1, c, h)
        cell.font = hf; cell.fill = navy; cell.alignment = Alignment(horizontal='center')
    for r, row in enumerate(rows, 2):
        for c, (k, _h) in enumerate(cols, 1):
            v = row.get(k)
            ws.cell(r, c, str(v) if k in ('run_ts', 'po_date', 'reason_at') and v is not None else v)
    for col in ws.columns:
        L = col[0].column_letter
        w = max((len(str(c.value or '')) for c in col), default=8)
        ws.column_dimensions[L].width = min(w + 2, 48)
    buf = _io.BytesIO(); wb.save(buf); buf.seek(0)
    stamp = _dt.datetime.now().strftime('%Y%m%d_%H%M%S')
    fname = f"tat_breaches_{f['status']}_{stamp}.xlsx"
    resp = HttpResponse(
        buf.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


@login_required
def sku_summary(request):
    """SKU-wise validation rollup — group by (Item No, EAN): qty + line-counts
    per status (OK / Mismatch / Not-in-master), MRP comparison, drill-down.
    READ-ONLY aggregation over the order_lines_full view (no DB changes)."""
    data = order_db.sku_summary(
        marketplace=request.GET.get('marketplace', '').strip(),
        q=request.GET.get('q', '').strip(),
        date_from=request.GET.get('from', '').strip(),
        date_to=request.GET.get('to', '').strip(),
        issues_only=request.GET.get('issues') == '1')
    if _is_ajax(request):
        return render(request, 'online_b2b/_sku_rows.html', {'d': data})
    return render(request, 'online_b2b/sku_summary.html', {'d': data})


@login_required
def sku_summary_lines(request):
    """Drill-down: the individual PO-lines behind one SKU (AJAX partial)."""
    data = order_db.sku_lines(request.GET.get('item_no', '').strip(),
                              request.GET.get('ean', '').strip())
    return render(request, 'online_b2b/_sku_drill.html', {'d': data})


@login_required
def export(request):
    """Stream the current filtered Orders view as an .xlsx."""
    import io
    from datetime import datetime

    from openpyxl import Workbook

    f = _filters(request)
    rows = order_db.orders_for_export(**f)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Online B2B Orders'
    cols = ['Run', 'Run Time', 'Marketplace', 'PO', 'Location', 'Warehouse',
            'PO Date', 'Exp Date', 'Type', 'Items', 'Qty', 'Order Value',
            'Days to Expiry']
    ws.append(cols)
    for r in rows:
        ws.append([
            r.get('run_id'), str(r.get('run_ts') or ''), r.get('marketplace'),
            r.get('po'), r.get('location'), r.get('warehouse'),
            str(r.get('po_date') or ''), str(r.get('exp_date') or ''),
            r.get('order_type'), r.get('items'), r.get('qty'), r.get('value'),
            r.get('days_to_expiry'),
        ])
    ws.freeze_panes = 'A2'
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"online_b2b_orders_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    resp = FileResponse(buf, as_attachment=True, filename=fname,
                        content_type='application/vnd.openxmlformats-'
                        'officedocument.spreadsheetml.sheet')
    return resp


_UPLOADS = _MEDIA / 'b2b_uploads'
_ERP = _MEDIA / 'b2b_erp'


def _token_dir(token: str) -> Path:
    """Resolve a token folder, guarding against path traversal."""
    base = _UPLOADS.resolve()
    d = (_UPLOADS / token).resolve()
    if d != base and base not in d.parents:
        raise Http404()
    return d


def _load_meta(token: str):
    d = _token_dir(token)
    p = d / 'meta.json'
    if not p.exists():
        return None, d
    try:
        return json.loads(p.read_text(encoding='utf-8')), d
    except Exception:
        return None, d


def _preview_sig(meta) -> str:
    """Cache key for a token's processed preview — everything that changes the
    engine output. Notably includes ``ean_fixes`` so an EAN-fix re-validate
    busts the cache automatically."""
    import hashlib
    key = json.dumps({
        'mp': meta.get('marketplace'), 'wh': meta.get('warehouse'),
        'mg': meta.get('margin_pct'), 'files': sorted(meta.get('files') or []),
        'fixes': meta.get('ean_fixes') or {},
    }, sort_keys=True)
    return hashlib.md5(key.encode()).hexdigest()


def _cached_preview(token, d, meta, force=False):
    """Run the engine preview ONCE per (files + EAN-fixes) and cache its JSON
    result on the token, so the review page — and every reload — renders
    instantly instead of re-running the engine. Re-runs only when the signature
    changes (new files / EAN fix) or ``force``."""
    cache = d / 'preview.json'
    sig = _preview_sig(meta)
    if not force and cache.exists():
        try:
            blob = json.loads(cache.read_text(encoding='utf-8'))
            if blob.get('sig') == sig and (blob.get('res') or {}).get('ok'):
                return blob['res']
        except Exception:  # noqa: BLE001 — corrupt cache → just re-run
            pass
    res = engine_bridge.preview(
        meta['marketplace'], [str(d / n) for n in meta['files']],
        warehouse=meta['warehouse'], margin_pct=meta['margin_pct'] / 100.0,
        ean_fixes=meta.get('ean_fixes'))
    if res.get('ok'):
        try:
            cache.write_text(json.dumps({'sig': sig, 'res': res}, default=str),
                             encoding='utf-8')
        except Exception:  # noqa: BLE001 — caching is best-effort
            pass
    return res


@login_required
def upload(request):
    """Phase 1: stash the uploaded PO(s) under a token and go to Review.
    Nothing is processed or written here."""
    if request.method == 'POST':
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            token = uuid.uuid4().hex[:12]
            up_dir = _UPLOADS / token
            up_dir.mkdir(parents=True, exist_ok=True)
            saved = []
            for f in form.cleaned_data['po_files']:
                dest = up_dir / Path(f.name).name
                with open(dest, 'wb') as out:
                    for chunk in f.chunks():
                        out.write(chunk)
                saved.append(dest.name)
            marketplace = form.cleaned_data['marketplace']
            # Blank margin → the marketplace's configured default landing rate
            # (Flipkart 77, Blink 70) — same as the Tkinter app's pre-fill.
            margin = (form.cleaned_data.get('margin_pct')
                      or engine_bridge.default_margin_pct(marketplace))
            meta = {
                'marketplace': marketplace,
                'warehouse': form.cleaned_data['warehouse'],
                'margin_pct': margin,
                'files': saved,
            }
            (up_dir / 'meta.json').write_text(json.dumps(meta), encoding='utf-8')
            # AJAX: do the import HERE (engine preview, cached) so the upload
            # page can show a real progress overlay + a definitive "✓ Imported"
            # completion before navigating — and the review page then loads from
            # the cache instantly. Non-JS clients fall back to the plain redirect
            # (review runs the same cached preview).
            if _is_ajax(request):
                from django.urls import reverse
                res = _cached_preview(token, up_dir, meta)
                if not res.get('ok'):
                    return JsonResponse(
                        {'ok': False, 'error': res.get('error', 'Import failed.')},
                        status=200)
                s = res.get('summary', {})
                return JsonResponse({
                    'ok': True, 'review_url': reverse('b2b_review', args=[token]),
                    'pos': s.get('pos', 0), 'lines': s.get('lines', 0),
                    'affected': s.get('affected', 0),
                    'warnings': len(res.get('warnings') or []),
                })
            return redirect('b2b_review', token=token)
        elif _is_ajax(request):
            errs = '; '.join(f"{k}: {v.as_text()}" for k, v in form.errors.items())
            return JsonResponse({'ok': False, 'error': errs or 'Invalid form.'},
                                status=200)
    else:
        form = UploadForm()
    return render(request, 'online_b2b/upload.html',
                  {'form': form,
                   'margin_defaults': json.dumps(engine_bridge.margin_defaults())})


@login_required
def review(request, token):
    """Phase 2: process in memory (no DB write) and show the review page."""
    meta, d = _load_meta(token)
    if not meta:
        raise Http404("Upload not found or expired.")
    # Load the cached preview (the AJAX upload already ran + cached it → instant);
    # falls back to a fresh run for non-JS uploads or a busted cache.
    res = _cached_preview(token, d, meta)
    has_preview = bool(res.get('output_path') and os.path.exists(res['output_path']))
    # Re-attach any saved decisions so the Affected rows show them (esp. when locked).
    decisions = meta.get('decisions') or {}
    for ln in res.get('affected') or []:
        key = f"{ln.get('po', '')}|{ln.get('item_no', '')}|{ln.get('ean', '')}"
        ln['decision'] = decisions.get(key, {})
    # Make the Line Items tab the FINAL ready-to-go view: attach each line's
    # disposition (the action taken on affected ones) so the operator sees OK /
    # Included / Override / Excluded per line before locking — same as the
    # post-lock Issues page, just earlier.
    for ln in res.get('lines') or []:
        key = f"{ln.get('po', '')}|{ln.get('item_no', '')}|{ln.get('ean', '')}"
        ln['decision'] = decisions.get(key, {})
    # NOT_IN_MASTER lines (wrong / unknown EAN) — the operator corrects the EAN
    # here, then re-validate resolves them against the DB master.
    nim_lines = [ln for ln in (res.get('affected') or [])
                 if ln.get('status') == 'NOT_IN_MASTER']
    # Auto-corrected lines: a previously-seen wrong EAN resolved through the
    # historical map. Flag them (TEMPORARY fix) + how many times that wrong EAN
    # was received before, so the team can escalate to the vendor.
    recv_counts = order_db.ean_correction_counts()
    auto_fixed = []
    for ln in (res.get('lines') or []):
        re = ln.get('received_ean')
        if re:
            ln['recv_count'] = recv_counts.get(re, 0)
            auto_fixed.append(ln)
    # Engine exceptions (vendor deals / overrides / remaps) — poke the operator
    # that these lines deviate from the flat marketplace rule.
    exc_count = sum(1 for ln in (res.get('lines') or [])
                    if (ln.get('exception_label') or '').strip())
    return render(request, 'online_b2b/review.html',
                  {'token': token, 'meta': meta, 'r': res,
                   'has_preview': has_preview,
                   'locked': bool(meta.get('locked')),
                   'run_id': meta.get('run_id'),
                   'exc_count': exc_count, 'nim_lines': nim_lines,
                   'auto_fixed': auto_fixed,
                   'margin': meta.get('margin_pct')})


@login_required
@require_POST
def confirm(request, token):
    """Phase 3: re-process and PUSH headers + lines to MySQL.

    Answers JSON when called via fetch (``X-Requested-With``) so the review page
    can lock in place with a progress bar — no full reload. Falls back to the
    classic redirect flow when JS is off."""
    ajax = request.headers.get('x-requested-with') == 'XMLHttpRequest'
    meta, d = _load_meta(token)
    if not meta:
        raise Http404("Upload not found or expired.")
    paths = [str(d / n) for n in meta['files']]

    # Per-affected-line operator decision (Action + Override CP + Remark),
    # keyed po|item|ean. Action ∈ INCLUDE / OVERRIDE / EXCLUDE.
    keys = request.POST.getlist('aff_key')
    acts = request.POST.getlist('aff_action')
    ocps = request.POST.getlist('aff_override_cp')
    rems = request.POST.getlist('aff_remark')
    actions = {}
    for i, k in enumerate(keys):
        a = (acts[i] if i < len(acts) else '').strip()
        ocp = (ocps[i] if i < len(ocps) else '').strip()
        r = (rems[i] if i < len(rems) else '').strip()
        if k and (a or r):
            actions[k] = {'action': a, 'remark': r, 'override_cp': ocp}

    res = engine_bridge.confirm(
        meta['marketplace'], paths, warehouse=meta['warehouse'],
        margin_pct=meta['margin_pct'] / 100.0, actions=actions,
        ean_fixes=meta.get('ean_fixes'))

    if not res.get('ok'):
        err = res.get('error', 'Push failed.')
        if ajax:
            return JsonResponse({'ok': False, 'error': err}, status=400)
        messages.error(request, err)
        return redirect('b2b_review', token=token)

    run_id = res.get('run_id')
    if run_id is None:
        msg = "All POs were already uploaded — nothing new to push."
        if ajax:
            from django.urls import reverse
            return JsonResponse({'ok': True, 'run_id': None, 'message': msg,
                                 'redirect': reverse('b2b_dashboard')})
        messages.info(request, msg)
        return redirect('b2b_dashboard')

    # Build the D365 dump ONCE at lock time and keep it next to the run sidecar
    # (web-owned, survives upload-temp cleanup) so it's retrievable later from
    # Orders → Run #N even if the review tab is closed. Never blocks the lock.
    d365_path = _persist_d365(
        run_id, meta['marketplace'], paths,
        meta['warehouse'], meta['margin_pct'] / 100.0, actions,
        ean_fixes=meta.get('ean_fixes'))
    _save_run_index(run_id, {
        'output_path': res.get('output_path'),
        'marketplace': meta['marketplace'],
        'summary': res.get('summary'),
        'warnings': res.get('warnings'),
        'd365_path': d365_path,
    })
    # Lock the decisions on the token so the review page now offers Generate D365.
    meta['decisions'] = actions
    meta['run_id'] = run_id
    meta['locked'] = True
    (d / 'meta.json').write_text(json.dumps(meta), encoding='utf-8')
    pos = res['summary']['pos']
    lines = res.get('lines_recorded', 0)
    if ajax:
        from django.urls import reverse
        return JsonResponse({
            'ok': True, 'run_id': run_id, 'pos': pos, 'lines': lines,
            'has_d365': bool(d365_path),
            'run_url': reverse('b2b_run_detail', args=[run_id]),
            'd365_url': reverse('b2b_generate_d365', args=[token]),
            'message': f"Locked & recorded {pos} PO(s), {lines} line(s).",
        })
    messages.success(
        request, f"🔒 Locked & recorded {pos} PO(s), {lines} line(s). "
        "Now generate the D365 dump.")
    return redirect('b2b_review', token=token)


@login_required
@require_POST
def generate_d365(request, token):
    """Build + download the ERP D365 dump from the LOCKED decisions (Excludes
    dropped, Overrides repriced). Engine + full SO Workbook untouched."""
    meta, d = _load_meta(token)
    if not meta:
        raise Http404("Upload not found or expired.")
    if not meta.get('locked'):
        messages.error(request, "Lock the decisions first, then generate D365.")
        return redirect('b2b_review', token=token)
    paths = [str(d / n) for n in meta['files']]
    out_path = d / f"{str(meta['marketplace']).lower()}_d365_decided.xlsx"
    res = engine_bridge.generate_d365(
        meta['marketplace'], paths, str(out_path),
        warehouse=meta['warehouse'], margin_pct=meta['margin_pct'] / 100.0,
        actions=meta.get('decisions') or {}, ean_fixes=meta.get('ean_fixes'))
    if not res.get('ok') or not os.path.exists(res.get('d365_path', '')):
        messages.error(request, res.get('error', 'D365 generation failed.'))
        return redirect('b2b_review', token=token)
    return FileResponse(
        open(res['d365_path'], 'rb'), as_attachment=True,
        filename=f"{meta['marketplace']}_D365_import.xlsx")


@login_required
@require_POST
def save_decision(request, token):
    """Persist ONE affected line's decision (Action / Override CP / Remark) to
    the upload session as the operator sets it — so decisions survive an EAN
    re-validate and the final lock just commits what's saved. Returns JSON."""
    meta, d = _load_meta(token)
    if not meta:
        return JsonResponse({'ok': False, 'error': 'expired'}, status=404)
    if meta.get('locked'):
        return JsonResponse({'ok': False, 'error': 'locked'})
    key = (request.POST.get('key') or '').strip()
    if not key:
        return JsonResponse({'ok': False, 'error': 'no key'})
    action = (request.POST.get('action') or '').strip()
    ocp = (request.POST.get('override_cp') or '').strip()
    remark = (request.POST.get('remark') or '').strip()
    decisions = dict(meta.get('decisions') or {})
    if action or remark or ocp:
        decisions[key] = {'action': action, 'remark': remark, 'override_cp': ocp}
    else:
        decisions.pop(key, None)
    meta['decisions'] = decisions
    (d / 'meta.json').write_text(json.dumps(meta), encoding='utf-8')
    return JsonResponse({'ok': True, 'saved': len(decisions)})


@login_required
@require_POST
def fix_ean(request, token):
    """Correct the EAN on NOT_IN_MASTER lines: validate each 'correct EAN'
    against the DB master, stash the fixes on the upload session, and re-run the
    review so they resolve. The wrong EAN is kept (becomes ``received_ean`` at
    lock). Persistent: a repeat wrong EAN auto-resolves on future POs."""
    meta, d = _load_meta(token)
    if not meta:
        raise Http404("Upload not found or expired.")
    if meta.get('locked'):
        messages.error(request, "Already locked — EANs can't be changed.")
        return redirect('b2b_review', token=token)
    from .services import item_master_loader as iml
    wrongs = request.POST.getlist('nim_ean')
    fixes = request.POST.getlist('nim_fix')
    ean_fixes = dict(meta.get('ean_fixes') or {})
    applied, errors = 0, []
    for i, w in enumerate(wrongs):
        w = (w or '').strip()
        c = (fixes[i] if i < len(fixes) else '').strip()
        if not w or not c:
            continue
        hit = iml.resolve_in_master(c)
        if not hit:
            errors.append(f"'{c}' is not in the item master — can't use it to "
                          f"correct EAN {w}.")
            continue
        ean_fixes[w] = hit['ean'] or hit['item_no']
        applied += 1
    if applied:
        meta['ean_fixes'] = ean_fixes
        (d / 'meta.json').write_text(json.dumps(meta), encoding='utf-8')
        messages.success(request, f"✓ Corrected {applied} EAN(s) — re-validating "
                         "against the master.")
    for e in errors:
        messages.error(request, e)
    return redirect('b2b_review', token=token)


@login_required
@require_POST
def discard(request, token):
    d = _token_dir(token)
    if d.exists():
        shutil.rmtree(d, ignore_errors=True)
    messages.info(request, "Upload discarded.")
    return redirect('b2b_dashboard')


def _full_name(name: str) -> str:
    """Clear download name for the FULL workbook (all sheets) so it's never
    confused with the headers-only ``*_d365.xlsx`` import package."""
    for s in ('_so_', '_to_'):
        if s in name:
            return name.replace(s, '_full_')
    return 'full_' + name


def _full_workbook(outdir: Path):
    """The full data workbook in a folder = the newest .xlsx that is NOT the
    ``*_d365.xlsx`` D365 import sibling."""
    if not outdir.exists():
        return None
    files = sorted([p for p in outdir.glob('*.xlsx')
                    if not p.stem.endswith('_d365')],
                   key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


@login_required
def review_download(request, token):
    """Download the FULL preview workbook (Summary / Validation / Raw Data /
    Headers / Lines) — explicitly NOT the headers-only *_d365.xlsx package."""
    d = _token_dir(token)
    f = _full_workbook(d / 'output')
    if not f:
        raise Http404("Preview workbook not found.")
    return FileResponse(open(f, 'rb'), as_attachment=True,
                        filename=_full_name(f.name))


# ── Bulk import (ERP Sales Orders) ──────────────────────────────────────

def _erp_file(token: str):
    base = _ERP.resolve()
    d = (_ERP / token).resolve()
    if d != base and base not in d.parents:
        raise Http404()
    nm = d / 'name.txt'
    if not nm.exists():
        return None
    return d / nm.read_text(encoding='utf-8').strip()


@login_required
def bulk_upload(request):
    """Upload an ERP 'Sales Orders' header export → stash → go to review."""
    if request.method == 'POST' and request.FILES.get('erp_file'):
        token = uuid.uuid4().hex[:12]
        up = _ERP / token
        up.mkdir(parents=True, exist_ok=True)
        f = request.FILES['erp_file']
        dest = up / Path(f.name).name
        with open(dest, 'wb') as out:
            for chunk in f.chunks():
                out.write(chunk)
        (up / 'name.txt').write_text(dest.name, encoding='utf-8')
        return redirect('b2b_bulk_review', token=token)
    return render(request, 'online_b2b/bulk_upload.html')


@login_required
def bulk_review(request, token):
    """Preview the parsed ERP rows (new vs already-imported) — no DB write."""
    fp = _erp_file(token)
    if not fp or not fp.exists():
        raise Http404("Upload not found or expired.")
    res = erp_import.preview(str(fp))
    return render(request, 'online_b2b/bulk_review.html',
                  {'token': token, 'r': res})


@login_required
@require_POST
def bulk_confirm(request, token):
    """Import the NEW ERP rows into the order DB (segment=Offline)."""
    fp = _erp_file(token)
    if not fp or not fp.exists():
        raise Http404("Upload not found or expired.")
    res = erp_import.do_import(str(fp))
    if not res.get('ok'):
        messages.error(request, res.get('error', 'Import failed.'))
        return redirect('b2b_bulk_review', token=token)
    if not res.get('imported'):
        messages.info(request, "All orders were already imported — nothing new.")
        return redirect('b2b_orders')
    messages.success(
        request, f"Imported {res['imported']} order(s) "
        f"({res.get('skipped', 0)} already present).")
    return redirect('b2b_orders')


@login_required
def run_detail(request, run_id):
    data = order_db.run_detail(int(run_id))
    idx = _load_run_index(run_id)
    has_file = bool(idx.get('output_path') and os.path.exists(idx['output_path']))
    has_d365 = bool(idx.get('d365_path') and os.path.exists(idx['d365_path']))
    return render(request, 'online_b2b/run_detail.html', {
        'run_id': run_id, 'd': data, 'idx': idx, 'has_file': has_file,
        'has_d365': has_d365,
    })


@login_required
def download_d365(request, run_id):
    """Re-download the decided D365 dump saved for this run at lock time
    (Excludes dropped, Overrides repriced). A derived static file — serving it
    reads nothing from and writes nothing to the DB."""
    idx = _load_run_index(run_id)
    path = idx.get('d365_path')
    if not path or not os.path.exists(path):
        raise Http404("No D365 dump saved for this run.")
    mkt = idx.get('marketplace', 'D365')
    return FileResponse(open(path, 'rb'), as_attachment=True,
                        filename=f"{mkt}_D365_import.xlsx")


@login_required
def download(request, run_id):
    """Download the FULL SO workbook for a run (all sheets) — the engine's main
    output, not the headers-only *_d365.xlsx package."""
    idx = _load_run_index(run_id)
    path = idx.get('output_path')
    if not path or not os.path.exists(path):
        raise Http404("Output file not found for this run.")
    return FileResponse(open(path, 'rb'), as_attachment=True,
                        filename=_full_name(os.path.basename(path)))


# ── Item Master (DB) ────────────────────────────────────────────────────────
# Upload the two ERP exports (Items + Item M.R.P.) → preview → rebuild the
# `item_master` table. The engine reads this DB master via DBMasterLoader — no
# manual Excel. The page also browses the live master (overview + search).

_IM_UPLOADS = _MEDIA / 'b2b_item_master'


def _im_token_dir(token: str) -> Path:
    base = _IM_UPLOADS.resolve()
    d = (_IM_UPLOADS / token).resolve()
    if d != base and base not in d.parents:
        raise Http404()
    return d


class ItemMasterView(LoginRequiredMixin, TemplateView):
    """`/b2b/item-master/` — status + upload/refresh + browsable overview."""
    template_name = 'online_b2b/item_master.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        from .services import item_master_loader as iml
        ctx['status'] = iml.status()
        ctx['overview'] = iml.list_items(self.request.GET.get('q', ''), limit=100)
        return ctx


@login_required
@require_POST
def item_master_upload(request):
    """Stash the two uploaded files under a token → preview. No DB write."""
    items_f = request.FILES.get('items_file')
    mrp_f = request.FILES.get('mrp_file')
    if not items_f or not mrp_f:
        messages.error(request, "Choose BOTH files — the Items export and the "
                                "Item M.R.P. export.")
        return redirect('b2b_item_master')
    token = uuid.uuid4().hex[:12]
    d = _IM_UPLOADS / token
    d.mkdir(parents=True, exist_ok=True)
    items_path = d / ('items' + Path(items_f.name).suffix)
    mrp_path = d / ('mrp' + Path(mrp_f.name).suffix)
    for f, dest in ((items_f, items_path), (mrp_f, mrp_path)):
        with open(dest, 'wb') as out:
            for chunk in f.chunks():
                out.write(chunk)
    (d / 'meta.json').write_text(json.dumps({
        'items_name': items_f.name, 'mrp_name': mrp_f.name,
        'items_path': items_path.name, 'mrp_path': mrp_path.name,
    }), encoding='utf-8')
    return redirect('b2b_item_master_preview', token=token)


@login_required
def item_master_preview(request, token):
    """Compute the master from the two files in-memory (no write) and show the
    stats / warnings / sample for the operator to confirm."""
    d = _im_token_dir(token)
    mp = d / 'meta.json'
    if not mp.exists():
        raise Http404("Upload not found or expired.")
    meta = json.loads(mp.read_text(encoding='utf-8'))
    from .services import item_master_loader as iml
    try:
        rows, stats, warnings = iml.build_rows(
            str(d / meta['items_path']), str(d / meta['mrp_path']))
    except Exception as e:  # noqa: BLE001
        messages.error(request, f"Could not read the files: {type(e).__name__}: {e}")
        return redirect('b2b_item_master')
    return render(request, 'online_b2b/item_master_preview.html', {
        'token': token, 'meta': meta, 'stats': stats, 'warnings': warnings,
        'sample': rows[:15], 'current': iml.status(),
        'diff': iml.diff_against_current(rows),
    })


@login_required
@require_POST
def item_master_confirm(request, token):
    """Rebuild item_master from the two files in one transaction (full replace).
    Seeds the durable Swiggy map once if empty. No engine change, no order data
    touched."""
    d = _im_token_dir(token)
    mp = d / 'meta.json'
    if not mp.exists():
        raise Http404("Upload not found or expired.")
    meta = json.loads(mp.read_text(encoding='utf-8'))
    from .services import item_master_loader as iml
    try:
        if not iml.load_swiggy_map():
            from online_po_processor.config.paths import get_bundled_master_path
            bm = get_bundled_master_path()
            if bm:
                iml.seed_swiggy_from_excel(str(bm))
        rows, stats, _ = iml.build_rows(
            str(d / meta['items_path']), str(d / meta['mrp_path']))
        res = iml.replace_item_master(rows)
    except Exception as e:  # noqa: BLE001
        messages.error(request, f"Rebuild failed: {type(e).__name__}: {e}")
        return redirect('b2b_item_master_preview', token=token)
    shutil.rmtree(d, ignore_errors=True)
    messages.success(
        request, f"✓ Item master rebuilt — {res['rows']} items now live "
        f"({stats['swiggy_mapped']} Swiggy-mapped · batch {res['batch_id']}).")
    return redirect('b2b_item_master')


@login_required
@require_POST
def item_master_discard(request, token):
    d = _im_token_dir(token)
    if d.exists():
        shutil.rmtree(d, ignore_errors=True)
    messages.info(request, "Upload discarded.")
    return redirect('b2b_item_master')


def _im_row_json(r: dict) -> dict:
    """JSON-safe item_master row (dates → str, Decimal → float)."""
    def s(v):
        return '' if v is None else str(v)
    return {
        'item_no': s(r.get('item_no')), 'ean': s(r.get('ean')),
        'description': s(r.get('description')), 'gst_code': s(r.get('gst_code')),
        'hsn': s(r.get('hsn')),
        'mrp': None if r.get('mrp') is None else float(r['mrp']),
        'mrp_start': s(r.get('mrp_start')), 'mrp_end': s(r.get('mrp_end')),
    }


@login_required
def item_master_search(request):
    """As-you-type search over item_master → JSON (item no / EAN / description /
    Swiggy). Read-only."""
    from .services import item_master_loader as iml
    res = iml.list_items(request.GET.get('q', ''), limit=100)
    return JsonResponse({
        'total': res['total'], 'shown': res['shown'], 'q': res['q'],
        'rows': [_im_row_json(r) for r in res['rows']],
    })


# ── GT Select (Offline parent — D365-finalised; import headers + lines) ─────

_GTS_UPLOADS = _MEDIA / 'b2b_gt_select'


def _gts_token_dir(token: str) -> Path:
    base = _GTS_UPLOADS.resolve()
    d = (_GTS_UPLOADS / token).resolve()
    if d != base and base not in d.parents:
        raise Http404()
    return d


class OfflineProcessView(LoginRequiredMixin, TemplateView):
    """Offline 'Process PO' — one entry that lists the Offline **parent**
    marketplaces (and MT's **children**), each routing to that channel's
    upload/import. Mirrors the online Process-PO marketplace picker so the two
    segments feel the same; replaces the scattered per-channel links."""
    template_name = 'online_b2b/offline_process.html'

    def get_context_data(self, **kwargs):
        from django.urls import reverse
        ctx = super().get_context_data(**kwargs)
        ctx['parents'] = [
            {'key': 'MT', 'name': 'MT · Modern Trade', 'tag': 'live',
             'desc': 'Retail chains — master-driven SO generation. Pick a child.',
             'children': [
                 {'key': 'SS', 'name': 'Shoppers Stop (SS)', 'tag': 'live',
                  'url': reverse('shoppers_stop')},
                 {'key': 'HG', 'name': 'HG', 'tag': 'soon'},
                 {'key': 'NT', 'name': 'NT', 'tag': 'soon'},
                 {'key': 'HB', 'name': 'HB', 'tag': 'soon'},
                 {'key': 'LL', 'name': 'LL', 'tag': 'soon'},
                 {'key': 'BN', 'name': 'BN', 'tag': 'soon'},
             ]},
            {'key': 'GT Mass', 'name': 'GT Mass', 'tag': 'live',
             'url': reverse('index'), 'desc': 'General trade mass — dump processing.',
             'children': []},
            {'key': 'GT Select', 'name': 'GT Select', 'tag': 'live',
             'url': reverse('b2b_gt_select'),
             'desc': 'D365-finalised — import Sales Orders + Lines.', 'children': []},
            {'key': 'EKA', 'name': 'EKA', 'tag': 'soon', 'desc': 'Coming soon.',
             'children': []},
            {'key': 'CSD', 'name': 'CSD', 'tag': 'soon', 'desc': 'Coming soon.',
             'children': []},
        ]
        return ctx


class GtSelectView(LoginRequiredMixin, TemplateView):
    """`/b2b/gt-select/` — upload the D365 Sales Orders + Sales Lines exports."""
    template_name = 'online_b2b/gt_select.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['kpis'] = order_db.segment_kpis('Offline')
        return ctx


@login_required
@require_POST
def gt_select_upload(request):
    """Stash the two D365 exports under a token → preview."""
    hf = request.FILES.get('headers_file')
    lf = request.FILES.get('lines_file')
    if not hf or not lf:
        messages.error(request, "Choose BOTH files — Sales Orders (headers) and "
                                "Sales Lines.")
        return redirect('b2b_gt_select')
    token = uuid.uuid4().hex[:12]
    d = _GTS_UPLOADS / token
    d.mkdir(parents=True, exist_ok=True)
    hp = d / ('headers' + Path(hf.name).suffix)
    lp = d / ('lines' + Path(lf.name).suffix)
    for f, dest in ((hf, hp), (lf, lp)):
        with open(dest, 'wb') as out:
            for chunk in f.chunks():
                out.write(chunk)
    (d / 'meta.json').write_text(json.dumps({
        'headers_name': hf.name, 'lines_name': lf.name,
        'headers_path': hp.name, 'lines_path': lp.name,
    }), encoding='utf-8')
    return redirect('b2b_gt_select_preview', token=token)


@login_required
def gt_select_preview(request, token):
    """Parse + join + dedup (no DB write) for the review page."""
    d = _gts_token_dir(token)
    mp = d / 'meta.json'
    if not mp.exists():
        raise Http404("Upload not found or expired.")
    meta = json.loads(mp.read_text(encoding='utf-8'))
    from .services import gt_select_import as gts
    pv = gts.preview(str(d / meta['headers_path']), str(d / meta['lines_path']))
    if not pv.get('ok'):
        messages.error(request, pv.get('error', 'Could not read the files.'))
        return redirect('b2b_gt_select')
    return render(request, 'online_b2b/gt_select_preview.html', {
        'token': token, 'meta': meta, 'r': pv,
    })


@login_required
@require_POST
def gt_select_confirm(request, token):
    """Import the NEW GT Select orders + lines under one IMPORT run."""
    d = _gts_token_dir(token)
    mp = d / 'meta.json'
    if not mp.exists():
        raise Http404("Upload not found or expired.")
    meta = json.loads(mp.read_text(encoding='utf-8'))
    from .services import gt_select_import as gts
    res = gts.do_import(str(d / meta['headers_path']), str(d / meta['lines_path']))
    if not res.get('ok'):
        messages.error(request, res.get('error', 'Import failed.'))
        return redirect('b2b_gt_select_preview', token=token)
    shutil.rmtree(d, ignore_errors=True)
    if not res.get('imported'):
        messages.info(request, "All orders were already imported — nothing new.")
        return redirect('b2b_orders')
    messages.success(
        request, f"✓ Imported {res['imported']} GT Select order(s) · "
        f"{res['lines']} line(s) ({res.get('skipped', 0)} already present).")
    if res.get('run_id'):
        return redirect('b2b_run_detail', run_id=res['run_id'])
    return redirect('b2b_orders')


@login_required
@require_POST
def gt_select_discard(request, token):
    d = _gts_token_dir(token)
    if d.exists():
        shutil.rmtree(d, ignore_errors=True)
    messages.info(request, "Upload discarded.")
    return redirect('b2b_gt_select')


@login_required
@require_POST
def item_master_add(request):
    """Add/overwrite ONE item by hand — written to item_master (flagged
    batch_id='manual', durable across rebuilds); a typed Swiggy SKU goes to
    channel_sku_map."""
    from .services import item_master_loader as iml
    res = iml.upsert_manual_item({
        'item_no': request.POST.get('item_no', ''),
        'ean': request.POST.get('ean', ''),
        'description': request.POST.get('description', ''),
        'gst_code': request.POST.get('gst_code', ''),
        'hsn': request.POST.get('hsn', ''),
        'mrp': request.POST.get('mrp', ''),
        'mrp_start': request.POST.get('mrp_start', ''),
        'mrp_end': request.POST.get('mrp_end', ''),
        'swiggy_sku_code': request.POST.get('swiggy_sku_code', ''),
    })
    if not res.get('ok'):
        messages.error(request, res.get('error', 'Could not add the item.'))
    else:
        messages.success(request, f"✓ Item {res['item_no']} saved to the master "
                         "(kept across rebuilds until the ERP export includes it).")
    return redirect('b2b_item_master')


# ── Ship-To Mapping (DB-backed; the bundled Ship to B2B.xlsx is retired) ─────
_STM_UPLOADS = _MEDIA / 'b2b_ship_to'


def _stm_token_dir(token: str) -> Path:
    base = _STM_UPLOADS.resolve()
    d = (_STM_UPLOADS / token).resolve()
    if d != base and base not in d.parents:
        raise Http404()
    return d


class ShipToView(LoginRequiredMixin, TemplateView):
    """`/b2b/ship-to/` — status + upload/replace + browse/search + add/delete."""
    template_name = 'online_b2b/ship_to.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        from .services import mapping_store as ms
        ctx['status'] = ms.status()
        ctx['overview'] = ms.list_mappings(
            self.request.GET.get('party', ''), self.request.GET.get('q', ''),
            limit=300)
        return ctx


@login_required
@require_POST
def ship_to_upload(request):
    """Stash an uploaded Ship-To B2B Excel under a token → preview. No DB write."""
    f = request.FILES.get('mapping_file')
    if not f:
        messages.error(request, "Choose the Ship-To B2B Excel to upload.")
        return redirect('b2b_ship_to')
    token = uuid.uuid4().hex[:12]
    d = _STM_UPLOADS / token
    d.mkdir(parents=True, exist_ok=True)
    dest = d / ('mapping' + Path(f.name).suffix)
    with open(dest, 'wb') as out:
        for chunk in f.chunks():
            out.write(chunk)
    (d / 'meta.json').write_text(json.dumps(
        {'name': f.name, 'path': dest.name}), encoding='utf-8')
    return redirect('b2b_ship_to_preview', token=token)


@login_required
def ship_to_preview(request, token):
    """Parse the uploaded Excel in-memory (no write) and show counts to confirm."""
    d = _stm_token_dir(token)
    mp = d / 'meta.json'
    if not mp.exists():
        raise Http404("Upload not found or expired.")
    meta = json.loads(mp.read_text(encoding='utf-8'))
    from .services import mapping_store as ms
    rows, stats, warnings = ms.build_rows(str(d / meta['path']))
    if not rows:
        messages.error(request, '; '.join(warnings) or "No rows parsed.")
        return redirect('b2b_ship_to')
    return render(request, 'online_b2b/ship_to_preview.html', {
        'token': token, 'meta': meta, 'stats': stats, 'warnings': warnings,
        'sample': rows[:15], 'current': ms.status(),
    })


@login_required
@require_POST
def ship_to_confirm(request, token):
    """Replace the Excel-sourced rows from the uploaded file (manual rows kept)."""
    d = _stm_token_dir(token)
    mp = d / 'meta.json'
    if not mp.exists():
        raise Http404("Upload not found or expired.")
    meta = json.loads(mp.read_text(encoding='utf-8'))
    from .services import mapping_store as ms
    try:
        rows, stats, _ = ms.build_rows(str(d / meta['path']))
        res = ms.replace_mapping(rows)
    except Exception as e:  # noqa: BLE001
        messages.error(request, f"Replace failed: {type(e).__name__}: {e}")
        return redirect('b2b_ship_to_preview', token=token)
    shutil.rmtree(d, ignore_errors=True)
    messages.success(request, f"✓ Ship-To mapping replaced — {res['rows']} rows "
                     f"across {stats['parties']} parties now live.")
    return redirect('b2b_ship_to')


@login_required
@require_POST
def ship_to_discard(request, token):
    d = _stm_token_dir(token)
    if d.exists():
        shutil.rmtree(d, ignore_errors=True)
    messages.info(request, "Upload discarded.")
    return redirect('b2b_ship_to')


@login_required
@require_POST
def ship_to_seed(request):
    """Re-seed the table from the bundled Ship to B2B.xlsx (one-click)."""
    from .services import mapping_store as ms
    res = ms.seed_from_bundled()
    if res.get('ok'):
        messages.success(request, f"✓ Seeded {res['rows']} rows from the bundled "
                         f"Ship-To B2B Excel.")
    else:
        messages.error(request, res.get('error', 'Seed failed.'))
    return redirect('b2b_ship_to')


@login_required
def ship_to_search(request):
    """As-you-type filter (party + text) → the table partial."""
    from .services import mapping_store as ms
    data = ms.list_mappings(request.GET.get('party', ''),
                            request.GET.get('q', ''), limit=300)
    return render(request, 'online_b2b/_ship_to_rows.html', {'overview': data})


@login_required
@require_POST
def ship_to_add(request):
    """Add ONE mapping row (durable manual)."""
    from .services import mapping_store as ms
    res = ms.add_mapping({k: request.POST.get(k, '') for k in (
        'party', 'del_location', 'cust_no', 'ship_to', 'name', 'address',
        'address2', 'postcode', 'city')})
    return JsonResponse(res, status=200 if res.get('ok') else 400)


@login_required
@require_POST
def ship_to_edit(request, row_id):
    from .services import mapping_store as ms
    res = ms.update_mapping(row_id, {k: request.POST.get(k, '') for k in (
        'party', 'del_location', 'cust_no', 'ship_to', 'name', 'address',
        'address2', 'postcode', 'city')})
    return JsonResponse(res, status=200 if res.get('ok') else 400)


@login_required
@require_POST
def ship_to_delete(request, row_id):
    from .services import mapping_store as ms
    return JsonResponse(ms.delete_mapping(row_id))
