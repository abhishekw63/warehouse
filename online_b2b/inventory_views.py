"""
online_b2b.inventory_views
==========================

Views for the **Inventory — Fill-Rate cockpit** (standalone module, like
``views_triangular`` / ``full_validation_views`` — delete these + the urls +
service + templates + sidebar link to remove the whole feature).

Flow: upload a D365 *Bin Contents* export → preview the per-warehouse
classification (sellable vs virtual bins, + any NEW/unknown bins) → confirm to
store a timestamped snapshot. The dashboard then shows stock-by-warehouse and the
fill-rate / OOS / tentative-billing rollup (PO-wise + MP-wise) against the
recorded orders. All additive; the frozen engine is untouched.
"""
from __future__ import annotations

import json
import uuid
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.http import JsonResponse
from django.shortcuts import redirect, render
from django.views.decorators.http import require_POST

from .services import inventory_store as store
from .services import inventory_fill as fill

_INV_UPLOADS = Path(settings.MEDIA_ROOT) / 'b2b_inventory'


# ── dashboard ───────────────────────────────────────────────────────────────
@login_required
def inventory(request):
    """Inventory cockpit: current stock per warehouse (timestamped) + the
    fill-rate / OOS / tentative-billing rollup for the chosen scope."""
    wh = (request.GET.get('wh') or '').strip()
    mp = (request.GET.get('mp') or '').strip()
    seg = (request.GET.get('seg') or '').strip()
    date_from = (request.GET.get('from') or '').strip()
    date_to = (request.GET.get('to') or '').strip()

    snaps = store.current_snapshots()
    # order snapshots by our warehouse registry, then any extras
    snap_cards = []
    for w in store.WAREHOUSES:
        s = snaps.get(w['code'])
        snap_cards.append({'wh': w, 'snap': s})
    extras = [c for k, c in snaps.items() if k not in store.WH_BY_CODE]

    # new-bin alerts across current snapshots (needs classification)
    alerts = []
    for code, s in snaps.items():
        nb = store.new_bins(s['snapshot_id'])
        if nb:
            alerts.append({'warehouse': code, 'name': store.wh_name(code),
                           'bins': nb, 'count': len(nb),
                           'qty': round(sum(b['qty'] for b in nb), 1)})

    fr = fill.fill_rate(date_from=date_from, date_to=date_to,
                        marketplace=mp, warehouse=wh, segment=seg)

    # AJAX filter change → return ONLY the fill-rate results partial (no reload).
    if request.GET.get('partial'):
        html = render(request, 'online_b2b/_inventory_fill.html', {'fr': fr})
        stock_as_of = fr.get('stock_as_of') or ''
        html['X-Stock-As-Of'] = stock_as_of
        return html

    # marketplaces present in the fill scope (for the filter dropdown)
    mp_options = sorted({g['label'] for g in fr.get('mps', [])})

    return render(request, 'online_b2b/inventory.html', {
        'snap_cards': snap_cards, 'extras': extras, 'alerts': alerts,
        'fr': fr, 'sel_wh': wh, 'sel_mp': mp, 'sel_seg': seg,
        'date_from': fr.get('date_from'), 'date_to': fr.get('date_to'),
        'warehouses': store.WAREHOUSES, 'mp_options': mp_options,
        'has_any_stock': bool(snaps),
    })


# ── upload → preview → confirm ───────────────────────────────────────────────
@login_required
@require_POST
def inventory_upload(request):
    """Stash an uploaded Bin Contents .xlsx under a token → preview. No DB write."""
    f = request.FILES.get('bin_file')
    if not f:
        messages.error(request, "Choose a D365 'Bin Contents' Excel to upload.")
        return redirect('b2b_inventory')
    token = uuid.uuid4().hex[:12]
    d = _INV_UPLOADS / token
    d.mkdir(parents=True, exist_ok=True)
    dest = d / ('bin' + (Path(f.name).suffix or '.xlsx'))
    with open(dest, 'wb') as out:
        for chunk in f.chunks():
            out.write(chunk)
    (d / 'meta.json').write_text(json.dumps({'name': f.name, 'path': dest.name}),
                                 encoding='utf-8')
    return redirect('b2b_inventory_preview', token=token)


def _staged(token):
    d = _INV_UPLOADS / token
    meta_f = d / 'meta.json'
    if not meta_f.exists():
        return None, None, None
    meta = json.loads(meta_f.read_text(encoding='utf-8'))
    return d, meta, d / meta['path']


@login_required
def inventory_preview(request, token):
    """Parse the staged file and show the per-warehouse classification preview."""
    d, meta, path = _staged(token)
    if not path or not path.exists():
        messages.error(request, "Upload expired — please re-upload.")
        return redirect('b2b_inventory')
    parsed = store.parse_bin_content(path)
    if not parsed['ok']:
        messages.error(request, f"Couldn't read the file: {parsed['error']}")
        return redirect('b2b_inventory')

    # shape per-WH cards for the template (+ flag unknown WH codes + new bins)
    cards = []
    for code, w in parsed['warehouses'].items():
        t = w['totals']
        newb = [b for b in w['bins'].values() if b['decision'] == 'new']
        cards.append({
            'code': code, 'name': store.wh_name(code),
            'known_wh': code in store.WH_BY_CODE,
            'totals': t,
            'new_bins': sorted(newb, key=lambda b: -b['qty']),
            'new_count': len(newb),
        })
    cards.sort(key=lambda c: (0 if c['known_wh'] else 1, c['code']))
    return render(request, 'online_b2b/inventory_preview.html', {
        'token': token, 'file_name': meta['name'], 'cards': cards,
        'file_rows': parsed['file_rows'],
    })


@login_required
@require_POST
def inventory_confirm(request, token):
    """Save the staged file's warehouses as new current snapshots."""
    d, meta, path = _staged(token)
    if not path or not path.exists():
        messages.error(request, "Upload expired — please re-upload.")
        return redirect('b2b_inventory')
    parsed = store.parse_bin_content(path)
    if not parsed['ok']:
        messages.error(request, f"Couldn't read the file: {parsed['error']}")
        return redirect('b2b_inventory')
    user = request.user.get_username()
    saved = []
    for code, w in parsed['warehouses'].items():
        res = store.save_snapshot(code, w, source_file=meta['name'], user=user)
        if res.get('ok'):
            saved.append(f"{store.wh_short(code)} ({res['item_count']} items)")
    # cleanup staged file
    try:
        for p in d.iterdir():
            p.unlink()
        d.rmdir()
    except OSError:
        pass
    if saved:
        messages.success(request, "Stock snapshot saved for: " + ", ".join(saved) + ".")
    else:
        messages.warning(request, "Nothing saved — no warehouse rows found in the file.")
    return redirect('b2b_inventory')


@login_required
def inventory_discard(request, token):
    d, _, _ = _staged(token)
    if d and d.exists():
        try:
            for p in d.iterdir():
                p.unlink()
            d.rmdir()
        except OSError:
            pass
    return redirect('b2b_inventory')


# ── bin rules (editable include/exclude list) ───────────────────────────────
@login_required
def inventory_bins(request):
    """Bin classification manager: the editable rule list + the current
    snapshots' per-bin audit (what we counted vs excluded vs flagged new)."""
    wh = (request.GET.get('wh') or '').strip()
    snaps = store.current_snapshots()
    audit = []
    target = None
    if snaps:
        target = wh if (wh and wh in snaps) else next(iter(snaps))
        audit = store.bin_audit(snaps[target]['snapshot_id'])
    summary = {'include': [0, 0.0], 'exclude': [0, 0.0], 'new': [0, 0.0]}
    for b in audit:
        s = summary[b['decision']]
        s[0] += 1
        s[1] += float(b['qty'] or 0)
    return render(request, 'online_b2b/inventory_bins.html', {
        'rules': store.load_rules(), 'audit': audit, 'summary': summary,
        'snaps': snaps, 'target_wh': target,
        'warehouses': store.WAREHOUSES,
    })


@login_required
@require_POST
def inventory_rule_add(request):
    res = store.add_rule(
        request.POST.get('pattern', ''), request.POST.get('match_type', 'prefix'),
        request.POST.get('decision', 'exclude'), request.POST.get('note', ''),
        user=request.user.get_username())
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return JsonResponse(res)
    if res.get('ok'):
        messages.success(request, "Bin rule saved.")
    else:
        messages.error(request, res.get('error', 'Could not save rule.'))
    return redirect('b2b_inventory_bins')


@login_required
@require_POST
def inventory_rule_delete(request, rule_id):
    store.delete_rule(rule_id)
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return JsonResponse({'ok': True})
    messages.success(request, "Bin rule removed.")
    return redirect('b2b_inventory_bins')
