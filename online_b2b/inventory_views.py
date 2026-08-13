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

import datetime as _dt
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

_INV_UPLOADS = Path(settings.MEDIA_ROOT) / 'b2b_inventory'

# Stock deducts as orders ship, so a snapshot goes stale — re-upload Bin Contents
# on this cadence. Cards past this age show a "refresh" nudge + a top banner.
STALE_HOURS = 6


def _snap_age_hours(captured_at):
    """Hours since a snapshot's captured_at (datetime or 'YYYY-MM-DD HH:MM:SS').
    None if unparseable."""
    if not captured_at:
        return None
    cap = captured_at if isinstance(captured_at, _dt.datetime) else None
    if cap is None:
        s = str(captured_at)[:19]
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M'):
            try:
                cap = _dt.datetime.strptime(s, fmt)
                break
            except ValueError:
                continue
    if cap is None:
        return None
    return max(0.0, (_dt.datetime.now() - cap).total_seconds() / 3600.0)


# ── dashboard ───────────────────────────────────────────────────────────────
@login_required
def inventory(request):
    """Inventory · Stock — pure available-stock view: how much of each SKU we hold
    per warehouse (AHD / BLR / North), from the current Bin-Contents snapshots.
    The fill-rate / OOS / tentative-billing analytics now live in the Fulfilment
    Cockpit; this page is stock only."""
    q = (request.GET.get('q') or '').strip()

    snaps = store.current_snapshots()
    # Fetch EVERY current snapshot's per-bin audit in ONE query (not one round-trip
    # per snapshot — each opens a fresh TLS connection to remote TiDB). Cards and
    # new-bin alerts are then both derived from this single fetch.
    audits = store.bin_audit_bulk([s['snapshot_id'] for s in snaps.values()])
    # order snapshots by our warehouse registry, then any extras
    snap_cards = []
    for w in store.WAREHOUSES:
        s = snaps.get(w['code'])
        card = {'wh': w, 'snap': s}
        if s:
            # per-bin audit → split into what we COUNTED vs EXCLUDED vs NEW, so
            # the operator can see exactly which bins the sellable qty came from.
            audit = audits.get(s['snapshot_id'], [])
            card['considered'] = [b for b in audit if b['decision'] == 'include']
            card['excluded'] = [b for b in audit if b['decision'] == 'exclude']
            card['new'] = [b for b in audit if b['decision'] == 'new']
            age = _snap_age_hours(s.get('captured_at'))
            card['age_hours'] = round(age, 1) if age is not None else None
            card['stale'] = age is not None and age >= STALE_HOURS
        snap_cards.append(card)
    # staleness banner: any current snapshot older than the refresh cadence
    stale = [c for c in snap_cards if c.get('stale')]
    max_age = max((c['age_hours'] for c in snap_cards
                   if c.get('age_hours') is not None), default=None)
    extras = [c for k, c in snaps.items() if k not in store.WH_BY_CODE]

    # new-bin alerts across current snapshots — derived from the SAME bulk audit
    # (no extra per-snapshot query).
    alerts = []
    for code, s in snaps.items():
        nb = [b for b in audits.get(s['snapshot_id'], []) if b['decision'] == 'new']
        if nb:
            alerts.append({'warehouse': code, 'name': store.wh_name(code),
                           'bins': nb, 'count': len(nb),
                           'qty': round(sum(b['qty'] for b in nb), 1)})

    # per-item available stock across the warehouses (AHD / BLR / North), one row
    # per SKU with a qty column per warehouse + total. Optional text filter.
    # Render ALL rows; the search box filters them client-side (no page reload),
    # so ``q`` is only the input's initial value (JS applies it on load).
    wh_codes = [w['code'] for w in store.WAREHOUSES]
    items = store.stock_by_item()
    stock_rows = [{
        'item_no': it['item_no'], 'ean': it['ean'],
        'description': it['description'], 'uom': it['uom'],
        'qtys': [round(it['wh'].get(c, 0.0)) for c in wh_codes],
        'total': round(it['total']),
    } for it in items]
    max_total = max((r['total'] for r in stock_rows), default=0) or 1
    total_units = sum(r['total'] for r in stock_rows)

    return render(request, 'online_b2b/inventory.html', {
        'snap_cards': snap_cards, 'extras': extras, 'alerts': alerts,
        'warehouses': store.WAREHOUSES, 'has_any_stock': bool(snaps),
        'stock_rows': stock_rows, 'q': q, 'item_count': len(stock_rows),
        'max_total': max_total, 'total_units': total_units,
        'stale_cards': stale, 'stale_hours': STALE_HOURS, 'max_age': max_age,
    })


# ── upload → preview → confirm ───────────────────────────────────────────────
@login_required
@require_POST
def inventory_upload(request):
    """Stash uploaded Bin Contents Excel file(s) under a token → preview. No DB
    write. Accepts MULTIPLE files in one go (e.g. one per warehouse) — they're
    parsed + merged by warehouse for a single preview/confirm."""
    files = request.FILES.getlist('bin_file')
    if not files:
        messages.error(request, "Choose one or more D365 'Bin Contents' Excel file(s).")
        return redirect('b2b_inventory')
    token = uuid.uuid4().hex[:12]
    d = _INV_UPLOADS / token
    d.mkdir(parents=True, exist_ok=True)
    saved = []
    for i, f in enumerate(files):
        dest = d / (f"bin_{i}" + (Path(f.name).suffix or '.xlsx'))
        with open(dest, 'wb') as out:
            for chunk in f.chunks():
                out.write(chunk)
        saved.append({'name': f.name, 'path': dest.name})
    (d / 'meta.json').write_text(json.dumps({'files': saved}), encoding='utf-8')
    return redirect('b2b_inventory_preview', token=token)


def _staged(token):
    d = _INV_UPLOADS / token
    meta_f = d / 'meta.json'
    if not meta_f.exists():
        return None, None
    meta = json.loads(meta_f.read_text(encoding='utf-8'))
    return d, meta


def _staged_files(meta):
    """Staged file list — supports the new multi-file meta ({'files':[…]}) and the
    legacy single-file meta ({'name','path'})."""
    if meta.get('files'):
        return meta['files']
    if meta.get('path'):
        return [{'name': meta.get('name', ''), 'path': meta['path']}]
    return []


def _merge_wh(a, b):
    """Merge parsed warehouse ``b`` into ``a`` when the SAME WH code appears in
    more than one uploaded file — sum bins/stock/totals, concat lines. Rare (each
    file is usually one warehouse); keeps multi-file uploads correct regardless."""
    for k, v in (b.get('totals') or {}).items():
        if isinstance(v, (int, float)):
            a['totals'][k] = a['totals'].get(k, 0) + v
    for bc, bd in (b.get('bins') or {}).items():
        if bc in a['bins']:
            a['bins'][bc]['lines'] += bd['lines']
            a['bins'][bc]['qty'] += bd['qty']
        else:
            a['bins'][bc] = bd
    for it, sd in (b.get('stock') or {}).items():
        if it in a['stock']:
            a['stock'][it]['qty'] += sd['qty']
        else:
            a['stock'][it] = sd
    a['lines'].extend(b.get('lines') or [])
    a['totals']['item_count'] = len(a['stock'])


def _parse_all(d, meta):
    """Parse EVERY staged file and merge their warehouses by code. Returns
    (warehouses, file_rows, file_names, errors)."""
    merged, file_rows, names, errors = {}, 0, [], []
    for fe in _staged_files(meta):
        p = d / fe['path']
        names.append(fe.get('name', ''))
        if not p.exists():
            errors.append((fe.get('name', ''), 'file missing'))
            continue
        parsed = store.parse_bin_content(p)
        if not parsed['ok']:
            errors.append((fe.get('name', ''), parsed['error']))
            continue
        file_rows += parsed.get('file_rows', 0)
        for code, w in parsed['warehouses'].items():
            if code in merged:
                _merge_wh(merged[code], w)
            else:
                merged[code] = w
    return merged, file_rows, names, errors


@login_required
def inventory_preview(request, token):
    """Parse the staged file(s) and show the per-warehouse classification preview
    (merged across all uploaded files)."""
    d, meta = _staged(token)
    if not meta:
        messages.error(request, "Upload expired — please re-upload.")
        return redirect('b2b_inventory')
    warehouses, file_rows, names, errors = _parse_all(d, meta)
    for nm, err in errors:
        messages.error(request, f"Couldn't read '{nm}': {err}")
    if not warehouses:
        return redirect('b2b_inventory')

    # shape per-WH cards for the template (+ flag unknown WH codes + new bins)
    cards = []
    for code, w in warehouses.items():
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
        'token': token, 'file_name': ', '.join(n for n in names if n),
        'file_names': [n for n in names if n], 'cards': cards,
        'file_rows': file_rows,
    })


@login_required
@require_POST
def inventory_confirm(request, token):
    """Save ALL staged files' warehouses as new current snapshots."""
    d, meta = _staged(token)
    if not meta:
        messages.error(request, "Upload expired — please re-upload.")
        return redirect('b2b_inventory')
    warehouses, _fr, names, errors = _parse_all(d, meta)
    for nm, err in errors:
        messages.error(request, f"Couldn't read '{nm}': {err}")
    src = ', '.join(n for n in names if n)
    user = request.user.get_username()
    saved = []
    for code, w in warehouses.items():
        res = store.save_snapshot(code, w, source_file=src, user=user)
        if res.get('ok'):
            saved.append(f"{store.wh_short(code)} ({res['item_count']} items)")
    # cleanup staged files
    try:
        for p in d.iterdir():
            p.unlink()
        d.rmdir()
    except OSError:
        pass
    if saved:
        messages.success(request, "Stock snapshot saved for: " + ", ".join(saved) + ".")
    else:
        messages.warning(request, "Nothing saved — no warehouse rows found in the file(s).")
    return redirect('b2b_inventory')


@login_required
def inventory_discard(request, token):
    d, _ = _staged(token)
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
        user=request.user.get_username(),
        warehouse=request.POST.get('warehouse', ''))
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
