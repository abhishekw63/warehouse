"""
online_b2b.full_validation_views  —  STANDALONE views for Full Validation.

Kept in its OWN module (not views.py) so the whole feature — this file +
services/full_validation.py + templates/online_b2b/full_validation.html + the 3
URL lines + one sidebar link — can be deleted in one go without touching the
rest of the app. Own media dir (``b2b_full_validation``). Read-only.
"""
from __future__ import annotations

import json
import uuid
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.http import FileResponse, Http404
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from .services import common

_UP = Path(settings.MEDIA_ROOT) / 'b2b_full_validation'


def _tok_dir(token: str) -> Path:
    return common.token_dir(_UP, token)


@login_required
def full_validation(request):
    """Upload form + last result (``?token=``)."""
    token = request.GET.get('token', '')
    result = None
    if token:
        rp = _tok_dir(token) / 'result.json'
        if rp.exists():
            result = json.loads(rp.read_text(encoding='utf-8'))
    ctx = {'token': token, 'result': result}
    if result and result.get('ok'):
        # on-screen line list = non-OK rows only (the full 1,900+ live in Excel)
        ctx['line_bad'] = [ln for ln in result.get('lines', []) if ln.get('status') != 'OK']
    return render(request, 'online_b2b/full_validation.html', ctx)


@login_required
@require_POST
def full_validation_run(request):
    """Save both D365 files → reconcile → stash result → redirect with token."""
    hf = request.FILES.get('headers_file')
    lf = request.FILES.get('lines_file')
    if not hf or not lf:
        messages.error(request, 'Upload BOTH files: D365 Sales Orders (headers) + Sales Lines.')
        return redirect('b2b_full_validation')
    token = uuid.uuid4().hex[:12]
    d = _UP / token
    d.mkdir(parents=True, exist_ok=True)
    hp, lp = d / 'headers.xlsx', d / 'lines.xlsx'
    for f, p in ((hf, hp), (lf, lp)):
        with open(p, 'wb') as out:
            for chunk in f.chunks():
                out.write(chunk)
    from .services import full_validation as fv
    res = fv.validate(str(hp), str(lp), excel_out=str(d / 'reconciliation.xlsx'))
    (d / 'result.json').write_text(json.dumps(res, default=str), encoding='utf-8')
    if not res.get('ok'):
        messages.error(request, res.get('error', 'Validation failed.'))
        return redirect('b2b_full_validation')
    return redirect(f"{reverse('b2b_full_validation')}?token={token}")


@login_required
def full_validation_download(request, token):
    """Serve the 3-tier reconciliation Excel."""
    xp = _tok_dir(token) / 'reconciliation.xlsx'
    if not xp.exists():
        raise Http404('Reconciliation file not found or expired.')
    return FileResponse(open(xp, 'rb'), as_attachment=True,
                        filename='D365_Full_Reconciliation.xlsx')
