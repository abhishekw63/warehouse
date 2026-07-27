"""
online_b2b.views_triangular  —  STANDALONE views for 3-way Triangular Validation.

Kept in its OWN module (not views.py) so the whole feature — this file +
services/triangular_validation.py + templates/online_b2b/triangular.html + the 3
URL lines + one nav link — can be deleted in one go without touching the rest of
the app. Own media dir (``b2b_triangular``). Read-only; wraps the untouched
``full_validation`` via ``triangular_validation``.
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

_UP = Path(settings.MEDIA_ROOT) / 'b2b_triangular'


def _tok_dir(token: str) -> Path:
    base = _UP.resolve()
    d = (_UP / token).resolve()
    if d != base and base not in d.parents:
        raise Http404()
    return d


@login_required
def triangular(request):
    """Upload form + last result (``?token=``)."""
    token = request.GET.get('token', '')
    result = None
    if token:
        rp = _tok_dir(token) / 'result.json'
        if rp.exists():
            result = json.loads(rp.read_text(encoding='utf-8'))
    ctx = {'token': token, 'result': result}
    if result and result.get('ok'):
        d = result['data']
        ctx['data'] = d
        # on-screen: flagged triangle rows first (all rows live in the Excel)
        ctx['tri_bad'] = [t for t in d.get('triangle', []) if not t.get('agree')]
    return render(request, 'online_b2b/triangular.html', ctx)


@login_required
@require_POST
def triangular_run(request):
    """Accept ONE batch of files (drop everything at once), auto-bifurcate them by
    content into D365 Headers / Lines / dump(s), reconcile (3-way), and stash the
    result + the file breakdown + Excel → redirect with a token. Also accepts the
    classic 3 separate inputs as a fallback."""
    from .services import triangular_validation as tv

    token = uuid.uuid4().hex[:12]
    d = _UP / token
    fdir = d / 'files'
    fdir.mkdir(parents=True, exist_ok=True)

    # Accept a batch of files AND/OR a whole folder (webkitdirectory), plus the
    # classic named inputs as a fallback. Each file is saved under its own index
    # sub-dir so identical basenames from different sub-folders never collide.
    saved = []
    batch = list(request.FILES.getlist('all_files')) + list(request.FILES.getlist('all_folder'))
    for key in ('headers_file', 'lines_file'):
        f = request.FILES.get(key)
        if f:
            batch.append(f)
    batch += request.FILES.getlist('dump_files')
    if not batch:
        messages.error(request, 'Pick some files or a whole folder — D365 Sales Orders + Sales Lines + the dump(s).')
        return redirect('b2b_triangular')
    for i, f in enumerate(batch):
        sub = fdir / str(i)
        sub.mkdir(parents=True, exist_ok=True)
        p = sub / Path(f.name).name            # basename only (drop any sub-path)
        with open(p, 'wb') as out:
            for chunk in f.chunks():
                out.write(chunk)
        saved.append(str(p))

    # auto-bifurcate: which file is headers / lines / dump / source / unknown
    cls = tv.classify_files(saved)
    if not cls['headers'] or not cls['lines']:
        detected = '; '.join(f"{s['name']} → {s['role']}" for s in cls['summary'])
        missing = []
        if not cls['headers']:
            missing.append('D365 Sales Orders (headers)')
        if not cls['lines']:
            missing.append('D365 Sales Lines')
        messages.error(request, f"Couldn't find: {', '.join(missing)}. Detected — {detected}")
        return redirect('b2b_triangular')

    res = tv.validate(cls['headers'], cls['lines'], cls['dumps'])
    if res.get('ok'):
        res['data']['files_detected'] = [f for f in cls['summary'] if f['role'] != 'unknown']
        res['data']['n_ignored'] = len(cls['unknown']) + len(cls['source'])
        res['data']['superseded'] = cls.get('superseded', [])
        try:
            tv.build_workbook(res['data'], str(d / 'triangular.xlsx'))
        except Exception as e:  # noqa: BLE001 — Excel is a convenience, never blocks the page
            res['data']['excel_error'] = f'{type(e).__name__}: {e}'
    (d / 'result.json').write_text(json.dumps(res, default=str), encoding='utf-8')
    if not res.get('ok'):
        messages.error(request, res.get('error', 'Validation failed.'))
        return redirect('b2b_triangular')
    return redirect(f"{reverse('b2b_triangular')}?token={token}")


@login_required
def triangular_download(request, token):
    """Serve the 360° triangular Excel."""
    xp = _tok_dir(token) / 'triangular.xlsx'
    if not xp.exists():
        raise Http404('Triangular file not found or expired.')
    return FileResponse(open(xp, 'rb'), as_attachment=True,
                        filename='D365_Triangular_Validation.xlsx')
