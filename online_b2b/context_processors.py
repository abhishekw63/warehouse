"""Build-version badge for the sidebar — so you can instantly tell whether a
server restart actually took effect (and whether the running process is stale).

Delivered as a CONTEXT PROCESSOR (not a template tag) on purpose: if this isn't
loaded yet on an old process, ``{{ build_info }}`` simply renders empty — it can
never crash a page. A new template tag would raise on the stale server.
"""

import datetime as _dt
import glob as _glob
import os as _os

from django.conf import settings
from django.utils.safestring import mark_safe

# When this module first loaded = when the running process picked up the Python
# code. Python modules only (re)load on a server (re)start → this is "boot".
_BOOT = _dt.datetime.now()
_WATCH = [
    "online_b2b/services",
    "online_b2b/views.py",
    "online_b2b/urls.py",
    "online_b2b/models.py",
]


def _latest_source_mtime():
    latest = _BOOT
    try:
        base = str(settings.BASE_DIR)
        for w in _WATCH:
            p = _os.path.join(base, w)
            files = (
                _glob.glob(_os.path.join(p, "**", "*.py"), recursive=True)
                if _os.path.isdir(p)
                else [p]
            )
            for f in files:
                m = _dt.datetime.fromtimestamp(_os.path.getmtime(f))
                if m > latest:
                    latest = m
    except Exception:  # noqa: BLE001
        pass
    return latest


def build_info(request):
    """Adds ``build_info`` — a small badge: green '✓ build <time>' when the
    running code is current, RED '⚠ restart needed' when a backend .py changed
    after boot (the process is stale)."""
    boot = _BOOT.strftime("%d %b %H:%M:%S")
    if _latest_source_mtime() > _BOOT:
        badge = (
            '<span class="build-badge stale" title="A backend .py changed after '
            f'this server started ({boot}) — RESTART to load it.">'
            "⚠ code changed · restart needed</span>"
        )
    else:
        badge = (
            f'<span class="build-badge ok" title="Running the latest code.">✓ build {boot}</span>'
        )
    return {"build_info": mark_safe(badge)}
