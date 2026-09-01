"""Build-version badge for the sidebar — so you can instantly tell WHICH build is
live and WHAT shipped in it (the latest update), plus whether a local process is
stale and needs a restart.

Delivered as a CONTEXT PROCESSOR (not a template tag) on purpose: if this isn't
loaded yet on an old process, ``{{ build_info }}`` simply renders empty — it can
never crash a page. A new template tag would raise on the stale server.

The badge is derived from GIT at boot (Render deploys are git checkouts, so
``.git`` + the ``git`` CLI are present): the latest commit's DATE and the current
BRANCH. Nothing to maintain by hand. If git is unavailable for any reason, it
degrades to the old boot-timestamp badge.

Deliberately NOT shown: commit subjects. This badge renders on every page for
every user, and commit messages are internal engineering notes — they routinely
name people, customers and internal decisions that have no business being
published in the product UI. Date + branch answers "which build am I on?"
without leaking the changelog.
"""

import datetime as _dt
import glob as _glob
import os as _os
import subprocess as _sp

from django.conf import settings
from django.utils.safestring import mark_safe
from django.utils.html import escape

# When this module first loaded = when the running process picked up the Python
# code. Python modules only (re)load on a server (re)start → this is "boot".
_BOOT = _dt.datetime.now()
_WATCH = ['online_b2b/services', 'online_b2b/views.py', 'online_b2b/urls.py',
          'online_b2b/models.py']

# Git-derived build facts, computed ONCE (lazily, then cached) — a git call per
# request would be wasteful and the answer can't change without a restart.
_BUILD_CACHE = None     # latest commit date (str) | False when git is unavailable
_BRANCH_CACHE = None    # current branch name (str) | '' when git is unavailable


def _git(args):
    """Run a git command in the project root; '' on any failure (never raises)."""
    try:
        out = _sp.run(['git', *args], cwd=str(settings.BASE_DIR),
                      capture_output=True, text=True, timeout=2)
        if out.returncode == 0:
            return out.stdout.strip()
    except Exception:  # noqa: BLE001 — git missing / not a repo / timeout
        pass
    return ''


def _build_date():
    """Latest commit date from git, cached; False if git is unavailable.

    Only the DATE is read — subjects are never pulled into the process, so there
    is no way for a commit message to reach the page.
    """
    global _BUILD_CACHE
    if _BUILD_CACHE is not None:
        return _BUILD_CACHE
    # %cd uses the commit's OWN timezone (commits are made from IST) → the date
    # reads correctly for the user regardless of the server's UTC clock.
    date = _git(['log', '-1', '--format=%cd', '--date=format:%d %b %Y']).strip()
    _BUILD_CACHE = date or False
    return _BUILD_CACHE


def _branch():
    """Current git branch, cached. '' when git isn't available.

    Shown because this repo runs two long-lived branches (the full app vs the
    limited build) — without it there's no way to tell, from the screen, which
    one you're looking at.
    """
    global _BRANCH_CACHE
    if _BRANCH_CACHE is None:
        _BRANCH_CACHE = _git(['rev-parse', '--abbrev-ref', 'HEAD'])
    return _BRANCH_CACHE


def _latest_source_mtime():
    latest = _BOOT
    try:
        base = str(settings.BASE_DIR)
        for w in _WATCH:
            p = _os.path.join(base, w)
            files = (_glob.glob(_os.path.join(p, '**', '*.py'), recursive=True)
                     if _os.path.isdir(p) else [p])
            for f in files:
                m = _dt.datetime.fromtimestamp(_os.path.getmtime(f))
                if m > latest:
                    latest = m
    except Exception:  # noqa: BLE001
        pass
    return latest


def build_info(request):
    """Adds ``build_info`` — the sidebar build badge.

      • ⚠ RED 'restart needed' when a backend .py changed after boot (local dev,
        stale process — checked first so it always wins).
      • ✓ GREEN 'Build <date>' + the git branch. No commit text (see module doc).
      • Fallback ✓ 'build <boot time>' when git isn't available.
    """
    boot = _BOOT.strftime('%d %b %H:%M:%S')

    # 1) Local staleness always wins — a changed .py means the live code isn't
    #    what's on disk, so the build log would be misleading. This is a DEV-only
    #    aid: on a deployed host (DEBUG off) the process always runs the checked-out
    #    code (a deploy restarts it), so we SKIP the per-request filesystem scan of
    #    every watched .py — it was pure wasted I/O + CPU on every page render.
    if settings.DEBUG and _latest_source_mtime() > _BOOT:
        return {'build_info': mark_safe(
            '<span class="build-badge stale" title="A backend .py changed after '
            f'this server started ({escape(boot)}) — RESTART to load it.">'
            '⚠ code changed · restart needed</span>')}

    date = _build_date()

    # 2) No git → old behaviour (boot timestamp), never crash.
    if not date:
        return {'build_info': mark_safe(
            '<span class="build-badge ok" title="Running the latest code.">'
            f'✓ build {escape(boot)}</span>')}

    branch = _branch()
    tip = f'branch: {branch}\n' if branch else ''
    tip += f'built {date}\nserver booted {boot}'
    chip = f'<span class="bb-b">{escape(branch)}</span>' if branch else ''
    badge = (f'<span class="build-badge ok build-log" title="{escape(tip)}">'
             f'<span class="bb-v">✓ Build {escape(date)}</span>{chip}</span>')
    return {'build_info': mark_safe(badge)}
