"""
core.project_map
================

Auto-generated "map of the whole system" for the Project Map page — the file
tree (apps → modules → templates), the real URL→view routes (introspected from
Django's resolver), and the DB models/tables (from the app registry). Pure
read-only introspection; generated from the live codebase so it never goes
stale.
"""
from __future__ import annotations

from pathlib import Path

from django.conf import settings

BASE = Path(settings.BASE_DIR)
APPS = ['core', 'online_b2b', 'offline', 'renee_cosmetics', 'online_po_management']
EXCLUDE = {'__pycache__', '.venv', 'node_modules', 'staticfiles', '.git'}
EXT_KIND = {'.py': 'py', '.html': 'tpl', '.css': 'css', '.js': 'js',
            '.json': 'cfg', '.md': 'doc'}
_MAXDEPTH = 4


# ── file tree ────────────────────────────────────────────────────────────
def _node(path: Path, depth: int = 0) -> dict | None:
    if path.is_dir():
        if path.name in EXCLUDE:
            return None
        children = []
        if depth < _MAXDEPTH:
            for c in sorted(path.iterdir(),
                            key=lambda x: (x.is_file(), x.name.lower())):
                n = _node(c, depth + 1)
                if n:
                    children.append(n)
        return {'name': path.name, 'type': 'dir', 'children': children,
                'count': _count(children)}
    kind = EXT_KIND.get(path.suffix.lower())
    if kind is None:
        return None
    try:
        lines = path.read_text(encoding='utf-8', errors='replace').count('\n') + 1
    except Exception:  # noqa: BLE001
        lines = 0
    return {'name': path.name, 'type': 'file', 'kind': kind, 'lines': lines}


def _count(children: list) -> int:
    return sum(1 if c['type'] == 'file' else c.get('count', 0) for c in children)


def app_tree() -> list:
    out = []
    for a in APPS:
        n = _node(BASE / a) if (BASE / a).exists() else None
        if n:
            n['frozen'] = (a == 'online_po_management')
            out.append(n)
    return out


# ── routes (URL → view) ──────────────────────────────────────────────────
def routes() -> list:
    from django.urls import get_resolver
    from django.urls.resolvers import URLPattern, URLResolver
    out: list = []

    def walk(patterns, prefix=''):
        for p in patterns:
            if isinstance(p, URLResolver):
                walk(p.url_patterns, prefix + str(p.pattern))
            elif isinstance(p, URLPattern):
                cb = p.callback
                if cb is None:
                    continue
                vc = getattr(cb, 'view_class', None)
                view = vc.__name__ if vc else getattr(cb, '__name__', '?')
                mod = getattr(cb, '__module__', '') or ''
                out.append({'url': '/' + prefix + str(p.pattern),
                            'name': p.name or '', 'view': view, 'module': mod,
                            'app': mod.split('.')[0]})

    try:
        walk(get_resolver().url_patterns)
    except Exception:  # noqa: BLE001
        pass
    out.sort(key=lambda r: (r['app'], r['url']))
    return out


# ── models / DB tables ───────────────────────────────────────────────────
def models() -> list:
    from django.apps import apps
    out = []
    try:
        for m in apps.get_models(include_auto_created=True):
            meta = m._meta
            out.append({'app': meta.app_label, 'model': m.__name__,
                        'table': meta.db_table, 'fields': len(meta.get_fields()),
                        'managed': meta.managed})
    except Exception:  # noqa: BLE001
        pass
    out.sort(key=lambda x: (x['app'], x['table']))
    return out


def summary() -> dict:
    tree = app_tree()

    def files_lines(node):
        if node['type'] == 'file':
            return 1, node.get('lines', 0)
        f = ln = 0
        for c in node.get('children', []):
            cf, cl = files_lines(c)
            f += cf
            ln += cl
        return f, ln

    tf = tl = 0
    for n in tree:
        f, ln = files_lines(n)
        tf += f
        tl += ln
    return {'apps': len(tree), 'files': tf, 'lines': tl,
            'routes': len(routes()), 'models': len(models())}
