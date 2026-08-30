"""
core.observability
===================

Request-timing / perf capture for the Dev · Health page. Fully ADDITIVE: a
single middleware that times every request and appends one JSON line to
``logs/perf.jsonl`` — it never touches the business DB and never alters a
response. Query count/time is measured with ``connection.execute_wrapper`` so it
works regardless of ``DEBUG``.
"""
from __future__ import annotations

import atexit
import json
import os
import re
import threading
import time
from contextlib import ExitStack
from datetime import datetime
from pathlib import Path

from django.conf import settings
from django.db import connections

PERF_LOG = Path(settings.BASE_DIR) / 'logs' / 'perf.jsonl'
_MAX_BYTES = 5_000_000          # cap the log; trim to the last _KEEP lines
_KEEP = 3000

# Buffer perf lines + flush in batches — turns a per-request open()+write() (and an
# inline size-trim that could rewrite the whole 5 MB file mid-request) into ONE write
# per _FLUSH_EVERY requests, off the 0.1-CPU hot path. Ephemeral telemetry only: a
# crash can lose up to a buffer's worth; the durable audit_log path is untouched.
_PERF_BUF: list = []
_PERF_LOCK = threading.Lock()
_FLUSH_EVERY = int(os.environ.get('PERF_FLUSH_EVERY', '25'))


def _flush_perf(force: bool = False) -> None:
    with _PERF_LOCK:
        if not _PERF_BUF or (len(_PERF_BUF) < _FLUSH_EVERY and not force):
            return
        lines, _PERF_BUF[:] = _PERF_BUF[:], []     # drain in-place (same list object)
    try:
        PERF_LOG.parent.mkdir(parents=True, exist_ok=True)
        with open(PERF_LOG, 'a', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')
        if PERF_LOG.stat().st_size > _MAX_BYTES:    # trim on flush, not per request
            tail = PERF_LOG.read_text(encoding='utf-8',
                                      errors='replace').splitlines()[-_KEEP:]
            PERF_LOG.write_text('\n'.join(tail) + '\n', encoding='utf-8')
    except Exception:  # noqa: BLE001
        pass


atexit.register(_flush_perf, True)                  # flush the tail on shutdown
_SKIP_PREFIX = ('/static/', '/media/')
# A GET slower than this (ms) is logged to the audit_log so slow PAGE loads are
# pinpointable in production (perf.jsonl is ephemeral on Render). Writes are always
# stamped regardless of speed. Tune/disable via env (0 = never log slow reads).
_SLOW_MS = float(os.environ.get('PERF_SLOW_MS', '1500'))


class PerfMiddleware:
    """Times each request + counts SQL across all DB connections, appends a perf
    record. Wraps nothing in the response path beyond a timer — safe to keep on."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        counter = {'n': 0, 't': 0.0}
        # Reset the app's RAW-pymysql query counter for this request so it can be
        # folded into `q` (Django's execute_wrapper below only sees ORM queries; the
        # real work is order_db's raw queries).
        try:
            from online_b2b.services import order_db as _odb
            _odb._q_reset()
        except Exception:  # noqa: BLE001
            _odb = None

        def _wrap(execute, sql, params, many, context):
            t = time.perf_counter()
            try:
                return execute(sql, params, many)
            finally:
                counter['n'] += 1
                counter['t'] += time.perf_counter() - t

        t0 = time.perf_counter()
        with ExitStack() as stack:
            for conn in connections.all():
                try:
                    stack.enter_context(conn.execute_wrapper(_wrap))
                except Exception:  # noqa: BLE001 — never block the request
                    pass
            response = self.get_response(request)
        dur_ms = (time.perf_counter() - t0) * 1000.0
        db_ms = counter['t'] * 1000.0                    # Django ORM query time
        try:                                             # + raw pymysql queries + time
            if _odb is not None:
                counter['n'] += _odb._q_count()
                db_ms += _odb._q_time_ms()
        except Exception:  # noqa: BLE001
            pass
        try:
            self._record(request, response, dur_ms, counter, db_ms)
        except Exception:  # noqa: BLE001 — observability must never break a page
            pass
        return response

    @staticmethod
    def _record(request, response, dur_ms, counter, db_ms=0.0):
        path = request.path or ''
        if path.startswith(_SKIP_PREFIX) or path == '/favicon.ico':
            return
        if getattr(response, 'streaming', False):
            size = 0
        else:
            size = len(getattr(response, 'content', b'') or b'')
        user = getattr(getattr(request, 'user', None), 'username', '') or ''
        rec = {
            'ts': datetime.now().isoformat(timespec='seconds'),
            'method': request.method, 'path': path,
            'status': getattr(response, 'status_code', 0),
            'ms': round(dur_ms, 1), 'q': counter['n'], 'db_ms': round(db_ms, 1),
            'qms': round(counter['t'] * 1000.0, 1), 'bytes': size, 'user': user,
        }
        # Buffer the telemetry line; the batched writer flushes to disk + trims.
        with _PERF_LOCK:
            _PERF_BUF.append(json.dumps(rec))
            n = len(_PERF_BUF)
        if n >= _FLUSH_EVERY:
            _flush_perf()
        # DURABLE capture (survives redeploys, visible in the Audit Log GUI): stamp
        # the elapsed ms onto the audit row a write already created, and — for a SLOW
        # page load with no audit row — write one so slow pages can be pinpointed.
        try:
            from core import audit
            aid = getattr(request, '_audit_id', None)
            if aid:
                audit.set_timing(aid, dur_ms, counter['n'], db_ms=db_ms)
            elif (request.method == 'GET' and _SLOW_MS and dur_ms >= _SLOW_MS
                  and getattr(getattr(request, 'user', None), 'is_authenticated', False)):
                m = getattr(request, 'resolver_match', None)
                try:
                    host = request.get_host()
                except Exception:  # noqa: BLE001
                    host = ''
                audit.log(user, 'GET', (m.url_name if m else '') or '', path,
                          '', 'slow load', ms=dur_ms, q=counter['n'], host=host,
                          db_ms=db_ms)
        except Exception:  # noqa: BLE001 — observability must never break a page
            pass


# ── read / aggregate (used by the Dev page) ──────────────────────────────
_HEX = re.compile(r'^[0-9a-f]{8,}$', re.I)


def norm_route(path: str) -> str:
    """Collapse ids/tokens so requests group into routes (``/x/<id>/``)."""
    out = []
    for seg in (path or '').split('/'):
        out.append(':id' if (seg.isdigit() or _HEX.match(seg)) else seg)
    return '/'.join(out)


def recent(limit: int = 800) -> list:
    if not PERF_LOG.exists():
        return []
    rows = []
    for ln in PERF_LOG.read_text(encoding='utf-8',
                                 errors='replace').splitlines()[-limit:]:
        try:
            rows.append(json.loads(ln))
        except Exception:  # noqa: BLE001
            pass
    return rows


def _flags(avg, mx, maxq, avg_bytes) -> list:
    f = []
    if mx > 1500 or avg > 500:
        f.append('slow')
    if maxq > 30:
        f.append('N+1?')
    if avg_bytes > 1_000_000:
        f.append('large')
    return f


def aggregate(rows: list) -> list:
    g: dict = {}
    for r in rows:
        key = (r.get('method', 'GET'), norm_route(r.get('path', '')))
        d = g.setdefault(key, {'method': key[0], 'route': key[1], 'ms': [],
                               'q': [], 'bytes': [], 'hits': 0, 'errors': 0})
        d['hits'] += 1
        d['ms'].append(r.get('ms', 0) or 0)
        d['q'].append(r.get('q', 0) or 0)
        d['bytes'].append(r.get('bytes', 0) or 0)
        if (r.get('status', 200) or 0) >= 400:
            d['errors'] += 1
    out = []
    for d in g.values():
        ms = sorted(d['ms'])
        n = len(ms)
        avg = sum(ms) / n if n else 0
        mx = max(ms) if ms else 0
        p95 = ms[int(0.95 * (n - 1))] if n else 0
        maxq = max(d['q']) if d['q'] else 0
        avg_b = (sum(d['bytes']) / n) if n else 0
        out.append({
            'method': d['method'], 'route': d['route'], 'hits': d['hits'],
            'avg_ms': round(avg, 1), 'p95_ms': round(p95, 1),
            'max_ms': round(mx, 1), 'avg_q': round(sum(d['q']) / n, 1) if n else 0,
            'max_q': maxq, 'avg_kb': round(avg_b / 1024, 1), 'errors': d['errors'],
            'flags': _flags(avg, mx, maxq, avg_b),
        })
    out.sort(key=lambda x: -x['max_ms'])
    return out


def kpis(rows: list) -> dict:
    n = len(rows)
    ms = [r.get('ms', 0) or 0 for r in rows]
    return {
        'requests': n,
        'avg_ms': round(sum(ms) / n, 1) if n else 0,
        'max_ms': round(max(ms), 1) if ms else 0,
        'errors': sum(1 for r in rows if (r.get('status', 200) or 0) >= 400),
        'queries': sum(r.get('q', 0) or 0 for r in rows),
    }
