"""
core.code_audit
===============

On-demand "all-angles" code audit for the Dev · Health page — runs ruff
(standards), AST metrics (big files / big functions), a light duplication scan,
and TODO/FIXME counts over the app code (NOT the frozen engine or .venv). Pure
analysis: it only reads files and writes a cached ``logs/audit.json``. Never
imports or mutates app modules, so it can't disturb the backend.
"""

from __future__ import annotations

import ast
import hashlib
import json
import subprocess
from datetime import datetime
from pathlib import Path

from django.conf import settings

BASE = Path(settings.BASE_DIR)
AUDIT_FILE = BASE / "logs" / "audit.json"
# App code we own — exclude the frozen engine, venv, migrations, vendored assets.
DIRS = ["core", "online_b2b", "offline", "renee_cosmetics"]
EXCLUDE = (
    "migrations",
    "__pycache__",
    ".venv",
    "online_po_management",
    "standalone_files",
    "node_modules",
    "staticfiles",
)


def _rel(p: Path) -> str:
    try:
        return str(p.relative_to(BASE)).replace("\\", "/")
    except ValueError:
        return str(p)


def _py_files():
    for d in DIRS:
        root = BASE / d
        if not root.exists():
            continue
        for p in root.rglob("*.py"):
            if any(x in p.parts for x in EXCLUDE):
                continue
            yield p


def _ruff() -> dict:
    exe = BASE / ".venv" / "Scripts" / "ruff.exe"
    if not exe.exists():
        exe = Path("ruff")  # fall back to PATH
    try:
        r = subprocess.run(
            [str(exe), "check", *DIRS, "--output-format", "json"],
            cwd=str(BASE),
            capture_output=True,
            text=True,
            timeout=180,
        )
        data = json.loads(r.stdout or "[]")
        findings = [
            {
                "file": _rel(Path(d.get("filename", ""))),
                "line": (d.get("location") or {}).get("row"),
                "code": d.get("code") or "",
                "msg": d.get("message") or "",
            }
            for d in data
        ]
        by_rule: dict = {}
        for f in findings:
            by_rule[f["code"]] = by_rule.get(f["code"], 0) + 1
        top = sorted(by_rule.items(), key=lambda x: -x[1])[:12]
        return {
            "available": True,
            "count": len(findings),
            "by_rule": top,
            "findings": findings[:400],
        }
    except Exception as e:  # noqa: BLE001
        return {"available": False, "error": str(e), "count": 0, "by_rule": [], "findings": []}


def _metrics() -> dict:
    big_files, big_funcs = [], []
    todos = total_lines = nfiles = 0
    for p in _py_files():
        try:
            src = p.read_text(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            continue
        nfiles += 1
        lines = src.count("\n") + 1
        total_lines += lines
        rel = _rel(p)
        todos += sum(1 for ln in src.splitlines() if "TODO" in ln or "FIXME" in ln or "XXX" in ln)
        if lines > 600:
            big_files.append({"file": rel, "lines": lines})
        try:
            tree = ast.parse(src)
        except Exception:  # noqa: BLE001
            continue
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                span = (node.end_lineno or node.lineno) - node.lineno
                if span > 80:
                    big_funcs.append(
                        {"file": rel, "func": node.name, "lines": span, "line": node.lineno}
                    )
    big_files.sort(key=lambda x: -x["lines"])
    big_funcs.sort(key=lambda x: -x["lines"])
    return {
        "files": nfiles,
        "total_lines": total_lines,
        "todos": todos,
        "big_files": big_files[:30],
        "big_funcs": big_funcs[:30],
    }


def _duplication() -> list:
    """Light copy-paste signal: hash 6-line windows of meaningful code; report
    windows that appear in 2+ places."""
    seen: dict = {}
    win = 6
    for p in _py_files():
        try:
            lines = p.read_text(encoding="utf-8", errors="replace").splitlines()
        except Exception:  # noqa: BLE001
            continue
        rel = _rel(p)
        norm = [ln.strip() for ln in lines]
        for i in range(max(0, len(norm) - win)):
            block = [x for x in norm[i : i + win] if x and not x.startswith("#")]
            if len(block) < win:
                continue
            h = hashlib.md5("\n".join(block).encode("utf-8")).hexdigest()
            seen.setdefault(h, {"sample": block[0][:90], "locs": []})
            seen[h]["locs"].append(f"{rel}:{i + 1}")
    groups = [
        v
        for v in seen.values()
        if len({loc.rsplit(":", 1)[0] for loc in v["locs"]}) >= 2 or len(v["locs"]) >= 3
    ]
    groups.sort(key=lambda v: -len(v["locs"]))
    return [
        {"count": len(g["locs"]), "sample": g["sample"], "locations": g["locs"][:6]}
        for g in groups[:25]
    ]


import re  # noqa: E402

# High-signal security patterns (clear-cut; low false-positive). (regex, label, sev)
_SEC_PATTERNS = [
    (re.compile(r"\beval\s*\("), "eval() on dynamic input", "high"),
    (re.compile(r"\bexec\s*\("), "exec() of dynamic code", "high"),
    (re.compile(r"shell\s*=\s*True"), "subprocess shell=True", "high"),
    (re.compile(r"\bmark_safe\s*\("), "mark_safe — XSS risk if unescaped", "med"),
    (re.compile(r"\bpickle\.load"), "pickle load — unsafe deserialization", "high"),
    (re.compile(r"verify\s*=\s*False"), "TLS verification disabled", "high"),
    (re.compile(r"yaml\.load\s*\((?!.*Loader)"), "yaml.load without SafeLoader", "high"),
    (re.compile(r"\.execute\(\s*[^)]*%[^)]*%"), "SQL built with %-formatting", "high"),
    (re.compile(r"\.execute\([^)]*\.format\("), "SQL built with .format()", "high"),
    (re.compile(r'ALLOWED_HOSTS\s*=\s*\[\s*["\']\*'), "ALLOWED_HOSTS = ['*']", "med"),
]


_SELF = _rel(Path(__file__))  # don't flag this scanner's own rules


def _security() -> list:
    hits = []
    for p in _py_files():
        rel = _rel(p)
        if rel == _SELF:  # the rule definitions live here
            continue
        try:
            lines = p.read_text(encoding="utf-8", errors="replace").splitlines()
        except Exception:  # noqa: BLE001
            continue
        for i, ln in enumerate(lines, 1):
            stripped = ln.lstrip()
            if stripped.startswith("#") or "re.compile(" in ln:
                continue
            for rx, label, sev in _SEC_PATTERNS:
                if rx.search(ln):
                    hits.append(
                        {
                            "file": rel,
                            "line": i,
                            "sev": sev,
                            "issue": label,
                            "code": ln.strip()[:120],
                        }
                    )
    order = {"high": 0, "med": 1, "low": 2}
    hits.sort(key=lambda h: order.get(h["sev"], 9))
    return hits[:60]


def run_audit() -> dict:
    res = {
        "ts": datetime.now().isoformat(timespec="seconds"),
        "ruff": _ruff(),
        "metrics": _metrics(),
        "duplication": _duplication(),
        "security": _security(),
    }
    try:
        AUDIT_FILE.parent.mkdir(parents=True, exist_ok=True)
        AUDIT_FILE.write_text(json.dumps(res), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    return res


def last_audit() -> dict | None:
    try:
        return json.loads(AUDIT_FILE.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return None
