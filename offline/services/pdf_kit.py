"""
offline.services.pdf_kit
========================

ADDITIVE, standalone PDF-parsing toolkit. It does **NOT** modify or import any
existing channel parser, engine, or backend — nothing here runs unless a caller
opts in. Channel parsers can migrate onto it incrementally.

Two pillars (the two failure modes that let the Naturals "Rose Glow" line slip):

  1. COLUMN-GEOMETRY extraction — read each cell by its x-position (word
     bounding boxes) instead of splitting a text line on whitespace. Immune to
     glued text / extra spaces / wraps, e.g. Naturals' ``ROSE GLOW 6ML33049990``
     where the name was stuck to the HSN with no space.

  2. RECONCILIATION GUARD — compare the parsed line-count / qty / value against
     the PO's own PRINTED totals and NEVER let a short parse pass silently.
     Channel-agnostic: even a perfectly parsed PDF gets a final sanity check.

Depends only on pdfplumber (already a project dependency).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

try:
    import pdfplumber  # already used across the app
except Exception:  # pragma: no cover - import guard only
    pdfplumber = None


# ── 1. Column-geometry extraction ────────────────────────────────────────────
def cluster_rows(words, y_tol: float = 3.0):
    """Group pdfplumber ``extract_words()`` output into visual rows by vertical
    position (``top``). Returns rows sorted top→bottom, each a list of words
    sorted left→right. Whitespace inside a cell is irrelevant — only geometry."""
    rows: list[dict] = []
    for w in sorted(words, key=lambda w: (round(w["top"], 1), w["x0"])):
        for r in rows:
            if abs(r["top"] - w["top"]) <= y_tol:
                r["words"].append(w)
                break
        else:
            rows.append({"top": w["top"], "words": [w]})
    for r in rows:
        r["words"].sort(key=lambda w: w["x0"])
    return rows


def cells_by_bands(row_words, bounds):
    """Bucket one row's words into columns by x-band. ``bounds`` = ascending x
    cut-points; N bounds ⇒ N+1 columns. A word joins the band its **mid-x**
    falls in — so a name glued to the next column's value still lands in the
    correct column as long as the glued chars start within the name's band.
    (The real fix for glued text is per-word banding: each word is placed
    independently, so ``6ML33049990`` splits only if it is two words; when it is
    one glued token we still keep it whole in the name cell and recover the HSN
    from the EAN/columns to its right — never dropping the row.)"""
    cols = [""] * (len(bounds) + 1)
    for w in row_words:
        mid = (w["x0"] + w["x1"]) / 2
        ci = 0
        while ci < len(bounds) and mid >= bounds[ci]:
            ci += 1
        cols[ci] = (cols[ci] + " " + w["text"]).strip()
    return cols


def bands_from_header(page, header_labels, y_tol: float = 3.0):
    """Infer column boundaries from a header row: locate each header label's
    x-start, set boundaries at the midpoints between consecutive labels. Returns
    ``(bounds, header_top)`` or ``(None, None)`` if the header wasn't found."""
    rows = cluster_rows(page.extract_words(), y_tol)
    low = [h.lower() for h in header_labels]
    best, best_hits = None, 0
    for r in rows:
        txt = " ".join(w["text"] for w in r["words"]).lower()
        hits = sum(1 for h in low if h in txt)
        if hits > best_hits:
            best, best_hits = r, hits
    if not best or best_hits < 2:
        return None, None
    xs = []
    for h in header_labels:
        for w in best["words"]:
            if w["text"].lower().startswith(h.lower()[:4]):
                xs.append(w["x0"])
                break
    xs = sorted(set(xs))
    bounds = [(xs[i] + xs[i + 1]) / 2 for i in range(len(xs) - 1)]
    return bounds, best["top"]


def line_item_rows(pdf_path, serial_re: str = r"^\d+$", min_cells: int = 5):
    """Yield candidate line-item rows across all pages using geometry: any
    clustered row whose FIRST cell is a bare serial number (1,2,3,…) AND whose
    SECOND cell looks like an item code (contains a digit) AND which has at least
    ``min_cells`` words. Those extra checks reject page footers like ``1 of 2``
    and tax/summary rows, while staying layout-agnostic — crucially, it does not
    depend on the spacing between the product name and the next column, so a
    glued name/HSN row is still counted. Returns ``[(serial:int, row_words)]``."""
    if pdfplumber is None:
        raise RuntimeError("pdfplumber not available")
    out = []
    pat = re.compile(serial_re)
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for r in cluster_rows(page.extract_words()):
                w = r["words"]
                if len(w) < min_cells:
                    continue                       # too few columns → footer/summary
                if not pat.match(w[0]["text"].strip()):
                    continue                       # first cell isn't a bare serial
                if not any(ch.isdigit() for ch in w[1]["text"]):
                    continue                       # 2nd cell must be an item-code-ish token
                out.append((int(w[0]["text"]), w))
    return out


# ── 2. Reconciliation guard ──────────────────────────────────────────────────
_TOTAL_QTY = re.compile(r"total\s*qty\s*:?\s*([0-9][0-9,]*)", re.I)
_GRAND = re.compile(
    r"(?:grand\s*total|net\s*value|net\s*payable|total\s*value)\s*:?\s*"
    r"([0-9][0-9,]*\.?\d*)", re.I)


def find_printed_totals(full_text: str) -> dict:
    """Best-effort scrape of a PO's own printed totals from its text. Returns
    ``{'qty': int|None, 'value': float|None, 'max_serial': int|None}``.
    ``max_serial`` = highest leading line number seen (a reliable expected
    line-count when the PO numbers its rows)."""
    qty = None
    m = _TOTAL_QTY.search(full_text)
    if m:
        qty = int(m.group(1).replace(",", ""))
    value = None
    vals = _GRAND.findall(full_text)
    if vals:
        value = max(float(v.replace(",", "")) for v in vals)
    serials = [int(x) for x in re.findall(r"(?m)^\s*(\d{1,3})\s+\d", full_text)]
    return {"qty": qty, "value": value, "max_serial": max(serials) if serials else None}


@dataclass
class ReconResult:
    ok: bool
    issues: list[str] = field(default_factory=list)
    parsed: dict = field(default_factory=dict)
    printed: dict = field(default_factory=dict)

    def raise_if_short(self):
        """Fail loud — the whole point. Call this in a parser to guarantee a
        short/over parse can never be emitted silently."""
        if not self.ok:
            raise ValueError("PDF reconciliation FAILED — do not emit this file:\n  "
                             + "\n  ".join(self.issues))


def reconcile(parsed_count: int, parsed_qty: int, parsed_value: float | None,
              printed: dict, value_tol: float = 1.0) -> ReconResult:
    """Compare parsed line-count / qty / value against the PO's printed totals.
    Any mismatch → ``ok=False`` with a human-readable issue list. Never silent."""
    issues = []
    exp_count = printed.get("max_serial")
    if exp_count is not None and parsed_count != exp_count:
        issues.append(f"line-count: parsed {parsed_count} ≠ printed {exp_count} "
                      f"({exp_count - parsed_count:+d})")
    exp_qty = printed.get("qty")
    if exp_qty is not None and parsed_qty != exp_qty:
        issues.append(f"qty: parsed {parsed_qty} ≠ printed {exp_qty} "
                      f"({exp_qty - parsed_qty:+d})")
    exp_val = printed.get("value")
    if exp_val is not None and parsed_value is not None \
            and abs(parsed_value - exp_val) > value_tol:
        issues.append(f"value: parsed {parsed_value:.2f} ≠ printed {exp_val:.2f} "
                      f"({exp_val - parsed_value:+.2f})")
    return ReconResult(ok=not issues, issues=issues,
                       parsed={"count": parsed_count, "qty": parsed_qty, "value": parsed_value},
                       printed=printed)
