"""
engine.purplle_parser
=====================

Parses the Purplle PO export into an engine-ready DataFrame.

Purplle's "dump" arrives with a ``.XLS`` extension but is actually a
TAB-SEPARATED text file (a SAP/ERP export), so a normal ``read_excel``
fails. This parser reads it as TSV and cleans the two quirks the raw
export carries:

  * **EAN Number** comes zero-padded with a trailing apostrophe, e.g.
    ``000008904473100590'`` — the real EAN is ``8904473100590``. We strip
    the apostrophe and leading zeros so the master lookup (by EAN) resolves.
  * **MRP / Price / Qty** carry trailing spaces (``'220.00 '``); the
    engine's float parsing tolerates those, so they're left as-is.

Columns (already what the ``Purplle`` config references — no renaming):
``PO Document Number, Item No, EAN Number, Sku, Material long text, MRP,
Price, Qty, Plant, Address, Storage location, Purchasing Group, PO Date,
Expiry Date``.

Registered in ``marketplace_engine.PDF_PARSERS`` under 'purplle' and
invoked when the config sets ``file_parser='purplle'`` (routed by config
key, not file extension).
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def _clean_ean(val) -> str:
    """``000008904473100590'`` → ``8904473100590``. Strips whitespace, a
    trailing apostrophe, and leading zeros. Empty string for blank/NaN."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    s = str(val).strip().rstrip("'").strip()
    s = s.lstrip('0')
    return s


def parse_purplle_export(filepath: str | Path) -> pd.DataFrame:
    """Read the tab-separated Purplle export → cleaned DataFrame."""
    filepath = Path(filepath)
    # latin-1: the export is single-byte; utf-8 can choke on stray bytes.
    df = pd.read_csv(filepath, sep='\t', dtype=str, encoding='latin-1')
    # Normalise headers (strip stray whitespace).
    df.columns = [str(c).strip() for c in df.columns]

    if 'EAN Number' not in df.columns:
        raise ValueError(
            f"{filepath.name}: not a Purplle export — 'EAN Number' column "
            f"missing. Found: {list(df.columns)[:10]}")

    # Clean the EAN in place so item_resolution='from_ean' resolves.
    df['EAN Number'] = df['EAN Number'].map(_clean_ean)

    # Drop fully-blank trailing rows (no PO and no EAN).
    po_col = 'PO Document Number'
    if po_col in df.columns:
        keep = ~(df[po_col].isna() | (df[po_col].astype(str).str.strip() == ''))
        df = df[keep | (df['EAN Number'] != '')].copy()

    if df.empty:
        raise ValueError(f"{filepath.name}: no Purplle line items found.")
    return df


def load_purplle_export_as_dataframe(filepath: str | Path) -> pd.DataFrame:
    """One-shot: parse the Purplle TSV → engine-ready DataFrame.

    Registered in ``marketplace_engine.PDF_PARSERS`` under 'purplle'."""
    return parse_purplle_export(filepath)
