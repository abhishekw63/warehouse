"""
online_po_processor.auto
========================

AUTO mode (v2.4.0) — headless batch processing.

This subpackage is a SEPARATE, parallel path to the Tkinter GUI
(``gui/app_window.py``, "Manual mode"). Manual mode is unchanged and
remains the default + fallback. Auto mode walks a
``Dump/Online/<marketplace>/`` tree and processes every file dropped in
each marketplace folder with the SAME engine + exporter the GUI uses —
just without any dialogs or per-file clicking.

Public entry point: :class:`auto_runner.AutoRunner`.
"""

__all__ = ["AutoRunner", "MarketplaceRun"]


def __getattr__(name):
    """Lazy auto re-exports so history_db can load on headless web hosts."""
    if name in {"AutoRunner", "MarketplaceRun"}:
        from online_po_processor.auto.auto_runner import AutoRunner, MarketplaceRun
        return {"AutoRunner": AutoRunner, "MarketplaceRun": MarketplaceRun}[name]
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
