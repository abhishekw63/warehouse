"""
online_po_processor
===================

Marketplace PO/punch file → ERP-importable Sales Order generator.

This package replaces the single-file ``standalone_po_processing.py`` script
(now retained as ``legacy_standalone.py`` for fallback). Logic is identical
through v1.4.0 — only the file layout changed.

Layout overview
---------------
::

    online_po_processor/
        config/      → constants, marketplace registry, paths/history helpers
        data/        → pure-data classes (SORow, ProcessingResult) and loaders
        engine/      → MarketplaceEngine — turns a punch file into result rows
        exporter/    → SOExporter + D365Exporter + per-sheet writers
        emailer/     → HTML report builder + SMTP sender
        gui/         → Tkinter UI (OnlinePOApp + dialogs)
        utils/       → cross-platform helpers (open_file)
        app.py       → bootstrap: expiry check + main() entry point

Quick start
-----------
The intended entry point is the top-level ``main.py`` in the project root::

    python main.py

That file does nothing more than::

    from online_po_processor.app import main
    main()

Public re-exports
-----------------
The most commonly imported names are exposed at package level for
convenience and to mirror what the legacy single-file module exported:

    >>> from online_po_processor import (
    ...     OnlinePOApp,
    ...     MARKETPLACE_CONFIGS,
    ...     SORow,
    ...     ProcessingResult,
    ... )
"""

__version__ = "1.5.1"
__all__ = [
    "__version__",
    # Re-exports for code that used to ``import standalone_po_processing as opp``
    "OnlinePOApp",
    "MARKETPLACE_CONFIGS",
    "MARKETPLACE_NAMES",
    "SORow",
    "ProcessingResult",
    "MasterLoader",
    "MappingLoader",
    "MarketplaceEngine",
    "SOExporter",
    "D365Exporter",
    "EmailBuilder",
    "EmailSender",
    "get_email_config",
    "main",
]

def __getattr__(name):
    """Lazy data-package re-exports.

    Importing submodules such as ``online_po_processor.data.models`` should stay
    usable on headless web hosts that do not provide Tkinter.
    """
    if name == "get_email_config":
        from online_po_processor.config.email_config import get_email_config
        return get_email_config
    if name in {"MARKETPLACE_CONFIGS", "MARKETPLACE_NAMES"}:
        from online_po_processor.config import marketplaces
        return getattr(marketplaces, name)
    if name == "MappingLoader":
        from online_po_processor.data.mapping_loader import MappingLoader
        return MappingLoader
    if name == "MasterLoader":
        from online_po_processor.data.master_loader import MasterLoader
        return MasterLoader
    if name in {"ProcessingResult", "SORow"}:
        from online_po_processor.data import models
        return getattr(models, name)
    if name in {"EmailBuilder", "EmailSender"}:
        from online_po_processor import emailer
        return getattr(emailer, name)
    if name == "MarketplaceEngine":
        from online_po_processor.engine.marketplace_engine import MarketplaceEngine
        return MarketplaceEngine
    if name == "D365Exporter":
        from online_po_processor.exporter.d365_exporter import D365Exporter
        return D365Exporter
    if name == "SOExporter":
        from online_po_processor.exporter.so_exporter import SOExporter
        return SOExporter
    if name == "OnlinePOApp":
        from online_po_processor.gui.app_window import OnlinePOApp
        return OnlinePOApp
    if name == "main":
        from online_po_processor.app import main
        return main
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
