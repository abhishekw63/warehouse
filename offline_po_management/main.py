"""
Offline PO Management — top-level launcher.

Mirrors ``online_po_management/main.py``: a thin entry point that defers to
the package's ``app.main()``. Run with ``python main.py`` from this folder.
"""

from offline_po_processor.app import main

if __name__ == "__main__":
    main()
