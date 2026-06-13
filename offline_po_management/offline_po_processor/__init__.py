"""
offline_po_processor
====================

The **Offline PO Management** package — the offline / general-trade
counterpart to ``online_po_processor``. Built to the same shape (a
``main.py`` launcher → ``app.main()``) so the two can be merged under one
**OMT (Order Management Tool)** umbrella in future.

Today it is a thin **launcher** over the offline channels (EKA, GT Mass,
MT Select), each of which is a self-contained standalone tool under
``channels/`` whose core logic is unchanged from the original
``standalone_files/`` scripts.

Layout
------
::

    offline_po_management/
        main.py                  → from offline_po_processor.app import main
        offline_po_processor/
            app.py               → main(): bootstrap + open the launcher
            config/channels.py   → the scalable Channel registry
            gui/launcher_window.py
        channels/
            eka/        gt_mass/        mt_select/

Add a channel: drop its tool under ``channels/<key>/`` and add one
``Channel(...)`` entry in ``config/channels.py``.
"""

__version__ = "0.1.0"
__all__ = ["__version__", "main"]

from offline_po_processor.app import main
