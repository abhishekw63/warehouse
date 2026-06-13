"""
offline_po_processor.app
========================

Application bootstrap — the ``main()`` entry point called by the top-level
``main.py`` launcher. Mirrors ``online_po_processor.app`` so the Online and
Offline sides share the same shape and can be merged under one OMT (Order
Management Tool) shell later.

For now it sets up logging and opens the channel launcher; everything else
is delegated to the per-channel standalone tools (run as subprocesses).
"""

from __future__ import annotations

import logging

from offline_po_processor.gui.launcher_window import LauncherWindow


def _configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s | %(levelname)s | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )


def main() -> None:
    """Entry point: configure logging, open the channel launcher."""
    _configure_logging()
    LauncherWindow().run()


if __name__ == '__main__':
    # Allow `python -m offline_po_processor.app` during development.
    main()
