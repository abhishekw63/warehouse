"""
Off Institutional (INCS) — offline channel.

The FOURTH offline channel, alongside EKA, GT Mass and MT Select. It is a
separate launcher entry / tool, NOT an item inside the MT Select dropdown.

Implementation: it reuses the MT Select engine (shared SO numbering, the
6-sheet D365 output workbook, and the shared history-DB push), but runs the
GUI LOCKED to the ``INST`` channel and presents as "Off Institutional".
Keeping one engine avoids duplicating thousands of lines; the launcher
treats this as its own channel.

What it does (INCS / Indian Naval Canteen Service POs):
  * reads ONLY the 'Shades' sheet of the INCS Excel,
  * drops OOS SKUs (0 demand AND 0 inventory),
  * regular order: qty = Demand units, Unit Price = Basic Price (exact),
  * tester order: qty = ceil(demand / 18), price 0.54, EXCLUDING nails,
  * Ship-to: the counter is picked in the GUI (Trupti / Shilpika / 9 IRSD /
    104 Area / R.K. Beach Road),
  * External Document No.: the Demand No when present, else our SO No,
  * pushes both orders (regular + tester) to the shared history DB via the
    "Push to DB" button (segment Offline, marketplace MT, label
    'Off Institutional').
"""
from __future__ import annotations

import sys
from pathlib import Path

# Reuse the MT Select engine — the INST channel definition lives there.
_MT_DIR = Path(__file__).resolve().parents[1] / 'mt_select'
if str(_MT_DIR) not in sys.path:
    sys.path.insert(0, str(_MT_DIR))

import standalone_mt_select_automation as mt   # noqa: E402


def main() -> None:
    """Open the GUI locked to the Off Institutional (INST) channel."""
    mt.run_gui(only_channel='INST',
               app_title='Off Institutional — INCS PO Processor')


if __name__ == '__main__':
    main()
