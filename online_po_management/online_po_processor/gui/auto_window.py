"""
gui.auto_window
===============

AUTO mode UI (v2.4.0) — a separate Toplevel window opened from the main
Manual window's "Auto" button. Manual mode is unchanged; this is purely
additive.

The window lets the operator:

* point at the ``Dump/Online`` folder (defaults to the OneDrive location),
* set a dispatch warehouse (AHD/BLR) **per marketplace**, and
* hit one button to process every marketplace folder unattended.

It drives :class:`online_po_processor.auto.auto_runner.AutoRunner` on a
background thread so the UI stays responsive and streams the runner's log
lines live. No per-file dialogs — the whole point of Auto mode.
"""

from __future__ import annotations

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk
from typing import Dict, Optional

from online_po_processor.auto.auto_runner import AutoRunner, summarize
from online_po_processor.config.marketplaces import (
    DEFAULT_WAREHOUSE,
    MARKETPLACE_NAMES,
    WAREHOUSE_DISPLAY_NAMES,
)


# Best-guess default for the Dump/Online root. Used only to pre-fill the
# folder box; the operator can Browse to anything. Falls back to '' when
# the well-known OneDrive path isn't present on this machine.
_DEFAULT_ROOTS = [
    r"D:\OneDrive - RENEE COSMETICS PRIVATE LIMITED\Dump\Online",
]


def _default_root() -> str:
    for p in _DEFAULT_ROOTS:
        if os.path.isdir(p):
            return p
    return ''


class AutoWindow:
    """Modeless Auto-mode window. One instance per click of the Auto button."""

    def __init__(self, parent: tk.Tk, master_path: Optional[str],
                 mapping_path: Optional[str]) -> None:
        self.parent = parent
        self.master_path = master_path
        self.mapping_path = mapping_path
        self._running = False

        self.win = tk.Toplevel(parent)
        self.win.title("Auto Mode — Batch Processing")
        self.win.geometry("760x640")
        self.win.transient(parent)

        self.root_var = tk.StringVar(value=_default_root())
        self.wh_vars: Dict[str, tk.StringVar] = {}
        self.last_consolidated: Optional[str] = None
        self.last_tracker: Optional[str] = None

        self._build()

    # ── UI ─────────────────────────────────────────────────────────────
    def _build(self) -> None:
        pad = {'padx': 8, 'pady': 4}

        head = tk.Label(
            self.win,
            text="Auto Mode — drop each marketplace's dump in its folder, "
                 "set its warehouse, then Run.",
            font=("Arial", 10, "bold"), justify='left', anchor='w',
        )
        head.pack(fill='x', **pad)

        # ── Folder row ──────────────────────────────────────────────────
        frow = tk.Frame(self.win)
        frow.pack(fill='x', **pad)
        tk.Label(frow, text="Dump / Online folder:", width=18, anchor='w'
                 ).pack(side='left')
        tk.Entry(frow, textvariable=self.root_var).pack(
            side='left', fill='x', expand=True, padx=4)
        tk.Button(frow, text="Browse…", command=self._browse).pack(side='left')

        # ── Per-marketplace warehouse grid ──────────────────────────────
        box = tk.LabelFrame(self.win, text="Warehouse per marketplace",
                            font=("Arial", 9, "bold"))
        box.pack(fill='x', **pad)

        # Bulk-set helpers so the operator isn't clicking 12 dropdowns when
        # most go to the same warehouse.
        bulk = tk.Frame(box)
        bulk.grid(row=0, column=0, columnspan=6, sticky='w', padx=6, pady=(4, 6))
        tk.Label(bulk, text="Set all to:").pack(side='left')
        for wh in WAREHOUSE_DISPLAY_NAMES:
            tk.Button(bulk, text=wh, width=5,
                      command=lambda w=wh: self._set_all(w)).pack(
                side='left', padx=2)

        # Two marketplaces per row to keep the window compact.
        per_row = 2
        for i, mp in enumerate(MARKETPLACE_NAMES):
            r = 1 + i // per_row
            c = (i % per_row) * 3
            var = tk.StringVar(value=DEFAULT_WAREHOUSE)
            self.wh_vars[mp] = var
            tk.Label(box, text=mp, width=14, anchor='w').grid(
                row=r, column=c, sticky='w', padx=(8, 2), pady=2)
            ttk.Combobox(box, textvariable=var, values=WAREHOUSE_DISPLAY_NAMES,
                         state='readonly', width=6).grid(
                row=r, column=c + 1, sticky='w', padx=(0, 16), pady=2)

        # ── Run button + status ─────────────────────────────────────────
        crow = tk.Frame(self.win)
        crow.pack(fill='x', **pad)
        self.run_btn = tk.Button(
            crow, text="▶  Run Auto (all folders)", bg="#00C853", fg='white',
            font=("Arial", 10, "bold"), command=self._on_run,
        )
        self.run_btn.pack(side='left')
        self.open_cons_btn = tk.Button(
            crow, text="📘  Open Consolidated", state='disabled',
            command=self._open_cons,
        )
        self.open_cons_btn.pack(side='left', padx=6)
        self.open_trk_btn = tk.Button(
            crow, text="📋  Open Tracker", state='disabled',
            command=self._open_trk,
        )
        self.open_trk_btn.pack(side='left', padx=2)
        tk.Button(
            crow, text="📜  View History", command=self._view_history,
        ).pack(side='left', padx=2)
        self.status_var = tk.StringVar(value="Idle")
        tk.Label(crow, textvariable=self.status_var, fg='blue').pack(
            side='left', padx=10)

        # ── Log ─────────────────────────────────────────────────────────
        logframe = tk.Frame(self.win)
        logframe.pack(fill='both', expand=True, **pad)
        scroll = tk.Scrollbar(logframe)
        scroll.pack(side='right', fill='y')
        self.log_text = tk.Text(logframe, height=16, wrap='word',
                                yscrollcommand=scroll.set,
                                font=("Consolas", 9))
        self.log_text.pack(side='left', fill='both', expand=True)
        scroll.config(command=self.log_text.yview)

    # ── helpers ────────────────────────────────────────────────────────
    def _browse(self) -> None:
        d = filedialog.askdirectory(
            title="Select the Dump/Online folder",
            initialdir=self.root_var.get() or os.path.expanduser("~"),
        )
        if d:
            self.root_var.set(d)

    def _set_all(self, wh: str) -> None:
        for var in self.wh_vars.values():
            var.set(wh)

    def _open_cons(self) -> None:
        if self.last_consolidated and os.path.exists(self.last_consolidated):
            from online_po_processor.utils import open_file
            open_file(self.last_consolidated)

    def _open_trk(self) -> None:
        if self.last_tracker and os.path.exists(self.last_tracker):
            from online_po_processor.utils import open_file
            open_file(self.last_tracker)

    def _view_history(self) -> None:
        """Export the full order history to a readable .xlsx and open it."""
        from online_po_processor.auto.history_db import (
            default_history_db_path, get_history_store, history_db_path,
        )
        from online_po_processor.utils import open_file
        # Prefer the DB for the chosen Dump/Online folder; fall back to the
        # default shared location (where Manual mode writes).
        root = self.root_var.get().strip()
        db_path = (history_db_path(root) if root and os.path.isdir(root)
                   else default_history_db_path())
        if not os.path.exists(db_path):
            self._log("· No history yet — run Auto or Manual at least once.")
            return
        out = os.path.join(os.path.dirname(str(db_path)), 'Order_History.xlsx')
        store = get_history_store(db_path)
        try:
            store.export_to_xlsx(out)
        finally:
            store.close()
        self._log(f"📜 History exported: {out}")
        open_file(out)

    def _log(self, msg: str) -> None:
        self.log_text.insert('end', msg + "\n")
        self.log_text.see('end')

    def _log_threadsafe(self, msg: str) -> None:
        # Called from the worker thread — marshal onto the Tk thread.
        self.win.after(0, lambda: self._log(msg))

    # ── run ────────────────────────────────────────────────────────────
    def _on_run(self) -> None:
        if self._running:
            return
        root = self.root_var.get().strip()
        if not root or not os.path.isdir(root):
            self._log("✗ Pick a valid Dump/Online folder first.")
            return
        if not self.master_path or not self.mapping_path:
            self._log("✗ Master / mapping file not loaded in the main "
                      "window — load them there first, then reopen Auto.")
            return

        wh_map = {mp: var.get() for mp, var in self.wh_vars.items()}

        self._running = True
        self.run_btn.config(state='disabled')
        self.status_var.set("Running…")
        self.log_text.delete('1.0', 'end')

        t = threading.Thread(
            target=self._worker, args=(root, wh_map), daemon=True)
        t.start()

    def _worker(self, root: str, wh_map: Dict[str, str]) -> None:
        try:
            runner = AutoRunner(
                self.master_path, self.mapping_path,
                warehouse=DEFAULT_WAREHOUSE, warehouse_map=wh_map,
                log=self._log_threadsafe,
            )
            runs = runner.run(root)

            # Consolidated workbook (combined Headers/Lines + roll-up
            # Summary/Validation + file-mapping overview). Best-effort:
            # a failure here must not lose the per-marketplace outputs
            # already written, so it's logged and the run still reports ok.
            cons_path = None
            trk_path = None
            if any(r.status == 'ok' for r in runs):
                from pathlib import Path
                from online_po_processor.auto.consolidated_exporter import (
                    export_consolidated, export_tracker_from_db,
                )
                from online_po_processor.auto.history_db import record_history

                # 1) Record to the history DB FIRST — it's the single source
                #    of truth and gives us the run_id the tracker reads back.
                run_id = None
                try:
                    info = record_history(runs, root)
                    run_id = info['run_id']
                    self._log_threadsafe(
                        f"📒 History: {info['new_orders']} new PO-line(s) "
                        + (f"recorded (run #{run_id})" if run_id
                           else "(nothing new to record)")
                        + (f" — {info['skipped']} already-uploaded removed "
                           f"(Skipped POs sheets)" if info['skipped'] else ""))
                except Exception as he:  # noqa: BLE001
                    self._log_threadsafe(
                        f"⚠ history record failed: "
                        f"{type(he).__name__}: {he}")

                # 2) Tracker — built FROM the DB (single source), only the
                #    NEW POs from this run (already-uploaded ones excluded).
                if run_id is not None:
                    try:
                        trk_path = export_tracker_from_db(
                            run_id, Path(root).parent)
                    except Exception as te:  # noqa: BLE001
                        self._log_threadsafe(
                            f"⚠ tracker file failed: "
                            f"{type(te).__name__}: {te}")

                # 3) Consolidated D365 + review workbook.
                try:
                    cons_path = export_consolidated(runs, root)
                except Exception as ce:  # noqa: BLE001
                    self._log_threadsafe(
                        f"⚠ consolidated file failed: "
                        f"{type(ce).__name__}: {ce}")

            self.win.after(0, lambda: self._done(runs, cons_path, trk_path))
        except Exception as e:  # noqa: BLE001 — surface, don't crash the UI thread
            msg = f"{type(e).__name__}: {e}"
            self.win.after(0, lambda: self._fail(msg))

    def _done(self, runs, cons_path=None, trk_path=None) -> None:
        self._log(summarize(runs))
        if cons_path:
            self.last_consolidated = str(cons_path)
            self._log(f"\n📘 Consolidated workbook: {cons_path}")
            self.open_cons_btn.config(state='normal')
        if trk_path:
            self.last_tracker = str(trk_path)
            self._log(f"📋 Tracker (Dump/Tracker/Online): {trk_path}")
            self.open_trk_btn.config(state='normal')
        ok = sum(1 for r in runs if r.status == 'ok')
        err = sum(1 for r in runs if r.status == 'error')
        self.status_var.set(f"Done — {ok} ok, {err} error")
        self._running = False
        self.run_btn.config(state='normal')

    def _fail(self, msg: str) -> None:
        self._log(f"✗ Run failed: {msg}")
        self.status_var.set("Failed")
        self._running = False
        self.run_btn.config(state='normal')
