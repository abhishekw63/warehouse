"""
offline_po_processor.gui.launcher_window
========================================

The **Offline PO Management launcher** — a small Tkinter window that lists
the registered channels (EKA / GT Mass / MT Select) and opens the one you
pick.

Design
------
Each channel is launched as an **independent subprocess**
(``python <channel script>`` with the channel's folder as the working
directory). That keeps the channels fully decoupled — their core logic is
untouched, each gets its own Tk main loop, and one crashing can't take the
launcher (or the others) down. It also mirrors how the Online tool stays a
self-contained app, so a future OMT merge can surface both sides from one
shell.

The launcher itself stays open so several channels can be opened in a
session.
"""

from __future__ import annotations

import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

from offline_po_processor.config.channels import (
    channel_script, channel_workdir, enabled_channels,
)


class LauncherWindow:
    """The channel-chooser window."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Offline PO Management")
        self.root.geometry("560x460")
        self.root.minsize(520, 400)
        self._build()

    # ── UI ─────────────────────────────────────────────────────────────
    def _build(self) -> None:
        tk.Label(
            self.root, text="Offline PO Management",
            font=("Arial", 16, "bold"),
        ).pack(pady=(18, 2))
        tk.Label(
            self.root, text="Pick a channel to open",
            font=("Arial", 10), fg="#555",
        ).pack(pady=(0, 12))

        body = tk.Frame(self.root)
        body.pack(fill="both", expand=True, padx=18)

        for ch in enabled_channels():
            self._channel_row(body, ch)

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(
            self.root, textvariable=self.status_var, fg="blue", anchor="w",
        ).pack(fill="x", side="bottom", padx=18, pady=8)

    def _channel_row(self, parent, ch) -> None:
        """One framed row per channel: a launch button + its description."""
        row = tk.Frame(parent, relief="groove", bd=1)
        row.pack(fill="x", pady=5)
        tk.Button(
            row, text=f"▶  {ch.name}", width=14,
            font=("Arial", 11, "bold"), bg="#1A237E", fg="white",
            command=lambda c=ch: self._launch(c),
        ).pack(side="left", padx=10, pady=10)
        tk.Label(
            row, text=ch.description, font=("Arial", 9),
            justify="left", anchor="w", wraplength=360,
        ).pack(side="left", padx=6)

    # ── launch ─────────────────────────────────────────────────────────
    def _launch(self, ch) -> None:
        """
        Open the channel as a detached subprocess so it runs with its own
        Tk loop, independent of this launcher. CWD is the channel's folder
        so its bundled-data resolution (script- and CWD-relative) works.
        """
        script = channel_script(ch)
        if not script.exists():
            messagebox.showerror(
                "Channel Missing",
                f"Could not find {ch.name}'s script:\n\n{script}",
            )
            return
        try:
            subprocess.Popen(
                [sys.executable, str(script)],
                cwd=str(channel_workdir(ch)),
            )
            self.status_var.set(f"Launched {ch.name} …")
        except Exception as e:  # noqa: BLE001 — surface, don't crash launcher
            messagebox.showerror(
                "Launch Failed",
                f"Could not launch {ch.name}:\n\n{type(e).__name__}: {e}",
            )
            self.status_var.set(f"Failed to launch {ch.name}")

    def run(self) -> None:
        self.root.mainloop()
