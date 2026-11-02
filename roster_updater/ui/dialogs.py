"""Capability correction dialog."""
from __future__ import annotations

from tkinter import messagebox, ttk
import tkinter as tk


class CapabilityDialog(tk.Toplevel):
    """Modal editor that collects a capability, source, and optional notes.

    The on_save callback performs persistence; this dialog owns only input UI.
    """
    def __init__(
        self,
        parent: tk.Misc,
        employee_name: str,
        logon_name: str,
        current: str,
        suggested: str,
        on_save,
    ) -> None:
        """Build and center the editor, prefilling a suggestion when one is available.

        on_save receives (capability, source, notes) when Save or Enter is used.
        Prefilling a suggestion does not itself save or approve anything.
        """
        super().__init__(parent)
        self.title("Capability correction")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.on_save = on_save

        outer = ttk.Frame(self, padding=18)
        outer.grid(row=0, column=0, sticky="nsew")

        ttk.Label(outer, text=employee_name, style="DialogTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 2)
        )
        ttk.Label(outer, text=logon_name, style="Muted.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 14)
        )

        ttk.Label(outer, text="Current capability").grid(row=2, column=0, sticky="w", padx=(0, 14), pady=5)
        ttk.Label(outer, text=current or "Blank").grid(row=2, column=1, sticky="w", pady=5)

        ttk.Label(outer, text="Manager suggestion").grid(row=3, column=0, sticky="w", padx=(0, 14), pady=5)
        ttk.Label(outer, text=suggested or "None").grid(row=3, column=1, sticky="w", pady=5)

        ttk.Label(outer, text="Approved capability").grid(row=4, column=0, sticky="w", padx=(0, 14), pady=5)
        self.capability_var = tk.StringVar(value=suggested or current)
        capability_entry = ttk.Entry(outer, width=34, textvariable=self.capability_var)
        capability_entry.grid(row=4, column=1, sticky="ew", pady=5)

        ttk.Label(outer, text="Source").grid(row=5, column=0, sticky="w", padx=(0, 14), pady=5)
        self.source_var = tk.StringVar(value="Manager suggestion" if suggested else "Slack/manual")
        source_box = ttk.Combobox(
            outer,
            width=31,
            textvariable=self.source_var,
            values=["Manager suggestion", "Slack/manual", "Other"],
            state="readonly",
        )
        source_box.grid(row=5, column=1, sticky="ew", pady=5)

        ttk.Label(outer, text="Notes").grid(row=6, column=0, sticky="nw", padx=(0, 14), pady=5)
        self.notes = tk.Text(outer, width=34, height=4, wrap="word")
        self.notes.grid(row=6, column=1, sticky="ew", pady=5)

        buttons = ttk.Frame(outer)
        buttons.grid(row=7, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="Save correction", style="Primary.TButton", command=self.save).pack(side="right")

        capability_entry.focus_set()
        self.bind("<Return>", lambda _event: self.save())
        self.bind("<Escape>", lambda _event: self.destroy())
        self.update_idletasks()
        x = parent.winfo_rootx() + max(20, (parent.winfo_width() - self.winfo_width()) // 2)
        y = parent.winfo_rooty() + max(20, (parent.winfo_height() - self.winfo_height()) // 2)
        self.geometry(f"+{x}+{y}")

    def save(self) -> None:
        """Require a nonblank capability, invoke the save callback, then close the dialog."""
        capability = self.capability_var.get().strip()
        if not capability:
            messagebox.showerror("Capability required", "Enter the approved capability.", parent=self)
            return
        self.on_save(capability, self.source_var.get(), self.notes.get("1.0", "end").strip())
        self.destroy()
