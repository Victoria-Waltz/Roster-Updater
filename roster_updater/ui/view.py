"""Window layout, styles, and review list presentation."""
from __future__ import annotations

from tkinter import messagebox, ttk
import os
import tkinter as tk

from ..models import ReviewItem
from ..review_categories import CATEGORY_COLORS, category_for_issue


APP_NAME = "TAIT Roster Updater"


class RosterView(tk.Tk):
    """Presentation methods inherited by RosterUpdaterApp.

    The application initializes state and supplies button callbacks. This
    class builds widgets and renders results; it does not persist roster data.
    """
    def _configure_styles(self) -> None:
        """Register reusable Tkinter styles and choose a platform-appropriate theme."""
        style = ttk.Style(self)
        try:
            style.theme_use("vista" if os.name == "nt" else "clam")
        except tk.TclError:
            pass
        style.configure(".", font=("Arial", 10))
        style.configure("Title.TLabel", font=("Arial", 19, "bold"), foreground="#17365D")
        style.configure("DialogTitle.TLabel", font=("Arial", 13, "bold"), foreground="#17365D")
        style.configure("Section.TLabel", font=("Arial", 11, "bold"), foreground="#17365D")
        style.configure("Muted.TLabel", foreground="#666666")
        style.configure("Primary.TButton", font=("Arial", 10, "bold"), padding=(12, 8))
        style.configure("Secondary.TButton", padding=(10, 7))
        style.configure("SummaryValue.TLabel", font=("Arial", 18, "bold"), foreground="#17365D")
        style.configure("SummaryName.TLabel", font=("Arial", 9), foreground="#555555")
        style.configure("Treeview", rowheight=28, font=("Arial", 9))
        style.configure("Treeview.Heading", font=("Arial", 9, "bold"))

    def _build_ui(self) -> None:
        """Build setup controls, file selection, summary cards, review table, and actions.

        Button commands refer to application methods; they run only when clicked.
        """
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main = ttk.Frame(self, padding=20)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(5, weight=1)

        ttk.Label(main, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            main,
            text="Upload Sven's file, review the exceptions, and generate the monthly roster report.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 16))

        # Shared setup identifies the official baseline used by everyone.
        setup = ttk.LabelFrame(main, text="Shared roster setup", padding=12)
        setup.grid(row=2, column=0, sticky="ew")
        setup.columnconfigure(1, weight=1)
        ttk.Label(setup, text="Shared folder").grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Entry(setup, textvariable=self.shared_folder_var, state="readonly").grid(row=0, column=1, sticky="ew")
        ttk.Button(setup, text="Choose folder", command=self.choose_shared_folder).grid(row=0, column=2, padx=(10, 0))
        ttk.Button(setup, text="Initialize baseline", command=self.initialize_baseline).grid(row=0, column=3, padx=(8, 0))
        ttk.Label(setup, textvariable=self.baseline_status_var, style="Muted.TLabel").grid(
            row=1, column=1, columnspan=3, sticky="w", pady=(7, 0)
        )

        # Monthly input is separate from the approved baseline.
        update = ttk.LabelFrame(main, text="Monthly update", padding=12)
        update.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        update.columnconfigure(1, weight=1)
        ttk.Label(update, text="Sven's new file").grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Entry(update, textvariable=self.pull_path_var, state="readonly").grid(row=0, column=1, sticky="ew")
        ttk.Button(update, text="Choose file", command=self.choose_monthly_pull).grid(row=0, column=2, padx=(10, 0))
        self.run_button = ttk.Button(update, text="Run update", style="Primary.TButton", command=self.run_update)
        self.run_button.grid(row=0, column=3, padx=(8, 0))

        # Each StringVar updates a summary label when comparison finishes.
        summary = ttk.Frame(main)
        summary.grid(row=4, column=0, sticky="ew", pady=(14, 12))
        for column in range(len(self.summary_vars)):
            summary.columnconfigure(column, weight=1)
        for column, (name, variable) in enumerate(self.summary_vars.items()):
            card = ttk.Frame(summary, padding=(10, 6))
            card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 4, 4))
            ttk.Label(card, textvariable=variable, style="SummaryValue.TLabel").pack(anchor="center")
            ttk.Label(card, text=name, style="SummaryName.TLabel").pack(anchor="center")

        review_frame = ttk.LabelFrame(main, text="Review items", padding=10)
        review_frame.grid(row=5, column=0, sticky="nsew")
        review_frame.columnconfigure(0, weight=1)
        review_frame.rowconfigure(0, weight=1)

        # Treeview columns show a subset of ReviewItem; review_by_iid retains
        # the full object for editing and applying suggestions.
        columns = ("employee", "logon", "issue", "current", "suggested", "manager", "changed")
        self.review_tree = ttk.Treeview(review_frame, columns=columns, show="headings", selectmode="browse")
        headings = {
            "employee": "Employee",
            "logon": "Logon Name",
            "issue": "Issue",
            "current": "Current",
            "suggested": "Suggested",
            "manager": "Manager",
            "changed": "Changed Fields",
        }
        widths = {"employee": 180, "logon": 130, "issue": 190, "current": 130, "suggested": 130, "manager": 180, "changed": 220}
        for column in columns:
            self.review_tree.heading(column, text=headings[column])
            self.review_tree.column(column, width=widths[column], minwidth=90, anchor="w")
        for category, color in CATEGORY_COLORS.items():
            self.review_tree.tag_configure(category, background=f"#{color}")
        self.review_tree.bind("<Double-1>", lambda _event: self.edit_capability())

        vertical = ttk.Scrollbar(review_frame, orient="vertical", command=self.review_tree.yview)
        horizontal = ttk.Scrollbar(review_frame, orient="horizontal", command=self.review_tree.xview)
        self.review_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.review_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")

        # The controller supplies these callbacks through inheritance.
        actions = ttk.Frame(main)
        actions.grid(row=6, column=0, sticky="ew", pady=(12, 0))
        ttk.Button(actions, text="Edit capability", command=self.edit_capability).pack(side="left")
        ttk.Button(actions, text="Apply suggestion", command=self.apply_suggestion).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Open shared folder", command=self.open_shared_folder).pack(side="left", padx=(8, 0))
        self.approve_button = ttk.Button(actions, text="Approve as new baseline", command=self.approve_month)
        self.approve_button.pack(side="right")
        self.export_button = ttk.Button(actions, text="Export report", style="Primary.TButton", command=self.export_report)
        self.export_button.pack(side="right", padx=(0, 8))

        status = ttk.Label(main, textvariable=self.status_var, style="Muted.TLabel")
        status.grid(row=7, column=0, sticky="w", pady=(8, 0))

    def _set_busy(self, busy: bool, message: str = "") -> None:
        """Toggle the main action buttons/cursor and optionally update status text.

        update_idletasks() repaints pending UI work; processing still runs on the
        GUI thread, so this is visual feedback rather than background execution.
        """
        self.run_button.configure(state="disabled" if busy else "normal")
        self.export_button.configure(state="disabled" if busy else "normal")
        self.approve_button.configure(state="disabled" if busy else "normal")
        self.configure(cursor="wait" if busy else "")
        if message:
            self.status_var.set(message)
        self.update_idletasks()

    def _display_result(self) -> None:
        """Replace displayed counts and review rows from self.result.

        Map each widget row ID to its ReviewItem so selection can retrieve the
        full finding, including values not shown in the table.
        """
        if not self.result:
            return
        for name, variable in self.summary_vars.items():
            variable.set(f"{self.result.summary.get(name, 0):,}")

        for item_id in self.review_tree.get_children():
            self.review_tree.delete(item_id)
        self.review_by_iid.clear()

        for index, item in enumerate(self.result.review_items):
            tag = category_for_issue(item.issue)
            item_id = f"review-{index}"
            self.review_tree.insert(
                "",
                "end",
                iid=item_id,
                values=(
                    item.display_name,
                    item.logon_name,
                    item.issue,
                    item.current_value,
                    item.suggested_value,
                    item.manager,
                    item.changed_fields,
                ),
                tags=(tag,),
            )
            self.review_by_iid[item_id] = item

    def _selected_review_item(self) -> ReviewItem | None:
        """Return the selected finding, or show a selection prompt and return None."""
        selected = self.review_tree.selection()
        if not selected:
            messagebox.showinfo("Select an employee", "Select an employee in the review list first.")
            return None
        return self.review_by_iid.get(selected[0])

