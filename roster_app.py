from __future__ import annotations

import json
import os
import sys
import webbrowser
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from roster_engine import (
    CAPABILITY_HEADER,
    KEY_HEADER,
    ComparisonResult,
    ReviewItem,
    RosterError,
    SharedRosterStore,
    clean,
    normalize_key,
    write_report,
)


APP_NAME = "TAIT Roster Updater"


def user_config_path() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    else:
        base = Path.home() / ".config"
    return base / "TAITRosterUpdater" / "config.json"


class ConfigStore:
    def __init__(self) -> None:
        self.path = user_config_path()

    def load(self) -> dict[str, str]:
        if not self.path.exists():
            return {}
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}
        return data if isinstance(data, dict) else {}

    def save(self, data: dict[str, str]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp = self.path.with_suffix(".tmp")
        temp.write_text(json.dumps(data, indent=2), encoding="utf-8")
        os.replace(temp, self.path)


class CapabilityDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        employee_name: str,
        logon_name: str,
        current: str,
        suggested: str,
        on_save,
    ) -> None:
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
        capability = self.capability_var.get().strip()
        if not capability:
            messagebox.showerror("Capability required", "Enter the approved capability.", parent=self)
            return
        self.on_save(capability, self.source_var.get(), self.notes.get("1.0", "end").strip())
        self.destroy()


class RosterUpdaterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1280x820")
        self.minsize(1050, 680)

        self.config_store = ConfigStore()
        self.config_data = self.config_store.load()
        self.store: SharedRosterStore | None = None
        self.result: ComparisonResult | None = None
        self.review_by_iid: dict[str, ReviewItem] = {}

        self.shared_folder_var = tk.StringVar(value=self.config_data.get("shared_folder", ""))
        self.pull_path_var = tk.StringVar()
        self.baseline_status_var = tk.StringVar(value="Choose the shared roster folder.")
        self.status_var = tk.StringVar(value="Ready")
        self.summary_vars = {
            "New pull employees": tk.StringVar(value="—"),
            "New employees": tk.StringVar(value="—"),
            "Missing from new pull": tk.StringVar(value="—"),
            "Employees with changes": tk.StringVar(value="—"),
            "Capability suggestions": tk.StringVar(value="—"),
            "Manager issues": tk.StringVar(value="—"),
        }

        self._configure_styles()
        self._build_ui()

        if self.shared_folder_var.get():
            self._set_shared_folder(Path(self.shared_folder_var.get()), save=False)

    def _configure_styles(self) -> None:
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

        update = ttk.LabelFrame(main, text="Monthly update", padding=12)
        update.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        update.columnconfigure(1, weight=1)
        ttk.Label(update, text="Sven's new file").grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Entry(update, textvariable=self.pull_path_var, state="readonly").grid(row=0, column=1, sticky="ew")
        ttk.Button(update, text="Choose file", command=self.choose_monthly_pull).grid(row=0, column=2, padx=(10, 0))
        self.run_button = ttk.Button(update, text="Run update", style="Primary.TButton", command=self.run_update)
        self.run_button.grid(row=0, column=3, padx=(8, 0))

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
        self.review_tree.tag_configure("capability", background="#FFF2CC")
        self.review_tree.tag_configure("manager", background="#FCE4D6")
        self.review_tree.tag_configure("missing", background="#F4CCCC")
        self.review_tree.tag_configure("new", background="#E2F0D9")
        self.review_tree.tag_configure("changed", background="#D9EAF7")
        self.review_tree.bind("<Double-1>", lambda _event: self.edit_capability())

        vertical = ttk.Scrollbar(review_frame, orient="vertical", command=self.review_tree.yview)
        horizontal = ttk.Scrollbar(review_frame, orient="horizontal", command=self.review_tree.xview)
        self.review_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.review_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")

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
        self.run_button.configure(state="disabled" if busy else "normal")
        self.export_button.configure(state="disabled" if busy else "normal")
        self.approve_button.configure(state="disabled" if busy else "normal")
        self.configure(cursor="wait" if busy else "")
        if message:
            self.status_var.set(message)
        self.update_idletasks()

    def choose_shared_folder(self) -> None:
        selected = filedialog.askdirectory(title="Choose the shared roster folder")
        if selected:
            self._set_shared_folder(Path(selected), save=True)

    def _set_shared_folder(self, folder: Path, save: bool) -> None:
        self.store = SharedRosterStore(folder)
        self.store.prepare()
        self.shared_folder_var.set(str(folder))
        if save:
            self.config_data["shared_folder"] = str(folder)
            self.config_store.save(self.config_data)
        self.refresh_baseline_status()

    def refresh_baseline_status(self) -> None:
        if not self.store:
            self.baseline_status_var.set("Choose the shared roster folder.")
            return
        if not self.store.has_baseline():
            self.baseline_status_var.set("No baseline found. Initialize it once using the approved roster workbook.")
            return
        try:
            baseline = self.store.load_baseline()
            modified = self.store.baseline_path.stat().st_mtime
            date_text = __import__("datetime").datetime.fromtimestamp(modified).strftime("%B %d, %Y at %I:%M %p")
            self.baseline_status_var.set(f"Baseline ready: {len(baseline.rows):,} employees · updated {date_text}")
        except RosterError as exc:
            self.baseline_status_var.set(str(exc))

    def initialize_baseline(self) -> None:
        if not self.store:
            messagebox.showinfo("Choose a folder", "Choose the shared roster folder first.")
            return
        source = filedialog.askopenfilename(
            title="Choose the approved baseline workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("CSV files", "*.csv")],
        )
        if not source:
            return
        if self.store.has_baseline() and not messagebox.askyesno(
            "Replace the baseline?",
            "A shared baseline already exists. The app will back it up before replacing it. Continue?",
        ):
            return
        try:
            self._set_busy(True, "Initializing the shared baseline...")
            baseline = self.store.initialize(source)
            self.status_var.set(f"Baseline initialized with {len(baseline.rows):,} employees.")
            self.refresh_baseline_status()
        except Exception as exc:
            self._show_error(exc)
        finally:
            self._set_busy(False)

    def choose_monthly_pull(self) -> None:
        selected = filedialog.askopenfilename(
            title="Choose Sven's new employee pull",
            filetypes=[("Employee files", "*.xlsx *.xlsm *.csv"), ("Excel files", "*.xlsx *.xlsm"), ("CSV files", "*.csv")],
        )
        if selected:
            self.pull_path_var.set(selected)
            self.status_var.set("New pull selected. Click Run update.")

    def run_update(self) -> None:
        if not self.store:
            messagebox.showinfo("Choose a folder", "Choose the shared roster folder first.")
            return
        if not self.store.has_baseline():
            messagebox.showinfo("Initialize baseline", "Initialize the shared baseline before running an update.")
            return
        source = self.pull_path_var.get().strip()
        if not source:
            self.choose_monthly_pull()
            source = self.pull_path_var.get().strip()
            if not source:
                return
        try:
            self._set_busy(True, "Comparing the new pull with the shared baseline...")
            self.result = self.store.compare(source)
            self._display_result()
            self.status_var.set(
                f"Comparison complete. {self.result.summary['Review items']:,} items are available for review."
            )
        except Exception as exc:
            self.result = None
            self._show_error(exc)
        finally:
            self._set_busy(False)

    def _display_result(self) -> None:
        if not self.result:
            return
        for name, variable in self.summary_vars.items():
            variable.set(f"{self.result.summary.get(name, 0):,}")

        for item_id in self.review_tree.get_children():
            self.review_tree.delete(item_id)
        self.review_by_iid.clear()

        for index, item in enumerate(self.result.review_items):
            tag = "changed"
            if "Missing from" in item.issue:
                tag = "missing"
            elif "Capability" in item.issue:
                tag = "capability"
            elif "Manager" in item.issue:
                tag = "manager"
            elif "New employee" in item.issue:
                tag = "new"
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
        selected = self.review_tree.selection()
        if not selected:
            messagebox.showinfo("Select an employee", "Select an employee in the review list first.")
            return None
        return self.review_by_iid.get(selected[0])

    def _current_employee_capability(self, logon_name: str) -> str:
        if not self.result:
            return ""
        key = normalize_key(logon_name)
        for employee in self.result.updated_roster:
            if normalize_key(employee.get(KEY_HEADER)) == key:
                return clean(employee.get(CAPABILITY_HEADER))
        return ""

    def edit_capability(self) -> None:
        if not self.store or not self.result:
            messagebox.showinfo("Run an update", "Run the monthly update before editing capabilities.")
            return
        item = self._selected_review_item()
        if not item:
            return
        if item.issue == "Missing from new pull":
            messagebox.showinfo("Not in the new pull", "Capability corrections cannot be applied to a departed employee.")
            return

        def save(capability: str, source: str, notes: str) -> None:
            try:
                self.store.save_override(item.logon_name, capability, source, notes)
                self.run_update()
                self.status_var.set(f"Saved {capability} for {item.display_name}.")
            except Exception as exc:
                self._show_error(exc)

        CapabilityDialog(
            self,
            item.display_name,
            item.logon_name,
            self._current_employee_capability(item.logon_name),
            item.suggested_value,
            save,
        )

    def apply_suggestion(self) -> None:
        if not self.store or not self.result:
            messagebox.showinfo("Run an update", "Run the monthly update first.")
            return
        item = self._selected_review_item()
        if not item:
            return
        if not item.suggested_value:
            messagebox.showinfo("No suggestion", "This review item does not have a capability suggestion.")
            return
        if not messagebox.askyesno(
            "Apply capability suggestion?",
            f"Set {item.display_name}'s capability to {item.suggested_value}?",
        ):
            return
        try:
            self.store.save_override(
                item.logon_name,
                item.suggested_value,
                "Manager suggestion",
                "Approved from the manager hierarchy in the roster updater.",
            )
            self.run_update()
            self.status_var.set(f"Applied the capability suggestion for {item.display_name}.")
        except Exception as exc:
            self._show_error(exc)

    def export_report(self) -> None:
        if not self.result:
            messagebox.showinfo("Run an update", "Run the monthly update before exporting the report.")
            return
        default_name = f"employee_roster_update_{self.result.created_at:%Y-%m-%d}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save the monthly roster report",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        try:
            self._set_busy(True, "Creating the Excel report...")
            write_report(path, self.result)
            self.status_var.set(f"Report saved to {path}")
            messagebox.showinfo("Report created", f"The report was saved to:\n{path}")
        except Exception as exc:
            self._show_error(exc)
        finally:
            self._set_busy(False)

    def approve_month(self) -> None:
        if not self.store or not self.result:
            messagebox.showinfo("Run an update", "Run the monthly update before approving it.")
            return
        warning = "This will save a backup and make the updated roster the new shared baseline."
        if self.result.unresolved_count:
            warning += f"\n\nThere are still {self.result.unresolved_count:,} capability or manager items to review."
        if not messagebox.askyesno("Approve this month?", warning):
            return
        try:
            self._set_busy(True, "Approving the month and creating backups...")
            report_path, backup_path = self.store.approve(self.result)
            self.refresh_baseline_status()
            self.status_var.set("The month was approved and the shared baseline was updated.")
            messagebox.showinfo(
                "Month approved",
                "The shared baseline was updated.\n\n"
                f"Report: {report_path}\n"
                f"Backup: {backup_path}",
            )
            self.result = None
            self.pull_path_var.set("")
        except Exception as exc:
            self._show_error(exc)
        finally:
            self._set_busy(False)

    def open_shared_folder(self) -> None:
        if not self.store:
            messagebox.showinfo("Choose a folder", "Choose the shared roster folder first.")
            return
        path = str(self.store.root)
        try:
            if os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f"open {json.dumps(path)}")
            else:
                webbrowser.open(Path(path).as_uri())
        except Exception as exc:
            self._show_error(exc)

    def _show_error(self, error: Exception) -> None:
        message = str(error)
        if not isinstance(error, RosterError):
            message = f"The app could not complete that step.\n\n{message}"
        self.status_var.set(message)
        messagebox.showerror("Roster update error", message)


def main() -> None:
    app = RosterUpdaterApp()
    app.mainloop()


if __name__ == "__main__":
    main()

