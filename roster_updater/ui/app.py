"""Application startup and user action handlers."""
from __future__ import annotations

from pathlib import Path
from tkinter import filedialog, messagebox
import json
import os
import sys
import tkinter as tk
import webbrowser

from ..constants import CAPABILITY_HEADER, KEY_HEADER
from ..models import ComparisonResult, ReviewItem, RosterError
from ..reports import write_report
from ..storage import SharedRosterStore
from ..utils import clean, normalize_key
from .config import ConfigStore
from .dialogs import CapabilityDialog
from .view import APP_NAME, RosterView


class RosterUpdaterApp(RosterView):
    """Coordinate user actions between the desktop widgets and roster services.

    self.store points to the shared data folder; self.result is the latest
    in-memory comparison. Local config remembers the folder between sessions.
    """
    def __init__(self) -> None:
        """Initialize window state, build the view, and restore the last shared folder."""
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1280x820")
        self.minsize(1050, 680)

        # Local preferences remember where shared data lives; employee data
        # is read through SharedRosterStore, not stored in this config.
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

        # Create state variables before widgets bind to those variables.
        self._configure_styles()
        self._build_ui()

        if self.shared_folder_var.get():
            self._set_shared_folder(Path(self.shared_folder_var.get()), save=False)

    def choose_shared_folder(self) -> None:
        """Ask for a shared data directory; cancellation leaves the selection alone."""
        selected = filedialog.askdirectory(title="Choose the shared roster folder")
        if selected:
            self._set_shared_folder(Path(selected), save=True)

    def _set_shared_folder(self, folder: Path, save: bool) -> None:
        """Create a store for folder, optionally remember it, and refresh baseline status."""
        self.store = SharedRosterStore(folder)
        self.store.prepare()
        self.shared_folder_var.set(str(folder))
        if save:
            self.config_data["shared_folder"] = str(folder)
            self.config_store.save(self.config_data)
        self.result = None
        self._clear_result_display()
        self.refresh_baseline_status()

    def refresh_baseline_status(self) -> None:
        """Display baseline availability, employee count, and last modification time."""
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
        """Ask for the first approved roster when the shared folder has no baseline.

        Delegate backup/import to the store and report errors in the window.
        """
        if not self.store:
            messagebox.showinfo("Choose a folder", "Choose the shared roster folder first.")
            return
        if self.store.has_baseline():
            messagebox.showinfo("Baseline already initialized", "For a reviewed Excel report, use Import reviewed roster as baseline. Initialize baseline is only for the first setup of a shared folder.")
            return
        source = filedialog.askopenfilename(
            title="Choose the approved baseline workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("CSV files", "*.csv")],
        )
        if not source:
            return
        if self.store.has_baseline():
            messagebox.showinfo("Baseline already initialized", "Another baseline appeared while selecting the file. Use the reviewed-import workflow for an existing baseline.")
            return
        try:
            self._set_busy(True, "Initializing the shared baseline...")
            baseline = self.store.initialize(source)
            self.result = None
            self._clear_result_display()
            self.status_var.set(f"Baseline initialized with {len(baseline.rows):,} employees.")
            self.refresh_baseline_status()
        except Exception as exc:
            self._show_error(exc)
        finally:
            self._set_busy(False)

    def choose_monthly_pull(self) -> None:
        """Select a supported input file without running a comparison yet."""
        selected = filedialog.askopenfilename(
            title="Choose Sven's new employee pull",
            filetypes=[("Employee files", "*.xlsx *.xlsm *.csv"), ("Excel files", "*.xlsx *.xlsm"), ("CSV files", "*.csv")],
        )
        if selected:
            self.pull_path_var.set(selected)
            self.status_var.set("New pull selected. Click Run update.")

    def run_update(self) -> None:
        """Validate setup, compare the selected pull, and refresh displayed findings.

        Prompt for a file if necessary. On failure, discard self.result and show
        the error; always restore the action buttons after the attempt.
        """
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
            # Comparing produces an in-memory proposal. Approval is a separate action.
            previous = self.result
            fresh = self.store.compare(source)
            # Retain session confirmations only while both input files are identical.
            # Capability edits rerun the comparison and keep these verified links.
            if (previous and previous.confirmed_logon_changes
                    and previous.baseline.file_hash == fresh.baseline.file_hash
                    and previous.monthly_pull.file_hash == fresh.monthly_pull.file_hash):
                fresh = self.store.compare(source, previous.confirmed_logon_changes)
            self.result = fresh
            self._display_result()
            self.status_var.set(
                f"Comparison complete. {self.result.summary['Review items']:,} items are available for review."
            )
        except Exception as exc:
            self.result = None
            self._show_error(exc)
        finally:
            self._set_busy(False)

    def _current_employee_capability(self, logon_name: str) -> str:
        """Find the proposed capability for a logon name, or return blank if unavailable."""
        if not self.result:
            return ""
        key = normalize_key(logon_name)
        for employee in self.result.updated_roster:
            if normalize_key(employee.get(KEY_HEADER)) == key:
                return clean(employee.get(CAPABILITY_HEADER))
        return ""

    def edit_capability(self) -> None:
        """Open the correction dialog for the selected active employee.

        The nested save callback persists the correction and reruns comparison so
        all affected findings and manager suggestions reflect the saved value.
        """
        if not self.store or not self.result:
            messagebox.showinfo("Run an update", "Run the monthly update before editing capabilities.")
            return
        item = self._selected_review_item()
        if not item:
            return
        if "Logon Name Change" in item.issue:
            messagebox.showinfo("Review identity first", "Use Confirm logon change for this finding; edit capabilities on the employee's other review rows.")
            return
        if item.issue == "Missing from new pull":
            messagebox.showinfo("Not in the new pull", "Capability corrections cannot be applied to a departed employee.")
            return

        def save(capability: str, source: str, notes: str) -> None:
            """Save the dialog values, recompute findings, and display any processing error."""
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
        """Confirm and save a selected manager suggestion, then rerun the comparison."""
        if not self.store or not self.result:
            messagebox.showinfo("Run an update", "Run the monthly update first.")
            return
        item = self._selected_review_item()
        if not item:
            return
        if item.issue != "Capability suggestion" or not item.suggested_value:
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

    def confirm_logon_change(self) -> None:
        """Ask a human to verify a unique candidate before linking its old/new records."""
        if not self.store or not self.result:
            messagebox.showinfo("Run an update", "Run the monthly update first.")
            return
        item = self._selected_review_item()
        if not item or item.issue != "Possible Logon Name Change":
            messagebox.showinfo("Select a candidate", "Select a Possible Logon Name Change finding first.")
            return
        candidate = next((candidate for candidate in self.result.logon_candidates
                          if normalize_key(candidate.old_logon) == normalize_key(item.current_value)
                          and normalize_key(candidate.new_logon) == normalize_key(item.logon_name)), None)
        if not candidate or candidate.ambiguous:
            messagebox.showinfo("Ambiguous identity", "Several records match. Verify and correct the source files before linking this employee.")
            return
        if not messagebox.askyesno(
            "Confirm the same employee?",
            f"Employee: {candidate.display_name}\nOld Logon Name: {candidate.old_logon}\n"
            f"New Logon Name: {candidate.new_logon}\n\nMatching details:\n" +
            "\n".join(candidate.evidence_details) + "\n\n" +
            "Confirm only after verifying these are the same person. The draft will use the old ERP Category "
            "and capability history. Saved corrections move when you approve the month in the app. "
            "Until then, Clear logon confirmations discards these links.",
        ):
            return
        try:
            self.result = self.store.confirm_logon_change(self.result, candidate.old_logon, candidate.new_logon)
            self._display_result()
            self.status_var.set(f"Confirmed the Logon Name change for {candidate.display_name}. Review the updated draft before approval.")
        except Exception as exc:
            self._show_error(exc)

    def clear_logon_confirmations(self) -> None:
        """Discard session-only links and return to a fresh, unlinked comparison."""
        if not self.store or not self.result or not self.result.confirmed_logon_changes:
            messagebox.showinfo("No confirmations", "There are no Logon Name confirmations to clear.")
            return
        if not messagebox.askyesno("Clear logon confirmations?", "Discard all identity links for this comparison? Saved data has not been changed by these confirmations."):
            return
        try:
            self.result = self.store.compare(self.result.monthly_pull.source_path)
            self._display_result()
            self.status_var.set("Logon confirmations cleared. New/departed findings restored.")
        except Exception as exc:
            self._show_error(exc)

    def export_report(self) -> None:
        """Ask for an XLSX destination and export the latest result without approving it."""
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

    def import_reviewed_roster(self) -> None:
        """Preview a saved Excel draft, ask for approval, then promote its actual edits.

        Does not require an in-memory comparison: export context preserves the
        baseline/correction snapshot and confirmed links across app restarts.
        """
        if not self.store or not self.store.has_baseline():
            messagebox.showinfo("Initialize baseline", "Choose a shared folder and initialize its baseline before importing a reviewed report.")
            return
        source = filedialog.askopenfilename(
            title="Select the saved reviewed workbook containing Roster Draft",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm")],
        )
        if not source:
            return
        try:
            self._set_busy(True, "Checking the reviewed workbook...")
            plan = self.store.prepare_reviewed_import(source)
            warning = (
                f"Reviewed file: {plan.reviewed.source_path.name}\n"
                f"Employees: {len(plan.reviewed.rows):,}\n"
                f"Fields edited in Excel: {len(plan.edits):,}\n"
                f"Confirmed username links: {len(plan.confirmed_links):,}\n\n"
                "This makes the saved Roster Draft values the official baseline, reconciles capability corrections, "
                "and keeps backups plus the reviewed workbook. Excel Review/Summary tabs remain export-time snapshots.\n\n"
            )
            if plan.unresolved_count:
                warning += f"The original export had {plan.unresolved_count:,} unresolved findings. Verify your review is complete.\n\n"
            warning += "Have you completed the review and approved this workbook as the next baseline?"
            if not messagebox.askyesno("Import approved reviewed roster?", warning):
                return
            report, backup = self.store.import_reviewed_roster(plan)
            self.result = None
            self._clear_result_display()
            self.pull_path_var.set("")
            self.refresh_baseline_status()
            self.status_var.set("Reviewed Excel roster imported as the new baseline. Saved corrections were reconciled.")
            messagebox.showinfo("Reviewed baseline imported", f"The saved Excel edits are now the baseline.\n\nReviewed workbook: {report}\nPrevious baseline backup: {backup}")
        except Exception as exc:
            self._show_error(exc)
        finally:
            self._set_busy(False)

    def approve_month(self) -> None:
        """Confirm approval, warn about unresolved findings, and delegate persistence.

        After success, clear the pending result and selected monthly file. The
        store performs the stale-baseline check before replacing shared data.
        """
        if not self.store or not self.result:
            messagebox.showinfo("Run an update", "Run the monthly update before approving it.")
            return
        warning = "This saves the app's current comparison as the baseline. It does not read Excel edits. If you edited Roster Draft, cancel and use Import reviewed roster as baseline instead."
        if self.result.unresolved_count:
            warning += f"\n\nThere are still {self.result.unresolved_count:,} capability, manager, or logon-name items to review."
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
            self._clear_result_display()
            self.pull_path_var.set("")
        except Exception as exc:
            self._show_error(exc)
        finally:
            self._set_busy(False)

    def open_shared_folder(self) -> None:
        """Open the selected data folder using the platform-specific file-opening route."""
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
        """Show expected roster errors directly and add context to unexpected errors."""
        message = str(error)
        if not isinstance(error, RosterError):
            message = f"The app could not complete that step.\n\n{message}"
        self.status_var.set(message)
        messagebox.showerror("Roster update error", message)


def main() -> None:
    """Construct the desktop application and start Tkinter's event loop."""
    app = RosterUpdaterApp()
    app.mainloop()


