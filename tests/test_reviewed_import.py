"""Exercise export -> Excel edit -> explicit reviewed import using fictional employees."""
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import csv
import os
import unittest
from unittest.mock import Mock, patch

from openpyxl import load_workbook

from roster_updater.constants import BASELINE_HEADERS, CORE_HEADERS
from roster_updater.models import RosterError
from roster_updater.reports import write_report
from roster_updater.review_context import CONTEXT_SHEET
from roster_updater.storage import SharedRosterStore
from scripts.make_sample_data import row, write


class ReviewedImportTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.manager = row(1, 'Alex Manager', 'alex', 'alex@example.test', 'Manager',
                           'Operations', 'Enable', 'Alex Manager', 'alex@example.test')
        self.old = row(2, 'Taylor Cooper', 'taylor.old', 'old@example.test', 'Designer',
                       'Design', 'Make', 'Alex Manager', 'alex@example.test', 'Design ERP')
        self.new = dict(self.old, **{'Logon Name': 'taylor.cooper', 'Email Address': 'new@example.test', 'Capability': ''})
        baseline = self.root / 'baseline.xlsx'
        self.monthly = self.root / 'monthly.xlsx'
        write(baseline, 'All Users', BASELINE_HEADERS, [self.manager, self.old])
        write(self.monthly, 'All Users', CORE_HEADERS, [self.manager, self.new])
        self.store = SharedRosterStore(self.root / 'shared')
        self.store.initialize(baseline)
        self.store.save_override('taylor.old', 'Create', 'Slack/manual', 'Confirmed earlier')
        result = self.store.compare(self.monthly)
        self.result = self.store.confirm_logon_change(result, 'taylor.old', 'taylor.cooper')
        self.report = write_report(self.root / 'reviewed.xlsx', self.result)

    def edit(self, changes=None):
        workbook = load_workbook(self.report)
        if changes is None:
            changes = {'G3': 'Senior Designer', 'H3': 'Creative', 'I3': 'Innovate', 'O3': 'Reviewed ERP'}
        for cell, value in changes.items():
            workbook['Roster Draft'][cell] = value
        workbook.save(self.report)
        workbook.close()

    def test_import_after_restart_honors_edits_and_confirmed_links(self):
        self.edit()
        restarted = SharedRosterStore(self.store.root)
        plan = restarted.prepare_reviewed_import(self.report)
        self.assertEqual(len(plan.edits), 4)
        self.assertEqual(plan.confirmed_links, {'taylor.cooper': 'taylor.old'})
        archive, backup = restarted.import_reviewed_roster(plan)
        self.assertTrue(archive.exists())
        self.assertTrue(backup.exists())
        actual = restarted.load_baseline().rows[1]
        self.assertEqual([actual[name] for name in ['Title', 'Department', 'Capability', 'ERP Category']],
                         ['Senior Designer', 'Creative', 'Innovate', 'Reviewed ERP'])
        corrections = restarted.load_overrides()
        self.assertNotIn('taylor.old', corrections)
        self.assertEqual(corrections['taylor.cooper']['capability'], 'Innovate')
        self.assertEqual(corrections['taylor.cooper']['source'], 'Reviewed Excel')
        self.assertEqual(restarted.compare(self.monthly).updated_roster[1]['Capability'], 'Innovate')
        self.assertEqual(archive.read_bytes(), self.report.read_bytes())
        self.assertTrue(list(restarted.history_dir.glob('reviewed_import_*.json')))
        with restarted.audit_path.open(newline='') as handle:
            entry = list(csv.DictReader(handle))[0]
        self.assertEqual(entry['Employee Count'], '2')
        self.assertEqual(entry['Source File'], 'reviewed.xlsx')

    def test_unchanged_capability_keeps_existing_correction_metadata(self):
        expected = self.store.load_overrides()['taylor.old']
        plan = self.store.prepare_reviewed_import(self.report)
        self.assertFalse(plan.edits)
        self.store.import_reviewed_roster(plan)
        self.assertEqual(self.store.load_overrides()['taylor.cooper'], expected)

    def test_blank_capability_removes_old_correction_and_stays_blank_in_baseline(self):
        self.edit({'I3': None})
        self.store.import_reviewed_roster(self.store.prepare_reviewed_import(self.report))
        self.assertEqual(self.store.load_baseline().rows[1]['Capability'], '')
        self.assertNotIn('taylor.old', self.store.load_overrides())
        self.assertNotIn('taylor.cooper', self.store.load_overrides())

    def test_new_key_correction_is_reconciled_with_excel_value(self):
        self.store.save_override('taylor.cooper', 'Lead', 'Other', 'Existing new-key correction')
        self.result = self.store.confirm_logon_change(self.store.compare(self.monthly), 'taylor.old', 'taylor.cooper')
        write_report(self.report, self.result)
        self.edit({'I3': 'Make'})
        self.store.import_reviewed_roster(self.store.prepare_reviewed_import(self.report))
        self.assertEqual(self.store.load_overrides()['taylor.cooper']['capability'], 'Make')

    def test_preview_and_cancel_do_not_change_shared_data(self):
        baseline, corrections = self.store.baseline_path.read_bytes(), self.store.overrides_path.read_bytes()
        self.edit()
        plan = self.store.prepare_reviewed_import(self.report)
        self.assertEqual(self.store.baseline_path.read_bytes(), baseline)
        self.assertEqual(self.store.overrides_path.read_bytes(), corrections)
        from roster_updater.ui.app import RosterUpdaterApp
        fake = SimpleNamespace(store=self.store, _set_busy=Mock(), _show_error=Mock())
        with patch('roster_updater.ui.app.filedialog.askopenfilename', return_value=str(self.report)), \
                patch('roster_updater.ui.app.messagebox.askyesno', return_value=False):
            RosterUpdaterApp.import_reviewed_roster(fake)
        self.assertEqual(self.store.baseline_path.read_bytes(), baseline)
        self.assertEqual(self.store.overrides_path.read_bytes(), corrections)
        fake._show_error.assert_not_called()

    def test_old_baseline_export_is_rejected(self):
        self.store.approve(self.result)
        with self.assertRaisesRegex(RosterError, 'older baseline'):
            self.store.prepare_reviewed_import(self.report)

    def test_correction_changes_since_export_are_rejected(self):
        self.store.save_override('alex', 'Lead', 'Manual')
        with self.assertRaisesRegex(RosterError, 'corrections changed since export'):
            self.store.prepare_reviewed_import(self.report)

    def test_workbook_edit_after_preview_is_rejected(self):
        plan = self.store.prepare_reviewed_import(self.report)
        self.edit()
        with self.assertRaisesRegex(RosterError, 'workbook changed after preview'):
            self.store.import_reviewed_roster(plan)

    def test_correction_or_baseline_change_after_preview_is_rejected(self):
        plan = self.store.prepare_reviewed_import(self.report)
        self.store.save_override('alex', 'Lead', 'Manual')
        with self.assertRaisesRegex(RosterError, 'corrections changed after preview'):
            self.store.import_reviewed_roster(plan)

    def test_missing_context_and_partial_rosters_are_rejected(self):
        workbook = load_workbook(self.report)
        del workbook[CONTEXT_SHEET]
        workbook.save(self.report)
        workbook.close()
        with self.assertRaisesRegex(RosterError, 'supported reviewed-roster export'):
            self.store.prepare_reviewed_import(self.report)
        write_report(self.report, self.result)
        self.edit({'E3': 'different.key'})
        with self.assertRaisesRegex(RosterError, 'retain every exported employee'):
            self.store.prepare_reviewed_import(self.report)

    def test_duplicate_keys_are_rejected(self):
        self.edit({'E3': 'alex'})
        with self.assertRaisesRegex(RosterError, 'Duplicate Logon'):
            self.store.prepare_reviewed_import(self.report)

    def test_missing_erp_column_is_not_silently_cleared(self):
        workbook = load_workbook(self.report)
        workbook['Roster Draft'].delete_cols(15)
        workbook.save(self.report)
        workbook.close()
        with self.assertRaisesRegex(RosterError, 'including ERP Category'):
            self.store.prepare_reviewed_import(self.report)

    def test_removed_rows_are_rejected(self):
        workbook = load_workbook(self.report)
        workbook['Roster Draft'].delete_rows(3)
        workbook.save(self.report)
        workbook.close()
        with self.assertRaisesRegex(RosterError, 'retain every exported employee'):
            self.store.prepare_reviewed_import(self.report)

    def test_failed_promotion_restores_baseline_corrections_and_audit(self):
        self.edit()
        plan = self.store.prepare_reviewed_import(self.report)
        baseline = self.store.baseline_path.read_bytes()
        corrections = self.store.overrides_path.read_bytes()
        replace = os.replace

        def fail_corrections(source, destination):
            if Path(destination) == self.store.overrides_path:
                raise PermissionError('Simulated locked corrections file')
            return replace(source, destination)

        with patch('roster_updater.reviewed_import.os.replace', side_effect=fail_corrections):
            with self.assertRaisesRegex(RosterError, 'previous baseline and corrections were restored'):
                self.store.import_reviewed_roster(plan)
        self.assertEqual(self.store.baseline_path.read_bytes(), baseline)
        self.assertEqual(self.store.overrides_path.read_bytes(), corrections)
        self.assertFalse(self.store.audit_path.exists())

    def test_context_is_chunked_and_keeps_four_visible_tabs(self):
        self.new['Office'] = 'Long office ' + 'x' * 20000
        write(self.monthly, 'All Users', CORE_HEADERS, [self.manager, self.new])
        result = self.store.confirm_logon_change(self.store.compare(self.monthly), 'taylor.old', 'taylor.cooper')
        write_report(self.report, result)
        workbook = load_workbook(self.report)
        self.addCleanup(workbook.close)
        self.assertEqual([sheet.title for sheet in workbook if sheet.sheet_state == 'visible'],
                         ['Summary', 'Review', 'Changes', 'Roster Draft'])
        self.assertEqual(workbook[CONTEXT_SHEET].sheet_state, 'veryHidden')
        self.assertGreater(workbook[CONTEXT_SHEET].max_row, 3)
        for row in workbook[CONTEXT_SHEET].iter_rows(min_row=2, max_col=1):
            self.assertLess(len(row[0].value), 32767)
            self.assertTrue(row[0].value.startswith('json:'))
        self.assertEqual(self.store.prepare_reviewed_import(self.report).reviewed.rows[1]['Office'], self.new['Office'])

    def test_context_changes_are_rejected(self):
        workbook = load_workbook(self.report)
        workbook[CONTEXT_SHEET]['A2'] = 'json:{}'
        workbook.save(self.report)
        workbook.close()
        with self.assertRaisesRegex(RosterError, 'review context changed'):
            self.store.prepare_reviewed_import(self.report)

    def test_supporting_identity_fields_can_be_corrected_after_confirmation(self):
        self.edit({'G3': 'New title', 'H3': 'New department', 'J3': 'New Manager', 'K3': 'manager@example.test'})
        plan = self.store.prepare_reviewed_import(self.report)
        self.assertEqual(plan.confirmed_links, {'taylor.cooper': 'taylor.old'})
        self.store.import_reviewed_roster(plan)
        self.assertEqual(self.store.load_baseline().rows[1]['Manager'], 'New Manager')

    def test_changed_baseline_after_preview_and_shared_lock_prevent_import(self):
        from roster_updater.storage import shared_lock
        plan = self.store.prepare_reviewed_import(self.report)
        with shared_lock(self.store.lock_path, 'Another user'):
            with self.assertRaisesRegex(RosterError, 'Another user'):
                self.store.import_reviewed_roster(plan)
        self.store.approve(self.result)
        with self.assertRaisesRegex(RosterError, 'baseline changed after preview'):
            self.store.import_reviewed_roster(plan)

    def test_rollback_restores_an_existing_audit_log(self):
        self.store.audit_path.write_text('Existing audit bytes\n', encoding='utf-8')
        plan = self.store.prepare_reviewed_import(self.report)
        before = self.store.audit_path.read_bytes()
        replace = os.replace

        def fail_receipt(source, destination):
            if Path(destination).name.startswith('reviewed_import_'):
                raise PermissionError('Simulated receipt failure')
            return replace(source, destination)

        with patch('roster_updater.reviewed_import.os.replace', side_effect=fail_receipt):
            with self.assertRaisesRegex(RosterError, 'were restored'):
                self.store.import_reviewed_roster(plan)
        self.assertEqual(self.store.audit_path.read_bytes(), before)
        self.assertEqual(self.store.load_baseline().rows[1]['Logon Name'], 'taylor.old')

    def test_ui_import_works_without_live_comparison_and_clears_old_display(self):
        from roster_updater.ui.app import RosterUpdaterApp
        self.edit()
        fake = SimpleNamespace(store=self.store, result=None, _set_busy=Mock(), _show_error=Mock(),
                               _clear_result_display=Mock(), pull_path_var=Mock(),
                               refresh_baseline_status=Mock(), status_var=Mock())
        with patch('roster_updater.ui.app.filedialog.askopenfilename', return_value=str(self.report)), \
                patch('roster_updater.ui.app.messagebox.askyesno', return_value=True), \
                patch('roster_updater.ui.app.messagebox.showinfo'):
            RosterUpdaterApp.import_reviewed_roster(fake)
        fake._show_error.assert_not_called()
        fake._clear_result_display.assert_called_once()
        fake.pull_path_var.set.assert_called_once_with('')
        self.assertEqual(self.store.load_baseline().rows[1]['Capability'], 'Innovate')


if __name__ == '__main__':
    unittest.main()
