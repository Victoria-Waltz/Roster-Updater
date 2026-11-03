"""Synthetic rename scenarios: review first, transfer history only after confirmation."""
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from openpyxl import load_workbook

from roster_updater.comparison import compare_rosters
from roster_updater.constants import BASELINE_HEADERS, CORE_HEADERS
from roster_updater.identity import find_logon_candidates
from roster_updater.models import RosterError
from roster_updater.reports import write_report
from roster_updater.storage import SharedRosterStore
from scripts.make_sample_data import row, write


class LogonChangeTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        # Fictional records model a surname/login/email change, not a new hire.
        self.manager = row(1, 'Alex Manager', 'alex', 'alex@example.test', 'Manager',
                           'Operations', 'Enable', 'Alex Manager', 'alex@example.test')
        self.old = row(2, 'Taylor Cooper', 'taylor.old', 'taylor.old@example.test', 'Designer',
                       'Design', 'Make', 'Alex Manager', 'alex@example.test', 'Design ERP')
        self.old['Last Name'] = 'Old'
        self.new = dict(self.old, **{'Logon Name': 'taylor.cooper', 'Last Name': 'Cooper',
                                    'Email Address': 'taylor.cooper@example.test', 'Capability': ''})
        self.baseline = self.root / 'baseline.xlsx'
        self.monthly = self.root / 'monthly.xlsx'
        self.write_files()
        self.store = SharedRosterStore(self.root / 'shared')
        self.store.initialize(self.baseline)

    def write_files(self, old_extra=(), new_extra=()):
        write(self.baseline, 'All Users', BASELINE_HEADERS, [self.manager, self.old, *old_extra])
        write(self.monthly, 'All Users', CORE_HEADERS, [self.manager, self.new, *new_extra])

    def confirm(self, result=None):
        return self.store.confirm_logon_change(result or self.store.compare(self.monthly),
                                               'taylor.old', 'taylor.cooper')

    def test_candidate_does_not_merge_or_transfer_data(self):
        self.store.save_override('taylor.old', 'Create', 'Slack/manual', 'Approved earlier')
        result = self.store.compare(self.monthly)
        self.assertEqual(result.summary['New employees'], 1)
        self.assertEqual(result.summary['Missing from new pull'], 1)
        self.assertEqual(len(result.logon_candidates), 1)
        candidate = result.logon_candidates[0]
        self.assertFalse(candidate.ambiguous)
        self.assertIn('Department', candidate.evidence)
        self.assertEqual(result.updated_roster[1]['ERP Category'], '')
        self.assertEqual(result.updated_roster[1]['Capability'], '')
        finding = next(item for item in result.review_items if item.issue == 'Possible Logon Name Change')
        self.assertEqual((finding.current_value, finding.suggested_value), ('taylor.old', 'taylor.cooper'))
        self.assertIn('Possible Logon Name Change', {item.issue for item in result.review_items})
        self.assertGreater(result.unresolved_count, 0)

    def test_confirmation_reclassifies_employee_and_retains_fields(self):
        self.store.save_override('taylor.old', 'Create', 'Slack/manual', 'Approved earlier')
        before = self.store.overrides_path.read_bytes()
        baseline_before = self.store.baseline_path.read_bytes()
        result = self.confirm()
        self.assertEqual(result.summary['New employees'], 0)
        self.assertEqual(result.summary['Missing from new pull'], 0)
        self.assertEqual(result.summary['Confirmed logon changes'], 1)
        self.assertEqual(result.updated_roster[1]['ERP Category'], 'Design ERP')
        self.assertEqual(result.updated_roster[1]['Capability'], 'Create')
        self.assertEqual(result.capability_sources['taylor.cooper'], 'Slack/manual')
        fields = {item.field_name for item in result.changes}
        self.assertTrue({'Logon Name', 'Last Name', 'Email Address'} <= fields)
        self.assertEqual(result.logon_candidates, [])
        self.assertEqual(self.store.overrides_path.read_bytes(), before)
        self.assertEqual(self.store.baseline_path.read_bytes(), baseline_before)
        # A fresh comparison (also used by Clear) discards the preview links.
        fresh = self.store.compare(self.monthly)
        self.assertEqual(fresh.summary['New employees'], 1)
        self.assertEqual(fresh.updated_roster[1]['ERP Category'], '')

    def test_approval_moves_correction_preserves_metadata_and_records_link(self):
        self.store.save_override('taylor.old', 'Create', 'Slack/manual', 'Keep these notes')
        original = self.store.load_overrides()['taylor.old']
        result = self.confirm()
        self.store.approve(result)
        overrides = self.store.load_overrides()
        self.assertNotIn('taylor.old', overrides)
        self.assertEqual(overrides['taylor.cooper'], original)
        self.assertTrue(list(self.store.history_dir.glob('corrections_before_logon_change_*.json')))
        self.assertTrue(list(self.store.history_dir.glob('confirmed_logon_changes_*.json')))
        next_month = self.store.compare(self.monthly)
        self.assertFalse(next_month.logon_candidates)
        self.assertEqual(next_month.updated_roster[1]['Capability'], 'Create')
        self.assertEqual(next_month.updated_roster[1]['ERP Category'], 'Design ERP')

    def test_new_key_correction_wins_over_old_key_correction(self):
        self.store.save_override('taylor.old', 'Create', 'Slack/manual')
        self.store.save_override('taylor.cooper', 'Innovate', 'Other', 'Newer choice')
        result = self.confirm()
        self.assertEqual(result.updated_roster[1]['Capability'], 'Innovate')
        self.store.approve(result)
        self.assertEqual(self.store.load_overrides()['taylor.cooper']['notes'], 'Newer choice')

    def test_baseline_capability_fallback_requires_confirmation(self):
        result = self.confirm()
        self.assertEqual(result.updated_roster[1]['Capability'], 'Make')
        self.assertEqual(result.capability_sources['taylor.cooper'], 'Baseline')

    def test_duplicate_names_and_multiple_candidates_are_ambiguous(self):
        duplicate = dict(self.old, **{'Logon Name': 'other.old'})
        self.write_files(old_extra=[duplicate])
        self.store.initialize(self.baseline)
        result = self.store.compare(self.monthly)
        self.assertEqual(len(result.logon_candidates), 2)
        self.assertTrue(all(candidate.ambiguous for candidate in result.logon_candidates))
        with self.assertRaisesRegex(RosterError, 'ambiguous'):
            self.confirm(result)
        # The inverse ambiguity (one old record, two incoming candidates) is also blocked.
        duplicate_new = dict(self.new, **{'Logon Name': 'other.new'})
        candidates = find_logon_candidates([self.old], [self.new, duplicate_new])
        self.assertEqual(len(candidates), 2)
        self.assertTrue(all(candidate.ambiguous for candidate in candidates))

    def test_weak_blank_and_already_matched_records_are_not_candidates(self):
        weak = dict(self.new, Title='', Department='', Manager='', **{'Manager Email Address': ''})
        self.assertFalse(find_logon_candidates([self.old], [weak]))
        blank = dict(self.old, **{'Display Name': ''})
        self.assertFalse(find_logon_candidates([blank], [dict(self.new, **{'Display Name': ''})]))
        self.assertFalse(find_logon_candidates([self.old], [self.old, self.new]))
        # Case/whitespace-only keys already match through normal key normalization.
        self.assertFalse(find_logon_candidates([self.old], [dict(self.old, **{'Logon Name': ' TAYLOR.OLD '})]))

    def test_invalid_and_duplicate_links_are_rejected(self):
        baseline = self.store.load_baseline()
        result = self.store.compare(self.monthly)
        with self.assertRaises(RosterError):
            compare_rosters(baseline, result.monthly_pull, {}, {'taylor.cooper': 'alex'})
        with self.assertRaises(RosterError):
            compare_rosters(baseline, result.monthly_pull, {},
                            {'taylor.cooper': 'taylor.old', 'TAYLOR.COOPER': 'taylor.old'})

    def test_changed_input_or_corrections_prevent_confirmation_and_approval(self):
        result = self.store.compare(self.monthly)
        self.store.save_override('taylor.old', 'Create', 'Manual')
        with self.assertRaisesRegex(RosterError, 'corrections changed'):
            self.confirm(result)
        confirmed = self.confirm()
        baseline_before = self.store.baseline_path.read_bytes()
        self.store.save_override('taylor.old', 'Lead', 'Manual')
        with self.assertRaisesRegex(RosterError, 'corrections changed'):
            self.store.approve(confirmed)
        self.assertEqual(self.store.baseline_path.read_bytes(), baseline_before)
        fresh = self.store.compare(self.monthly)
        write(self.monthly, 'All Users', CORE_HEADERS, [self.manager, dict(self.new, Title='Changed')])
        with self.assertRaisesRegex(RosterError, 'monthly file changed'):
            self.confirm(fresh)

    def test_report_shows_pending_evidence_and_confirmed_old_new_values(self):
        result = self.store.compare(self.monthly)
        path = write_report(self.root / 'pending.xlsx', result)
        workbook = load_workbook(path)
        self.addCleanup(workbook.close)
        rows = list(workbook['Review'].iter_rows(min_row=5, values_only=True))
        finding = next(row for row in rows if row[2] == 'Possible Logon Name Change')
        self.assertEqual(finding[3:5], ('taylor.old', 'taylor.cooper'))
        self.assertIn('Matching evidence', finding[10])
        self.assertEqual(finding[11], 'Logon name review')
        path = write_report(self.root / 'confirmed.xlsx', self.confirm(result))
        workbook = load_workbook(path)
        self.addCleanup(workbook.close)
        changes = list(workbook['Changes'].iter_rows(min_row=5, values_only=True))
        self.assertTrue(any(row[2:5] == ('Logon Name', 'taylor.old', 'taylor.cooper') for row in changes))

    def test_ui_requires_confirmation_and_does_not_apply_logon_as_capability(self):
        from roster_updater.ui.app import RosterUpdaterApp
        result = self.store.compare(self.monthly)
        item = next(item for item in result.review_items if item.issue == 'Possible Logon Name Change')
        fake = SimpleNamespace(store=self.store, result=result, _selected_review_item=Mock(return_value=item),
                               _display_result=Mock(), status_var=Mock(), _show_error=Mock())
        with patch('roster_updater.ui.app.messagebox.askyesno', return_value=False):
            RosterUpdaterApp.confirm_logon_change(fake)
        self.assertIs(fake.result, result)
        with patch('roster_updater.ui.app.messagebox.askyesno', return_value=True):
            RosterUpdaterApp.confirm_logon_change(fake)
        self.assertEqual(fake.result.confirmed_logon_changes, {'taylor.cooper': 'taylor.old'})
        fake._show_error.assert_not_called()
        fake.store = Mock()
        with patch('roster_updater.ui.app.messagebox.showinfo'):
            RosterUpdaterApp.apply_suggestion(fake)
        fake.store.save_override.assert_not_called()

    def test_ui_rerun_preserves_links_until_input_files_change(self):
        from roster_updater.ui.app import RosterUpdaterApp
        fake = SimpleNamespace(store=self.store, result=self.confirm(),
                               pull_path_var=Mock(), _set_busy=Mock(), _display_result=Mock(),
                               status_var=Mock(), _show_error=Mock())
        fake.pull_path_var.get.return_value = str(self.monthly)
        self.store.save_override('taylor.cooper', 'Lead', 'Manual')
        RosterUpdaterApp.run_update(fake)
        self.assertEqual(fake.result.confirmed_logon_changes, {'taylor.cooper': 'taylor.old'})
        self.assertEqual(fake.result.updated_roster[1]['Capability'], 'Lead')
        # New bytes mean a new monthly comparison; do not retain session links.
        write(self.monthly, 'All Users', CORE_HEADERS, [self.manager, dict(self.new, Title='Senior Designer')])
        RosterUpdaterApp.run_update(fake)
        self.assertEqual(fake.result.confirmed_logon_changes, {})
        fake._show_error.assert_not_called()

    def test_ui_clear_discards_preview_and_baseline_changes_block_confirmation(self):
        from roster_updater.ui.app import RosterUpdaterApp
        confirmed = self.confirm()
        fake = SimpleNamespace(store=self.store, result=confirmed, _display_result=Mock(),
                               status_var=Mock(), _show_error=Mock())
        with patch('roster_updater.ui.app.messagebox.askyesno', return_value=True):
            RosterUpdaterApp.clear_logon_confirmations(fake)
        self.assertEqual(fake.result.confirmed_logon_changes, {})
        self.assertEqual(fake.result.summary['New employees'], 1)
        # A competing baseline promotion invalidates an older identity proposal.
        self.store.approve(confirmed)
        with self.assertRaisesRegex(RosterError, 'baseline changed'):
            self.confirm(fake.result)


if __name__ == '__main__':
    unittest.main()
