"""Regression coverage for the monthly roster workflow (no real employee data)."""
import csv
from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

import roster_engine
from roster_updater.constants import BASELINE_HEADERS, CORE_HEADERS
from roster_updater.excel import write_baseline
from roster_updater.loading import load_roster
from roster_updater.models import RosterError
from roster_updater.reports import create_report_from_files, write_report
from roster_updater.storage import SharedRosterStore, shared_lock
from scripts.make_sample_data import row, write


class RosterWorkflowTests(unittest.TestCase):
    """Exercise the public monthly workflow with synthetic files and isolated storage."""
    def setUp(self):
        """Create a temporary baseline, monthly pull, and shared folder for each test.

        Register cleanup immediately so temporary files are removed even on failure.
        """
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.baseline = self.root / 'baseline.xlsx'
        self.monthly = self.root / 'monthly.xlsx'
        manager = row(1, 'Alex Manager', 'alex', 'alex@example.test', 'Manager',
                      'Operations', 'Enable', 'Alex Manager', 'alex@example.test', 'ERP')
        departed = row(2, 'Casey Departed', 'casey', 'casey@example.test', 'Designer',
                       'Design', 'Make', 'Alex Manager', 'alex@example.test')
        newcomer = row(2, 'Dana New', 'dana', 'dana@example.test', 'Coordinator',
                      'Operations', '', 'Alex Manager', 'alex@example.test')
        write(self.baseline, 'All_Users_Active_Info (1)', BASELINE_HEADERS, [manager, departed])
        self.monthly_rows = [dict(manager, Title='Senior Manager', Capability=''), newcomer]
        write(self.monthly, 'All Users', CORE_HEADERS, self.monthly_rows)
        self.store = SharedRosterStore(self.root / 'shared')
        self.store.initialize(self.baseline)

    def test_comparison_preserves_capability_and_flags_changes(self):
        """Check carry-forward, retained ERP data, new/departed counts, and unapplied suggestions."""
        result = self.store.compare(self.monthly)
        self.assertEqual(result.summary['New employees'], 1)
        self.assertEqual(result.summary['Missing from new pull'], 1)
        self.assertEqual(result.updated_roster[0]['Capability'], 'Enable')
        self.assertEqual(result.updated_roster[0]['ERP Category'], 'ERP')
        self.assertEqual(result.updated_roster[1]['Capability'], '')
        suggestion = next(item for item in result.review_items if item.issue == 'Capability suggestion')
        self.assertEqual((suggestion.logon_name, suggestion.suggested_value), ('dana', 'Enable'))
        self.assertTrue(any(change.field_name == 'Title' for change in result.changes))

    def test_correction_is_persisted_and_can_be_removed(self):
        """Reopen storage to prove corrections survive in JSON and removal restores fallback."""
        self.store.save_override('DANA', 'Make', 'Slack/manual', 'Approved')
        reopened = SharedRosterStore(self.store.root)
        result = reopened.compare(self.monthly)
        self.assertEqual(result.updated_roster[1]['Capability'], 'Make')
        self.assertEqual(result.capability_sources['dana'], 'Slack/manual')
        reopened.remove_override('dana')
        self.assertEqual(reopened.compare(self.monthly).updated_roster[1]['Capability'], '')

    def test_approval_creates_report_backup_and_audit(self):
        """Verify approval promotes rows and produces the expected backup, report, and audit entry."""
        result = self.store.compare(self.monthly)
        report, backup = self.store.approve(result)
        self.assertTrue(report.exists())
        self.assertTrue(backup.exists())
        self.assertEqual(len(load_roster(backup).rows), 2)
        self.assertEqual(self.store.load_baseline().rows[1]['Logon Name'], 'dana')
        with self.store.audit_path.open(newline='') as handle:
            records = list(csv.DictReader(handle))
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]['New Employees'], '1')
        workbook = load_workbook(report)
        self.addCleanup(workbook.close)
        self.assertEqual(workbook.sheetnames, ['Summary', 'Review', 'Changes', 'Updated Roster'])
        self.assertEqual(workbook['Updated Roster']['P1'].value, 'Capability Source')

    def test_stale_comparison_cannot_replace_new_baseline(self):
        """Ensure a result based on an older baseline cannot be approved again after promotion."""
        result = self.store.compare(self.monthly)
        self.store.approve(result)
        with self.assertRaisesRegex(RosterError, 'baseline changed'):
            self.store.approve(result)

    def test_lock_blocks_concurrent_updates_and_is_released(self):
        """Verify an existing lock blocks a second writer and is removed on normal exit."""
        with shared_lock(self.store.lock_path, 'Test'):
            with self.assertRaisesRegex(RosterError, 'Another user'):
                with shared_lock(self.store.lock_path, 'Conflicting update'):
                    self.fail('Concurrent lock was acquired')
        self.assertFalse(self.store.lock_path.exists())

    def test_csv_loading_and_duplicate_validation(self):
        """Accept a valid CSV and reject it after a duplicate employee key is appended."""
        source = self.root / 'monthly.csv'
        with source.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=CORE_HEADERS, extrasaction='ignore')
            writer.writeheader()
            writer.writerows(self.monthly_rows)
        self.assertEqual(len(load_roster(source).rows), 2)
        with source.open('a', newline='', encoding='utf-8') as handle:
            csv.DictWriter(handle, fieldnames=CORE_HEADERS, extrasaction='ignore').writerow(self.monthly_rows[0])
        with self.assertRaisesRegex(RosterError, 'Duplicate Logon Name'):
            load_roster(source)

    def test_file_report_api_and_legacy_imports(self):
        """Check old imports still resolve to the organized implementation and export a report."""
        self.assertIs(roster_engine.SharedRosterStore, SharedRosterStore)
        self.assertIs(roster_engine.create_report_from_files, create_report_from_files)
        report = self.root / 'export.xlsx'
        result = create_report_from_files(self.baseline, self.monthly, report)
        self.assertTrue(report.exists())
        self.assertEqual(len(result.updated_roster), 2)

    def test_desktop_entry_point_imports_without_starting_window(self):
        """Import the desktop entry point without launching Tkinter or needing a display."""
        import roster_app
        from roster_updater.ui.app import RosterUpdaterApp
        from roster_updater.ui.view import RosterView
        self.assertIs(roster_app.RosterUpdaterApp, RosterUpdaterApp)
        self.assertTrue(issubclass(RosterUpdaterApp, RosterView))

    def assert_table_filters(self, path, expected_tables):
        """Inspect saved workbook XML, not just the objects used to build it.

        Table sheets must have one table-owned filter and no sheet filter.
        expected_tables maps worksheet names to expected table ranges.
        """
        namespace = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        workbook = load_workbook(path)
        self.addCleanup(workbook.close)
        with ZipFile(path) as archive:
            for index, sheet in enumerate(workbook.worksheets, start=1):
                xml = ElementTree.fromstring(archive.read(f'xl/worksheets/sheet{index}.xml'))
                if sheet.title not in expected_tables:
                    continue
                expected_range = expected_tables[sheet.title]
                with self.subTest(sheet=sheet.title):
                    self.assertIsNone(xml.find('s:autoFilter', namespace))
                    self.assertIsNone(sheet.auto_filter.ref)
                    self.assertEqual(len(sheet.tables), 1)
                    table = next(iter(sheet.tables.values()))
                    self.assertEqual(table.ref, expected_range)
                    self.assertIsNotNone(table.autoFilter)
                    self.assertEqual(table.autoFilter.ref, expected_range)
                    first_column, header_row, last_column, _ = range_boundaries(expected_range)
                    self.assertEqual(
                        [column.name for column in table.tableColumns],
                        [str(sheet.cell(header_row, column).value)
                         for column in range(first_column, last_column + 1)],
                    )
            table_paths = [name for name in archive.namelist() if name.startswith('xl/tables/') and name.endswith('.xml')]
            self.assertEqual(len(table_paths), len(expected_tables))
            for name in table_paths:
                xml = ElementTree.fromstring(archive.read(name))
                filters = xml.findall('s:autoFilter', namespace)
                self.assertEqual(len(filters), 1)
                self.assertEqual(filters[0].get('ref'), xml.get('ref'))

    def test_report_filters_belong_to_tables(self):
        """Preserve all three report tables/filters without sheet-level duplicates."""
        result = self.store.compare(self.monthly)
        report = write_report(self.root / 'filters.xlsx', result)
        self.assert_table_filters(report, {
            'Review': f'A4:K{4 + len(result.review_items)}',
            'Changes': f'A4:E{4 + len(result.changes)}',
            'Updated Roster': f'A1:P{1 + len(result.updated_roster)}',
        })
        workbook = load_workbook(report)
        self.addCleanup(workbook.close)
        self.assertEqual(workbook['Summary'].auto_filter.ref, f'A5:B{5 + len(result.summary)}')
        self.assertEqual(workbook['Updated Roster']['B2'].value, 'Alex Manager')

    def test_baseline_filter_belongs_to_table(self):
        """Apply the same filter rule to baseline files created during setup/approval."""
        self.assert_table_filters(self.store.baseline_path, {'Current Baseline': 'A1:O3'})
        self.assertEqual(len(load_roster(self.store.baseline_path).rows), 2)

    def test_empty_report_sections_and_baseline_have_no_tables_or_filters(self):
        """Empty sections keep headings without creating invalid header-only tables."""
        result = replace(self.store.compare(self.monthly), review_items=[], changes=[], updated_roster=[])
        report = write_report(self.root / 'empty_report.xlsx', result)
        baseline = write_baseline(self.root / 'empty_baseline.xlsx', [])
        for path, titles in [(report, ['Review', 'Changes', 'Updated Roster']),
                             (baseline, ['Current Baseline'])]:
            workbook = load_workbook(path)
            self.addCleanup(workbook.close)
            for title in titles:
                with self.subTest(file=path.name, sheet=title):
                    self.assertFalse(workbook[title].tables)
                    self.assertIsNone(workbook[title].auto_filter.ref)


if __name__ == '__main__':
    unittest.main()

