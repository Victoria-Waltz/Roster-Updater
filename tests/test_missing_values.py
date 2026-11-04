"""Missing markers must not become field changes after Excel export/import."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import load_workbook

from roster_updater.comparison import compare_rosters
from roster_updater.constants import BASELINE_HEADERS, CORE_HEADERS
from roster_updater.excel import write_baseline
from roster_updater.loading import load_roster
from roster_updater.reports import write_report
from roster_updater.utils import normalize_comparison_value
from scripts.make_sample_data import row, write


class MissingValueTests(unittest.TestCase):
    def test_only_missing_markers_are_equivalent(self):
        for value in (None, '', ' ', '-', "'-", "''-", ' – ', '—',
                      float('nan'), 'NaN', '\u200b-\ufeff', '\u00a0-\u00a0'):
            with self.subTest(value=repr(value)):
                self.assertEqual(normalize_comparison_value(value), '')
        for value in ('0', 'Nana', 'Anne-Marie', "O'Neil", '-Team', "'Team"):
            with self.subTest(value=value):
                self.assertEqual(normalize_comparison_value(value), value.casefold())

    def test_excel_roundtrip_reports_city_only_and_preserves_real_changes(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            person = row(1, 'Sample Employee', 'sample', 'sample@example.test',
                         'Designer', 'Design', '-', '-', '-')
            person['City'] = '-'
            baseline_path = root / 'baseline.xlsx'
            # Reproduce the actual baseline writer's dash escaping.
            write_baseline(baseline_path, [person])
            monthly = dict(person, City='Lititz')
            monthly_path = root / 'monthly.xlsx'
            write(monthly_path, 'All Users', CORE_HEADERS, [monthly])
            result = compare_rosters(load_roster(baseline_path), load_roster(monthly_path))
            self.assertEqual([change.field_name for change in result.changes], ['City'])
            self.assertEqual(result.summary['Changed fields'], 1)
            self.assertEqual(result.summary['Employees with changes'], 1)
            changed = [item for item in result.review_items
                       if item.issue == 'Employee information changed']
            self.assertEqual(changed[0].changed_fields, 'City')
            report_path = write_report(root / 'report.xlsx', result)
            workbook = load_workbook(report_path)
            try:
                self.assertEqual(workbook['Changes'].max_row, 5)
                self.assertEqual(workbook['Changes']['C5'].value, 'City')
                self.assertEqual(workbook['Changes']['E5'].value, 'Lititz')
            finally:
                workbook.close()
            # Real values replacing a missing manager still need review.
            monthly['Manager'] = 'New Manager'
            monthly['Manager Email Address'] = 'manager@example.test'
            monthly['Capability'] = 'Create'
            write(monthly_path, 'All Users', CORE_HEADERS, [monthly])
            result = compare_rosters(load_roster(baseline_path), load_roster(monthly_path))
            self.assertEqual({c.field_name for c in result.changes},
                             {'City', 'Manager', 'Manager Email Address', 'Capability'})

    def test_missing_capability_does_not_report_intentional_deletion(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            person = row(1, 'Sample Employee', 'sample', 'sample@example.test',
                         'Designer', 'Design', 'Create', 'Sample Employee', 'sample@example.test')
            write(root / 'baseline.xlsx', 'All Users', BASELINE_HEADERS, [person])
            for marker in ('-', "'-", 'NaN', ''):
                write(root / 'monthly.xlsx', 'All Users', CORE_HEADERS,
                      [dict(person, Capability=marker)])
                result = compare_rosters(load_roster(root / 'baseline.xlsx'),
                                         load_roster(root / 'monthly.xlsx'))
                self.assertFalse(result.changes)


if __name__ == '__main__':
    unittest.main()
