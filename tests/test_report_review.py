"""Verify the exported workbook supports focused review and a separate editable draft."""
from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import load_workbook

from roster_updater.comparison import compare_rosters
from roster_updater.constants import BASELINE_HEADERS, CORE_HEADERS
from roster_updater.loading import load_roster
from roster_updater.models import ReviewItem
from roster_updater.reports import write_report
from roster_updater.review_categories import CATEGORY_COLORS, CATEGORY_LABELS, category_for_issue
from roster_updater.storage import SharedRosterStore
from scripts.make_sample_data import row, write


class ReportReviewTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.alex = row(1, 'Alex Manager', 'alex', 'alex@example.test', 'Manager',
                        'Operations', 'Enable', 'Alex Manager', 'alex@example.test', 'ERP')
        self.blake = row(2, 'Blake Builder', 'blake', 'blake@example.test', 'Technician',
                         'Production', 'Make', 'Alex Manager', 'alex@example.test')
        baseline = self.root / 'baseline.xlsx'
        monthly = self.root / 'monthly.xlsx'
        write(baseline, 'All Users', BASELINE_HEADERS, [self.alex, self.blake])
        write(monthly, 'All Users', CORE_HEADERS, [self.alex, dict(self.blake, Title='Senior Technician')])
        self.result = compare_rosters(load_roster(baseline), load_roster(monthly))

    def export(self, result=None):
        path = write_report(self.root / 'review.xlsx', result or self.result)
        workbook = load_workbook(path)
        self.addCleanup(workbook.close)
        return workbook

    def test_unchanged_people_stay_in_draft_only(self):
        workbook = self.export()
        self.assertEqual(workbook.sheetnames, ['Summary', 'Review', 'Changes', 'Roster Draft'])
        self.assertEqual(workbook['Review']['A5'].value, 'blake')
        self.assertEqual(workbook['Review'].max_row, 5)
        self.assertEqual(workbook['Changes']['A5'].value, 'blake')
        self.assertEqual([workbook['Changes'].cell(5, c).value for c in (3, 4, 5)],
                         ['Title', 'Technician', 'Senior Technician'])
        self.assertEqual(workbook['Changes'].max_row, 5)
        self.assertEqual(workbook['Roster Draft'].max_row, 3)
        self.assertEqual(workbook['Roster Draft']['A1'].comment.author, 'TAIT Roster Updater')

    def test_categories_colors_and_filters_match_dashboard(self):
        findings = [
            ('Employee information changed', 'changed'),
            ('Missing from new pull', 'missing'),
            ('New employee', 'new'),
            ('Missing manager information', 'manager'),
            ('Manager not in active pull', 'manager'),
            ('Missing capability', 'capability'),
            ('Capability suggestion', 'capability'),
        ]
        items = [ReviewItem('blake', 'Blake Builder', issue) for issue, _ in findings]
        workbook = self.export(replace(self.result, review_items=items))
        review = workbook['Review']
        self.assertEqual(review['L4'].value, 'Category')
        self.assertEqual(next(iter(review.tables.values())).autoFilter.ref, 'A4:L11')
        self.assertIsNone(review.auto_filter.ref)
        for row_number, (issue, category) in enumerate(findings, start=5):
            self.assertEqual(category_for_issue(issue), category)
            self.assertEqual(review.cell(row_number, 12).value, CATEGORY_LABELS[category])
            for cell in review[row_number]:
                self.assertEqual(cell.fill.fgColor.rgb[-6:], CATEGORY_COLORS[category])
        self.assertEqual(workbook['Changes']['F5'].value, CATEGORY_LABELS['changed'])
        self.assertEqual(workbook['Changes']['E5'].fill.fgColor.rgb[-6:], CATEGORY_COLORS['changed'])
        summary_values = [cell.value for row in workbook['Summary'] for cell in row]
        for label in CATEGORY_LABELS.values():
            self.assertIn(label, summary_values)

    def test_no_changes_has_clear_empty_messages_and_complete_draft(self):
        result = compare_rosters(self.result.baseline, self.result.baseline)
        workbook = self.export(result)
        self.assertEqual(workbook['Review']['A5'].value, 'No changes or review issues found.')
        self.assertEqual(workbook['Changes']['A5'].value, 'No employee field changes found.')
        self.assertFalse(workbook['Review'].tables)
        self.assertFalse(workbook['Changes'].tables)
        self.assertEqual(workbook['Roster Draft'].max_row, 3)

    def test_reviewed_draft_can_be_loaded_as_a_baseline(self):
        workbook = self.export()
        draft = workbook['Roster Draft']
        draft['I3'] = 'Create'
        reviewed = self.root / 'reviewed.xlsx'
        workbook.save(reviewed)
        store = SharedRosterStore(self.root / 'shared')
        baseline = store.initialize(reviewed)
        self.assertEqual(len(baseline.rows), 2)
        self.assertEqual(baseline.rows[1]['Capability'], 'Create')
        self.assertEqual(baseline.rows[0]['ERP Category'], 'ERP')

    def test_guide_follows_counts_and_explains_workbook_edits(self):
        rows = [dict(self.alex, **{'Logon Name': f'person-{i}', 'Department': f'Department {i}'})
                for i in range(45)]
        workbook = self.export(replace(self.result, updated_roster=rows))
        summary = workbook['Summary']
        guide = next(cell for row in summary for cell in row if cell.value == 'Workbook guide')
        self.assertGreater(guide.row, 5 + 45)
        text = ' '.join(str(cell.value) for row in summary for cell in row if cell.value)
        self.assertIn('Workbook edits are not read automatically', text)
        self.assertIn('Initialize baseline', text)
        self.assertIn("does not import edits", text)


if __name__ == '__main__':
    unittest.main()
