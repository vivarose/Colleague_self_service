from datetime import datetime
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import warnings

import pandas as pd

from refresh_catalogue import EST, catalog_keys, parse_course, parse_schedule, refresh
from seat_allocations import allocation_key, merge_allocations, parse_allocations
from course_info_to_html import display_course_info


class SeatAllocationTests(unittest.TestCase):
    def test_normalizes_writing_intensive_and_sections_without_merging_labs(self):
        self.assertEqual(allocation_key(' bio ', '101W', '01'), ('BIO', '101W', '1'))
        self.assertEqual(allocation_key('BIO', '101', '1.0'), ('BIO', '101', '1'))
        self.assertNotEqual(allocation_key('BIO', '101L', '1'), ('BIO', '101', '1'))
        self.assertNotEqual(allocation_key('BIO', '101W', '1'), ('BIO', '101', '1'))

    def test_parses_open_and_rejects_missing_or_duplicate_rows(self):
        row = 'BIO-101W-01 BIOLOGY 18 9 9 0 0'
        self.assertEqual(parse_allocations(row)[0]['Fr Cap'], '9')
        self.assertEqual(parse_allocations('GOVT-230-01 DATA 25 OPEN OPEN 2 2')[0]['Fr Cap'], 'OPEN')
        for text in ('', row + '\n' + row, row + '\nBIO-102-01 BROKEN'):
            with self.subTest(text=text), self.assertRaises(ValueError):
                parse_allocations(text)

    def test_merge_uses_normalized_keys_and_preserves_rows(self):
        schedule = pd.DataFrame([
            {'Subject Code': 'BIO', 'Course Number+': '101W', 'Section': 1},
            {'Subject Code': 'BIO', 'Course Number+': '101L', 'Section': 1},
            {'Subject Code': 'BIO', 'Course Number+': '101', 'Section': 1},
        ])
        seats = pd.DataFrame(parse_allocations('BIO-101W-01 BIOLOGY 18 9 9 0 0'))
        merged = merge_allocations(schedule, seats)
        self.assertEqual(len(merged), 3)
        self.assertEqual(merged.iloc[0]['Fr Cap'], '9')
        self.assertTrue(pd.isna(merged.iloc[1]['Fr Cap']))
        self.assertTrue(pd.isna(merged.iloc[2]['Fr Cap']))
        with self.assertRaises(ValueError):
            merge_allocations(schedule, pd.concat([seats, seats]))

    def test_source_discrepancy_warns_without_changing_values(self):
        with warnings.catch_warnings(record=True) as caught:
            row = parse_allocations('BIO-101-01 BIOLOGY 18 4 4 4 4')[0]
        self.assertTrue(caught)
        self.assertEqual(row['Sec Capacity'], '18')
        self.assertEqual(row['Fr Cap'], '4')


class RefreshParserTests(unittest.TestCase):
    def schedule(self, code='BIO101W01', time='09:00AM-09:50AM'):
        headers = ['Course', 'Title', 'Day', 'Time', 'Room', 'Credit', 'Instructor', 'Method', 'Max Cap', 'Restr']
        cells = [code, 'BIOLOGY<br/>NOTE: First years only.', 'MWF', time, 'SCI 100', '1', 'A Person', 'LEC', '18', 'R']
        return ('<h3>Printable Course Schedule: 27/SP</h3><table><tr>' +
                ''.join(f'<th>{header}</th>' for header in headers) + '</tr><tr>' +
                ''.join(f'<td>{cell}</td>' for cell in cells) + '</tr></table>')

    def test_schedule_preserves_section_and_separates_notes_and_times(self):
        row = parse_schedule(self.schedule())[0]
        self.assertEqual(row['COURSE_NO'], 'BIO101W01')
        self.assertEqual(row['COMMENTS'], 'First years only.')
        self.assertEqual(row['TIME1'], '09:00AM 09:50AM')

    def test_schedule_title_survives_a_missing_catalog_entry(self):
        rendered = display_course_info({
            'Subject Code': 'BIO', 'Course Number+': '101', 'Section': 1,
            'Course Name': float('nan'), 'TITLE': 'NEW BIOLOGY COURSE'
        })
        self.assertIn('New Biology Course', rendered)

    def test_schedule_rejects_wrong_term_or_unknown_time(self):
        for page in (self.schedule().replace('27/SP', '27/FA'), self.schedule(time='bad')):
            with self.assertRaises(ValueError):
                parse_schedule(page)

    def test_incomplete_feed_preserves_last_good_published_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pd.DataFrame([
                {'Subject Code': 'BIO', 'Course Number+': str(number), 'Section': 1}
                for number in range(20)
            ]).to_excel(root / 'merged_course_schedule.xlsx', index=False)
            (root / 'courses.html').write_text('last good page')
            original = (root / 'merged_course_schedule.xlsx').read_bytes()
            with patch('refresh_catalogue.download', return_value=self.schedule()) as download:
                with self.assertRaisesRegex(ValueError, 'shrank by over 10%'):
                    refresh(root, datetime.now(EST).date().isoformat())
            self.assertEqual(download.call_count, 1)
            self.assertEqual((root / 'merged_course_schedule.xlsx').read_bytes(), original)
            self.assertEqual((root / 'courses.html').read_text(), 'last good page')

    def test_catalog_combined_numbers_and_narratives(self):
        self.assertEqual(catalog_keys('AFRST-267/367'), [('AFRST', '267'), ('AFRST', '367')])
        content = '''
        <div id="main"><h1><span>BIO-101</span> Biology</h1>
        <div class="desc"><p>A description.</p></div>
        <div class="sc_prereqs">Prerequisites: BIO-100</div>
        <div class="sc_coreqs">Corequisite: BIO-101L</div>
        <h3>Notes</h3>New note.</div>'''
        row = parse_course(content, ('BIO', '101'), 'https://example.test')
        self.assertEqual(row['Prerequisite Narrative'], 'BIO-100')
        self.assertEqual(row['Corequisite Narrative'], 'BIO-101L')
        self.assertEqual(row['Notes'], 'New note.')
        with self.assertRaises(ValueError):
            parse_course(content, ('BIO', '102'), 'https://example.test')


if __name__ == '__main__':
    unittest.main()
