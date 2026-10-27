import html
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

from catalogue_snapshot import build_snapshot, publish_snapshot, record_id


class CatalogSnapshotTests(unittest.TestCase):
    def row(self, number='101', **changes):
        return {
            'Subject Code': 'BIO', 'Course Number+': number, 'Section': 1,
            'TITLE': 'Biology', 'INSTRUCTOR1': None, **changes,
        }

    def snapshot(self, rows, **changes):
        return build_snapshot(rows, '2027-spring', 'Spring 2027', **changes)

    def test_hash_ignores_row_order_and_generation_time(self):
        rows = [self.row(), self.row('102')]
        first = self.snapshot(rows, generated_at='2026-10-27T00:00:00+00:00')
        second = self.snapshot(list(reversed(rows)), generated_at='2026-10-28T00:00:00+00:00')
        self.assertEqual(first['snapshot_id'], second['snapshot_id'])

    def test_normalizes_missing_values_whitespace_and_section_numbers(self):
        first = self.snapshot([self.row()])
        second = self.snapshot([self.row(**{'Section': '01', 'TITLE': ' Biology ', 'INSTRUCTOR1': float('nan')})])
        self.assertEqual(first['snapshot_id'], second['snapshot_id'])

    def test_content_changes_additions_and_removals_change_hash(self):
        baseline = self.snapshot([self.row(), self.row('102')])['snapshot_id']
        for rows in ([self.row()], [self.row(), self.row('102'), self.row('103')],
                     [self.row(TITLE='Updated Biology'), self.row('102')]):
            with self.subTest(rows=rows):
                self.assertNotEqual(baseline, self.snapshot(rows)['snapshot_id'])

    def test_ambiguous_duplicates_preserve_all_variants(self):
        result = self.snapshot([self.row(), self.row(TITLE='Different course')])
        self.assertEqual(len(result['records']), 1)
        self.assertEqual(len(result['records'][0]['variants']), 2)
        self.assertEqual(result['identity_issues'][0]['variant_count'], 2)

    def test_exact_duplicates_do_not_create_false_changes(self):
        self.assertEqual(self.snapshot([self.row()])['snapshot_id'],
                         self.snapshot([self.row(), self.row()])['snapshot_id'])

    def test_lab_changes_are_retained_and_change_snapshot(self):
        lecture = self.row()
        lab = self.row('101L', **{'DAY1': 'M', 'TIME1-start': '01:00PM'})
        baseline = self.snapshot([lecture, lab])
        changed = self.snapshot([lecture, {**lab, 'DAY1': 'T'}])
        self.assertEqual(len(baseline['records']), 2)
        self.assertNotEqual(baseline['snapshot_id'], changed['snapshot_id'])

    def test_terms_are_part_of_identity_and_content(self):
        baseline = self.snapshot([self.row()])
        changed = build_snapshot([self.row()], '2027-fall', 'Fall 2027')
        self.assertNotEqual(baseline['snapshot_id'], changed['snapshot_id'])
        self.assertNotEqual(baseline['records'][0]['id'], changed['records'][0]['id'])

    def test_record_id_uses_snapshot_identity_normalization(self):
        first = record_id(self.row(**{'Section': 1}), '2027-spring')
        normalized = record_id(
            self.row(**{'Subject Code': ' bio ', 'Section': '01'}), '2027-spring'
        )
        self.assertEqual(first, normalized)

    @unittest.skipUnless(shutil.which('node'), 'Node.js is required for browser comparison tests')
    def test_browser_comparison_classifies_selected_changes_and_catalog_additions(self):
        source = Path(__file__).resolve().parents[1]
        script = f"""
const assert = require('node:assert/strict');
const {{ compareCatalogSnapshots }} = require({json.dumps(str(source / 'filter.js'))});
const row = title => ({{ 'Subject Code': 'BIO', 'Course Number+': '101', Section: '1', TITLE: title }});
const record = (id, ...variants) => ({{ id, variants }});
const baseline = {{ term: '2027-spring', schema_version: 1, records: [
  record('same', row('Same')), record('changed', row('Old')),
  record('removed', row('Removed')), record('ambiguous', row('A'), row('B'))
]}};
const current = {{ term: '2027-spring', schema_version: 1, records: [
  record('same', row('Same')), record('changed', row('New')),
  record('ambiguous', row('A')), record('added', row('Added'))
]}};
const result = compareCatalogSnapshots(baseline, current, ['same', 'changed', 'removed', 'ambiguous']);
assert.equal(result.unchanged, 1);
assert.equal(result.changed.length, 1);
assert.equal(result.removed.length, 1);
assert.equal(result.added.length, 1);
assert.equal(result.ambiguous.length, 1);
assert.throws(() => compareCatalogSnapshots(baseline, {{ ...current, term: '2027-fall' }}, []), /same term/);
"""
        result = subprocess.run(
            ['node', '-e', script], capture_output=True, text=True, cwd=source
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_ambiguous_variants_are_stable_across_row_order(self):
        rows = [self.row(), self.row(TITLE='Different course')]
        self.assertEqual(self.snapshot(rows)['snapshot_id'], self.snapshot(list(reversed(rows)))['snapshot_id'])

    def test_rejects_missing_identity_empty_catalog_and_missing_term(self):
        for rows in ([], [self.row(**{'Section': None})]):
            with self.subTest(rows=rows), self.assertRaises(ValueError):
                self.snapshot(rows)
        with self.assertRaises(ValueError):
            build_snapshot([self.row()], '', '')

    def test_archive_is_immutable_and_validated(self):
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            first = self.snapshot([self.row()], generated_at='2026-10-27T00:00:00+00:00')
            path = publish_snapshot(first, directory)
            original = path.read_bytes()
            second = self.snapshot([self.row()], generated_at='2026-10-28T00:00:00+00:00')
            self.assertEqual(path, publish_snapshot(second, directory))
            self.assertEqual(path.read_bytes(), original)
            corrupted = json.loads(original)
            corrupted['records'][0]['variants'][0]['TITLE'] = 'Corrupted'
            path.write_text(json.dumps(corrupted), encoding='utf-8')
            with self.assertRaisesRegex(ValueError, 'integrity'):
                publish_snapshot(second, directory)


class CatalogBuilderTests(unittest.TestCase):
    def test_build_from_another_directory_and_preserve_last_good_output(self):
        import pandas as pd

        source = Path(__file__).resolve().parents[1]
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            project = root / 'catalog project'
            project.mkdir()
            for name in ('course_info_to_html.py', 'catalogue_snapshot.py', 'htmltemplate.html', 'filter.js'):
                shutil.copy2(source / name, project / name)
            spreadsheet = project / 'merged_course_schedule.xlsx'
            course = {'Subject Code': 'BIO', 'Course Number+': '101', 'Section': 1,
                      'TITLE': 'Biology', 'DAY1': 'M', 'TIME1-start': '09:00AM',
                      'TIME1-end': '10:00AM'}
            lab = {'Subject Code': 'BIO', 'Course Number+': '101L', 'Section': 1,
                   'TITLE': 'Biology Lab', 'DAY1': 'T', 'TIME1-start': '09:00AM',
                   'TIME1-end': '10:00AM'}
            pd.DataFrame([course, lab]).to_excel(spreadsheet, index=False)
            command = [sys.executable, str(project / 'course_info_to_html.py')]
            result = subprocess.run(command, cwd=root, capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            output = project / 'courses.html'
            original = output.read_bytes()
            latest = project / 'catalog-snapshots' / 'latest.json'
            metadata = json.loads(latest.read_bytes())
            archive = project / metadata['snapshot_path']
            self.assertEqual(json.loads(archive.read_text())['snapshot_id'], metadata['snapshot_id'])
            self.assertIn(metadata['snapshot_id'].encode(), original)
            self.assertNotIn(b'{{CATALOG_', original)
            self.assertIsNone(metadata['source_updated_at'])
            self.assertIsNone(metadata['source_retrieved_at'])
            self.assertEqual(len(metadata['source_file_sha256']), 64)
            self.assertFalse((root / 'courses.html').exists())
            output_text = output.read_text(encoding='utf-8')
            self.assertIn('data-record-ids=', output_text)
            self.assertIn('id="catalogSnapshot"', output_text)
            self.assertIn('Print Corrigendum', output_text)
            section_ids = [record_id(course, '2027-spring'), record_id(lab, '2027-spring')]
            self.assertIn(html.escape(json.dumps(section_ids), quote=True), output_text)
            original_archive = archive.read_bytes()
            repeated = subprocess.run(command, cwd=project, capture_output=True, text=True)
            self.assertEqual(repeated.returncode, 0, repeated.stdout + repeated.stderr)
            self.assertEqual(json.loads(latest.read_text())['snapshot_id'], metadata['snapshot_id'])
            self.assertEqual(archive.read_bytes(), original_archive)
            last_good_output = output.read_bytes()
            last_good_latest = latest.read_bytes()
            pd.DataFrame([{'Subject Code': 'BIO', 'Course Number+': '101', 'Section': None}]).to_excel(spreadsheet, index=False)
            failed = subprocess.run(command, cwd=root, capture_output=True, text=True)
            self.assertNotEqual(failed.returncode, 0)
            self.assertIn('Missing section identity', failed.stdout)
            self.assertEqual(output.read_bytes(), last_good_output)
            self.assertEqual(latest.read_bytes(), last_good_latest)
            spreadsheet.unlink()
            missing = subprocess.run(command, cwd=root, capture_output=True, text=True)
            self.assertNotEqual(missing.returncode, 0)
            self.assertEqual(output.read_bytes(), last_good_output)
            self.assertEqual(latest.read_bytes(), last_good_latest)


if __name__ == '__main__':
    unittest.main()