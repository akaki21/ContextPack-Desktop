import json
from pathlib import Path
import tempfile
import unittest

from contextpack.core.review import initialize, append_event, read_events, verify_package_review
from contextpack.core.package_cache import file_hash


class ReviewTests(unittest.TestCase):
    def fixture(self, root):
        source = root / 'scan.pdf'
        source.write_bytes(b'unchanged original fixture')
        package = root / 'package'
        package.mkdir()
        (package / 'page-text.md').write_text('14.5 m2', encoding='utf-8')
        (package / 'page.png').write_bytes(b'image must not be copied')
        pointer = initialize(package, source, root / '.contextpack-reviews')
        directory = Path(pointer['journal_directory'])
        entry = {'record_id': 'cell-1', 'source_sha256': file_hash(source), 'source_revision': 'fixture edition',
                 'locator': {'page': 1, 'table': 'A', 'cell': 'r2c3'}, 'original': '14.5 m2',
                 'corrected': '14,5 m²', 'status': 'source_verified', 'reviewer': 'test reviewer',
                 'method': 'fixture comparison', 'raw_snapshot': pointer['raw_snapshot']}
        return source, package, directory, entry

    def test_rerun_preserves_events_and_raw_ocr_and_does_not_copy_images(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, package, directory, entry = self.fixture(root)
            append_event(directory, entry)
            before = (directory / 'events.jsonl').read_bytes()
            initialize(package, source, root / '.contextpack-reviews')
            self.assertEqual(before, (directory / 'events.jsonl').read_bytes())
            self.assertEqual(len(list((directory / 'raw').iterdir())), 1)
            raw = directory / 'raw' / entry['raw_snapshot']
            self.assertEqual((raw / 'page-text.md').read_text(), '14.5 m2')
            self.assertFalse((raw / 'page.png').exists())
            (package / 'page-text.md').write_text('new extraction')
            initialize(package, source, root / '.contextpack-reviews')
            self.assertEqual(len(list((directory / 'raw').iterdir())), 2)
            self.assertEqual(before, (directory / 'events.jsonl').read_bytes())
            self.assertEqual(source.read_bytes(), b'unchanged original fixture')

    def test_wrong_source_missing_locator_and_unverified_corrections_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, _, directory, entry = self.fixture(Path(tmp))
            for update in ({'source_sha256': 'different'}, {'locator': {}}, {'status': 'needs_review'},
                           {'reviewer': ''}, {'raw_snapshot': '../escape'}, {'event_sha256': 'forged'},
                           {'locator': {'region': [0, 0, 10, 10]}}):
                with self.subTest(update=update), self.assertRaises(ValueError):
                    append_event(directory, dict(entry, **update))
            self.assertEqual(read_events(directory), [])

    def test_status_history_is_append_only_and_corruption_is_detected(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, _, directory, entry = self.fixture(Path(tmp))
            append_event(directory, dict(entry, status='needs_review', corrected=None))
            append_event(directory, entry)
            self.assertEqual([e['status'] for e in read_events(directory)], ['needs_review', 'source_verified'])
            p = directory / 'events.jsonl'
            p.write_text(p.read_text(encoding='utf-8').replace('14,5', '15,5'), encoding='utf-8')
            with self.assertRaises(ValueError):
                append_event(directory, entry)

    def test_modified_raw_snapshot_is_never_overwritten(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, package, directory, entry = self.fixture(root)
            (directory / 'raw' / entry['raw_snapshot'] / 'page-text.md').write_text('human edit')
            with self.assertRaises(ValueError):
                initialize(package, source, root / '.contextpack-reviews')

    def test_source_revision_gets_a_separate_journal(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, package, directory, entry = self.fixture(root)
            append_event(directory, entry)
            source.write_bytes(b'new revision')
            second = initialize(package, source, root / '.contextpack-reviews')
            self.assertNotEqual(str(directory), second['journal_directory'])
            self.assertEqual(len(read_events(directory)), 1)

    def test_cache_review_integrity_allows_valid_append_but_rejects_missing_raw(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, package, directory, entry = self.fixture(Path(tmp))
            append_event(directory, entry)
            verify_package_review(package, file_hash(source))
            (directory / 'raw' / entry['raw_snapshot'] / 'page-text.md').unlink()
            with self.assertRaises(OSError):
                verify_package_review(package, file_hash(source))


if __name__ == '__main__':
    unittest.main()
