from pathlib import Path
import json
import tempfile
import unittest
import subprocess
import sys

from contextpack.core.package_cache import pipeline_signature, seal_cache, validate_cache


class PackageCacheTests(unittest.TestCase):
    def test_cli_accepts_utf8_settings_independent_of_windows_stdin_encoding(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / 'requirements.txt').write_text('example==1', encoding='utf-8')
            settings = {'path': 'ქართული/ნორმა.pdf'}
            command = [sys.executable, '-B', str(Path(__file__).resolve().parents[1] / 'package-cache.py'), 'signature', str(root)]
            result = subprocess.run(command, input=json.dumps(settings, ensure_ascii=False).encode('utf-8'), capture_output=True, check=True)
            self.assertEqual(result.stdout.decode().strip(), pipeline_signature(root, settings))

    def make_package(self, root):
        package = root / 'package'
        package.mkdir()
        (package / 'manifest.json').write_text(json.dumps({'source': {'sha256': 'sourcehash'}, 'settings': {'ocr': True}}), encoding='utf-8')
        (package / 'source.md').write_text('Extracted text', encoding='utf-8')
        seal_cache(package, 'signature')
        return package

    def test_unchanged_package_reuses_but_changed_input_or_settings_do_not(self):
        with tempfile.TemporaryDirectory() as temporary:
            package = self.make_package(Path(temporary))
            self.assertTrue(validate_cache(package, 'signature', 'sourcehash')[0])
            self.assertFalse(validate_cache(package, 'newsettings', 'sourcehash')[0])
            self.assertFalse(validate_cache(package, 'signature', 'different-source')[0])

    def test_missing_edited_or_extra_output_invalidates_reuse(self):
        for mutation in ('missing', 'edited', 'extra'):
            with self.subTest(mutation=mutation), tempfile.TemporaryDirectory() as temporary:
                package = self.make_package(Path(temporary))
                if mutation == 'missing':
                    (package / 'source.md').unlink()
                elif mutation == 'edited':
                    (package / 'source.md').write_text('Human edit', encoding='utf-8')
                else:
                    (package / 'notes.md').write_text('Keep this', encoding='utf-8')
                self.assertFalse(validate_cache(package, 'signature', 'sourcehash')[0])

    def test_manifest_changes_and_corruption_do_not_count_as_valid_cache(self):
        with tempfile.TemporaryDirectory() as temporary:
            package = self.make_package(Path(temporary))
            path = package / 'manifest.json'
            manifest = json.loads(path.read_text(encoding='utf-8'))
            manifest['settings']['ocr'] = False
            path.write_text(json.dumps(manifest), encoding='utf-8')
            self.assertFalse(validate_cache(package, 'signature', 'sourcehash')[0])
            path.write_text('{broken', encoding='utf-8')
            self.assertFalse(validate_cache(package, 'signature', 'sourcehash')[0])

    def test_pipeline_code_and_settings_changes_invalidate_signature(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / 'requirements.txt').write_text('example==1', encoding='utf-8')
            script = root / 'processor.py'
            script.write_text('version = 1', encoding='utf-8')
            first = pipeline_signature(root, {'language': 'rus'})
            self.assertNotEqual(first, pipeline_signature(root, {'language': 'eng'}))
            script.write_text('version = 2', encoding='utf-8')
            self.assertNotEqual(first, pipeline_signature(root, {'language': 'rus'}))


if __name__ == '__main__':
    unittest.main()
