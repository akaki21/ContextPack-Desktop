from pathlib import Path
import json
import tempfile
import unittest

from bs4 import BeautifulSoup
from contextpack.html.extractor import asset_signature, extract_html, extract_table


class HtmlExtractionTests(unittest.TestCase):
    def extract(self, html: str, directory: Path):
        source = directory / 'norm.html'
        source.write_text(html, encoding='utf-8')
        before = source.read_bytes()
        output = directory / 'package'
        metrics = extract_html(source, output)
        self.assertEqual(source.read_bytes(), before)
        return metrics, json.loads((output / 'tables.json').read_text(encoding='utf-8'))['tables'], (output / 'source.md').read_text(encoding='utf-8')

    def test_multilevel_headers_remain_aligned_with_exact_values_units_and_notes(self):
        html = '''<meta charset="utf-8"><h1 id="chapter">Дороги</h1><p><a href="#note">Примечание</a></p>
        <table id="t27"><tr><th rowspan="2">Материал</th><th colspan="2">Толщина</th></tr>
        <tr><th>15 см</th><th>20 см</th></tr><tr><td>Битум, т</td><td>14,5</td><td>−0,25</td></tr>
        <tr><td>Вода, м<sup>3</sup></td><td colspan="2">По проекту</td></tr></table>
        <p><a name="note"></a>Условия: ≤ 600; - // -; 1,6 маш.-ч.</p>'''
        with tempfile.TemporaryDirectory() as temporary:
            metrics, tables, markdown = self.extract(html, Path(temporary))
        table = tables[0]
        self.assertEqual(table['columns'], 3)
        self.assertEqual(table['grid'][0][0], table['grid'][1][0])
        self.assertEqual(table['grid'][0][1], table['grid'][0][2])
        self.assertEqual(table['grid'][3][1], table['grid'][3][2])
        self.assertEqual([c['text'] for c in table['cells'] if c['row'] == 3], ['Битум, т', '14,5', '−0,25'])
        self.assertIn('<sup>3</sup>', markdown)
        self.assertIn('rowspan="2"', markdown)
        self.assertIn('id="note"', markdown)
        self.assertIn('](#note)', markdown)
        self.assertIn('≤ 600', markdown)
        self.assertEqual(metrics['missing_source_targets'], [])

    def test_malformed_office_html_is_repaired_and_layout_wrappers_do_not_swallow_text(self):
        with tempfile.TemporaryDirectory() as temporary:
            metrics, tables, markdown = self.extract('''<meta charset="utf-8"><table><tr><td><p>Technical section</p></td></tr></table>
            <table><tr><td rowspan=2>A<td>B<tr><td>14,5</table><a name="tail"></a>Notes''', Path(temporary))
        self.assertIn('Technical section', markdown)
        self.assertEqual(metrics['table_count'], 1)
        self.assertEqual(tables[0]['grid'][0][0], tables[0]['grid'][1][0])
        self.assertIn('id="tail"', markdown)

    def test_rowspan_zero_stops_at_its_row_group(self):
        soup = BeautifulSoup('<table><thead><tr><td rowspan=0>A<td>B<tr><td>C</thead><tbody><tr><td>D<td>E</tbody></table>', 'html5lib')
        table = extract_table(soup.table, 1)
        self.assertEqual(table['cells'][0]['rowspan'], 2)
        self.assertNotEqual(table['grid'][1][0], table['grid'][2][0])

    def test_unavailable_assets_and_invalid_source_links_are_reported_without_network(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            metrics, _, markdown = self.extract('''<a href="#missing">X</a><script>alert(1)</script>
            <img src="https://example.invalid/private.png"><img src="../outside.png">
            <a href="javascript:alert(1)">No script</a>''', root)
        self.assertEqual(metrics['missing_source_targets'], ['missing'])
        self.assertGreaterEqual(len(metrics['warnings']), 3)
        self.assertNotIn('<script', markdown)
        self.assertNotIn('javascript:', markdown)
        self.assertNotIn('![](https:', markdown)

    def test_image_changes_invalidate_signature_and_copy_mode_keeps_relative_original_assets(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / 'norm.html'
            source.write_text('<img src="figure.png">', encoding='utf-8')
            image = root / 'figure.png'
            image.write_bytes(b'first image fixture')
            first = asset_signature(source)
            output = root / 'package'
            (output / 'original').mkdir(parents=True)
            extract_html(source, output)
            self.assertEqual((output / 'original/figure.png').read_bytes(), image.read_bytes())
            image.write_bytes(b'second image fixture')
            self.assertNotEqual(asset_signature(source), first)


if __name__ == '__main__':
    unittest.main()
