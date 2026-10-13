from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from PIL import Image
from contextpack.pdf.extraction import extract_pdf, text_needs_ocr


class PdfExtractionTests(unittest.TestCase):
    def make_pdf(self, root: Path):
        source = root / 'three-pages.pdf'
        images = [Image.new('RGB', (100, 100), 'white') for _ in range(3)]
        try:
            images[0].save(source, 'PDF', save_all=True, append_images=images[1:])
        finally:
            for image in images:
                image.close()
        return source

    def test_png_budget_never_discards_page_text_or_page_provenance(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = self.make_pdf(root)
            result = extract_pdf(source, root / 'pages', page_text=root / 'text.md', max_render_pages=1, dpi=96)
            self.assertEqual(result['page_count'], 3)
            self.assertEqual(result['rendered_pages'], 1)
            self.assertEqual(len(result['pages']), 3)
            self.assertIn('source-page: 3', (root / 'text.md').read_text(encoding='utf-8'))

    def test_recovery_runs_even_when_saved_images_are_disabled(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = self.make_pdf(root)
            recovered = 'Текст страницы с числом 14,5 и единицей измерения м3.'
            with patch('contextpack.pdf.extraction.ocr_page', return_value=(recovered, {'attempted': True, 'status': 'extracted'})) as ocr:
                result = extract_pdf(source, root / 'pages', page_text=root / 'text.md', render_mode='None', tesseract='test')
            self.assertEqual(ocr.call_count, 3)
            self.assertEqual(result['fallback_ocr_pages'], [1, 2, 3])
            self.assertEqual(result['rendered_pages'], 0)
            self.assertEqual(result['sparse_pages'], [])
            self.assertTrue(all(p['review_required'] for p in result['pages']))
            self.assertIn(recovered, (root / 'text.md').read_text(encoding='utf-8'))

    def test_failed_ocr_stays_explicitly_unresolved(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = self.make_pdf(root)
            with patch('contextpack.pdf.extraction.ocr_page', return_value=('', {'attempted': True, 'status': 'failed'})):
                result = extract_pdf(source, root / 'pages', render_mode='None', tesseract='test')
            self.assertEqual(result['ocr_failed_pages'], [1, 2, 3])
            self.assertEqual(result['sparse_pages'], [1, 2, 3])
            self.assertEqual(result['total_characters'], 0)

    def test_cancellation_stops_before_processing_pages(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = self.make_pdf(root)
            token = root / 'cancel'
            token.touch()
            with self.assertRaises(InterruptedError):
                extract_pdf(source, root / 'pages', cancel_file=token)
            self.assertTrue(source.exists())

    def test_sparse_or_damaged_text_is_detected_per_page(self):
        self.assertTrue(text_needs_ocr(''))
        self.assertTrue(text_needs_ocr('\ufffd' * 100))
        self.assertTrue(text_needs_ocr('\ue000' * 100))
        self.assertFalse(text_needs_ocr('Readable русский ქართული text. ' * 4))


if __name__ == '__main__':
    unittest.main()
