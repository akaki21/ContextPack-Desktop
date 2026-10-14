from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
RENDERER = ROOT / "render-pdf-pages.py"


class PdfRendererTests(unittest.TestCase):
    def test_page_limit_preserves_pdf_and_skips_png_explosion(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "many-pages.pdf"
            output = root / "pages"
            metrics = root / "metrics.json"
            images = [Image.new("RGB", (100, 100), "white") for _ in range(3)]
            try:
                images[0].save(source, "PDF", save_all=True, append_images=images[1:])
            finally:
                for image in images:
                    image.close()

            subprocess.run(
                [
                    sys.executable,
                    str(RENDERER),
                    str(source),
                    str(output),
                    "--metrics",
                    str(metrics),
                    "--max-pages",
                    "2",
                ],
                check=True,
            )

            payload = json.loads(metrics.read_text(encoding="utf-8"))
            self.assertEqual(payload["page_count"], 3)
            self.assertTrue(payload["render_skipped"])
            self.assertEqual(payload["rendered_pages"], 0)
            self.assertTrue(source.exists())
            self.assertEqual(list(output.glob("page-*.png")), [])
