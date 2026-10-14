from __future__ import annotations

import argparse
import json
from pathlib import Path

import pypdfium2 as pdfium
from contextpack.pdf.extraction import text_needs_ocr


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect PDF text coverage for OCR auto-detection.")
    parser.add_argument("input_pdf")
    args = parser.parse_args()
    source = Path(args.input_pdf).resolve()
    document = pdfium.PdfDocument(str(source))
    counts: list[int] = []
    sparse: list[int] = []
    for index in range(len(document)):
        page = document[index]
        text_page = page.get_textpage()
        text = text_page.get_text_range().strip()
        counts.append(len(text))
        if text_needs_ocr(text):
            sparse.append(index + 1)
        text_page.close()
        page.close()
    document.close()
    needs_ocr = bool(sparse)
    print(json.dumps({"page_count": len(counts), "characters_per_page": counts, "sparse_pages": sparse, "needs_ocr": needs_ocr}))


if __name__ == "__main__":
    main()
