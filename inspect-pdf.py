from __future__ import annotations

import argparse
import json
from pathlib import Path

import pypdfium2 as pdfium


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect PDF text coverage for OCR auto-detection.")
    parser.add_argument("input_pdf")
    args = parser.parse_args()
    source = Path(args.input_pdf).resolve()
    document = pdfium.PdfDocument(str(source))
    counts: list[int] = []
    for index in range(len(document)):
        page = document[index]
        text_page = page.get_textpage()
        counts.append(len(text_page.get_text_range().strip()))
        text_page.close()
        page.close()
    document.close()
    sparse = [index + 1 for index, count in enumerate(counts) if count < 40]
    needs_ocr = bool(counts) and (sum(counts) / len(counts) < 80 or len(sparse) / len(counts) >= 0.5)
    print(json.dumps({"page_count": len(counts), "characters_per_page": counts, "sparse_pages": sparse, "needs_ocr": needs_ocr}))


if __name__ == "__main__":
    main()
