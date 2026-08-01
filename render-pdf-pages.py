from __future__ import annotations

import argparse
import json
from pathlib import Path

import pypdfium2 as pdfium


def normalize_text(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()


def main() -> None:
    parser = argparse.ArgumentParser(description="Render PDF pages and preserve page-level text provenance.")
    parser.add_argument("input_pdf")
    parser.add_argument("output_dir")
    parser.add_argument("--dpi", type=int, default=180)
    parser.add_argument("--page-text")
    parser.add_argument("--metrics")
    args = parser.parse_args()

    if not 72 <= args.dpi <= 600:
        raise SystemExit("DPI must be between 72 and 600")

    source = Path(args.input_pdf).resolve()
    destination = Path(args.output_dir).resolve()
    destination.mkdir(parents=True, exist_ok=True)

    document = pdfium.PdfDocument(str(source))
    page_count = len(document)
    digits = max(3, len(str(page_count)))
    scale = args.dpi / 72.0
    page_text_parts = [f"# Page-aware text — {source.name}\n"]
    metrics: list[dict[str, object]] = []

    for index in range(page_count):
        page_number = index + 1
        page = document[index]
        text_page = page.get_textpage()
        text = normalize_text(text_page.get_text_range())
        character_count = len(text)
        word_count = len(text.split())
        sparse = character_count < 40
        metrics.append(
            {
                "page": page_number,
                "characters": character_count,
                "words": word_count,
                "sparse_text": sparse,
            }
        )
        page_text_parts.extend(
            [
                f"<!-- source-page: {page_number} -->",
                f"## Page {page_number}",
                "",
                text if text else "_No extractable text on this page._",
                "",
            ]
        )

        bitmap = page.render(scale=scale)
        image = bitmap.to_pil().convert("RGB")
        output = destination / f"page-{page_number:0{digits}d}.png"
        image.save(output, format="PNG", optimize=True)
        print(output.name)
        image.close()
        bitmap.close()
        text_page.close()
        page.close()

    if args.page_text:
        Path(args.page_text).write_text("\n".join(page_text_parts), encoding="utf-8")
    if args.metrics:
        payload = {
            "source": source.name,
            "page_count": page_count,
            "total_characters": sum(int(item["characters"]) for item in metrics),
            "sparse_pages": [int(item["page"]) for item in metrics if item["sparse_text"]],
            "pages": metrics,
        }
        Path(args.metrics).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    document.close()
    print(f"Rendered {page_count} pages at {args.dpi} DPI")


if __name__ == "__main__":
    main()
