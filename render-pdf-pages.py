from __future__ import annotations

import argparse
from pathlib import Path

import pypdfium2 as pdfium


def main() -> None:
    parser = argparse.ArgumentParser(description="Render every PDF page to a PNG image.")
    parser.add_argument("input_pdf")
    parser.add_argument("output_dir")
    parser.add_argument("--dpi", type=int, default=180)
    args = parser.parse_args()

    source = Path(args.input_pdf).resolve()
    destination = Path(args.output_dir).resolve()
    destination.mkdir(parents=True, exist_ok=True)

    document = pdfium.PdfDocument(str(source))
    digits = max(3, len(str(len(document))))
    scale = args.dpi / 72.0

    for index in range(len(document)):
        page = document[index]
        bitmap = page.render(scale=scale)
        image = bitmap.to_pil().convert("RGB")
        output = destination / f"page-{index + 1:0{digits}d}.png"
        image.save(output, format="PNG", optimize=True)
        print(output.name)

    print(f"Rendered {len(document)} pages at {args.dpi} DPI")


if __name__ == "__main__":
    main()
