from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from PIL import Image, ImageDraw, ImageFont


def load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = (
        Path(r"C:\Windows\Fonts\arial.ttf"),
        Path(r"C:\Windows\Fonts\calibri.ttf"),
    )
    for candidate in candidates:
        if candidate.is_file():
            return ImageFont.truetype(str(candidate), size=size)
    return ImageFont.load_default()


def create_ocr_image(path: Path) -> Image.Image:
    image = Image.new("RGB", (1800, 600), "white")
    draw = ImageDraw.Draw(image)
    title_font = load_font(72)
    body_font = load_font(48)
    draw.text((100, 100), "CONTEXT PACK OCR", fill="black", font=title_font)
    draw.text((100, 260), "Invoice number: 12345", fill="black", font=body_font)
    draw.text((100, 360), "Total amount: 987.65", fill="black", font=body_font)
    image.save(path, dpi=(300, 300))
    return image


def create_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Estimate"
    sheet.append(["Item", "Quantity", "Unit price", "Total"])
    sheet.append(["Concrete", 12, 95.5, "=B2*C2"])
    sheet.append(["Rebar", 3, 780, "=B3*C3"])
    sheet["A1"].font = Font(bold=True)
    sheet["B1"].font = Font(bold=True)
    sheet["C1"].font = Font(bold=True)
    sheet["D1"].font = Font(bold=True)
    sheet.freeze_panes = "A2"
    sheet.print_title_rows = "1:1"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    for column, width in {"A": 24, "B": 12, "C": 14, "D": 14}.items():
        sheet.column_dimensions[column].width = width
    workbook.save(path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("output_dir")
    args = parser.parse_args()

    output = Path(args.output_dir).resolve()
    output.mkdir(parents=True, exist_ok=True)
    image = create_ocr_image(output / "sample-scan.png")
    image.save(output / "sample-scan.pdf", "PDF", resolution=300.0)
    create_workbook(output / "sample.xlsx")
    (output / "sample.md").write_text(
        "# ContextPack Markdown fixture\n\nA small bilingual check: English / ქართული.\n",
        encoding="utf-8",
    )
    print(output)


if __name__ == "__main__":
    main()
