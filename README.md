# ContextPack Desktop

ContextPack Desktop turns PDFs, scans, spreadsheets, and Office documents into AI-ready context packages. It combines lightweight Markdown for fast reading with originals, formulas, structure metadata, and page images for accurate verification.

> Markdown for reading. Originals for accuracy. Images for visual evidence.

[ქართული დოკუმენტაცია](README.ka.md)

## Features

- Convert PDF, DOCX, PPTX, XLSX, HTML, CSV, JSON, and other MarkItDown-supported formats to Markdown.
- OCR Georgian and English scans with Tesseract and OCRmyPDF.
- Build complete PDF packages containing Markdown, the source PDF, rendered page images, and bilingual AI handoff prompts.
- Build complete Excel packages containing displayed values, formulas, cached results, workbook metadata, the original workbook, rendered PDF/PNG sheets, and bilingual AI handoff prompts.
- Keep source documents local. `input`, generated `output`, the virtual environment, and downloaded OCR models are excluded from Git.
- Run from any folder on a Windows computer; no user-specific paths are embedded in the scripts.

## Requirements

- Windows 10 or 11
- PowerShell 5.1 or newer
- Python 3.10 or newer
- Internet access during initial setup
- Microsoft Excel, installed and activated, only for Excel PDF/PNG rendering

Tesseract 5 is required for OCR. `setup.ps1` detects an existing installation or installs the current UB Mannheim Windows package through `winget`. Tesseract itself does not publish an official installer for newer Windows versions; its documentation points Windows users to UB Mannheim builds.

## Install on a new computer

```powershell
git clone https://github.com/akaki21/ContextPack-Desktop.git
cd ContextPack-Desktop
Set-ExecutionPolicy -Scope Process Bypass
.\setup.ps1
.\check-environment.ps1
```

Setup performs the following local actions:

1. creates `.venv` inside the project;
2. installs pinned Python dependencies;
3. detects or installs Tesseract;
4. copies the installed Tesseract support files into the local `tessdata` folder;
5. downloads official fast English, Georgian, and orientation models when missing;
6. creates `input` and `output`.

If Tesseract is managed separately:

```powershell
.\setup.ps1 -SkipTesseractInstall
```

Set `CONTEXTPACK_TESSERACT` to a custom `tesseract.exe` path if automatic detection cannot find it.

## Quick start

You do not need to activate `.venv`; every script calls the project environment directly.

```powershell
.\convert-to-markdown.ps1 ".\input\document.docx"
```

Generated files are written to `output`.

## Commands

### Standard document to Markdown

```powershell
.\convert-to-markdown.ps1 ".\input\document.pdf"
```

Use this for text-based PDFs and supported Office/data formats.

### Scanned PDF with OCR

```powershell
.\ocr-and-convert.ps1 ".\input\scan.pdf"
```

Creates a searchable `_ocr.pdf` and a Markdown version. OCR uses `kat+eng`.

### Single image with OCR

```powershell
.\ocr-image.ps1 ".\input\page.png"
```

Creates `output\page.txt`.

### Complete PDF package

```powershell
.\pdf-package.ps1 ".\input\document.pdf"
```

For a scan:

```powershell
.\pdf-package.ps1 ".\input\scan.pdf" -Ocr
```

For small text or technical drawings:

```powershell
.\pdf-package.ps1 ".\input\drawing.pdf" -Dpi 240
```

The package contains the source PDF, Markdown, PNG pages, a manifest, and English/Georgian AI handoff files.

### Complete Excel package

```powershell
.\excel-package.ps1 ".\input\workbook.xlsx"
```

The package contains:

- the original workbook;
- `workbook-info.md` for sheet structure and notable workbook features;
- `formulas.md` for formulas, cached values, and number formats;
- `values.md` for stored values;
- rendered workbook PDF and page PNGs;
- `AI-HANDOFF.md` and `AI-HANDOFF.ka.md`.

Excel must have recalculated and saved the workbook for cached formula results to be current.

## Efficient AI handoff

Start with the smallest useful context:

1. Give the AI the generated Markdown and describe the exact goal.
2. Name relevant PDF pages, Excel sheets, periods, columns, or cell ranges.
3. Add the original only when exact verification or editing is needed.
4. Add page images only when layout, drawings, signatures, stamps, or difficult tables matter.

Each full package includes a ready-to-use handoff prompt. See [AI-HANDOFF.md](AI-HANDOFF.md) for general guidance.

## Privacy

ContextPack processes files locally. It does not upload documents to an AI service. You decide which generated files or originals to share afterward. Review sensitive documents before sharing them with any third party.

## Test the installation

```powershell
.\convert-to-markdown.ps1 ".\examples\sample.csv"
```

Expected result: `output\sample.md`.

## Troubleshooting

- PowerShell blocks scripts: run `Set-ExecutionPolicy -Scope Process Bypass` in the current terminal.
- Tesseract is installed in a custom location: set `$env:CONTEXTPACK_TESSERACT = 'D:\path\to\tesseract.exe'`.
- OCR output has errors: verify names, numbers, and critical facts against the source image/PDF.
- Excel rendering fails: confirm desktop Microsoft Excel is installed, activated, and can open the workbook.
- `ffmpeg` warning: it does not affect PDF, Excel, Word, or image processing; it matters only for audio/video features.

## License

[MIT](LICENSE)
