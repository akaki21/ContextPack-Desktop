# Document-library workflow

ContextPack 2.3 prepares local evidence for later review, explanations and search. It does not decide whether a standard is current or legally applicable, translate engineering requirements, or certify extracted numbers.

## Desktop settings

Run `setup.ps1` once after updating, then open `Start-ContextPack-GUI.cmd`. Choose a PDF, HTML or image and open **OCR and PDF/HTML storage settings…**.

| Setting | Meaning |
|---|---|
| OCR languages | Choose Russian, English, Georgian or a listed combination. Interface language and OCR language are independent. |
| PDF images: All | Save pages up to the image limit. |
| PDF images: Sparse | Save only pages whose text needs review. |
| PDF images: None | Save no PNGs. OCR may still render temporary images. |
| Maximum images | Default 100; 0 means unlimited. Text extraction visits every page regardless. |
| Copy original | Include an unchanged original; use for portable PDF/HTML packages. |
| Reference original | Store its path/hash instead; the original must stay available there. |
| Reprocess | Bypass reuse of an unchanged PDF/HTML package. |

```powershell
.\contextpack.ps1 '.\input\norm.pdf' -Mode Auto -OcrLanguages rus+eng -PdfRenderMode Sparse -MaxRenderPages 30 -SourcePolicy Reference -OutputDirectory 'D:\Library\processed'
.\contextpack.ps1 '.\input\norm.html' -SourcePolicy Reference -OutputDirectory 'D:\Library\processed'
```

`Full` creates a package without PDF OCR. `Ocr` requests recognition. `Fast` uses the old lightweight conversion: it does not provide structured HTML tables, reuse or provenance reports. New reference/reuse controls apply to PDF/HTML packages, not Excel or image-only output.

## Review the result

Read `quality-report.md` first. HTML `tables.json` stores one-based row/column positions, merged-cell spans, value strings and nearby context. Decimal commas, units and ditto marks are not interpreted. Verify the correct header and notes against the original before calculations. Markdown viewers without HTML table support can use JSON and the source HTML.

PDF `page-text.md` and main Markdown contain source-page markers. `pdf-metrics.json` identifies unresolved pages and recovered OCR. PDF Markdown is not a structured numerical table. Original PDF and selected images remain visual evidence.

Sparse/damaged pages left after searchable-PDF processing are rendered at up to 300 DPI and recognized with Tesseract. `ocr-pages/` stores text and word-box TSV files. Confidence values are diagnostics, not a correctness guarantee. Recovered words are not inserted into the searchable PDF; use Markdown and the provenance files.

HTML images outside the source folder, external images, unsupported/oversized images and missing images are reported in `assets.json`; no remote images are fetched. If a figure matters, inspect the original.

## Reuse and originals

Reuse checks source, settings, processing code/dependencies and all generated-file hashes. Selected OCR models/binary and local HTML image changes also invalidate it. `-Force` bypasses reuse. When a generated package was edited, rebuilding retains it as `.previous-<id>`. Keep reviewed guides and project notes outside generated folders.

Sources remain unchanged. Changes detected during processing prevent publication. A source inside its destination package is rejected. Reference packages save space but are not self-contained. Reuse is an integrity check, not engineering validation.

## Development

Run `tests/test-static.ps1`, Python unittest discovery, `tests/test-common.ps1`, `tests/test-library.ps1` and `tests/test-e2e.ps1`. Isolated checkouts can set `CONTEXTPACK_PYTHON` to an existing interpreter and `CONTEXTPACK_TESSDATA` to an existing model/config directory. Do not commit local paths. Default installs use `.venv` and `tessdata`.

The Russian model is pinned to the same [official tessdata_fast revision](https://github.com/tesseract-ocr/tessdata_fast/tree/87416418657359cb625c412a48b6e1d6d41c29bd) as the other models, with SHA-256 verification in setup. HTML explicitly uses the HTML5 parser; [Beautiful Soup documents why parser choice matters](https://www.crummy.com/software/BeautifulSoup/bs4/doc/#differences-between-parsers).
