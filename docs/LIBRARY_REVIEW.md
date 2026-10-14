# Economical PDF packages and durable review (2.3.1)

For a library with originals already stored locally:

```powershell
.\contextpack.ps1 document.pdf -Mode Auto -OcrLanguages rus+eng -PdfRenderMode None -SourcePolicy Reference -OcrPdfPolicy Temporary -OutputDirectory E:\Library\pilot
```

`Temporary` performs searchable-PDF OCR in a job-local scratch directory on the selected output volume, extracts text, and removes that OCR PDF and OCR scratch before publishing. OCRmyPDF's TEMP/TMP are also scoped there. Page-image fallback still works with `None`; its temporary rasters are removed, while text and available TSV word boxes remain. Failure/cancellation removes the incomplete build. Process termination or power loss can leave a hidden incomplete build, which is not a completed package.

`Keep` is the default and retains a successful OCR PDF. `Copy` still supplies an original copy for portable packages. These choices are independent. `Reference` requires the original path to remain accessible. The standalone legacy `ocr-and-convert.ps1` still retains its OCR PDF; use the package route for economical storage.

For deeply nested Windows libraries, set `$env:CONTEXTPACK_TEMP = 'E:\Engineering-Assistant\99_Temp\ocr'` before launching the CLI or GUI. The folder must be writable. This explicitly selects a short scratch root for OCRmyPDF and fallback Tesseract; only job-specific children are removed. Without it, scratch is inside the output build. Native OCR tools may fail at long paths even when Python can read the source. Set `TEMP`/`TMP` on E: too for other tools; no models/environment need reinstalling.

`None` controls PDF page PNG export only. Excel retains its existing workbook copy, data, print-layout PDFs and PNGs according to `ExcelRenderMode`. HTML retains structured tables/anchors and locally available image assets; `Reference` avoids copying the original HTML but does not disable those assets. Fast mode bypasses package metadata/review features. There is no universal text-only switch.

For a limited PDF pilot, pass `-PdfPages '1-3,9,15'` in Auto, Full or Ocr mode. Omitted means every page. The original is neither split nor modified; source page numbers remain unchanged. `pdf-metrics.json` records total source `page_count`, `extracted_pages` and `coverage`. Page selection is currently CLI-only. Sparse PNG selection is based on text in the extraction input (possibly an OCR PDF). Invalid/out-of-range selections fail without publishing a package.

GUI: Advanced processing options includes **PDF OCR copy: Keep / Temporary** in Georgian and English. For economical PDF storage select No saved images, Reference and Temporary. Existing defaults remain unchanged. Policy, page selection, language, source hash, code and model hashes contribute to cache identity. Manifest `ocr_pdf` and `page_images` are null when no such artifact is retained.

## Review and correction history

PDF and HTML packages include `review.json`, linking to the sibling `.contextpack-reviews/<source SHA-256>/`. This folder is outside the replaceable extraction package. It contains:

- `source.json`: original identity; edition must be established from the source, not guessed.
- `raw/<snapshot SHA-256>/`: immutable extracted text/table/TSV/metadata snapshots. No source PDF or page PNG is copied. Identical extractions reuse a snapshot; changed extraction keeps earlier snapshots.
- `events.jsonl`: append-only review history with a hash chain.
- `CHECKLIST.md`: practical verification steps.

`page-text.md` is the selected extracted text; `input-page-text.md` retains text before fallback OCR. The original PDF remains authoritative. Fallback `ocr-pages/*.txt` and TSV also remain raw. Corrections are journal overlays and never silently alter OCR. A later event with the same `record_id` supersedes that record while preserving its history.

Each event records `record_id`, `source_sha256`, `source_revision`, `locator` (page/clause/table/cell/region; regions require coordinate system), `original`, `corrected`, `status`, `reviewer`, `method`, and `raw_snapshot`. Numeric values remain strings. Unknown revisions should explicitly say “not established”. Optional contextual notes may be added. Generated event ID, UTC time, previous hash and event hash are added automatically.

Statuses are **extracted**, **needs_review**, **source_verified**. Only source_verified may contain a non-null correction; an unchanged checked value is repeated in corrected. Unclear values stay needs_review with corrected=null. Verification applies only to the locator recorded, never the whole page/document by implication.

```powershell
python review-journal.py append E:\Library\pilot\.contextpack-reviews\<hash> entry.json
python review-journal.py check E:\Library\pilot\.contextpack-reviews\<hash>
```

Open the original context and check decimals, signs, inequalities, powers/subscripts, units, row/column headings, merged cells and notes. TSV confidence and coordinates cover fallback OCR pages only. They are not an accuracy guarantee and may use a rotated/deskewed text-source coordinate system. Every page starts unverified, including pages with readable native text.

Cache reuse verifies generated files and referenced raw snapshot/journal integrity. Valid review appends do not invalidate extraction reuse. Forced conversion preserves the journal. A changed source hash gets a separate journal; old corrections are not applied to a new revision. A modified raw snapshot or damaged journal stops reuse/rebuild rather than overwriting evidence. A stale `.writer.lock` after a crash requires investigation before removing that lock. Hashes detect modification, not authorship or engineering correctness.

Move/back up the sibling `.contextpack-reviews` folder alongside portable packages if the review history is needed. `relative_to_output_root` remains usable after moving them together; the absolute first-path fields are provenance. Review history and private source material belong in the user's local library, not the public code repository.
