# ContextPack package format v2.3

Packages are finalized after extraction completes. Unresolved OCR/coverage problems are reported explicitly; a generated package is not a validated engineering source. The GUI surfaces manifest warnings when opening the result.

By default packages are written under the project `output` folder. GUI jobs ask for a destination before every run, and CLI callers can pass `-OutputDirectory`. Atomic temporary and final package paths remain inside the selected destination.

Common files:

- `manifest.json` — machine-readable provenance, SHA-256 source identity, outputs, settings, and warnings.
- `quality-report.md` — human-readable review priorities.
- `AI-HANDOFF.md` / `AI-HANDOFF.ka.md` — English and Georgian AI instructions.
- the unchanged source file, or (PDF/HTML only, opt-in) `source-reference.json` with its original path and hash. Reference packages depend on that external file.

PDF packages additionally contain:

- main page-aware Markdown (same text as `page-text.md`, including page-image OCR recovery);
- `page-text.md` with `<!-- source-page: N -->` provenance markers;
- selected `pages/page-NNN.png` (default maximum 100 for PDF packages; every page still receives text extraction);
- optional searchable OCR PDF;
- `pdf-metrics.json` with per-page input/recovered character counts, sparse/failed OCR flags, saved image references, and OCR confidence/word-box locations. The text input can be the searchable OCR PDF; `input_text_characters` means before page-image recovery, not necessarily native text in the original;
- optional `ocr-pages/page-NNN.txt` and `.tsv` from sparse-page recovery. These recovered words are not inserted into the OCR PDF.

HTML packages contain `source.md`, `tables.json`, `anchors.json`, `assets.json`, `html-metrics.json` and the common package metadata. Markdown uses HTML tables to retain merged headers/cells. `tables.json` stores row/column positions (one-based), rowspan/colspan, unchanged value strings, inline markup, nearby context and a rectangular grid of cell IDs. A merged cell occupies several grid slots with the same ID. No decimal/unit/ditto-mark interpretation is applied. Single-cell layout wrappers are flattened into text. Nested tables have separate records; every table and the source document must still be checked for engineering use.

HTML originals are under `original/` in Copy mode. Safe local raster images within the document folder are copied with their relative paths. External, missing, oversized, vector, or out-of-folder images are reported and are never downloaded. Reference mode stores available local images under `assets/`. Scripts and active markup are removed from the generated reading view, while the original is unchanged. HTML5 parsing repairs malformed source markup; original markup remains authoritative.

Excel packages additionally contain:

- `workbook-info.md`;
- lightweight `values.md` and `formulas.md` indexes;
- per-sheet files under `sheets-data/NN-sheet-name/`;
- authoritative PDF/PNG under `rendered-sheets/workbook-layout/` when requested;
- guarded convenience PDF/PNG under `rendered-sheets/auto-layout/` when requested;
- `print-layout-report.json` with per-sheet AutoFit decisions and preserved settings;
- `excel-metrics.json`.

`Both` is the default Excel render mode and therefore creates two visual render sets. `Workbook` and `AutoFit` create only their selected set. AutoFit uses populated-cell bounds and adaptively splits wide tables across horizontal pages (roughly eight populated columns per page) while leaving page height unlimited. It skips risky sheets (hidden/empty, over the width limit, manual page breaks, or drawing objects). The default width limit is 60 populated columns and can be set from 10 to 200 with `-MaxAutoFitColumns` on `excel-package.ps1`. Merged cells do not block AutoFit but are recorded as a visual-review warning.

The source workbook is opened read-only, with macros, events, and automatic link updates disabled, and is never saved. `manifest.json` records the selected render mode, DPI, AutoFit width limit, safety settings, and which render folders exist. `print-layout-report.json` records the decision, inferred print area, preserved print titles, manual page-break counts, and drawing-object count for every worksheet and requested layout.

## Minimal AI handoff

Do not treat every generated file as mandatory input. Start with common metadata and the lightweight text indexes, then open only relevant sheet files or page images. For Excel, the source workbook and `workbook-layout` are authoritative; `auto-layout` is a convenience view. When Codex can access the package locally, the folder path plus a clear task is sufficient.

PDF/HTML packages include a `cache` record in `manifest.json`. Reuse requires matching source SHA-256, settings, processing code/dependency versions, selected OCR models/binary (PDF), HTML image dependencies, manifest content and every generated file hash. Reuse is local integrity checking, not engineering validation. `-Force` bypasses reuse. When generated files or the manifest have changed, rebuilding retains the old directory as `.previous-<id>`; a healthy generated package can be atomically replaced. User notes belong outside generated packages.

Different sources sharing a base name receive a source-hash suffix. Occupied short-hash destinations are checked before use. A source inside its destination package is rejected. Failed jobs do not replace a valid package. Excel retains its established package/rebuild workflow; the new cache and Reference controls apply to PDF/HTML only.

`render-pdf-pages.py --max-pages` retains the legacy all-or-none PNG safety behavior used by Excel, but now preserves all page text even when PNG rendering is skipped. `--max-render-pages` is a separate saved-image budget; the PDF wrapper sets it to 100 by default. Renders are bounded to 24 million pixels; actual DPI is recorded.


For 2.3.1 economical PDF storage, selected-page extraction and durable review history, see [Library storage and review](LIBRARY_REVIEW.md).
