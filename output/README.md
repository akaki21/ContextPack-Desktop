# Output folder

Generated files are written here. Complete packages include `manifest.json`, `quality-report.md`, source provenance, and bilingual AI handoff prompts. Start with these lightweight files, then open originals or selected page images only when verification is needed.

- Text summary: provide the main `.md`.
- PDF verification: add the source/OCR PDF.
- Visual verification: add only relevant PNG pages.
- Excel analysis: provide `workbook-info.md`, `formulas.md`, `values.md`, and `print-layout-report.json`; use `workbook-layout`/the workbook for exact checks and `auto-layout` only for convenient reading.

PDF packages include page-aware text. Excel packages include per-sheet data under `sheets-data`. Full packages are finalized atomically, so failed builds do not replace valid results.
