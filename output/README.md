# Output folder

Generated files are written here. Complete packages include `manifest.json`, `quality-report.md`, source provenance, and bilingual AI handoff prompts. Start with these lightweight files, then open originals or selected page images only when verification is needed.

When Codex is on the same computer, give it the package folder path and the task. There is no need to attach every file.

- Text summary: provide the main `.md`.
- PDF verification: add the source/OCR PDF.
- Visual verification: add only relevant PNG pages.
- Excel analysis: provide `workbook-info.md`, `formulas.md`, `values.md`, `print-layout-report.json`, and only the relevant `sheets-data` files; use `workbook-layout`/the workbook for exact checks and `auto-layout` only for convenient reading.

PDF packages include page-aware text. Excel packages include per-sheet data under `sheets-data`. The default Excel `Both` mode stores two PDF/PNG render sets, so it uses more disk space; choose `Workbook` when only the authoritative print view is needed. Full packages are finalized atomically, so failed builds do not replace valid results.
