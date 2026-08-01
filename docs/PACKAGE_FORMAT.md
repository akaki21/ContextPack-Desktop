# ContextPack package format v1

Every complete package is finalized only after all processing steps succeed.

Common files:

- `manifest.json` — machine-readable provenance, SHA-256 source identity, outputs, settings, and warnings.
- `quality-report.md` — human-readable review priorities.
- `AI-HANDOFF.md` / `AI-HANDOFF.ka.md` — English and Georgian AI instructions.
- the unchanged source file.

PDF packages additionally contain:

- main MarkItDown output;
- `page-text.md` with `<!-- source-page: N -->` provenance markers;
- `pages/page-NNN.png`;
- optional searchable OCR PDF;
- `pdf-metrics.json`.

Excel packages additionally contain:

- `workbook-info.md`;
- lightweight `values.md` and `formulas.md` indexes;
- per-sheet files under `sheets-data/NN-sheet-name/`;
- rendered workbook PDF and PNG pages;
- `excel-metrics.json`.

When an existing package has the same source hash, a successful rebuild replaces it. A different source with the same base name receives an eight-character SHA-256 suffix, preventing accidental mixing.
