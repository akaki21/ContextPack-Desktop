# ContextPack package format v2.1

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
- authoritative PDF/PNG under `rendered-sheets/workbook-layout/` when requested;
- guarded convenience PDF/PNG under `rendered-sheets/auto-layout/` when requested;
- `print-layout-report.json` with per-sheet AutoFit decisions and preserved settings;
- `excel-metrics.json`.

`Both` is the default Excel render mode. AutoFit uses populated-cell bounds, fits to one page wide with unlimited page height, and skips risky sheets (hidden/empty, over the width limit, manual page breaks, or drawing objects). The source workbook is opened read-only, with macros, events, and automatic link updates disabled, and is never saved.

When an existing package has the same source hash, a successful rebuild replaces it. A different source with the same base name receives an eight-character SHA-256 suffix, preventing accidental mixing.
