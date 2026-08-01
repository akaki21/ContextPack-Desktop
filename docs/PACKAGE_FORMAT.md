# ContextPack package format v2.1

Every complete package is finalized only after all processing steps succeed.

By default packages are written under the project `output` folder. GUI jobs ask for a destination before every run, and CLI callers can pass `-OutputDirectory`. Atomic temporary and final package paths remain inside the selected destination.

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

`Both` is the default Excel render mode and therefore creates two visual render sets. `Workbook` and `AutoFit` create only their selected set. AutoFit uses populated-cell bounds and adaptively splits wide tables across horizontal pages (roughly eight populated columns per page) while leaving page height unlimited. It skips risky sheets (hidden/empty, over the width limit, manual page breaks, or drawing objects). The default width limit is 60 populated columns and can be set from 10 to 200 with `-MaxAutoFitColumns` on `excel-package.ps1`. Merged cells do not block AutoFit but are recorded as a visual-review warning.

The source workbook is opened read-only, with macros, events, and automatic link updates disabled, and is never saved. `manifest.json` records the selected render mode, DPI, AutoFit width limit, safety settings, and which render folders exist. `print-layout-report.json` records the decision, inferred print area, preserved print titles, manual page-break counts, and drawing-object count for every worksheet and requested layout.

## Minimal AI handoff

Do not treat every generated file as mandatory input. Start with common metadata and the lightweight text indexes, then open only relevant sheet files or page images. For Excel, the source workbook and `workbook-layout` are authoritative; `auto-layout` is a convenience view. When Codex can access the package locally, the folder path plus a clear task is sufficient.

When an existing package has the same source hash, a successful rebuild replaces it. A different source with the same base name receives an eight-character SHA-256 suffix, preventing accidental mixing.
