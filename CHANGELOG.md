# Changelog

All notable user-facing changes to ContextPack Desktop are recorded here.

[ქართული ვერსია](CHANGELOG.ka.md)

## 2.3.1 — 2026-10-14 (source update)

- Optional temporary OCR PDFs, short scratch root, selected original page numbers and accurate output manifests.
- Source-hash-bound raw snapshots and append-only review journals survive cache reuse and forced regeneration.
- Georgian/English GUI OCR storage option and practical verification checklist.
- Portable Keep/Copy behavior and independent Excel/HTML rendering remain available.

## 2.3.0 — 2026-10-14 (source update)

- Added selectable English/Georgian/Russian OCR and a checksum-pinned Russian model.
- Added HTML packages with original anchors, merged-cell grids, local assets and unverified quality reports.
- PDF Auto now detects individual sparse/damaged pages. Added image OCR recovery, text/TSV provenance and explicit unresolved-page diagnostics.
- Added PDF image budgets, original references and verified PDF/HTML package reuse with code/settings/dependency invalidation.
- Preserve edited previous packages; reject source-overwriting text conversion and package destinations containing the input.
- Added GUI storage/OCR settings and fail-with-error behavior when no completed output is returned.
- Surface quality warnings in the completion dialog; handle native-tool stderr progress correctly under the Windows PowerShell GUI runner.
- Added regression and Windows PowerShell integration tests. Fixed JSON-array enumeration in the existing Excel E2E test under Windows PowerShell 5.
- Added runtime/tessdata overrides for isolated development without copying installed environments.

## 2.2.0 — 2026-08-01

- Added a dependency-free Tkinter desktop GUI and double-click Windows launcher.
- Added structured UTF-8 job events, stage-based progress, detailed logs, and cooperative cancellation.
- Added environment check/repair, result-folder opening, and AI-task copy actions.
- Added a destination-folder prompt before every GUI conversion and portable `-OutputDirectory` routing across the backend.
- Added live Georgian/English localization for the complete GUI, including dialogs, validation, progress messages, and copied AI tasks.
- Exposed the Excel AutoFit column limit through the main router.
- Added GUI core tests and beginner-focused English/Georgian guides.
- Added real bilingual GUI screenshots, a 60-second demo, and a checked-in sample conversion result.

## 2.1.0 — 2026-08-01

- Added Excel render modes: `Workbook`, `AutoFit`, and the default `Both`.
- Preserved the workbook's author-defined print layout as the authoritative render.
- Added a guarded one-page-wide AutoFit convenience render with unlimited page height.
- Added per-sheet AutoFit safety checks for hidden/empty sheets, wide populated ranges, manual page breaks, and drawing objects.
- Added `print-layout-report.json` and print-layout warnings in `quality-report.md`.
- Disabled Excel events and automatic external-link updates; workbooks remain read-only and are never saved.
- Added populated-range bounds and drawing/merged-cell metrics to Excel extraction.
- Expanded English and Georgian guidance for choosing files to provide to an AI.

## 2.0.0 — 2026-08-01

- Added the auto-routing `contextpack.ps1` command.
- Added atomic builds, SHA-256 source identity, and collision-safe package names.
- Added page-aware PDF text and per-sheet Excel values/formulas.
- Added `manifest.json`, `quality-report.md`, and bilingual AI handoff files.
- Added forced macro disabling and bounded sparse extraction for large Excel ranges.
- Added verified OCR model downloads and automated tests.
