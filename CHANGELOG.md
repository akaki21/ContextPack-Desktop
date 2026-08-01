# Changelog

All notable user-facing changes to ContextPack Desktop are recorded here.

[ქართული ვერსია](CHANGELOG.ka.md)

## 2.2.0 — 2026-08-01

- Added a dependency-free Tkinter desktop GUI and double-click Windows launcher.
- Added structured UTF-8 job events, stage-based progress, detailed logs, and cooperative cancellation.
- Added environment check/repair, result-folder opening, and AI-task copy actions.
- Added a destination-folder prompt before every GUI conversion and portable `-OutputDirectory` routing across the backend.
- Added live Georgian/English localization for the complete GUI, including dialogs, validation, progress messages, and copied AI tasks.
- Exposed the Excel AutoFit column limit through the main router.
- Added GUI core tests and beginner-focused English/Georgian guides.

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
