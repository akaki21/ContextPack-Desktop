# Contributing

Issues and pull requests are welcome. Keep generated documents, personal data, `.venv`, and OCR model binaries out of commits.

Changes must preserve these safety and compatibility guarantees:

- atomic package finalization and source-hash collision protection;
- bounded/sparse Excel extraction for extreme dimensions;
- Excel read-only opening with macros, events, and automatic external-link updates disabled;
- no save operation against the source workbook;
- authoritative `workbook-layout` plus guarded `auto-layout` behavior;
- per-sheet AutoFit decisions in `print-layout-report.json`;
- source-page provenance for PDF text and images;
- portable paths with no user-specific Windows directories.

When output structure, commands, defaults, or safety behavior changes, update `README.md`, `README.ka.md`, `docs/PACKAGE_FORMAT.md`, `CHANGELOG.md`, and `CHANGELOG.ka.md` in the same pull request.

```powershell
.\setup.ps1
.\check-environment.ps1
.\tests\test-static.ps1
.\.venv\Scripts\python.exe -m unittest discover -s tests -p "test_*.py"
.\contextpack.ps1 ".\examples\sample.csv"
```
