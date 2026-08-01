# Contributing

Issues and pull requests are welcome. Keep generated documents, personal data, `.venv`, and OCR model binaries out of commits. Preserve atomic package semantics, forced macro disabling, source hashes, and bounded/sparse Excel extraction.

```powershell
.\setup.ps1
.\check-environment.ps1
.\tests\test-static.ps1
.\.venv\Scripts\python.exe -m unittest discover -s tests -p "test_*.py"
.\contextpack.ps1 ".\examples\sample.csv"
```
