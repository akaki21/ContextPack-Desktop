"""Detect local tools used by ContextPack processing pipelines."""

from __future__ import annotations

import importlib.util
import os
import shutil
from pathlib import Path

from .paths import PROJECT_ROOT, RUNNER


def preferred_output_directory() -> Path:
    """Choose a familiar existing folder for the output picker."""

    home = Path.home()
    candidates = [
        Path(os.environ.get("OneDrive", "")) / "Desktop" if os.environ.get("OneDrive") else None,
        home / "OneDrive" / "Desktop",
        home / "Desktop",
    ]
    return next((candidate for candidate in candidates if candidate is not None and candidate.is_dir()), home)


def find_tesseract() -> Path | None:
    """Find Tesseract using configuration, PATH, then common install paths."""

    candidates: list[Path] = []
    configured = os.environ.get("CONTEXTPACK_TESSERACT")
    if configured:
        candidates.append(Path(configured))
    located = shutil.which("tesseract.exe")
    if located:
        candidates.append(Path(located))
    for variable, relative in (
        ("ProgramFiles", Path("Tesseract-OCR") / "tesseract.exe"),
        ("ProgramFiles(x86)", Path("Tesseract-OCR") / "tesseract.exe"),
        ("LOCALAPPDATA", Path("Programs") / "Tesseract-OCR" / "tesseract.exe"),
    ):
        base = os.environ.get(variable)
        if base:
            candidates.append(Path(base) / relative)
    return next((candidate for candidate in candidates if candidate.is_file()), None)


def excel_is_available() -> bool:
    """Check whether Microsoft Excel exposes its Windows COM registration."""

    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
            return True
    except (ImportError, FileNotFoundError, OSError):
        return False


def environment_summary() -> dict[str, bool]:
    """Report whether the engine, OCR, and Excel rendering are ready."""

    models_ready = all(
        (PROJECT_ROOT / "tessdata" / f"{language}.traineddata").is_file()
        for language in ("kat", "eng", "osd")
    )
    return {
        "engine": RUNNER.is_file() and importlib.util.find_spec("markitdown") is not None,
        "ocr": find_tesseract() is not None and models_ready and importlib.util.find_spec("ocrmypdf") is not None,
        "excel": excel_is_available(),
    }

