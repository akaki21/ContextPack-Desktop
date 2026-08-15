"""Recognize the processing route for an input file.

Only formats with a dedicated pipeline are listed here. Every other extension
uses the generic document-to-Markdown route.
"""

from pathlib import Path


EXCEL_EXTENSIONS = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"})
IMAGE_EXTENSIONS = frozenset({".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"})


def classify_input(path: Path) -> str:
    """Return the internal processor name selected for *path*."""

    extension = path.suffix.lower()
    if extension == ".pdf":
        return "pdf"
    if extension in EXCEL_EXTENSIONS:
        return "excel"
    if extension in IMAGE_EXTENSIONS:
        return "image"
    return "document"

