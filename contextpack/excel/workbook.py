"""Workbook-level naming and metadata helpers."""

from __future__ import annotations

import re
from typing import Any


def safe_sheet_folder(index: int, title: str) -> str:
    """Create a stable Windows-safe folder name for a worksheet."""

    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", title).strip().rstrip(".")
    safe = safe[:60] or "sheet"
    return f"{index:02d}-{safe}"


def calculation_mode(workbook: Any) -> str:
    """Return Excel's calculation mode without assuming calcPr is present."""

    calculation = getattr(workbook, "calculation", None)
    mode = getattr(calculation, "calcMode", None)
    return str(mode) if mode else "unspecified"

