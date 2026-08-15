"""Data passed from the Desktop interface to a processing job."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class JobOptions:
    """A validated snapshot of the processing choices made in the GUI."""

    input_file: Path
    mode: str = "Auto"
    dpi: int = 180
    excel_render_mode: str = "Both"
    max_autofit_columns: int = 60
    output_directory: Path | None = None

