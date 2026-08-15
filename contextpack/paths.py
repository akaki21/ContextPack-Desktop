"""Stable paths shared by the GUI and its background runner."""

from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
RUNNER = PROJECT_ROOT / "contextpack-gui-runner.ps1"
OUTPUT_ROOT = PROJECT_ROOT / "output"

