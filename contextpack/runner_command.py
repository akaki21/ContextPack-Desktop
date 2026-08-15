"""Build the argument list used to start the PowerShell GUI runner."""

from __future__ import annotations

import os
import shutil
from pathlib import Path

from .job_options import JobOptions
from .paths import RUNNER


def powershell_executable() -> str:
    """Find Windows PowerShell without relying on shell command parsing."""

    system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    candidate = system_root / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
    return str(candidate if candidate.exists() else (shutil.which("powershell.exe") or "powershell.exe"))


def build_runner_command(
    action: str,
    *,
    options: JobOptions | None = None,
    cancel_file: Path | None = None,
) -> list[str]:
    """Return a subprocess-safe argument list for a runner action."""

    command = [
        powershell_executable(),
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(RUNNER),
        "-Action",
        action,
    ]
    if action == "Convert":
        if options is None:
            raise ValueError("Convert requires JobOptions")
        command.extend(
            [
                "-InputFile",
                str(options.input_file),
                "-Mode",
                options.mode,
                "-Dpi",
                str(options.dpi),
                "-ExcelRenderMode",
                options.excel_render_mode,
                "-MaxAutoFitColumns",
                str(options.max_autofit_columns),
            ]
        )
        if options.output_directory is not None:
            command.extend(["-OutputDirectory", str(options.output_directory)])
    if cancel_file is not None:
        command.extend(["-CancelFile", str(cancel_file)])
    return command

