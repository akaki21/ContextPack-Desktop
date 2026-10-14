"""Run the PowerShell worker without blocking the Tkinter event loop."""

from __future__ import annotations

import queue
import subprocess
import tempfile
import threading
import uuid
from pathlib import Path


JobEventQueue = queue.Queue[tuple[str, object]]


def new_cancel_file() -> Path:
    """Return a unique token path used for cooperative cancellation."""

    return Path(tempfile.gettempdir()) / f"contextpack-gui-{uuid.uuid4().hex}.cancel"


def _forward_process_output(process: subprocess.Popen[str], events: JobEventQueue) -> None:
    """Forward worker output and its final exit code to the GUI event queue."""

    assert process.stdout is not None
    for line in process.stdout:
        events.put(("line", line.rstrip("\r\n")))
    events.put(("done", process.wait()))


def launch_background_job(
    command: list[str],
    *,
    working_directory: Path,
    events: JobEventQueue,
) -> tuple[subprocess.Popen[str], threading.Thread]:
    """Start a hidden worker process and a daemon thread that reads its output."""

    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    process = subprocess.Popen(
        command,
        cwd=working_directory,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
        creationflags=creation_flags,
    )
    worker = threading.Thread(target=_forward_process_output, args=(process, events), daemon=True)
    worker.start()
    return process, worker


def request_cancellation(cancel_file: Path) -> None:
    """Ask the runner to stop at its next safe operation boundary."""

    if not cancel_file.exists():
        cancel_file.write_text("cancel", encoding="utf-8")

