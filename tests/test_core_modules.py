from __future__ import annotations

import tempfile
import unittest
import queue
from pathlib import Path
from unittest.mock import Mock, patch

from contextpack.file_types import classify_input
from contextpack.job_options import JobOptions
from contextpack.job_controller import launch_background_job, new_cancel_file, request_cancellation
from contextpack.validation import validation_error_key


class CoreModuleTests(unittest.TestCase):
    """Protect logic that no longer belongs to the GUI window."""

    def test_file_type_classification_is_case_insensitive(self) -> None:
        self.assertEqual(classify_input(Path("report.PDF")), "pdf")
        self.assertEqual(classify_input(Path("budget.XLSM")), "excel")
        self.assertEqual(classify_input(Path("scan.WEBP")), "image")
        self.assertEqual(classify_input(Path("notes.docx")), "document")

    def test_validation_returns_language_independent_error_keys(self) -> None:
        self.assertEqual(
            validation_error_key(JobOptions(Path("missing.pdf"))),
            "validation.file_missing",
        )

        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "notes.docx"
            source.write_bytes(b"test")
            self.assertEqual(
                validation_error_key(JobOptions(source, mode="Ocr")),
                "validation.ocr_scope",
            )

    def test_cancellation_uses_an_explicit_unique_token_file(self) -> None:
        first = new_cancel_file()
        second = new_cancel_file()
        self.assertNotEqual(first, second)
        self.assertTrue(first.name.startswith("contextpack-gui-"))
        self.assertEqual(first.suffix, ".cancel")

        with tempfile.TemporaryDirectory() as temporary:
            cancel_file = Path(temporary) / "job.cancel"
            request_cancellation(cancel_file)
            request_cancellation(cancel_file)
            self.assertEqual(cancel_file.read_text(encoding="utf-8"), "cancel")

    def test_background_job_forwards_output_and_completion(self) -> None:
        process = Mock()
        process.stdout = ["first line\n", "second line\r\n"]
        process.wait.return_value = 0
        events: queue.Queue[tuple[str, object]] = queue.Queue()

        with patch("contextpack.job_controller.subprocess.Popen", return_value=process):
            launched_process, worker = launch_background_job(
                ["runner", "--example"],
                working_directory=Path.cwd(),
                events=events,
            )
        worker.join(timeout=2)

        self.assertIs(launched_process, process)
        self.assertFalse(worker.is_alive())
        self.assertEqual(events.get_nowait(), ("line", "first line"))
        self.assertEqual(events.get_nowait(), ("line", "second line"))
        self.assertEqual(events.get_nowait(), ("done", 0))


if __name__ == "__main__":
    unittest.main()
