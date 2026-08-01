from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from contextpack_gui import (
    EXCEL_MODE_LABELS,
    MODE_LABELS,
    TEXT,
    ContextPackGui,
    JobOptions,
    build_runner_command,
    classify_input,
    make_ai_prompt,
    parse_runner_event,
    preferred_output_directory,
    translate_runner_status,
    validate_options,
)


class GuiCoreTests(unittest.TestCase):
    def test_classifies_supported_inputs(self) -> None:
        self.assertEqual(classify_input(Path("report.pdf")), "pdf")
        self.assertEqual(classify_input(Path("budget.XLSX")), "excel")
        self.assertEqual(classify_input(Path("scan.png")), "image")
        self.assertEqual(classify_input(Path("notes.docx")), "document")

    def test_validates_ocr_scope_and_numeric_bounds(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "notes.docx"
            source.write_bytes(b"test")
            self.assertIn("OCR", validate_options(JobOptions(source, mode="Ocr")) or "")
            self.assertIn("DPI", validate_options(JobOptions(source, dpi=500)) or "")
            self.assertIn("AutoFit", validate_options(JobOptions(source, max_autofit_columns=5)) or "")

    def test_builds_argument_safe_runner_command(self) -> None:
        source = Path(r"C:\Work Files\ანგარიში.xlsx")
        output_directory = Path(r"C:\Destination Folder")
        command = build_runner_command(
            "Convert",
            options=JobOptions(source, mode="Full", dpi=240, excel_render_mode="Workbook", max_autofit_columns=80, output_directory=output_directory),
            cancel_file=Path(r"C:\Temp\cancel token"),
        )
        self.assertIn(str(source), command)
        self.assertEqual(command[command.index("-Dpi") + 1], "240")
        self.assertEqual(command[command.index("-ExcelRenderMode") + 1], "Workbook")
        self.assertEqual(command[command.index("-MaxAutoFitColumns") + 1], "80")
        self.assertEqual(command[command.index("-OutputDirectory") + 1], str(output_directory))

    def test_parses_only_structured_runner_events(self) -> None:
        event = {
            "contextpack_event": 1,
            "type": "result",
            "message": "done",
            "output_path": r"C:\output\package",
        }
        self.assertEqual(parse_runner_event(json.dumps(event)), event)
        self.assertIsNone(parse_runner_event("ordinary tool output"))
        self.assertIsNone(parse_runner_event('{"type":"result"}'))

    def test_ai_prompt_contains_local_result_and_review_order(self) -> None:
        prompt = make_ai_prompt(Path(r"C:\output\package"))
        self.assertIn(r"C:\output\package", prompt)
        self.assertIn("manifest.json", prompt)
        self.assertIn("quality-report.md", prompt)

        english_prompt = make_ai_prompt(Path(r"C:\output\package"), "en")
        self.assertIn("Review this ContextPack result", english_prompt)
        self.assertIn("Goal:", english_prompt)

    def test_localization_catalogs_are_complete_and_keep_internal_mode_values(self) -> None:
        self.assertEqual(set(TEXT["ka"]), set(TEXT["en"]))
        self.assertEqual(set(MODE_LABELS["ka"].values()), set(MODE_LABELS["en"].values()))
        self.assertEqual(set(EXCEL_MODE_LABELS["ka"].values()), set(EXCEL_MODE_LABELS["en"].values()))
        self.assertEqual(translate_runner_status("Package finalized.", "en"), "Package finalized.")
        self.assertEqual(translate_runner_status("Package finalized.", "ka"), "პაკეტი საბოლოოდ შეიქმნა.")

    def test_validation_can_return_english_messages(self) -> None:
        missing = JobOptions(Path("missing.pdf"))
        self.assertEqual(validate_options(missing, "en"), "The selected file no longer exists.")

    def test_preferred_output_directory_exists(self) -> None:
        self.assertTrue(preferred_output_directory().is_dir())

    def test_cancelled_output_picker_does_not_start_conversion(self) -> None:
        gui = Mock(spec=ContextPackGui)
        gui.root = Mock()
        gui.last_output_directory = Path.home()
        gui.status_var = Mock()
        gui._read_options.return_value = JobOptions(Path("report.pdf"))

        with patch("contextpack_gui.filedialog.askdirectory", return_value=""):
            ContextPackGui._start_conversion(gui)

        gui._start_job.assert_not_called()
        gui._set_status.assert_called_once_with("status.output_cancelled")

    def test_selected_output_directory_is_forwarded_to_job(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            destination = Path(temporary).resolve()
            gui = Mock(spec=ContextPackGui)
            gui.root = Mock()
            gui.last_output_directory = Path.home()
            gui.status_var = Mock()
            gui.open_button = Mock()
            gui.copy_button = Mock()
            original = JobOptions(Path("report.pdf"))
            gui._read_options.return_value = original

            with patch("contextpack_gui.filedialog.askdirectory", return_value=str(destination)):
                ContextPackGui._start_conversion(gui)

            forwarded = gui._start_job.call_args.args[1]
            self.assertEqual(forwarded.output_directory, destination)
            self.assertEqual(gui.last_output_directory, destination)

    def test_language_switch_preserves_processing_mode(self) -> None:
        gui = Mock(spec=ContextPackGui)
        gui.language = "ka"
        gui.language_var = Mock()
        gui.mode_var = Mock()
        gui.excel_mode_var = Mock()
        gui.mode_combo = Mock()
        gui.excel_combo = Mock()
        gui.process = None
        gui.language_var.get.return_value = "English"
        gui.mode_var.get.return_value = "სრული პაკეტი — OCR-ის გარეშე"
        gui.excel_mode_var.get.return_value = "Workbook — ზუსტი და მცირე"
        gui.localized_widgets = []

        ContextPackGui._change_language(gui)

        self.assertEqual(gui.language, "en")
        gui.mode_var.set.assert_called_once_with("Full package — without OCR")
        gui.excel_mode_var.set.assert_called_once_with("Workbook — exact and compact")


if __name__ == "__main__":
    unittest.main()
