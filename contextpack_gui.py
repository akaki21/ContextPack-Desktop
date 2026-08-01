from __future__ import annotations

import importlib.util
import json
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import uuid
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


ROOT = Path(__file__).resolve().parent
RUNNER = ROOT / "contextpack-gui-runner.ps1"
OUTPUT_ROOT = ROOT / "output"
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}

LANGUAGE_LABELS = {"ქართული": "ka", "English": "en"}
MODE_LABELS = {
    "ka": {
        "ავტომატური — რეკომენდებული": "Auto",
        "სწრაფი — მხოლოდ ტექსტი": "Fast",
        "სრული პაკეტი — OCR-ის გარეშე": "Full",
        "OCR — სკანირებული PDF": "Ocr",
    },
    "en": {
        "Automatic — recommended": "Auto",
        "Fast — text only": "Fast",
        "Full package — without OCR": "Full",
        "OCR — scanned PDF": "Ocr",
    },
}
EXCEL_MODE_LABELS = {
    "ka": {
        "ორივე ხედი — რეკომენდებული": "Both",
        "Workbook — ზუსტი და მცირე": "Workbook",
        "AutoFit — მხოლოდ დამხმარე ხედი": "AutoFit",
    },
    "en": {
        "Both views — recommended": "Both",
        "Workbook — exact and compact": "Workbook",
        "AutoFit — helper view only": "AutoFit",
    },
}

TEXT = {
    "ka": {
        "subtitle": "ფაილების მომზადება AI-სთვის — მარტივად და ლოკალურად",
        "file.section": "1. აირჩიე ფაილი",
        "file.hint": "PDF, Excel, სურათი ან Office დოკუმენტი",
        "file.browse": "არჩევა…",
        "options.section": "2. აირჩიე დამუშავების რეჟიმი",
        "options.hint": "თუ არ ხარ დარწმუნებული, დატოვე რეკომენდებული მნიშვნელობები",
        "options.mode": "ძირითადი რეჟიმი",
        "options.excel": "Excel-ის ხედები",
        "options.columns": "AutoFit სვეტების ზღვარი",
        "action.hint": "3. გაუშვი და დაელოდე შედეგს",
        "action.start": "დამუშავების დაწყება",
        "action.cancel": "გაუქმება",
        "details.title": "მიმდინარეობა და დეტალები",
        "details.check": "გარემოს შემოწმება",
        "details.setup": "Setup / Repair",
        "result.open": "შედეგის გახსნა",
        "result.copy": "AI დავალების კოპირება",
        "result.local": "ყველა დამუშავება სრულდება შენს კომპიუტერზე",
        "environment.engine": "ძრავა",
        "status.ready": "ფაილის ასარჩევად მზადაა",
        "status.output_cancelled": "შენახვის ადგილის არჩევა გაუქმდა",
        "status.starting": "დავალება იწყება…",
        "status.success": "დამუშავება წარმატებით დასრულდა",
        "status.environment_success": "გარემოს მოქმედება წარმატებით დასრულდა",
        "status.cancelled": "დავალება გაუქმდა",
        "status.failed": "დამუშავება შეცდომით დასრულდა",
        "status.cancel_requested": "გაუქმების მოთხოვნა მიღებულია — მიმდინარე უსაფრთხო ეტაპს ველოდებით…",
        "status.prompt_copied": "AI-სთვის დავალების ტექსტი დაკოპირდა",
        "file.none": "ჯერ ფაილი არ არის არჩეული",
        "file.missing": "ფაილი ვერ მოიძებნა",
        "file.detected": "ამოცნობილია: {kind} • {name}",
        "kind.pdf": "PDF დოკუმენტი",
        "kind.excel": "Excel workbook",
        "kind.image": "სურათი — OCR",
        "kind.document": "დოკუმენტი / მონაცემთა ფაილი",
        "dialog.choose_file": "აირჩიე დასამუშავებელი ფაილი",
        "dialog.all_supported": "ყველა მხარდაჭერილი ფაილი",
        "dialog.images": "სურათები",
        "dialog.all_files": "ყველა ფაილი",
        "dialog.choose_output": "სად შეინახოს ContextPack-ის შედეგი?",
        "dialog.no_file.title": "ფაილი არ არის არჩეული",
        "dialog.no_file.body": "ჯერ აირჩიე დასამუშავებელი ფაილი.",
        "dialog.invalid.title": "არასწორი პარამეტრი",
        "dialog.invalid.body": "DPI და სვეტების ზღვარი უნდა იყოს მთელი რიცხვი.",
        "dialog.cannot_start": "დავალება ვერ იწყება",
        "dialog.no_excel.title": "Excel ვერ მოიძებნა",
        "dialog.no_excel.body": "სრული Excel პაკეტის ვიზუალური რენდერისთვის Microsoft Excel საჭიროა. გაუშვი Environment Check ან Setup / Repair.",
        "dialog.no_ocr.title": "OCR გარემო მზად არაა",
        "dialog.no_ocr.body": "OCR-ისთვის საჭირო კომპონენტი ან ენის მოდელი აკლია. გაუშვი Setup / Repair.",
        "dialog.setup.body": "ეს მოქმედება შეამოწმებს და საჭიროების შემთხვევაში დააინსტალირებს პროექტის დამოკიდებულებებს. გაგრძელდეს?",
        "dialog.busy.title": "დავალება მიმდინარეობს",
        "dialog.busy.body": "ჯერ მიმდინარე დავალების დასრულებას დაელოდე.",
        "dialog.launch_failed": "გაშვება ვერ მოხერხდა",
        "dialog.ready.title": "მზადაა",
        "dialog.ready.body": "ContextPack-ის შედეგი მზადაა.\n\n{path}",
        "dialog.error.title": "შეცდომა",
        "dialog.error.body": "დავალება ვერ დასრულდა. დეტალები იხილე ქვედა ჟურნალში.",
        "dialog.result_missing": "შედეგი ვერ მოიძებნა",
        "dialog.close.title": "დავალება ჯერ მიმდინარეობს",
        "dialog.close.body": "გაუქმდეს მიმდინარე დავალება და პროგრამა უსაფრთხო გაჩერების შემდეგ დაიხუროს?",
        "validation.file_missing": "არჩეული ფაილი აღარ არსებობს.",
        "validation.mode": "არჩეული რეჟიმი უცნობია.",
        "validation.dpi": "DPI უნდა იყოს 96-დან 300-მდე.",
        "validation.excel_mode": "Excel-ის ხედის რეჟიმი უცნობია.",
        "validation.columns": "AutoFit-ის სვეტების ზღვარი უნდა იყოს 10-დან 200-მდე.",
        "validation.ocr_scope": "OCR რეჟიმი გამოიყენება მხოლოდ PDF-სა და სურათებზე.",
        "validation.output_missing": "შენახვისთვის არჩეული საქაღალდე აღარ არსებობს.",
        "log.selected_file": "არჩეული ფაილი: {path}",
        "log.output": "შენახვის ადგილი: {path}",
        "log.mode": "რეჟიმი: {mode} | DPI: {dpi} | Excel: {excel}",
        "log.error": "შეცდომა: {message}",
        "log.cancelled": "დავალება უსაფრთხოდ გაუქმდა.",
        "log.cancel_requested": "Cancel მოთხოვნილია. მიმდინარე ოპერაცია უსაფრთხოდ დასრულდება და შემდეგ დავალება გაჩერდება.",
        "prompt.line1": "გაეცანი ContextPack-ის ამ შედეგს: {path}",
        "prompt.line2": "ჯერ წაიკითხე manifest.json და quality-report.md, შემდეგ გახსენი მხოლოდ დავალებისთვის საჭირო ფაილები.",
        "prompt.goal": "მიზანი: [აქ ჩაწერე რა უნდა გაკეთდეს]",
    },
    "en": {
        "subtitle": "Prepare files for AI — simply and locally",
        "file.section": "1. Choose a file",
        "file.hint": "PDF, Excel, image, or Office document",
        "file.browse": "Browse…",
        "options.section": "2. Choose a processing mode",
        "options.hint": "Keep the recommended settings unless you need something specific",
        "options.mode": "Main mode",
        "options.excel": "Excel views",
        "options.columns": "AutoFit column limit",
        "action.hint": "3. Start and wait for the result",
        "action.start": "Start processing",
        "action.cancel": "Cancel",
        "details.title": "Progress and details",
        "details.check": "Check environment",
        "details.setup": "Setup / Repair",
        "result.open": "Open result",
        "result.copy": "Copy AI task",
        "result.local": "All processing stays on your computer",
        "environment.engine": "Engine",
        "status.ready": "Ready to choose a file",
        "status.output_cancelled": "Output folder selection was cancelled",
        "status.starting": "Starting the job…",
        "status.success": "Processing completed successfully",
        "status.environment_success": "Environment action completed successfully",
        "status.cancelled": "Job cancelled",
        "status.failed": "Processing failed",
        "status.cancel_requested": "Cancellation requested — waiting for the current safe boundary…",
        "status.prompt_copied": "AI task text copied",
        "file.none": "No file selected yet",
        "file.missing": "File not found",
        "file.detected": "Detected: {kind} • {name}",
        "kind.pdf": "PDF document",
        "kind.excel": "Excel workbook",
        "kind.image": "Image — OCR",
        "kind.document": "Document / data file",
        "dialog.choose_file": "Choose a file to process",
        "dialog.all_supported": "All supported files",
        "dialog.images": "Images",
        "dialog.all_files": "All files",
        "dialog.choose_output": "Where should ContextPack save the result?",
        "dialog.no_file.title": "No file selected",
        "dialog.no_file.body": "Choose a file to process first.",
        "dialog.invalid.title": "Invalid setting",
        "dialog.invalid.body": "DPI and the column limit must be whole numbers.",
        "dialog.cannot_start": "Cannot start the job",
        "dialog.no_excel.title": "Excel not found",
        "dialog.no_excel.body": "Microsoft Excel is required to render the complete visual Excel package. Run Check environment or Setup / Repair.",
        "dialog.no_ocr.title": "OCR environment is not ready",
        "dialog.no_ocr.body": "An OCR component or language model is missing. Run Setup / Repair.",
        "dialog.setup.body": "This checks the project dependencies and installs anything missing. Continue?",
        "dialog.busy.title": "A job is running",
        "dialog.busy.body": "Wait for the current job to finish first.",
        "dialog.launch_failed": "Could not start",
        "dialog.ready.title": "Ready",
        "dialog.ready.body": "Your ContextPack result is ready.\n\n{path}",
        "dialog.error.title": "Error",
        "dialog.error.body": "The job could not be completed. See the log below for details.",
        "dialog.result_missing": "Result not found",
        "dialog.close.title": "A job is still running",
        "dialog.close.body": "Cancel the current job and close after it stops safely?",
        "validation.file_missing": "The selected file no longer exists.",
        "validation.mode": "The selected mode is unknown.",
        "validation.dpi": "DPI must be between 96 and 300.",
        "validation.excel_mode": "The Excel view mode is unknown.",
        "validation.columns": "The AutoFit column limit must be between 10 and 200.",
        "validation.ocr_scope": "OCR mode can only be used with PDFs and images.",
        "validation.output_missing": "The selected output folder no longer exists.",
        "log.selected_file": "Selected file: {path}",
        "log.output": "Output folder: {path}",
        "log.mode": "Mode: {mode} | DPI: {dpi} | Excel: {excel}",
        "log.error": "Error: {message}",
        "log.cancelled": "The job was cancelled safely.",
        "log.cancel_requested": "Cancellation requested. The current operation will finish safely, then the job will stop.",
        "prompt.line1": "Review this ContextPack result: {path}",
        "prompt.line2": "Read manifest.json and quality-report.md first, then open only the files needed for the task.",
        "prompt.goal": "Goal: [describe what should be done]",
    },
}

STATUS_TRANSLATIONS = {
    "Checking the environment...": {"ka": "გარემო მოწმდება…", "en": "Checking the environment…"},
    "Environment check passed.": {"ka": "გარემო გამართულია.", "en": "Environment check passed."},
    "Installing or repairing dependencies...": {"ka": "კომპონენტები ყენდება ან ახლდება…", "en": "Installing or repairing dependencies…"},
    "Verifying the environment...": {"ka": "გარემო საბოლოოდ მოწმდება…", "en": "Verifying the environment…"},
    "Setup and environment check completed.": {"ka": "Setup დასრულდა და გარემო გამართულია.", "en": "Setup and environment check completed."},
    "Validating options and preparing the job...": {"ka": "პარამეტრები მოწმდება და დავალება მზადდება…", "en": "Validating options and preparing the job…"},
    "ContextPack is processing the file...": {"ka": "ContextPack ფაილს ამუშავებს…", "en": "ContextPack is processing the file…"},
    "PDF inspection finished.": {"ka": "PDF-ის შემოწმება დასრულდა.", "en": "PDF inspection finished."},
    "Workbook data extracted.": {"ka": "Excel-ის მონაცემები ამოღებულია.", "en": "Workbook data extracted."},
    "Visual pages rendered.": {"ka": "ვიზუალური გვერდები შექმნილია.", "en": "Visual pages rendered."},
    "Package finalized.": {"ka": "პაკეტი საბოლოოდ შეიქმნა.", "en": "Package finalized."},
    "Output finalized.": {"ka": "შედეგი საბოლოოდ შეიქმნა.", "en": "Output finalized."},
    "The job completed successfully.": {"ka": "დამუშავება წარმატებით დასრულდა.", "en": "The job completed successfully."},
    "The job was cancelled safely.": {"ka": "დავალება უსაფრთხოდ გაუქმდა.", "en": "The job was cancelled safely."},
}

COLORS = {
    "background": "#0b1220",
    "card": "#162033",
    "card_alt": "#111a2b",
    "border": "#2b3a55",
    "text": "#f8fafc",
    "muted": "#9aa9bf",
    "accent": "#38bdf8",
    "accent_active": "#0ea5e9",
    "success": "#34d399",
    "warning": "#fbbf24",
    "danger": "#fb7185",
}


@dataclass(frozen=True)
class JobOptions:
    input_file: Path
    mode: str = "Auto"
    dpi: int = 180
    excel_render_mode: str = "Both"
    max_autofit_columns: int = 60
    output_directory: Path | None = None


def powershell_executable() -> str:
    system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    candidate = system_root / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
    return str(candidate if candidate.exists() else (shutil.which("powershell.exe") or "powershell.exe"))


def classify_input(path: Path) -> str:
    extension = path.suffix.lower()
    if extension == ".pdf":
        return "pdf"
    if extension in EXCEL_EXTENSIONS:
        return "excel"
    if extension in IMAGE_EXTENSIONS:
        return "image"
    return "document"


def translate(language: str, key: str, **values: Any) -> str:
    language = language if language in TEXT else "ka"
    template = TEXT[language].get(key, TEXT["en"].get(key, key))
    return template.format(**values)


def label_for_value(labels: dict[str, dict[str, str]], language: str, value: str) -> str:
    return next(label for label, code in labels[language].items() if code == value)


def translate_runner_status(message: str, language: str) -> str:
    return STATUS_TRANSLATIONS.get(message, {}).get(language, message)


def validate_options(options: JobOptions, language: str = "ka") -> str | None:
    if not options.input_file.is_file():
        return translate(language, "validation.file_missing")
    if options.mode not in {"Auto", "Fast", "Full", "Ocr"}:
        return translate(language, "validation.mode")
    if not 96 <= options.dpi <= 300:
        return translate(language, "validation.dpi")
    if options.excel_render_mode not in {"Workbook", "AutoFit", "Both"}:
        return translate(language, "validation.excel_mode")
    if not 10 <= options.max_autofit_columns <= 200:
        return translate(language, "validation.columns")
    if options.mode == "Ocr" and classify_input(options.input_file) not in {"pdf", "image"}:
        return translate(language, "validation.ocr_scope")
    if options.output_directory is not None and not options.output_directory.is_dir():
        return translate(language, "validation.output_missing")
    return None


def build_runner_command(
    action: str,
    *,
    options: JobOptions | None = None,
    cancel_file: Path | None = None,
) -> list[str]:
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


def parse_runner_event(line: str) -> dict[str, Any] | None:
    try:
        payload = json.loads(line)
    except json.JSONDecodeError:
        return None
    if not isinstance(payload, dict) or payload.get("contextpack_event") != 1:
        return None
    return payload


def make_ai_prompt(output_path: Path, language: str = "ka") -> str:
    return "\n".join(
        (
            translate(language, "prompt.line1", path=output_path),
            translate(language, "prompt.line2"),
            translate(language, "prompt.goal"),
        )
    )


def preferred_output_directory() -> Path:
    home = Path.home()
    candidates = [
        Path(os.environ.get("OneDrive", "")) / "Desktop" if os.environ.get("OneDrive") else None,
        home / "OneDrive" / "Desktop",
        home / "Desktop",
    ]
    return next((candidate for candidate in candidates if candidate is not None and candidate.is_dir()), home)


def find_tesseract() -> Path | None:
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
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
            return True
    except (ImportError, FileNotFoundError, OSError):
        return False


def environment_summary() -> dict[str, bool]:
    models_ready = all((ROOT / "tessdata" / f"{language}.traineddata").is_file() for language in ("kat", "eng", "osd"))
    return {
        "engine": RUNNER.is_file() and importlib.util.find_spec("markitdown") is not None,
        "ocr": find_tesseract() is not None and models_ready and importlib.util.find_spec("ocrmypdf") is not None,
        "excel": excel_is_available(),
    }


class ContextPackGui:
    def __init__(self, root: tk.Tk, initial_file: Path | None = None) -> None:
        self.root = root
        self.events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.process: subprocess.Popen[str] | None = None
        self.worker: threading.Thread | None = None
        self.cancel_file: Path | None = None
        self.result_path: Path | None = None
        self.current_action = ""
        self.close_when_done = False
        self.environment = environment_summary()
        self.last_output_directory = preferred_output_directory()
        self.language = "ka"
        self.localized_widgets: list[tuple[tk.Widget, str]] = []
        self.status_token: tuple[str, str] = ("key", "status.ready")

        self.language_var = tk.StringVar(value="ქართული")
        self.file_var = tk.StringVar(value=str(initial_file) if initial_file else "")
        self.mode_var = tk.StringVar(value=next(iter(MODE_LABELS[self.language])))
        self.excel_mode_var = tk.StringVar(value=next(iter(EXCEL_MODE_LABELS[self.language])))
        self.dpi_var = tk.IntVar(value=180)
        self.max_columns_var = tk.IntVar(value=60)
        self.status_var = tk.StringVar()
        self.file_kind_var = tk.StringVar()
        self.environment_var = tk.StringVar()

        self._configure_window()
        self._configure_styles()
        self._build_interface()
        self._refresh_status()
        self._refresh_environment_label()
        self._on_file_changed()
        if self.process is not None:
            self.excel_combo.configure(state="disabled")
            self.columns_spin.configure(state="disabled")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._poll_events)

    def _configure_window(self) -> None:
        self.root.title("ContextPack Desktop — Local GUI")
        self.root.geometry("1080x760")
        self.root.minsize(880, 650)
        self.root.configure(bg=COLORS["background"])
        if os.name == "nt":
            self.root.state("zoomed")
        try:
            self.root.tk.call("tk", "scaling", 1.08)
        except tk.TclError:
            pass

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background=COLORS["background"])
        style.configure("Card.TFrame", background=COLORS["card"])
        style.configure("AltCard.TFrame", background=COLORS["card_alt"])
        style.configure("TLabel", background=COLORS["background"], foreground=COLORS["text"], font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI Semibold", 24), foreground=COLORS["text"])
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground=COLORS["muted"])
        style.configure("Section.TLabel", background=COLORS["card"], font=("Segoe UI Semibold", 12), foreground=COLORS["text"])
        style.configure("Card.TLabel", background=COLORS["card"], foreground=COLORS["text"])
        style.configure("Hint.TLabel", background=COLORS["card"], foreground=COLORS["muted"], font=("Segoe UI", 9))
        style.configure("Status.TLabel", background=COLORS["card_alt"], foreground=COLORS["muted"], font=("Segoe UI", 9))
        style.configure("TButton", font=("Segoe UI Semibold", 10), padding=(12, 8))
        style.configure("Accent.TButton", background=COLORS["accent"], foreground="#062033", borderwidth=0, padding=(18, 10))
        style.map("Accent.TButton", background=[("active", COLORS["accent_active"]), ("disabled", COLORS["border"])])
        style.configure("Secondary.TButton", background=COLORS["border"], foreground=COLORS["text"], borderwidth=0)
        style.map("Secondary.TButton", background=[("active", "#3b4d6b")])
        style.configure("Danger.TButton", background=COLORS["danger"], foreground="#3f0a18", borderwidth=0)
        style.configure("TCombobox", fieldbackground=COLORS["card_alt"], background=COLORS["border"], foreground=COLORS["text"], arrowcolor=COLORS["text"], padding=7)
        style.map("TCombobox", fieldbackground=[("readonly", COLORS["card_alt"]), ("disabled", COLORS["card"])], foreground=[("readonly", COLORS["text"]), ("disabled", COLORS["muted"])])
        style.configure("TSpinbox", fieldbackground=COLORS["card_alt"], foreground=COLORS["text"], arrowcolor=COLORS["text"], padding=7)
        style.configure("Horizontal.TProgressbar", troughcolor=COLORS["card_alt"], background=COLORS["accent"], bordercolor=COLORS["card_alt"], lightcolor=COLORS["accent"], darkcolor=COLORS["accent"])

    def _t(self, key: str, **values: Any) -> str:
        return translate(self.language, key, **values)

    def _localized_label(self, parent: tk.Widget, key: str, **options: Any) -> ttk.Label:
        widget = ttk.Label(parent, text=self._t(key), **options)
        self.localized_widgets.append((widget, key))
        return widget

    def _localized_button(self, parent: tk.Widget, key: str, **options: Any) -> ttk.Button:
        widget = ttk.Button(parent, text=self._t(key), **options)
        self.localized_widgets.append((widget, key))
        return widget

    def _set_status(self, key: str) -> None:
        self.status_token = ("key", key)
        self._refresh_status()

    def _set_runner_status(self, message: str) -> None:
        self.status_token = ("runner", message)
        self._refresh_status()

    def _refresh_status(self) -> None:
        token_type, value = self.status_token
        rendered = self._t(value) if token_type == "key" else translate_runner_status(value, self.language)
        self.status_var.set(rendered)

    def _change_language(self, _event: tk.Event[Any] | None = None) -> None:
        new_language = LANGUAGE_LABELS.get(self.language_var.get(), "ka")
        if new_language == self.language:
            return
        mode_code = MODE_LABELS[self.language].get(self.mode_var.get(), "Auto")
        excel_code = EXCEL_MODE_LABELS[self.language].get(self.excel_mode_var.get(), "Both")
        self.language = new_language
        self.mode_combo.configure(values=list(MODE_LABELS[self.language]))
        self.excel_combo.configure(values=list(EXCEL_MODE_LABELS[self.language]))
        self.mode_var.set(label_for_value(MODE_LABELS, self.language, mode_code))
        self.excel_mode_var.set(label_for_value(EXCEL_MODE_LABELS, self.language, excel_code))
        for widget, key in self.localized_widgets:
            widget.configure(text=self._t(key))
        self._refresh_status()
        self._refresh_environment_label()
        self._on_file_changed()

    def _build_interface(self) -> None:
        shell = ttk.Frame(self.root, padding=(24, 18))
        shell.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(4, weight=1)

        header = ttk.Frame(shell)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="ContextPack Desktop", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self._localized_label(header, "subtitle", style="Subtitle.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.language_combo = ttk.Combobox(header, state="readonly", values=list(LANGUAGE_LABELS), textvariable=self.language_var, width=11)
        self.language_combo.grid(row=0, column=2, sticky="e")
        self.language_combo.bind("<<ComboboxSelected>>", self._change_language)
        self.environment_label = ttk.Label(header, textvariable=self.environment_var, style="Subtitle.TLabel")
        self.environment_label.grid(row=1, column=1, columnspan=2, sticky="e")

        file_card = ttk.Frame(shell, style="Card.TFrame", padding=15)
        file_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        file_card.columnconfigure(0, weight=1)
        self._localized_label(file_card, "file.section", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        self._localized_label(file_card, "file.hint", style="Hint.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 10))

        entry_frame = ttk.Frame(file_card, style="Card.TFrame")
        entry_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
        entry_frame.columnconfigure(0, weight=1)
        self.file_entry = tk.Entry(
            entry_frame,
            textvariable=self.file_var,
            bg=COLORS["card_alt"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            font=("Segoe UI", 10),
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
        )
        self.file_entry.grid(row=0, column=0, sticky="ew", ipady=9)
        self.file_entry.bind("<FocusOut>", lambda _event: self._on_file_changed())
        self._localized_button(entry_frame, "file.browse", style="Secondary.TButton", command=self._choose_file).grid(row=0, column=1, padx=(10, 0))
        ttk.Label(file_card, textvariable=self.file_kind_var, style="Hint.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 0))

        options_card = ttk.Frame(shell, style="Card.TFrame", padding=15)
        options_card.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        options_card.columnconfigure(0, weight=1)
        options_card.columnconfigure(1, weight=1)
        self._localized_label(options_card, "options.section", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        self._localized_label(options_card, "options.hint", style="Hint.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 12))

        self._localized_label(options_card, "options.mode", style="Card.TLabel").grid(row=2, column=0, sticky="w")
        self.mode_combo = ttk.Combobox(options_card, state="readonly", values=list(MODE_LABELS[self.language]), textvariable=self.mode_var, width=27)
        self.mode_combo.grid(row=3, column=0, sticky="ew", padx=(0, 10), pady=(4, 0))

        self._localized_label(options_card, "options.excel", style="Card.TLabel").grid(row=2, column=1, sticky="w")
        self.excel_combo = ttk.Combobox(options_card, state="disabled", values=list(EXCEL_MODE_LABELS[self.language]), textvariable=self.excel_mode_var, width=27)
        self.excel_combo.grid(row=3, column=1, sticky="ew", pady=(4, 0))

        ttk.Label(options_card, text="DPI", style="Card.TLabel").grid(row=4, column=0, sticky="w", pady=(12, 0))
        self.dpi_spin = ttk.Spinbox(options_card, from_=96, to=300, increment=12, textvariable=self.dpi_var, width=8)
        self.dpi_spin.grid(row=5, column=0, sticky="ew", padx=(0, 10), pady=(4, 0))

        self._localized_label(options_card, "options.columns", style="Card.TLabel").grid(row=4, column=1, sticky="w", pady=(12, 0))
        self.columns_spin = ttk.Spinbox(options_card, from_=10, to=200, increment=5, textvariable=self.max_columns_var, width=8, state="disabled")
        self.columns_spin.grid(row=5, column=1, sticky="ew", pady=(4, 0))

        action_card = ttk.Frame(shell, style="AltCard.TFrame", padding=15)
        action_card.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        action_card.columnconfigure(1, weight=1)
        self._localized_label(action_card, "action.hint", style="Status.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        self.run_button = self._localized_button(action_card, "action.start", style="Accent.TButton", command=self._start_conversion)
        self.run_button.grid(row=1, column=0, rowspan=2, sticky="nsw", pady=(10, 0))
        self.status_label = ttk.Label(action_card, textvariable=self.status_var, style="Status.TLabel")
        self.status_label.grid(row=1, column=1, sticky="sw", padx=(16, 12), pady=(10, 4))
        self.progress = ttk.Progressbar(action_card, mode="determinate", maximum=100)
        self.progress.grid(row=2, column=1, sticky="ew", padx=(16, 12), pady=(0, 1))
        self.cancel_button = self._localized_button(action_card, "action.cancel", style="Danger.TButton", command=self._request_cancel, state="disabled")
        self.cancel_button.grid(row=1, column=2, rowspan=2, sticky="nse", pady=(10, 0))

        details = ttk.Frame(shell, style="Card.TFrame", padding=14)
        details.grid(row=4, column=0, sticky="nsew")
        details.columnconfigure(0, weight=1)
        details.rowconfigure(1, weight=1)
        toolbar = ttk.Frame(details, style="Card.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self._localized_label(toolbar, "details.title", style="Section.TLabel").pack(side="left")
        self._localized_button(toolbar, "details.check", style="Secondary.TButton", command=lambda: self._start_utility("Check")).pack(side="right")
        self._localized_button(toolbar, "details.setup", style="Secondary.TButton", command=self._confirm_setup).pack(side="right", padx=(0, 8))

        log_frame = tk.Frame(details, bg=COLORS["card"])
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log = tk.Text(
            log_frame,
            bg="#0a1020",
            fg="#cbd5e1",
            insertbackground=COLORS["text"],
            relief="flat",
            font=("Cascadia Mono", 9),
            wrap="word",
            height=5,
            padx=12,
            pady=10,
            state="disabled",
        )
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)

        result_bar = ttk.Frame(details, style="Card.TFrame")
        result_bar.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.open_button = self._localized_button(result_bar, "result.open", style="Secondary.TButton", command=self._open_result, state="disabled")
        self.open_button.pack(side="left")
        self.copy_button = self._localized_button(result_bar, "result.copy", style="Secondary.TButton", command=self._copy_ai_prompt, state="disabled")
        self.copy_button.pack(side="left", padx=(8, 0))
        self._localized_label(result_bar, "result.local", style="Hint.TLabel").pack(side="right")

    def _refresh_environment_label(self) -> None:
        def mark(value: bool) -> str:
            return "✓" if value else "!"

        self.environment_var.set(
            f"{self._t('environment.engine')} {mark(self.environment['engine'])}   OCR {mark(self.environment['ocr'])}   Excel {mark(self.environment['excel'])}"
        )

    def _choose_file(self) -> None:
        selected = filedialog.askopenfilename(
            title=self._t("dialog.choose_file"),
            filetypes=[
                (self._t("dialog.all_supported"), "*.pdf *.xlsx *.xlsm *.xltx *.xltm *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp *.docx *.pptx *.csv *.txt"),
                ("PDF", "*.pdf"),
                ("Excel", "*.xlsx *.xlsm *.xltx *.xltm"),
                (self._t("dialog.images"), "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp"),
                (self._t("dialog.all_files"), "*.*"),
            ],
        )
        if selected:
            self.file_var.set(selected)
            self._on_file_changed()

    def _on_file_changed(self) -> None:
        path = Path(self.file_var.get().strip().strip('"')) if self.file_var.get().strip() else None
        is_excel = bool(path and classify_input(path) == "excel")
        self.excel_combo.configure(state="readonly" if is_excel else "disabled")
        self.columns_spin.configure(state="normal" if is_excel else "disabled")
        if not path:
            self.file_kind_var.set(self._t("file.none"))
        elif not path.exists():
            self.file_kind_var.set(self._t("file.missing"))
        else:
            kind = classify_input(path)
            self.file_kind_var.set(self._t("file.detected", kind=self._t(f"kind.{kind}"), name=path.name))

    def _read_options(self) -> JobOptions | None:
        raw_path = self.file_var.get().strip().strip('"')
        if not raw_path:
            messagebox.showwarning(self._t("dialog.no_file.title"), self._t("dialog.no_file.body"))
            return None
        try:
            dpi = int(self.dpi_var.get())
            max_columns = int(self.max_columns_var.get())
        except (tk.TclError, ValueError):
            messagebox.showerror(self._t("dialog.invalid.title"), self._t("dialog.invalid.body"))
            return None
        options = JobOptions(
            input_file=Path(raw_path).resolve(),
            mode=MODE_LABELS[self.language][self.mode_var.get()],
            dpi=dpi,
            excel_render_mode=EXCEL_MODE_LABELS[self.language][self.excel_mode_var.get()],
            max_autofit_columns=max_columns,
        )
        error = validate_options(options, self.language)
        if error:
            messagebox.showerror(self._t("dialog.cannot_start"), error)
            return None
        kind = classify_input(options.input_file)
        if kind == "excel" and options.mode != "Fast" and not self.environment["excel"]:
            messagebox.showerror(self._t("dialog.no_excel.title"), self._t("dialog.no_excel.body"))
            return None
        if (options.mode == "Ocr" or kind == "image") and not self.environment["ocr"]:
            messagebox.showerror(self._t("dialog.no_ocr.title"), self._t("dialog.no_ocr.body"))
            return None
        return options

    def _start_conversion(self) -> None:
        options = self._read_options()
        if options is None:
            return
        selected_output = filedialog.askdirectory(
            parent=self.root,
            title=self._t("dialog.choose_output"),
            initialdir=str(self.last_output_directory),
            mustexist=True,
        )
        if not selected_output:
            self._set_status("status.output_cancelled")
            return
        self.last_output_directory = Path(selected_output).resolve()
        options = replace(options, output_directory=self.last_output_directory)
        self.result_path = None
        self.open_button.configure(state="disabled")
        self.copy_button.configure(state="disabled")
        self._clear_log()
        self._append_log(self._t("log.selected_file", path=options.input_file))
        self._append_log(self._t("log.output", path=options.output_directory))
        self._append_log(self._t("log.mode", mode=options.mode, dpi=options.dpi, excel=options.excel_render_mode))
        self._start_job("Convert", options)

    def _confirm_setup(self) -> None:
        if self.process is not None:
            return
        approved = messagebox.askyesno(
            self._t("details.setup"),
            self._t("dialog.setup.body"),
        )
        if approved:
            self._clear_log()
            self._start_utility("Setup")

    def _start_utility(self, action: str) -> None:
        if self.process is not None:
            messagebox.showinfo(self._t("dialog.busy.title"), self._t("dialog.busy.body"))
            return
        self.result_path = None
        self.open_button.configure(state="disabled")
        self.copy_button.configure(state="disabled")
        self._start_job(action, None)

    def _start_job(self, action: str, options: JobOptions | None) -> None:
        self.current_action = action
        self.cancel_file = Path(tempfile.gettempdir()) / f"contextpack-gui-{uuid.uuid4().hex}.cancel"
        command = build_runner_command(action, options=options, cancel_file=self.cancel_file)
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        try:
            self.process = subprocess.Popen(
                command,
                cwd=ROOT,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=creation_flags,
            )
        except OSError as error:
            self.process = None
            messagebox.showerror(self._t("dialog.launch_failed"), str(error))
            return

        self._set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self._set_status("status.starting")
        self.worker = threading.Thread(target=self._read_process_output, daemon=True)
        self.worker.start()

    def _read_process_output(self) -> None:
        assert self.process is not None and self.process.stdout is not None
        for line in self.process.stdout:
            self.events.put(("line", line.rstrip("\r\n")))
        return_code = self.process.wait()
        self.events.put(("done", return_code))

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "line":
                    self._handle_line(str(payload))
                elif kind == "done":
                    self._finish_job(int(payload))
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)

    def _handle_line(self, line: str) -> None:
        event = parse_runner_event(line)
        if event is None:
            self._append_log(line)
            return
        event_type = str(event.get("type", ""))
        message = str(event.get("message", ""))
        progress = event.get("progress")
        output_path = str(event.get("output_path", "")).strip()

        if event_type == "log":
            self._append_log(message)
            return
        if message:
            self._set_runner_status(message)
            self._append_log(translate_runner_status(message, self.language))
        if isinstance(progress, int):
            self.progress.stop()
            self.progress.configure(mode="determinate", value=progress)
        if output_path:
            self.result_path = Path(output_path)
        if event_type == "error":
            self._append_log(self._t("log.error", message=message))
        elif event_type == "cancelled":
            self._append_log(self._t("log.cancelled"))

    def _finish_job(self, return_code: int) -> None:
        action = self.current_action
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self._set_busy(False)
        self.process = None
        self.worker = None
        self.current_action = ""

        if return_code == 0:
            self.progress.configure(value=100)
            if action == "Convert":
                self._set_status("status.success")
                if self.result_path is None:
                    self.result_path = OUTPUT_ROOT
                self.open_button.configure(state="normal")
                self.copy_button.configure(state="normal")
                messagebox.showinfo(self._t("dialog.ready.title"), self._t("dialog.ready.body", path=self.result_path))
            else:
                self.environment = environment_summary()
                self._refresh_environment_label()
                self._set_status("status.environment_success")
        elif return_code == 2:
            self.progress.configure(value=0)
            self._set_status("status.cancelled")
        else:
            self.progress.configure(value=0)
            self._set_status("status.failed")
            messagebox.showerror(self._t("dialog.error.title"), self._t("dialog.error.body"))

        if self.close_when_done:
            self.root.destroy()

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.run_button.configure(state=state)
        self.cancel_button.configure(state="normal" if busy else "disabled")
        self.file_entry.configure(state=state)
        self.mode_combo.configure(state="disabled" if busy else "readonly")
        self.dpi_spin.configure(state=state)
        if busy:
            self.excel_combo.configure(state="disabled")
            self.columns_spin.configure(state="disabled")
        else:
            self._on_file_changed()

    def _request_cancel(self) -> None:
        if self.process is None or self.cancel_file is None:
            return
        if not self.cancel_file.exists():
            self.cancel_file.write_text("cancel", encoding="utf-8")
        self.cancel_button.configure(state="disabled")
        self._set_status("status.cancel_requested")
        self._append_log(self._t("log.cancel_requested"))

    def _open_result(self) -> None:
        target = self.result_path or OUTPUT_ROOT
        if target.is_file():
            target = target.parent
        if not target.exists():
            messagebox.showwarning(self._t("dialog.result_missing"), str(target))
            return
        os.startfile(target)  # type: ignore[attr-defined]

    def _copy_ai_prompt(self) -> None:
        target = self.result_path or OUTPUT_ROOT
        if target.is_file():
            target = target.parent
        self.root.clipboard_clear()
        self.root.clipboard_append(make_ai_prompt(target, self.language))
        self.root.update()
        self._set_status("status.prompt_copied")

    def _append_log(self, message: str) -> None:
        if not message:
            return
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _on_close(self) -> None:
        if self.process is None:
            self.root.destroy()
            return
        approved = messagebox.askyesno(
            self._t("dialog.close.title"),
            self._t("dialog.close.body"),
        )
        if approved:
            self.close_when_done = True
            self._request_cancel()


def initial_file_from_arguments() -> Path | None:
    if len(sys.argv) < 2:
        return None
    candidate = Path(sys.argv[1].strip('"')).expanduser()
    return candidate.resolve() if candidate.is_file() else None


def main() -> None:
    root = tk.Tk()
    ContextPackGui(root, initial_file_from_arguments())
    root.mainloop()


if __name__ == "__main__":
    main()
