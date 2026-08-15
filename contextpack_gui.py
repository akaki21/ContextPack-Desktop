from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from dataclasses import replace
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from contextpack.file_types import EXCEL_EXTENSIONS, IMAGE_EXTENSIONS, classify_input
from contextpack.environment import environment_summary, excel_is_available, find_tesseract, preferred_output_directory
from contextpack.gui.theme import COLORS, configure_gui_styles
from contextpack.gui.layout import build_main_layout
from contextpack.job_options import JobOptions
from contextpack.job_controller import launch_background_job, new_cancel_file, request_cancellation
from contextpack.localization import (
    EXCEL_MODE_LABELS,
    LANGUAGE_LABELS,
    MODE_LABELS,
    STATUS_TRANSLATIONS,
    TEXT,
    label_for_value,
    translate,
    translate_runner_status,
)
from contextpack.paths import OUTPUT_ROOT, RUNNER
from contextpack.runner_command import build_runner_command, powershell_executable
from contextpack.runner_events import parse_runner_event
from contextpack.validation import validation_error_key


ROOT = Path(__file__).resolve().parent


def validate_options(options: JobOptions, language: str = "ka") -> str | None:
    error_key = validation_error_key(options)
    return translate(language, error_key) if error_key else None


def make_ai_prompt(output_path: Path, language: str = "ka") -> str:
    return "\n".join(
        (
            translate(language, "prompt.line1", path=output_path),
            translate(language, "prompt.line2"),
            translate(language, "prompt.goal"),
        )
    )


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
        configure_gui_styles(self.root)

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
        build_main_layout(self)

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
        self.cancel_file = new_cancel_file()
        command = build_runner_command(action, options=options, cancel_file=self.cancel_file)
        try:
            self.process, self.worker = launch_background_job(
                command,
                working_directory=ROOT,
                events=self.events,
            )
        except OSError as error:
            self.process = None
            self.worker = None
            messagebox.showerror(self._t("dialog.launch_failed"), str(error))
            return

        self._set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self._set_status("status.starting")

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
        request_cancellation(self.cancel_file)
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
