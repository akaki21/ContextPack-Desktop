"""Build the visual sections of the Desktop window.

The functions receive the application object because Tkinter variables and
callbacks are owned by it. Keeping each section separate makes it clear which
widgets belong together without moving any job-processing behavior here.
"""

from __future__ import annotations

from typing import Any

import tkinter as tk
from tkinter import ttk, messagebox

from ..localization import EXCEL_MODE_LABELS, LANGUAGE_LABELS, MODE_LABELS, OCR_LANGUAGE_LABELS, PDF_RENDER_LABELS, SOURCE_POLICY_LABELS, OCR_PDF_POLICY_LABELS, label_for_value
from .theme import COLORS


def build_header(app: Any, shell: ttk.Frame) -> None:
    header = ttk.Frame(shell)
    header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
    header.columnconfigure(0, weight=1)
    ttk.Label(header, text="ContextPack Desktop", style="Title.TLabel").grid(row=0, column=0, sticky="w")
    app._localized_label(header, "subtitle", style="Subtitle.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))
    app.language_combo = ttk.Combobox(header, state="readonly", values=list(LANGUAGE_LABELS), textvariable=app.language_var, width=11)
    app.language_combo.grid(row=0, column=2, sticky="e")
    app.language_combo.bind("<<ComboboxSelected>>", app._change_language)
    app.environment_label = ttk.Label(header, textvariable=app.environment_var, style="Subtitle.TLabel")
    app.environment_label.grid(row=1, column=1, columnspan=2, sticky="e")


def build_file_section(app: Any, shell: ttk.Frame) -> None:
    file_card = ttk.Frame(shell, style="Card.TFrame", padding=15)
    file_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
    file_card.columnconfigure(0, weight=1)
    app._localized_label(file_card, "file.section", style="Section.TLabel").grid(row=0, column=0, sticky="w")
    app._localized_label(file_card, "file.hint", style="Hint.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 10))

    entry_frame = ttk.Frame(file_card, style="Card.TFrame")
    entry_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
    entry_frame.columnconfigure(0, weight=1)
    app.file_entry = tk.Entry(
        entry_frame,
        textvariable=app.file_var,
        bg=COLORS["card_alt"],
        fg=COLORS["text"],
        insertbackground=COLORS["text"],
        relief="flat",
        font=("Segoe UI", 10),
        highlightthickness=1,
        highlightbackground=COLORS["border"],
        highlightcolor=COLORS["accent"],
    )
    app.file_entry.grid(row=0, column=0, sticky="ew", ipady=9)
    app.file_entry.bind("<FocusOut>", lambda _event: app._on_file_changed())
    app._localized_button(entry_frame, "file.browse", style="Secondary.TButton", command=app._choose_file).grid(row=0, column=1, padx=(10, 0))
    ttk.Label(file_card, textvariable=app.file_kind_var, style="Hint.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 0))


def build_options_section(app: Any, shell: ttk.Frame) -> None:
    options_card = ttk.Frame(shell, style="Card.TFrame", padding=15)
    options_card.grid(row=2, column=0, sticky="ew", pady=(0, 10))
    options_card.columnconfigure(0, weight=1)
    options_card.columnconfigure(1, weight=1)
    app._localized_label(options_card, "options.section", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
    app._localized_label(options_card, "options.hint", style="Hint.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 12))

    app._localized_label(options_card, "options.mode", style="Card.TLabel").grid(row=2, column=0, sticky="w")
    app.mode_combo = ttk.Combobox(options_card, state="readonly", values=list(MODE_LABELS[app.language]), textvariable=app.mode_var, width=27)
    app.mode_combo.grid(row=3, column=0, sticky="ew", padx=(0, 10), pady=(4, 0))
    app._localized_label(options_card, "options.excel", style="Card.TLabel").grid(row=2, column=1, sticky="w")
    app.excel_combo = ttk.Combobox(options_card, state="disabled", values=list(EXCEL_MODE_LABELS[app.language]), textvariable=app.excel_mode_var, width=27)
    app.excel_combo.grid(row=3, column=1, sticky="ew", pady=(4, 0))

    ttk.Label(options_card, text="DPI", style="Card.TLabel").grid(row=4, column=0, sticky="w", pady=(12, 0))
    app.dpi_spin = ttk.Spinbox(options_card, from_=96, to=300, increment=12, textvariable=app.dpi_var, width=8)
    app.dpi_spin.grid(row=5, column=0, sticky="ew", padx=(0, 10), pady=(4, 0))
    app._localized_label(options_card, "options.columns", style="Card.TLabel").grid(row=4, column=1, sticky="w", pady=(12, 0))
    app.columns_spin = ttk.Spinbox(options_card, from_=10, to=200, increment=5, textvariable=app.max_columns_var, width=8, state="disabled")
    app.columns_spin.grid(row=5, column=1, sticky="ew", pady=(4, 0))
    app.advanced_button = app._localized_button(options_card, 'options.advanced', style='Secondary.TButton', command=lambda: show_processing_options(app))
    app.advanced_button.grid(row=6, column=0, columnspan=2, sticky='w', pady=(12, 0))


def show_processing_options(app: Any) -> tk.Toplevel:
    dialog = tk.Toplevel(app.root)
    dialog.title(app._t('options.advanced'))
    dialog.transient(app.root)
    dialog.resizable(False, False)
    frame = ttk.Frame(dialog, padding=20)
    frame.grid(sticky='nsew')
    language = tk.StringVar(dialog, value=next(label for label, code in OCR_LANGUAGE_LABELS.items() if code == app.ocr_languages_var.get()))
    render = tk.StringVar(dialog, value=label_for_value(PDF_RENDER_LABELS, app.language, app.pdf_render_var.get()))
    policy = tk.StringVar(dialog, value=label_for_value(SOURCE_POLICY_LABELS, app.language, app.source_policy_var.get()))
    ocr_pdf = tk.StringVar(dialog, value=label_for_value(OCR_PDF_POLICY_LABELS, app.language, app.ocr_pdf_policy_var.get()))
    limit = tk.StringVar(dialog, value=str(app.max_render_pages_var.get()))
    force = tk.BooleanVar(dialog, value=app.force_var.get())
    for row, (key, variable, labels) in enumerate([
        ('options.ocr_languages', language, OCR_LANGUAGE_LABELS),
        ('options.pdf_render', render, PDF_RENDER_LABELS[app.language]),
        ('options.source_policy', policy, SOURCE_POLICY_LABELS[app.language]),
        ('options.ocr_pdf_policy', ocr_pdf, OCR_PDF_POLICY_LABELS[app.language]),
    ]):
        ttk.Label(frame, text=app._t(key)).grid(row=row * 2, column=0, sticky='w', pady=(8, 3))
        ttk.Combobox(frame, textvariable=variable, values=list(labels), state='readonly', width=58).grid(row=row * 2 + 1, column=0, sticky='ew')
    ttk.Label(frame, text=app._t('options.render_limit')).grid(row=8, column=0, sticky='w', pady=(12, 3))
    ttk.Spinbox(frame, from_=0, to=100000, textvariable=limit).grid(row=9, column=0, sticky='ew')
    ttk.Checkbutton(frame, text=app._t('options.force'), variable=force).grid(row=10, column=0, sticky='w', pady=12)
    ttk.Label(frame, text=app._t('options.reference_hint'), wraplength=540).grid(row=11, column=0, sticky='w')

    def save() -> None:
        try:
            value = int(limit.get())
            if not 0 <= value <= 100000:
                raise ValueError
        except ValueError:
            messagebox.showerror(app._t('dialog.invalid.title'), app._t('validation.render_limit'), parent=dialog)
            return
        app.ocr_languages_var.set(OCR_LANGUAGE_LABELS[language.get()])
        app.pdf_render_var.set(PDF_RENDER_LABELS[app.language][render.get()])
        app.source_policy_var.set(SOURCE_POLICY_LABELS[app.language][policy.get()])
        app.ocr_pdf_policy_var.set(OCR_PDF_POLICY_LABELS[app.language][ocr_pdf.get()])
        app.max_render_pages_var.set(value)
        app.force_var.set(force.get())
        dialog.destroy()

    ttk.Label(frame, text=app._t('options.library_hint'), wraplength=540).grid(row=12, column=0, sticky='w', pady=(8, 0))
    ttk.Button(frame, text=app._t('options.apply'), command=save).grid(row=13, column=0, sticky='e', pady=(16, 0))
    dialog.grab_set()
    return dialog


def build_action_section(app: Any, shell: ttk.Frame) -> None:
    action_card = ttk.Frame(shell, style="AltCard.TFrame", padding=15)
    action_card.grid(row=3, column=0, sticky="ew", pady=(0, 10))
    action_card.columnconfigure(1, weight=1)
    app._localized_label(action_card, "action.hint", style="Status.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
    app.run_button = app._localized_button(action_card, "action.start", style="Accent.TButton", command=app._start_conversion)
    app.run_button.grid(row=1, column=0, rowspan=2, sticky="nsw", pady=(10, 0))
    app.status_label = ttk.Label(action_card, textvariable=app.status_var, style="Status.TLabel")
    app.status_label.grid(row=1, column=1, sticky="sw", padx=(16, 12), pady=(10, 4))
    app.progress = ttk.Progressbar(action_card, mode="determinate", maximum=100)
    app.progress.grid(row=2, column=1, sticky="ew", padx=(16, 12), pady=(0, 1))
    app.cancel_button = app._localized_button(action_card, "action.cancel", style="Danger.TButton", command=app._request_cancel, state="disabled")
    app.cancel_button.grid(row=1, column=2, rowspan=2, sticky="nse", pady=(10, 0))


def build_details_section(app: Any, shell: ttk.Frame) -> None:
    details = ttk.Frame(shell, style="Card.TFrame", padding=14)
    details.grid(row=4, column=0, sticky="nsew")
    details.columnconfigure(0, weight=1)
    details.rowconfigure(1, weight=1)
    toolbar = ttk.Frame(details, style="Card.TFrame")
    toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
    app._localized_label(toolbar, "details.title", style="Section.TLabel").pack(side="left")
    app._localized_button(toolbar, "details.check", style="Secondary.TButton", command=lambda: app._start_utility("Check")).pack(side="right")
    app._localized_button(toolbar, "details.setup", style="Secondary.TButton", command=app._confirm_setup).pack(side="right", padx=(0, 8))

    log_frame = tk.Frame(details, bg=COLORS["card"])
    log_frame.grid(row=1, column=0, sticky="nsew")
    log_frame.rowconfigure(0, weight=1)
    log_frame.columnconfigure(0, weight=1)
    app.log = tk.Text(log_frame, bg="#0a1020", fg="#cbd5e1", insertbackground=COLORS["text"], relief="flat", font=("Cascadia Mono", 9), wrap="word", height=5, padx=12, pady=10, state="disabled")
    app.log.grid(row=0, column=0, sticky="nsew")
    scroll = ttk.Scrollbar(log_frame, orient="vertical", command=app.log.yview)
    scroll.grid(row=0, column=1, sticky="ns")
    app.log.configure(yscrollcommand=scroll.set)

    result_bar = ttk.Frame(details, style="Card.TFrame")
    result_bar.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    app.open_button = app._localized_button(result_bar, "result.open", style="Secondary.TButton", command=app._open_result, state="disabled")
    app.open_button.pack(side="left")
    app.copy_button = app._localized_button(result_bar, "result.copy", style="Secondary.TButton", command=app._copy_ai_prompt, state="disabled")
    app.copy_button.pack(side="left", padx=(8, 0))
    app._localized_label(result_bar, "result.local", style="Hint.TLabel").pack(side="right")


def build_main_layout(app: Any) -> None:
    """Build and connect all visual sections in their display order."""

    shell = ttk.Frame(app.root, padding=(24, 18))
    shell.grid(row=0, column=0, sticky="nsew")
    app.root.rowconfigure(0, weight=1)
    app.root.columnconfigure(0, weight=1)
    shell.columnconfigure(0, weight=1)
    shell.rowconfigure(4, weight=1)
    build_header(app, shell)
    build_file_section(app, shell)
    build_options_section(app, shell)
    build_action_section(app, shell)
    build_details_section(app, shell)
