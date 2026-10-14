"""Rules that must pass before a processing job can start."""

from __future__ import annotations

import re

from .file_types import classify_input
from .job_options import JobOptions


def validation_error_key(options: JobOptions) -> str | None:
    """Return a translation key for the first invalid option, or ``None``.

    Returning a key instead of display text keeps validation independent from
    the GUI language and makes the rules straightforward to unit-test.
    """

    if not options.input_file.is_file():
        return "validation.file_missing"
    if options.mode not in {"Auto", "Fast", "Full", "Ocr"}:
        return "validation.mode"
    if not 96 <= options.dpi <= 300:
        return "validation.dpi"
    if options.excel_render_mode not in {"Workbook", "AutoFit", "Both"}:
        return "validation.excel_mode"
    if not 10 <= options.max_autofit_columns <= 200:
        return "validation.columns"
    if not re.fullmatch(r'(eng|kat|rus)(\+(eng|kat|rus))*', options.ocr_languages):
        return 'validation.ocr_languages'
    if options.pdf_render_mode not in {'All', 'Sparse', 'None'}:
        return 'validation.pdf_render_mode'
    if not 0 <= options.max_render_pages <= 100000:
        return 'validation.render_limit'
    if options.ocr_pdf_policy not in {'Keep', 'Temporary'}:
        return 'validation.ocr_pdf_policy'
    if options.source_policy not in {'Copy', 'Reference'}:
        return 'validation.source_policy'
    if options.mode == "Ocr" and classify_input(options.input_file) not in {"pdf", "image"}:
        return "validation.ocr_scope"
    if options.output_directory is not None and not options.output_directory.is_dir():
        return "validation.output_missing"
    return None
