from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from contextpack.excel.analysis import analyze_sheet
from contextpack.excel.markdown import display
from contextpack.excel.workbook import calculation_mode, safe_sheet_folder


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_workbook")
    parser.add_argument("output_dir")
    args = parser.parse_args()

    source = Path(args.input_workbook).resolve()
    output = Path(args.output_dir).resolve()
    output.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output / source.name)

    keep_vba = source.suffix.lower() in {".xlsm", ".xltm"}
    formulas_book = load_workbook(source, data_only=False, read_only=False, keep_vba=keep_vba, keep_links=True)
    values_book = load_workbook(source, data_only=True, read_only=False, keep_vba=keep_vba, keep_links=True)
    sheets_root = output / "sheets-data"
    sheets_root.mkdir(parents=True, exist_ok=True)

    values_index = [f"# Displayed values index — {source.name}", ""]
    formulas_index = [f"# Formulas index — {source.name}", ""]
    info_parts = [f"# Workbook information — {source.name}", ""]
    total_formulas = 0
    total_errors = 0
    quality_warnings: list[str] = []
    sheet_metrics: list[dict[str, Any]] = []

    external_links = len(getattr(formulas_book, "_external_links", []))
    if external_links:
        quality_warnings.append(f"Workbook contains {external_links} external link(s).")

    for index, formula_ws in enumerate(formulas_book.worksheets, 1):
        value_ws = values_book[formula_ws.title]
        analysis = analyze_sheet(formula_ws, value_ws)
        folder_name = safe_sheet_folder(index, formula_ws.title)
        sheet_dir = sheets_root / folder_name
        sheet_dir.mkdir(parents=True, exist_ok=True)

        if analysis.warning:
            quality_warnings.append(analysis.warning)
        total_formulas += len(analysis.formulas)
        total_errors += len(analysis.cached_errors)

        values_path = sheet_dir / "values.md"
        formulas_path = sheet_dir / "formulas.md"
        values_path.write_text(f"# Values — {formula_ws.title}\n\n{analysis.values_markdown}", encoding="utf-8")
        formula_lines = [f"# Formulas — {formula_ws.title}", ""]
        if analysis.formulas:
            formula_lines.extend(["| Cell | Formula | Cached result | Number format |", "| --- | --- | --- | --- |"])
            formula_lines.extend(
                f"| {record.coordinate} | {display(record.formula)} | {display(record.cached_result)} | {display(record.number_format)} |"
                for record in analysis.formulas
            )
        else:
            formula_lines.append("_No formulas._")
        formulas_path.write_text("\n".join(formula_lines) + "\n", encoding="utf-8")

        relative_values = f"sheets-data/{folder_name}/values.md"
        relative_formulas = f"sheets-data/{folder_name}/formulas.md"
        values_index.append(f"- [{index}. {formula_ws.title}]({relative_values})")
        formulas_index.append(f"- [{index}. {formula_ws.title}]({relative_formulas}) — {len(analysis.formulas)} formula(s)")
        info_parts.extend(
            [
                f"## {index}. {formula_ws.title}",
                f"- Visibility: {formula_ws.sheet_state}",
                f"- Populated cells: {analysis.populated_cells}",
                f"- Populated bounds: A1:{get_column_letter(analysis.max_column)}{analysis.max_row}" if analysis.max_row and analysis.max_column else "- Populated bounds: empty",
                f"- Output mode: {analysis.output_mode}",
                f"- Formulas: {len(analysis.formulas)}",
                f"- Cached formula errors: {len(analysis.cached_errors)}",
                f"- Merged ranges: {analysis.merged_ranges}",
                f"- Hidden rows / columns: {analysis.hidden_rows} / {analysis.hidden_columns}",
                f"- Charts / embedded images: {analysis.charts} / {analysis.images}",
                f"- Values: [{relative_values}]({relative_values})",
                f"- Formulas: [{relative_formulas}]({relative_formulas})",
                "",
            ]
        )
        sheet_metrics.append(analysis.metrics(index=index, title=formula_ws.title, visibility=formula_ws.sheet_state))

    summary = [
        f"- Sheets: {len(formulas_book.worksheets)}",
        f"- Defined names: {len(formulas_book.defined_names)}",
        f"- External links: {external_links}",
        f"- Calculation mode: {calculation_mode(formulas_book)}",
        f"- Total formulas: {total_formulas}",
        f"- Cached formula errors: {total_errors}",
        "",
    ]
    info_parts[2:2] = summary
    (output / "values.md").write_text("\n".join(values_index) + "\n", encoding="utf-8")
    (output / "formulas.md").write_text("\n".join(formulas_index) + "\n", encoding="utf-8")
    (output / "workbook-info.md").write_text("\n".join(info_parts), encoding="utf-8")

    quality = [
        f"# Quality report — {source.name}",
        "",
        f"- Sheets analyzed: {len(sheet_metrics)}",
        f"- Total formulas: {total_formulas}",
        f"- Cached formula errors: {total_errors}",
        f"- External links: {external_links}",
        "",
        "## Warnings",
        "",
    ]
    quality.extend(f"- {warning}" for warning in quality_warnings)
    if not quality_warnings:
        quality.append("- No structural warnings detected.")
    quality.extend(["", "Cached values may be stale if Excel did not recalculate and save the workbook before packaging.", ""])
    (output / "quality-report.md").write_text("\n".join(quality), encoding="utf-8")
    (output / "excel-metrics.json").write_text(
        json.dumps(
            {
                "sheet_count": len(sheet_metrics),
                "total_formulas": total_formulas,
                "cached_formula_errors": total_errors,
                "external_links": external_links,
                "warnings": quality_warnings,
                "sheets": sheet_metrics,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    formulas_book.close()
    values_book.close()
    print(f"Extracted {len(sheet_metrics)} sheets and {total_formulas} formulas")


if __name__ == "__main__":
    main()
