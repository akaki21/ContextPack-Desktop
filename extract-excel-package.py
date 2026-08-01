from __future__ import annotations

import argparse
import shutil
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def actual_bounds(ws) -> tuple[int, int]:
    max_row = 0
    max_col = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                max_row = max(max_row, cell.row)
                max_col = max(max_col, cell.column)
    return max_row, max_col


def display(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return text.replace("|", "\\|").replace("\n", "<br>")


def sheet_table(ws, max_row: int, max_col: int) -> str:
    if max_row == 0 or max_col == 0:
        return "_Empty sheet._\n"
    rows: list[str] = []
    headers = [get_column_letter(i) for i in range(1, max_col + 1)]
    rows.append("| Row | " + " | ".join(headers) + " |")
    rows.append("| ---: | " + " | ".join("---" for _ in headers) + " |")
    for row_idx in range(1, max_row + 1):
        values = [display(ws.cell(row_idx, col_idx).value) for col_idx in range(1, max_col + 1)]
        rows.append(f"| {row_idx} | " + " | ".join(values) + " |")
    return "\n".join(rows) + "\n"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_workbook")
    parser.add_argument("output_dir")
    args = parser.parse_args()

    source = Path(args.input_workbook).resolve()
    output = Path(args.output_dir).resolve()
    output.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output / source.name)

    formulas_book = load_workbook(source, data_only=False, read_only=False)
    values_book = load_workbook(source, data_only=True, read_only=False)

    values_parts = [f"# Displayed values - {source.name}\n"]
    formulas_parts = [f"# Formulas and cached results - {source.name}\n"]
    info_parts = [f"# Workbook information - {source.name}\n"]
    info_parts.append(f"- Sheets: {len(formulas_book.worksheets)}")
    info_parts.append(f"- Defined names: {len(formulas_book.defined_names)}")
    info_parts.append(f"- Calculation mode: {formulas_book.calculation.calcMode or 'unspecified'}")
    info_parts.append("")

    total_formulas = 0
    total_errors = 0
    for index, formula_ws in enumerate(formulas_book.worksheets, 1):
        value_ws = values_book[formula_ws.title]
        max_row, max_col = actual_bounds(formula_ws)
        formulas = []
        cached_errors = []

        for row in formula_ws.iter_rows(min_row=1, max_row=max_row or 1, min_col=1, max_col=max_col or 1):
            for cell in row:
                if cell.data_type == "f" or (isinstance(cell.value, str) and cell.value.startswith("=")):
                    cached = value_ws[cell.coordinate].value
                    formulas.append((cell.coordinate, cell.value, cached, cell.number_format))
                    if isinstance(cached, str) and cached.startswith("#"):
                        cached_errors.append((cell.coordinate, cached))

        total_formulas += len(formulas)
        total_errors += len(cached_errors)
        hidden_rows = sum(1 for dim in formula_ws.row_dimensions.values() if dim.hidden)
        hidden_cols = sum(1 for dim in formula_ws.column_dimensions.values() if dim.hidden)

        info_parts.extend([
            f"## {index}. {formula_ws.title}",
            f"- Visibility: {formula_ws.sheet_state}",
            f"- Populated range: A1:{get_column_letter(max_col)}{max_row}" if max_row and max_col else "- Populated range: empty",
            f"- Formulas: {len(formulas)}",
            f"- Merged ranges: {len(formula_ws.merged_cells.ranges)}",
            f"- Hidden rows: {hidden_rows}",
            f"- Hidden columns: {hidden_cols}",
            f"- Charts: {len(formula_ws._charts)}",
            f"- Embedded images: {len(formula_ws._images)}",
        ])
        if formula_ws.merged_cells.ranges:
            info_parts.append("- Merges: " + ", ".join(str(r) for r in formula_ws.merged_cells.ranges))
        if cached_errors:
            info_parts.append("- Cached formula errors: " + ", ".join(f"{c}={v}" for c, v in cached_errors))
        info_parts.append("")

        values_parts.extend([f"## {index}. {formula_ws.title}\n", sheet_table(value_ws, max_row, max_col)])
        formulas_parts.append(f"## {index}. {formula_ws.title}\n")
        if formulas:
            formulas_parts.extend([
                "| Cell | Formula | Cached result | Number format |",
                "| --- | --- | --- | --- |",
            ])
            for coordinate, formula, cached, number_format in formulas:
                formulas_parts.append(
                    f"| {coordinate} | {display(formula)} | {display(cached)} | {display(number_format)} |"
                )
        else:
            formulas_parts.append("_No formulas._")
        formulas_parts.append("")

    info_parts.insert(4, f"- Total formulas: {total_formulas}")
    info_parts.insert(5, f"- Cached formula errors: {total_errors}")

    (output / "values.md").write_text("\n".join(values_parts), encoding="utf-8")
    (output / "formulas.md").write_text("\n".join(formulas_parts), encoding="utf-8")
    (output / "workbook-info.md").write_text("\n".join(info_parts), encoding="utf-8")
    print(f"Extracted {len(formulas_book.worksheets)} sheets and {total_formulas} formulas")


if __name__ == "__main__":
    main()

