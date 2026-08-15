from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from openpyxl import load_workbook

from contextpack.excel.analysis import analyze_sheet
from contextpack.excel.reporting import WorkbookReport
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

    external_links = len(getattr(formulas_book, "_external_links", []))
    report = WorkbookReport(source.name, external_links)

    for index, formula_ws in enumerate(formulas_book.worksheets, 1):
        value_ws = values_book[formula_ws.title]
        analysis = analyze_sheet(formula_ws, value_ws)
        folder_name = safe_sheet_folder(index, formula_ws.title)
        sheet_dir = sheets_root / folder_name
        report.add_sheet(
            index=index,
            worksheet=formula_ws,
            folder_name=folder_name,
            sheet_directory=sheet_dir,
            analysis=analysis,
        )

    report.write_summary(
        output,
        defined_names=len(formulas_book.defined_names),
        calculation_mode=calculation_mode(formulas_book),
    )
    formulas_book.close()
    values_book.close()
    print(f"Extracted {len(report.sheet_metrics)} sheets and {report.total_formulas} formulas")


if __name__ == "__main__":
    main()
