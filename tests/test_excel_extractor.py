from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

from contextpack.excel.analysis import analyze_sheet
from contextpack.excel.extractor import extract_workbook
from contextpack.excel.cells import populated_cells
from contextpack.excel.markdown import display, sparse_table
from contextpack.excel.workbook import calculation_mode, safe_sheet_folder


ROOT = Path(__file__).resolve().parents[1]
EXTRACTOR = ROOT / "extract-excel-package.py"


class MockWorkbookWithoutCalculation:
    """Minimal workbook shape used to verify the defensive metadata fallback."""


class ExcelExtractorTests(unittest.TestCase):
    def run_extractor(self, workbook_path: Path, output_path: Path) -> None:
        subprocess.run([sys.executable, str(EXTRACTOR), str(workbook_path), str(output_path)], check=True)

    def test_excel_helpers_are_safe_and_deterministic(self) -> None:
        self.assertEqual(display("first|second\nthird"), "first\\|second<br>third")
        self.assertEqual(safe_sheet_folder(2, 'Bad:/Name*.'), "02-Bad__Name_")

        workbook = Workbook()
        sheet = workbook.active
        sheet["B2"] = "second"
        sheet["A1"] = "first"
        cells = populated_cells(sheet)
        self.assertEqual([cell.coordinate for cell in cells], ["B2", "A1"])
        rendered = sparse_table(cells)
        self.assertLess(rendered.index("A1"), rendered.index("B2"))
        self.assertEqual(calculation_mode(MockWorkbookWithoutCalculation()), "unspecified")

    def test_sheet_analysis_collects_formula_and_layout_risks(self) -> None:
        formulas_book = Workbook()
        formula_sheet = formulas_book.active
        formula_sheet.title = "Analysis"
        formula_sheet["B2"] = "=1/0"
        formula_sheet.row_dimensions[2].hidden = True
        formula_sheet.column_dimensions["B"].hidden = True
        formula_sheet.merge_cells("A1:B1")

        values_book = Workbook()
        value_sheet = values_book.active
        value_sheet.title = "Analysis"
        value_sheet["B2"] = "#DIV/0!"

        analysis = analyze_sheet(formula_sheet, value_sheet)
        metrics = analysis.metrics(index=1, title="Analysis", visibility="visible")

        self.assertEqual(len(analysis.formulas), 1)
        self.assertEqual(analysis.formulas[0].coordinate, "B2")
        self.assertEqual(analysis.cached_errors, (("B2", "#DIV/0!"),))
        self.assertEqual(metrics["hidden_rows"], 1)
        self.assertEqual(metrics["hidden_columns"], 1)
        self.assertEqual(metrics["merged_ranges"], 1)
        self.assertEqual(metrics["cached_formula_errors"], 1)

    def test_extractor_closes_first_workbook_if_second_open_fails(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "sample.xlsx"
            source.write_bytes(b"mock workbook")
            formulas_book = Mock()

            with patch(
                "contextpack.excel.extractor.load_workbook",
                side_effect=[formulas_book, RuntimeError("second open failed")],
            ):
                with self.assertRaisesRegex(RuntimeError, "second open failed"):
                    extract_workbook(source, root / "output")

            formulas_book.close.assert_called_once_with()

    def test_splits_values_and_formulas_per_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "sample.xlsx"
            output = root / "package"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Summary"
            sheet["A1"] = 4
            sheet["B1"] = 5
            sheet["C1"] = "=A1+B1"
            workbook.save(source)
            self.run_extractor(source, output)

            metrics = json.loads((output / "excel-metrics.json").read_text(encoding="utf-8"))
            self.assertEqual(metrics["sheet_count"], 1)
            self.assertEqual(metrics["total_formulas"], 1)
            self.assertEqual(metrics["sheets"][0]["min_row"], 1)
            self.assertEqual(metrics["sheets"][0]["min_column"], 1)
            self.assertEqual(metrics["sheets"][0]["populated_column_span"], 3)
            self.assertEqual(metrics["sheets"][0]["merged_ranges"], 0)
            self.assertTrue((output / "sheets-data" / "01-Summary" / "values.md").exists())
            self.assertIn("=A1+B1", (output / "sheets-data" / "01-Summary" / "formulas.md").read_text(encoding="utf-8"))

    def test_extreme_dimension_uses_sparse_output(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "sparse.xlsx"
            output = root / "package"
            workbook = Workbook()
            sheet = workbook.active
            sheet["A1"] = "start"
            sheet["XFD1048576"] = "end"
            workbook.save(source)
            self.run_extractor(source, output)

            metrics = json.loads((output / "excel-metrics.json").read_text(encoding="utf-8"))
            self.assertEqual(metrics["sheets"][0]["output_mode"], "sparse")
            self.assertEqual(metrics["sheets"][0]["populated_column_span"], 16_384)
            values = (output / "sheets-data" / "01-Sheet" / "values.md").read_text(encoding="utf-8")
            self.assertIn("XFD1048576", values)
            self.assertLess(len(values), 10_000)

    def test_reports_drawing_risk_metrics(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "chart.xlsx"
            output = root / "package"
            workbook = Workbook()
            sheet = workbook.active
            for row in range(1, 5):
                sheet.append([f"Row {row}", row])
            chart = BarChart()
            chart.add_data(Reference(sheet, min_col=2, min_row=1, max_row=4))
            sheet.add_chart(chart, "D2")
            workbook.save(source)
            self.run_extractor(source, output)

            metrics = json.loads((output / "excel-metrics.json").read_text(encoding="utf-8"))
            self.assertEqual(metrics["sheets"][0]["charts"], 1)
            self.assertEqual(metrics["sheets"][0]["images"], 0)

    def test_missing_calculation_properties_are_reported_as_unspecified(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            original = root / "original.xlsx"
            source = root / "without-calculation-properties.xlsx"
            output = root / "package"
            workbook = Workbook()
            workbook.active["A1"] = "valid workbook without calcPr"
            workbook.save(original)

            with zipfile.ZipFile(original) as input_archive, zipfile.ZipFile(source, "w") as output_archive:
                for item in input_archive.infolist():
                    content = input_archive.read(item.filename)
                    if item.filename == "xl/workbook.xml":
                        text = content.decode("utf-8")
                        text = text.replace(
                            '<calcPr calcId="124519" fullCalcOnLoad="1"/>',
                            "",
                        )
                        content = text.encode("utf-8")
                    output_archive.writestr(item, content)

            self.run_extractor(source, output)

            info = (output / "workbook-info.md").read_text(encoding="utf-8")
            self.assertIn("- Calculation mode: unspecified", info)


if __name__ == "__main__":
    unittest.main()
