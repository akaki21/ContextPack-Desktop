"""Analyze one worksheet without writing files or controlling Microsoft Excel."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .cells import populated_cells
from .markdown import rectangular_table, sparse_table


RECTANGULAR_CELL_LIMIT = 250_000
SPARSE_CELL_LIMIT = 500_000


@dataclass(frozen=True)
class FormulaRecord:
    """A formula and the last cached result stored in the workbook file."""

    coordinate: str
    formula: Any
    cached_result: Any
    number_format: str


@dataclass(frozen=True)
class SheetAnalysis:
    """Everything the package writer needs to describe one worksheet."""

    populated_cells: int
    max_row: int
    max_column: int
    min_row: int
    min_column: int
    populated_row_span: int
    populated_column_span: int
    output_mode: str
    values_markdown: str
    formulas: tuple[FormulaRecord, ...]
    cached_errors: tuple[tuple[str, str], ...]
    hidden_rows: int
    hidden_columns: int
    charts: int
    images: int
    merged_ranges: int
    warning: str | None = None

    def metrics(self, *, index: int, title: str, visibility: str) -> dict[str, Any]:
        """Return the stable JSON representation consumed by the renderer."""

        return {
            "index": index,
            "title": title,
            "visibility": visibility,
            "populated_cells": self.populated_cells,
            "max_row": self.max_row,
            "max_column": self.max_column,
            "min_row": self.min_row,
            "min_column": self.min_column,
            "populated_row_span": self.populated_row_span,
            "populated_column_span": self.populated_column_span,
            "output_mode": self.output_mode,
            "formulas": len(self.formulas),
            "cached_formula_errors": len(self.cached_errors),
            "hidden_rows": self.hidden_rows,
            "hidden_columns": self.hidden_columns,
            "charts": self.charts,
            "images": self.images,
            "merged_ranges": self.merged_ranges,
        }


def analyze_sheet(formula_worksheet: Any, value_worksheet: Any) -> SheetAnalysis:
    """Collect values, formulas, dimensions, and visual-risk indicators."""

    formula_cells = populated_cells(formula_worksheet)
    value_cells = populated_cells(value_worksheet)
    max_row = max((cell.row for cell in formula_cells), default=0)
    max_column = max((cell.column for cell in formula_cells), default=0)
    min_row = min((cell.row for cell in formula_cells), default=0)
    min_column = min((cell.column for cell in formula_cells), default=0)
    populated_row_span = max_row - min_row + 1 if max_row else 0
    populated_column_span = max_column - min_column + 1 if max_column else 0
    rectangular_size = max_row * max_column

    if len(formula_cells) > SPARSE_CELL_LIMIT:
        raise RuntimeError(
            f"Sheet {formula_worksheet.title!r} contains more than {SPARSE_CELL_LIMIT:,} populated/stored cells; "
            "refine the workbook before packaging."
        )

    use_sparse = rectangular_size > RECTANGULAR_CELL_LIMIT
    warning = None
    if use_sparse:
        warning = (
            f"Sheet {formula_worksheet.title!r} uses sparse output because its rectangular range "
            f"contains {rectangular_size:,} cells."
        )
        values_markdown = sparse_table(value_cells)
    elif max_row and max_column:
        values_markdown = rectangular_table(value_worksheet, max_row, max_column)
    else:
        values_markdown = "_Empty sheet._\n"

    formulas: list[FormulaRecord] = []
    cached_errors: list[tuple[str, str]] = []
    for cell in formula_cells:
        if cell.data_type == "f" or (isinstance(cell.value, str) and cell.value.startswith("=")):
            cached = value_worksheet[cell.coordinate].value
            formulas.append(FormulaRecord(cell.coordinate, cell.value, cached, cell.number_format))
            if isinstance(cached, str) and cached.startswith("#"):
                cached_errors.append((cell.coordinate, cached))

    return SheetAnalysis(
        populated_cells=len(formula_cells),
        max_row=max_row,
        max_column=max_column,
        min_row=min_row,
        min_column=min_column,
        populated_row_span=populated_row_span,
        populated_column_span=populated_column_span,
        output_mode="sparse" if use_sparse else "rectangular",
        values_markdown=values_markdown,
        formulas=tuple(formulas),
        cached_errors=tuple(cached_errors),
        hidden_rows=sum(1 for dimension in formula_worksheet.row_dimensions.values() if dimension.hidden),
        hidden_columns=sum(1 for dimension in formula_worksheet.column_dimensions.values() if dimension.hidden),
        charts=len(getattr(formula_worksheet, "_charts", [])),
        images=len(getattr(formula_worksheet, "_images", [])),
        merged_ranges=len(formula_worksheet.merged_cells.ranges),
        warning=warning,
    )

