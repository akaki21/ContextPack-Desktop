"""Read the cells that are actually stored in an openpyxl worksheet."""

from __future__ import annotations

from typing import Any


def instantiated_cells(worksheet: Any) -> list[Any]:
    """Return stored cells without expanding a worksheet's declared bounds.

    Some workbooks claim the full Excel grid as their used range. Iterating that
    rectangle would allocate millions of empty cells, so the internal stored-cell
    mapping is preferred when openpyxl exposes it.
    """

    cells = getattr(worksheet, "_cells", None)
    if isinstance(cells, dict):
        return list(cells.values())
    return [cell for row in worksheet.iter_rows() for cell in row]


def populated_cells(worksheet: Any) -> list[Any]:
    """Return only stored cells that contain a value or formula."""

    return [cell for cell in instantiated_cells(worksheet) if cell.value is not None]

