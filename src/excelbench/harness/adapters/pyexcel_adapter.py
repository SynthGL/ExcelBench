"""Adapter for pyexcel library (read+write, value-only)."""

import re
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

import pyexcel

from excelbench.harness.adapters.base import ExcelAdapter
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)


def _get_version() -> str:
    try:
        from importlib.metadata import version
        return version("pyexcel")
    except Exception:
        return "unknown"


def _parse_cell_ref(cell: str) -> tuple[int, int]:
    """Parse a cell reference like 'A1' to (row_0based, col_0based)."""
    match = re.match(r"([A-Z]+)(\d+)", cell.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell}")
    col_str, row_str = match.groups()
    row = int(row_str) - 1
    col = 0
    for char in col_str:
        col = col * 26 + (ord(char) - ord("A") + 1)
    col -= 1
    return row, col


class PyexcelAdapter(ExcelAdapter):
    """Adapter for pyexcel library (read+write, value-only).

    pyexcel is a meta-library wrapping openpyxl (via pyexcel-xlsx).
    It exposes cell values only — no formatting, conditional formatting,
    comments, or images.
    """

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="pyexcel",
            version=_get_version(),
            language="python",
            capabilities={"read", "write"},
        )

    # =========================================================================
    # Read Operations
    # =========================================================================

    def open_workbook(self, path: Path) -> Any:
        return pyexcel.get_book(file_name=str(path))

    def close_workbook(self, workbook: Any) -> None:
        pass

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return workbook.sheet_names()

    def read_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellValue:
        row_idx, col_idx = _parse_cell_ref(cell)
        try:
            ws = workbook.sheet_by_name(sheet)
        except (KeyError, AttributeError):
            return CellValue(type=CellType.BLANK)

        if row_idx >= ws.number_of_rows():
            return CellValue(type=CellType.BLANK)
        row_data = ws.row_at(row_idx)
        if col_idx >= len(row_data):
            return CellValue(type=CellType.BLANK)

        value = row_data[col_idx]

        if value is None or value == "":
            return CellValue(type=CellType.BLANK)

        if isinstance(value, bool):
            return CellValue(type=CellType.BOOLEAN, value=value)

        if isinstance(value, (int, float)):
            return CellValue(type=CellType.NUMBER, value=value)

        if isinstance(value, datetime):
            is_midnight = (
                value.hour == 0 and value.minute == 0
                and value.second == 0 and value.microsecond == 0
            )
            if is_midnight:
                return CellValue(type=CellType.DATE, value=value.date())
            return CellValue(type=CellType.DATETIME, value=value)

        if isinstance(value, date) and not isinstance(value, datetime):
            return CellValue(type=CellType.DATE, value=value)

        if isinstance(value, time):
            return CellValue(
                type=CellType.DATETIME,
                value=datetime.combine(date.today(), value),
            )

        if isinstance(value, str):
            if value in ("#N/A", "#NULL!", "#NAME?", "#REF!"):
                return CellValue(type=CellType.ERROR, value=value)
            if value.startswith("#") and value.endswith("!"):
                return CellValue(type=CellType.ERROR, value=value)
            if value.startswith("="):
                return CellValue(type=CellType.FORMULA, value=value, formula=value)
            return CellValue(type=CellType.STRING, value=value)

        return CellValue(type=CellType.STRING, value=str(value))

    def read_cell_format(self, workbook: Any, sheet: str, cell: str) -> CellFormat:
        return CellFormat()

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        return BorderInfo()

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        return None

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        return None

    # =========================================================================
    # Tier 2 Read Operations
    # =========================================================================

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        return []

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_data_validations(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_images(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_comments(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_freeze_panes(self, workbook: Any, sheet: str) -> dict:
        return {}

    # =========================================================================
    # Write Operations
    # =========================================================================

    def create_workbook(self) -> dict:
        return {"sheets": {}, "_order": []}

    def add_sheet(self, workbook: dict, name: str) -> None:
        if name not in workbook["sheets"]:
            workbook["sheets"][name] = {}
            workbook["_order"].append(name)

    def write_cell_value(
        self,
        workbook: dict,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        if sheet not in workbook["sheets"]:
            workbook["sheets"][sheet] = {}
            workbook["_order"].append(sheet)

        row_idx, col_idx = _parse_cell_ref(cell)

        raw_value: Any = ""
        if value.type == CellType.BLANK:
            raw_value = ""
        elif value.type == CellType.FORMULA:
            raw_value = value.formula or value.value or ""
        elif value.type == CellType.BOOLEAN:
            raw_value = bool(value.value)
        elif value.type == CellType.NUMBER:
            raw_value = value.value
        elif value.type in (CellType.DATE, CellType.DATETIME):
            raw_value = value.value
        elif value.type == CellType.ERROR:
            raw_value = str(value.value)
        else:
            raw_value = str(value.value) if value.value is not None else ""

        workbook["sheets"][sheet][(row_idx, col_idx)] = raw_value

    def write_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        pass  # pyexcel does not support formatting

    def write_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        pass  # pyexcel does not support borders

    def set_row_height(self, workbook: Any, sheet: str, row: int, height: float) -> None:
        pass

    def set_column_width(self, workbook: Any, sheet: str, column: str, width: float) -> None:
        pass

    # =========================================================================
    # Tier 2 Write Operations
    # =========================================================================

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        pass

    def add_conditional_format(self, workbook: Any, sheet: str, rule: dict) -> None:
        pass

    def add_data_validation(self, workbook: Any, sheet: str, validation: dict) -> None:
        pass

    def add_hyperlink(self, workbook: Any, sheet: str, link: dict) -> None:
        pass

    def add_image(self, workbook: Any, sheet: str, image: dict) -> None:
        pass

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: dict) -> None:
        pass

    def add_comment(self, workbook: Any, sheet: str, comment: dict) -> None:
        pass

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: dict) -> None:
        pass

    def save_workbook(self, workbook: dict, path: Path) -> None:
        book_dict = {}
        for name in workbook["_order"]:
            cells = workbook["sheets"].get(name, {})
            if not cells:
                book_dict[name] = [[""]]
                continue
            max_row = max(r for r, _ in cells.keys()) + 1
            max_col = max(c for _, c in cells.keys()) + 1
            grid = [[""] * max_col for _ in range(max_row)]
            for (r, c), val in cells.items():
                grid[r][c] = val
            book_dict[name] = grid

        book = pyexcel.Book(book_dict)
        book.save_as(str(path))
