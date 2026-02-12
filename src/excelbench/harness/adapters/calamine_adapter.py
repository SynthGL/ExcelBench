"""Adapter for python-calamine library (read-only, Rust-backed)."""

import re
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

from python_calamine import CalamineWorkbook

from excelbench.harness.adapters.base import ReadOnlyAdapter
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)

JSONDict = dict[str, Any]


def _get_version() -> str:
    """Get python-calamine version."""
    try:
        from importlib.metadata import version

        return version("python-calamine")
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


def _convert_value(value: Any) -> CellValue:
    """Convert a raw calamine Python value to a CellValue."""
    if value is None or value == "":
        return CellValue(type=CellType.BLANK)

    # Check bool BEFORE int (bool is subclass of int in Python)
    if isinstance(value, bool):
        return CellValue(type=CellType.BOOLEAN, value=value)

    if isinstance(value, (int, float)):
        return CellValue(type=CellType.NUMBER, value=value)

    # Check date before datetime (date is not a subclass of datetime,
    # but calamine may return either)
    if isinstance(value, datetime):
        is_midnight = (
            value.hour == 0
            and value.minute == 0
            and value.second == 0
            and value.microsecond == 0
        )
        if is_midnight:
            return CellValue(type=CellType.DATE, value=value.date())
        return CellValue(type=CellType.DATETIME, value=value)

    if isinstance(value, date) and not isinstance(value, datetime):
        return CellValue(type=CellType.DATE, value=value)

    if isinstance(value, time):
        return CellValue(type=CellType.DATETIME, value=datetime.combine(date.today(), value))

    if isinstance(value, str):
        # Error values — includes #N/A (no trailing !)
        if value in ("#N/A", "#NULL!", "#NAME?", "#REF!"):
            return CellValue(type=CellType.ERROR, value=value)
        if value.startswith("#") and value.endswith("!"):
            return CellValue(type=CellType.ERROR, value=value)

        # Formulas — calamine generally evaluates formulas and returns
        # computed values, but if a string starts with = it's a formula
        if value.startswith("="):
            return CellValue(type=CellType.FORMULA, value=value, formula=value)

        return CellValue(type=CellType.STRING, value=value)

    # Fallback
    return CellValue(type=CellType.STRING, value=str(value))


class CalamineAdapter(ReadOnlyAdapter):
    """Adapter for python-calamine library (read-only).

    python-calamine is a Rust-backed library that provides fast Excel reading.
    It evaluates formulas and returns computed values, so formula text is NOT
    preserved. Formatting information is also not available.
    """

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="python-calamine",
            version=_get_version(),
            language="python",
            capabilities={"read"},
        )

    @property
    def supported_read_extensions(self) -> set[str]:
        return {".xlsx", ".xls"}

    # =========================================================================
    # Read Operations
    # =========================================================================

    def open_workbook(self, path: Path) -> CalamineWorkbook:
        return CalamineWorkbook.from_path(str(path))

    def close_workbook(self, workbook: Any) -> None:
        pass  # No close needed

    def get_sheet_names(self, workbook: CalamineWorkbook) -> list[str]:
        return workbook.sheet_names

    def read_cell_value(
        self,
        workbook: CalamineWorkbook,
        sheet: str,
        cell: str,
    ) -> CellValue:
        sheet_data = workbook.get_sheet_by_name(sheet)
        rows = sheet_data.to_python()
        row_idx, col_idx = _parse_cell_ref(cell)

        # Out of bounds → blank
        if row_idx >= len(rows):
            return CellValue(type=CellType.BLANK)
        row = rows[row_idx]
        if col_idx >= len(row):
            return CellValue(type=CellType.BLANK)

        return _convert_value(row[col_idx])

    def read_sheet_values(
        self,
        workbook: CalamineWorkbook,
        sheet: str,
        cell_range: str | None = None,
    ) -> list[list[CellValue]]:
        """Bulk read all values from a sheet (or a rectangular sub-range).

        Optional helper used by performance workloads.  Calls to_python()
        once and converts the entire grid, avoiding the per-cell overhead
        of read_cell_value().
        """
        sheet_data = workbook.get_sheet_by_name(sheet)
        rows = sheet_data.to_python()

        if cell_range:
            clean = cell_range.replace("$", "").upper()
            if ":" in clean:
                a, b = clean.split(":", 1)
            else:
                a, b = clean, clean
            r0, c0 = _parse_cell_ref(a)
            r1, c1 = _parse_cell_ref(b)
            if r1 < r0:
                r0, r1 = r1, r0
            if c1 < c0:
                c0, c1 = c1, c0
            rows = [
                row[c0: c1 + 1]
                for row in rows[r0: r1 + 1]
                if len(row) > c0
            ]

        return [[_convert_value(v) for v in row] for row in rows]

    def read_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        return CellFormat()  # No formatting support

    def read_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        return BorderInfo()  # No border support

    def read_row_height(
        self,
        workbook: Any,
        sheet: str,
        row: int,
    ) -> float | None:
        return None  # Not available

    def read_column_width(
        self,
        workbook: Any,
        sheet: str,
        column: str,
    ) -> float | None:
        return None  # Not available

    # =========================================================================
    # Tier 2 Read Operations
    # =========================================================================

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        return []

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_data_validations(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_images(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_comments(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_freeze_panes(self, workbook: Any, sheet: str) -> JSONDict:
        return {}
