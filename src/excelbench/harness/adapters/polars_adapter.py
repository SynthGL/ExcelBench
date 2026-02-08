"""Adapter for polars library (read-only, value-only).

polars is a Rust-backed DataFrame library that uses calamine internally
for .xlsx reads.  Like pandas, it exposes cell values only — all
formatting is lost in the DataFrame abstraction.  Additionally, polars'
columnar model forces mixed-type columns to a common supertype (usually
String), so we must parse string representations back to native types.

This adapter measures the abstraction cost of polars vs. both pandas
and the raw calamine adapter.
"""

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import polars as pl

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
    try:
        from importlib.metadata import version

        return version("polars")
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


class PolarsAdapter(ReadOnlyAdapter):
    """Adapter for polars library (read-only, value-only).

    polars wraps calamine (Rust) behind ``pl.read_excel``.  Mixed-type
    columns are coerced to String, so we parse values back to native
    Python types for cell-level comparison.
    """

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="polars",
            version=_get_version(),
            language="python",
            capabilities={"read"},
        )

    @property
    def supported_read_extensions(self) -> set[str]:
        return {".xlsx"}

    # =========================================================================
    # Read Operations
    # =========================================================================

    def open_workbook(self, path: Path) -> Any:
        frames = pl.read_excel(
            path,
            sheet_id=0,  # 0 = all sheets
            has_header=False,
            infer_schema_length=0,  # don't infer — keep as String
            raise_if_empty=False,
        )
        if isinstance(frames, pl.DataFrame):
            # Single sheet — wrap in dict
            frames = {"Sheet1": frames}
        return {"frames": frames, "path": path}

    def close_workbook(self, workbook: Any) -> None:
        pass

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return list(workbook["frames"].keys())

    def read_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellValue:
        row_idx, col_idx = _parse_cell_ref(cell)
        frames: dict[str, pl.DataFrame] = workbook["frames"]

        if sheet not in frames:
            return CellValue(type=CellType.BLANK)

        df = frames[sheet]
        if row_idx >= len(df) or col_idx >= len(df.columns):
            return CellValue(type=CellType.BLANK)

        value = df.item(row_idx, col_idx)

        if value is None:
            return CellValue(type=CellType.BLANK)

        # polars may return native types for homogeneous columns
        if isinstance(value, bool):
            return CellValue(type=CellType.BOOLEAN, value=value)

        if isinstance(value, (int, float)):
            if isinstance(value, float) and (value != value):  # NaN check
                return CellValue(type=CellType.BLANK)
            return CellValue(type=CellType.NUMBER, value=value)

        if isinstance(value, date) and not isinstance(value, datetime):
            return CellValue(type=CellType.DATE, value=value)

        if isinstance(value, datetime):
            if (
                value.hour == 0
                and value.minute == 0
                and value.second == 0
                and value.microsecond == 0
            ):
                return CellValue(type=CellType.DATE, value=value.date())
            return CellValue(type=CellType.DATETIME, value=value)

        # String path — parse back to native types
        if isinstance(value, str):
            if value == "":
                return CellValue(type=CellType.BLANK)

            # Error values
            if value in ("#N/A", "#NULL!", "#NAME?", "#REF!"):
                return CellValue(type=CellType.ERROR, value=value)
            if value.startswith("#") and value.endswith("!"):
                return CellValue(type=CellType.ERROR, value=value)

            # Formula
            if value.startswith("="):
                return CellValue(type=CellType.FORMULA, value=value, formula=value)

            # Boolean
            if value.lower() == "true":
                return CellValue(type=CellType.BOOLEAN, value=True)
            if value.lower() == "false":
                return CellValue(type=CellType.BOOLEAN, value=False)

            # Number
            try:
                num = float(value)
                if num == int(num) and "." not in value and "e" not in value.lower():
                    return CellValue(type=CellType.NUMBER, value=int(value))
                return CellValue(type=CellType.NUMBER, value=num)
            except ValueError:
                pass

            # Datetime (polars stringifies as "YYYY-MM-DD HH:MM:SS")
            try:
                dt = datetime.fromisoformat(value)
                if (
                    dt.hour == 0
                    and dt.minute == 0
                    and dt.second == 0
                    and dt.microsecond == 0
                ):
                    return CellValue(type=CellType.DATE, value=dt.date())
                return CellValue(type=CellType.DATETIME, value=dt)
            except ValueError:
                pass

            # Date only (YYYY-MM-DD)
            try:
                d = date.fromisoformat(value)
                return CellValue(type=CellType.DATE, value=d)
            except ValueError:
                pass

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
