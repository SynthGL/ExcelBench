"""Adapter for calamine (styled) via the Rust PyO3 extension.

This adapter exercises the Rust calamine crate through our CalamineStyledBook
PyO3 binding which includes style-aware reading (format, borders, dimensions).
It is read-only and supports .xlsx files only (the styled API requires the
Xlsx<R> reader, not the format-sniffing open_workbook_auto).
"""

from pathlib import Path
from typing import Any

from excelbench.harness.adapters.base import ReadOnlyAdapter
from excelbench.harness.adapters.rust_adapter_utils import (
    cell_value_from_payload,
    dict_to_border,
    dict_to_format,
    get_rust_backend_version,
)
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)

JSONDict = dict[str, Any]

try:
    import wolfxl._rust as _excelbench_rust
except ImportError as e:  # pragma: no cover
    raise ImportError("wolfxl._rust calamine-styled backend unavailable") from e

if getattr(_excelbench_rust, "CalamineStyledBook", None) is None:  # pragma: no cover
    raise ImportError("wolfxl._rust built without calamine (styled) backend")


class RustCalamineStyledAdapter(ReadOnlyAdapter):
    """Adapter for the Rust calamine crate (with style support) via PyO3."""

    def __init__(self) -> None:
        self._cell_cache: dict[tuple[int, str, str], CellValue] = {}

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="calamine-styled",
            version=get_rust_backend_version("calamine"),
            language="rust",
            capabilities={"read"},
        )

    @property
    def supported_read_extensions(self) -> set[str]:
        # CalamineStyledBook uses Xlsx<R> directly — only .xlsx supported.
        return {".xlsx"}

    def open_workbook(self, path: Path) -> Any:
        import wolfxl._rust as rust

        m: Any = rust
        cls = getattr(m, "CalamineStyledBook")
        return cls.open(str(path))

    def close_workbook(self, workbook: Any) -> None:
        wb_id = id(workbook)
        self._cell_cache = {k: v for k, v in self._cell_cache.items() if k[0] != wb_id}

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return [str(name) for name in workbook.sheet_names()]

    def read_cell_value(self, workbook: Any, sheet: str, cell: str) -> CellValue:
        key = (id(workbook), sheet, cell)
        cached = self._cell_cache.get(key)
        if cached is not None:
            return cached
        payload = workbook.read_cell_value(sheet, cell)
        if not isinstance(payload, dict):
            result = CellValue(type=CellType.STRING, value=str(payload))
        else:
            result = cell_value_from_payload(payload)
        self._cell_cache[key] = result
        return result

    def read_sheet_values(
        self,
        workbook: Any,
        sheet: str,
        cell_range: str | None = None,
    ) -> list[list[CellValue]]:
        """Bulk read all values from a sheet via CalamineStyledBook.read_sheet_values()."""
        raw = workbook.read_sheet_values(sheet, cell_range)
        return [
            [
                cell_value_from_payload(v)
                if isinstance(v, dict)
                else CellValue(type=CellType.BLANK)
                for v in row
            ]
            for row in raw
        ]

    def read_sheet_values_raw(
        self,
        workbook: Any,
        sheet: str,
        cell_range: str | None = None,
    ) -> list[list[Any]]:
        """Return raw Rust FFI output without cell_value_from_payload() wrapping."""
        result: list[list[Any]] = workbook.read_sheet_values(sheet, cell_range)
        return result

    def read_cell_format(self, workbook: Any, sheet: str, cell: str) -> CellFormat:
        payload = workbook.read_cell_format(sheet, cell)
        if not isinstance(payload, dict) or not payload:
            return CellFormat()
        return dict_to_format(payload)

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        payload = workbook.read_cell_border(sheet, cell)
        if not isinstance(payload, dict) or not payload:
            return BorderInfo()
        return dict_to_border(payload)

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        value = workbook.read_row_height(sheet, row)
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        return None

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        # Rust binding strips Excel padding before returning.
        value = workbook.read_column_width(sheet, column)
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        return None

    # =========================================================================
    # Tier 2 Read Operations
    # =========================================================================

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        result = workbook.read_merged_ranges(sheet)
        if isinstance(result, list):
            return [str(x) for x in result]
        return []

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_conditional_formats(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_data_validations(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_data_validations(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_named_ranges(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_named_ranges(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_tables(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_hyperlinks(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_images(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_comments(self, workbook: Any, sheet: str) -> list[JSONDict]:
        result = workbook.read_comments(sheet)
        if isinstance(result, list):
            return [dict(x) for x in result if isinstance(x, dict)]
        return []

    def read_freeze_panes(self, workbook: Any, sheet: str) -> JSONDict:
        return dict(workbook.read_freeze_panes(sheet))
