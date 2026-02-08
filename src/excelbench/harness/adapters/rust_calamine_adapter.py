"""Adapter for calamine via excelbench_rust (PyO3).

This adapter exercises the Rust calamine crate through our local PyO3 extension
module. It is read-only.
"""

from pathlib import Path
from typing import Any

from excelbench.harness.adapters.base import ReadOnlyAdapter
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)

from excelbench.harness.adapters.rust_adapter_utils import (
    cell_value_from_payload,
    get_rust_backend_version,
)

JSONDict = dict[str, Any]


# Optional dependency guard: ensure the extension exists and the backend was
# compiled in (feature-flagged builds can omit individual backends).
try:
    import excelbench_rust as _excelbench_rust
except ImportError as e:  # pragma: no cover
    raise ImportError("excelbench_rust calamine backend unavailable") from e

if getattr(_excelbench_rust, "CalamineBook", None) is None:  # pragma: no cover
    raise ImportError("excelbench_rust built without calamine backend")


class RustCalamineAdapter(ReadOnlyAdapter):
    """Adapter for the Rust calamine crate via our PyO3 module."""

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="calamine",
            version=get_rust_backend_version("calamine"),
            language="rust",
            capabilities={"read"},
        )

    @property
    def supported_read_extensions(self) -> set[str]:
        return {".xlsx", ".xls"}

    def open_workbook(self, path: Path) -> Any:
        import excelbench_rust

        m: Any = excelbench_rust
        book_cls = getattr(m, "CalamineBook")
        return book_cls.open(str(path))

    def close_workbook(self, workbook: Any) -> None:
        # No explicit close needed for the current binding.
        return

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return [str(name) for name in workbook.sheet_names()]

    def read_cell_value(self, workbook: Any, sheet: str, cell: str) -> CellValue:
        payload = workbook.read_cell_value(sheet, cell)
        if not isinstance(payload, dict):
            return CellValue(type=CellType.STRING, value=str(payload))
        return cell_value_from_payload(payload)

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
