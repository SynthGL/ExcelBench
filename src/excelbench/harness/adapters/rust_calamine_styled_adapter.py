"""Adapter for calamine (styled) via excelbench_rust (PyO3).

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

# Excel OOXML stores column widths with font-metric padding included.
# Standard paddings for common default fonts:
_CALIBRI_WIDTH_PADDING = 0.83203125   # Calibri 11pt (most .xlsx files)
_ALT_WIDTH_PADDING = 0.7109375        # Some alternate fonts
_WIDTH_TOLERANCE = 0.0005

try:
    import excelbench_rust as _excelbench_rust
except ImportError as e:  # pragma: no cover
    raise ImportError("excelbench_rust calamine-styled backend unavailable") from e

if getattr(_excelbench_rust, "CalamineStyledBook", None) is None:  # pragma: no cover
    raise ImportError("excelbench_rust built without calamine (styled) backend")


class RustCalamineStyledAdapter(ReadOnlyAdapter):
    """Adapter for the Rust calamine crate (with style support) via PyO3."""

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
        import excelbench_rust

        return excelbench_rust.CalamineStyledBook.open(str(path))

    def close_workbook(self, workbook: Any) -> None:
        return

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return [str(name) for name in workbook.sheet_names()]

    def read_cell_value(self, workbook: Any, sheet: str, cell: str) -> CellValue:
        payload = workbook.read_cell_value(sheet, cell)
        if not isinstance(payload, dict):
            return CellValue(type=CellType.STRING, value=str(payload))
        return cell_value_from_payload(payload)

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
        return workbook.read_row_height(sheet, row)

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        raw = workbook.read_column_width(sheet, column)
        if raw is None:
            return None
        # Strip Excel's font-metric padding from the OOXML stored width.
        frac = raw % 1
        for padding in (_CALIBRI_WIDTH_PADDING, _ALT_WIDTH_PADDING):
            if abs(frac - padding) < _WIDTH_TOLERANCE:
                adjusted = raw - padding
                if adjusted >= 0:
                    raw = adjusted
                break
        return round(raw, 4)

    # =========================================================================
    # Tier 2 Read Operations — not yet implemented in the Rust binding
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
