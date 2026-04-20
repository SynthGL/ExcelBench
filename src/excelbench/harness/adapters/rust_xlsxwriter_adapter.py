"""Adapter for rust_xlsxwriter via the Rust PyO3 extension.

Supports Tier 0/1 cell values, formulas, and formatting (text, background,
borders, alignment, number formats, dimensions).
"""

from pathlib import Path
from typing import Any

from excelbench.harness.adapters.base import WriteOnlyAdapter
from excelbench.harness.adapters.rust_adapter_utils import (
    border_to_dict,
    format_to_dict,
    get_rust_backend_version,
    payload_from_cell_value,
)
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellValue,
    LibraryInfo,
)

JSONDict = dict[str, Any]


# Optional dependency guard: if the extension module is missing or compiled
# without the rust_xlsxwriter backend, importing this adapter should fail so
# that the adapter registry can skip it.
try:
    import wolfxl._rust as _excelbench_rust
except ImportError as e:  # pragma: no cover
    raise ImportError("wolfxl._rust rust_xlsxwriter backend unavailable") from e

if getattr(_excelbench_rust, "RustXlsxWriterBook", None) is None:  # pragma: no cover
    raise ImportError("wolfxl._rust built without rust_xlsxwriter backend")


class RustXlsxWriterAdapter(WriteOnlyAdapter):
    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="rust_xlsxwriter",
            version=get_rust_backend_version("rust_xlsxwriter"),
            language="rust",
            capabilities={"write"},
        )

    def create_workbook(self) -> Any:
        import wolfxl._rust as rust

        m: Any = rust
        cls = getattr(m, "RustXlsxWriterBook")
        return cls()

    def add_sheet(self, workbook: Any, name: str) -> None:
        workbook.add_sheet(name)

    def write_cell_value(self, workbook: Any, sheet: str, cell: str, value: CellValue) -> None:
        payload = payload_from_cell_value(value)
        workbook.write_cell_value(sheet, cell, payload)

    def write_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        d = format_to_dict(format)
        if d:
            workbook.write_cell_format(sheet, cell, d)

    def write_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        d = border_to_dict(border)
        if d:
            workbook.write_cell_border(sheet, cell, d)

    def write_sheet_values(
        self,
        workbook: Any,
        sheet: str,
        start_cell: str,
        values: list[list[Any]],
    ) -> None:
        """Bulk write a grid of values via RustXlsxWriterBook.write_sheet_values()."""
        workbook.write_sheet_values(sheet, start_cell, values)

    def set_row_height(self, workbook: Any, sheet: str, row: int, height: float) -> None:
        workbook.set_row_height(sheet, row, height)

    def set_column_width(self, workbook: Any, sheet: str, column: str, width: float) -> None:
        workbook.set_column_width(sheet, column, width)

    # =========================================================================
    # Tier 2 Write Operations
    # =========================================================================

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        workbook.merge_cells(sheet, cell_range)

    def add_conditional_format(self, workbook: Any, sheet: str, rule: JSONDict) -> None:
        workbook.add_conditional_format(sheet, rule)

    def add_data_validation(self, workbook: Any, sheet: str, validation: JSONDict) -> None:
        workbook.add_data_validation(sheet, validation)

    def add_named_range(self, workbook: Any, sheet: str, named_range: JSONDict) -> None:
        workbook.add_named_range(sheet, named_range)

    def add_table(self, workbook: Any, sheet: str, table: JSONDict) -> None:
        workbook.add_table(sheet, table)

    def add_hyperlink(self, workbook: Any, sheet: str, link: JSONDict) -> None:
        workbook.add_hyperlink(sheet, link)

    def add_image(self, workbook: Any, sheet: str, image: JSONDict) -> None:
        return

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: JSONDict) -> None:
        raise NotImplementedError("rust_xlsxwriter pivot tables not implemented")

    def add_comment(self, workbook: Any, sheet: str, comment: JSONDict) -> None:
        workbook.add_comment(sheet, comment)

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        workbook.set_freeze_panes(sheet, settings)

    def save_workbook(self, workbook: Any, path: Path) -> None:
        workbook.save(str(path))
