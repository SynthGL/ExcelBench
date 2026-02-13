"""Adapter for umya-spreadsheet via excelbench_rust (PyO3).

This adapter is read/write (xlsx). Supports Tier 0/1 cell values, formulas,
formatting (text, background, borders, alignment, number formats, dimensions),
Tier 2 features (merged cells, freeze panes, comments, hyperlinks, images,
data validation, conditional formatting), and Tier 3 features (named ranges,
tables).
"""

from pathlib import Path
from typing import Any

from excelbench.harness.adapters.base import ExcelAdapter
from excelbench.harness.adapters.rust_adapter_utils import (
    border_to_dict,
    cell_value_from_payload,
    dict_to_border,
    dict_to_format,
    format_to_dict,
    get_rust_backend_version,
    payload_from_cell_value,
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
    import excelbench_rust as _excelbench_rust
except ImportError as e:  # pragma: no cover
    raise ImportError("excelbench_rust umya backend unavailable") from e

if getattr(_excelbench_rust, "UmyaBook", None) is None:  # pragma: no cover
    raise ImportError("excelbench_rust built without umya backend")


class UmyaAdapter(ExcelAdapter):
    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="umya-spreadsheet",
            version=get_rust_backend_version("umya-spreadsheet"),
            language="rust",
            capabilities={"read", "write"},
        )

    @property
    def supported_read_extensions(self) -> set[str]:
        return {".xlsx"}

    # =========================================================================
    # Read
    # =========================================================================

    def open_workbook(self, path: Path) -> Any:
        import excelbench_rust

        m: Any = excelbench_rust
        cls = getattr(m, "UmyaBook")
        return cls.open(str(path))

    def close_workbook(self, workbook: Any) -> None:
        return

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return [str(n) for n in workbook.sheet_names()]

    def read_cell_value(self, workbook: Any, sheet: str, cell: str) -> CellValue:
        payload = workbook.read_cell_value(sheet, cell)
        if not isinstance(payload, dict):
            return CellValue(type=CellType.STRING, value=str(payload))
        return cell_value_from_payload(payload)

    def read_cell_format(self, workbook: Any, sheet: str, cell: str) -> CellFormat:
        d = workbook.read_cell_format(sheet, cell)
        if not isinstance(d, dict) or not d:
            return CellFormat()
        return dict_to_format(d)

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        d = workbook.read_cell_border(sheet, cell)
        if not isinstance(d, dict) or not d:
            return BorderInfo()
        return dict_to_border(d)

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        return workbook.read_row_height(sheet, row - 1)

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        return workbook.read_column_width(sheet, column)

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        return workbook.read_merged_ranges(sheet)

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return workbook.read_conditional_formats(sheet)

    def read_data_validations(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return workbook.read_data_validations(sheet)

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return workbook.read_hyperlinks(sheet)

    def read_images(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return workbook.read_images(sheet)

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_comments(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return workbook.read_comments(sheet)

    def read_freeze_panes(self, workbook: Any, sheet: str) -> JSONDict:
        return workbook.read_freeze_panes(sheet)

    def read_named_ranges(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return workbook.read_named_ranges(sheet)

    def read_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return workbook.read_tables(sheet)

    # =========================================================================
    # Write
    # =========================================================================

    def create_workbook(self) -> Any:
        import excelbench_rust

        m: Any = excelbench_rust
        cls = getattr(m, "UmyaBook")
        return cls()

    def add_sheet(self, workbook: Any, name: str) -> None:
        workbook.add_sheet(name)

    def write_cell_value(self, workbook: Any, sheet: str, cell: str, value: CellValue) -> None:
        workbook.write_cell_value(sheet, cell, payload_from_cell_value(value))

    def write_cell_format(self, workbook: Any, sheet: str, cell: str, format: CellFormat) -> None:
        d = format_to_dict(format)
        if d:
            workbook.write_cell_format(sheet, cell, d)

    def write_cell_border(self, workbook: Any, sheet: str, cell: str, border: BorderInfo) -> None:
        d = border_to_dict(border)
        if d:
            workbook.write_cell_border(sheet, cell, d)

    def set_row_height(self, workbook: Any, sheet: str, row: int, height: float) -> None:
        workbook.set_row_height(sheet, row - 1, height)

    def set_column_width(self, workbook: Any, sheet: str, column: str, width: float) -> None:
        workbook.set_column_width(sheet, column, width)

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        workbook.merge_cells(sheet, cell_range)

    def add_conditional_format(self, workbook: Any, sheet: str, rule: JSONDict) -> None:
        if not rule:
            return
        workbook.add_conditional_format(sheet, rule)

    def add_data_validation(self, workbook: Any, sheet: str, validation: JSONDict) -> None:
        if not validation:
            return
        workbook.add_data_validation(sheet, validation)

    def add_hyperlink(self, workbook: Any, sheet: str, link: JSONDict) -> None:
        if not link:
            return
        workbook.add_hyperlink(sheet, link)

    def add_image(self, workbook: Any, sheet: str, image: JSONDict) -> None:
        if not image:
            return
        workbook.add_image(sheet, image)

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: JSONDict) -> None:
        raise NotImplementedError("umya pivot tables not implemented")

    def add_comment(self, workbook: Any, sheet: str, comment: JSONDict) -> None:
        if not comment:
            return
        workbook.add_comment(sheet, comment)

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        if not settings:
            return
        workbook.set_freeze_panes(sheet, settings)

    def add_named_range(self, workbook: Any, sheet: str, named_range: JSONDict) -> None:
        if not named_range:
            return
        workbook.add_named_range(sheet, named_range)

    def add_table(self, workbook: Any, sheet: str, table: JSONDict) -> None:
        if not table:
            return
        workbook.add_table(sheet, table)

    def save_workbook(self, workbook: Any, path: Path) -> None:
        workbook.save(str(path))
