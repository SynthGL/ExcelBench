"""Adapter for umya-spreadsheet via excelbench_rust (PyO3).

This adapter is read/write (xlsx).

Current scope: Tier 1 cell values + formulas, plus Tier 2 read/write operations
including merged cells, conditional formatting, data validation, hyperlinks,
images, comments, and freeze panes.
"""

from pathlib import Path
from typing import Any

from excelbench.harness.adapters.base import ExcelAdapter
from excelbench.harness.adapters.rust_adapter_utils import (
    cell_value_from_payload,
    get_rust_backend_version,
    payload_from_border_info,
    payload_from_cell_format,
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
        payload = workbook.read_cell_format(sheet, cell)
        if not isinstance(payload, dict):
            return CellFormat()
        return CellFormat(
            bold=payload.get("bold"),
            italic=payload.get("italic"),
            underline=payload.get("underline"),
            strikethrough=payload.get("strikethrough"),
            font_name=payload.get("font_name"),
            font_size=payload.get("font_size"),
            font_color=payload.get("font_color"),
            bg_color=payload.get("bg_color"),
            number_format=payload.get("number_format"),
            h_align=payload.get("h_align"),
            v_align=payload.get("v_align"),
            wrap=payload.get("wrap"),
            rotation=payload.get("rotation"),
            indent=payload.get("indent"),
        )

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        return BorderInfo()

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        return None

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        return None

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        return [str(r) for r in workbook.read_merged_ranges(sheet)]

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[JSONDict]:
        rules = workbook.read_conditional_formats(sheet)
        return list(rules) if isinstance(rules, list) else []

    def read_data_validations(self, workbook: Any, sheet: str) -> list[JSONDict]:
        vals = workbook.read_data_validations(sheet)
        return list(vals) if isinstance(vals, list) else []

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[JSONDict]:
        links = workbook.read_hyperlinks(sheet)
        return list(links) if isinstance(links, list) else []

    def read_images(self, workbook: Any, sheet: str) -> list[JSONDict]:
        images = workbook.read_images(sheet)
        return list(images) if isinstance(images, list) else []

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_comments(self, workbook: Any, sheet: str) -> list[JSONDict]:
        comments = workbook.read_comments(sheet)
        return list(comments) if isinstance(comments, list) else []

    def read_freeze_panes(self, workbook: Any, sheet: str) -> JSONDict:
        cfg = workbook.read_freeze_panes(sheet)
        return dict(cfg) if isinstance(cfg, dict) else {}

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
        workbook.write_cell_format(sheet, cell, payload_from_cell_format(format))

    def write_cell_border(self, workbook: Any, sheet: str, cell: str, border: BorderInfo) -> None:
        workbook.write_cell_border(sheet, cell, payload_from_border_info(border))

    def set_row_height(self, workbook: Any, sheet: str, row: int, height: float) -> None:
        workbook.set_row_height(sheet, row, height)

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

    def save_workbook(self, workbook: Any, path: Path) -> None:
        workbook.save(str(path))
