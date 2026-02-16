"""wolfxl — hybrid Rust adapter: calamine (read) + rust_xlsxwriter (write).

Combines the fastest Rust Excel reader (calamine with style support) and the
fastest Rust writer (rust_xlsxwriter) into a single full-fidelity R+W adapter.
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
    raise ImportError("excelbench_rust unavailable — wolfxl requires it") from e

if getattr(_excelbench_rust, "CalamineStyledBook", None) is None:  # pragma: no cover
    raise ImportError("excelbench_rust built without calamine backend")
if getattr(_excelbench_rust, "RustXlsxWriterBook", None) is None:  # pragma: no cover
    raise ImportError("excelbench_rust built without rust_xlsxwriter backend")


class WolfxlAdapter(ExcelAdapter):
    """Hybrid adapter: calamine-styled reads + rust_xlsxwriter writes."""

    def __init__(self) -> None:
        # Python-side cell cache: avoids FFI on repeated reads of the same cell.
        # Keyed by (workbook_id, sheet, cell) → CellValue.
        self._cell_cache: dict[tuple[int, str, str], CellValue] = {}

    @property
    def info(self) -> LibraryInfo:
        cal_ver = get_rust_backend_version("calamine")
        rxw_ver = get_rust_backend_version("rust_xlsxwriter")
        return LibraryInfo(
            name="wolfxl",
            version=f"cal={cal_ver}+rxw={rxw_ver}",
            language="rust",
            capabilities={"read", "write"},
        )

    @property
    def supported_read_extensions(self) -> set[str]:
        return {".xlsx"}

    # =========================================================================
    # Read — delegates to CalamineStyledBook
    # =========================================================================

    def open_workbook(self, path: Path) -> Any:
        import excelbench_rust

        m: Any = excelbench_rust
        return getattr(m, "CalamineStyledBook").open(str(path))

    def close_workbook(self, workbook: Any) -> None:
        # Evict cached cells for this workbook.
        wb_id = id(workbook)
        self._cell_cache = {
            k: v for k, v in self._cell_cache.items() if k[0] != wb_id
        }

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
            [cell_value_from_payload(v) if isinstance(v, dict) else CellValue(type=CellType.BLANK)
             for v in row]
            for row in raw
        ]

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
        value = workbook.read_column_width(sheet, column)
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        return None

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

    # =========================================================================
    # Write — delegates to RustXlsxWriterBook
    # =========================================================================

    def create_workbook(self) -> Any:
        import excelbench_rust

        m: Any = excelbench_rust
        return getattr(m, "RustXlsxWriterBook")()

    def add_sheet(self, workbook: Any, name: str) -> None:
        workbook.add_sheet(name)

    def write_cell_value(self, workbook: Any, sheet: str, cell: str, value: CellValue) -> None:
        payload = payload_from_cell_value(value)
        workbook.write_cell_value(sheet, cell, payload)

    def write_sheet_values(
        self,
        workbook: Any,
        sheet: str,
        start_cell: str,
        values: list[list[Any]],
    ) -> None:
        """Bulk write a grid of values via RustXlsxWriterBook.write_sheet_values()."""
        workbook.write_sheet_values(sheet, start_cell, values)

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
        raise NotImplementedError("wolfxl pivot tables not implemented")

    def add_comment(self, workbook: Any, sheet: str, comment: JSONDict) -> None:
        workbook.add_comment(sheet, comment)

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        workbook.set_freeze_panes(sheet, settings)

    def save_workbook(self, workbook: Any, path: Path) -> None:
        workbook.save(str(path))
