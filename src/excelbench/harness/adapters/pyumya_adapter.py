"""Adapter for pyumya (Rust-backed, read/write).

This adapter exists primarily to validate pyumya correctness against ExcelBench
fixtures as pyumya expands formatting coverage.
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any

import pyumya

from excelbench.harness.adapters.base import ExcelAdapter
from excelbench.models import (
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)

JSONDict = dict[str, Any]


def _get_version() -> str:
    try:
        from importlib.metadata import version

        return version("pyumya")
    except Exception:
        try:
            import pyumya._rust as _rust

            return str(getattr(_rust, "__version__", "unknown"))
        except Exception:
            return "unknown"


def _is_error_token(value: str) -> bool:
    return value == "#N/A" or (value.startswith("#") and value.endswith("!"))


def _to_rgb_no_hash(value: str) -> str:
    s = value.strip()
    if s.startswith("#"):
        s = s[1:]
    s = s.upper()
    if len(s) == 8:
        s = s[2:]
    return s


def _to_rgb_hash(value: str) -> str:
    s = _to_rgb_no_hash(value)
    if len(s) != 6:
        return "#000000"
    return f"#{s}"


def _border_style(style: str) -> BorderStyle:
    try:
        return BorderStyle(style)
    except Exception:
        return BorderStyle.NONE


class PyumyaAdapter(ExcelAdapter):
    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="pyumya",
            version=_get_version(),
            language="python",
            capabilities={"read", "write"},
        )

    # =========================================================================
    # Read
    # =========================================================================

    def open_workbook(self, path: Path) -> Any:
        return pyumya.load_workbook(str(path))

    def close_workbook(self, workbook: Any) -> None:
        # pyumya has no explicit close()
        return None

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return [str(n) for n in workbook.sheetnames]

    def read_cell_value(self, workbook: Any, sheet: str, cell: str) -> CellValue:
        ws = workbook[sheet]
        value = ws[cell].value

        if value is None:
            return CellValue(type=CellType.BLANK)

        if isinstance(value, bool):
            return CellValue(type=CellType.BOOLEAN, value=value)

        if isinstance(value, (int, float)):
            return CellValue(type=CellType.NUMBER, value=value)

        if isinstance(value, date) and not isinstance(value, datetime):
            return CellValue(type=CellType.DATE, value=value)

        if isinstance(value, datetime):
            return CellValue(type=CellType.DATETIME, value=value)

        if isinstance(value, str):
            if value.startswith("="):
                return CellValue(type=CellType.FORMULA, value=value, formula=value)
            if _is_error_token(value):
                return CellValue(type=CellType.ERROR, value=value)
            return CellValue(type=CellType.STRING, value=value)

        return CellValue(type=CellType.STRING, value=str(value))

    def read_cell_format(self, workbook: Any, sheet: str, cell: str) -> CellFormat:
        ws = workbook[sheet]
        c = ws[cell]

        font = c.font
        fill = c.fill
        align = c.alignment

        font_color = _to_rgb_hash(getattr(font, "color", "000000")) if font else None
        bg_color = _to_rgb_hash(getattr(fill, "fgColor", "000000")) if fill else None
        if fill and getattr(fill, "fill_type", "none") == "none":
            bg_color = None

        nf = getattr(c, "number_format", None)
        if isinstance(nf, str) and nf == "General":
            nf = None

        return CellFormat(
            bold=getattr(font, "bold", None),
            italic=getattr(font, "italic", None),
            underline=(
                getattr(font, "underline", None) if font and font.underline != "none" else None
            ),
            strikethrough=getattr(font, "strikethrough", None),
            font_name=getattr(font, "name", None),
            font_size=getattr(font, "size", None),
            font_color=font_color,
            bg_color=bg_color,
            number_format=nf,
            h_align=getattr(align, "horizontal", None),
            v_align=getattr(align, "vertical", None),
            wrap=getattr(align, "wrap_text", None),
            rotation=(
                getattr(align, "text_rotation", None)
                if getattr(align, "text_rotation", None) not in (0, None)
                else None
            ),
            indent=(
                getattr(align, "indent", None)
                if getattr(align, "indent", None) not in (0, None)
                else None
            ),
        )

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        ws = workbook[sheet]
        b = ws[cell].border

        def edge(side: Any) -> BorderEdge | None:
            if side is None:
                return None
            style = str(getattr(side, "style", "none"))
            if not style or style == "none":
                return None
            color = _to_rgb_hash(str(getattr(side, "color", "000000")))
            return BorderEdge(style=_border_style(style), color=color)

        diag_side = getattr(b, "diagonal", None)
        diag_edge = edge(diag_side)
        diag_up = bool(getattr(b, "diagonalUp", False))
        diag_down = bool(getattr(b, "diagonalDown", False))

        return BorderInfo(
            top=edge(getattr(b, "top", None)),
            bottom=edge(getattr(b, "bottom", None)),
            left=edge(getattr(b, "left", None)),
            right=edge(getattr(b, "right", None)),
            diagonal_up=diag_edge if diag_edge is not None and diag_up else None,
            diagonal_down=diag_edge if diag_edge is not None and diag_down else None,
        )

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        ws = workbook[sheet]
        v = ws.row_dimensions[row].height
        return None if v is None else float(v)

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        ws = workbook[sheet]
        v = ws.column_dimensions[column].width
        if v is None:
            return None
        try:
            width_f = float(v)
        except Exception:
            return None
        # Excel and third-party libraries add font-metric padding to stored
        # column widths (e.g. +0.83203125 for Calibri 11pt).
        frac = width_f % 1
        for padding in (0.83203125, 0.7109375):
            if abs(frac - padding) < 0.01:
                width_f = width_f - padding
                break
        return round(width_f, 4)

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        ws = workbook[sheet]
        return [str(r) for r in ws.merged_cells.ranges]

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
        ws = workbook[sheet]
        top_left = ws.freeze_panes
        if top_left:
            return {"mode": "freeze", "top_left_cell": str(top_left)}
        # pyumya currently doesn't expose split pane metadata.
        return {"mode": "none"}

    # =========================================================================
    # Write
    # =========================================================================

    def create_workbook(self) -> Any:
        # ExcelBench write harness expects to fully control sheet creation.
        # pyumya defaults to including a starter worksheet (openpyxl-style),
        # so request an empty workbook when supported.
        try:
            return pyumya.Workbook(remove_default_sheet=True)
        except TypeError:
            return pyumya.Workbook()

    def add_sheet(self, workbook: Any, name: str) -> None:
        workbook.create_sheet(name)

    def write_cell_value(self, workbook: Any, sheet: str, cell: str, value: CellValue) -> None:
        ws = workbook[sheet]
        c = ws[cell]

        if value.type == CellType.BLANK:
            c.value = None
        elif value.type == CellType.BOOLEAN:
            c.value = bool(value.value)
        elif value.type == CellType.NUMBER:
            c.value = value.value
        elif value.type == CellType.STRING:
            c.value = "" if value.value is None else str(value.value)
        elif value.type == CellType.ERROR:
            c.value = str(value.value)
        elif value.type == CellType.FORMULA:
            f = value.formula or value.value
            c.value = "" if f is None else str(f)
        elif value.type == CellType.DATE:
            c.value = value.value
        elif value.type == CellType.DATETIME:
            c.value = value.value
        else:
            c.value = "" if value.value is None else str(value.value)

    def write_cell_format(self, workbook: Any, sheet: str, cell: str, format: CellFormat) -> None:
        ws = workbook[sheet]
        c = ws[cell]

        # Font
        if any(
            v is not None
            for v in [
                format.bold,
                format.italic,
                format.underline,
                format.strikethrough,
                format.font_name,
                format.font_size,
                format.font_color,
            ]
        ):
            existing = c.font

            name = (
                format.font_name
                if format.font_name is not None
                else getattr(existing, "name", "Calibri")
            )
            size_obj = (
                format.font_size
                if format.font_size is not None
                else getattr(existing, "size", 11.0)
            )
            try:
                size = float(size_obj) if size_obj is not None else 11.0
            except Exception:
                size = 11.0

            color_obj = (
                format.font_color
                if format.font_color is not None
                else getattr(existing, "color", "000000")
            )

            c.font = pyumya.Font(
                name=str(name),
                size=size,
                bold=bool(
                    format.bold if format.bold is not None else getattr(existing, "bold", False)
                ),
                italic=bool(
                    format.italic
                    if format.italic is not None
                    else getattr(existing, "italic", False)
                ),
                underline=str(
                    format.underline
                    if format.underline is not None
                    else getattr(existing, "underline", "none")
                ),
                strikethrough=bool(
                    format.strikethrough
                    if format.strikethrough is not None
                    else getattr(existing, "strikethrough", False)
                ),
                color=_to_rgb_no_hash(str(color_obj)),
            )

        # Fill
        if format.bg_color is not None:
            c.fill = pyumya.PatternFill(fgColor=_to_rgb_no_hash(format.bg_color))

        # Number format
        if format.number_format is not None:
            c.number_format = str(format.number_format)

        # Alignment
        if any(
            v is not None
            for v in [format.h_align, format.v_align, format.wrap, format.rotation, format.indent]
        ):
            existing = c.alignment
            c.alignment = pyumya.Alignment(
                horizontal=str(format.h_align or getattr(existing, "horizontal", "general")),
                vertical=str(format.v_align or getattr(existing, "vertical", "bottom")),
                wrap_text=bool(
                    format.wrap
                    if format.wrap is not None
                    else getattr(existing, "wrap_text", False)
                ),
                text_rotation=int(
                    format.rotation
                    if format.rotation is not None
                    else getattr(existing, "text_rotation", 0)
                ),
                indent=int(
                    format.indent if format.indent is not None else getattr(existing, "indent", 0)
                ),
            )

    def write_cell_border(self, workbook: Any, sheet: str, cell: str, border: BorderInfo) -> None:
        ws = workbook[sheet]
        c = ws[cell]

        def side(edge: BorderEdge | None) -> pyumya.Side:
            if edge is None:
                return pyumya.Side()
            return pyumya.Side(style=edge.style.value, color=_to_rgb_no_hash(edge.color))

        diag = border.diagonal_up or border.diagonal_down
        diag_up = border.diagonal_up is not None
        diag_down = border.diagonal_down is not None
        c.border = pyumya.Border(
            left=side(border.left),
            right=side(border.right),
            top=side(border.top),
            bottom=side(border.bottom),
            diagonal=side(diag),
            diagonalUp=diag_up,
            diagonalDown=diag_down,
        )

    def set_row_height(self, workbook: Any, sheet: str, row: int, height: float) -> None:
        ws = workbook[sheet]
        ws.row_dimensions[row].height = float(height)

    def set_column_width(self, workbook: Any, sheet: str, column: str, width: float) -> None:
        ws = workbook[sheet]
        ws.column_dimensions[column].width = float(width)

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        ws = workbook[sheet]
        ws.merge_cells(str(cell_range))

    def add_conditional_format(self, workbook: Any, sheet: str, rule: JSONDict) -> None:
        return

    def add_data_validation(self, workbook: Any, sheet: str, validation: JSONDict) -> None:
        return

    def add_hyperlink(self, workbook: Any, sheet: str, link: JSONDict) -> None:
        return

    def add_image(self, workbook: Any, sheet: str, image: JSONDict) -> None:
        return

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: JSONDict) -> None:
        raise NotImplementedError("pyumya pivot tables not implemented")

    def add_comment(self, workbook: Any, sheet: str, comment: JSONDict) -> None:
        return

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        ws = workbook[sheet]
        cfg = settings.get("freeze", settings)
        mode = cfg.get("mode")
        if mode == "freeze":
            ws.freeze_panes = cfg.get("top_left_cell")
            return
        if mode == "split":
            # pyumya doesn't expose split pane metadata yet; treat as a no-op so
            # freeze-pane basics can still pass in ExcelBench.
            return
        ws.freeze_panes = None

    def save_workbook(self, workbook: Any, path: Path) -> None:
        workbook.save(str(path))
