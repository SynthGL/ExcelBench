"""Adapter for xlrd library (read-only, .xls format only)."""

import re
from pathlib import Path
from typing import Any

import xlrd
from xlrd import Book
from xlrd.sheet import Sheet

from excelbench.harness.adapters.base import ReadOnlyAdapter
from excelbench.models import (
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)


def _get_version() -> str:
    return xlrd.__version__


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


# xlrd border style index → ExcelBench BorderStyle
_BORDER_STYLE_MAP: dict[int, BorderStyle] = {
    0: BorderStyle.NONE,
    1: BorderStyle.THIN,
    2: BorderStyle.MEDIUM,
    3: BorderStyle.DASHED,
    4: BorderStyle.DOTTED,
    5: BorderStyle.THICK,
    6: BorderStyle.DOUBLE,
    7: BorderStyle.HAIR,
    8: BorderStyle.MEDIUM_DASHED,
    9: BorderStyle.DASH_DOT,
    10: BorderStyle.MEDIUM_DASH_DOT,
    11: BorderStyle.DASH_DOT_DOT,
    12: BorderStyle.MEDIUM_DASH_DOT_DOT,
    13: BorderStyle.SLANT_DASH_DOT,
}

# xlrd horizontal alignment index → name
_H_ALIGN_MAP: dict[int, str] = {
    1: "left",
    2: "center",
    3: "right",
    4: "fill",
    5: "justify",
    6: "centerContinuous",
    7: "distributed",
}

# xlrd vertical alignment index → name
_V_ALIGN_MAP: dict[int, str] = {
    0: "top",
    1: "center",
    2: "bottom",
    3: "justify",
    4: "distributed",
}


def _color_to_hex(book: Book, colour_index: int) -> str | None:
    """Convert xlrd colour index to hex string."""
    if colour_index is None or colour_index == 0x7FFF or colour_index == 64:
        return None
    colour_map = book.colour_map
    if colour_index in colour_map:
        rgb = colour_map[colour_index]
        if rgb is not None:
            return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    return None


class XlrdAdapter(ReadOnlyAdapter):
    """Adapter for xlrd library (read-only, .xls format).

    xlrd >=2.0 only supports .xls (BIFF) format. It will raise an
    error when attempting to open .xlsx files.
    """

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="xlrd",
            version=_get_version(),
            language="python",
            capabilities={"read"},
        )

    # =========================================================================
    # Read Operations
    # =========================================================================

    def open_workbook(self, path: Path) -> Book:
        return xlrd.open_workbook(str(path), formatting_info=True)

    def close_workbook(self, workbook: Any) -> None:
        workbook.release_resources()

    def get_sheet_names(self, workbook: Book) -> list[str]:
        return workbook.sheet_names()

    def read_cell_value(
        self,
        workbook: Book,
        sheet: str,
        cell: str,
    ) -> CellValue:
        sh: Sheet = workbook.sheet_by_name(sheet)
        row_idx, col_idx = _parse_cell_ref(cell)

        if row_idx >= sh.nrows or col_idx >= sh.ncols:
            return CellValue(type=CellType.BLANK)

        cell_type = sh.cell_type(row_idx, col_idx)
        value = sh.cell_value(row_idx, col_idx)

        # xlrd cell types:
        # 0 = XL_CELL_EMPTY, 1 = XL_CELL_TEXT, 2 = XL_CELL_NUMBER
        # 3 = XL_CELL_DATE, 4 = XL_CELL_BOOLEAN, 5 = XL_CELL_ERROR, 6 = XL_CELL_BLANK
        if cell_type in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
            return CellValue(type=CellType.BLANK)

        if cell_type == xlrd.XL_CELL_TEXT:
            if isinstance(value, str):
                if value in ("#N/A", "#NULL!", "#NAME?", "#REF!"):
                    return CellValue(type=CellType.ERROR, value=value)
                if value.startswith("#") and value.endswith("!"):
                    return CellValue(type=CellType.ERROR, value=value)
                if value.startswith("="):
                    return CellValue(type=CellType.FORMULA, value=value, formula=value)
            return CellValue(type=CellType.STRING, value=value)

        if cell_type == xlrd.XL_CELL_NUMBER:
            return CellValue(type=CellType.NUMBER, value=value)

        if cell_type == xlrd.XL_CELL_DATE:
            from datetime import date, datetime

            date_tuple = xlrd.xldate_as_tuple(value, workbook.datemode)
            year, month, day, hour, minute, second = date_tuple
            if hour == 0 and minute == 0 and second == 0:
                return CellValue(type=CellType.DATE, value=date(year, month, day))
            return CellValue(
                type=CellType.DATETIME,
                value=datetime(year, month, day, hour, minute, second),
            )

        if cell_type == xlrd.XL_CELL_BOOLEAN:
            return CellValue(type=CellType.BOOLEAN, value=bool(value))

        if cell_type == xlrd.XL_CELL_ERROR:
            error_map = {
                0: "#NULL!",
                7: "#DIV/0!",
                15: "#VALUE!",
                23: "#REF!",
                29: "#NAME?",
                36: "#NUM!",
                42: "#N/A",
            }
            return CellValue(
                type=CellType.ERROR,
                value=error_map.get(int(value), f"#ERROR({value})"),
            )

        return CellValue(type=CellType.STRING, value=str(value))

    def read_cell_format(
        self,
        workbook: Book,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        sh = workbook.sheet_by_name(sheet)
        row_idx, col_idx = _parse_cell_ref(cell)

        if row_idx >= sh.nrows or col_idx >= sh.ncols:
            return CellFormat()

        xf_index = sh.cell_xf_index(row_idx, col_idx)
        xf = workbook.xf_list[xf_index]
        font = workbook.font_list[xf.font_index]

        # Font attributes
        bold = True if font.bold else None
        italic = True if font.italic else None
        strikethrough = True if font.struck_out else None
        font_name = font.name if font.name else None
        font_size = font.height / 20.0 if font.height else None

        underline = None
        if font.underline_type == 1:
            underline = "single"
        elif font.underline_type == 2:
            underline = "double"
        elif font.underline_type == 33:
            underline = "singleAccounting"
        elif font.underline_type == 34:
            underline = "doubleAccounting"

        font_color = _color_to_hex(workbook, font.colour_index)

        # Background color
        bg_color = None
        if xf.background and xf.background.pattern_colour_index:
            bg_color = _color_to_hex(workbook, xf.background.pattern_colour_index)

        # Number format
        number_format = None
        if xf.format_key in workbook.format_map:
            fmt = workbook.format_map[xf.format_key]
            if fmt.format_str and fmt.format_str != "General":
                number_format = fmt.format_str

        # Alignment
        align = xf.alignment
        h_align = _H_ALIGN_MAP.get(align.hor_align)
        v_align = _V_ALIGN_MAP.get(align.vert_align)
        wrap = True if align.text_wrapped else None
        rotation = align.rotation if align.rotation not in (0, 255) else None
        indent = align.indent_level if align.indent_level else None

        return CellFormat(
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            font_name=font_name,
            font_size=font_size,
            font_color=font_color,
            bg_color=bg_color,
            number_format=number_format,
            h_align=h_align,
            v_align=v_align,
            wrap=wrap,
            rotation=rotation,
            indent=indent,
        )

    def read_cell_border(
        self,
        workbook: Book,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        sh = workbook.sheet_by_name(sheet)
        row_idx, col_idx = _parse_cell_ref(cell)

        if row_idx >= sh.nrows or col_idx >= sh.ncols:
            return BorderInfo()

        xf_index = sh.cell_xf_index(row_idx, col_idx)
        xf = workbook.xf_list[xf_index]
        border = xf.border

        def make_edge(style_idx: int, colour_idx: int) -> BorderEdge | None:
            style = _BORDER_STYLE_MAP.get(style_idx, BorderStyle.NONE)
            if style == BorderStyle.NONE:
                return None
            color = _color_to_hex(workbook, colour_idx) or "#000000"
            return BorderEdge(style=style, color=color)

        return BorderInfo(
            top=make_edge(border.top_line_style, border.top_colour_index),
            bottom=make_edge(border.bottom_line_style, border.bottom_colour_index),
            left=make_edge(border.left_line_style, border.left_colour_index),
            right=make_edge(border.right_line_style, border.right_colour_index),
            diagonal_up=make_edge(border.diag_line_style, border.diag_colour_index)
            if border.diag_up
            else None,
            diagonal_down=make_edge(border.diag_line_style, border.diag_colour_index)
            if border.diag_down
            else None,
        )

    def read_row_height(
        self,
        workbook: Book,
        sheet: str,
        row: int,
    ) -> float | None:
        sh = workbook.sheet_by_name(sheet)
        row_idx = row - 1
        if row_idx >= sh.nrows:
            return None
        rowinfo = sh.rowinfo_map.get(row_idx)
        if rowinfo and rowinfo.height:
            return rowinfo.height / 20.0
        return None

    def read_column_width(
        self,
        workbook: Book,
        sheet: str,
        column: str,
    ) -> float | None:
        sh = workbook.sheet_by_name(sheet)
        col_idx = 0
        for char in column.upper():
            col_idx = col_idx * 26 + (ord(char) - ord("A") + 1)
        col_idx -= 1
        colinfo = sh.colinfo_map.get(col_idx)
        if colinfo and colinfo.width:
            return colinfo.width / 256.0
        return None

    # =========================================================================
    # Tier 2 Read Operations
    # =========================================================================

    def read_merged_ranges(self, workbook: Book, sheet: str) -> list[str]:
        sh = workbook.sheet_by_name(sheet)
        ranges: list[str] = []
        for rlo, rhi, clo, chi in sh.merged_cells:
            start = f"{_col_letter(clo + 1)}{rlo + 1}"
            end = f"{_col_letter(chi)}{rhi}"
            ranges.append(f"{start}:{end}")
        return ranges

    def read_conditional_formats(self, workbook: Book, sheet: str) -> list[dict]:
        return []  # xlrd has limited CF support

    def read_data_validations(self, workbook: Book, sheet: str) -> list[dict]:
        return []  # Not available in xlrd

    def read_hyperlinks(self, workbook: Book, sheet: str) -> list[dict]:
        sh = workbook.sheet_by_name(sheet)
        links: list[dict] = []
        for link in sh.hyperlink_list:
            cell = f"{_col_letter(link.fcolx + 1)}{link.frowx + 1}"
            target = link.url_or_path
            display = link.desc or sh.cell_value(link.frowx, link.fcolx)
            links.append({
                "cell": cell,
                "target": target,
                "display": display,
                "tooltip": link.textmark if link.textmark else None,
                "internal": bool(link.textmark and not link.url_or_path),
            })
        return links

    def read_images(self, workbook: Book, sheet: str) -> list[dict]:
        return []  # xlrd does not support image reading

    def read_pivot_tables(self, workbook: Book, sheet: str) -> list[dict]:
        return []  # xlrd does not support pivot table reading

    def read_comments(self, workbook: Book, sheet: str) -> list[dict]:
        sh = workbook.sheet_by_name(sheet)
        comments: list[dict] = []
        note_map = getattr(sh, "cell_note_map", {})
        for (row_idx, col_idx), note in note_map.items():
            cell = f"{_col_letter(col_idx + 1)}{row_idx + 1}"
            comments.append({
                "cell": cell,
                "text": note.text,
                "author": note.author,
                "threaded": False,
            })
        return comments

    def read_freeze_panes(self, workbook: Book, sheet: str) -> dict:
        sh = workbook.sheet_by_name(sheet)
        result: dict[str, Any] = {}
        if sh.frozen_row_count or sh.frozen_col_count:
            result["mode"] = "freeze"
            top_row = sh.frozen_row_count or 0
            left_col = sh.frozen_col_count or 0
            result["top_left_cell"] = f"{_col_letter(left_col + 1)}{top_row + 1}"
        return result


def _col_letter(index: int) -> str:
    """Convert 1-based column index to letter(s)."""
    result = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result
