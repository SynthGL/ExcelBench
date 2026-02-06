"""Adapter for xlwt library (write-only, .xls BIFF8 format)."""

import re
from datetime import date as _date
from datetime import datetime as _datetime
from pathlib import Path
from typing import Any

import xlwt

from excelbench.harness.adapters.base import WriteOnlyAdapter
from excelbench.models import (
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)


def _get_version() -> str:
    try:
        from importlib.metadata import version
        return version("xlwt")
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


def _col_to_index(column: str) -> int:
    col = 0
    for char in column.upper():
        col = col * 26 + (ord(char) - ord("A") + 1)
    return col - 1


# xlwt border style constants
_BORDER_MAP: dict[BorderStyle, int] = {
    BorderStyle.NONE: xlwt.Borders.NO_LINE,
    BorderStyle.THIN: xlwt.Borders.THIN,
    BorderStyle.MEDIUM: xlwt.Borders.MEDIUM,
    BorderStyle.DASHED: xlwt.Borders.DASHED,
    BorderStyle.DOTTED: xlwt.Borders.DOTTED,
    BorderStyle.THICK: xlwt.Borders.THICK,
    BorderStyle.DOUBLE: xlwt.Borders.DOUBLE,
    BorderStyle.HAIR: xlwt.Borders.HAIR,
    BorderStyle.MEDIUM_DASHED: xlwt.Borders.MEDIUM_DASHED,
    BorderStyle.DASH_DOT: xlwt.Borders.THIN_DASH_DOTTED,
    BorderStyle.MEDIUM_DASH_DOT: xlwt.Borders.MEDIUM_DASH_DOTTED,
    BorderStyle.DASH_DOT_DOT: xlwt.Borders.THIN_DASH_DOT_DOTTED,
    BorderStyle.MEDIUM_DASH_DOT_DOT: xlwt.Borders.MEDIUM_DASH_DOT_DOTTED,
    BorderStyle.SLANT_DASH_DOT: xlwt.Borders.SLANTED_MEDIUM_DASH_DOTTED,
}

# Map hex colors to xlwt palette indices (56-color palette)
_COLOUR_MAP: dict[str, int] = {}
for _name, _idx in xlwt.Style.colour_map.items():
    _COLOUR_MAP[_name] = _idx


def _hex_to_xlwt_colour(hex_color: str) -> int:
    """Map a hex color string to the nearest xlwt palette index.

    xlwt uses a 56-color indexed palette inherited from BIFF8.
    We find the closest match by Euclidean distance in RGB space.
    """
    hex_color = hex_color.lstrip("#").upper()
    if len(hex_color) != 6:
        return 0x40  # Default (black)

    target_r = int(hex_color[0:2], 16)
    target_g = int(hex_color[2:4], 16)
    target_b = int(hex_color[4:6], 16)

    # Known hex→palette shortcuts for exact matches
    exact: dict[str, str] = {
        "FF0000": "red",
        "00FF00": "green",
        "0000FF": "blue",
        "FFFF00": "yellow",
        "FF00FF": "pink",
        "00FFFF": "turquoise",
        "FFFFFF": "white",
        "000000": "black",
        "808080": "gray50",
        "C0C0C0": "gray25",
        "800000": "dark_red",
        "008000": "dark_green",
        "000080": "dark_blue",
        "808000": "olive_green",
        "800080": "purple_ega",
        "008080": "dark_teal",
        "FFA500": "orange",
        "FFC0CB": "rose",
        "ADD8E6": "light_blue",
    }
    short = hex_color.upper()
    if short in exact and exact[short] in _COLOUR_MAP:
        return _COLOUR_MAP[exact[short]]

    # Brute-force nearest colour in palette
    # xlwt.Style.colour_map maps name→index; we need index→RGB from
    # the default palette.  xlwt doesn't expose RGB directly, but the
    # BIFF8 default palette is well-known.  For simplicity, return the
    # closest named colour by trying the standard set.
    palette_rgb: dict[str, tuple[int, int, int]] = {
        "black": (0, 0, 0),
        "white": (255, 255, 255),
        "red": (255, 0, 0),
        "green": (0, 255, 0),
        "blue": (0, 0, 255),
        "yellow": (255, 255, 0),
        "pink": (255, 0, 255),
        "turquoise": (0, 255, 255),
        "dark_red": (128, 0, 0),
        "dark_green": (0, 128, 0),
        "dark_blue": (0, 0, 128),
        "olive_green": (128, 128, 0),
        "purple_ega": (128, 0, 128),
        "dark_teal": (0, 128, 128),
        "gray50": (128, 128, 128),
        "gray25": (192, 192, 192),
        "orange": (255, 165, 0),
        "coral": (255, 128, 128),
        "light_blue": (173, 216, 230),
        "light_green": (204, 255, 204),
        "light_yellow": (255, 255, 153),
        "sky_blue": (0, 204, 255),
        "rose": (255, 153, 204),
        "tan": (255, 204, 153),
        "periwinkle": (153, 153, 255),
        "ice_blue": (204, 255, 255),
        "ivory": (255, 255, 204),
        "lavender": (204, 153, 255),
        "gold": (255, 204, 0),
        "aqua": (51, 204, 204),
        "lime": (153, 204, 0),
        "plum": (153, 51, 102),
        "indigo": (51, 51, 153),
        "ocean_blue": (51, 102, 255),
        "brown": (153, 51, 0),
        "dark_purple": (51, 51, 153),
        "teal": (0, 128, 128),
        "gray80": (51, 51, 51),
        "gray40": (150, 150, 150),
    }

    best_name = "black"
    best_dist = float("inf")
    for name, (r, g, b) in palette_rgb.items():
        dist = (target_r - r) ** 2 + (target_g - g) ** 2 + (target_b - b) ** 2
        if dist < best_dist:
            best_dist = dist
            best_name = name
    return _COLOUR_MAP.get(best_name, 0x40)


# Horizontal alignment name → xlwt constant
_H_ALIGN_MAP: dict[str, int] = {
    "general": xlwt.Alignment.HORZ_GENERAL,
    "left": xlwt.Alignment.HORZ_LEFT,
    "center": xlwt.Alignment.HORZ_CENTER,
    "right": xlwt.Alignment.HORZ_RIGHT,
    "fill": xlwt.Alignment.HORZ_FILLED,
    "justify": xlwt.Alignment.HORZ_JUSTIFIED,
    "centerContinuous": xlwt.Alignment.HORZ_CENTER_ACROSS_SEL,
    "distributed": xlwt.Alignment.HORZ_DISTRIBUTED,
}

# Vertical alignment name → xlwt constant
_V_ALIGN_MAP: dict[str, int] = {
    "top": xlwt.Alignment.VERT_TOP,
    "center": xlwt.Alignment.VERT_CENTER,
    "bottom": xlwt.Alignment.VERT_BOTTOM,
    "justify": xlwt.Alignment.VERT_JUSTIFIED,
    "distributed": xlwt.Alignment.VERT_DISTRIBUTED,
}


class XlwtAdapter(WriteOnlyAdapter):
    """Adapter for xlwt library (write-only, .xls BIFF8 format).

    xlwt can write values, text formatting, borders, merged cells,
    freeze panes, row heights, and column widths.  It cannot write
    conditional formatting, data validation, hyperlinks, images,
    comments, or pivot tables.
    """

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="xlwt",
            version=_get_version(),
            language="python",
            capabilities={"write"},
        )

    @property
    def output_extension(self) -> str:
        return ".xls"

    # =========================================================================
    # Write Operations
    # =========================================================================

    def create_workbook(self) -> xlwt.Workbook:
        return xlwt.Workbook()

    def add_sheet(self, workbook: xlwt.Workbook, name: str) -> None:
        workbook.add_sheet(name, cell_overwrite_ok=True)

    def _get_sheet(self, workbook: xlwt.Workbook, name: str) -> Any:
        return workbook.get_sheet(name)

    def write_cell_value(
        self,
        workbook: xlwt.Workbook,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        ws = self._get_sheet(workbook, sheet)
        row, col = _parse_cell_ref(cell)

        if value.type == CellType.BLANK:
            ws.write(row, col, "")
        elif value.type == CellType.FORMULA:
            formula_str = value.formula or value.value or ""
            if formula_str.startswith("="):
                formula_str = formula_str[1:]
            ws.write(row, col, xlwt.Formula(formula_str))
        elif value.type == CellType.BOOLEAN:
            ws.write(row, col, bool(value.value))
        elif value.type == CellType.NUMBER:
            ws.write(row, col, value.value)
        elif value.type == CellType.DATE:
            dt = value.value
            if isinstance(dt, _date) and not isinstance(dt, _datetime):
                dt = _datetime.combine(dt, _datetime.min.time())
            style = xlwt.XFStyle()
            style.num_format_str = "YYYY-MM-DD"
            ws.write(row, col, dt, style)
        elif value.type == CellType.DATETIME:
            style = xlwt.XFStyle()
            style.num_format_str = "YYYY-MM-DD HH:MM:SS"
            ws.write(row, col, value.value, style)
        elif value.type == CellType.ERROR:
            ws.write(row, col, str(value.value))
        else:
            ws.write(row, col, str(value.value) if value.value is not None else "")

    def write_cell_format(
        self,
        workbook: xlwt.Workbook,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        ws = self._get_sheet(workbook, sheet)
        row, col = _parse_cell_ref(cell)

        style = self._build_style(format)

        # Read existing value (xlwt doesn't expose cell values once written,
        # so we re-write with the style; value may be lost if written before format).
        # For the benchmark, write_cell_value is always called before write_cell_format,
        # so we read the pending value from the worksheet row data.
        existing = ""
        try:
            existing_row = ws._Worksheet__rows.get(row)
            if existing_row:
                existing_cell = existing_row._Row__cells.get(col)
                if existing_cell:
                    # Re-write with new style
                    ws.write(row, col, existing_cell.number, style)
                    return
        except (AttributeError, KeyError):
            pass

        ws.write(row, col, existing, style)

    def write_cell_border(
        self,
        workbook: xlwt.Workbook,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        ws = self._get_sheet(workbook, sheet)
        row, col = _parse_cell_ref(cell)

        style = xlwt.XFStyle()
        borders = xlwt.Borders()

        if border.top:
            borders.top = _BORDER_MAP.get(border.top.style, xlwt.Borders.THIN)
            borders.top_colour = _hex_to_xlwt_colour(border.top.color)
        if border.bottom:
            borders.bottom = _BORDER_MAP.get(border.bottom.style, xlwt.Borders.THIN)
            borders.bottom_colour = _hex_to_xlwt_colour(border.bottom.color)
        if border.left:
            borders.left = _BORDER_MAP.get(border.left.style, xlwt.Borders.THIN)
            borders.left_colour = _hex_to_xlwt_colour(border.left.color)
        if border.right:
            borders.right = _BORDER_MAP.get(border.right.style, xlwt.Borders.THIN)
            borders.right_colour = _hex_to_xlwt_colour(border.right.color)
        if border.diagonal_up or border.diagonal_down:
            diag = border.diagonal_up or border.diagonal_down
            borders.diag = _BORDER_MAP.get(diag.style, xlwt.Borders.THIN)
            borders.diag_colour = _hex_to_xlwt_colour(diag.color)
            if border.diagonal_up:
                borders.need_diag1 = xlwt.Borders.NEED_DIAG1
            if border.diagonal_down:
                borders.need_diag2 = xlwt.Borders.NEED_DIAG2

        style.borders = borders

        # Re-write existing value with border style
        existing = ""
        try:
            existing_row = ws._Worksheet__rows.get(row)
            if existing_row:
                existing_cell = existing_row._Row__cells.get(col)
                if existing_cell:
                    ws.write(row, col, existing_cell.number, style)
                    return
        except (AttributeError, KeyError):
            pass
        ws.write(row, col, existing, style)

    def _build_style(self, fmt: CellFormat) -> xlwt.XFStyle:
        style = xlwt.XFStyle()
        font = xlwt.Font()

        if fmt.bold:
            font.bold = True
        if fmt.italic:
            font.italic = True
        if fmt.underline:
            underline_map = {
                "single": xlwt.Font.UNDERLINE_SINGLE,
                "double": xlwt.Font.UNDERLINE_DOUBLE,
                "singleAccounting": xlwt.Font.UNDERLINE_SINGLE_ACC,
                "doubleAccounting": xlwt.Font.UNDERLINE_DOUBLE_ACC,
            }
            font.underline = underline_map.get(fmt.underline, xlwt.Font.UNDERLINE_SINGLE)
        if fmt.strikethrough:
            font.struck_out = True
        if fmt.font_name:
            font.name = fmt.font_name
        if fmt.font_size:
            font.height = int(fmt.font_size * 20)
        if fmt.font_color:
            font.colour_index = _hex_to_xlwt_colour(fmt.font_color)

        style.font = font

        if fmt.bg_color:
            pattern = xlwt.Pattern()
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN
            pattern.pattern_fore_colour = _hex_to_xlwt_colour(fmt.bg_color)
            style.pattern = pattern

        if fmt.number_format:
            style.num_format_str = fmt.number_format

        alignment = xlwt.Alignment()
        if fmt.h_align:
            alignment.horz = _H_ALIGN_MAP.get(fmt.h_align, xlwt.Alignment.HORZ_GENERAL)
        if fmt.v_align:
            alignment.vert = _V_ALIGN_MAP.get(fmt.v_align, xlwt.Alignment.VERT_BOTTOM)
        if fmt.wrap:
            alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
        if fmt.rotation is not None:
            alignment.rota = fmt.rotation
        if fmt.indent is not None:
            alignment.inde = fmt.indent
        style.alignment = alignment

        return style

    def set_row_height(
        self,
        workbook: xlwt.Workbook,
        sheet: str,
        row: int,
        height: float,
    ) -> None:
        ws = self._get_sheet(workbook, sheet)
        # xlwt uses twips (1/20 of a point) for row height
        ws.row(row - 1).height_mismatch = True
        ws.row(row - 1).height = int(height * 20)

    def set_column_width(
        self,
        workbook: xlwt.Workbook,
        sheet: str,
        column: str,
        width: float,
    ) -> None:
        ws = self._get_sheet(workbook, sheet)
        col_idx = _col_to_index(column)
        # xlwt uses 1/256th of zero-character width
        ws.col(col_idx).width = int(width * 256)

    # =========================================================================
    # Tier 2 Write Operations
    # =========================================================================

    def merge_cells(self, workbook: xlwt.Workbook, sheet: str, cell_range: str) -> None:
        ws = self._get_sheet(workbook, sheet)
        start, end = cell_range.replace("$", "").split(":")
        r1, c1 = _parse_cell_ref(start)
        r2, c2 = _parse_cell_ref(end)
        ws.write_merge(r1, r2, c1, c2, "")

    def add_conditional_format(self, workbook: Any, sheet: str, rule: dict) -> None:
        pass  # xlwt does not support conditional formatting

    def add_data_validation(self, workbook: Any, sheet: str, validation: dict) -> None:
        pass  # xlwt does not support data validation

    def add_hyperlink(self, workbook: Any, sheet: str, link: dict) -> None:
        pass  # xlwt does not support hyperlinks via write_url

    def add_image(self, workbook: Any, sheet: str, image: dict) -> None:
        pass  # xlwt does not support images

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: dict) -> None:
        pass  # xlwt does not support pivot tables

    def add_comment(self, workbook: Any, sheet: str, comment: dict) -> None:
        pass  # xlwt does not support comments

    def set_freeze_panes(self, workbook: xlwt.Workbook, sheet: str, settings: dict) -> None:
        ws = self._get_sheet(workbook, sheet)
        cfg = settings.get("freeze", settings)
        mode = cfg.get("mode")
        if mode == "freeze" and cfg.get("top_left_cell"):
            row, col = _parse_cell_ref(cfg["top_left_cell"])
            ws.set_panes_frozen(True)
            ws.set_horz_split_pos(row)
            ws.set_vert_split_pos(col)

    def save_workbook(self, workbook: xlwt.Workbook, path: Path) -> None:
        workbook.save(str(path))
