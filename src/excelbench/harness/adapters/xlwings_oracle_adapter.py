"""Excel oracle adapter using xlwings (read-only)."""

from datetime import date, datetime
from pathlib import Path
from typing import Any

import xlwings as xw

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


class XlLineStyle:
    CONTINUOUS = 1
    DASH = -4115
    DASH_DOT = 4
    DASH_DOT_DOT = 5
    DOT = -4118
    DOUBLE = -4119
    NONE = -4142
    SLANT_DASH_DOT = 13


class XlBorderWeight:
    HAIRLINE = 1
    THIN = 2
    MEDIUM = -4138
    THICK = 4


class XlBordersIndex:
    EDGE_LEFT = 7
    EDGE_TOP = 8
    EDGE_BOTTOM = 9
    EDGE_RIGHT = 10
    DIAGONAL_DOWN = 5
    DIAGONAL_UP = 6


H_ALIGN_MAP = {
    -4131: "left",  # xlLeft
    -4108: "center",  # xlCenter
    -4152: "right",  # xlRight
    -4130: "justify",  # xlJustify
    -4117: "distributed",  # xlDistributed
    7: "centerContinuous",
    1: "general",
}

V_ALIGN_MAP = {
    -4160: "top",  # xlTop
    -4108: "center",  # xlCenter
    -4107: "bottom",  # xlBottom
    -4130: "justify",  # xlJustify
    -4117: "distributed",  # xlDistributed
}

MAC_H_ALIGN_MAP = {
    "k.horizontal_align_general": "general",
    "k.horizontal_align_left": "left",
    "k.horizontal_align_center": "center",
    "k.horizontal_align_right": "right",
    "k.horizontal_align_justify": "justify",
    "k.horizontal_align_distributed": "distributed",
    "k.horizontal_align_center_across_selection": "centerContinuous",
}

MAC_V_ALIGN_MAP = {
    "k.vertical_alignment_top": "top",
    "k.vertical_alignment_center": "center",
    "k.vertical_alignment_bottom": "bottom",
    "k.vertical_alignment_justify": "justify",
    "k.vertical_alignment_distributed": "distributed",
}


def _int_to_hex(color_int: int | None) -> str | None:
    if color_int is None:
        return None
    if isinstance(color_int, (list, tuple)) and len(color_int) == 3:
        r, g, b = color_int
        return f"#{int(r):02X}{int(g):02X}{int(b):02X}"
    if isinstance(color_int, float):
        color_int = int(color_int)
    if color_int < 0:
        return None
    r = color_int & 0xFF
    g = (color_int >> 8) & 0xFF
    b = (color_int >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


def _map_underline(value: object) -> str | None:
    if value in (None, False, 0):
        return None
    if value == 2:
        return "single"
    if value == -4119:
        return "double"
    if str(value) == "k.underline_style_single":
        return "single"
    if str(value) == "k.underline_style_double":
        return "double"
    if str(value) == "k.underline_style_none":
        return None
    return str(value)


def _cell_address(row: int, col: int) -> str:
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return f"{result}{row}"


class ExcelOracleAdapter(ReadOnlyAdapter):
    """Read-only adapter backed by Excel via xlwings."""

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="excel_oracle",
            version="excel",
            language="excel",
            capabilities={"read"},
        )

    def open_workbook(self, path: Path) -> Any:
        return xw.Book(str(path))

    def close_workbook(self, workbook: Any) -> None:
        try:
            workbook.close()
        except Exception:
            pass

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return [s.name for s in workbook.sheets]

    def read_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellValue:
        ws = workbook.sheets[sheet]
        rng = ws.range(cell)
        value = rng.value
        formula = rng.formula

        if value is None:
            return CellValue(type=CellType.BLANK)

        if isinstance(value, bool):
            return CellValue(type=CellType.BOOLEAN, value=value)

        if isinstance(value, (int, float)):
            return CellValue(type=CellType.NUMBER, value=value)

        if isinstance(value, date) and not isinstance(value, datetime):
            return CellValue(type=CellType.DATE, value=value)

        if isinstance(value, datetime):
            if (
                value.hour == 0
                and value.minute == 0
                and value.second == 0
                and value.microsecond == 0
            ):
                return CellValue(type=CellType.DATE, value=value.date())
            return CellValue(type=CellType.DATETIME, value=value)

        if isinstance(value, str):
            if value.startswith("#"):
                return CellValue(type=CellType.ERROR, value=value)
            if formula:
                return CellValue(type=CellType.FORMULA, value=value, formula=str(formula))
            return CellValue(type=CellType.STRING, value=value)

        if formula:
            return CellValue(type=CellType.FORMULA, value=value, formula=str(formula))

        return CellValue(type=CellType.STRING, value=str(value))

    def read_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        ws = workbook.sheets[sheet]
        rng = ws.range(cell)
        is_mac = False
        try:
            font = rng.api.Font
            font_color = (
                _int_to_hex(font.Color) if getattr(font, "Color", None) is not None else None
            )
            underline = _map_underline(getattr(font, "Underline", None))
            bold = bool(font.Bold) if font.Bold is not None else None
            italic = bool(font.Italic) if font.Italic is not None else None
            strikethrough = bool(font.Strikethrough) if font.Strikethrough is not None else None
            font_name = font.Name if font.Name else None
            font_size = font.Size if font.Size else None
        except Exception:
            is_mac = True
            font = rng.api.font_object
            font_color = _int_to_hex(font.color.get())
            underline = _map_underline(font.underline.get())
            bold = bool(font.bold.get())
            italic = bool(font.italic.get())
            strikethrough = bool(font.strikethrough.get())
            font_name = font.font_name.get()
            if str(font_name) == "k.missing_value":
                font_name = None
            font_size = font.font_size.get()

        bg_color = None
        if not is_mac:
            if (
                hasattr(rng.api, "Interior")
                and getattr(rng.api.Interior, "Color", None) is not None
            ):
                bg_color = _int_to_hex(rng.api.Interior.Color)
        else:
            try:
                bg_color = _int_to_hex(rng.api.interior_object.color.get())
            except Exception:
                bg_color = None

        if is_mac:
            h_align = MAC_H_ALIGN_MAP.get(str(rng.api.horizontal_alignment.get()))
            v_align = MAC_V_ALIGN_MAP.get(str(rng.api.vertical_alignment.get()))
            wrap = rng.api.wrap_text.get()
            rotation = rng.api.text_orientation.get()
            indent = rng.api.indent_level.get()
        else:
            h_align = H_ALIGN_MAP.get(getattr(rng.api, "HorizontalAlignment", None))
            v_align = V_ALIGN_MAP.get(getattr(rng.api, "VerticalAlignment", None))
            wrap = getattr(rng.api, "WrapText", None)
            rotation = getattr(rng.api, "Orientation", None)
            indent = getattr(rng.api, "IndentLevel", None)

        return CellFormat(
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            font_name=font_name,
            font_size=font_size,
            font_color=font_color,
            bg_color=bg_color,
            number_format=rng.number_format if hasattr(rng, "number_format") else None,
            h_align=h_align,
            v_align=v_align,
            wrap=bool(wrap) if wrap is not None else None,
            rotation=(int(rotation) if rotation not in (None, 0, "k.missing_value") else None),
            indent=int(indent) if indent not in (None, 0) else None,
        )

    def read_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        ws = workbook.sheets[sheet]
        rng = ws.range(cell)

        def parse_border(index: int) -> BorderEdge | None:
            border = rng.api.Borders(index)
            line_style = border.LineStyle
            if line_style in (None, 0, XlLineStyle.NONE):
                return None

            if line_style == XlLineStyle.DOUBLE:
                style = BorderStyle.DOUBLE
            elif line_style == XlLineStyle.DASH:
                style = BorderStyle.DASHED
            elif line_style == XlLineStyle.DOT:
                style = BorderStyle.DOTTED
            elif line_style == XlLineStyle.DASH_DOT:
                style = BorderStyle.DASH_DOT
            elif line_style == XlLineStyle.DASH_DOT_DOT:
                style = BorderStyle.DASH_DOT_DOT
            elif line_style == XlLineStyle.SLANT_DASH_DOT:
                style = BorderStyle.SLANT_DASH_DOT
            else:
                weight = border.Weight
                weight_map = {
                    XlBorderWeight.HAIRLINE: BorderStyle.HAIR,
                    XlBorderWeight.THIN: BorderStyle.THIN,
                    XlBorderWeight.MEDIUM: BorderStyle.MEDIUM,
                    XlBorderWeight.THICK: BorderStyle.THICK,
                }
                style = weight_map.get(weight, BorderStyle.THIN)

            color = _int_to_hex(border.Color) or "#000000"
            return BorderEdge(style=style, color=color)

        return BorderInfo(
            top=parse_border(XlBordersIndex.EDGE_TOP),
            bottom=parse_border(XlBordersIndex.EDGE_BOTTOM),
            left=parse_border(XlBordersIndex.EDGE_LEFT),
            right=parse_border(XlBordersIndex.EDGE_RIGHT),
            diagonal_up=parse_border(XlBordersIndex.DIAGONAL_UP),
            diagonal_down=parse_border(XlBordersIndex.DIAGONAL_DOWN),
        )

    def read_row_height(
        self,
        workbook: Any,
        sheet: str,
        row: int,
    ) -> float | None:
        ws = workbook.sheets[sheet]
        return ws.range(f"{row}:{row}").row_height

    def read_column_width(
        self,
        workbook: Any,
        sheet: str,
        column: str,
    ) -> float | None:
        ws = workbook.sheets[sheet]
        return ws.range(f"{column}:{column}").column_width

    # =========================================================================
    # Tier 2 Read Operations
    # =========================================================================

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        ws = workbook.sheets[sheet]
        used = ws.api.UsedRange
        rows = used.Rows.Count
        cols = used.Columns.Count
        start_row = used.Row
        start_col = used.Column
        merges: set[str] = set()
        for r in range(start_row, start_row + rows):
            for c in range(start_col, start_col + cols):
                cell = ws.api.Cells(r, c)
                try:
                    if cell.MergeCells:
                        area = cell.MergeArea
                        addr = area.Address(False, False)
                        merges.add(addr)
                except Exception:
                    continue
        return sorted(merges)

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[dict]:
        ws = workbook.sheets[sheet]
        rules: list[dict] = []
        try:
            fc = ws.api.UsedRange.FormatConditions
            count = fc.Count
        except Exception:
            return rules
        type_map = {
            1: "cellIs",
            2: "expression",
            3: "colorScale",
            4: "dataBar",
            6: "iconSet",
        }
        op_map = {
            5: "greaterThan",
            6: "greaterThanOrEqual",
            7: "lessThan",
            8: "lessThanOrEqual",
            3: "equal",
            4: "notEqual",
            1: "between",
            2: "notBetween",
        }
        for i in range(1, count + 1):
            rule = fc.Item(i)
            entry = {
                "range": rule.AppliesTo.Address(False, False),
                "rule_type": type_map.get(rule.Type, str(rule.Type)),
                "operator": op_map.get(getattr(rule, "Operator", None), None),
                "formula": getattr(rule, "Formula1", None),
                "priority": getattr(rule, "Priority", None),
                "stop_if_true": bool(getattr(rule, "StopIfTrue", None))
                if getattr(rule, "StopIfTrue", None) is not None
                else None,
                "format": {},
            }
            try:
                color = _int_to_hex(rule.Interior.Color)
                if color:
                    entry["format"]["bg_color"] = color
            except Exception:
                pass
            rules.append(entry)
        return rules

    def read_data_validations(self, workbook: Any, sheet: str) -> list[dict]:
        ws = workbook.sheets[sheet]
        used = ws.api.UsedRange
        rows = used.Rows.Count
        cols = used.Columns.Count
        start_row = used.Row
        start_col = used.Column
        validations: list[dict] = []
        seen: set[tuple] = set()
        for r in range(start_row, start_row + rows):
            for c in range(start_col, start_col + cols):
                cell = ws.api.Cells(r, c)
                try:
                    val = cell.Validation
                    if val is None or val.Type in (0, None):
                        continue
                    key = (cell.Address(False, False), val.Type, val.Formula1, val.Formula2)
                    if key in seen:
                        continue
                    seen.add(key)
                    validations.append(
                        {
                            "range": cell.Address(False, False),
                            "validation_type": _excel_validation_type(val.Type),
                            "operator": _excel_validation_operator(val.Operator),
                            "formula1": val.Formula1,
                            "formula2": val.Formula2,
                            "allow_blank": bool(val.IgnoreBlank)
                            if val.IgnoreBlank is not None
                            else None,
                            "show_input": bool(val.ShowInput)
                            if val.ShowInput is not None
                            else None,
                            "show_error": bool(val.ShowError)
                            if val.ShowError is not None
                            else None,
                            "prompt_title": val.InputTitle,
                            "prompt": val.InputMessage,
                            "error_title": val.ErrorTitle,
                            "error": val.ErrorMessage,
                        }
                    )
                except Exception:
                    continue
        return validations

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[dict]:
        ws = workbook.sheets[sheet]
        links: list[dict] = []
        try:
            hyperlinks = ws.api.Hyperlinks
            for i in range(1, hyperlinks.Count + 1):
                link = hyperlinks.Item(i)
                address = link.Address or link.SubAddress
                links.append(
                    {
                        "cell": link.Range.Address(False, False),
                        "target": address,
                        "display": link.TextToDisplay,
                        "tooltip": link.ScreenTip,
                        "internal": bool(link.SubAddress),
                    }
                )
        except Exception:
            pass
        return links

    def read_images(self, workbook: Any, sheet: str) -> list[dict]:
        ws = workbook.sheets[sheet]
        images: list[dict] = []
        try:
            shapes = ws.api.Shapes
            for i in range(1, shapes.Count + 1):
                shape = shapes.Item(i)
                cell = shape.TopLeftCell.Address(False, False)
                images.append(
                    {
                        "cell": cell,
                        "path": None,
                        "anchor": "twoCell",
                        "offset": [
                            int(shape.Left - shape.TopLeftCell.Left),
                            int(shape.Top - shape.TopLeftCell.Top),
                        ],
                        "alt_text": getattr(shape, "AlternativeText", None),
                    }
                )
        except Exception:
            pass
        return images

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[dict]:
        ws = workbook.sheets[sheet]
        pivots: list[dict] = []
        try:
            pts = ws.api.PivotTables()
            for i in range(1, pts.Count + 1):
                pt = pts.Item(i)
                row_fields = []
                col_fields = []
                data_fields = []
                page_fields = []
                try:
                    rf = pt.RowFields()
                    row_fields = [rf.Item(j).Name for j in range(1, rf.Count + 1)]
                except Exception:
                    pass
                try:
                    cf = pt.ColumnFields()
                    col_fields = [cf.Item(j).Name for j in range(1, cf.Count + 1)]
                except Exception:
                    pass
                try:
                    df = pt.DataFields()
                    data_fields = [df.Item(j).Name for j in range(1, df.Count + 1)]
                except Exception:
                    pass
                try:
                    pf = pt.PageFields()
                    page_fields = [pf.Item(j).Name for j in range(1, pf.Count + 1)]
                except Exception:
                    pass
                grouped = None
                try:
                    grouped = bool(pt.PivotFields("Date").Grouped)
                except Exception:
                    grouped = None
                pivots.append(
                    {
                        "name": pt.Name,
                        "source_range": pt.SourceData,
                        "target_cell": pt.TableRange2.Address(False, False),
                        "row_fields": row_fields,
                        "column_fields": col_fields,
                        "data_fields": data_fields,
                        "filter_fields": page_fields,
                        "grouped": grouped,
                    }
                )
        except Exception:
            pass
        return pivots

    def read_comments(self, workbook: Any, sheet: str) -> list[dict]:
        ws = workbook.sheets[sheet]
        comments: list[dict] = []
        try:
            legacy = ws.api.Comments
            for i in range(1, legacy.Count + 1):
                c = legacy.Item(i)
                comments.append(
                    {
                        "cell": c.Parent.Address(False, False),
                        "text": c.Text(),
                        "author": c.Author,
                        "threaded": False,
                    }
                )
        except Exception:
            pass
        try:
            threaded = ws.api.CommentsThreaded
            for i in range(1, threaded.Count + 1):
                c = threaded.Item(i)
                comments.append(
                    {
                        "cell": c.Parent.Address(False, False),
                        "text": c.Text,
                        "author": c.Author.Name if hasattr(c.Author, "Name") else None,
                        "threaded": True,
                    }
                )
        except Exception:
            pass
        return comments

    def read_freeze_panes(self, workbook: Any, sheet: str) -> dict:
        ws = workbook.sheets[sheet]
        try:
            ws.api.Activate()
            window = workbook.app.api.ActiveWindow
            if window.FreezePanes:
                return {
                    "mode": "freeze",
                    "top_left_cell": _cell_address(window.ScrollRow, window.ScrollColumn),
                }
            if window.SplitRow or window.SplitColumn:
                return {
                    "mode": "split",
                    "x_split": int(window.SplitColumn),
                    "y_split": int(window.SplitRow),
                }
        except Exception:
            return {}
        return {}


def _excel_validation_type(code: int | None) -> str | None:
    mapping = {
        1: "whole",
        2: "decimal",
        3: "list",
        4: "date",
        5: "time",
        6: "textLength",
        7: "custom",
    }
    return mapping.get(code)


def _excel_validation_operator(code: int | None) -> str | None:
    mapping = {
        1: "between",
        2: "notBetween",
        3: "equal",
        4: "notEqual",
        5: "greaterThan",
        6: "lessThan",
        7: "greaterThanOrEqual",
        8: "lessThanOrEqual",
    }
    return mapping.get(code)
