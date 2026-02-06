"""Adapter for openpyxl library."""

from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.datavalidation import DataValidation

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

# Formulas that produce known error values (openpyxl returns formula, not error)
ERROR_FORMULA_MAP = {
    "=1/0": "#DIV/0!",
    "=NA()": "#N/A",
    '="text"+1': "#VALUE!",
}


def _col_letter(index: int) -> str:
    result = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def _get_version() -> str:
    """Get openpyxl version."""
    return openpyxl.__version__


class OpenpyxlAdapter(ExcelAdapter):
    """Adapter for openpyxl library (read/write support)."""

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="openpyxl",
            version=_get_version(),
            language="python",
            capabilities={"read", "write"},
        )

    # =========================================================================
    # Read Operations
    # =========================================================================

    def open_workbook(self, path: Path) -> Workbook:
        """Open a workbook for reading."""
        return openpyxl.load_workbook(str(path), data_only=False)

    def close_workbook(self, workbook: Any) -> None:
        """Close an opened workbook."""
        workbook.close()

    def get_sheet_names(self, workbook: Workbook) -> list[str]:
        """Get list of sheet names in a workbook."""
        return workbook.sheetnames

    def read_cell_value(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
    ) -> CellValue:
        """Read the value of a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]

        # Handle different value types
        value = c.value

        if value is None:
            return CellValue(type=CellType.BLANK)

        if isinstance(value, bool):
            return CellValue(type=CellType.BOOLEAN, value=value)

        if isinstance(value, (int, float)):
            return CellValue(type=CellType.NUMBER, value=value)

        # Check date BEFORE datetime since datetime is a subclass of date
        if isinstance(value, date) and not isinstance(value, datetime):
            return CellValue(type=CellType.DATE, value=value)

        if isinstance(value, datetime):
            # Check if this is a "date" (time component is midnight)
            # Excel stores dates as datetimes with 00:00:00 time
            if (
                value.hour == 0
                and value.minute == 0
                and value.second == 0
                and value.microsecond == 0
            ):
                return CellValue(type=CellType.DATE, value=value.date())
            return CellValue(type=CellType.DATETIME, value=value)

        if isinstance(value, str):
            # Check if it's an error value
            if value.startswith("#") and value.endswith("!"):
                return CellValue(type=CellType.ERROR, value=value)

            # Check if there's a formula
            if c.data_type == "f" or (hasattr(c, "value") and str(c.value).startswith("=")):
                # Check if this formula produces a known error value
                formula_str = str(c.value)
                if formula_str and not formula_str.startswith("="):
                    formula_str = f"={formula_str}"
                if not formula_str.startswith("=") and value:
                    formula_str = str(value)
                if formula_str in ERROR_FORMULA_MAP:
                    return CellValue(type=CellType.ERROR, value=ERROR_FORMULA_MAP[formula_str])
                return CellValue(
                    type=CellType.FORMULA,
                    value=value,
                    formula=formula_str,
                )

            return CellValue(type=CellType.STRING, value=value)

        # Fallback to string
        return CellValue(type=CellType.STRING, value=str(value))

    def read_cell_format(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        """Read the formatting of a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]
        font: Font = c.font

        # Convert color to hex
        font_color = None
        if font.color and font.color.rgb:
            rgb = font.color.rgb
            if isinstance(rgb, str) and len(rgb) >= 6:
                # Handle ARGB format (8 chars) or RGB format (6 chars)
                if len(rgb) == 8:
                    font_color = f"#{rgb[2:]}"  # Skip alpha
                else:
                    font_color = f"#{rgb}"

        # Get background color
        bg_color = None
        fill = c.fill
        if fill and fill.patternType == "solid" and fill.fgColor and fill.fgColor.rgb:
            rgb = fill.fgColor.rgb
            if isinstance(rgb, str) and len(rgb) >= 6:
                if len(rgb) == 8:
                    bg_color = f"#{rgb[2:]}"
                else:
                    bg_color = f"#{rgb}"

        # Map underline
        underline = None
        if font.underline:
            underline_map = {
                "single": "single",
                "double": "double",
                "singleAccounting": "singleAccounting",
                "doubleAccounting": "doubleAccounting",
            }
            underline = underline_map.get(font.underline, font.underline)

        alignment = c.alignment
        h_align = alignment.horizontal if alignment and alignment.horizontal else None
        v_align = alignment.vertical if alignment and alignment.vertical else None
        wrap = alignment.wrap_text if alignment and alignment.wrap_text else None
        rotation = (
            alignment.text_rotation
            if alignment and alignment.text_rotation not in (0, None)
            else None
        )
        indent = alignment.indent if alignment and alignment.indent else None

        return CellFormat(
            bold=font.bold if font.bold else None,
            italic=font.italic if font.italic else None,
            underline=underline,
            strikethrough=font.strike if font.strike else None,
            font_name=font.name if font.name else None,
            font_size=font.size if font.size else None,
            font_color=font_color,
            bg_color=bg_color,
            number_format=c.number_format if c.number_format else None,
            h_align=h_align,
            v_align=v_align,
            wrap=wrap,
            rotation=rotation,
            indent=indent,
        )

    def read_cell_border(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        """Read the border information of a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]
        border: Border = c.border

        def parse_side(side: Side | None) -> BorderEdge | None:
            if side is None or side.style is None:
                return None

            # Map openpyxl style to our style
            style_map = {
                "thin": BorderStyle.THIN,
                "medium": BorderStyle.MEDIUM,
                "thick": BorderStyle.THICK,
                "double": BorderStyle.DOUBLE,
                "dashed": BorderStyle.DASHED,
                "dotted": BorderStyle.DOTTED,
                "hair": BorderStyle.HAIR,
                "mediumDashed": BorderStyle.MEDIUM_DASHED,
                "dashDot": BorderStyle.DASH_DOT,
                "mediumDashDot": BorderStyle.MEDIUM_DASH_DOT,
                "dashDotDot": BorderStyle.DASH_DOT_DOT,
                "mediumDashDotDot": BorderStyle.MEDIUM_DASH_DOT_DOT,
                "slantDashDot": BorderStyle.SLANT_DASH_DOT,
            }

            style = style_map.get(side.style, BorderStyle.THIN)

            # Get color
            color = "#000000"
            if side.color and side.color.rgb:
                rgb = side.color.rgb
                if isinstance(rgb, str) and len(rgb) >= 6:
                    if len(rgb) == 8:
                        color = f"#{rgb[2:]}"
                    else:
                        color = f"#{rgb}"

            return BorderEdge(style=style, color=color)

        return BorderInfo(
            top=parse_side(border.top),
            bottom=parse_side(border.bottom),
            left=parse_side(border.left),
            right=parse_side(border.right),
            diagonal_up=parse_side(border.diagonal) if border.diagonalUp else None,
            diagonal_down=parse_side(border.diagonal) if border.diagonalDown else None,
        )

    def read_row_height(
        self,
        workbook: Workbook,
        sheet: str,
        row: int,
    ) -> float | None:
        ws = workbook[sheet]
        return ws.row_dimensions[row].height

    def read_column_width(
        self,
        workbook: Workbook,
        sheet: str,
        column: str,
    ) -> float | None:
        ws = workbook[sheet]
        width = ws.column_dimensions[column].width
        if width is None:
            return None
        # Excel stores column width as display_width + 0.83203125 padding
        # (for Calibri 11pt default font). Strip the padding to return
        # the display character width that matches Excel's UI value.
        frac = width % 1
        if abs(frac - 0.83203125) < 0.001:
            width = width - 0.83203125
        return width

    # =========================================================================
    # Tier 2 Read Operations
    # =========================================================================

    def read_merged_ranges(self, workbook: Workbook, sheet: str) -> list[str]:
        ws = workbook[sheet]
        return [str(rng) for rng in ws.merged_cells.ranges]

    def read_conditional_formats(self, workbook: Workbook, sheet: str) -> list[dict]:
        ws = workbook[sheet]
        rules: list[dict] = []
        cf_rules = getattr(ws.conditional_formatting, "_cf_rules", {})
        for sqref, rule_list in cf_rules.items():
            range_value = None
            if hasattr(sqref, "sqref"):
                range_value = str(sqref.sqref)
            else:
                range_value = str(sqref)
                if range_value.startswith("<ConditionalFormatting"):
                    range_value = (
                        range_value.replace("<ConditionalFormatting", "").replace(">", "").strip()
                    )
            for rule in rule_list:
                entry: dict[str, Any] = {
                    "range": range_value,
                    "rule_type": getattr(rule, "type", None),
                    "operator": getattr(rule, "operator", None),
                    "formula": None,
                    "priority": getattr(rule, "priority", None),
                    "stop_if_true": getattr(rule, "stopIfTrue", None),
                    "format": {},
                }
                if getattr(rule, "formula", None):
                    entry["formula"] = rule.formula[0] if rule.formula else None
                dxf = getattr(rule, "dxf", None)
                if dxf and dxf.fill and dxf.fill.fgColor and dxf.fill.fgColor.rgb:
                    rgb = dxf.fill.fgColor.rgb
                    if len(rgb) == 8:
                        entry["format"]["bg_color"] = f"#{rgb[2:]}"
                    elif len(rgb) == 6:
                        entry["format"]["bg_color"] = f"#{rgb}"
                if dxf and dxf.font and dxf.font.color and dxf.font.color.rgb:
                    rgb = dxf.font.color.rgb
                    if len(rgb) == 8:
                        entry["format"]["font_color"] = f"#{rgb[2:]}"
                    elif len(rgb) == 6:
                        entry["format"]["font_color"] = f"#{rgb}"
                rules.append(entry)
        return rules

    def read_data_validations(self, workbook: Workbook, sheet: str) -> list[dict]:
        ws = workbook[sheet]
        validations: list[dict] = []
        dv = getattr(ws, "data_validations", None)
        if not dv:
            return validations
        for entry in dv.dataValidation:
            validations.append(
                {
                    "range": str(entry.sqref),
                    "validation_type": entry.type,
                    "operator": entry.operator,
                    "formula1": entry.formula1,
                    "formula2": entry.formula2,
                    "allow_blank": entry.allow_blank,
                    "show_input": entry.showInputMessage,
                    "show_error": entry.showErrorMessage,
                    "prompt_title": entry.promptTitle,
                    "prompt": entry.prompt,
                    "error_title": entry.errorTitle,
                    "error": entry.error,
                }
            )
        return validations

    def read_hyperlinks(self, workbook: Workbook, sheet: str) -> list[dict]:
        ws = workbook[sheet]
        links: list[dict] = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.hyperlink:
                    target = cell.hyperlink.target or cell.hyperlink.location
                    links.append(
                        {
                            "cell": cell.coordinate,
                            "target": target,
                            "display": cell.value,
                            "tooltip": cell.hyperlink.tooltip,
                            "internal": bool(cell.hyperlink.location) if cell.hyperlink else False,
                        }
                    )
        return links

    def read_images(self, workbook: Workbook, sheet: str) -> list[dict]:
        ws = workbook[sheet]
        images: list[dict] = []
        for img in getattr(ws, "_images", []):
            anchor = getattr(img, "anchor", None)
            anchor_type = None
            cell = None
            offset = None
            if isinstance(anchor, str):
                anchor_type = "oneCell"
                cell = anchor
            if anchor is not None and hasattr(anchor, "_from"):
                anchor_type = "oneCell"
                cell = f"{_col_letter(anchor._from.col + 1)}{anchor._from.row + 1}"
                offset = [anchor._from.colOff, anchor._from.rowOff]
            if anchor is not None and hasattr(anchor, "_to"):
                anchor_type = "twoCell"
            images.append(
                {
                    "cell": cell,
                    "path": getattr(img, "path", None) or getattr(img, "_path", None),
                    "anchor": anchor_type,
                    "offset": offset,
                    "alt_text": getattr(img, "title", None),
                }
            )
        return images

    def read_pivot_tables(self, workbook: Workbook, sheet: str) -> list[dict]:
        ws = workbook[sheet]
        pivots: list[dict] = []
        pivot_list = getattr(ws, "_pivots", []) or []
        for pivot in pivot_list:
            pivots.append(
                {
                    "name": getattr(pivot, "name", None),
                    "source_range": getattr(pivot, "cache", None).cacheSource
                    if getattr(pivot, "cache", None)
                    else None,
                    "target_cell": getattr(pivot, "location", None),
                }
            )
        return pivots

    def read_comments(self, workbook: Workbook, sheet: str) -> list[dict]:
        ws = workbook[sheet]
        comments: list[dict] = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.comment:
                    comments.append(
                        {
                            "cell": cell.coordinate,
                            "text": cell.comment.text,
                            "author": cell.comment.author,
                            "threaded": False,
                        }
                    )
        return comments

    def read_freeze_panes(self, workbook: Workbook, sheet: str) -> dict:
        ws = workbook[sheet]
        result: dict[str, Any] = {}
        if ws.freeze_panes:
            result["mode"] = "freeze"
            result["top_left_cell"] = (
                ws.freeze_panes.coordinate
                if hasattr(ws.freeze_panes, "coordinate")
                else str(ws.freeze_panes)
            )
        pane = getattr(ws.sheet_view, "pane", None)
        if pane and pane.state == "split" and (pane.xSplit or pane.ySplit):
            result["mode"] = "split"
            result["x_split"] = int(pane.xSplit) if pane.xSplit is not None else None
            result["y_split"] = int(pane.ySplit) if pane.ySplit is not None else None
            if pane.topLeftCell:
                result["top_left_cell"] = pane.topLeftCell
            if pane.activePane:
                result["active_pane"] = pane.activePane
        return result

    # =========================================================================
    # Write Operations
    # =========================================================================

    def create_workbook(self) -> Workbook:
        """Create a new workbook."""
        wb = Workbook()
        # Remove default sheet to allow explicit sheet creation
        if wb.sheetnames:
            default_sheet = wb.active
            wb.remove(default_sheet)
        return wb

    def add_sheet(self, workbook: Workbook, name: str) -> None:
        """Add a new sheet to a workbook."""
        workbook.create_sheet(name)

    def write_cell_value(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        """Write a value to a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]

        if value.type == CellType.BLANK:
            c.value = None
        elif value.type == CellType.FORMULA:
            c.value = value.formula or value.value
        elif value.type == CellType.ERROR:
            # Write a formula that produces the error
            error_formulas = {
                "#DIV/0!": "=1/0",
                "#N/A": "=NA()",
                "#VALUE!": '="text"+1',
                "#REF!": "=#REF!",
                "#NAME?": "=_undefined_name_",
                "#NUM!": "=SQRT(-1)",
                "#NULL!": "=A1:A2 B1:B2",
            }
            c.value = error_formulas.get(value.value, value.value)
        else:
            c.value = value.value

    def write_cell_format(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        """Apply formatting to a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]

        # Build font kwargs
        font_kwargs = {}

        if format.bold is not None:
            font_kwargs["bold"] = format.bold
        if format.italic is not None:
            font_kwargs["italic"] = format.italic
        if format.underline is not None:
            font_kwargs["underline"] = format.underline
        if format.strikethrough is not None:
            font_kwargs["strike"] = format.strikethrough
        if format.font_name is not None:
            font_kwargs["name"] = format.font_name
        if format.font_size is not None:
            font_kwargs["size"] = format.font_size
        if format.font_color is not None:
            from openpyxl.styles import Color

            # Remove # prefix if present
            hex_color = format.font_color.lstrip("#")
            font_kwargs["color"] = Color(rgb=f"FF{hex_color}")

        if font_kwargs:
            c.font = Font(**font_kwargs)

        # Apply background color
        if format.bg_color is not None:
            hex_color = format.bg_color.lstrip("#")
            c.fill = PatternFill(
                start_color=f"FF{hex_color}",
                end_color=f"FF{hex_color}",
                fill_type="solid",
            )

        if format.number_format is not None:
            c.number_format = format.number_format

        align_kwargs = {}
        if format.h_align is not None:
            align_kwargs["horizontal"] = format.h_align
        if format.v_align is not None:
            align_kwargs["vertical"] = format.v_align
        if format.wrap is not None:
            align_kwargs["wrap_text"] = format.wrap
        if format.rotation is not None:
            align_kwargs["text_rotation"] = format.rotation
        if format.indent is not None:
            align_kwargs["indent"] = format.indent
        if align_kwargs:
            c.alignment = Alignment(**align_kwargs)

    def write_cell_border(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        """Apply border to a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]

        def make_side(edge: BorderEdge | None) -> Side:
            if edge is None:
                return Side()

            # Map our style to openpyxl style
            style_map = {
                BorderStyle.NONE: None,
                BorderStyle.THIN: "thin",
                BorderStyle.MEDIUM: "medium",
                BorderStyle.THICK: "thick",
                BorderStyle.DOUBLE: "double",
                BorderStyle.DASHED: "dashed",
                BorderStyle.DOTTED: "dotted",
                BorderStyle.HAIR: "hair",
                BorderStyle.MEDIUM_DASHED: "mediumDashed",
                BorderStyle.DASH_DOT: "dashDot",
                BorderStyle.MEDIUM_DASH_DOT: "mediumDashDot",
                BorderStyle.DASH_DOT_DOT: "dashDotDot",
                BorderStyle.MEDIUM_DASH_DOT_DOT: "mediumDashDotDot",
                BorderStyle.SLANT_DASH_DOT: "slantDashDot",
            }

            style = style_map.get(edge.style)
            if style is None:
                return Side()

            hex_color = edge.color.lstrip("#")
            from openpyxl.styles import Color

            return Side(style=style, color=Color(rgb=f"FF{hex_color}"))

        # Determine diagonal settings
        diagonal_side = Side()
        diagonal_up = False
        diagonal_down = False

        if border.diagonal_up:
            diagonal_side = make_side(border.diagonal_up)
            diagonal_up = True
        if border.diagonal_down:
            diagonal_side = make_side(border.diagonal_down)
            diagonal_down = True

        c.border = Border(
            left=make_side(border.left),
            right=make_side(border.right),
            top=make_side(border.top),
            bottom=make_side(border.bottom),
            diagonal=diagonal_side,
            diagonalUp=diagonal_up,
            diagonalDown=diagonal_down,
        )

    def save_workbook(self, workbook: Workbook, path: Path) -> None:
        """Save a workbook to a file."""
        workbook.save(str(path))

    def set_row_height(
        self,
        workbook: Workbook,
        sheet: str,
        row: int,
        height: float,
    ) -> None:
        ws = workbook[sheet]
        ws.row_dimensions[row].height = height

    def set_column_width(
        self,
        workbook: Workbook,
        sheet: str,
        column: str,
        width: float,
    ) -> None:
        ws = workbook[sheet]
        ws.column_dimensions[column].width = width

    # =========================================================================
    # Tier 2 Write Operations
    # =========================================================================

    def merge_cells(self, workbook: Workbook, sheet: str, cell_range: str) -> None:
        ws = workbook[sheet]
        ws.merge_cells(cell_range)

    def add_conditional_format(self, workbook: Workbook, sheet: str, rule: dict) -> None:
        ws = workbook[sheet]
        cf = rule.get("cf_rule", rule)
        range_ref = cf.get("range")
        rule_type = cf.get("rule_type")
        formula = cf.get("formula")
        operator = cf.get("operator")
        fmt = cf.get("format") or {}

        dxf = None
        if fmt:
            fill = None
            font = None
            if fmt.get("bg_color"):
                hex_color = fmt["bg_color"].lstrip("#")
                fill = PatternFill(
                    start_color=f"FF{hex_color}", end_color=f"FF{hex_color}", fill_type="solid"
                )
            if fmt.get("font_color"):
                hex_color = fmt["font_color"].lstrip("#")
                font = Font(color=f"FF{hex_color}")
            if fill or font:
                from openpyxl.styles import DifferentialStyle

                dxf = DifferentialStyle(fill=fill, font=font)

        if rule_type in ("cellIs", "cellIsRule"):
            op_map = {
                "greaterThan": "greaterThan",
                "lessThan": "lessThan",
                "between": "between",
                "equal": "equal",
                "notEqual": "notEqual",
                "greaterThanOrEqual": "greaterThanOrEqual",
                "lessThanOrEqual": "lessThanOrEqual",
            }
            op = op_map.get(operator, operator)
            from openpyxl.formatting.rule import CellIsRule

            rule_obj = CellIsRule(operator=op, formula=[formula], dxf=dxf)
            ws.conditional_formatting.add(range_ref, rule_obj)
        elif rule_type in ("expression", "formula"):
            rule_obj = FormulaRule(formula=[formula], dxf=dxf)
            ws.conditional_formatting.add(range_ref, rule_obj)
        elif rule_type == "colorScale":
            rule_obj = ColorScaleRule(
                start_type="min",
                start_color="FFAA0000",
                mid_type="percentile",
                mid_value=50,
                mid_color="FFFFFF00",
                end_type="max",
                end_color="FF00AA00",
            )
            ws.conditional_formatting.add(range_ref, rule_obj)
        elif rule_type == "dataBar":
            rule_obj = DataBarRule(
                start_type="min", end_type="max", color="FF638EC6", showValue=True
            )
            ws.conditional_formatting.add(range_ref, rule_obj)

    def add_data_validation(self, workbook: Workbook, sheet: str, validation: dict) -> None:
        ws = workbook[sheet]
        v = validation.get("validation", validation)
        dv = DataValidation(
            type=v.get("validation_type"),
            operator=v.get("operator"),
            formula1=v.get("formula1"),
            formula2=v.get("formula2"),
            allow_blank=v.get("allow_blank"),
            showInputMessage=v.get("show_input"),
            showErrorMessage=v.get("show_error"),
            promptTitle=v.get("prompt_title"),
            prompt=v.get("prompt"),
            errorTitle=v.get("error_title"),
            error=v.get("error"),
        )
        ws.add_data_validation(dv)
        dv.add(v.get("range"))

    def add_hyperlink(self, workbook: Workbook, sheet: str, link: dict) -> None:
        ws = workbook[sheet]
        data = link.get("hyperlink", link)
        cell = data.get("cell")
        target = data.get("target")
        display = data.get("display")
        tooltip = data.get("tooltip")
        internal = data.get("internal")
        c = ws[cell]
        if display is not None:
            c.value = display
        if internal:
            location = target.lstrip("#") if target else target
            c.hyperlink = Hyperlink(ref=cell, location=location)
        else:
            c.hyperlink = target
        if c.hyperlink and tooltip is not None:
            c.hyperlink.tooltip = tooltip

    def add_image(self, workbook: Workbook, sheet: str, image: dict) -> None:
        ws = workbook[sheet]
        data = image.get("image", image)
        path = data.get("path")
        cell = data.get("cell")
        if not path or not cell:
            return
        img = Image(path)
        ws.add_image(img, cell)

    def add_pivot_table(self, workbook: Workbook, sheet: str, pivot: dict) -> None:
        raise NotImplementedError("openpyxl does not support pivot table creation")

    def add_comment(self, workbook: Workbook, sheet: str, comment: dict) -> None:
        ws = workbook[sheet]
        data = comment.get("comment", comment)
        cell = data.get("cell")
        text = data.get("text")
        author = data.get("author") or ""
        if cell and text is not None:
            ws[cell].comment = Comment(text, author)

    def set_freeze_panes(self, workbook: Workbook, sheet: str, settings: dict) -> None:
        ws = workbook[sheet]
        cfg = settings.get("freeze", settings)
        mode = cfg.get("mode")
        if mode == "freeze":
            ws.freeze_panes = cfg.get("top_left_cell")
        elif mode == "split":
            ws.freeze_panes = None
            pane = ws.sheet_view.pane
            pane.xSplit = cfg.get("x_split")
            pane.ySplit = cfg.get("y_split")
            pane.topLeftCell = cfg.get("top_left_cell")
            pane.activePane = cfg.get("active_pane")
