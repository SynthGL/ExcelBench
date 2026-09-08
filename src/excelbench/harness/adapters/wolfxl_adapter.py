"""Adapter for WolfXL's public openpyxl-compatible Python API."""

import posixpath
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import wolfxl

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

# Formulas that produce known error values. WolfXL writes formulas rather than
# cached error values, so preserve the benchmark's normal error-write contract.
ERROR_FORMULA_MAP = {
    "=1/0": "#DIV/0!",
    "=NA()": "#N/A",
    '="text"+1': "#VALUE!",
}


def _color_to_hex(color: Any) -> str | None:
    """Return a normalized RGB value from an openpyxl-compatible color."""
    if color is None:
        return None

    if isinstance(color, str):
        # WolfXL exposes fill colors as plain ARGB strings where openpyxl
        # returns Color objects; both spellings normalize the same way.
        rgb: Any = color
    else:
        rgb = getattr(color, "rgb", None)
    if isinstance(rgb, str) and len(rgb) >= 6:
        return f"#{rgb[2:]}" if len(rgb) == 8 else f"#{rgb}"

    value = getattr(rgb, "value", None)
    if isinstance(value, str) and len(value) >= 6:
        return f"#{value[2:]}" if len(value) == 8 else f"#{value}"
    return None


def _col_letter(index: int) -> str:
    """Return the one-based column index as an A1 column letter."""
    result = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def _cell_value_from_wolfxl_cell(cell: Any) -> CellValue:
    """Convert a WolfXL cell proxy into ExcelBench's typed value model."""
    value = getattr(cell, "value", None)
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
        if value in {"#N/A", "#NULL!", "#NAME?", "#REF!"} or (
            value.startswith("#") and value.endswith("!")
        ):
            return CellValue(type=CellType.ERROR, value=value)
        if getattr(cell, "data_type", None) == "f" or value.startswith("="):
            formula = value if value.startswith("=") else f"={value}"
            if formula in ERROR_FORMULA_MAP:
                return CellValue(type=CellType.ERROR, value=ERROR_FORMULA_MAP[formula])
            return CellValue(type=CellType.FORMULA, value=value, formula=formula)
        return CellValue(type=CellType.STRING, value=value)
    return CellValue(type=CellType.STRING, value=str(value))


def _read_images_from_xlsx(path: Path, sheet_name: str) -> list[JSONDict]:
    """Read image anchors from OOXML drawing parts.

    WolfXL exposes images through its openpyxl-compatible API, but fixture image
    paths remain OOXML package paths. Reading drawing relationships directly
    preserves those package-relative paths for ExcelBench comparison.
    """
    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }

    def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
        return ET.fromstring(zf.read(name))

    def rel_targets(zf: zipfile.ZipFile, rels_path: str) -> dict[str, str]:
        root = read_xml(zf, rels_path)
        return {
            rel_id: target
            for rel in root.findall("rel:Relationship", ns)
            if (rel_id := rel.attrib.get("Id")) and (target := rel.attrib.get("Target"))
        }

    def resolve(base_part: str, target: str) -> str:
        if target.startswith("/"):
            return target.lstrip("/")
        return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))

    def rels_for(part: str) -> str:
        directory, filename = posixpath.split(part)
        return posixpath.join(directory, "_rels", f"{filename}.rels")

    def cell_from_marker(marker: ET.Element) -> str:
        column = int(marker.findtext("xdr:col", "0", ns)) + 1
        row = int(marker.findtext("xdr:row", "0", ns)) + 1
        return f"{_col_letter(column)}{row}"

    with zipfile.ZipFile(path) as zf:
        workbook = read_xml(zf, "xl/workbook.xml")
        workbook_rels = rel_targets(zf, "xl/_rels/workbook.xml.rels")
        sheet_part: str | None = None
        for sheet in workbook.findall("main:sheets/main:sheet", ns):
            if sheet.attrib.get("name") != sheet_name:
                continue
            rel_id = sheet.attrib.get(f"{{{ns['r']}}}id")
            target = workbook_rels.get(str(rel_id)) if rel_id else None
            if target:
                sheet_part = resolve("xl/workbook.xml", target)
                break
        if sheet_part is None:
            return []

        try:
            sheet_rels = rel_targets(zf, rels_for(sheet_part))
        except KeyError:
            return []

        images: list[JSONDict] = []
        for drawing_target in sheet_rels.values():
            if "drawing" not in drawing_target:
                continue
            drawing_part = resolve(sheet_part, drawing_target)
            drawing = read_xml(zf, drawing_part)
            try:
                drawing_rels = rel_targets(zf, rels_for(drawing_part))
            except KeyError:
                drawing_rels = {}
            for anchor_tag, anchor_name in (
                ("xdr:oneCellAnchor", "oneCell"),
                ("xdr:twoCellAnchor", "twoCell"),
            ):
                for anchor in drawing.findall(anchor_tag, ns):
                    marker = anchor.find("xdr:from", ns)
                    blip = anchor.find(".//a:blip", ns)
                    if marker is None or blip is None:
                        continue
                    embed = blip.attrib.get(f"{{{ns['r']}}}embed")
                    media_target = drawing_rels.get(str(embed)) if embed else None
                    if media_target:
                        images.append(
                            {
                                "cell": cell_from_marker(marker),
                                "path": f"/{resolve(drawing_part, media_target)}",
                                "anchor": anchor_name,
                            }
                        )
        return images


class WolfxlAdapter(ExcelAdapter):
    """Read and write XLSX workbooks through WolfXL's public Python API."""

    def __init__(self) -> None:
        self._workbook_paths: dict[int, Path] = {}

    @classmethod
    def is_available(cls) -> bool:
        """Return whether WolfXL 2's required public API is importable."""
        try:
            from wolfxl import (
                Alignment,
                Border,
                Font,
                PatternFill,
                Side,
                Workbook,
                load_workbook,
            )
            from wolfxl.chart import BarChart, LineChart, Reference
            from wolfxl.comments import Comment
            from wolfxl.drawing.image import Image
            from wolfxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
            from wolfxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
            from wolfxl.worksheet.datavalidation import DataValidation
            from wolfxl.worksheet.hyperlink import Hyperlink
            from wolfxl.worksheet.table import Table, TableColumn, TableStyleInfo

            required = (
                Alignment,
                AnchorMarker,
                BarChart,
                Border,
                ColorScaleRule,
                Comment,
                DataBarRule,
                DataValidation,
                Font,
                FormulaRule,
                Hyperlink,
                Image,
                LineChart,
                PatternFill,
                Reference,
                Side,
                Table,
                TableColumn,
                TableStyleInfo,
                TwoCellAnchor,
                Workbook,
                load_workbook,
            )
            return bool(required) and int(str(wolfxl.__version__).split(".", 1)[0]) >= 2
        except (ImportError, ValueError):
            return False

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="wolfxl",
            version=str(wolfxl.__version__),
            language="python",
            capabilities={"read", "write", "modify"},
        )

    @property
    def supported_read_extensions(self) -> set[str]:
        return {".xlsx"}

    # =========================================================================
    # Read operations
    # =========================================================================

    def open_workbook(self, path: Path) -> Any:
        workbook = wolfxl.load_workbook(path, data_only=False)
        self._workbook_paths[id(workbook)] = path
        return workbook

    def close_workbook(self, workbook: Any) -> None:
        self._workbook_paths.pop(id(workbook), None)
        close = getattr(workbook, "close", None)
        if callable(close):
            close()

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return [str(name) for name in workbook.sheetnames]

    def read_sheet_values(
        self, workbook: Any, sheet: str, cell_range: str | None = None
    ) -> list[list[CellValue]]:
        worksheet = workbook[sheet]
        rows = worksheet[cell_range] if cell_range else worksheet.iter_rows()
        return [[_cell_value_from_wolfxl_cell(cell) for cell in row] for row in rows]

    def read_cell_value(self, workbook: Any, sheet: str, cell: str) -> CellValue:
        return _cell_value_from_wolfxl_cell(workbook[sheet][cell])

    def read_cell_format(self, workbook: Any, sheet: str, cell: str) -> CellFormat:
        current = workbook[sheet][cell]
        font = current.font
        fill = current.fill
        alignment = current.alignment
        return CellFormat(
            bold=font.bold or None,
            italic=font.italic or None,
            underline=font.underline or None,
            strikethrough=font.strike or None,
            font_name=font.name or None,
            font_size=font.size or None,
            font_color=_color_to_hex(getattr(font, "color", None)),
            bg_color=(
                _color_to_hex(getattr(fill, "fgColor", None))
                if getattr(fill, "patternType", None) == "solid"
                else None
            ),
            number_format=current.number_format or None,
            h_align=getattr(alignment, "horizontal", None) or None,
            v_align=getattr(alignment, "vertical", None) or None,
            wrap=getattr(alignment, "wrap_text", None) or None,
            rotation=(
                getattr(alignment, "text_rotation", None)
                if getattr(alignment, "text_rotation", None) not in (0, None)
                else None
            ),
            indent=getattr(alignment, "indent", None) or None,
        )

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        border = workbook[sheet][cell].border

        def parse_side(side: Any) -> BorderEdge | None:
            if side is None or getattr(side, "style", None) is None:
                return None
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
            return BorderEdge(
                style=style_map.get(side.style, BorderStyle.THIN),
                color=_color_to_hex(getattr(side, "color", None)) or "#000000",
            )

        return BorderInfo(
            top=parse_side(border.top),
            bottom=parse_side(border.bottom),
            left=parse_side(border.left),
            right=parse_side(border.right),
            diagonal_up=parse_side(border.diagonal) if border.diagonalUp else None,
            diagonal_down=parse_side(border.diagonal) if border.diagonalDown else None,
        )

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        height = workbook[sheet].row_dimensions[row].height
        return float(height) if isinstance(height, (int, float)) else None

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        width = workbook[sheet].column_dimensions[column].width
        try:
            return float(width) if width is not None else None
        except (TypeError, ValueError):
            return None

    # =========================================================================
    # Tier 2 reads
    # =========================================================================

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        return [str(cell_range) for cell_range in workbook[sheet].merged_cells.ranges]

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[JSONDict]:
        conditional_formatting = workbook[sheet].conditional_formatting
        rules_by_range = getattr(conditional_formatting, "_cf_rules", None)
        if rules_by_range is None:
            self.unsupported_operation(
                "read_conditional_formats",
                "WolfXL's installed compatibility layer does not expose conditional rules.",
            )
        rules: list[JSONDict] = []
        for sqref, rule_list in rules_by_range.items():
            range_value = str(getattr(sqref, "sqref", sqref))
            for rule in rule_list:
                entry: JSONDict = {
                    "range": range_value,
                    "rule_type": getattr(rule, "type", None),
                    "operator": getattr(rule, "operator", None),
                    "formula": rule.formula[0]
                    if getattr(rule, "formula", None)
                    else None,
                    "priority": getattr(rule, "priority", None),
                    "stop_if_true": getattr(rule, "stopIfTrue", None),
                    "format": {},
                }
                dxf = getattr(rule, "dxf", None)
                if dxf is not None:
                    if bg := _color_to_hex(
                        getattr(getattr(dxf, "fill", None), "fgColor", None)
                    ):
                        entry["format"]["bg_color"] = bg
                    if font_color := _color_to_hex(
                        getattr(getattr(dxf, "font", None), "color", None)
                    ):
                        entry["format"]["font_color"] = font_color
                rules.append(entry)
        return rules

    def read_data_validations(self, workbook: Any, sheet: str) -> list[JSONDict]:
        validations = getattr(workbook[sheet].data_validations, "dataValidation", None)
        if validations is None:
            self.unsupported_operation(
                "read_data_validations",
                "WolfXL's installed compatibility layer does not expose data validations.",
            )
        return [
            {
                "range": str(entry.sqref),
                "validation_type": entry.type,
                "operator": entry.operator or ("between" if entry.formula2 else None),
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
            for entry in validations
        ]

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[JSONDict]:
        links: list[JSONDict] = []
        for row in workbook[sheet].iter_rows():
            for cell in row:
                hyperlink = cell.hyperlink
                if hyperlink is None:
                    continue
                if hyperlink.target and hyperlink.location:
                    target = f"{hyperlink.target}#{hyperlink.location}"
                    internal = False
                elif hyperlink.target:
                    target = hyperlink.target
                    internal = False
                else:
                    target = hyperlink.location
                    internal = True
                links.append(
                    {
                        "cell": cell.coordinate,
                        "target": target,
                        "display": cell.value,
                        "tooltip": hyperlink.tooltip,
                        "internal": internal,
                    }
                )
        return links

    def read_images(self, workbook: Any, sheet: str) -> list[JSONDict]:
        path = self._workbook_paths.get(id(workbook))
        if path is None:
            self.unsupported_operation(
                "read_images",
                "Image package paths are available only for workbooks opened from a file.",
            )
        return _read_images_from_xlsx(path, sheet)

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        pivots = getattr(workbook[sheet], "pivot_tables", None)
        if pivots is None:
            self.unsupported_operation(
                "read_pivot_tables",
                "WolfXL's installed compatibility layer does not expose pivot tables.",
            )
        results: list[JSONDict] = []
        for pivot in pivots:
            cache = getattr(pivot, "cache", None)
            cache_source = getattr(cache, "cacheSource", None)
            worksheet_source = getattr(cache_source, "worksheetSource", None)
            source_ref = getattr(worksheet_source, "ref", None)
            source_sheet = getattr(worksheet_source, "sheet", None)
            source_range: str | None
            if source_sheet and source_ref:
                source_range = f"{source_sheet}!{source_ref}"
            else:
                source_range = source_ref or getattr(cache_source, "ref", None)
            location = getattr(pivot, "location", None)
            target_cell = getattr(location, "ref", None) or location
            if target_cell and "!" not in str(target_cell):
                target_cell = f"{sheet}!{target_cell}"
            results.append(
                {
                    "name": getattr(pivot, "name", None),
                    "source_range": source_range,
                    "target_cell": target_cell,
                }
            )
        return results

    def read_comments(self, workbook: Any, sheet: str) -> list[JSONDict]:
        comments: list[JSONDict] = []
        for row in workbook[sheet].iter_rows():
            for cell in row:
                if cell.comment is not None:
                    comments.append(
                        {
                            "cell": cell.coordinate,
                            "text": cell.comment.text,
                            "author": cell.comment.author,
                            "threaded": False,
                        }
                    )
        return comments

    def read_freeze_panes(self, workbook: Any, sheet: str) -> JSONDict:
        worksheet = workbook[sheet]
        result: JSONDict = {}
        if worksheet.freeze_panes:
            result["mode"] = "freeze"
            result["top_left_cell"] = str(worksheet.freeze_panes)
        pane = getattr(worksheet.sheet_view, "pane", None)
        if pane is not None and pane.state == "split" and (pane.xSplit or pane.ySplit):
            result["mode"] = "split"
            result["x_split"] = int(pane.xSplit) if pane.xSplit is not None else None
            result["y_split"] = int(pane.ySplit) if pane.ySplit is not None else None
            if pane.topLeftCell:
                result["top_left_cell"] = pane.topLeftCell
            if pane.activePane:
                result["active_pane"] = pane.activePane
        return result

    # =========================================================================
    # Tier 3 reads
    # =========================================================================

    def read_named_ranges(self, workbook: Any, sheet: str) -> list[JSONDict]:
        def normalize_refers_to(value: Any) -> str:
            raw = str(value or "").lstrip("=")
            if "!" not in raw:
                return raw
            sheet_name, address = raw.split("!", 1)
            if sheet_name.startswith("'") and sheet_name.endswith("'"):
                sheet_name = sheet_name[1:-1].replace("''", "'")
            return f"{sheet_name}!{address}"

        results: list[JSONDict] = []
        seen: set[tuple[str, str, str]] = set()
        worksheet = workbook[sheet]
        sheet_id = workbook.worksheets.index(worksheet)
        for key in workbook.defined_names:
            name = workbook.defined_names.get(key)
            if name is None:
                continue
            local_sheet_id = getattr(name, "localSheetId", None)
            if local_sheet_id is not None and int(local_sheet_id) != sheet_id:
                continue
            scope = "sheet" if local_sheet_id is not None else "workbook"
            item: JSONDict = {
                "name": str(getattr(name, "name", key)),
                "scope": scope,
                "refers_to": normalize_refers_to(getattr(name, "attr_text", None)),
            }
            signature = (str(item["name"]), str(item["scope"]), str(item["refers_to"]))
            if signature not in seen:
                seen.add(signature)
                results.append(item)
        for key in worksheet.defined_names:
            name = worksheet.defined_names.get(key)
            if name is None:
                continue
            item = {
                "name": str(getattr(name, "name", key)),
                "scope": "sheet",
                "refers_to": normalize_refers_to(getattr(name, "attr_text", None)),
            }
            signature = (str(item["name"]), str(item["scope"]), str(item["refers_to"]))
            if signature not in seen:
                seen.add(signature)
                results.append(item)
        return results

    def read_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        worksheet = workbook[sheet]
        tables = getattr(worksheet, "tables", None)
        if tables is None:
            self.unsupported_operation(
                "read_tables",
                "WolfXL's installed compatibility layer does not expose tables.",
            )
        results: list[JSONDict] = []
        for table in tables.values():
            columns = [
                str(column.name)
                for column in (getattr(table, "tableColumns", []) or [])
                if column.name is not None
            ]
            if not columns:
                from wolfxl.utils.cell import range_boundaries

                boundaries = range_boundaries(str(table.ref))
                min_col = int(boundaries[0] or 0)
                min_row = int(boundaries[1] or 0)
                max_col = int(boundaries[2] or 0)
                columns = [
                    ""
                    if worksheet.cell(row=min_row, column=column).value is None
                    else str(worksheet.cell(row=min_row, column=column).value)
                    for column in range(min_col, max_col + 1)
                ]
            results.append(
                {
                    "name": getattr(table, "displayName", None)
                    or getattr(table, "name", None),
                    "ref": getattr(table, "ref", None),
                    "header_row": getattr(table, "headerRowCount", 1) != 0,
                    "totals_row": (getattr(table, "totalsRowCount", 0) or 0) > 0,
                    "style": (
                        table.tableStyleInfo.name
                        if getattr(table, "tableStyleInfo", None)
                        else None
                    ),
                    "columns": columns,
                    "autofilter": getattr(table, "autoFilter", None) is not None,
                }
            )
        return results

    # =========================================================================
    # Write operations
    # =========================================================================

    def create_workbook(self) -> Any:
        workbook = wolfxl.Workbook()
        if workbook.sheetnames:
            workbook.remove(workbook.active)
        return workbook

    def add_sheet(self, workbook: Any, name: str) -> None:
        workbook.create_sheet(name)

    def write_sheet_values(
        self, workbook: Any, sheet: str, start_cell: str, values: list[list[Any]]
    ) -> None:
        from wolfxl.utils.cell import column_index_from_string, coordinate_from_string

        column, row = coordinate_from_string(start_cell)
        worksheet = workbook[sheet]
        start_column = column_index_from_string(column)
        for row_offset, values_row in enumerate(values):
            for column_offset, value in enumerate(values_row):
                if value is not None:
                    worksheet.cell(
                        row=row + row_offset,
                        column=start_column + column_offset,
                        value=value,
                    )

    def write_cell_value(
        self, workbook: Any, sheet: str, cell: str, value: CellValue
    ) -> None:
        target = workbook[sheet][cell]
        if value.type == CellType.BLANK:
            target.value = None
        elif value.type == CellType.FORMULA:
            target.value = value.formula or value.value
        elif value.type == CellType.ERROR:
            error_formulas = {
                "#DIV/0!": "=1/0",
                "#N/A": "=NA()",
                "#VALUE!": '="text"+1',
                "#REF!": "=#REF!",
                "#NAME?": "=_undefined_name_",
                "#NUM!": "=SQRT(-1)",
                "#NULL!": "=A1:A2 B1:B2",
            }
            target.value = error_formulas.get(str(value.value), value.value)
        else:
            target.value = value.value

    def write_sheet_formats(
        self,
        workbook: Any,
        sheet: str,
        start_cell: str,
        formats: list[list[dict[str, Any] | None]],
    ) -> None:
        from wolfxl.utils.cell import column_index_from_string, coordinate_from_string

        column, row = coordinate_from_string(start_cell)
        start_column = column_index_from_string(column)
        for row_offset, format_row in enumerate(formats):
            for column_offset, format_data in enumerate(format_row):
                if format_data is not None:
                    self.write_cell_format(
                        workbook,
                        sheet,
                        f"{_col_letter(start_column + column_offset)}{row + row_offset}",
                        CellFormat(**format_data),
                    )

    def write_cell_format(
        self, workbook: Any, sheet: str, cell: str, format: CellFormat
    ) -> None:
        from wolfxl import Alignment, Color, Font, PatternFill

        target = workbook[sheet][cell]
        font_kwargs: JSONDict = {}
        for field, attr in (
            ("bold", "bold"),
            ("italic", "italic"),
            ("underline", "underline"),
            ("strikethrough", "strike"),
            ("font_name", "name"),
            ("font_size", "size"),
        ):
            if (field_value := getattr(format, field)) is not None:
                font_kwargs[attr] = field_value
        if format.font_color is not None:
            font_kwargs["color"] = Color(rgb=f"FF{format.font_color.lstrip('#')}")
        if font_kwargs:
            target.font = Font(**font_kwargs)
        if format.bg_color is not None:
            color = f"FF{format.bg_color.lstrip('#')}"
            target.fill = PatternFill(
                start_color=color, end_color=color, fill_type="solid"
            )
        if format.number_format is not None:
            target.number_format = format.number_format
        alignment_kwargs: JSONDict = {}
        for field, attr in (
            ("h_align", "horizontal"),
            ("v_align", "vertical"),
            ("wrap", "wrap_text"),
            ("rotation", "text_rotation"),
            ("indent", "indent"),
        ):
            if (field_value := getattr(format, field)) is not None:
                alignment_kwargs[attr] = field_value
        if alignment_kwargs:
            target.alignment = Alignment(**alignment_kwargs)

    def write_sheet_borders(
        self,
        workbook: Any,
        sheet: str,
        start_cell: str,
        borders: list[list[dict[str, Any] | None]],
    ) -> None:
        from wolfxl.utils.cell import column_index_from_string, coordinate_from_string

        from excelbench.harness.adapters.rust_adapter_utils import dict_to_border

        column, row = coordinate_from_string(start_cell)
        start_column = column_index_from_string(column)
        for row_offset, border_row in enumerate(borders):
            for column_offset, border_data in enumerate(border_row):
                if border_data is not None:
                    self.write_cell_border(
                        workbook,
                        sheet,
                        f"{_col_letter(start_column + column_offset)}{row + row_offset}",
                        dict_to_border(border_data),
                    )

    def write_cell_border(
        self, workbook: Any, sheet: str, cell: str, border: BorderInfo
    ) -> None:
        from wolfxl import Border, Color, Side

        def make_side(edge: BorderEdge | None) -> Any:
            if edge is None:
                return Side()
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
            return (
                Side()
                if style is None
                else Side(style=style, color=Color(rgb=f"FF{edge.color.lstrip('#')}"))
            )

        diagonal = make_side(border.diagonal_up or border.diagonal_down)
        workbook[sheet][cell].border = Border(
            left=make_side(border.left),
            right=make_side(border.right),
            top=make_side(border.top),
            bottom=make_side(border.bottom),
            diagonal=diagonal,
            diagonalUp=border.diagonal_up is not None,
            diagonalDown=border.diagonal_down is not None,
        )

    def set_row_height(
        self, workbook: Any, sheet: str, row: int, height: float
    ) -> None:
        workbook[sheet].row_dimensions[row].height = height

    def set_column_width(
        self, workbook: Any, sheet: str, column: str, width: float
    ) -> None:
        workbook[sheet].column_dimensions[column].width = width

    # =========================================================================
    # Tier 2 writes
    # =========================================================================

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        workbook[sheet].merge_cells(cell_range)

    def add_conditional_format(self, workbook: Any, sheet: str, rule: JSONDict) -> None:
        from wolfxl import Font, PatternFill
        from wolfxl.formatting.rule import (
            CellIsRule,
            ColorScaleRule,
            DataBarRule,
            FormulaRule,
        )

        cfg = rule.get("cf_rule", rule)
        range_ref = cfg.get("range")
        if not isinstance(range_ref, str):
            raise ValueError("add_conditional_format requires range")
        fmt = cfg.get("format") or {}
        fill = (
            PatternFill(
                start_color=f"FF{str(fmt['bg_color']).lstrip('#')}",
                end_color=f"FF{str(fmt['bg_color']).lstrip('#')}",
                fill_type="solid",
            )
            if fmt.get("bg_color")
            else None
        )
        font = (
            Font(color=f"FF{str(fmt['font_color']).lstrip('#')}")
            if fmt.get("font_color")
            else None
        )
        rule_type = cfg.get("rule_type")
        formula = cfg.get("formula")
        if rule_type in {"cellIs", "cellIsRule"}:
            rule_obj: Any = CellIsRule(
                operator=cfg.get("operator"),
                formula=[str(formula)],
                fill=fill,
                font=font,
                stopIfTrue=cfg.get("stop_if_true", False),
            )
        elif rule_type in {"expression", "formula"}:
            rule_obj = FormulaRule(
                formula=[str(formula)],
                fill=fill,
                font=font,
                stopIfTrue=cfg.get("stop_if_true", False),
            )
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
        elif rule_type == "dataBar":
            rule_obj = DataBarRule(
                start_type="min", end_type="max", color="FF638EC6", showValue=True
            )
        else:
            self.unsupported_operation(
                "add_conditional_format",
                f"WolfXL adapter does not support rule type {rule_type!r}.",
            )
        if cfg.get("priority") is not None:
            rule_obj.priority = cfg["priority"]
        workbook[sheet].conditional_formatting.add(range_ref, rule_obj)

    def add_data_validation(
        self, workbook: Any, sheet: str, validation: JSONDict
    ) -> None:
        from wolfxl.worksheet.datavalidation import DataValidation

        cfg = validation.get("validation", validation)
        cell_range = cfg.get("range")
        if not isinstance(cell_range, str):
            raise ValueError("add_data_validation requires range")
        data_validation = DataValidation(
            type=cfg.get("validation_type"),
            operator=cfg.get("operator"),
            formula1=cfg.get("formula1"),
            formula2=cfg.get("formula2"),
            allow_blank=cfg.get("allow_blank"),
            showInputMessage=cfg.get("show_input"),
            showErrorMessage=cfg.get("show_error"),
            promptTitle=cfg.get("prompt_title"),
            prompt=cfg.get("prompt"),
            errorTitle=cfg.get("error_title"),
            error=cfg.get("error"),
        )
        workbook[sheet].add_data_validation(data_validation)
        data_validation.add(cell_range)

    def add_hyperlink(self, workbook: Any, sheet: str, link: JSONDict) -> None:
        from wolfxl.worksheet.hyperlink import Hyperlink

        cfg = link.get("hyperlink", link)
        cell = cfg.get("cell")
        if not isinstance(cell, str):
            raise ValueError("add_hyperlink requires cell")
        target = cfg.get("target")
        current = workbook[sheet][cell]
        if cfg.get("display") is not None:
            current.value = cfg["display"]
        if cfg.get("internal"):
            current.hyperlink = Hyperlink(
                ref=cell, location=str(target).lstrip("#") if target else None
            )
        else:
            current.hyperlink = Hyperlink(
                ref=cell, target=str(target) if target else None
            )
        if current.hyperlink is not None and cfg.get("tooltip") is not None:
            current.hyperlink.tooltip = str(cfg["tooltip"])

    def add_image(self, workbook: Any, sheet: str, image: JSONDict) -> None:
        from wolfxl.drawing.image import Image

        cfg = image.get("image", image)
        path = cfg.get("path")
        cell = cfg.get("cell")
        if not isinstance(path, str) or not isinstance(cell, str):
            raise ValueError("add_image requires path and cell")
        workbook[sheet].add_image(Image(path), cell)

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: JSONDict) -> None:
        self.unsupported_operation(
            "add_pivot_table",
            (
                "WolfXL requires a concrete registered pivot cache; ExcelBench's "
                "generic pivot payload cannot construct one."
            ),
        )

    def add_comment(self, workbook: Any, sheet: str, comment: JSONDict) -> None:
        from wolfxl.comments import Comment

        cfg = comment.get("comment", comment)
        cell = cfg.get("cell")
        text = cfg.get("text")
        if not isinstance(cell, str) or text is None:
            raise ValueError("add_comment requires cell and text")
        workbook[sheet][cell].comment = Comment(str(text), str(cfg.get("author") or ""))

    def add_named_range(self, workbook: Any, sheet: str, named_range: JSONDict) -> None:
        from wolfxl.workbook.defined_name import DefinedName

        cfg = named_range.get("named_range", named_range)
        name = cfg.get("name")
        refers_to = cfg.get("refers_to")
        if not isinstance(name, str) or not isinstance(refers_to, str):
            raise ValueError("add_named_range requires name and refers_to")
        defined_name = DefinedName(
            name, attr_text=refers_to if refers_to.startswith("=") else f"={refers_to}"
        )
        if cfg.get("scope", "workbook") == "sheet":
            workbook[sheet].defined_names.add(defined_name)
        else:
            workbook.defined_names.add(defined_name)

    def add_table(self, workbook: Any, sheet: str, table: JSONDict) -> None:
        from wolfxl.worksheet.filters import AutoFilter
        from wolfxl.worksheet.table import Table, TableColumn, TableStyleInfo

        cfg = table.get("table", table)
        name = cfg.get("name")
        cell_range = cfg.get("ref")
        if not isinstance(name, str) or not isinstance(cell_range, str):
            raise ValueError("add_table requires name and ref")
        result = Table(displayName=name, ref=cell_range)
        if cfg.get("header_row") is False:
            result.headerRowCount = 0
        if cfg.get("totals_row"):
            result.totalsRowCount = 1
        if (style_name := cfg.get("style")) is not None:
            result.tableStyleInfo = TableStyleInfo(
                name=str(style_name),
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
        columns = cfg.get("columns")
        if isinstance(columns, list) and columns:
            result.tableColumns = [
                TableColumn(id=index, name=str(column))
                for index, column in enumerate(columns, start=1)
            ]
        if cfg.get("autofilter"):
            result.autoFilter = AutoFilter(ref=cell_range)
        workbook[sheet].add_table(result)

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        cfg = settings.get("freeze", settings)
        worksheet = workbook[sheet]
        if cfg.get("mode") == "freeze":
            worksheet.freeze_panes = cfg.get("top_left_cell")
            return
        if cfg.get("mode") == "split":
            from wolfxl.worksheet.views import Pane

            worksheet.freeze_panes = None
            pane = worksheet.sheet_view.pane or Pane()
            worksheet.sheet_view.pane = pane
            if cfg.get("x_split") is not None:
                pane.xSplit = cfg["x_split"]
            if cfg.get("y_split") is not None:
                pane.ySplit = cfg["y_split"]
            if cfg.get("top_left_cell") is not None:
                pane.topLeftCell = cfg["top_left_cell"]
            if cfg.get("active_pane") is not None:
                pane.activePane = cfg["active_pane"]
            pane.state = "split"

    # =========================================================================
    # Tier 4 operations
    # =========================================================================

    def read_sheet_protection(self, workbook: Any, sheet: str) -> JSONDict:
        protection = workbook[sheet].protection
        return {
            "protected": bool(protection.sheet),
            "password_hash_present": bool(protection.password or protection.hashValue),
            "format_cells": protection.formatCells,
            "insert_rows": protection.insertRows,
            "select_locked_cells": protection.selectLockedCells,
            "select_unlocked_cells": protection.selectUnlockedCells,
            "sort": protection.sort,
            "auto_filter": protection.autoFilter,
        }

    def set_sheet_protection(
        self, workbook: Any, sheet: str, settings: JSONDict
    ) -> None:
        cfg = settings.get("protection", settings)
        protection = workbook[sheet].protection
        for key, attr in (
            ("protected", "sheet"),
            ("format_cells", "formatCells"),
            ("insert_rows", "insertRows"),
            ("select_locked_cells", "selectLockedCells"),
            ("select_unlocked_cells", "selectUnlockedCells"),
            ("sort", "sort"),
            ("auto_filter", "autoFilter"),
        ):
            if cfg.get(key) is not None:
                setattr(protection, attr, bool(cfg[key]))
        if cfg.get("password"):
            protection.set_password(str(cfg["password"]))

    def read_page_setup(self, workbook: Any, sheet: str) -> JSONDict:
        worksheet = workbook[sheet]
        page_setup = worksheet.page_setup
        return {
            "orientation": page_setup.orientation,
            "fit_to_width": page_setup.fitToWidth,
            "fit_to_height": page_setup.fitToHeight,
            "scale": page_setup.scale,
            "print_title_rows": worksheet.print_title_rows,
            "header_center": worksheet.oddHeader.center.text,
            "footer_center": worksheet.oddFooter.center.text,
        }

    def set_page_setup(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        cfg = settings.get("page_setup", settings)
        worksheet = workbook[sheet]
        page_setup = worksheet.page_setup
        if cfg.get("orientation") is not None:
            page_setup.orientation = str(cfg["orientation"])
        if cfg.get("fit_to_width") is not None:
            page_setup.fitToWidth = int(cfg["fit_to_width"])
        if cfg.get("fit_to_height") is not None:
            page_setup.fitToHeight = int(cfg["fit_to_height"])
        if cfg.get("scale") is not None:
            page_setup.scale = int(cfg["scale"])
        if cfg.get("print_title_rows") is not None:
            # WolfXL 2.1's public setter accepts the unanchored openpyxl
            # compatibility spelling ("1:2"), while ExcelBench's contract
            # uses the OOXML-style "$1:$2".
            worksheet.print_title_rows = str(cfg["print_title_rows"]).replace("$", "")
        if cfg.get("header_center") is not None:
            worksheet.oddHeader.center.text = str(cfg["header_center"])
        if cfg.get("footer_center") is not None:
            worksheet.oddFooter.center.text = str(cfg["footer_center"])

    def read_chart_anchors(self, workbook: Any, sheet: str) -> list[JSONDict]:
        def cell_from_marker(marker: Any) -> str | None:
            if marker is None:
                return None
            return f"{_col_letter(marker.col + 1)}{marker.row + 1}"

        charts = getattr(workbook[sheet], "_charts", None)
        if charts is None:
            self.unsupported_operation(
                "read_chart_anchors",
                "WolfXL's installed compatibility layer does not expose charts.",
            )
        results: list[JSONDict] = []
        for chart in charts:
            anchor = getattr(chart, "anchor", None)
            # In WolfXL 2.1 a loaded chart's public anchor is its default
            # placement cell, while the persisted DrawingML anchor is retained
            # in the openpyxl-compatible private _anchor field.
            if isinstance(anchor, str):
                anchor = getattr(chart, "_anchor", None)
            anchor_name = type(anchor).__name__ if anchor is not None else ""
            results.append(
                {
                    "type": type(chart).__name__.replace("Chart", "").lower() or None,
                    "anchor_type": (
                        "twoCell"
                        if anchor_name == "TwoCellAnchor"
                        else "oneCell"
                        if anchor_name == "OneCellAnchor"
                        else "absolute"
                        if anchor_name == "AbsoluteAnchor"
                        else None
                    ),
                    "from": cell_from_marker(getattr(anchor, "_from", None)),
                    "to": cell_from_marker(getattr(anchor, "to", None))
                    if anchor_name == "TwoCellAnchor"
                    else None,
                }
            )
        return results

    def add_chart_with_anchor(self, workbook: Any, sheet: str, chart: JSONDict) -> None:
        from wolfxl.chart import BarChart, LineChart, Reference
        from wolfxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
        from wolfxl.utils.cell import coordinate_to_tuple

        cfg = chart.get("chart", chart)
        data_ref = cfg.get("data_ref")
        if not isinstance(data_ref, str) or not data_ref:
            raise ValueError("add_chart_with_anchor requires data_ref")
        worksheet = workbook[sheet]
        series_chart: Any = (
            LineChart() if str(cfg.get("type", "bar")).lower() == "line" else BarChart()
        )
        qualified_data_ref = (
            data_ref if "!" in data_ref else f"{worksheet.title}!{data_ref}"
        )
        series_chart.add_data(
            Reference(worksheet, range_string=qualified_data_ref),
            titles_from_data=False,
        )
        if (categories_ref := cfg.get("categories_ref")) is not None:
            categories = str(categories_ref)
            qualified_categories = (
                categories if "!" in categories else f"{worksheet.title}!{categories}"
            )
            series_chart.set_categories(
                Reference(worksheet, range_string=qualified_categories)
            )
        from_row, from_column = coordinate_to_tuple(str(cfg.get("from") or "B2"))
        to_row, to_column = coordinate_to_tuple(str(cfg.get("to") or "H12"))
        series_chart.anchor = TwoCellAnchor(
            _from=AnchorMarker(col=from_column - 1, row=from_row - 1),
            to=AnchorMarker(col=to_column - 1, row=to_row - 1),
        )
        worksheet.add_chart(series_chart)

    def save_workbook(self, workbook: Any, path: Path) -> None:
        workbook.save(path)
