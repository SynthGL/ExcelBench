"""Adapter for the :mod:`aspose.cells_foss` library."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any

try:
    import aspose.cells_foss as aspose
    from aspose.cells_foss.data_validation import (
        DataValidationOperator,
        DataValidationType,
    )
    from aspose.cells_foss.workbook_properties import DefinedName
except ImportError as exc:  # pragma: no cover - exercised by the adapter registry.
    raise ImportError(
        "AsposeCellsFossAdapter requires the optional 'aspose-cells-foss' package."
    ) from exc

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

_ERROR_FORMULA_MAP = {
    "=1/0": "#DIV/0!",
    "=NA()": "#N/A",
    '="text"+1': "#VALUE!",
    "=#REF!": "#REF!",
    "=_undefined_name_": "#NAME?",
    "=SQRT(-1)": "#NUM!",
    "=A1:A2 B1:B2": "#NULL!",
}
_ERROR_VALUE_FORMULA_MAP = {value: key for key, value in _ERROR_FORMULA_MAP.items()}


_VALIDATION_TYPES = {
    "none": DataValidationType.NONE,
    "whole": DataValidationType.WHOLE_NUMBER,
    "wholeNumber": DataValidationType.WHOLE_NUMBER,
    "decimal": DataValidationType.DECIMAL,
    "list": DataValidationType.LIST,
    "date": DataValidationType.DATE,
    "time": DataValidationType.TIME,
    "textLength": DataValidationType.TEXT_LENGTH,
    "text_length": DataValidationType.TEXT_LENGTH,
    "custom": DataValidationType.CUSTOM,
}
_VALIDATION_TYPE_NAMES = {
    DataValidationType.NONE: "none",
    DataValidationType.WHOLE_NUMBER: "whole",
    DataValidationType.DECIMAL: "decimal",
    DataValidationType.LIST: "list",
    DataValidationType.DATE: "date",
    DataValidationType.TIME: "time",
    DataValidationType.TEXT_LENGTH: "textLength",
    DataValidationType.CUSTOM: "custom",
}
_VALIDATION_OPERATORS = {
    "between": DataValidationOperator.BETWEEN,
    "notBetween": DataValidationOperator.NOT_BETWEEN,
    "equal": DataValidationOperator.EQUAL,
    "notEqual": DataValidationOperator.NOT_EQUAL,
    "greaterThan": DataValidationOperator.GREATER_THAN,
    "lessThan": DataValidationOperator.LESS_THAN,
    "greaterThanOrEqual": DataValidationOperator.GREATER_THAN_OR_EQUAL,
    "lessThanOrEqual": DataValidationOperator.LESS_THAN_OR_EQUAL,
}
_VALIDATION_OPERATOR_NAMES = {
    value: key for key, value in _VALIDATION_OPERATORS.items()
}


def _color_to_hex(value: Any) -> str | None:
    """Normalize the library's RGB/ARGB colors to the benchmark's #RRGGBB form."""
    if not isinstance(value, str):
        return None
    raw = value.lstrip("#").upper()
    if len(raw) == 8:
        raw = raw[2:]
    if len(raw) != 6 or any(char not in "0123456789ABCDEF" for char in raw):
        return None
    return f"#{raw}"


def _color_to_argb(value: str) -> str:
    """Convert a benchmark #RRGGBB color to the library's AARRGGBB form."""
    raw = value.lstrip("#").upper()
    if len(raw) == 6:
        return f"FF{raw}"
    if len(raw) == 8:
        return raw
    raise ValueError(f"Expected a six- or eight-digit hex color, got {value!r}")


def _cell_value_from_aspose_cell(cell: Any) -> CellValue:
    """Convert an aspose-cells-foss cell to the benchmark's typed value model."""
    formula = getattr(cell, "formula", None)
    value = getattr(cell, "value", None)
    if isinstance(formula, str) and formula:
        formula_text = formula if formula.startswith("=") else f"={formula}"
        if formula_text in _ERROR_FORMULA_MAP:
            return CellValue(
                type=CellType.ERROR, value=_ERROR_FORMULA_MAP[formula_text]
            )
        return CellValue(type=CellType.FORMULA, value=value, formula=formula_text)
    if value is None:
        return CellValue(type=CellType.BLANK)
    if isinstance(value, bool):
        return CellValue(type=CellType.BOOLEAN, value=value)
    if isinstance(value, (int, float)):
        return CellValue(type=CellType.NUMBER, value=value)
    if isinstance(value, datetime):
        if (
            value.hour == 0
            and value.minute == 0
            and value.second == 0
            and value.microsecond == 0
        ):
            return CellValue(type=CellType.DATE, value=value.date())
        return CellValue(type=CellType.DATETIME, value=value)
    if isinstance(value, date):
        return CellValue(type=CellType.DATE, value=value)
    if isinstance(value, str):
        if value.startswith("#") and value.endswith("!") or value == "#N/A":
            return CellValue(type=CellType.ERROR, value=value)
        return CellValue(type=CellType.STRING, value=value)
    return CellValue(type=CellType.STRING, value=str(value))


def _edge_from_aspose_border(border: Any) -> BorderEdge | None:
    """Convert one aspose-cells-foss border side to a benchmark border edge."""
    line_style = getattr(border, "line_style", "none")
    if not isinstance(line_style, str) or line_style == "none":
        return None
    try:
        style = BorderStyle(line_style)
    except ValueError:
        style = BorderStyle.THIN
    return BorderEdge(
        style=style, color=_color_to_hex(getattr(border, "color", None)) or "#000000"
    )


def _set_aspose_border_edge(target: Any, edge: BorderEdge | None) -> None:
    """Apply a benchmark border edge to an aspose-cells-foss border side."""
    if edge is None:
        target.line_style = "none"
        return
    target.line_style = edge.style.value
    target.color = _color_to_argb(edge.color)


def _coordinates(cell_range: str) -> tuple[int, int, int, int]:
    """Return zero-based top-left and bottom-right coordinates for an A1 range."""
    normalized = cell_range.replace("$", "").upper()
    start, separator, end = normalized.partition(":")
    if not separator:
        end = start
    start_row, start_col = aspose.Cells.coordinate_from_string(start)
    end_row, end_col = aspose.Cells.coordinate_from_string(end)
    return start_row - 1, start_col - 1, end_row - 1, end_col - 1


def _center_section(value: Any) -> str | None:
    """Extract the center section from an OOXML header/footer string."""
    if not isinstance(value, str) or "&C" not in value:
        return None
    center = value.split("&C", 1)[1]
    for marker in ("&L", "&R"):
        center = center.split(marker, 1)[0]
    return center or None


class AsposeCellsFossAdapter(ExcelAdapter):
    """Read/write Excel adapter for aspose-cells-foss 26.7.0 and compatible releases."""

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="aspose-cells-foss",
            version=str(getattr(aspose, "__version__", "unknown")),
            language="python",
            capabilities={"read", "write"},
        )

    def open_workbook(self, path: Path) -> Any:
        """Open an XLSX workbook for reading and modification."""
        return aspose.Workbook(str(path))

    def close_workbook(self, workbook: Any) -> None:
        """Release a workbook; aspose-cells-foss keeps no open file handle."""
        del workbook

    def get_sheet_names(self, workbook: Any) -> list[str]:
        """Return worksheet names in workbook order."""
        return [str(worksheet.name) for worksheet in workbook.worksheets]

    def read_sheet_values(
        self, workbook: Any, sheet: str, cell_range: str | None = None
    ) -> list[list[CellValue]]:
        """Read a rectangular cell range as typed benchmark values."""
        worksheet = workbook.get_worksheet(sheet)
        if worksheet is None:
            raise ValueError(f"Worksheet {sheet!r} not found")
        if cell_range is None:
            cells = worksheet.cells
            if not cells._cells:
                return []
            rows_and_columns = [
                aspose.Cells.coordinate_from_string(ref) for ref in cells._cells
            ]
            min_row = min(row for row, _ in rows_and_columns)
            max_row = max(row for row, _ in rows_and_columns)
            min_col = min(column for _, column in rows_and_columns)
            max_col = max(column for _, column in rows_and_columns)
        else:
            start_row, start_col, end_row, end_col = _coordinates(cell_range)
            min_row, max_row = start_row + 1, end_row + 1
            min_col, max_col = start_col + 1, end_col + 1
        return [
            [_cell_value_from_aspose_cell(cell) for cell in row]
            for row in worksheet.cells.iter_rows(
                min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
            )
        ]

    def read_cell_value(self, workbook: Any, sheet: str, cell: str) -> CellValue:
        """Read a typed value, preserving formula text."""
        return _cell_value_from_aspose_cell(workbook.get_worksheet(sheet).cells[cell])

    def read_cell_format(self, workbook: Any, sheet: str, cell: str) -> CellFormat:
        """Read a cell's font, fill, number-format, and alignment properties."""
        style = workbook.get_worksheet(sheet).cells[cell].style
        font = style.font
        fill = style.fill
        alignment = style.alignment
        return CellFormat(
            bold=True if font.bold else None,
            italic=True if font.italic else None,
            underline="single" if font.underline else None,
            strikethrough=True if font.strikethrough else None,
            font_name=str(font.name) if font.name else None,
            font_size=float(font.size) if font.size else None,
            font_color=_color_to_hex(font.color),
            bg_color=_color_to_hex(fill.foreground_color)
            if fill.pattern_type == "solid"
            else None,
            number_format=str(style.number_format) if style.number_format else None,
            h_align=alignment.horizontal if alignment.horizontal != "general" else None,
            v_align=alignment.vertical if alignment.vertical != "bottom" else None,
            wrap=True if alignment.wrap_text else None,
            rotation=alignment.text_rotation if alignment.text_rotation else None,
            indent=alignment.indent if alignment.indent else None,
        )

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        """Read all four and diagonal borders from a cell."""
        borders = workbook.get_worksheet(sheet).cells[cell].style.borders
        return BorderInfo(
            top=_edge_from_aspose_border(borders.top),
            bottom=_edge_from_aspose_border(borders.bottom),
            left=_edge_from_aspose_border(borders.left),
            right=_edge_from_aspose_border(borders.right),
            diagonal_up=_edge_from_aspose_border(borders.diagonal)
            if borders.diagonal_up
            else None,
            diagonal_down=_edge_from_aspose_border(borders.diagonal)
            if borders.diagonal_down
            else None,
        )

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        """Read a 1-indexed row's height in points."""
        value = workbook.get_worksheet(sheet).cells.get_row_height(row - 1)
        return float(value) if isinstance(value, (int, float)) else None

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        """Read a column's display width in characters."""
        value = workbook.get_worksheet(sheet).cells.get_column_width(column)
        return float(value) if isinstance(value, (int, float)) else None

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        """Read merged A1 ranges."""
        return workbook.get_worksheet(sheet).cells.get_merged_cells()

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[JSONDict]:
        """Read conditional-format rules exposed by the FOSS API."""
        rules: list[JSONDict] = []
        for conditional_format in workbook.get_worksheet(sheet).conditional_formats:
            rule_type = conditional_format.type
            if rule_type == "cellValue":
                rule_type = "cellIs"
            elif rule_type == "formula":
                rule_type = "expression"
            rules.append(
                {
                    "range": conditional_format.range,
                    "rule_type": rule_type,
                    "operator": conditional_format.operator,
                    "formula": conditional_format.formula
                    if conditional_format.formula is not None
                    else conditional_format.formula1,
                    "priority": conditional_format.priority,
                    "stop_if_true": conditional_format.stop_if_true,
                    "format": {
                        **(
                            {
                                "bg_color": _color_to_hex(
                                    conditional_format.fill.foreground_color
                                )
                            }
                            if conditional_format.fill.pattern_type == "solid"
                            and _color_to_hex(conditional_format.fill.foreground_color)
                            else {}
                        ),
                        **(
                            {"font_color": _color_to_hex(conditional_format.font.color)}
                            if _color_to_hex(conditional_format.font.color)
                            not in (None, "#000000")
                            else {}
                        ),
                    },
                }
            )
        return rules

    def read_data_validations(self, workbook: Any, sheet: str) -> list[JSONDict]:
        """Read data-validation rules."""
        validations: list[JSONDict] = []
        for validation in workbook.get_worksheet(sheet).data_validations:
            validation_type = _VALIDATION_TYPE_NAMES.get(
                validation.type, str(validation.type)
            )
            operator = _VALIDATION_OPERATOR_NAMES.get(
                validation.operator, str(validation.operator)
            )
            validations.append(
                {
                    "range": validation.sqref,
                    "validation_type": validation_type,
                    "operator": operator,
                    "formula1": validation.formula1,
                    "formula2": validation.formula2,
                    "allow_blank": validation.allow_blank,
                    "show_input": validation.show_input_message,
                    "show_error": validation.show_error_message,
                    "prompt_title": validation.input_title,
                    "prompt": validation.input_message,
                    "error_title": validation.error_title,
                    "error": validation.error_message,
                }
            )
        return validations

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[JSONDict]:
        """Read worksheet hyperlinks with internal/external targets normalized."""
        worksheet = workbook.get_worksheet(sheet)
        return [
            {
                "cell": link.range,
                "target": link.sub_address if link.sub_address else link.address,
                "display": link.text_to_display or worksheet.cells[link.range].value,
                "tooltip": link.screen_tip or None,
                "internal": bool(link.sub_address),
            }
            for link in worksheet.hyperlinks
        ]

    def read_images(self, workbook: Any, sheet: str) -> list[JSONDict]:
        """Reject image reads because one-cell drawing anchors are not surfaced reliably."""
        self.unsupported_operation(
            "read_images",
            "aspose-cells-foss does not preserve the source one-cell versus two-cell anchor kind.",
        )

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        """Reject pivot reads because the FOSS API has no pivot-table model."""
        self.unsupported_operation(
            "read_pivot_tables", "aspose-cells-foss exposes no pivot-table API."
        )

    def read_comments(self, workbook: Any, sheet: str) -> list[JSONDict]:
        """Read legacy cell comments."""
        worksheet = workbook.get_worksheet(sheet)
        return [
            {
                "cell": cell_ref,
                "text": comment["text"],
                "author": comment["author"],
                "threaded": False,
            }
            for cell_ref, cell in worksheet.cells._cells.items()
            if (comment := cell.comment) is not None
        ]

    def read_freeze_panes(self, workbook: Any, sheet: str) -> JSONDict:
        """Read frozen or split pane settings."""
        pane = workbook.get_worksheet(sheet).properties.pane
        if pane.state is None:
            return {}
        result: JSONDict = {
            "mode": "freeze" if pane.state in {"frozen", "frozenSplit"} else "split"
        }
        if pane.x_split is not None:
            result["x_split"] = int(pane.x_split)
        if pane.y_split is not None:
            result["y_split"] = int(pane.y_split)
        if pane.top_left_cell is not None:
            result["top_left_cell"] = pane.top_left_cell
        if pane.active_pane is not None:
            result["active_pane"] = pane.active_pane
        return result

    def read_named_ranges(self, workbook: Any, sheet: str) -> list[JSONDict]:
        """Read workbook names plus names local to the requested worksheet."""
        sheet_index = self.get_sheet_names(workbook).index(sheet)
        return [
            {
                "name": defined_name.name,
                "scope": "sheet"
                if defined_name.local_sheet_id is not None
                else "workbook",
                "refers_to": str(defined_name.refers_to).lstrip("="),
            }
            for defined_name in workbook.properties.defined_names
            if defined_name.local_sheet_id is None
            or defined_name.local_sheet_id == sheet_index
        ]

    def read_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        """Read structured table definitions."""
        return [
            {
                "name": table.display_name or table.name,
                "ref": table.ref,
                "header_row": table.has_headers,
                "totals_row": table.show_totals_row,
                "style": table.table_style_info.name
                if table.table_style_info
                else None,
                "columns": [column.name for column in table.columns],
                "autofilter": table.show_auto_filter,
            }
            for table in workbook.get_worksheet(sheet).tables
        ]

    def read_sheet_protection(self, workbook: Any, sheet: str) -> JSONDict:
        """Read raw OOXML-equivalent sheet-protection flags."""
        protection = workbook.get_worksheet(sheet).properties.protection
        return {
            "protected": bool(protection.sheet),
            "password_hash_present": bool(
                protection.password
                or protection.hash_value
                or protection.algorithm_name
            ),
            "format_cells": protection.format_cells,
            "insert_rows": protection.insert_rows,
            "select_locked_cells": protection.select_locked_cells,
            "select_unlocked_cells": protection.select_unlocked_cells,
            "sort": protection.sort,
            "auto_filter": protection.auto_filter,
        }

    def read_page_setup(self, workbook: Any, sheet: str) -> JSONDict:
        """Read print settings supported by the FOSS worksheet-properties API."""
        worksheet = workbook.get_worksheet(sheet)
        page_setup = worksheet.properties.page_setup
        header_footer = worksheet.properties.header_footer
        return {
            "orientation": page_setup.orientation,
            "fit_to_width": page_setup.fit_to_width,
            "fit_to_height": page_setup.fit_to_height,
            "scale": page_setup.scale,
            "print_title_rows": None,
            "header_center": _center_section(header_footer.odd_header),
            "footer_center": _center_section(header_footer.odd_footer),
        }

    def read_chart_anchors(self, workbook: Any, sheet: str) -> list[JSONDict]:
        """Read charts and their explicit two-cell drawing anchors."""
        results: list[JSONDict] = []
        for chart in workbook.get_worksheet(sheet).charts:
            chart_type = getattr(chart.type, "name", str(chart.type)).lower()
            results.append(
                {
                    "type": chart_type,
                    "anchor_type": "twoCell",
                    "from": aspose.Cells.coordinate_to_string(
                        chart._upper_left_row + 1, chart._upper_left_column + 1
                    ),
                    "to": aspose.Cells.coordinate_to_string(
                        chart._lower_right_row + 1, chart._lower_right_column + 1
                    ),
                }
            )
        return results

    def create_workbook(self) -> Any:
        """Create an empty workbook without an implicit worksheet."""
        workbook = aspose.Workbook()
        workbook.remove_worksheet(0)
        return workbook

    def add_sheet(self, workbook: Any, name: str) -> None:
        """Add a named worksheet."""
        workbook.add_worksheet(name)

    def write_sheet_values(
        self, workbook: Any, sheet: str, start_cell: str, values: list[list[Any]]
    ) -> None:
        """Write a rectangular grid of raw values for throughput workloads."""
        start_row, start_col = aspose.Cells.coordinate_from_string(start_cell)
        cells = workbook.get_worksheet(sheet).cells
        for row_offset, values_row in enumerate(values):
            for column_offset, raw_value in enumerate(values_row):
                if raw_value is not None:
                    cells.cell(
                        start_row + row_offset, start_col + column_offset
                    ).value = raw_value

    def write_cell_value(
        self, workbook: Any, sheet: str, cell: str, value: CellValue
    ) -> None:
        """Write a typed cell value, including formulas and formula-produced errors."""
        target = workbook.get_worksheet(sheet).cells[cell]
        target.formula = None
        if value.type == CellType.BLANK:
            target.value = None
        elif value.type == CellType.FORMULA:
            formula = value.formula if value.formula is not None else value.value
            if not isinstance(formula, str):
                raise ValueError("Formula CellValue must provide a formula string")
            target.formula = formula if formula.startswith("=") else f"={formula}"
            target.value = None
        elif value.type == CellType.ERROR:
            formula = _ERROR_VALUE_FORMULA_MAP.get(str(value.value))
            if formula is None:
                self.unsupported_operation(
                    "write_cell_value",
                    f"Cannot serialize Excel error {value.value!r} without a formula.",
                )
            target.formula = formula
            target.value = None
        else:
            target.value = value.value

    def write_cell_format(
        self, workbook: Any, sheet: str, cell: str, format: CellFormat
    ) -> None:
        """Apply benchmark formatting fields the FOSS style model can represent."""
        style = workbook.get_worksheet(sheet).cells[cell].style
        if format.bold is not None:
            style.font.bold = format.bold
        if format.italic is not None:
            style.font.italic = format.italic
        if format.underline is not None:
            if format.underline not in {"single", True}:
                self.unsupported_operation(
                    "write_cell_format",
                    (
                        "aspose-cells-foss supports only a single underline, "
                        "not accounting or double variants."
                    ),
                )
            style.font.underline = True
        if format.strikethrough is not None:
            style.font.strikethrough = format.strikethrough
        if format.font_name is not None:
            style.font.name = format.font_name
        if format.font_size is not None:
            style.font.size = format.font_size
        if format.font_color is not None:
            style.font.color = _color_to_argb(format.font_color)
        if format.bg_color is not None:
            style.fill.set_solid_fill(_color_to_argb(format.bg_color))
        if format.number_format is not None:
            style.number_format = format.number_format
        if format.h_align is not None:
            style.alignment.horizontal = format.h_align
        if format.v_align is not None:
            style.alignment.vertical = format.v_align
        if format.wrap is not None:
            style.alignment.wrap_text = format.wrap
        if format.rotation is not None:
            style.alignment.text_rotation = format.rotation
        if format.indent is not None:
            style.alignment.indent = format.indent

    def write_cell_border(
        self, workbook: Any, sheet: str, cell: str, border: BorderInfo
    ) -> None:
        """Apply four-side and diagonal borders."""
        borders = workbook.get_worksheet(sheet).cells[cell].style.borders
        _set_aspose_border_edge(borders.top, border.top)
        _set_aspose_border_edge(borders.bottom, border.bottom)
        _set_aspose_border_edge(borders.left, border.left)
        _set_aspose_border_edge(borders.right, border.right)
        diagonal = border.diagonal_up or border.diagonal_down
        _set_aspose_border_edge(borders.diagonal, diagonal)
        borders.diagonal_up = border.diagonal_up is not None
        borders.diagonal_down = border.diagonal_down is not None

    def set_row_height(
        self, workbook: Any, sheet: str, row: int, height: float
    ) -> None:
        """Set a 1-indexed row height in points."""
        workbook.get_worksheet(sheet).cells.set_row_height(row - 1, height)

    def set_column_width(
        self, workbook: Any, sheet: str, column: str, width: float
    ) -> None:
        """Set a column width in display characters."""
        workbook.get_worksheet(sheet).cells.set_column_width(column, width)

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        """Merge an A1 range."""
        workbook.get_worksheet(sheet).cells.merge_range(cell_range)

    def add_conditional_format(self, workbook: Any, sheet: str, rule: JSONDict) -> None:
        """Add a conditional-format rule."""
        config = rule.get("cf_rule", rule)
        rule_type = str(config.get("rule_type") or "")
        type_mapping = {
            "cellIs": "cellValue",
            "cellIsRule": "cellValue",
            "expression": "formula",
        }
        aspose_type = type_mapping.get(rule_type, rule_type)
        if not aspose_type:
            raise ValueError("Conditional-format rule requires rule_type")
        conditional_format = workbook.get_worksheet(sheet).conditional_formats.add()
        conditional_format.type = aspose_type
        conditional_format.range = config.get("range")
        if config.get("operator") is not None:
            conditional_format.operator = config["operator"]
        formula = config.get("formula")
        if aspose_type == "formula":
            conditional_format.formula = formula
        elif formula is not None:
            conditional_format.formula1 = formula
        if config.get("priority") is not None:
            conditional_format.priority = int(config["priority"])
        if config.get("stop_if_true") is not None:
            conditional_format.stop_if_true = bool(config["stop_if_true"])
        style = config.get("format") or {}
        if style.get("bg_color"):
            conditional_format.fill.set_solid_fill(
                _color_to_argb(str(style["bg_color"]))
            )
        if style.get("font_color"):
            conditional_format.font.color = _color_to_argb(str(style["font_color"]))

    def add_data_validation(
        self, workbook: Any, sheet: str, validation: JSONDict
    ) -> None:
        """Add an Excel data-validation rule."""
        config = validation.get("validation", validation)
        raw_type = str(config.get("validation_type") or "none")
        validation_type = _VALIDATION_TYPES.get(raw_type)
        if validation_type is None:
            self.unsupported_operation(
                "add_data_validation", f"Unsupported data-validation type {raw_type!r}."
            )
        raw_operator = str(config.get("operator") or "between")
        operator = _VALIDATION_OPERATORS.get(raw_operator)
        if operator is None:
            self.unsupported_operation(
                "add_data_validation",
                f"Unsupported data-validation operator {raw_operator!r}.",
            )
        target = workbook.get_worksheet(sheet).data_validations.add(
            str(config.get("range") or ""),
            validation_type,
            operator,
            config.get("formula1"),
            config.get("formula2"),
        )
        if config.get("allow_blank") is not None:
            target.allow_blank = bool(config["allow_blank"])
        if config.get("show_input") is not None:
            target.show_input_message = bool(config["show_input"])
        if config.get("show_error") is not None:
            target.show_error_message = bool(config["show_error"])
        if config.get("prompt_title") is not None:
            target.input_title = str(config["prompt_title"])
        if config.get("prompt") is not None:
            target.input_message = str(config["prompt"])
        if config.get("error_title") is not None:
            target.error_title = str(config["error_title"])
        if config.get("error") is not None:
            target.error_message = str(config["error"])

    def add_hyperlink(self, workbook: Any, sheet: str, link: JSONDict) -> None:
        """Add an internal or external hyperlink."""
        config = link.get("hyperlink", link)
        cell = str(config.get("cell") or "")
        target = str(config.get("target") or "")
        internal = bool(config.get("internal"))
        if not cell or not target:
            raise ValueError("Hyperlinks require both cell and target")
        hyperlink = workbook.get_worksheet(sheet).hyperlinks.add(
            cell,
            "" if internal else target,
            target.lstrip("#") if internal else "",
            str(config.get("display") or ""),
            str(config.get("tooltip") or ""),
        )
        if config.get("display") is not None:
            workbook.get_worksheet(sheet).cells[cell].value = config["display"]
        del hyperlink

    def add_image(self, workbook: Any, sheet: str, image: JSONDict) -> None:
        """Reject image writes because the adapter cannot create one-cell anchors."""
        self.unsupported_operation(
            "add_image",
            (
                "aspose-cells-foss writes only two-cell picture anchors, "
                "while the benchmark requires one-cell anchors."
            ),
        )

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: JSONDict) -> None:
        """Reject pivot writes because the FOSS API has no pivot-table model."""
        self.unsupported_operation(
            "add_pivot_table", "aspose-cells-foss exposes no pivot-table API."
        )

    def add_comment(self, workbook: Any, sheet: str, comment: JSONDict) -> None:
        """Add a legacy note/comment to a cell."""
        config = comment.get("comment", comment)
        cell = config.get("cell")
        text = config.get("text")
        if cell is None or text is None:
            raise ValueError("Comments require both cell and text")
        workbook.get_worksheet(sheet).cells[str(cell)].set_comment(
            str(text), str(config.get("author") or "")
        )

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        """Set frozen or split worksheet panes."""
        config = settings.get("freeze", settings)
        mode = config.get("mode")
        if mode not in {"freeze", "split"}:
            self.unsupported_operation(
                "set_freeze_panes",
                f"Unsupported pane mode {mode!r}; expected 'freeze' or 'split'.",
            )
        pane = workbook.get_worksheet(sheet).properties.pane
        pane.state = "frozen" if mode == "freeze" else "split"
        pane.x_split = config.get("x_split")
        pane.y_split = config.get("y_split")
        pane.top_left_cell = config.get("top_left_cell")
        pane.active_pane = config.get("active_pane")

    def add_named_range(self, workbook: Any, sheet: str, named_range: JSONDict) -> None:
        """Add a workbook- or sheet-scoped named range."""
        config = named_range.get("named_range", named_range)
        name = config.get("name")
        refers_to = config.get("refers_to")
        if not name or not refers_to:
            raise ValueError("Named ranges require name and refers_to")
        scope = config.get("scope") or "workbook"
        local_sheet_id = (
            self.get_sheet_names(workbook).index(sheet) if scope == "sheet" else None
        )
        workbook.properties.defined_names.add(
            DefinedName(str(name), str(refers_to).lstrip("="), local_sheet_id)
        )

    def add_table(self, workbook: Any, sheet: str, table: JSONDict) -> None:
        """Add a structured table and apply its supported metadata."""
        config = table.get("table", table)
        name = config.get("name")
        cell_range = config.get("ref") or config.get("range")
        if not name or not cell_range:
            raise ValueError("Tables require both name and ref")
        worksheet = workbook.get_worksheet(sheet)
        created = worksheet.tables.add_with_range(
            str(cell_range), str(name), bool(config.get("header_row", True))
        )
        created.show_totals_row = bool(config.get("totals_row", False))
        created.show_auto_filter = bool(config.get("autofilter", True))
        if config.get("style") is not None:
            created.table_style_info.name = str(config["style"])
        columns = config.get("columns")
        if isinstance(columns, list) and len(columns) == len(created.columns):
            for column, name_value in zip(created.columns, columns, strict=True):
                column.name = str(name_value)

    def set_sheet_protection(
        self, workbook: Any, sheet: str, settings: JSONDict
    ) -> None:
        """Apply raw OOXML-equivalent sheet protection settings."""
        config = settings.get("protection", settings)
        protection = workbook.get_worksheet(sheet).properties.protection
        property_names = {
            "protected": "sheet",
            "format_cells": "format_cells",
            "insert_rows": "insert_rows",
            "select_locked_cells": "select_locked_cells",
            "select_unlocked_cells": "select_unlocked_cells",
            "sort": "sort",
            "auto_filter": "auto_filter",
        }
        for input_name, property_name in property_names.items():
            if config.get(input_name) is not None:
                setattr(protection, property_name, bool(config[input_name]))
        if config.get("password") is not None:
            protection.password = str(config["password"])

    def set_page_setup(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        """Apply supported page setup and centered header/footer text."""
        config = settings.get("page_setup", settings)
        if config.get("print_title_rows") is not None:
            self.unsupported_operation(
                "set_page_setup", "aspose-cells-foss has no print-title-rows API."
            )
        worksheet = workbook.get_worksheet(sheet)
        page_setup = worksheet.properties.page_setup
        for key in ("orientation", "fit_to_width", "fit_to_height", "scale"):
            if config.get(key) is not None:
                setattr(page_setup, key, config[key])
        if config.get("header_center") is not None:
            worksheet.properties.header_footer.odd_header = (
                f"&C{config['header_center']}"
            )
        if config.get("footer_center") is not None:
            worksheet.properties.header_footer.odd_footer = (
                f"&C{config['footer_center']}"
            )

    def add_chart_with_anchor(self, workbook: Any, sheet: str, chart: JSONDict) -> None:
        """Add a line or bar chart with a two-cell anchor and first data series."""
        config = chart.get("chart", chart)
        chart_type_name = str(config.get("type") or "bar").upper()
        chart_type = getattr(aspose.ChartType, chart_type_name, None)
        if chart_type not in {aspose.ChartType.BAR, aspose.ChartType.LINE}:
            self.unsupported_operation(
                "add_chart_with_anchor",
                f"Unsupported chart type {config.get('type')!r}.",
            )
        data_ref = config.get("data_ref")
        if not data_ref:
            raise ValueError("add_chart_with_anchor requires data_ref")
        from_row, from_column, to_row, to_column = _coordinates(
            f"{config.get('from') or 'B2'}:{config.get('to') or 'H12'}"
        )
        charts = workbook.get_worksheet(sheet).charts
        chart_index = charts.add(chart_type, from_row, from_column, to_row, to_column)
        series_chart = charts[chart_index]
        series_chart.add_series(str(data_ref), config.get("categories_ref"))

    def save_workbook(self, workbook: Any, path: Path) -> None:
        """Save a workbook as XLSX."""
        workbook.save(str(path))
