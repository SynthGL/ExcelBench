"""Write-oriented adapter backed by the Excelize external helper."""

from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
from typing import Any

from excelbench.harness.adapters.base import WriteOnlyAdapter
from excelbench.harness.external_oracles import (
    ExternalOracleRequest,
    ExternalOracleTool,
    external_oracle_catalog,
    run_external_oracle,
)
from excelbench.models import BorderInfo, CellFormat, CellType, CellValue, LibraryInfo

WorkbookData = dict[str, Any]


class ExcelizeAdapter(WriteOnlyAdapter):
    """Cross-language write-first adapter backed by the Excelize Go helper."""

    @classmethod
    def repo_root(cls) -> Path:
        return Path(__file__).resolve().parents[4]

    @classmethod
    def tool(cls) -> ExternalOracleTool:
        return external_oracle_catalog(repo_root=cls.repo_root())["excelize"]

    @classmethod
    def is_available(cls) -> bool:
        return cls.tool().is_available()

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="excelize",
            version="go-helper",
            language="go",
            capabilities={"write"},
        )

    def create_workbook(self) -> WorkbookData:
        return {
            "sheets": [],
            "cells": [],
            "formats": [],
            "borders": [],
            "columns": [],
            "row_heights": [],
            "merges": [],
            "validations": [],
            "hyperlinks": [],
            "comments": [],
            "panes": [],
            "named_ranges": [],
            "tables": [],
            "conditional_formats": [],
            "charts": [],
            "pivots": [],
            "slicers": [],
            "pictures": [],
        }

    def add_sheet(self, workbook: WorkbookData, name: str) -> None:
        if not any(sheet["name"] == name for sheet in workbook["sheets"]):
            workbook["sheets"].append({"name": name})

    def _ensure_sheet(self, workbook: WorkbookData, sheet: str) -> None:
        self.add_sheet(workbook, sheet)

    def write_cell_value(
        self,
        workbook: WorkbookData,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        entry: dict[str, Any] = {"sheet": sheet, "cell": cell}
        if value.type == CellType.FORMULA:
            entry["type"] = "formula"
            entry["formula"] = value.formula or str(value.value or "")
            entry["value"] = value.value
        elif value.type == CellType.BLANK or (value.type == CellType.STRING and value.value == ""):
            entry["type"] = "blank"
            entry["value"] = None
        else:
            entry["type"] = str(value.type)
            entry["value"] = self._serialize_value(value)
        workbook["cells"].append(entry)

    def write_cell_format(
        self,
        workbook: WorkbookData,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        workbook["formats"].append(
            {
                "sheet": sheet,
                "cell": cell,
                "bold": format.bold,
                "italic": format.italic,
                "underline": format.underline,
                "strikethrough": format.strikethrough,
                "font_name": format.font_name,
                "font_size": format.font_size,
                "font_color": format.font_color,
                "bg_color": format.bg_color,
                "number_format": format.number_format,
                "h_align": format.h_align,
                "v_align": format.v_align,
                "wrap": format.wrap,
                "rotation": format.rotation,
                "indent": format.indent,
            }
        )

    def write_cell_border(
        self,
        workbook: WorkbookData,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        workbook["borders"].append({"sheet": sheet, "cell": cell, "border": asdict(border)})

    def set_row_height(
        self,
        workbook: WorkbookData,
        sheet: str,
        row: int,
        height: float,
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        workbook["row_heights"].append({"sheet": sheet, "row": row, "height": height})

    def set_column_width(
        self,
        workbook: WorkbookData,
        sheet: str,
        column: str,
        width: float,
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        workbook["columns"].append({"sheet": sheet, "start": column, "end": column, "width": width})

    def merge_cells(self, workbook: WorkbookData, sheet: str, cell_range: str) -> None:
        self._ensure_sheet(workbook, sheet)
        workbook["merges"].append({"sheet": sheet, "range": cell_range})

    def add_conditional_format(
        self, workbook: WorkbookData, sheet: str, rule: dict[str, Any]
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = rule.get("cf_rule", rule)
        workbook["conditional_formats"].append(
            {
                "sheet": sheet,
                "range": payload.get("range"),
                "type": self._map_conditional_format_type(payload.get("rule_type")),
                "criteria": self._map_conditional_format_criteria(
                    payload.get("rule_type"), payload.get("operator"), payload.get("formula")
                ),
                "value": self._map_conditional_format_value(payload),
                "stop_if_true": payload.get("stop_if_true", False),
                "bg_color": (payload.get("format") or {}).get("bg_color", ""),
            }
        )

    def add_data_validation(
        self, workbook: WorkbookData, sheet: str, validation: dict[str, Any]
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = validation.get("validation", validation)
        workbook["validations"].append({"sheet": sheet, **payload})

    def add_hyperlink(self, workbook: WorkbookData, sheet: str, link: dict[str, Any]) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = link.get("hyperlink", link)
        workbook["hyperlinks"].append({"sheet": sheet, **payload})

    def add_image(self, workbook: WorkbookData, sheet: str, image: dict[str, Any]) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = image.get("image", image)
        workbook["pictures"].append({"sheet": sheet, **payload})

    def add_pivot_table(self, workbook: WorkbookData, sheet: str, pivot: dict[str, Any]) -> None:
        payload = pivot.get("pivot", pivot)
        target_cell = payload.get("target_cell")
        target_range = payload.get("range")
        if target_range is None and target_cell is not None:
            target_range = f"{sheet}!{target_cell}:{target_cell}"
        workbook["pivots"].append(
            {
                "data_range": payload.get("source_range") or payload.get("data_range"),
                "range": target_range,
                "name": payload.get("name"),
                "rows": payload.get("rows")
                or [{"name": name} for name in payload.get("row_fields", [])],
                "columns": payload.get("columns")
                or [{"name": name} for name in payload.get("column_fields", [])],
                "data": payload.get("data")
                or [{"name": name, "subtotal": "Sum"} for name in payload.get("data_fields", [])],
                "filters": payload.get("filters")
                or [{"name": name} for name in payload.get("filter_fields", [])],
            }
        )

    def add_comment(self, workbook: WorkbookData, sheet: str, comment: dict[str, Any]) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = comment.get("comment", comment)
        workbook["comments"].append({"sheet": sheet, **payload})

    def set_freeze_panes(
        self, workbook: WorkbookData, sheet: str, settings: dict[str, Any]
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = settings.get("freeze", settings)
        workbook["panes"].append({"sheet": sheet, **payload})

    def add_table(self, workbook: WorkbookData, sheet: str, table: dict[str, Any]) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = table.get("table", table)
        workbook["tables"].append(
            {
                "sheet": sheet,
                "range": payload.get("ref") or payload.get("range"),
                "name": payload.get("name"),
                "style": payload.get("style"),
                "show_header_row": payload.get("header_row", True),
                "show_row_stripes": True,
                "totals_row": payload.get("totals_row", False),
                "columns": payload.get("columns", []),
                "autofilter": payload.get("autofilter", False),
            }
        )

    def add_named_range(
        self, workbook: WorkbookData, sheet: str, named_range: dict[str, Any]
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = named_range.get("named_range", named_range)
        workbook["named_ranges"].append({"sheet": sheet, **payload})

    def save_workbook(self, workbook: WorkbookData, path: Path) -> None:
        tool = self.tool()
        result = run_external_oracle(
            tool,
            ExternalOracleRequest(
                fixture_id="excelize-adapter",
                operation="write_fixture",
                output_path=path,
                payload=workbook,
            ),
            timeout_seconds=120,
        )
        if result.skipped:
            raise FileNotFoundError(result.notes or "Excelize helper unavailable")
        if not result.passed:
            message = result.payload.get("message") or result.stderr or result.stdout
            raise RuntimeError(str(message))

    def _serialize_value(self, value: CellValue) -> Any:
        if value.type == CellType.BLANK:
            return None
        if value.type == CellType.DATE and value.value is not None:
            return value.value.isoformat()
        if value.type == CellType.DATETIME and value.value is not None:
            return value.value.isoformat(timespec="seconds")
        return value.value

    def _map_conditional_format_type(self, value: Any) -> str:
        rule_type = str(value or "")
        if rule_type == "colorScale":
            return "3_color_scale"
        if rule_type == "dataBar":
            return "data_bar"
        if rule_type == "iconSet":
            return "icon_set"
        if rule_type == "expression":
            return "formula"
        return "cell"

    def _map_conditional_format_criteria(self, rule_type: Any, operator: Any, formula: Any) -> str:
        if str(rule_type or "") == "expression":
            return str(formula or "").lstrip("=")
        mapping = {
            "greaterThan": ">",
            "greaterThanOrEqual": ">=",
            "lessThan": "<",
            "lessThanOrEqual": "<=",
            "equal": "==",
            "notEqual": "!=",
            "between": "between",
            "notBetween": "not between",
        }
        return mapping.get(str(operator or ""), ">")

    def _map_conditional_format_value(self, payload: dict[str, Any]) -> str:
        if str(payload.get("rule_type") or "") == "expression":
            return ""
        return (payload.get("formula") or payload.get("value") or "0").lstrip("=")
