"""Write-only adapter backed by the Apache POI external helper."""

from __future__ import annotations

from dataclasses import asdict
from datetime import date, datetime
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


class ApachePoiAdapter(WriteOnlyAdapter):
    """Minimal write-only adapter using the Apache POI helper.

    This is a cross-language context adapter, not part of the Python-first hero
    benchmark path yet. It currently targets generic workbook generation for a
    focused subset of write operations.
    """

    VERSION = "5.5.1"

    @classmethod
    def repo_root(cls) -> Path:
        return Path(__file__).resolve().parents[4]

    @classmethod
    def tool(cls) -> ExternalOracleTool:
        return external_oracle_catalog(repo_root=cls.repo_root())["apache-poi"]

    @classmethod
    def is_available(cls) -> bool:
        return cls.tool().is_available()

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="apache-poi",
            version=self.VERSION,
            language="java",
            capabilities={"write"},
        )

    def create_workbook(self) -> WorkbookData:
        return {
            "sheets": [],
            "values": [],
            "formats": [],
            "borders": [],
            "row_heights": [],
            "column_widths": [],
            "merges": [],
            "validations": [],
            "hyperlinks": [],
            "comments": [],
            "freeze_panes": [],
            "tables": [],
            "images": [],
            "named_ranges": [],
        }

    def add_sheet(self, workbook: WorkbookData, name: str) -> None:
        if name not in workbook["sheets"]:
            workbook["sheets"].append(name)

    def _ensure_sheet(self, workbook: WorkbookData, sheet: str) -> None:
        if sheet not in workbook["sheets"]:
            workbook["sheets"].append(sheet)

    def write_cell_value(
        self,
        workbook: WorkbookData,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        workbook["values"].append(
            {
                "sheet": sheet,
                "cell": cell,
                "type": str(value.type),
                "value": self._serialize_cell_value(value),
                "formula": value.formula or "",
            }
        )

    def write_cell_format(
        self,
        workbook: WorkbookData,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        workbook["formats"].append({"sheet": sheet, "cell": cell, **asdict(format)})

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
        workbook["column_widths"].append({"sheet": sheet, "column": column, "width": width})

    def merge_cells(self, workbook: WorkbookData, sheet: str, cell_range: str) -> None:
        self._ensure_sheet(workbook, sheet)
        workbook["merges"].append({"sheet": sheet, "range": cell_range})

    def add_conditional_format(
        self, workbook: WorkbookData, sheet: str, rule: dict[str, Any]
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = rule.get("cf_rule", rule)
        workbook.setdefault("conditional_formats", []).append({"sheet": sheet, **payload})

    def add_data_validation(
        self, workbook: WorkbookData, sheet: str, validation: dict[str, Any]
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = validation.get("validation", validation)
        if payload.get("values") and not payload.get("formula1"):
            payload = {
                **payload,
                "validation_type": payload.get("validation_type", "list"),
                "formula1": '"' + ",".join(str(v) for v in payload["values"]) + '"',
            }
        workbook["validations"].append({"sheet": sheet, **payload})

    def add_hyperlink(self, workbook: WorkbookData, sheet: str, link: dict[str, Any]) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = link.get("hyperlink", link)
        workbook["hyperlinks"].append({"sheet": sheet, **payload})

    def add_image(self, workbook: WorkbookData, sheet: str, image: dict[str, Any]) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = image.get("image", image)
        path = payload.get("path")
        resolved = str(Path(path).resolve()) if path else path
        workbook["images"].append({"sheet": sheet, **payload, "path": resolved})

    def add_pivot_table(self, workbook: WorkbookData, sheet: str, pivot: dict[str, Any]) -> None:
        raise NotImplementedError("apache-poi adapter does not yet support pivot writes")

    def add_comment(self, workbook: WorkbookData, sheet: str, comment: dict[str, Any]) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = comment.get("comment", comment)
        workbook["comments"].append({"sheet": sheet, **payload})

    def set_freeze_panes(
        self, workbook: WorkbookData, sheet: str, settings: dict[str, Any]
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = settings.get("freeze", settings)
        mode = str(payload.get("mode") or "freeze")
        x_split = int(payload.get("x_split", 0) or 0)
        y_split = int(payload.get("y_split", 0) or 0)
        top_left = payload.get("top_left_cell")
        if mode == "freeze" and top_left and x_split == 0 and y_split == 0:
            from openpyxl.utils.cell import coordinate_to_tuple

            row, col = coordinate_to_tuple(str(top_left))
            x_split = max(col - 1, 0)
            y_split = max(row - 1, 0)
        workbook["freeze_panes"].append(
            {
                "sheet": sheet,
                "mode": mode,
                "x_split": x_split,
                "y_split": y_split,
                "top_left_cell": top_left,
            }
        )

    def add_table(self, workbook: WorkbookData, sheet: str, table: dict[str, Any]) -> None:
        self._ensure_sheet(workbook, sheet)
        payload = table.get("table", table)
        workbook["tables"].append({"sheet": sheet, **payload})

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
                fixture_id="apache-poi-adapter",
                operation="write_adapter_workbook",
                output_path=path,
                payload=workbook,
            ),
        )
        if result.skipped:
            raise FileNotFoundError(result.notes or "Apache POI helper unavailable")
        if not result.passed:
            message = result.payload.get("message") or result.stderr or result.stdout
            raise RuntimeError(str(message))

    def _serialize_cell_value(self, value: CellValue) -> str:
        if value.type == CellType.BLANK:
            return ""
        if value.type == CellType.STRING:
            return "" if value.value is None else str(value.value)
        if value.type == CellType.NUMBER:
            return str(value.value)
        if value.type == CellType.BOOLEAN:
            return "true" if bool(value.value) else "false"
        if value.type == CellType.ERROR:
            return str(value.value)
        if value.type == CellType.FORMULA:
            return str(value.value or value.formula or "")
        if value.type == CellType.DATE and isinstance(value.value, date):
            return value.value.isoformat()
        if value.type == CellType.DATETIME and isinstance(value.value, datetime):
            return value.value.isoformat(timespec="seconds")
        return "" if value.value is None else str(value.value)
