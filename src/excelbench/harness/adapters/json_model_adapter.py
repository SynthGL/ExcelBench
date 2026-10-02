"""Shared adapter base for libraries driven through an external JSON helper.

Non-Python libraries (SheetJS, ExcelJS, LibreOffice UNO) cannot hand Python a
live workbook object. Their helpers instead speak a small JSON protocol over
stdin/stdout (the same transport as ``external_oracles``):

``describe``
    Returns ``{"library", "version", "capabilities", "unsupported_write"}``.
``read_model`` (``input_path``)
    Loads the workbook with the library's own read API and returns
    ``{"model": ...}``: a normalized snapshot of everything the library
    exposed. Values the library does not surface are simply absent, and
    features it cannot read at all are listed in ``model["unsupported"]``.
``write_model`` (``output_path``, ``payload.ops``)
    Replays the recorded adapter calls through the library's write API.
``mutate`` (``input_path``, ``output_path``, ``payload.mutations``)
    Opens an existing workbook, sets cell values, and saves it.

The Python side only translates the snapshot into ExcelBench models; it never
fills in values the helper did not report.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, ClassVar

from excelbench.harness.adapters.base import ExcelAdapter
from excelbench.harness.external_oracles import (
    ExternalOracleRequest,
    ExternalOracleTool,
    run_external_oracle,
)
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

_FORMAT_FIELDS = (
    "bold",
    "italic",
    "underline",
    "strikethrough",
    "font_name",
    "font_size",
    "font_color",
    "bg_color",
    "number_format",
    "h_align",
    "v_align",
    "wrap",
    "rotation",
    "indent",
)
_BORDER_EDGES = ("top", "bottom", "left", "right", "diagonal_up", "diagonal_down")
# Stored column widths carry font-metric padding; strip the same known paddings
# the openpyxl adapter strips so every reader reports display character width.
_COLUMN_WIDTH_PADDINGS = (0.83203125, 0.7109375)


class HelperError(RuntimeError):
    """Raised when an external adapter helper reports a failure."""


@dataclass
class ModelWorkbook:
    """Read-side handle: the normalized snapshot returned by ``read_model``."""

    model: JSONDict
    sheets: dict[str, JSONDict] = field(init=False)

    def __post_init__(self) -> None:
        self.sheets = {
            str(sheet["name"]): sheet for sheet in self.model.get("sheets", [])
        }


@dataclass
class OpLog:
    """Write-side handle: adapter calls recorded for one ``write_model`` replay."""

    ops: list[JSONDict] = field(default_factory=list)


def _json_safe(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat(timespec="seconds")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(k): _json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(v) for v in value]
    if hasattr(value, "value") and type(value).__module__.startswith("excelbench"):
        return value.value
    return value


def _cell_value_from_model(entry: JSONDict | None) -> CellValue:
    if not entry:
        return CellValue(type=CellType.BLANK)
    cell_type = CellType(str(entry.get("type", "blank")))
    value = entry.get("value")
    formula = entry.get("formula")
    if cell_type == CellType.BLANK:
        return CellValue(type=CellType.BLANK)
    if cell_type == CellType.DATE and isinstance(value, str):
        return CellValue(type=cell_type, value=date.fromisoformat(value[:10]))
    if cell_type == CellType.DATETIME and isinstance(value, str):
        return CellValue(type=cell_type, value=datetime.fromisoformat(value[:19]))
    if cell_type == CellType.FORMULA:
        text = str(formula if formula is not None else value or "")
        if text and not text.startswith("="):
            text = f"={text}"
        return CellValue(type=cell_type, value=text, formula=text)
    return CellValue(type=cell_type, value=value)


def _border_edge(entry: Any) -> BorderEdge | None:
    if not isinstance(entry, dict):
        return None
    style = entry.get("style")
    if not style or style == "none":
        return None
    try:
        border_style = BorderStyle(str(style))
    except ValueError:
        border_style = BorderStyle.THIN
    color = entry.get("color") or "#000000"
    return BorderEdge(style=border_style, color=str(color))


def _normalize_column_width(width: Any) -> float | None:
    if not isinstance(width, (int, float)):
        return None
    width_f = float(width)
    frac = width_f % 1
    for padding in _COLUMN_WIDTH_PADDINGS:
        if abs(frac - padding) < 0.01:
            width_f -= padding
            break
    return round(width_f, 4)


def _normalize_refers_to(value: Any) -> str:
    if value is None:
        return ""
    raw = str(value).lstrip("=")
    if "!" not in raw:
        return raw
    sheet_part, addr = raw.split("!", 1)
    if sheet_part.startswith("'") and sheet_part.endswith("'") and len(sheet_part) >= 2:
        sheet_part = sheet_part[1:-1].replace("''", "'")
    return f"{sheet_part}!{addr}"


class JsonModelAdapter(ExcelAdapter):
    """Adapter whose reads/writes are executed by an external JSON helper."""

    LIBRARY_NAME: ClassVar[str] = ""
    LANGUAGE: ClassVar[str] = "external"
    TIMEOUT_SECONDS: ClassVar[float] = 120.0
    _describe_cache: ClassVar[dict[str, JSONDict]] = {}

    @classmethod
    def tool(cls) -> ExternalOracleTool:
        """Return the helper descriptor."""
        raise NotImplementedError

    @classmethod
    def is_available(cls) -> bool:
        """Return whether the helper runtime and dependencies are installed."""
        return cls.tool().is_available()

    @classmethod
    def call_helper(
        cls,
        operation: str,
        *,
        payload: JSONDict | None = None,
        input_path: Path | None = None,
        output_path: Path | None = None,
    ) -> JSONDict:
        """Run one helper operation and return its JSON payload."""
        tool = cls.tool()
        result = run_external_oracle(
            tool,
            ExternalOracleRequest(
                fixture_id=f"{tool.name}-adapter",
                operation=operation,
                payload=payload or {},
                input_path=input_path,
                output_path=output_path,
            ),
            timeout_seconds=cls.TIMEOUT_SECONDS,
        )
        if not result.passed:
            message = (
                result.payload.get("message")
                or result.notes
                or result.stderr.strip()
                or result.stdout.strip()
                or "helper failed without output"
            )
            raise HelperError(f"{tool.name} {operation} failed: {message}")
        return result.payload

    @classmethod
    def describe(cls) -> JSONDict:
        """Return cached library identity and write-support declarations."""
        key = cls.__name__
        if key not in cls._describe_cache:
            cls._describe_cache[key] = cls.call_helper("describe")
        return cls._describe_cache[key]

    @property
    def name(self) -> str:
        # Static so adapter selection never spawns the helper.
        return self.LIBRARY_NAME

    @property
    def info(self) -> LibraryInfo:
        described = self.describe()
        if described.get("library") != self.LIBRARY_NAME:
            raise HelperError(
                f"helper identifies as {described.get('library')!r}, expected {self.LIBRARY_NAME!r}"
            )
        return LibraryInfo(
            name=self.LIBRARY_NAME,
            version=str(described["version"]),
            language=self.LANGUAGE,
            capabilities=set(described.get("capabilities", ["read", "write"])),
        )

    # =========================================================================
    # Read side
    # =========================================================================

    def open_workbook(self, path: Path) -> ModelWorkbook:
        payload = self.call_helper("read_model", input_path=Path(path).resolve())
        return ModelWorkbook(payload["model"])

    def close_workbook(self, workbook: Any) -> None:
        return None

    def get_sheet_names(self, workbook: ModelWorkbook) -> list[str]:
        return list(workbook.sheets)

    def _sheet(self, workbook: ModelWorkbook, sheet: str) -> JSONDict:
        try:
            return workbook.sheets[sheet]
        except KeyError as exc:
            raise KeyError(f"Worksheet not found: {sheet}") from exc

    def _require_read(self, workbook: ModelWorkbook, feature: str) -> None:
        reason = workbook.model.get("unsupported", {}).get(feature)
        if reason:
            self.unsupported_operation(feature, str(reason))

    def _sheet_list(
        self, workbook: ModelWorkbook, sheet: str, key: str
    ) -> list[JSONDict]:
        self._require_read(workbook, key)
        return [dict(item) for item in self._sheet(workbook, sheet).get(key, [])]

    def _cell_entry(self, workbook: ModelWorkbook, sheet: str, cell: str) -> JSONDict:
        entry: JSONDict = (
            self._sheet(workbook, sheet).get("cells", {}).get(cell.upper(), {})
        )
        return entry

    def read_cell_value(
        self, workbook: ModelWorkbook, sheet: str, cell: str
    ) -> CellValue:
        self._require_read(workbook, "cell_values")
        return _cell_value_from_model(
            self._cell_entry(workbook, sheet, cell).get("value")
        )

    def read_cell_format(
        self, workbook: ModelWorkbook, sheet: str, cell: str
    ) -> CellFormat:
        self._require_read(workbook, "cell_format")
        fmt = self._cell_entry(workbook, sheet, cell).get("format") or {}
        return CellFormat(**{key: fmt.get(key) for key in _FORMAT_FIELDS})

    def read_cell_border(
        self, workbook: ModelWorkbook, sheet: str, cell: str
    ) -> BorderInfo:
        self._require_read(workbook, "cell_border")
        border = self._cell_entry(workbook, sheet, cell).get("border") or {}
        return BorderInfo(
            **{edge: _border_edge(border.get(edge)) for edge in _BORDER_EDGES}
        )

    def read_row_height(
        self, workbook: ModelWorkbook, sheet: str, row: int
    ) -> float | None:
        self._require_read(workbook, "row_heights")
        height = self._sheet(workbook, sheet).get("row_heights", {}).get(str(row))
        return float(height) if isinstance(height, (int, float)) else None

    def read_column_width(
        self, workbook: ModelWorkbook, sheet: str, column: str
    ) -> float | None:
        self._require_read(workbook, "column_widths")
        widths = self._sheet(workbook, sheet).get("column_widths", {})
        return _normalize_column_width(widths.get(column.upper()))

    def read_merged_ranges(self, workbook: ModelWorkbook, sheet: str) -> list[str]:
        self._require_read(workbook, "merged_ranges")
        return [
            str(rng) for rng in self._sheet(workbook, sheet).get("merged_ranges", [])
        ]

    def read_conditional_formats(
        self, workbook: ModelWorkbook, sheet: str
    ) -> list[JSONDict]:
        return self._sheet_list(workbook, sheet, "conditional_formats")

    def read_data_validations(
        self, workbook: ModelWorkbook, sheet: str
    ) -> list[JSONDict]:
        return self._sheet_list(workbook, sheet, "data_validations")

    def read_hyperlinks(self, workbook: ModelWorkbook, sheet: str) -> list[JSONDict]:
        return self._sheet_list(workbook, sheet, "hyperlinks")

    def read_images(self, workbook: ModelWorkbook, sheet: str) -> list[JSONDict]:
        return self._sheet_list(workbook, sheet, "images")

    def read_pivot_tables(self, workbook: ModelWorkbook, sheet: str) -> list[JSONDict]:
        return self._sheet_list(workbook, sheet, "pivot_tables")

    def read_comments(self, workbook: ModelWorkbook, sheet: str) -> list[JSONDict]:
        return self._sheet_list(workbook, sheet, "comments")

    def read_freeze_panes(self, workbook: ModelWorkbook, sheet: str) -> JSONDict:
        self._require_read(workbook, "freeze_panes")
        return dict(self._sheet(workbook, sheet).get("freeze_panes") or {})

    def read_named_ranges(self, workbook: ModelWorkbook, sheet: str) -> list[JSONDict]:
        self._require_read(workbook, "named_ranges")
        out: list[JSONDict] = []
        for item in workbook.model.get("named_ranges", []):
            scope = item.get("scope") or "workbook"
            if scope == "sheet" and item.get("sheet") != sheet:
                continue
            out.append(
                {
                    "name": str(item.get("name")),
                    "scope": scope,
                    "refers_to": _normalize_refers_to(item.get("refers_to")),
                }
            )
        return out

    def read_tables(self, workbook: ModelWorkbook, sheet: str) -> list[JSONDict]:
        return self._sheet_list(workbook, sheet, "tables")

    def read_sheet_protection(self, workbook: ModelWorkbook, sheet: str) -> JSONDict:
        self._require_read(workbook, "sheet_protection")
        return dict(self._sheet(workbook, sheet).get("sheet_protection") or {})

    def read_page_setup(self, workbook: ModelWorkbook, sheet: str) -> JSONDict:
        self._require_read(workbook, "page_setup")
        return dict(self._sheet(workbook, sheet).get("page_setup") or {})

    def read_chart_anchors(self, workbook: ModelWorkbook, sheet: str) -> list[JSONDict]:
        return self._sheet_list(workbook, sheet, "chart_anchors")

    # =========================================================================
    # Write side
    # =========================================================================

    def _record(self, workbook: OpLog, op: str, **fields: Any) -> None:
        reason = self.describe().get("unsupported_write", {}).get(op)
        if reason:
            self.unsupported_operation(op, str(reason))
        workbook.ops.append({"op": op, **_json_safe(fields)})

    def create_workbook(self) -> OpLog:
        return OpLog()

    def add_sheet(self, workbook: OpLog, name: str) -> None:
        self._record(workbook, "add_sheet", name=name)

    def write_cell_value(
        self, workbook: OpLog, sheet: str, cell: str, value: CellValue
    ) -> None:
        self._record(
            workbook,
            "cell_value",
            sheet=sheet,
            cell=cell,
            type=value.type.value,
            value=value.value,
            formula=value.formula,
        )

    def write_cell_format(
        self, workbook: OpLog, sheet: str, cell: str, format: CellFormat
    ) -> None:
        fmt = {key: getattr(format, key) for key in _FORMAT_FIELDS}
        self._record(
            workbook,
            "cell_format",
            sheet=sheet,
            cell=cell,
            format={k: v for k, v in fmt.items() if v is not None},
        )

    def write_cell_border(
        self, workbook: OpLog, sheet: str, cell: str, border: BorderInfo
    ) -> None:
        edges: JSONDict = {}
        for edge_name in _BORDER_EDGES:
            edge = getattr(border, edge_name)
            if edge is not None and edge.style != BorderStyle.NONE:
                edges[edge_name] = {"style": edge.style.value, "color": edge.color}
        self._record(workbook, "cell_border", sheet=sheet, cell=cell, border=edges)

    def set_row_height(
        self, workbook: OpLog, sheet: str, row: int, height: float
    ) -> None:
        self._record(workbook, "row_height", sheet=sheet, row=row, height=height)

    def set_column_width(
        self, workbook: OpLog, sheet: str, column: str, width: float
    ) -> None:
        self._record(workbook, "column_width", sheet=sheet, column=column, width=width)

    def merge_cells(self, workbook: OpLog, sheet: str, cell_range: str) -> None:
        self._record(workbook, "merge", sheet=sheet, range=cell_range)

    def add_conditional_format(
        self, workbook: OpLog, sheet: str, rule: JSONDict
    ) -> None:
        self._record(
            workbook, "conditional_format", sheet=sheet, rule=rule.get("cf_rule", rule)
        )

    def add_data_validation(
        self, workbook: OpLog, sheet: str, validation: JSONDict
    ) -> None:
        self._record(
            workbook,
            "data_validation",
            sheet=sheet,
            validation=validation.get("validation", validation),
        )

    def add_hyperlink(self, workbook: OpLog, sheet: str, link: JSONDict) -> None:
        self._record(
            workbook, "hyperlink", sheet=sheet, link=link.get("hyperlink", link)
        )

    def add_image(self, workbook: OpLog, sheet: str, image: JSONDict) -> None:
        data = dict(image.get("image", image))
        if data.get("path"):
            data["path"] = str(Path(str(data["path"])).resolve())
        self._record(workbook, "image", sheet=sheet, image=data)

    def add_pivot_table(self, workbook: OpLog, sheet: str, pivot: JSONDict) -> None:
        self._record(workbook, "pivot", sheet=sheet, pivot=pivot.get("pivot", pivot))

    def add_comment(self, workbook: OpLog, sheet: str, comment: JSONDict) -> None:
        self._record(
            workbook, "comment", sheet=sheet, comment=comment.get("comment", comment)
        )

    def set_freeze_panes(self, workbook: OpLog, sheet: str, settings: JSONDict) -> None:
        self._record(
            workbook, "freeze", sheet=sheet, settings=settings.get("freeze", settings)
        )

    def add_named_range(
        self, workbook: OpLog, sheet: str, named_range: JSONDict
    ) -> None:
        self._record(
            workbook,
            "named_range",
            sheet=sheet,
            named_range=named_range.get("named_range", named_range),
        )

    def add_table(self, workbook: OpLog, sheet: str, table: JSONDict) -> None:
        self._record(workbook, "table", sheet=sheet, table=table.get("table", table))

    def set_sheet_protection(
        self, workbook: OpLog, sheet: str, settings: JSONDict
    ) -> None:
        self._record(
            workbook,
            "protection",
            sheet=sheet,
            settings=settings.get("protection", settings),
        )

    def set_page_setup(self, workbook: OpLog, sheet: str, settings: JSONDict) -> None:
        self._record(
            workbook,
            "page_setup",
            sheet=sheet,
            settings=settings.get("page_setup", settings),
        )

    def add_chart_with_anchor(
        self, workbook: OpLog, sheet: str, chart: JSONDict
    ) -> None:
        self._record(workbook, "chart", sheet=sheet, chart=chart.get("chart", chart))

    def save_workbook(self, workbook: OpLog, path: Path) -> None:
        output = Path(path).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        self.call_helper(
            "write_model", payload={"ops": workbook.ops}, output_path=output
        )
        if not output.is_file():
            raise HelperError(
                f"{self.name} write_model reported success but wrote no file"
            )

    # =========================================================================
    # Modify side (template-mutation lane)
    # =========================================================================

    @classmethod
    def mutate_file(
        cls, template: Path, output: Path, mutations: list[JSONDict]
    ) -> None:
        """Open ``template`` with the library, set cell values, save to ``output``."""
        output = Path(output).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        cls.call_helper(
            "mutate",
            payload={"mutations": json.loads(json.dumps(_json_safe(mutations)))},
            input_path=Path(template).resolve(),
            output_path=output,
        )
        if not output.is_file():
            raise HelperError(
                f"{cls.__name__} mutate reported success but wrote no file"
            )
