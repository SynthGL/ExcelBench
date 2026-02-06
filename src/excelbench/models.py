"""Core data models for ExcelBench."""

from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum
from typing import Any


class CellType(StrEnum):
    STRING = "string"
    NUMBER = "number"
    DATE = "date"
    DATETIME = "datetime"
    BOOLEAN = "boolean"
    ERROR = "error"
    BLANK = "blank"
    FORMULA = "formula"


class OperationType(StrEnum):
    READ = "read"
    WRITE = "write"


class Importance(StrEnum):
    BASIC = "basic"
    EDGE = "edge"


class BorderStyle(StrEnum):
    NONE = "none"
    THIN = "thin"
    MEDIUM = "medium"
    THICK = "thick"
    DOUBLE = "double"
    DASHED = "dashed"
    DOTTED = "dotted"
    HAIR = "hair"
    MEDIUM_DASHED = "mediumDashed"
    DASH_DOT = "dashDot"
    MEDIUM_DASH_DOT = "mediumDashDot"
    DASH_DOT_DOT = "dashDotDot"
    MEDIUM_DASH_DOT_DOT = "mediumDashDotDot"
    SLANT_DASH_DOT = "slantDashDot"


@dataclass
class CellValue:
    """Represents a cell's value and type."""

    type: CellType
    value: Any = None
    formula: str | None = None  # If type is FORMULA, this holds the formula string


@dataclass
class CellFormat:
    """Represents text formatting for a cell."""

    bold: bool | None = None
    italic: bool | None = None
    underline: str | None = None  # "single", "double", etc.
    strikethrough: bool | None = None
    font_name: str | None = None
    font_size: float | None = None
    font_color: str | None = None  # Hex color like "#FF0000"
    bg_color: str | None = None  # Background/fill color
    number_format: str | None = None
    h_align: str | None = None
    v_align: str | None = None
    wrap: bool | None = None
    rotation: int | None = None
    indent: int | None = None


@dataclass
class BorderEdge:
    """Represents one edge of a cell border."""

    style: BorderStyle = BorderStyle.NONE
    color: str = "#000000"


@dataclass
class BorderInfo:
    """Represents all borders of a cell."""

    top: BorderEdge | None = None
    bottom: BorderEdge | None = None
    left: BorderEdge | None = None
    right: BorderEdge | None = None
    diagonal_up: BorderEdge | None = None
    diagonal_down: BorderEdge | None = None


@dataclass
class TestCase:
    """A single test case within a feature test file."""

    id: str
    label: str
    row: int
    expected: dict[str, Any]
    sheet: str | None = None
    cell: str | None = None
    importance: Importance = Importance.BASIC


@dataclass
class TestFile:
    """Metadata for a generated test file."""

    path: str
    feature: str
    tier: int
    file_format: str | None = None
    test_cases: list[TestCase] = field(default_factory=list)


@dataclass
class Manifest:
    """Index of all generated test files."""

    generated_at: datetime
    excel_version: str
    generator_version: str
    file_format: str = "xlsx"
    files: list[TestFile] = field(default_factory=list)


@dataclass
class TestResult:
    """Result of testing a single test case."""

    test_case_id: str
    operation: OperationType
    passed: bool
    expected: dict[str, Any]
    actual: dict[str, Any]
    notes: str | None = None
    importance: Importance | None = None
    label: str | None = None


# =============================================================================
# Tier 2 schema helpers
# =============================================================================


@dataclass
class MergeSpec:
    """Represents a merged cell range and expectations."""

    range: str
    top_left_value: Any | None = None
    non_top_left_nonempty: int | None = None
    top_left_bg_color: str | None = None
    non_top_left_bg_color: str | None = None

    def to_expected(self) -> dict[str, Any]:
        expected: dict[str, Any] = {
            "merged_range": self.range,
        }
        if self.top_left_value is not None:
            expected["top_left_value"] = self.top_left_value
        if self.non_top_left_nonempty is not None:
            expected["non_top_left_nonempty"] = self.non_top_left_nonempty
        if self.top_left_bg_color is not None:
            expected["top_left_bg_color"] = self.top_left_bg_color
        if self.non_top_left_bg_color is not None:
            expected["non_top_left_bg_color"] = self.non_top_left_bg_color
        return expected


@dataclass
class ConditionalFormatSpec:
    """Represents a conditional formatting rule expectation."""

    range: str
    rule_type: str
    operator: str | None = None
    formula: str | None = None
    priority: int | None = None
    stop_if_true: bool | None = None
    format: dict[str, Any] | None = None

    def to_expected(self) -> dict[str, Any]:
        expected: dict[str, Any] = {
            "cf_rule": {
                "range": self.range,
                "rule_type": self.rule_type,
            }
        }
        if self.operator is not None:
            expected["cf_rule"]["operator"] = self.operator
        if self.formula is not None:
            expected["cf_rule"]["formula"] = self.formula
        if self.priority is not None:
            expected["cf_rule"]["priority"] = self.priority
        if self.stop_if_true is not None:
            expected["cf_rule"]["stop_if_true"] = self.stop_if_true
        if self.format is not None:
            expected["cf_rule"]["format"] = self.format
        return expected


@dataclass
class DataValidationSpec:
    """Represents a data validation rule expectation."""

    range: str
    validation_type: str
    operator: str | None = None
    formula1: str | None = None
    formula2: str | None = None
    allow_blank: bool | None = None
    show_input: bool | None = None
    show_error: bool | None = None
    prompt_title: str | None = None
    prompt: str | None = None
    error_title: str | None = None
    error: str | None = None

    def to_expected(self) -> dict[str, Any]:
        expected: dict[str, Any] = {
            "validation": {
                "range": self.range,
                "validation_type": self.validation_type,
            }
        }
        v = expected["validation"]
        if self.operator is not None:
            v["operator"] = self.operator
        if self.formula1 is not None:
            v["formula1"] = self.formula1
        if self.formula2 is not None:
            v["formula2"] = self.formula2
        if self.allow_blank is not None:
            v["allow_blank"] = self.allow_blank
        if self.show_input is not None:
            v["show_input"] = self.show_input
        if self.show_error is not None:
            v["show_error"] = self.show_error
        if self.prompt_title is not None:
            v["prompt_title"] = self.prompt_title
        if self.prompt is not None:
            v["prompt"] = self.prompt
        if self.error_title is not None:
            v["error_title"] = self.error_title
        if self.error is not None:
            v["error"] = self.error
        return expected


@dataclass
class HyperlinkSpec:
    """Represents a hyperlink expectation."""

    cell: str
    target: str
    display: str | None = None
    tooltip: str | None = None
    internal: bool | None = None

    def to_expected(self) -> dict[str, Any]:
        expected: dict[str, Any] = {
            "hyperlink": {
                "cell": self.cell,
                "target": self.target,
            }
        }
        if self.display is not None:
            expected["hyperlink"]["display"] = self.display
        if self.tooltip is not None:
            expected["hyperlink"]["tooltip"] = self.tooltip
        if self.internal is not None:
            expected["hyperlink"]["internal"] = self.internal
        return expected


@dataclass
class ImageSpec:
    """Represents an image expectation."""

    cell: str
    path: str
    anchor: str | None = None
    offset: tuple[int, int] | None = None
    alt_text: str | None = None

    def to_expected(self) -> dict[str, Any]:
        expected: dict[str, Any] = {
            "image": {
                "cell": self.cell,
                "path": self.path,
            }
        }
        if self.anchor is not None:
            expected["image"]["anchor"] = self.anchor
        if self.offset is not None:
            expected["image"]["offset"] = list(self.offset)
        if self.alt_text is not None:
            expected["image"]["alt_text"] = self.alt_text
        return expected


@dataclass
class PivotSpec:
    """Represents a pivot table expectation."""

    name: str
    source_range: str
    target_cell: str
    row_fields: list[str]
    column_fields: list[str]
    data_fields: list[str]
    filter_fields: list[str] | None = None

    def to_expected(self) -> dict[str, Any]:
        expected = {
            "pivot": {
                "name": self.name,
                "source_range": self.source_range,
                "target_cell": self.target_cell,
                "row_fields": self.row_fields,
                "column_fields": self.column_fields,
                "data_fields": self.data_fields,
            }
        }
        if self.filter_fields is not None:
            expected["pivot"]["filter_fields"] = self.filter_fields
        return expected


@dataclass
class CommentSpec:
    """Represents a comment/note expectation."""

    cell: str
    text: str
    author: str | None = None
    threaded: bool | None = None

    def to_expected(self) -> dict[str, Any]:
        expected: dict[str, Any] = {
            "comment": {
                "cell": self.cell,
                "text": self.text,
            }
        }
        if self.author is not None:
            expected["comment"]["author"] = self.author
        if self.threaded is not None:
            expected["comment"]["threaded"] = self.threaded
        return expected


@dataclass
class FreezePaneSpec:
    """Represents freeze/split panes expectation."""

    mode: str
    top_left_cell: str | None = None
    x_split: int | None = None
    y_split: int | None = None
    active_pane: str | None = None

    def to_expected(self) -> dict[str, Any]:
        expected: dict[str, Any] = {
            "freeze": {
                "mode": self.mode,
            }
        }
        if self.top_left_cell is not None:
            expected["freeze"]["top_left_cell"] = self.top_left_cell
        if self.x_split is not None:
            expected["freeze"]["x_split"] = self.x_split
        if self.y_split is not None:
            expected["freeze"]["y_split"] = self.y_split
        if self.active_pane is not None:
            expected["freeze"]["active_pane"] = self.active_pane
        return expected


@dataclass
class FeatureScore:
    """Fidelity score for a feature."""

    feature: str
    library: str
    read_score: int | None = None  # 0-3, None if not applicable
    write_score: int | None = None  # 0-3, None if not applicable
    test_results: list[TestResult] = field(default_factory=list)
    notes: str | None = None


@dataclass
class LibraryInfo:
    """Information about a library being tested."""

    name: str
    version: str
    language: str  # "python" or "rust"
    capabilities: set[str] = field(default_factory=set)  # {"read", "write"}


@dataclass
class BenchmarkMetadata:
    """Metadata about a benchmark run."""

    benchmark_version: str
    run_date: datetime
    excel_version: str
    platform: str
    profile: str = "xlsx"


@dataclass
class BenchmarkResults:
    """Complete benchmark results."""

    metadata: BenchmarkMetadata
    libraries: dict[str, LibraryInfo] = field(default_factory=dict)
    scores: list[FeatureScore] = field(default_factory=list)
