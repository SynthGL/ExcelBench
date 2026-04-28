"""Generate local external-oracle fixture packs."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from excelbench.harness.external_oracles import (
    ExternalOracleRequest,
    ExternalOracleResult,
    external_oracle_catalog,
    run_external_oracle,
)

JSONDict = dict[str, Any]


@dataclass(frozen=True)
class ExternalFixtureSpec:
    """Definition of a local external-oracle fixture.

    Args:
        fixture_id: Stable fixture identifier.
        tool: Source helper name.
        filename: Workbook filename to generate.
        payload: JSON payload for the source helper.
        expected_parts: OOXML package parts expected in the generated workbook.
        notes: Human-readable reason this fixture exists.
    """

    fixture_id: str
    tool: str
    filename: str
    payload: JSONDict
    expected_parts: tuple[str, ...]
    notes: str


@dataclass(frozen=True)
class FixtureGenerationResult:
    """Result for one generated external fixture."""

    fixture_id: str
    tool: str
    workbook_path: Path
    write_result: ExternalOracleResult
    expected_parts: tuple[str, ...]
    missing_parts: tuple[str, ...]
    validations: tuple[ExternalOracleResult, ...] = field(default_factory=tuple)

    @property
    def passed(self) -> bool:
        """Return whether generation and requested validations passed."""
        return (
            self.write_result.passed
            and not self.missing_parts
            and all(result.passed or result.skipped for result in self.validations)
        )

    def to_json_dict(self, output_root: Path) -> JSONDict:
        """Convert the result to a manifest entry."""
        return {
            "fixture_id": self.fixture_id,
            "tool": self.tool,
            "workbook": str(self.workbook_path.relative_to(output_root)),
            "passed": self.passed,
            "expected_parts": list(self.expected_parts),
            "missing_parts": list(self.missing_parts),
            "write_result": _oracle_result_to_json(self.write_result),
            "validations": [_oracle_result_to_json(result) for result in self.validations],
        }


def excelize_fixture_specs() -> list[ExternalFixtureSpec]:
    """Return the initial Excelize fixture pack."""
    return [
        ExternalFixtureSpec(
            fixture_id="excelize_sales_pivot_slicer_chart",
            tool="excelize",
            filename="excelize-sales-pivot-slicer-chart.xlsx",
            payload={
                "sheets": [{"name": "Data"}, {"name": "Pivot"}],
                "cells": _sales_cells(),
                "columns": [{"sheet": "Data", "start": "A", "end": "C", "width": 16}],
                "tables": [{"sheet": "Data", "range": "A1:C6", "name": "SalesTable"}],
                "conditional_formats": [
                    {"sheet": "Data", "range": "C2:C6", "type": "3_color_scale"},
                    {"sheet": "Data", "range": "C2:C6", "type": "data_bar"},
                    {
                        "sheet": "Data",
                        "range": "C2:C6",
                        "type": "icon_set",
                        "icon_style": "3TrafficLights1",
                    },
                ],
                "charts": [
                    {
                        "sheet": "Data",
                        "cell": "E2",
                        "type": "col",
                        "title": "Sales by row",
                        "categories": "Data!$A$2:$A$6",
                        "values": "Data!$C$2:$C$6",
                        "show_values": True,
                        "alt_text": "Sales by row chart",
                    }
                ],
                "pivots": [
                    {
                        "data_range": "Data!A1:C6",
                        "range": "Pivot!A3:E12",
                        "name": "SalesPivot",
                        "rows": [{"name": "Region"}],
                        "columns": [{"name": "Product"}],
                        "data": [{"name": "Sales", "subtotal": "Sum"}],
                        "show_row_stripes": True,
                    }
                ],
                "slicers": [
                    {
                        "sheet": "Data",
                        "name": "Region",
                        "cell": "E20",
                        "table_sheet": "Data",
                        "table_name": "SalesTable",
                        "caption": "Region",
                    }
                ],
                "pictures": [{"sheet": "Data", "cell": "H2", "name": "Pixel"}],
            },
            expected_parts=(
                "xl/tables/table1.xml",
                "xl/pivotTables/pivotTable1.xml",
                "xl/pivotCache/pivotCacheDefinition1.xml",
                "xl/slicers/slicer1.xml",
                "xl/slicerCaches/slicerCache1.xml",
                "xl/charts/chart1.xml",
                "xl/drawings/drawing1.xml",
            ),
            notes="Pivots, slicers, charts, CF, tables, and drawings in one workbook.",
        ),
        ExternalFixtureSpec(
            fixture_id="excelize_chart_points_formula_cf",
            tool="excelize",
            filename="excelize-chart-points-formula-cf.xlsx",
            payload={
                "sheets": [{"name": "Metrics"}],
                "cells": [
                    {"sheet": "Metrics", "cell": "A1", "value": "Month"},
                    {"sheet": "Metrics", "cell": "B1", "value": "Revenue"},
                    {"sheet": "Metrics", "cell": "C1", "value": "Margin"},
                    {"sheet": "Metrics", "cell": "A2", "value": "Jan"},
                    {"sheet": "Metrics", "cell": "B2", "value": 1200},
                    {"sheet": "Metrics", "cell": "C2", "value": 0.31},
                    {"sheet": "Metrics", "cell": "A3", "value": "Feb"},
                    {"sheet": "Metrics", "cell": "B3", "value": 900},
                    {"sheet": "Metrics", "cell": "C3", "value": 0.24},
                    {"sheet": "Metrics", "cell": "A4", "value": "Mar"},
                    {"sheet": "Metrics", "cell": "B4", "value": 1450},
                    {"sheet": "Metrics", "cell": "C4", "value": 0.36},
                    {
                        "sheet": "Metrics",
                        "cell": "B5",
                        "type": "formula",
                        "formula": "SUM(B2:B4)",
                    },
                ],
                "conditional_formats": [
                    {"sheet": "Metrics", "range": "B2:B4", "type": "data_bar"},
                    {
                        "sheet": "Metrics",
                        "range": "C2:C4",
                        "type": "cell",
                        "criteria": ">",
                        "value": "0.3",
                    },
                ],
                "charts": [
                    {
                        "sheet": "Metrics",
                        "cell": "E2",
                        "type": "col",
                        "title": "Revenue",
                        "series": [
                            {
                                "name": "Metrics!$B$1",
                                "categories": "Metrics!$A$2:$A$4",
                                "values": "Metrics!$B$2:$B$4",
                                "data_points": [
                                    {"index": 0, "fill_color": "4472C4"},
                                    {"index": 1, "fill_color": "ED7D31"},
                                    {"index": 2, "fill_color": "70AD47"},
                                ],
                            }
                        ],
                        "show_values": True,
                    }
                ],
            },
            expected_parts=(
                "xl/charts/chart1.xml",
                "xl/drawings/drawing1.xml",
                "xl/worksheets/sheet1.xml",
            ),
            notes="Per-point chart styling, formulas, and conditional-formatting rules.",
        ),
    ]


def closedxml_fixture_specs() -> list[ExternalFixtureSpec]:
    """Return the initial ClosedXML fixture pack."""
    return [
        ExternalFixtureSpec(
            fixture_id="closedxml_pivot_cf_table",
            tool="closedxml",
            filename="closedxml-pivot-cf-table.xlsx",
            payload={
                "sheets": [{"name": "Data"}, {"name": "Pivot"}],
                "cells": [
                    {"sheet": "Data", "cell": "A1", "value": "Region"},
                    {"sheet": "Data", "cell": "B1", "value": "Product"},
                    {"sheet": "Data", "cell": "C1", "value": "Sales"},
                    {"sheet": "Data", "cell": "A2", "value": "West"},
                    {"sheet": "Data", "cell": "B2", "value": "Widgets"},
                    {"sheet": "Data", "cell": "C2", "value": 120},
                    {"sheet": "Data", "cell": "A3", "value": "East"},
                    {"sheet": "Data", "cell": "B3", "value": "Services"},
                    {"sheet": "Data", "cell": "C3", "value": 95},
                    {"sheet": "Data", "cell": "A4", "value": "West"},
                    {"sheet": "Data", "cell": "B4", "value": "Services"},
                    {"sheet": "Data", "cell": "C4", "value": 140},
                ],
                "tables": [{"sheet": "Data", "range": "A1:C4", "name": "ClosedXmlSales"}],
                "conditional_formats": [
                    {"sheet": "Data", "range": "C2:C4", "type": "3_color_scale"},
                    {"sheet": "Data", "range": "C2:C4", "type": "data_bar"},
                ],
                "pivots": [
                    {
                        "data_range": "Data!A1:C4",
                        "cell": "Pivot!A3",
                        "name": "ClosedXmlPivot",
                        "rows": [{"name": "Region"}],
                        "columns": [{"name": "Product"}],
                        "data": [{"name": "Sales"}],
                    }
                ],
            },
            expected_parts=(
                "xl/tables/table1.xml",
                "xl/pivotTables/pivotTable.xml",
                "pivotCache/pivotCacheDefinition1.xml",
                "xl/worksheets/sheet1.xml",
            ),
            notes="ClosedXML pivot/cache/table/conditional-formatting package layout.",
        ),
        ExternalFixtureSpec(
            fixture_id="closedxml_rich_comment_protection",
            tool="closedxml",
            filename="closedxml-rich-comment-protection.xlsx",
            payload={
                "sheets": [{"name": "Review"}],
                "cells": [
                    {"sheet": "Review", "cell": "A1", "value": "Finding"},
                    {"sheet": "Review", "cell": "B1", "value": "Status"},
                    {"sheet": "Review", "cell": "A2", "value": "Revenue cutoff"},
                    {"sheet": "Review", "cell": "B2", "value": "Open"},
                ],
                "rich_text": [
                    {
                        "sheet": "Review",
                        "cell": "A4",
                        "runs": [
                            {"text": "Priority: ", "bold": True, "font_color": "#C00000"},
                            {"text": "management response needed", "italic": True},
                        ],
                    }
                ],
                "comments": [
                    {
                        "sheet": "Review",
                        "cell": "B2",
                        "text": "Tie this status to final PBC evidence.",
                        "author": "ClosedXML Oracle",
                    }
                ],
                "protection": [{"sheet": "Review", "password": "audit"}],
            },
            expected_parts=(
                "xl/sharedStrings.xml",
                "xl/comments1.xml",
                "xl/drawings/vmldrawing.vml",
                "xl/worksheets/sheet1.xml",
            ),
            notes="ClosedXML rich text, legacy comments, and sheet protection.",
        ),
    ]


def npoi_fixture_specs() -> list[ExternalFixtureSpec]:
    """Return the initial NPOI fixture pack."""
    return [
        ExternalFixtureSpec(
            fixture_id="npoi_formula_comment_merge_protection",
            tool="npoi",
            filename="npoi-formula-comment-merge-protection.xlsx",
            payload={
                "sheets": [{"name": "NPOI"}],
                "cells": [
                    {"sheet": "NPOI", "cell": "A1", "value": "Account"},
                    {"sheet": "NPOI", "cell": "B1", "value": "Amount"},
                    {"sheet": "NPOI", "cell": "A2", "value": "Revenue"},
                    {"sheet": "NPOI", "cell": "B2", "value": 1250},
                    {"sheet": "NPOI", "cell": "A3", "value": "COGS"},
                    {"sheet": "NPOI", "cell": "B3", "value": -400},
                    {"sheet": "NPOI", "cell": "A4", "value": "Gross profit"},
                    {"sheet": "NPOI", "cell": "B4", "type": "formula", "formula": "SUM(B2:B3)"},
                    {"sheet": "NPOI", "cell": "D1", "value": "Merged review header"},
                ],
                "rich_text": [
                    {
                        "sheet": "NPOI",
                        "cell": "D3",
                        "runs": [
                            {"text": "NPOI ", "bold": True},
                            {"text": "rich text", "italic": True},
                        ],
                    }
                ],
                "comments": [
                    {
                        "sheet": "NPOI",
                        "cell": "B4",
                        "text": "Formula result should preserve calc metadata.",
                        "author": "NPOI Oracle",
                    }
                ],
                "merged_ranges": [{"sheet": "NPOI", "range": "D1:F1"}],
                "protection": [{"sheet": "NPOI", "password": "audit"}],
            },
            expected_parts=(
                "xl/sharedStrings.xml",
                "xl/comments1.xml",
                "xl/drawings/vmlDrawing1.vml",
                "xl/worksheets/sheet1.xml",
            ),
            notes="NPOI formula, rich text, legacy comments, merged range, and protection.",
        )
    ]


def external_fixture_specs() -> list[ExternalFixtureSpec]:
    """Return all implemented external fixture specifications."""
    return [*excelize_fixture_specs(), *closedxml_fixture_specs(), *npoi_fixture_specs()]


def generate_external_fixture_pack(
    output_root: Path,
    *,
    repo_root: Path,
    include_validators: bool = True,
    timeout_seconds: float = 180.0,
) -> list[FixtureGenerationResult]:
    """Generate the local external fixture pack and write ``manifest.json``."""
    output_root = output_root.resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    catalog = external_oracle_catalog(repo_root=repo_root)
    results: list[FixtureGenerationResult] = []

    for spec in external_fixture_specs():
        if not catalog[spec.tool].is_available():
            continue
        workbook_path = output_root / spec.filename
        write_result = run_external_oracle(
            catalog[spec.tool],
            ExternalOracleRequest(
                fixture_id=spec.fixture_id,
                operation="write_fixture",
                output_path=workbook_path,
                payload=spec.payload,
            ),
            timeout_seconds=timeout_seconds,
        )
        missing_parts = (
            _missing_parts(workbook_path, spec.expected_parts) if write_result.passed else ()
        )
        validations: list[ExternalOracleResult] = []
        if include_validators and write_result.passed:
            validations.extend(
                [
                    run_external_oracle(
                        catalog["libreoffice"],
                        ExternalOracleRequest(
                            fixture_id=spec.fixture_id,
                            operation="open_save_validate",
                            input_path=workbook_path,
                            output_path=output_root / "validated" / spec.filename,
                            payload={},
                        ),
                        timeout_seconds=timeout_seconds,
                    ),
                    run_external_oracle(
                        catalog["libreoffice"],
                        ExternalOracleRequest(
                            fixture_id=spec.fixture_id,
                            operation="render_validate",
                            input_path=workbook_path,
                            output_path=output_root / "pdf" / f"{workbook_path.stem}.pdf",
                            payload={},
                        ),
                        timeout_seconds=timeout_seconds,
                    ),
                ]
            )
        results.append(
            FixtureGenerationResult(
                fixture_id=spec.fixture_id,
                tool=spec.tool,
                workbook_path=workbook_path,
                write_result=write_result,
                expected_parts=spec.expected_parts,
                missing_parts=missing_parts,
                validations=tuple(validations),
            )
        )

    manifest = {
        "generated_at": datetime.now(UTC).isoformat(),
        "output_root": str(output_root),
        "fixtures": [result.to_json_dict(output_root) for result in results],
    }
    manifest_text = json.dumps(manifest, indent=2, sort_keys=True) + "\n"
    (output_root / "manifest.json").write_text(manifest_text)
    return results


def _sales_cells() -> list[JSONDict]:
    headers = ["Region", "Product", "Sales"]
    rows: list[tuple[str, str, int]] = [
        ("West", "Widgets", 120),
        ("East", "Widgets", 95),
        ("West", "Services", 140),
        ("East", "Services", 160),
        ("Central", "Widgets", 115),
    ]
    cells: list[JSONDict] = [
        {"sheet": "Data", "cell": f"{column}1", "value": value}
        for column, value in zip(("A", "B", "C"), headers, strict=True)
    ]
    for row_index, row in enumerate(rows, start=2):
        cells.extend(
            [
                {"sheet": "Data", "cell": f"A{row_index}", "value": row[0]},
                {"sheet": "Data", "cell": f"B{row_index}", "value": row[1]},
                {"sheet": "Data", "cell": f"C{row_index}", "value": row[2]},
            ]
        )
    return cells


def _missing_parts(workbook_path: Path, expected_parts: tuple[str, ...]) -> tuple[str, ...]:
    if not workbook_path.exists():
        return expected_parts
    with ZipFile(workbook_path) as workbook_zip:
        names = set(workbook_zip.namelist())
    return tuple(part for part in expected_parts if part not in names)


def _oracle_result_to_json(result: ExternalOracleResult) -> JSONDict:
    return {
        "tool_name": result.tool_name,
        "passed": result.passed,
        "skipped": result.skipped,
        "returncode": result.returncode,
        "payload": result.payload,
        "notes": result.notes,
    }
