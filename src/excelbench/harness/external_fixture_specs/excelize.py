"""Excelize external fixture specifications."""

from __future__ import annotations

from excelbench.harness.external_fixture_specs.base import ExternalFixtureSpec, JSONDict


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
            readback_probes=(
                {"kind": "cell_value", "sheet": "Data", "cell": "A1", "expected": "Region"},
                {"kind": "cell_value", "sheet": "Data", "cell": "C6", "expected": 115},
                {
                    "kind": "zip_contains",
                    "part": "xl/worksheets/sheet1.xml",
                    "contains": "conditionalFormatting",
                    "label": "conditional formatting survives",
                },
                {
                    "kind": "conditional_formatting",
                    "sheet": "Data",
                    "sqref": "C2:C6",
                    "type": "dataBar",
                    "priority": 2,
                },
                {
                    "kind": "zip_contains",
                    "part": "xl/worksheets/sheet1.xml",
                    "contains": "slicerList",
                    "label": "slicer extension survives",
                },
                {
                    "kind": "table_metadata",
                    "sheet": "Data",
                    "name": "SalesTable",
                    "ref": "A1:C6",
                    "style": "TableStyleMedium9",
                },
                {
                    "kind": "relationship_target",
                    "part": "xl/drawings/_rels/drawing1.xml.rels",
                    "target": "../media/image1.png",
                    "type_contains": "/image",
                    "label": "image drawing relationship survives",
                },
            ),
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
            readback_probes=(
                {"kind": "cell_value", "sheet": "Metrics", "cell": "A1", "expected": "Month"},
                {
                    "kind": "cell_formula",
                    "sheet": "Metrics",
                    "cell": "B5",
                    "expected": "=SUM(B2:B4)",
                },
                {
                    "kind": "zip_contains",
                    "part": "xl/worksheets/sheet1.xml",
                    "contains": "conditionalFormatting",
                    "label": "conditional formatting survives",
                },
                {
                    "kind": "conditional_formatting",
                    "sheet": "Metrics",
                    "sqref": "C2:C4",
                    "type": "cellIs",
                    "priority": 2,
                    "operator": "greaterThan",
                    "formula": "0.3",
                },
            ),
        ),
    ]


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
