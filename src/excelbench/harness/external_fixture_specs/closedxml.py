"""ClosedXML external fixture specifications."""

from __future__ import annotations

from excelbench.harness.external_fixture_specs.base import ExternalFixtureSpec


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
            readback_probes=(
                {"kind": "cell_value", "sheet": "Data", "cell": "A1", "expected": "Region"},
                {"kind": "cell_value", "sheet": "Data", "cell": "C4", "expected": 140},
                {
                    "kind": "zip_contains",
                    "part": "xl/worksheets/sheet1.xml",
                    "contains": "conditionalFormatting",
                    "label": "conditional formatting survives",
                },
                {
                    "kind": "conditional_formatting",
                    "sheet": "Data",
                    "sqref": "C2:C4",
                    "type": "dataBar",
                    "priority": 2,
                },
                {
                    "kind": "zip_contains",
                    "part": "xl/tables/table1.xml",
                    "contains": "ClosedXmlSales",
                    "label": "table name survives",
                },
                {
                    "kind": "table_metadata",
                    "sheet": "Data",
                    "name": "ClosedXmlSales",
                    "ref": "A1:C4",
                    "style": "TableStyleMedium2",
                },
            ),
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
            readback_probes=(
                {"kind": "cell_value", "sheet": "Review", "cell": "A1", "expected": "Finding"},
                {
                    "kind": "comment_text",
                    "sheet": "Review",
                    "cell": "B2",
                    "contains": "Tie this status",
                },
                {
                    "kind": "zip_contains",
                    "part": "xl/worksheets/sheet1.xml",
                    "contains": "sheetProtection",
                    "label": "sheet protection survives",
                },
                {
                    "kind": "zip_contains",
                    "part": "xl/sharedStrings.xml",
                    "contains": ":r>",
                    "label": "rich text runs survive",
                },
            ),
        ),
    ]
