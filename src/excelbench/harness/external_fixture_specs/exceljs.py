"""ExcelJS external fixture specifications."""

from __future__ import annotations

from excelbench.harness.external_fixture_specs.base import ExternalFixtureSpec


def exceljs_fixture_specs() -> list[ExternalFixtureSpec]:
    """Return the initial ExcelJS fixture pack."""
    return [
        ExternalFixtureSpec(
            fixture_id="exceljs_table_validation_image_comment",
            tool="exceljs",
            filename="exceljs-table-validation-image-comment.xlsx",
            payload={
                "sheets": [
                    {
                        "name": "ExcelJS",
                        "freeze_panes": {"x_split": 1, "y_split": 1},
                    }
                ],
                "columns": [
                    {"sheet": "ExcelJS", "column": "A", "width": 18},
                    {"sheet": "ExcelJS", "column": "B", "width": 14},
                    {"sheet": "ExcelJS", "column": "D", "width": 24},
                ],
                "cells": [
                    {"sheet": "ExcelJS", "cell": "A1", "value": "Metric", "font": {"bold": True}},
                    {"sheet": "ExcelJS", "cell": "B1", "value": "Value", "font": {"bold": True}},
                    {"sheet": "ExcelJS", "cell": "A2", "value": "Revenue"},
                    {"sheet": "ExcelJS", "cell": "B2", "value": 1200, "num_fmt": "$#,##0"},
                    {"sheet": "ExcelJS", "cell": "A3", "value": "COGS"},
                    {"sheet": "ExcelJS", "cell": "B3", "value": -450, "num_fmt": "$#,##0"},
                    {"sheet": "ExcelJS", "cell": "A4", "value": "Gross profit"},
                    {
                        "sheet": "ExcelJS",
                        "cell": "B4",
                        "type": "formula",
                        "formula": "SUM(B2:B3)",
                        "result": 750,
                        "num_fmt": "$#,##0",
                    },
                    {"sheet": "ExcelJS", "cell": "D1", "value": "Merged review header"},
                ],
                "rich_text": [
                    {
                        "sheet": "ExcelJS",
                        "cell": "D2",
                        "runs": [
                            {"text": "ExcelJS ", "bold": True, "font_color": "#4472C4"},
                            {"text": "rich text", "italic": True},
                        ],
                    }
                ],
                "comments": [
                    {
                        "sheet": "ExcelJS",
                        "cell": "B4",
                        "text": "Formula result should survive package round trips.",
                    }
                ],
                "hyperlinks": [
                    {
                        "sheet": "ExcelJS",
                        "cell": "D4",
                        "text": "ExcelJS project",
                        "url": "https://github.com/exceljs/exceljs",
                        "tooltip": "ExcelJS project",
                    }
                ],
                "merged_ranges": [{"sheet": "ExcelJS", "range": "D1:F1"}],
                "data_validations": [
                    {
                        "sheet": "ExcelJS",
                        "cell": "C2",
                        "type": "list",
                        "formulae": ['"Open,Closed,Review"'],
                    }
                ],
                "tables": [
                    {
                        "sheet": "ExcelJS",
                        "name": "ExcelJsReviewTable",
                        "ref": "F1:G4",
                        "columns": [{"name": "Item"}, {"name": "Status"}],
                        "rows": [
                            ["Revenue", "Open"],
                            ["COGS", "Closed"],
                            ["Gross profit", "Review"],
                        ],
                    }
                ],
                "images": [{"sheet": "ExcelJS", "range": "D6:E8"}],
                "protection": [{"sheet": "ExcelJS", "password": "audit"}],
            },
            expected_parts=(
                "xl/sharedStrings.xml",
                "xl/comments1.xml",
                "xl/drawings/vmlDrawing1.vml",
                "xl/drawings/drawing1.xml",
                "xl/media/image1.png",
                "xl/tables/table1.xml",
                "xl/worksheets/sheet1.xml",
            ),
            notes=(
                "ExcelJS table, formula, data validation, rich text, comment, hyperlink, "
                "image, merge, freeze panes, and sheet protection."
            ),
            readback_probes=(
                {"kind": "cell_value", "sheet": "ExcelJS", "cell": "A1", "expected": "Metric"},
                {
                    "kind": "cell_formula",
                    "sheet": "ExcelJS",
                    "cell": "B4",
                    "expected": "=SUM(B2:B3)",
                },
                {
                    "kind": "comment_text",
                    "sheet": "ExcelJS",
                    "cell": "B4",
                    "contains": "Formula result",
                },
                {"kind": "merged_range", "sheet": "ExcelJS", "range": "D1:F1"},
                {
                    "kind": "zip_contains",
                    "part": "xl/worksheets/sheet1.xml",
                    "contains": "dataValidations",
                    "label": "data validation survives",
                },
                {
                    "kind": "zip_contains",
                    "part": "xl/tables/table1.xml",
                    "contains": "ExcelJsReviewTable",
                    "label": "table name survives",
                },
            ),
        )
    ]
