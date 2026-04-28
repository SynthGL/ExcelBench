"""NPOI external fixture specifications."""

from __future__ import annotations

from excelbench.harness.external_fixture_specs.base import ExternalFixtureSpec


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
            readback_probes=(
                {"kind": "cell_value", "sheet": "NPOI", "cell": "A1", "expected": "Account"},
                {
                    "kind": "cell_formula",
                    "sheet": "NPOI",
                    "cell": "B4",
                    "expected": "=SUM(B2:B3)",
                },
                {
                    "kind": "comment_text",
                    "sheet": "NPOI",
                    "cell": "B4",
                    "contains": "Formula result",
                },
                {"kind": "merged_range", "sheet": "NPOI", "range": "D1:F1"},
                {
                    "kind": "zip_contains",
                    "part": "xl/worksheets/sheet1.xml",
                    "contains": "sheetProtection",
                    "label": "sheet protection survives",
                },
                {
                    "kind": "sheet_protection",
                    "sheet": "NPOI",
                    "expected": {"sheet": True, "objects": True, "scenarios": True},
                },
                {
                    "kind": "rich_text_runs",
                    "part": "xl/sharedStrings.xml",
                    "min_runs": 2,
                    "contains": ["NPOI ", "rich text"],
                    "label": "rich text runs survive",
                },
            ),
        )
    ]
