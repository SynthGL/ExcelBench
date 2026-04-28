"""Apache POI external fixture specifications."""

from __future__ import annotations

from excelbench.harness.external_fixture_specs.base import ExternalFixtureSpec


def apache_poi_fixture_specs() -> list[ExternalFixtureSpec]:
    """Return the initial Apache POI fixture pack."""
    return [
        ExternalFixtureSpec(
            fixture_id="apache_poi_table_validation_image_comment",
            tool="apache-poi",
            filename="apache-poi-table-validation-image-comment.xlsx",
            payload={},
            expected_parts=(
                "xl/sharedStrings.xml",
                "xl/comments1.xml",
                "xl/drawings/vmlDrawing0.vml",
                "xl/drawings/drawing1.xml",
                "xl/media/image1.png",
                "xl/tables/table1.xml",
                "xl/worksheets/sheet1.xml",
            ),
            notes=(
                "Apache POI table, formula, data validation, rich text, comment, hyperlink, "
                "image, merge, freeze panes, and sheet protection."
            ),
            readback_probes=(
                {"kind": "cell_value", "sheet": "POI", "cell": "A1", "expected": "Metric"},
                {
                    "kind": "cell_formula",
                    "sheet": "POI",
                    "cell": "B4",
                    "expected": "=SUM(B2:B3)",
                },
                {
                    "kind": "cell_style",
                    "sheet": "POI",
                    "cell": "B2",
                    "expected": {"number_format": "$#,##0"},
                },
                {
                    "kind": "comment_text",
                    "sheet": "POI",
                    "cell": "B4",
                    "contains": "POI formula",
                },
                {
                    "kind": "hyperlink_target",
                    "sheet": "POI",
                    "cell": "D4",
                    "target": "https://poi.apache.org/",
                },
                {
                    "kind": "data_validation",
                    "sheet": "POI",
                    "cell": "C2",
                    "type": "list",
                    "formula1": '"Open,Closed,Review"',
                },
                {"kind": "merged_range", "sheet": "POI", "range": "D1:F1"},
                {
                    "kind": "table_metadata",
                    "sheet": "POI",
                    "name": "PoiReviewTable",
                    "ref": "F1:G4",
                    "style": "TableStyleMedium2",
                },
                {
                    "kind": "zip_contains",
                    "part": "xl/worksheets/sheet1.xml",
                    "contains": "dataValidations",
                    "label": "data validation survives",
                },
                {
                    "kind": "zip_contains",
                    "part": "xl/tables/table1.xml",
                    "contains": "PoiReviewTable",
                    "label": "table name survives",
                },
            ),
        )
    ]
