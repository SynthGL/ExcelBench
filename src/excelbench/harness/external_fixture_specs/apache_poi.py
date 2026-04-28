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
        )
    ]
