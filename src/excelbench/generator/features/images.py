"""Generator for images/embedded objects test cases (Tier 2)."""

import sys
from pathlib import Path

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import ImageSpec, Importance, TestCase


class ImagesGenerator(FeatureGenerator):
    """Generates test cases for embedded images."""

    feature_name = "images"
    tier = 2
    filename = "14_images.xlsx"

    def __init__(self) -> None:
        self._use_openpyxl = sys.platform == "darwin"
        self._ops: list[dict[str, object]] = []

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        test_cases: list[TestCase] = []
        row = 2

        image_dir = Path("fixtures/images")
        png_path = (image_dir / "sample.png").resolve()
        jpg_path = (image_dir / "sample.jpg").resolve()

        # One-cell anchor image
        label = "Image: one-cell anchor"
        cell = "B2"
        if not self._use_openpyxl:
            anchor = sheet.range(cell)
            sheet.pictures.add(
                png_path,
                name="png_one_cell",
                left=anchor.left,
                top=anchor.top,
                width=60,
                height=60,
            )
            expected = ImageSpec(
                cell=cell,
                path=str(png_path),
                anchor="oneCell",
            ).to_expected()
        else:
            self._ops.append({"cell": cell, "path": str(png_path)})
            expected = ImageSpec(
                cell=cell,
                path=str(png_path),
                anchor="oneCell",
            ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="image_one_cell", label=label, row=row, expected=expected))
        row += 1

        # Two-cell anchor image with offset
        label = "Image: two-cell anchor with offset"
        cell = "D6"
        if not self._use_openpyxl:
            anchor = sheet.range(cell)
            sheet.pictures.add(
                jpg_path,
                name="jpg_two_cell",
                left=anchor.left + 8,
                top=anchor.top + 6,
                width=120,
                height=80,
            )
            expected = ImageSpec(
                cell=cell,
                path=str(jpg_path),
                anchor="twoCell",
                offset=(8, 6),
            ).to_expected()
        else:
            self._ops.append({"cell": cell, "path": str(jpg_path)})
            expected = ImageSpec(
                cell=cell,
                path=str(jpg_path),
                anchor="oneCell",
            ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(
            TestCase(
                id="image_two_cell_offset",
                label=label,
                row=row,
                expected=expected,
                importance=Importance.EDGE,
            )
        )

        return test_cases

    def post_process(self, output_path: Path) -> None:
        if not self._use_openpyxl or not self._ops:
            return
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image

        wb = load_workbook(output_path)
        ws = wb[self.feature_name]

        for op in self._ops:
            cell = op.get("cell")
            path = op.get("path")
            if not isinstance(cell, str) or not isinstance(path, str):
                continue
            img = Image(path)
            ws.add_image(img, cell)

        wb.save(output_path)
