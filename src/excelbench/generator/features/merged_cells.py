"""Generator for merged cells test cases (Tier 2)."""

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import Importance, MergeSpec, TestCase


class MergedCellsGenerator(FeatureGenerator):
    """Generates test cases for merged cells."""

    feature_name = "merged_cells"
    tier = 2
    filename = "10_merged_cells.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        test_cases: list[TestCase] = []
        row = 2

        # Simple horizontal merge
        label = "Merge horizontal B2:D2"
        merge_range = "B2:D2"
        sheet.range("B2").value = "Merged"
        sheet.range(merge_range).merge()
        expected = MergeSpec(range=merge_range, top_left_value="Merged").to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="merge_horizontal", label=label, row=row, expected=expected))
        row += 1

        # Vertical merge
        label = "Merge vertical B3:B5"
        merge_range = "B3:B5"
        sheet.range("B3").value = "Vertical"
        sheet.range(merge_range).merge()
        expected = MergeSpec(range=merge_range, top_left_value="Vertical").to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="merge_vertical", label=label, row=row, expected=expected))
        row += 1

        # Value not in top-left before merge
        label = "Merge with non-top-left value"
        merge_range = "B6:D6"
        sheet.range("C6").value = "OffTop"
        sheet.range(merge_range).merge()
        expected = MergeSpec(
            range=merge_range,
            top_left_value="OffTop",
            non_top_left_nonempty=0,
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(
            id="merge_value_off_top_left",
            label=label,
            row=row,
            expected=expected,
            importance=Importance.EDGE,
        ))
        row += 1

        # Formatting inheritance on top-left
        label = "Merge with top-left fill"
        merge_range = "B7:D7"
        sheet.range("B7").color = (255, 0, 0)
        sheet.range("B7").value = "Fill"
        sheet.range(merge_range).merge()
        expected = MergeSpec(
            range=merge_range,
            top_left_value="Fill",
            top_left_bg_color="#FF0000",
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(
            id="merge_top_left_fill",
            label=label,
            row=row,
            expected=expected,
            importance=Importance.EDGE,
        ))

        return test_cases
