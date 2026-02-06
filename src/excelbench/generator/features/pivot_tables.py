"""Generator for pivot table test cases (Tier 2)."""

import shutil
import sys
from pathlib import Path

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import Importance, PivotSpec, TestCase


class PivotTablesGenerator(FeatureGenerator):
    """Generates test cases for pivot tables."""

    feature_name = "pivot_tables"
    tier = 2
    filename = "15_pivot_tables.xlsx"

    def __init__(self) -> None:
        self._fixture_path = Path("fixtures/excel/tier2/15_pivot_tables.xlsx")

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        if sys.platform == "darwin":
            if self._fixture_path.exists():
                return self._fixture_test_cases(sheet)
            print("  Pivot fixture not found; skipping pivot tests on macOS.")
            return []

        wb = sheet.book
        data_sheet = wb.sheets.add("Data")
        pivot_sheet = wb.sheets.add("Pivot")

        # Seed data
        data = [
            ["Region", "Product", "Date", "Sales"],
            ["North", "A", "2026-01-05", 100],
            ["North", "B", "2026-01-08", 150],
            ["South", "A", "2026-02-03", 90],
            ["South", "B", "2026-02-10", 120],
            ["West", "A", "2026-03-15", 200],
        ]
        data_sheet.range("A1").value = data

        source_range = data_sheet.range("A1:D6").api
        dest = pivot_sheet.range("B3").api

        # Create pivot cache and table
        pivot_cache = wb.api.PivotCaches().Create(SourceType=1, SourceData=source_range)
        pivot_table = pivot_cache.CreatePivotTable(TableDestination=dest, TableName="SalesPivot")

        # Field layout
        pivot_table.PivotFields("Region").Orientation = 1  # xlRowField
        pivot_table.PivotFields("Product").Orientation = 2  # xlColumnField
        pivot_table.PivotFields("Date").Orientation = 3  # xlPageField

        # Data field (sum)
        pivot_table.AddDataField(pivot_table.PivotFields("Sales"), "Sum of Sales", -4157)
        # Add count field
        pivot_table.AddDataField(pivot_table.PivotFields("Sales"), "Count of Sales", -4112)

        # Calculated field
        try:
            pivot_table.CalculatedFields().Add("SalesPlus10", "=Sales+10")
            pivot_table.AddDataField(pivot_table.PivotFields("SalesPlus10"), "SalesPlus10", -4157)
            calc_added = True
        except Exception:
            calc_added = False

        # Attempt date grouping (month)
        grouped = False
        try:
            pivot_table.PivotFields("Date").Group(True, True, 30)
            grouped = True
        except Exception:
            grouped = False

        test_cases: list[TestCase] = []
        row = 2

        # Basic pivot structure
        expected = PivotSpec(
            name="SalesPivot",
            source_range="Data!A1:D6",
            target_cell="Pivot!B3",
            row_fields=["Region"],
            column_fields=["Product"],
            data_fields=["Sum of Sales"],
            filter_fields=["Date"],
        ).to_expected()
        self.write_test_case(sheet, row, "Pivot: basic layout", expected)
        test_cases.append(
            TestCase(
                id="pivot_basic",
                label="Pivot: basic layout",
                row=row,
                expected=expected,
                sheet="Pivot",
            )
        )
        row += 1

        # Multiple value fields
        expected = PivotSpec(
            name="SalesPivot",
            source_range="Data!A1:D6",
            target_cell="Pivot!B3",
            row_fields=["Region"],
            column_fields=["Product"],
            data_fields=["Sum of Sales", "Count of Sales"],
            filter_fields=["Date"],
        ).to_expected()
        self.write_test_case(sheet, row, "Pivot: multiple value fields", expected)
        test_cases.append(
            TestCase(
                id="pivot_multi_values",
                label="Pivot: multiple value fields",
                row=row,
                expected=expected,
                sheet="Pivot",
                importance=Importance.EDGE,
            )
        )
        row += 1

        # Calculated field
        if calc_added:
            expected = PivotSpec(
                name="SalesPivot",
                source_range="Data!A1:D6",
                target_cell="Pivot!B3",
                row_fields=["Region"],
                column_fields=["Product"],
                data_fields=["SalesPlus10"],
                filter_fields=["Date"],
            ).to_expected()
            self.write_test_case(sheet, row, "Pivot: calculated field", expected)
            test_cases.append(
                TestCase(
                    id="pivot_calc_field",
                    label="Pivot: calculated field",
                    row=row,
                    expected=expected,
                    sheet="Pivot",
                    importance=Importance.EDGE,
                )
            )
            row += 1

        # Grouping
        expected = {
            "pivot": {
                "name": "SalesPivot",
                "grouped": grouped,
            }
        }
        self.write_test_case(sheet, row, "Pivot: date grouping", expected)
        test_cases.append(
            TestCase(
                id="pivot_grouping",
                label="Pivot: date grouping",
                row=row,
                expected=expected,
                sheet="Pivot",
                importance=Importance.EDGE,
            )
        )

        return test_cases

    def post_process(self, output_path: Path) -> None:
        if sys.platform != "darwin":
            return
        if self._fixture_path.exists():
            shutil.copyfile(self._fixture_path, output_path)

    def _fixture_test_cases(self, sheet: xw.Sheet) -> list[TestCase]:
        test_cases: list[TestCase] = []
        row = 2

        expected = PivotSpec(
            name="SalesPivot",
            source_range="Data!A1:D6",
            target_cell="Pivot!B3",
            row_fields=["Region"],
            column_fields=["Product"],
            data_fields=["Sum of Sales"],
            filter_fields=["Date"],
        ).to_expected()
        self.write_test_case(sheet, row, "Pivot: basic layout", expected)
        test_cases.append(
            TestCase(
                id="pivot_basic",
                label="Pivot: basic layout",
                row=row,
                expected=expected,
                sheet="Pivot",
            )
        )
        row += 1

        expected = PivotSpec(
            name="SalesPivot",
            source_range="Data!A1:D6",
            target_cell="Pivot!B3",
            row_fields=["Region"],
            column_fields=["Product"],
            data_fields=["Sum of Sales", "Count of Sales"],
            filter_fields=["Date"],
        ).to_expected()
        self.write_test_case(sheet, row, "Pivot: multiple value fields", expected)
        test_cases.append(
            TestCase(
                id="pivot_multi_values",
                label="Pivot: multiple value fields",
                row=row,
                expected=expected,
                sheet="Pivot",
                importance=Importance.EDGE,
            )
        )

        return test_cases
