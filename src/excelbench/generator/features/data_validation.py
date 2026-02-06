"""Generator for data validation test cases (Tier 2)."""

import sys
from pathlib import Path

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import DataValidationSpec, Importance, TestCase


class DataValidationGenerator(FeatureGenerator):
    """Generates test cases for data validation."""

    feature_name = "data_validation"
    tier = 2
    filename = "12_data_validation.xlsx"

    def __init__(self) -> None:
        self._use_openpyxl = sys.platform == "darwin"
        self._dv_ops: list[dict[str, object]] = []

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        test_cases: list[TestCase] = []
        row = 2

        # 1) List from CSV string
        label = "DV: list from CSV"
        cell = "B2"
        if not self._use_openpyxl:
            sheet.range(cell).api.Validation.Delete()
            sheet.range(cell).api.Validation.Add(
                Type=3,  # xlValidateList
                AlertStyle=1,
                Operator=1,
                Formula1='"Red,Green,Blue"',
            )
        self._record_dv(
            range=cell,
            validation_type="list",
            formula1='"Red,Green,Blue"',
            allow_blank=True,
        )
        expected = DataValidationSpec(
            range=cell,
            validation_type="list",
            formula1='"Red,Green,Blue"',
            allow_blank=True,
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="dv_list_csv", label=label, row=row, expected=expected))
        row += 1

        # 2) List from range
        label = "DV: list from range"
        sheet.range("D2:D4").value = ["A", "B", "C"]
        cell = "B3"
        if not self._use_openpyxl:
            sheet.range(cell).api.Validation.Delete()
            sheet.range(cell).api.Validation.Add(
                Type=3,
                AlertStyle=1,
                Operator=1,
                Formula1="=$D$2:$D$4",
            )
        self._record_dv(
            range=cell,
            validation_type="list",
            formula1="=$D$2:$D$4",
        )
        expected = DataValidationSpec(
            range=cell,
            validation_type="list",
            formula1="=$D$2:$D$4",
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(
            TestCase(
                id="dv_list_range",
                label=label,
                row=row,
                expected=expected,
                importance=Importance.EDGE,
            )
        )
        row += 1

        # 3) Cross-sheet list via named range
        label = "DV: cross-sheet named range"
        list_sheet = sheet.book.sheets.add("Lists")
        list_sheet.range("A1:A3").value = ["North", "South", "West"]
        sheet.book.names.add("RegionList", "=Lists!$A$1:$A$3")
        cell = "B4"
        if not self._use_openpyxl:
            sheet.range(cell).api.Validation.Delete()
            sheet.range(cell).api.Validation.Add(
                Type=3,
                AlertStyle=1,
                Operator=1,
                Formula1="=RegionList",
            )
        self._record_dv(
            range=cell,
            validation_type="list",
            formula1="=RegionList",
        )
        expected = DataValidationSpec(
            range=cell,
            validation_type="list",
            formula1="=RegionList",
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(
            TestCase(
                id="dv_cross_sheet",
                label=label,
                row=row,
                expected=expected,
                importance=Importance.EDGE,
            )
        )
        row += 1

        # 4) Custom formula
        label = "DV: custom formula"
        sheet.range("C5").value = 5
        cell = "B5"
        if not self._use_openpyxl:
            sheet.range(cell).api.Validation.Delete()
            sheet.range(cell).api.Validation.Add(
                Type=7,  # xlValidateCustom
                AlertStyle=1,
                Operator=1,
                Formula1="=B5>C5",
            )
        self._record_dv(
            range=cell,
            validation_type="custom",
            formula1="=B5>C5",
        )
        expected = DataValidationSpec(
            range=cell,
            validation_type="custom",
            formula1="=B5>C5",
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(
            TestCase(
                id="dv_custom_formula",
                label=label,
                row=row,
                expected=expected,
                importance=Importance.EDGE,
            )
        )
        row += 1

        # 5) Whole number between with error message
        label = "DV: whole number with error"
        cell = "B6"
        if not self._use_openpyxl:
            sheet.range(cell).api.Validation.Delete()
            sheet.range(cell).api.Validation.Add(
                Type=1,  # xlValidateWholeNumber
                AlertStyle=2,  # xlValidAlertStop
                Operator=1,  # xlBetween
                Formula1="1",
                Formula2="10",
            )
            sheet.range(cell).api.Validation.IgnoreBlank = False
            sheet.range(cell).api.Validation.ErrorTitle = "Invalid"
            sheet.range(cell).api.Validation.ErrorMessage = "Enter 1-10"
        self._record_dv(
            range=cell,
            validation_type="whole",
            operator="between",
            formula1="1",
            formula2="10",
            allow_blank=False,
            error_title="Invalid",
            error="Enter 1-10",
        )
        expected = DataValidationSpec(
            range=cell,
            validation_type="whole",
            operator="between",
            formula1="1",
            formula2="10",
            allow_blank=False,
            error_title="Invalid",
            error="Enter 1-10",
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="dv_whole_between", label=label, row=row, expected=expected))

        return test_cases

    def _record_dv(self, **kwargs: object) -> None:
        self._dv_ops.append(kwargs)

    def post_process(self, output_path: Path) -> None:
        if not self._use_openpyxl or not self._dv_ops:
            return
        from openpyxl import load_workbook
        from openpyxl.worksheet.datavalidation import DataValidation

        wb = load_workbook(output_path)
        ws = wb[self.feature_name]

        for op in self._dv_ops:
            dv = DataValidation(
                type=str(op.get("validation_type")),
                operator=op.get("operator"),
                formula1=op.get("formula1"),
                formula2=op.get("formula2"),
                allow_blank=op.get("allow_blank", True),
            )
            error_title = op.get("error_title")
            error = op.get("error")
            if isinstance(error_title, str):
                dv.errorTitle = error_title
            if isinstance(error, str):
                dv.error = error
            target = op.get("range")
            if isinstance(target, str):
                dv.add(target)
                ws.add_data_validation(dv)

        wb.save(output_path)
