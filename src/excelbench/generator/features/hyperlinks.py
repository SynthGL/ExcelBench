"""Generator for hyperlink test cases (Tier 2)."""

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import HyperlinkSpec, Importance, TestCase


class HyperlinksGenerator(FeatureGenerator):
    """Generates test cases for hyperlinks."""

    feature_name = "hyperlinks"
    tier = 2
    filename = "13_hyperlinks.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        test_cases: list[TestCase] = []
        row = 2

        # External hyperlink
        label = "Hyperlink: external URL"
        cell = "B2"
        sheet.range(cell).add_hyperlink(
            "https://example.com/docs",
            "Example Docs",
            "Go to docs",
        )
        expected = HyperlinkSpec(
            cell=cell,
            target="https://example.com/docs",
            display="Example Docs",
            tooltip="Go to docs",
            internal=False,
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="link_external", label=label, row=row, expected=expected))
        row += 1

        # Internal hyperlink to another sheet
        label = "Hyperlink: internal sheet"
        target_sheet = sheet.book.sheets.add("Targets")
        target_sheet.range("A1").value = "Target"
        cell = "B3"
        sheet.range(cell).add_hyperlink(
            "#'Targets'!A1",
            "Go Target",
            "Jump to target",
        )
        expected = HyperlinkSpec(
            cell=cell,
            target="Targets!A1",
            display="Go Target",
            tooltip="Jump to target",
            internal=True,
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(
            id="link_internal",
            label=label,
            row=row,
            expected=expected,
            importance=Importance.EDGE,
        ))
        row += 1

        # Mailto hyperlink
        label = "Hyperlink: mailto"
        cell = "B4"
        sheet.range(cell).add_hyperlink(
            "mailto:test@example.com",
            "Email",
            "Send email",
        )
        expected = HyperlinkSpec(
            cell=cell,
            target="mailto:test@example.com",
            display="Email",
            tooltip="Send email",
            internal=False,
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="link_mailto", label=label, row=row, expected=expected))
        row += 1

        # Long URL with encoding
        label = "Hyperlink: long encoded URL"
        cell = "B5"
        url = "https://example.com/search?q=excel%20bench&sort=desc#section-2"
        sheet.range(cell).add_hyperlink(url, "Search", "Encoded URL")
        expected = HyperlinkSpec(
            cell=cell,
            target=url,
            display="Search",
            tooltip="Encoded URL",
            internal=False,
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(
            id="link_long",
            label=label,
            row=row,
            expected=expected,
            importance=Importance.EDGE,
        ))

        return test_cases
