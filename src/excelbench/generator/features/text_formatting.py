"""Generator for text formatting test cases."""

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import TestCase


class TextFormattingGenerator(FeatureGenerator):
    """Generates test cases for text formatting.

    Tests: bold, italic, underline, strikethrough, fonts, colors.
    """

    feature_name = "text_formatting"
    tier = 1
    filename = "03_text_formatting.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        """Generate text formatting test cases."""
        self.setup_header(sheet)

        test_cases = []
        row = 2

        # Basic formatting
        test_cases.append(self._test_bold(sheet, row))
        row += 1

        test_cases.append(self._test_italic(sheet, row))
        row += 1

        test_cases.append(self._test_underline_single(sheet, row))
        row += 1

        test_cases.append(self._test_underline_double(sheet, row))
        row += 1

        test_cases.append(self._test_strikethrough(sheet, row))
        row += 1

        test_cases.append(self._test_bold_italic(sheet, row))
        row += 1

        # Font sizes
        test_cases.append(self._test_font_size_8(sheet, row))
        row += 1

        test_cases.append(self._test_font_size_14(sheet, row))
        row += 1

        test_cases.append(self._test_font_size_24(sheet, row))
        row += 1

        test_cases.append(self._test_font_size_36(sheet, row))
        row += 1

        # Font families
        test_cases.append(self._test_font_arial(sheet, row))
        row += 1

        test_cases.append(self._test_font_times(sheet, row))
        row += 1

        test_cases.append(self._test_font_courier(sheet, row))
        row += 1

        # Font colors
        test_cases.append(self._test_color_red(sheet, row))
        row += 1

        test_cases.append(self._test_color_blue(sheet, row))
        row += 1

        test_cases.append(self._test_color_green(sheet, row))
        row += 1

        test_cases.append(self._test_color_custom(sheet, row))
        row += 1

        # Combined formatting
        test_cases.append(self._test_combined(sheet, row))
        row += 1

        return test_cases

    def _test_bold(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Bold"
        expected = {"bold": True}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Bold Text"
        cell.font.bold = True

        return TestCase(id="bold", label=label, row=row, expected=expected)

    def _test_italic(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Italic"
        expected = {"italic": True}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Italic Text"
        cell.font.italic = True

        return TestCase(id="italic", label=label, row=row, expected=expected)

    def _test_underline_single(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Underline - single"
        expected = {"underline": "single"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Underlined Text"
        # xlwings uses 2 for single underline
        self._set_underline(cell, 2)  # xlUnderlineStyleSingle

        return TestCase(id="underline_single", label=label, row=row, expected=expected)

    def _test_underline_double(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Underline - double"
        expected = {"underline": "double"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Double Underlined"
        # xlwings uses -4119 for double underline
        self._set_underline(cell, -4119)  # xlUnderlineStyleDouble

        return TestCase(id="underline_double", label=label, row=row, expected=expected)

    def _test_strikethrough(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Strikethrough"
        expected = {"strikethrough": True}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Strikethrough Text"
        self._set_strikethrough(cell, True)

        return TestCase(id="strikethrough", label=label, row=row, expected=expected)

    def _test_bold_italic(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Bold + Italic"
        expected = {"bold": True, "italic": True}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Bold Italic Text"
        cell.font.bold = True
        cell.font.italic = True

        return TestCase(id="bold_italic", label=label, row=row, expected=expected)

    def _test_font_size_8(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font size 8"
        expected = {"font_size": 8}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Size 8"
        cell.font.size = 8

        return TestCase(id="font_size_8", label=label, row=row, expected=expected)

    def _test_font_size_14(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font size 14"
        expected = {"font_size": 14}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Size 14"
        cell.font.size = 14

        return TestCase(id="font_size_14", label=label, row=row, expected=expected)

    def _test_font_size_24(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font size 24"
        expected = {"font_size": 24}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Size 24"
        cell.font.size = 24

        return TestCase(id="font_size_24", label=label, row=row, expected=expected)

    def _test_font_size_36(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font size 36"
        expected = {"font_size": 36}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Size 36"
        cell.font.size = 36

        return TestCase(id="font_size_36", label=label, row=row, expected=expected)

    def _test_font_arial(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font - Arial"
        expected = {"font_name": "Arial"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Arial Font"
        cell.font.name = "Arial"

        return TestCase(id="font_arial", label=label, row=row, expected=expected)

    def _test_font_times(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font - Times New Roman"
        expected = {"font_name": "Times New Roman"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Times New Roman"
        cell.font.name = "Times New Roman"

        return TestCase(id="font_times", label=label, row=row, expected=expected)

    def _test_font_courier(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font - Courier New"
        expected = {"font_name": "Courier New"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Courier New"
        cell.font.name = "Courier New"

        return TestCase(id="font_courier", label=label, row=row, expected=expected)

    def _test_color_red(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font color - red"
        expected = {"font_color": "#FF0000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Red Text"
        cell.font.color = (255, 0, 0)

        return TestCase(id="color_red", label=label, row=row, expected=expected)

    def _test_color_blue(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font color - blue"
        expected = {"font_color": "#0000FF"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Blue Text"
        cell.font.color = (0, 0, 255)

        return TestCase(id="color_blue", label=label, row=row, expected=expected)

    def _test_color_green(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font color - green"
        expected = {"font_color": "#00FF00"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Green Text"
        cell.font.color = (0, 255, 0)

        return TestCase(id="color_green", label=label, row=row, expected=expected)

    def _test_color_custom(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Font color - custom (#8B4513)"
        expected = {"font_color": "#8B4513"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Custom Color"
        cell.font.color = (139, 69, 19)  # Saddle brown

        return TestCase(id="color_custom", label=label, row=row, expected=expected)

    def _test_combined(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Combined - bold, 16pt, red"
        expected = {"bold": True, "font_size": 16, "font_color": "#FF0000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Combined Formatting"
        cell.font.bold = True
        cell.font.size = 16
        cell.font.color = (255, 0, 0)

        return TestCase(id="combined", label=label, row=row, expected=expected)

    @staticmethod
    def _set_underline(cell: xw.Range, value: int | bool) -> None:
        try:
            cell.api.Font.Underline = value
        except Exception:
            cell.api.font_object.underline.set(value)

    @staticmethod
    def _set_strikethrough(cell: xw.Range, value: bool) -> None:
        try:
            cell.api.Font.Strikethrough = value
        except Exception:
            cell.api.font_object.strikethrough.set(value)
