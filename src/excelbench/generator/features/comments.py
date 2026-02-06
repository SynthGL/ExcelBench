"""Generator for comments/notes test cases (Tier 2)."""

import sys

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import CommentSpec, Importance, TestCase


class CommentsGenerator(FeatureGenerator):
    """Generates test cases for comments and notes."""

    feature_name = "comments"
    tier = 2
    filename = "16_comments.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        is_mac = sys.platform == "darwin"

        test_cases: list[TestCase] = []
        row = 2

        # Legacy note
        label = "Comment: legacy note"
        cell = "B2"
        sheet.range(cell).value = "Note"
        if is_mac:
            sheet.range(cell).api.add_comment(comment_text="Legacy note")
        else:
            sheet.range(cell).api.AddComment("Legacy note")
            sheet.range(cell).api.Comment.Author = "ExcelBench"
        expected = CommentSpec(
            cell=cell,
            text="Legacy note",
            author=None if is_mac else "ExcelBench",
            threaded=False,
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="comment_legacy", label=label, row=row, expected=expected))
        row += 1

        # Threaded comment (if supported)
        label = "Comment: threaded"
        cell = "B3"
        sheet.range(cell).value = "Thread"
        threaded_added = False
        if not is_mac:
            try:
                sheet.range(cell).api.CommentThreaded.Add("Threaded comment")
                threaded_added = True
            except Exception:
                threaded_added = False
        if not threaded_added:
            if is_mac:
                sheet.range(cell).api.add_comment(comment_text="Threaded fallback")
            else:
                sheet.range(cell).api.AddComment("Threaded fallback")
        expected = CommentSpec(
            cell=cell,
            text="Threaded comment" if threaded_added else "Threaded fallback",
            author=None,
            threaded=threaded_added,
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(
            TestCase(
                id="comment_threaded",
                label=label,
                row=row,
                expected=expected,
                importance=Importance.EDGE,
            )
        )
        row += 1

        # Multiple authors (legacy)
        label = "Comment: second author"
        cell = "B4"
        sheet.range(cell).value = "Note 2"
        if is_mac:
            sheet.range(cell).api.add_comment(comment_text="Another note")
        else:
            sheet.range(cell).api.AddComment("Another note")
            sheet.range(cell).api.Comment.Author = "Alice"
        expected = CommentSpec(
            cell=cell,
            text="Another note",
            author=None if is_mac else "Alice",
            threaded=False,
        ).to_expected()
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(
            TestCase(
                id="comment_author",
                label=label,
                row=row,
                expected=expected,
                importance=Importance.EDGE,
            )
        )

        return test_cases
