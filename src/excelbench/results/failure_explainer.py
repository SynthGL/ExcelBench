"""Human-readable explanations for benchmark diagnostics."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from excelbench.models import Diagnostic, DiagnosticCategory, TestResult

JSONDict = dict[str, Any]


@dataclass(frozen=True)
class FailureExplanation:
    """Readable failure explanation attached to diagnostics and reports."""

    code: str
    summary: str
    probable_cause: str
    next_step: str
    tag: str

    def to_json_dict(self) -> JSONDict:
        return {
            "code": self.code,
            "summary": self.summary,
            "probable_cause": self.probable_cause,
            "next_step": self.next_step,
            "tag": self.tag,
        }


def explain_test_failure(result: TestResult) -> FailureExplanation | None:
    """Classify a failed test result into a concise explanation."""
    if result.passed:
        return None
    for diagnostic in result.diagnostics:
        explanation = explain_diagnostic(diagnostic, result.expected, result.actual)
        if explanation is not None:
            return explanation
    return _classify_payload(result.expected, result.actual)


def explain_diagnostic(
    diagnostic: Diagnostic,
    expected: JSONDict | None = None,
    actual: JSONDict | None = None,
) -> FailureExplanation | None:
    """Classify a diagnostic plus optional expected/actual payloads."""
    if diagnostic.category == DiagnosticCategory.UNSUPPORTED_FEATURE:
        return FailureExplanation(
            code="unsupported_feature",
            summary="adapter reported this operation as unsupported",
            probable_cause=diagnostic.probable_cause
            or "library/adapter does not implement the requested feature surface",
            next_step="treat as unsupported capability, not a semantic regression",
            tag="unsupported",
        )

    if diagnostic.root_cause_code and diagnostic.suggested_next_step:
        return FailureExplanation(
            code=diagnostic.root_cause_code,
            summary=diagnostic.probable_cause or diagnostic.adapter_message,
            probable_cause=diagnostic.probable_cause or diagnostic.adapter_message,
            next_step=diagnostic.suggested_next_step,
            tag=diagnostic.root_cause_code or "uncategorized",
        )
    message = f"{diagnostic.adapter_message} {diagnostic.probable_cause or ''}".lower()
    payload_text = f"{expected or {}} {actual or {}}".lower()
    return _classify_text(message, payload_text) or _classify_payload(expected or {}, actual or {})


def enrich_diagnostic(
    diagnostic: Diagnostic,
    *,
    expected: JSONDict | None = None,
    actual: JSONDict | None = None,
) -> Diagnostic:
    """Populate root cause fields when a known failure archetype matches."""
    explanation = explain_diagnostic(diagnostic, expected, actual)
    if explanation is None:
        return diagnostic
    if diagnostic.probable_cause is None:
        diagnostic.probable_cause = explanation.probable_cause
    diagnostic.root_cause_code = diagnostic.root_cause_code or explanation.code
    diagnostic.suggested_next_step = diagnostic.suggested_next_step or explanation.next_step
    return diagnostic


def render_why_failed(results: list[tuple[str, str, TestResult]]) -> str:
    """Render a compact WHY_FAILED.md body."""
    lines = ["# Why Failed", ""]
    failures = [(feature, library, tr) for feature, library, tr in results if not tr.passed]
    if not failures:
        lines.extend(["No failed test cases recorded.", ""])
        return "\n".join(lines)
    for feature, library, tr in failures:
        explanation = explain_test_failure(tr)
        lines.append(f"## {feature} / {library} / {tr.operation.value} / {tr.test_case_id}")
        lines.append("")
        if explanation is None:
            lines.append("- What failed: benchmark expectation did not match actual output.")
            lines.append("- Likely cause: no known failure archetype matched this payload.")
            lines.append("- Next debug surface: inspect expected/actual payloads in results.json.")
        else:
            lines.append(f"- What failed: {explanation.summary}")
            lines.append(f"- Likely cause: {explanation.probable_cause}")
            lines.append(f"- Next debug surface: {explanation.next_step}")
        lines.append("")
    return "\n".join(lines)


def _classify_payload(expected: JSONDict, actual: JSONDict) -> FailureExplanation | None:
    expected_keys = set(expected)
    actual_keys = set(actual)
    text = f"{expected} {actual}".lower()
    if "error" in actual:
        return FailureExplanation(
            "exception",
            "adapter raised instead of returning comparable workbook data",
            "adapter/runtime error interrupted the assertion path",
            "open the diagnostic adapter message and reproduce the adapter call directly",
            "exception",
        )
    if "formula" in expected_keys or "formula" in actual_keys or "cached" in text:
        return FailureExplanation(
            "formula_cache_or_formula_drift",
            "formula text or cached formula result drifted",
            "formula preservation and cached-value handling differ between libraries",
            "inspect formula XML and cached value handling for the target cell",
            "formula",
        )
    if {"bg_color", "font_color", "number_format", "format"} & expected_keys or "style" in text:
        return FailureExplanation(
            "style_drift",
            "cell style metadata changed",
            "style normalization, default style handling, or writer formatting support "
            "is incomplete",
            "diff styles.xml and the adapter's cell-format read/write path",
            "style",
        )
    return _classify_text("", text)


def _classify_text(message: str, payload_text: str) -> FailureExplanation | None:
    text = f"{message} {payload_text}"
    rules = [
        (
            ("not implemented", "unsupported", "not supported", "read-only", "write-only"),
            FailureExplanation(
                "unsupported_feature",
                "adapter does not support this feature surface",
                "the library or adapter has no implementation for this operation",
                "check adapter capability gating before treating this as semantic drift",
                "unsupported",
            ),
        ),
        (
            ("fill", "color", "black", "font", "number_format", "style"),
            FailureExplanation(
                "style_drift",
                "cell style metadata changed",
                "formatting was dropped, defaulted, or normalized differently",
                "inspect styles.xml and the adapter's style mapping",
                "style",
            ),
        ),
        (
            ("table", "totals", "autofilter"),
            FailureExplanation(
                "table_metadata_drift",
                "table metadata changed",
                "table XML, totals-row state, or auto-filter metadata was not preserved",
                "inspect xl/tables/table*.xml and worksheet table relationships",
                "table",
            ),
        ),
        (
            ("image", "drawing", "anchor", "media"),
            FailureExplanation(
                "drawing_or_image_drift",
                "image or drawing relationship changed",
                "media part may exist but worksheet drawing rels or anchors do not match",
                "inspect drawing XML, drawing rels, and xl/media package parts",
                "drawing",
            ),
        ),
        (
            ("named", "definedname", "scope"),
            FailureExplanation(
                "named_range_scope_drift",
                "named range metadata changed",
                "workbook-level vs sheet-level defined-name scope was lost or rewritten",
                "inspect workbook.xml definedNames and localSheetId values",
                "named_range",
            ),
        ),
        (
            ("merge", "merged"),
            FailureExplanation(
                "merge_drift",
                "merged-cell metadata changed",
                "merged range XML or non-anchor cell handling differs",
                "inspect mergeCells in the worksheet XML and subordinate cell behavior",
                "merge",
            ),
        ),
        (
            ("validation", "data_validation", "sqref"),
            FailureExplanation(
                "data_validation_drift",
                "data validation metadata changed",
                "validation type, formula, or target range was dropped or rewritten",
                "inspect dataValidations in the worksheet XML",
                "validation",
            ),
        ),
        (
            ("hyperlink", "relationship_target"),
            FailureExplanation(
                "hyperlink_drift",
                "hyperlink target or display metadata changed",
                "hyperlink rel target, tooltip, or internal location was not preserved",
                "inspect worksheet hyperlinks and worksheet rels",
                "hyperlink",
            ),
        ),
        (
            ("comment", "vml"),
            FailureExplanation(
                "comment_drift",
                "comment metadata changed",
                "legacy comment text, author, or VML relationship was not preserved",
                "inspect comments XML plus VML drawing relationships",
                "comment",
            ),
        ),
        (
            ("freeze", "pane"),
            FailureExplanation(
                "freeze_pane_drift",
                "freeze pane settings changed",
                "pane split/top-left metadata was dropped or normalized incorrectly",
                "inspect sheetViews/pane in the worksheet XML",
                "freeze_pane",
            ),
        ),
    ]
    for needles, explanation in rules:
        if any(needle in text for needle in needles):
            return explanation
    return None
