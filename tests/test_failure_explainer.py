from __future__ import annotations

from excelbench.models import (
    Diagnostic,
    DiagnosticCategory,
    DiagnosticLocation,
    DiagnosticSeverity,
    OperationType,
    TestResult,
)
from excelbench.results.failure_explainer import explain_diagnostic, explain_test_failure


def _diagnostic(message: str) -> Diagnostic:
    return Diagnostic(
        category=DiagnosticCategory.DATA_MISMATCH,
        severity=DiagnosticSeverity.ERROR,
        location=DiagnosticLocation(feature="styles", operation=OperationType.READ),
        adapter_message=message,
    )


def test_explain_diagnostic_classifies_style_drift() -> None:
    explanation = explain_diagnostic(
        _diagnostic("Expected fill color #FFFF00, actual default black"),
        {"bg_color": "#FFFF00"},
        {"bg_color": "#000000"},
    )

    assert explanation is not None
    assert explanation.code == "style_drift"
    assert explanation.tag == "style"


def test_explain_test_failure_classifies_formula_payload() -> None:
    result = TestResult(
        test_case_id="formula",
        operation=OperationType.READ,
        passed=False,
        expected={"formula": "=SUM(A1:A2)"},
        actual={"value": None},
        diagnostics=[],
    )

    explanation = explain_test_failure(result)

    assert explanation is not None
    assert explanation.code == "formula_cache_or_formula_drift"
    assert explanation.tag == "formula"
