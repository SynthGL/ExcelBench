from typing import Any

from excelbench.harness.runner import compare_results

JSONDict = dict[str, Any]


def test_compare_results_color_case_insensitive() -> None:
    expected: JSONDict = {"font_color": "#ff0000"}
    actual: JSONDict = {"font_color": "#FF0000"}
    assert compare_results(expected, actual)


def test_compare_results_missing_expected_none() -> None:
    expected: JSONDict = {"border_top": None}
    actual: JSONDict = {}
    assert compare_results(expected, actual)


def test_compare_results_numeric_tolerance() -> None:
    expected: JSONDict = {"value": 1.0}
    actual: JSONDict = {"value": 1.00005}
    assert compare_results(expected, actual)

    actual_fail: JSONDict = {"value": 1.01}
    assert not compare_results(expected, actual_fail)


def test_compare_results_list_order_insensitive() -> None:
    expected: JSONDict = {"items": [{"cell": "B2"}, {"cell": "B3"}]}
    actual: JSONDict = {"items": [{"cell": "B3"}, {"cell": "B2"}, {"cell": "B4"}]}
    assert compare_results(expected, actual)


def test_compare_results_nested_dicts() -> None:
    expected: JSONDict = {"rule": {"range": "B2:B6", "type": "cellIs"}}
    actual: JSONDict = {"rule": {"range": "B2:B6", "type": "cellIs", "priority": 1}}
    assert compare_results(expected, actual)


def test_compare_results_error_payload_fails() -> None:
    expected: JSONDict = {"type": "string", "value": "x"}
    actual: JSONDict = {"error": "boom"}
    assert not compare_results(expected, actual)
