"""Tests for optional external spreadsheet oracle helpers."""

from __future__ import annotations

import sys
from pathlib import Path

from excelbench.harness.external_oracles import (
    ExternalOracleRequest,
    ExternalOracleTool,
    external_oracle_catalog,
    run_external_oracle,
)


def test_catalog_lists_planned_external_oracles() -> None:
    catalog = external_oracle_catalog()

    assert set(catalog) == {"excelize", "libreoffice", "apache-poi", "closedxml"}
    assert "pivots" in catalog["excelize"].capabilities
    assert "open_save_validate" in catalog["libreoffice"].capabilities


def test_missing_oracle_helper_is_structured_skip() -> None:
    tool = ExternalOracleTool(
        name="missing",
        command=("excelbench-helper-that-does-not-exist",),
        language="test",
        homepage="https://example.invalid",
        capabilities=frozenset({"write"}),
    )

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="fixture",
            operation="write_fixture",
            payload={"cells": []},
        ),
    )

    assert result.skipped is True
    assert result.passed is False
    assert result.returncode is None
    assert result.notes is not None
    assert "not found" in result.notes


def test_oracle_request_serializes_paths() -> None:
    request = ExternalOracleRequest(
        fixture_id="pivot-cache",
        operation="open_save_validate",
        payload={"feature": "pivot_tables"},
        input_path=Path("input.xlsx"),
        output_path=Path("output.xlsx"),
    )

    assert request.to_json_dict() == {
        "fixture_id": "pivot-cache",
        "operation": "open_save_validate",
        "payload": {"feature": "pivot_tables"},
        "input_path": "input.xlsx",
        "output_path": "output.xlsx",
    }


def test_successful_oracle_helper_parses_json_stdout() -> None:
    tool = ExternalOracleTool(
        name="python-json-helper",
        command=(
            sys.executable,
            "-c",
            (
                "import json, sys; "
                "request = json.load(sys.stdin); "
                "print(json.dumps({'fixture_id': request['fixture_id'], 'ok': True}))"
            ),
        ),
        language="python",
        homepage="https://example.invalid",
        capabilities=frozenset({"write"}),
    )

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="rich-text",
            operation="write_fixture",
            payload={"cells": [{"cell": "A1", "value": "Hello"}]},
        ),
    )

    assert result.passed is True
    assert result.skipped is False
    assert result.returncode == 0
    assert result.payload == {"fixture_id": "rich-text", "ok": True}


def test_nonzero_oracle_helper_is_failure() -> None:
    tool = ExternalOracleTool(
        name="python-failing-helper",
        command=(sys.executable, "-c", "import sys; sys.stderr.write('boom'); sys.exit(2)"),
        language="python",
        homepage="https://example.invalid",
        capabilities=frozenset({"write"}),
    )

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="broken",
            operation="write_fixture",
            payload={},
        ),
    )

    assert result.passed is False
    assert result.skipped is False
    assert result.returncode == 2
    assert result.stderr == "boom"


def test_invalid_json_stdout_is_failure_payload() -> None:
    tool = ExternalOracleTool(
        name="python-invalid-json-helper",
        command=(sys.executable, "-c", "print('not json')"),
        language="python",
        homepage="https://example.invalid",
        capabilities=frozenset({"write"}),
    )

    result = run_external_oracle(
        tool,
        ExternalOracleRequest(
            fixture_id="bad-json",
            operation="write_fixture",
            payload={},
        ),
    )

    assert result.passed is False
    assert result.payload["error"] == "invalid_json_stdout"

