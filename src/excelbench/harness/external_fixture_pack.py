"""Generate local external-oracle fixture packs."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from zipfile import ZipFile

from excelbench.harness.external_fixture_specs import (
    ExternalFixtureSpec,
    closedxml_fixture_specs,
    excelize_fixture_specs,
    exceljs_fixture_specs,
    external_fixture_specs,
    npoi_fixture_specs,
)
from excelbench.harness.external_fixture_specs.base import JSONDict
from excelbench.harness.external_oracles import (
    ExternalOracleRequest,
    ExternalOracleResult,
    external_oracle_catalog,
    run_external_oracle,
)

__all__ = [
    "ExternalFixtureSpec",
    "FixtureGenerationResult",
    "closedxml_fixture_specs",
    "exceljs_fixture_specs",
    "excelize_fixture_specs",
    "external_fixture_specs",
    "generate_external_fixture_pack",
    "npoi_fixture_specs",
]


@dataclass(frozen=True)
class FixtureGenerationResult:
    """Result for one generated external fixture."""

    fixture_id: str
    tool: str
    workbook_path: Path
    write_result: ExternalOracleResult
    expected_parts: tuple[str, ...]
    missing_parts: tuple[str, ...]
    validations: tuple[ExternalOracleResult, ...] = field(default_factory=tuple)

    @property
    def passed(self) -> bool:
        """Return whether generation and requested validations passed."""
        return (
            self.write_result.passed
            and not self.missing_parts
            and all(result.passed or result.skipped for result in self.validations)
        )

    def to_json_dict(self, output_root: Path) -> JSONDict:
        """Convert the result to a manifest entry."""
        return {
            "fixture_id": self.fixture_id,
            "tool": self.tool,
            "workbook": str(self.workbook_path.relative_to(output_root)),
            "passed": self.passed,
            "expected_parts": list(self.expected_parts),
            "missing_parts": list(self.missing_parts),
            "write_result": _oracle_result_to_json(self.write_result),
            "validations": [_oracle_result_to_json(result) for result in self.validations],
        }


def generate_external_fixture_pack(
    output_root: Path,
    *,
    repo_root: Path,
    include_validators: bool = True,
    timeout_seconds: float = 180.0,
) -> list[FixtureGenerationResult]:
    """Generate the local external fixture pack and write ``manifest.json``."""
    output_root = output_root.resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    catalog = external_oracle_catalog(repo_root=repo_root)
    results: list[FixtureGenerationResult] = []

    for spec in external_fixture_specs():
        if not catalog[spec.tool].is_available():
            continue
        workbook_path = output_root / spec.filename
        write_result = run_external_oracle(
            catalog[spec.tool],
            ExternalOracleRequest(
                fixture_id=spec.fixture_id,
                operation="write_fixture",
                output_path=workbook_path,
                payload=spec.payload,
            ),
            timeout_seconds=timeout_seconds,
        )
        missing_parts = (
            _missing_parts(workbook_path, spec.expected_parts) if write_result.passed else ()
        )
        validations: list[ExternalOracleResult] = []
        if include_validators and write_result.passed:
            validations.extend(
                [
                    run_external_oracle(
                        catalog["libreoffice"],
                        ExternalOracleRequest(
                            fixture_id=spec.fixture_id,
                            operation="open_save_validate",
                            input_path=workbook_path,
                            output_path=output_root / "validated" / spec.filename,
                            payload={},
                        ),
                        timeout_seconds=timeout_seconds,
                    ),
                    run_external_oracle(
                        catalog["libreoffice"],
                        ExternalOracleRequest(
                            fixture_id=spec.fixture_id,
                            operation="render_validate",
                            input_path=workbook_path,
                            output_path=output_root / "pdf" / f"{workbook_path.stem}.pdf",
                            payload={},
                        ),
                        timeout_seconds=timeout_seconds,
                    ),
                ]
            )
        results.append(
            FixtureGenerationResult(
                fixture_id=spec.fixture_id,
                tool=spec.tool,
                workbook_path=workbook_path,
                write_result=write_result,
                expected_parts=spec.expected_parts,
                missing_parts=missing_parts,
                validations=tuple(validations),
            )
        )

    manifest = {
        "generated_at": datetime.now(UTC).isoformat(),
        "output_root": str(output_root),
        "fixtures": [result.to_json_dict(output_root) for result in results],
    }
    manifest_text = json.dumps(manifest, indent=2, sort_keys=True) + "\n"
    (output_root / "manifest.json").write_text(manifest_text)
    return results


def _missing_parts(workbook_path: Path, expected_parts: tuple[str, ...]) -> tuple[str, ...]:
    if not workbook_path.exists():
        return expected_parts
    with ZipFile(workbook_path) as workbook_zip:
        names = set(workbook_zip.namelist())
    return tuple(part for part in expected_parts if part not in names)


def _oracle_result_to_json(result: ExternalOracleResult) -> JSONDict:
    return {
        "tool_name": result.tool_name,
        "passed": result.passed,
        "skipped": result.skipped,
        "returncode": result.returncode,
        "payload": result.payload,
        "notes": result.notes,
    }
