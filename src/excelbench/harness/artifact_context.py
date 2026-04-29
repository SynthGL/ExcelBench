"""Chart and macro artifact context lanes."""

from __future__ import annotations

import json
import platform
import shutil
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from zipfile import ZipFile

JSONDict = dict[str, Any]


@dataclass(frozen=True)
class ArtifactValidation:
    fixture: str
    passed: bool
    skipped: bool
    checks: dict[str, bool]
    details: dict[str, Any]
    error: str | None = None

    def to_json_dict(self) -> JSONDict:
        return {
            "fixture": self.fixture,
            "passed": self.passed,
            "skipped": self.skipped,
            "checks": self.checks,
            "details": self.details,
            "error": self.error,
        }


def run_chart_context(fixture_dir: Path, output_dir: Path) -> JSONDict:
    """Validate chart-bearing OOXML artifacts structurally."""
    fixtures = sorted(Path(fixture_dir).rglob("*.xlsx"))
    validations = [_validate_chart_fixture(path) for path in fixtures]
    payload = _artifact_payload("charts", fixture_dir, validations)
    _write_artifact_outputs(output_dir, "charts", payload)
    return payload


def run_macro_context(
    fixture_dir: Path,
    output_dir: Path,
    *,
    preserve_with_wolfxl: bool = True,
) -> JSONDict:
    """Validate macro-bearing OOXML artifacts structurally."""
    fixtures = sorted(Path(fixture_dir).rglob("*.xlsm"))
    validations = [
        _validate_macro_fixture(path, output_dir, preserve_with_wolfxl) for path in fixtures
    ]
    if not fixtures:
        validations = [
            ArtifactValidation(
                fixture=str(Path(fixture_dir)),
                passed=False,
                skipped=True,
                checks={},
                details={},
                error="No .xlsm fixtures found.",
            )
        ]
    payload = _artifact_payload("macros", fixture_dir, validations)
    _write_artifact_outputs(output_dir, "macros", payload)
    return payload


def _validate_chart_fixture(path: Path) -> ArtifactValidation:
    try:
        with ZipFile(path) as workbook_zip:
            names = set(workbook_zip.namelist())
            chart_parts = sorted(name for name in names if name.startswith("xl/charts/"))
            drawing_parts = sorted(name for name in names if name.startswith("xl/drawings/drawing"))
            drawing_rels = sorted(name for name in names if name.startswith("xl/drawings/_rels/"))
            worksheet_rels = sorted(
                name for name in names if name.startswith("xl/worksheets/_rels/")
            )
            chart_xml_mentions = 0
            range_mentions = 0
            for part in chart_parts:
                xml = workbook_zip.read(part).decode("utf-8", errors="ignore")
                chart_xml_mentions += int("<c:chart" in xml or "<chart" in xml)
                range_mentions += int("!" in xml and ("<c:f>" in xml or "<f>" in xml))
        checks = {
            "chart_parts_present": bool(chart_parts),
            "drawing_parts_present": bool(drawing_parts),
            "drawing_relationships_present": bool(drawing_rels),
            "worksheet_relationships_present": bool(worksheet_rels),
            "chart_xml_present": chart_xml_mentions > 0,
            "chart_references_present": range_mentions > 0,
        }
        return ArtifactValidation(
            fixture=str(path),
            passed=all(checks.values()),
            skipped=False,
            checks=checks,
            details={
                "chart_parts": chart_parts,
                "drawing_parts": drawing_parts,
                "drawing_rels": drawing_rels,
                "worksheet_rels": worksheet_rels,
            },
        )
    except Exception as exc:
        return ArtifactValidation(
            fixture=str(path),
            passed=False,
            skipped=False,
            checks={},
            details={},
            error=f"{type(exc).__name__}: {exc}",
        )


def _validate_macro_fixture(
    path: Path,
    output_dir: Path,
    preserve_with_wolfxl: bool,
) -> ArtifactValidation:
    try:
        with ZipFile(path) as workbook_zip:
            before = set(workbook_zip.namelist())
            content_types = workbook_zip.read("[Content_Types].xml").decode(
                "utf-8", errors="ignore"
            )
            rels_text = _read_optional(workbook_zip, "xl/_rels/workbook.xml.rels")
        checks = {
            "vba_project_present": "xl/vbaProject.bin" in before,
            "content_type_present": "vbaProject" in content_types,
            "relationship_present": "vbaProject" in rels_text,
        }
        details: JSONDict = {
            "macro_parts": sorted(name for name in before if "vba" in name.lower())
        }
        if preserve_with_wolfxl:
            preserve_checks, preserve_details = _wolfxl_macro_preservation(path, output_dir)
            checks.update(preserve_checks)
            details.update(preserve_details)
        return ArtifactValidation(
            fixture=str(path),
            passed=all(checks.values()),
            skipped=False,
            checks=checks,
            details=details,
        )
    except Exception as exc:
        return ArtifactValidation(
            fixture=str(path),
            passed=False,
            skipped=False,
            checks={},
            details={},
            error=f"{type(exc).__name__}: {exc}",
        )


def _wolfxl_macro_preservation(path: Path, output_dir: Path) -> tuple[dict[str, bool], JSONDict]:
    try:
        import wolfxl

        out_dir = Path(output_dir) / "wolfxl-modified"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / path.name
        shutil.copy2(path, out_path)
        workbook = wolfxl.load_workbook(out_path, modify=True)
        try:
            workbook.save(out_path)
        finally:
            close = getattr(workbook, "close", None)
            if close is not None:
                close()
        with ZipFile(out_path) as workbook_zip:
            after = set(workbook_zip.namelist())
        return (
            {
                "wolfxl_modify_save_completed": True,
                "wolfxl_preserved_vba_project": "xl/vbaProject.bin" in after,
            },
            {"wolfxl_modified_workbook": str(out_path)},
        )
    except Exception as exc:
        return (
            {
                "wolfxl_modify_save_completed": False,
                "wolfxl_preserved_vba_project": False,
            },
            {"wolfxl_error": f"{type(exc).__name__}: {exc}"},
        )


def _artifact_payload(
    kind: str,
    fixture_dir: Path,
    validations: list[ArtifactValidation],
) -> JSONDict:
    return {
        "kind": kind,
        "generated_at": datetime.now(UTC).isoformat(),
        "platform": f"{platform.system()}-{platform.machine()}",
        "fixture_dir": str(fixture_dir),
        "results": [validation.to_json_dict() for validation in validations],
        "passed": all(item.passed or item.skipped for item in validations),
    }


def _write_artifact_outputs(output_dir: Path, kind: str, payload: JSONDict) -> None:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "results.json").write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n")
    (output_dir / "README.md").write_text(_render_artifact_readme(payload))
    (output_dir / "CONTEXT.md").write_text(_render_artifact_context(kind))


def _render_artifact_readme(payload: JSONDict) -> str:
    title = "Chart" if payload["kind"] == "charts" else "Macro"
    lines = [
        f"# ExcelBench {title} Artifact Context",
        "",
        f"- Generated: `{payload['generated_at']}`",
        f"- Fixture dir: `{payload['fixture_dir']}`",
        f"- Passed: `{payload['passed']}`",
        "",
        "| Fixture | Status | Checks |",
        "|---------|--------|--------|",
    ]
    for row in payload["results"]:
        status = "skipped" if row["skipped"] else ("passed" if row["passed"] else "failed")
        checks = ", ".join(f"{k}:{v}" for k, v in row["checks"].items()) or row.get("error") or "-"
        lines.append(f"| {Path(row['fixture']).name} | {status} | {checks} |")
    lines.append("")
    return "\n".join(lines)


def _render_artifact_context(kind: str) -> str:
    return "\n".join(
        [
            f"# {kind.title()} Context",
            "",
            "This is an artifact lane, not a normal scored benchmark lane.",
            "It records package-level evidence for advanced workbook capabilities.",
            "",
        ]
    )


def _read_optional(workbook_zip: ZipFile, part: str) -> str:
    try:
        return workbook_zip.read(part).decode("utf-8", errors="ignore")
    except KeyError:
        return ""
