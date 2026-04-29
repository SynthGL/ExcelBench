"""Roundtrip/idempotence context lane."""

from __future__ import annotations

import json
import platform
import shutil
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from excelbench.generator.generate import load_manifest
from excelbench.harness.adapters import ExcelAdapter
from excelbench.harness.semantic_diff import write_diff_artifacts

JSONDict = dict[str, Any]


@dataclass(frozen=True)
class RoundtripResult:
    adapter: str
    workbook: str
    cycle: int
    passed: bool
    skipped: bool
    delta_count: int
    category_counts: dict[str, int]
    output_workbook: str | None = None
    diff_dir: str | None = None
    error: str | None = None

    def to_json_dict(self) -> JSONDict:
        return {
            "adapter": self.adapter,
            "workbook": self.workbook,
            "cycle": self.cycle,
            "passed": self.passed,
            "skipped": self.skipped,
            "delta_count": self.delta_count,
            "category_counts": self.category_counts,
            "output_workbook": self.output_workbook,
            "diff_dir": self.diff_dir,
            "error": self.error,
        }


def run_roundtrip_context(
    test_dir: Path,
    output_dir: Path,
    *,
    adapters: list[ExcelAdapter],
    cycles: int = 2,
) -> list[RoundtripResult]:
    """Run roundtrip drift checks for selected adapters and fixture workbooks."""
    test_dir = Path(test_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    workbooks = _discover_workbooks(test_dir)
    results: list[RoundtripResult] = []

    for adapter in adapters:
        for workbook_path in workbooks:
            results.extend(
                _roundtrip_one_adapter(adapter, workbook_path, output_dir, cycles=cycles)
            )

    payload = {
        "generated_at": datetime.now(UTC).isoformat(),
        "platform": f"{platform.system()}-{platform.machine()}",
        "test_dir": str(test_dir),
        "cycles": cycles,
        "results": [result.to_json_dict() for result in results],
        "passed": all(result.passed or result.skipped for result in results),
    }
    (output_dir / "roundtrip.json").write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n")
    (output_dir / "ROUNDTRIP.md").write_text(_render_roundtrip_markdown(payload))
    (output_dir / "CONTEXT.md").write_text(_render_roundtrip_context(adapters, cycles))
    return results


def _roundtrip_one_adapter(
    adapter: ExcelAdapter,
    workbook_path: Path,
    output_dir: Path,
    *,
    cycles: int,
) -> list[RoundtripResult]:
    if not _can_roundtrip(adapter):
        return [
            RoundtripResult(
                adapter=adapter.name,
                workbook=str(workbook_path),
                cycle=0,
                passed=False,
                skipped=True,
                delta_count=0,
                category_counts={},
                error=(
                    "Adapter does not expose an openpyxl-compatible read-modify-save path "
                    "for semantic roundtrip checks."
                ),
            )
        ]

    current = workbook_path
    results: list[RoundtripResult] = []
    for cycle in range(1, cycles + 1):
        cycle_dir = output_dir / "workbooks" / adapter.name
        cycle_dir.mkdir(parents=True, exist_ok=True)
        out_path = cycle_dir / f"{workbook_path.stem}.cycle{cycle}{workbook_path.suffix}"
        try:
            _save_roundtrip(adapter, current, out_path)
            diff_dir = output_dir / "diffs" / adapter.name / workbook_path.stem / f"cycle{cycle}"
            diff = write_diff_artifacts(workbook_path, out_path, diff_dir)
            results.append(
                RoundtripResult(
                    adapter=adapter.name,
                    workbook=str(workbook_path),
                    cycle=cycle,
                    passed=_roundtrip_passed(diff.category_counts()),
                    skipped=False,
                    delta_count=len(diff.deltas),
                    category_counts=diff.category_counts(),
                    output_workbook=str(out_path),
                    diff_dir=str(diff_dir),
                )
            )
            current = out_path
        except Exception as exc:
            results.append(
                RoundtripResult(
                    adapter=adapter.name,
                    workbook=str(workbook_path),
                    cycle=cycle,
                    passed=False,
                    skipped=False,
                    delta_count=0,
                    category_counts={},
                    output_workbook=str(out_path),
                    error=f"{type(exc).__name__}: {exc}",
                )
            )
            break
    return results


def _discover_workbooks(test_dir: Path) -> list[Path]:
    manifest_path = test_dir / "manifest.json"
    if manifest_path.exists():
        manifest = load_manifest(manifest_path)
        return [test_dir / f.path for f in manifest.files if (test_dir / f.path).suffix == ".xlsx"]
    return sorted(test_dir.rglob("*.xlsx"))


def _roundtrip_passed(category_counts: dict[str, int]) -> bool:
    """Package hash churn is informational; workbook-semantic drift fails."""
    return not {category for category in category_counts if category != "package_parts"}


def _can_roundtrip(adapter: ExcelAdapter) -> bool:
    return adapter.name in {"openpyxl", "wolfxl"}


def _save_roundtrip(adapter: ExcelAdapter, input_path: Path, output_path: Path) -> None:
    if adapter.name == "openpyxl":
        import openpyxl

        workbook = openpyxl.load_workbook(input_path, data_only=False)
        try:
            workbook.save(output_path)
        finally:
            workbook.close()
        return
    if adapter.name == "wolfxl":
        import wolfxl

        shutil.copy2(input_path, output_path)
        workbook = wolfxl.load_workbook(output_path, modify=True)
        try:
            workbook.save(output_path)
        finally:
            close = getattr(workbook, "close", None)
            if close is not None:
                close()
        return
    raise NotImplementedError(f"{adapter.name} has no roundtrip implementation")


def _render_roundtrip_markdown(payload: JSONDict) -> str:
    rows = payload["results"]
    lines = [
        "# ExcelBench Roundtrip Context",
        "",
        f"- Generated: `{payload['generated_at']}`",
        f"- Cycles: `{payload['cycles']}`",
        f"- Passed: `{payload['passed']}`",
        "",
        "| Adapter | Workbook | Cycle | Status | Deltas | Categories |",
        "|---------|----------|-------|--------|--------|------------|",
    ]
    for row in rows:
        if row["skipped"]:
            status = "skipped"
        elif row["passed"]:
            status = "passed"
        else:
            status = "failed"
        categories = ", ".join(f"{k}:{v}" for k, v in row["category_counts"].items())
        lines.append(
            f"| {row['adapter']} | {Path(row['workbook']).name} | {row['cycle']} | "
            f"{status} | {row['delta_count']} | {categories or '-'} |"
        )
    lines.append("")
    return "\n".join(lines)


def _render_roundtrip_context(adapters: list[ExcelAdapter], cycles: int) -> str:
    adapter_names = ", ".join(adapter.name for adapter in adapters)
    return "\n".join(
        [
            "# Roundtrip Context",
            "",
            "This lane measures semantic drift after repeated open/save cycles.",
            "",
            f"- Requested adapters: {adapter_names}",
            f"- Cycles: {cycles}",
            "- Unsupported adapters are explicit skips, not silent passes.",
            "- Diffs are generated with the semantic workbook diff tool.",
            "",
        ]
    )
