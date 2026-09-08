"""Rendering for surgical workbook mutation benchmark results."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def _format_number(value: object, suffix: str = "") -> str:
    """Format an optional benchmark number for a Markdown cell."""
    if isinstance(value, int | float):
        return f"{value:.3f}{suffix}"
    return "—"


def _verdict(result: dict[str, Any]) -> str:
    """Return a concise status label for an engine result."""
    status = result.get("status")
    if status == "unavailable":
        return "Unavailable"
    if status == "failed":
        return "Failed"
    if status == "integrity-failed":
        return "Mutation integrity failed"
    if result.get("preservation_score") == 100:
        return "Preserved"
    return "Loss detected"


def render_mutation_report(results: dict[str, Any], output_dir: Path) -> None:
    """Write ``results.json`` and a compact Markdown comparison report."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "results.json").write_text(
        json.dumps(results, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )

    metadata = results.get("metadata", {})
    engine_results = results.get("engines", {})
    lines = [
        "# ExcelBench Mutation Results",
        "",
        f"*Generated: {metadata.get('timestamp', 'unknown')}*",
        f"*Template SHA-256: {metadata.get('template_sha256', 'unknown')}*",
        f"*Repeats: {metadata.get('repeats', 'unknown')}*",
        "",
        "## Comparison",
        "",
        "| Engine | Wall ms | Peak RSS KB | Preservation score | Lost parts | Verdict |",
        "|--------|---------|-------------|--------------------|------------|---------|",
    ]
    for engine_name in sorted(engine_results):
        result = engine_results[engine_name]
        details = result.get("details", {})
        lost_parts = details.get("lost_parts", [])
        lost_count = len(lost_parts) if isinstance(lost_parts, list) else "—"
        score = result.get("preservation_score")
        score_text = _format_number(score) if isinstance(score, int | float) else "—"
        lines.append(
            f"| {engine_name} | {_format_number(result.get('wall_ms_median'))} | "
            f"{_format_number(result.get('peak_rss_kb_median'))} | {score_text} | "
            f"{lost_count} | {_verdict(result)} |"
        )
    lines.extend(
        [
            "",
            "## Notes",
            "",
            (
                "Wall time and peak RSS are measured in isolated subprocesses "
                "with macOS `/usr/bin/time -l`."
            ),
            (
                "Preservation compares each output package against the original template. "
                "It does not measure formula recalculation or rendering."
            ),
            "",
        ]
    )
    (output_dir / "README.md").write_text("\n".join(lines), encoding="utf-8")
