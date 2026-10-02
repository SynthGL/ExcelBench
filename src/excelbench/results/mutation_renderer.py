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


def mutation_verdict(result: dict[str, Any]) -> str:
    """Return the verdict for one engine row.

    Preserved / Changed <features> / Edits missing / Failed, plus Unavailable for
    engines that were not run.
    """
    status = result.get("status")
    if status == "unavailable":
        return "Unavailable"
    if status == "failed":
        return "Failed"
    if not result.get("edits_applied"):
        return "Edits missing"
    if result.get("preserved"):
        return "Preserved"
    changed = result.get("features_changed") or {}
    return "Changed " + ", ".join(changed)


def _changed_text(result: dict[str, Any]) -> str:
    """Summarize unexpected differences per feature for a Markdown cell."""
    changed = result.get("features_changed")
    if not isinstance(changed, dict):
        return "—"
    if not changed:
        return "none"
    return ", ".join(f"{feature} ({count})" for feature, count in changed.items())


def _integrity_text(result: dict[str, Any]) -> str:
    """Summarize package integrity issues (unscored) for a Markdown cell."""
    integrity = result.get("details", {}).get("integrity")
    if not isinstance(integrity, dict):
        return "—"
    issues = len(integrity.get("dangling_rels", [])) + len(
        integrity.get("missing_content_types", [])
    )
    return "ok" if issues == 0 else f"{issues} issue(s)"


def render_mutation_report(results: dict[str, Any], output_dir: Path) -> None:
    """Write ``results.json`` and a compact Markdown comparison report."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    engine_results = {
        name: {**result, "verdict": mutation_verdict(result)}
        for name, result in results.get("engines", {}).items()
    }
    persisted = {**results, "engines": engine_results}
    (output_dir / "results.json").write_text(
        json.dumps(persisted, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )

    metadata = results.get("metadata", {})
    lines = [
        "# ExcelBench Mutation Results",
        "",
        f"*Generated: {metadata.get('timestamp', 'unknown')}*",
        f"*Template SHA-256: {metadata.get('template_sha256', 'unknown')}*",
        f"*Repeats: {metadata.get('repeats', 'unknown')}*",
        "",
        "## Comparison",
        "",
        (
            "| Engine | Wall ms | Peak RSS KB | Features preserved | Unexpected changes "
            "| Edits applied | Integrity | Verdict |"
        ),
        (
            "|--------|---------|-------------|--------------------|--------------------"
            "|---------------|-----------|---------|"
        ),
    ]
    for engine_name in sorted(engine_results):
        result = engine_results[engine_name]
        edits_applied = result.get("edits_applied")
        edits_text = "—" if edits_applied is None else ("yes" if edits_applied else "no")
        lines.append(
            f"| {engine_name} | {_format_number(result.get('wall_ms_median'))} | "
            f"{_format_number(result.get('peak_rss_kb_median'))} | "
            f"{result.get('feature_preservation') or '—'} | {_changed_text(result)} | "
            f"{edits_text} | {_integrity_text(result)} | {result['verdict']} |"
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
                "Preservation compares the output's workbook content model against the "
                "template's model with the two declared cell edits applied. Part names, "
                "relationship ids, XML serialization and document metadata are not "
                "content, so an equivalent rewrite scores the same as a byte copy. "
                "Cached results of formula cells are recalculation state and are not "
                "compared; formula text is. Features preserved counts template features "
                "with no unexpected difference. Integrity (dangling relationships, parts "
                "without a content type) is reported separately and is not scored. "
                "This does not measure formula recalculation or rendering."
            ),
            "",
        ]
    )
    (output_dir / "README.md").write_text("\n".join(lines), encoding="utf-8")
