"""Render formula recalculation comparison results."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def render_calc_results(results: dict[str, Any], output_dir: Path) -> None:
    """Write calculation results as JSON and a concise Markdown report."""
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "results.json").write_text(
        json.dumps(results, indent=2, sort_keys=True) + "\n"
    )
    (output_dir / "README.md").write_text(_render_markdown(results))


def _render_markdown(results: dict[str, Any]) -> str:
    engines = results.get("engines", {})
    lines = [
        "# ExcelBench Formula Recalculation Results",
        "",
        f"Oracle: {results.get('oracle') or 'unknown'}",
        "",
        "Formula values are compared with absolute tolerance 1e-6 or relative tolerance 1e-9 "
        "for numbers; strings and booleans compare exactly.",
        "",
        "| Engine | Status | Matched/total | First mismatches |",
        "| --- | --- | --- | --- |",
    ]
    if not isinstance(engines, dict):
        engines = {}
    for name, result in engines.items():
        if not isinstance(result, dict):
            continue
        mismatches = result.get("mismatched_cells", [])
        mismatch_text = _mismatch_text(mismatches)
        lines.append(
            "| "
            f"{name} | {result.get('status', 'unknown')} | "
            f"{result.get('matched', 0)}/{result.get('total', 0)} | {mismatch_text} |"
        )
        reason = result.get("reason")
        if reason:
            lines.extend(["", f"{name}: {reason}"])
    lines.append("")
    return "\n".join(lines)


def _mismatch_text(mismatches: Any) -> str:
    if not isinstance(mismatches, list) or not mismatches:
        return "—"
    rendered: list[str] = []
    for mismatch in mismatches[:10]:
        if not isinstance(mismatch, dict):
            continue
        rendered.append(
            f"{mismatch.get('cell')}: expected {mismatch.get('expected')!r}, "
            f"actual {mismatch.get('actual')!r}"
        )
    return "<br>".join(rendered).replace("|", "\\|") or "—"
