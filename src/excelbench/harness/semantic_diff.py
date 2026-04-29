"""Structured semantic diffs for Excel workbooks."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from excelbench.harness.workbook_snapshot import WorkbookSnapshot, snapshot_workbook, write_snapshot

JSONDict = dict[str, Any]


@dataclass(frozen=True)
class SemanticDelta:
    """One semantic difference between two workbook snapshots."""

    category: str
    path: str
    left: Any
    right: Any

    def to_json_dict(self) -> JSONDict:
        return {
            "category": self.category,
            "path": self.path,
            "left": self.left,
            "right": self.right,
        }


@dataclass(frozen=True)
class WorkbookDiff:
    """Structured diff result grouped by semantic category."""

    left: str
    right: str
    deltas: tuple[SemanticDelta, ...] = field(default_factory=tuple)

    @property
    def passed(self) -> bool:
        return not self.deltas

    def category_counts(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for delta in self.deltas:
            counts[delta.category] = counts.get(delta.category, 0) + 1
        return dict(sorted(counts.items()))

    def grouped(self) -> dict[str, list[JSONDict]]:
        grouped: dict[str, list[JSONDict]] = {}
        for delta in self.deltas:
            grouped.setdefault(delta.category, []).append(delta.to_json_dict())
        return dict(sorted(grouped.items()))

    def to_json_dict(self) -> JSONDict:
        return {
            "left": self.left,
            "right": self.right,
            "passed": self.passed,
            "delta_count": len(self.deltas),
            "category_counts": self.category_counts(),
            "deltas": [delta.to_json_dict() for delta in self.deltas],
        }


def diff_workbooks(left: Path, right: Path) -> WorkbookDiff:
    """Snapshot and compare two workbooks."""
    return compare_snapshots(snapshot_workbook(left), snapshot_workbook(right))


def compare_snapshots(left: WorkbookSnapshot, right: WorkbookSnapshot) -> WorkbookDiff:
    """Compare two already-built snapshots."""
    deltas: list[SemanticDelta] = []
    categories = sorted(set(left.categories) | set(right.categories))
    for category in categories:
        _collect_deltas(
            category=category,
            path=category,
            left=left.categories.get(category),
            right=right.categories.get(category),
            deltas=deltas,
        )
    return WorkbookDiff(left=left.workbook, right=right.workbook, deltas=tuple(deltas))


def write_diff_artifacts(left: Path, right: Path, output_dir: Path) -> WorkbookDiff:
    """Write snapshot and diff artifacts for two workbooks."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    left_snapshot = snapshot_workbook(left)
    right_snapshot = snapshot_workbook(right)
    diff = compare_snapshots(left_snapshot, right_snapshot)

    write_snapshot(left_snapshot, output_dir / "left.snapshot.json")
    write_snapshot(right_snapshot, output_dir / "right.snapshot.json")
    (output_dir / "summary.json").write_text(
        json.dumps(diff.to_json_dict(), indent=2, sort_keys=True) + "\n"
    )
    (output_dir / "summary.md").write_text(render_diff_markdown(diff))

    categories_dir = output_dir / "categories"
    categories_dir.mkdir(exist_ok=True)
    for category, items in diff.grouped().items():
        (categories_dir / f"{category}.json").write_text(
            json.dumps(items, indent=2, sort_keys=True) + "\n"
        )
    return diff


def render_diff_markdown(diff: WorkbookDiff) -> str:
    lines = [
        "# Workbook Semantic Diff",
        "",
        f"- Left: `{diff.left}`",
        f"- Right: `{diff.right}`",
        f"- Passed: `{diff.passed}`",
        f"- Delta count: `{len(diff.deltas)}`",
        "",
    ]
    if not diff.deltas:
        lines.extend(["No semantic deltas detected.", ""])
        return "\n".join(lines)

    lines.extend(["## Category Summary", "", "| Category | Deltas |", "|----------|--------|"])
    for category, count in diff.category_counts().items():
        lines.append(f"| {category} | {count} |")
    lines.append("")
    lines.extend(["## Deltas", ""])
    for delta in diff.deltas[:200]:
        lines.append(
            f"- `{delta.category}` `{delta.path}`: "
            f"`{_short(delta.left)}` -> `{_short(delta.right)}`"
        )
    if len(diff.deltas) > 200:
        lines.append(f"- ... {len(diff.deltas) - 200} additional deltas omitted from markdown")
    lines.append("")
    return "\n".join(lines)


def _collect_deltas(
    *,
    category: str,
    path: str,
    left: Any,
    right: Any,
    deltas: list[SemanticDelta],
) -> None:
    if isinstance(left, dict) and isinstance(right, dict):
        for key in sorted(set(left) | set(right), key=str):
            _collect_deltas(
                category=category,
                path=f"{path}.{key}",
                left=left.get(key),
                right=right.get(key),
                deltas=deltas,
            )
        return
    if isinstance(left, list) and isinstance(right, list):
        if left == right:
            return
        max_len = max(len(left), len(right))
        for index in range(max_len):
            _collect_deltas(
                category=category,
                path=f"{path}[{index}]",
                left=left[index] if index < len(left) else None,
                right=right[index] if index < len(right) else None,
                deltas=deltas,
            )
        return
    if left != right:
        deltas.append(SemanticDelta(category=category, path=path, left=left, right=right))


def _short(value: Any) -> str:
    text = json.dumps(value, sort_keys=True, default=str)
    return text if len(text) <= 180 else f"{text[:177]}..."
