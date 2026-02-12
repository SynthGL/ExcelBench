"""Results rendering to various output formats."""

import json
import subprocess
from pathlib import Path
from typing import Any

from excelbench.models import BenchmarkResults, Diagnostic, FeatureScore, OperationType, TestResult

# Feature tier assignments
_FEATURE_TIERS: dict[str, tuple[int, str]] = {
    "cell_values": (0, "Basic Values"),
    "formulas": (0, "Basic Values"),
    "text_formatting": (1, "Formatting"),
    "background_colors": (1, "Formatting"),
    "number_formats": (1, "Formatting"),
    "alignment": (1, "Formatting"),
    "borders": (1, "Formatting"),
    "dimensions": (1, "Formatting"),
    "multiple_sheets": (0, "Basic Values"),
    "merged_cells": (2, "Advanced"),
    "conditional_formatting": (2, "Advanced"),
    "data_validation": (2, "Advanced"),
    "hyperlinks": (2, "Advanced"),
    "images": (2, "Advanced"),
    "pivot_tables": (2, "Advanced"),
    "comments": (2, "Advanced"),
    "freeze_panes": (2, "Advanced"),
    "named_ranges": (3, "Workbook Metadata"),
    "tables": (3, "Workbook Metadata"),
}

_TIER_LABELS = {
    0: "Tier 0 — Basic Values",
    1: "Tier 1 — Formatting",
    2: "Tier 2 — Advanced",
    3: "Tier 3 — Workbook Metadata",
}

# Short display names for headline matrix
_SHORT_NAMES: dict[str, str] = {
    "openpyxl": "openpyxl",
    "xlsxwriter": "xlsxwriter",
    "xlsxwriter-constmem": "xlsx-constmem",
    "openpyxl-readonly": "opxl-readonly",
    "python-calamine": "calamine",
    "pylightxl": "pylightxl",
    "pyexcel": "pyexcel",
    "pandas": "pandas",
    "polars": "polars",
    "tablib": "tablib",
    "xlrd": "xlrd",
    "xlwt": "xlwt",
}

# Short display names for features in headline matrix
_SHORT_FEATURE_NAMES: dict[str, str] = {
    "cell_values": "Cell Values",
    "formulas": "Formulas",
    "multiple_sheets": "Sheets",
    "text_formatting": "Text Fmt",
    "background_colors": "Bg Colors",
    "number_formats": "Num Fmt",
    "alignment": "Alignment",
    "borders": "Borders",
    "dimensions": "Dimensions",
    "merged_cells": "Merged",
    "conditional_formatting": "Cond Fmt",
    "data_validation": "Validation",
    "hyperlinks": "Hyperlinks",
    "images": "Images",
    "comments": "Comments",
    "freeze_panes": "Freeze",
    "pivot_tables": "Pivots",
    "named_ranges": "Named Ranges",
    "tables": "Tables",
}


def render_results(results: BenchmarkResults, output_dir: Path) -> None:
    """Render results to all output formats."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    render_json(results, output_dir / "results.json")
    render_markdown(results, output_dir / "README.md")
    render_csv(results, output_dir / "matrix.csv")
    _append_history(results, output_dir)


def render_json(results: BenchmarkResults, path: Path) -> None:
    """Render results to JSON."""
    data = {
        "metadata": {
            "benchmark_version": results.metadata.benchmark_version,
            "run_date": results.metadata.run_date.isoformat(),
            "excel_version": results.metadata.excel_version,
            "platform": results.metadata.platform,
            "profile": results.metadata.profile,
        },
        "libraries": {
            name: {
                "name": info.name,
                "version": info.version,
                "language": info.language,
                "capabilities": list(info.capabilities),
            }
            for name, info in results.libraries.items()
        },
        "results": [
            {
                "feature": score.feature,
                "library": score.library,
                "scores": {
                    "read": score.read_score,
                    "write": score.write_score,
                },
                "test_cases": _group_test_cases(score.test_results),
                "notes": score.notes,
            }
            for score in results.scores
        ],
    }

    with open(path, "w") as f:
        json.dump(data, f, indent=2)


def render_markdown(results: BenchmarkResults, path: Path) -> None:
    """Render results to markdown summary."""
    lines: list[str] = []

    # Header
    lines.append("# ExcelBench Results")
    lines.append("")
    lines.append(f"*Generated: {results.metadata.run_date.strftime('%Y-%m-%d %H:%M UTC')}*")
    lines.append(f"*Profile: {results.metadata.profile}*")
    lines.append(f"*Excel Version: {results.metadata.excel_version}*")
    lines.append(f"*Platform: {results.metadata.platform}*")
    lines.append("")

    # Build lookups used across sections
    features = sorted(set(s.feature for s in results.scores))
    libraries = sorted(results.libraries.keys())

    score_lookup: dict[tuple[str, str], FeatureScore] = {}
    for score_entry in results.scores:
        score_lookup[(score_entry.feature, score_entry.library)] = score_entry

    # ── T0-1: Headline matrix (condensed) ──
    lines.extend(_render_headline_matrix(features, libraries, score_lookup))

    # ── T0-2: Library tier list ──
    lines.extend(_render_tier_list(results, features, libraries, score_lookup))

    # Legend
    lines.append("## Score Legend")
    lines.append("")
    lines.append("| Score | Meaning |")
    lines.append("|-------|---------|")
    lines.append("| 🟢 3 | Complete — full fidelity |")
    lines.append("| 🟡 2 | Functional — works for common cases |")
    lines.append("| 🟠 1 | Minimal — basic recognition only |")
    lines.append("| 🔴 0 | Unsupported — errors or data loss |")
    lines.append("| ➖ | Not applicable |")
    lines.append("")

    # Full summary table — grouped by tier (unchanged structure)
    lines.append("## Full Results Matrix")
    lines.append("")

    header = "| Feature |"
    separator = "|---------|"
    for lib in libraries:
        caps = results.libraries[lib].capabilities
        if "read" in caps and "write" in caps:
            header += f" {lib} (R) | {lib} (W) |"
            separator += "------------|------------|"
        elif "read" in caps:
            header += f" {lib} (R) |"
            separator += "------------|"
        else:
            header += f" {lib} (W) |"
            separator += "------------|"

    # Group features by tier
    tier_features: dict[int, list[str]] = {0: [], 1: [], 2: [], 3: []}
    for feature_name in features:
        tier = _FEATURE_TIERS.get(feature_name, (2, "Advanced"))[0]
        tier_features[tier].append(feature_name)

    for tier_num in sorted(tier_features.keys()):
        tier_list = tier_features[tier_num]
        if not tier_list:
            continue
        lines.append(f"**{_TIER_LABELS[tier_num]}**")
        lines.append("")
        lines.append(header)
        lines.append(separator)

        for feature in tier_list:
            row = f"| [{feature}](#{feature}-details) |"
            for lib in libraries:
                score = score_lookup.get((feature, lib))
                caps = results.libraries[lib].capabilities
                if "read" in caps:
                    if score and score.read_score is not None:
                        row += f" {score_emoji(score.read_score)} |"
                    else:
                        row += " ➖ |"
                if "write" in caps:
                    if score and score.write_score is not None:
                        row += f" {score_emoji(score.write_score)} |"
                    else:
                        row += " ➖ |"
            lines.append(row)
        lines.append("")

    # ── T0-3: Deduplicated notes ──
    lines.extend(_render_notes_deduped(results))

    # Statistics section
    lines.extend(_render_statistics(results, libraries, features, score_lookup))

    # Library details
    lines.append("## Libraries Tested")
    lines.append("")
    for name, info in sorted(results.libraries.items()):
        caps_str = ", ".join(sorted(info.capabilities))
        lines.append(f"- **{name}** v{info.version} ({info.language}) - {caps_str}")
    lines.append("")

    # Diagnostic summaries
    lines.append("## Diagnostics Summary")
    lines.append("")
    lines.extend(_render_diagnostics_summary(results))

    # Detailed results per feature
    lines.append("## Detailed Results")
    lines.append("")

    for feature in features:
        lines.append(f'<a id="{feature}-details"></a>')
        lines.append(f"### {feature}")
        lines.append("")

        for lib in libraries:
            score = score_lookup.get((feature, lib))
            if not score:
                continue

            # Compact header: **lib** — Read: emoji | Write: emoji
            parts = [f"**{lib}**"]
            score_parts = []
            if score.read_score is not None:
                score_parts.append(f"Read: {score_emoji(score.read_score)}")
            if score.write_score is not None:
                score_parts.append(f"Write: {score_emoji(score.write_score)}")
            if score_parts:
                parts.append(" — " + " | ".join(score_parts))
            lines.append("".join(parts))

            if score.notes:
                lines.append(f"- Notes: {score.notes}")

            # Show per-test breakdown table if there are failures
            failed = [tr for tr in score.test_results if not tr.passed]
            if failed:
                lines.append("")
                lines.extend(_render_per_test_table(score))

            lines.append("")

    # Footer
    lines.append("---")
    lines.append(f"*Benchmark version: {results.metadata.benchmark_version}*")

    with open(path, "w") as fp:
        fp.write("\n".join(lines))


# ── T0-1: Headline matrix ──────────────────────────────────────────────────


def _render_headline_matrix(
    features: list[str],
    libraries: list[str],
    score_lookup: dict[tuple[str, str], FeatureScore],
) -> list[str]:
    """Render a condensed overview matrix: one column per library, best of R/W."""
    lines: list[str] = []

    # Filter out libraries that have NO scored features (e.g. xlrd on xlsx profile)
    scored_libs: list[str] = []
    for lib in libraries:
        has_score = any(
            (s := score_lookup.get((f, lib)))
            and (s.read_score is not None or s.write_score is not None)
            for f in features
        )
        if has_score:
            scored_libs.append(lib)

    # Filter out features where every library is N/A (e.g. pivot_tables)
    scored_features: list[str] = []
    for feat in features:
        has_score = any(
            (s := score_lookup.get((feat, lib)))
            and (s.read_score is not None or s.write_score is not None)
            for lib in scored_libs
        )
        if has_score:
            scored_features.append(feat)

    if not scored_libs or not scored_features:
        return lines

    lines.append("## Overview")
    lines.append("")
    lines.append(
        "> Condensed view — shows the **best score** across read/write for each library. "
        "See [Full Results Matrix](#full-results-matrix) for the complete R/W breakdown."
    )
    lines.append("")

    # Build header with short names
    header = "| Feature |"
    sep = "|---------|"
    for lib in scored_libs:
        short = _SHORT_NAMES.get(lib, lib)
        header += f" {short} |"
        sep += ":-:|"

    # Group features by tier
    tier_features: dict[int, list[str]] = {0: [], 1: [], 2: [], 3: []}
    for feat in scored_features:
        tier = _FEATURE_TIERS.get(feat, (2, "Advanced"))[0]
        tier_features[tier].append(feat)

    for tier_num in sorted(tier_features.keys()):
        tier_list = tier_features[tier_num]
        if not tier_list:
            continue
        lines.append(f"**{_TIER_LABELS[tier_num]}**")
        lines.append("")
        lines.append(header)
        lines.append(sep)

        for feat in tier_list:
            short_feat = _SHORT_FEATURE_NAMES.get(feat, feat)
            row = f"| {short_feat} |"
            for lib in scored_libs:
                score = score_lookup.get((feat, lib))
                best = _best_score(score)
                row += f" {_score_icon(best)} |"
            lines.append(row)
        lines.append("")

    return lines


def _best_score(score: FeatureScore | None) -> int | None:
    """Return the best (max) of read and write scores, or None if both missing."""
    if score is None:
        return None
    r = score.read_score
    w = score.write_score
    if r is not None and w is not None:
        return max(r, w)
    return r if r is not None else w


def _score_icon(score: int | None) -> str:
    """Compact emoji-only score (no number) for headline matrix."""
    if score is None:
        return "➖"
    if score == 3:
        return "🟢"
    if score == 2:
        return "🟡"
    if score == 1:
        return "🟠"
    return "🔴"


# ── T0-2: Library tier list ────────────────────────────────────────────────


def _render_tier_list(
    results: BenchmarkResults,
    features: list[str],
    libraries: list[str],
    score_lookup: dict[tuple[str, str], FeatureScore],
) -> list[str]:
    """Render a tier list grouping libraries by capability level."""
    lines: list[str] = []

    # Compute stats per library: best green features (max of R green, W green)
    lib_stats: list[tuple[str, int, int, str]] = []  # (lib, best_green, total_scored, caps_str)
    for lib in libraries:
        caps = results.libraries[lib].capabilities
        has_rw = "read" in caps and "write" in caps
        caps_label = "R+W" if has_rw else ("R" if "read" in caps else "W")

        r_green = 0
        w_green = 0
        r_scored = 0
        w_scored = 0
        for feat in features:
            score = score_lookup.get((feat, lib))
            if not score:
                continue
            if score.read_score is not None:
                r_scored += 1
                if score.read_score == 3:
                    r_green += 1
            if score.write_score is not None:
                w_scored += 1
                if score.write_score == 3:
                    w_green += 1

        best_green = max(r_green, w_green)
        total_scored = max(r_scored, w_scored)
        if total_scored == 0:
            continue
        lib_stats.append((lib, best_green, total_scored, caps_label))

    if not lib_stats:
        return lines

    # Sort by best_green descending
    lib_stats.sort(key=lambda x: (-x[1], x[0]))

    # Assign tiers
    tier_defs = [
        ("S", "Full Fidelity", lambda g, t: g >= t and t > 0),
        ("A", "Near-Complete", lambda g, t: g >= t * 0.8 and g < t),
        ("B", "Partial", lambda g, _t: g >= 4),
        ("C", "Basic", lambda g, _t: g >= 1),
        ("D", "Values Only", lambda g, _t: g == 0),
    ]

    lines.append("## Library Tiers")
    lines.append("")
    lines.append(
        "> Libraries ranked by their best capability (max of read/write green features)."
    )
    lines.append("")
    lines.append("| Tier | Library | Caps | Green Features | Summary |")
    lines.append("|:----:|---------|:----:|:--------------:|---------|")

    for lib, best_green, total_scored, caps_label in lib_stats:
        tier_label = "D"
        for t_label, _, predicate in tier_defs:
            if predicate(best_green, total_scored):
                tier_label = t_label
                break
        summary = _lib_summary(lib, best_green, total_scored)
        lines.append(
            f"| **{tier_label}** | {lib} | {caps_label} | "
            f"{best_green}/{total_scored} | {summary} |"
        )

    lines.append("")
    return lines


def _lib_summary(lib: str, green: int, total: int) -> str:
    """One-line summary for tier list."""
    summaries: dict[str, str] = {
        "openpyxl": "Reference adapter — full read + write fidelity",
        "xlsxwriter": "Best write-only option — full formatting support",
        "xlsxwriter-constmem": "Memory-optimized write — loses images, comments, row height",
        "openpyxl-readonly": "Streaming read — loses all formatting metadata",
        "python-calamine": "Fast Rust-backed reader — cell values + sheet names only",
        "pylightxl": "Lightweight — basic values, no formatting API",
        "pyexcel": "Meta-library wrapping openpyxl — preserves error values",
        "pandas": "DataFrame abstraction — errors coerced to NaN on read",
        "polars": "Rust DataFrame reader — columnar type coercion drops fidelity",
        "tablib": "Dataset wrapper — matches pyexcel on fidelity",
        "xlrd": "Legacy .xls reader — not applicable to .xlsx",
        "xlwt": "Legacy .xls writer — basic formatting subset",
    }
    return summaries.get(lib, f"{green}/{total} features with full fidelity")


# ── T0-3: Deduplicated notes ──────────────────────────────────────────────


def _render_notes_deduped(results: BenchmarkResults) -> list[str]:
    """Render notes section with deduplication: group repeated note texts."""
    lines: list[str] = []

    # Collect (note_text -> set of features)
    note_features: dict[str, list[str]] = {}
    for score_entry in results.scores:
        if score_entry.notes:
            text = score_entry.notes
            note_features.setdefault(text, [])
            if score_entry.feature not in note_features[text]:
                note_features[text].append(score_entry.feature)

    if not note_features:
        return lines

    lines.append("## Notes")
    lines.append("")
    for text, feats in sorted(note_features.items(), key=lambda x: x[0]):
        if len(feats) <= 3:
            feat_str = ", ".join(feats)
        else:
            feat_str = f"{feats[0]}, {feats[1]}, ... ({len(feats)} features)"
        lines.append(f"- **{feat_str}**: {text}")
    lines.append("")
    return lines


# ── Statistics ─────────────────────────────────────────────────────────────


def _render_statistics(
    results: BenchmarkResults,
    libraries: list[str],
    features: list[str],
    score_lookup: dict[tuple[str, str], FeatureScore],
) -> list[str]:
    """Render pass-rate statistics section."""
    lines = ["## Statistics", ""]
    lines.append("| Library | Mode | Tests | Passed | Failed | Pass Rate | Green Features |")
    lines.append("|---------|------|-------|--------|--------|-----------|----------------|")

    for lib in libraries:
        caps = results.libraries[lib].capabilities
        for mode in ["read", "write"]:
            if mode not in caps:
                continue
            op = OperationType.READ if mode == "read" else OperationType.WRITE
            total = 0
            passed = 0
            green = 0
            total_features = 0

            for feature in features:
                score = score_lookup.get((feature, lib))
                if not score:
                    continue
                mode_score = score.read_score if mode == "read" else score.write_score
                if mode_score is None:
                    continue
                total_features += 1
                if mode_score == 3:
                    green += 1
                for tr in score.test_results:
                    if tr.operation == op:
                        total += 1
                        if tr.passed:
                            passed += 1

            if total == 0:
                continue
            failed = total - passed
            pct = f"{passed / total * 100:.0f}%" if total else "—"
            mode_label = "R" if mode == "read" else "W"
            lines.append(
                f"| {lib} | {mode_label} | {total} | {passed} | {failed} "
                f"| {pct} | {green}/{total_features} |"
            )
    lines.append("")
    return lines


# ── Per-test table ─────────────────────────────────────────────────────────


def _render_per_test_table(score: FeatureScore) -> list[str]:
    """Render per-test-case breakdown table for a feature/library with failures."""
    lines: list[str] = []

    # Collect unique test IDs preserving order
    test_ids: list[str] = []
    seen: set[str] = set()
    for tr in score.test_results:
        if tr.test_case_id not in seen:
            test_ids.append(tr.test_case_id)
            seen.add(tr.test_case_id)

    # Build lookup: (test_id, op) -> TestResult
    lookup: dict[tuple[str, str], TestResult] = {}
    for tr in score.test_results:
        lookup[(tr.test_case_id, tr.operation.value)] = tr

    has_read = any(tr.operation == OperationType.READ for tr in score.test_results)
    has_write = any(tr.operation == OperationType.WRITE for tr in score.test_results)

    header = "| Test | Importance |"
    sep = "|------|-----------|"
    if has_read:
        header += " Read |"
        sep += "------|"
    if has_write:
        header += " Write |"
        sep += "-------|"
    lines.append(header)
    lines.append(sep)

    for tid in test_ids:
        read_tr = lookup.get((tid, "read"))
        write_tr = lookup.get((tid, "write"))
        importance = ""
        label = tid
        if read_tr:
            importance = read_tr.importance.value if read_tr.importance else ""
            label = read_tr.label or tid
        elif write_tr:
            importance = write_tr.importance.value if write_tr.importance else ""
            label = write_tr.label or tid

        row = f"| {label} | {importance} |"
        if has_read:
            if read_tr:
                row += " ✅ |" if read_tr.passed else " ❌ |"
            else:
                row += " — |"
        if has_write:
            if write_tr:
                row += " ✅ |" if write_tr.passed else " ❌ |"
            else:
                row += " — |"
        lines.append(row)

    return lines


# ── CSV ────────────────────────────────────────────────────────────────────


def render_csv(results: BenchmarkResults, path: Path) -> None:
    """Render results to CSV."""
    lines = ["library,feature,read_score,write_score"]

    for score in results.scores:
        read = score.read_score if score.read_score is not None else ""
        write = score.write_score if score.write_score is not None else ""
        lines.append(f"{score.library},{score.feature},{read},{write}")

    with open(path, "w") as f:
        f.write("\n".join(lines))


# ── Utilities ──────────────────────────────────────────────────────────────


def score_emoji(score: int | None) -> str:
    """Convert score to emoji representation."""
    if score is None:
        return "➖"
    elif score == 3:
        return "🟢 3"
    elif score == 2:
        return "🟡 2"
    elif score == 1:
        return "🟠 1"
    else:
        return "🔴 0"


def _group_test_cases(test_results: list[TestResult]) -> dict[str, Any]:
    grouped: dict[str, Any] = {}
    for tr in test_results:
        entry = grouped.setdefault(tr.test_case_id, {})
        entry[tr.operation.value] = {
            "passed": tr.passed,
            "expected": tr.expected,
            "actual": tr.actual,
            "notes": tr.notes,
            "diagnostics": [_diagnostic_to_json(d) for d in tr.diagnostics],
            "importance": tr.importance.value if tr.importance else None,
            "label": tr.label,
        }
    return grouped


def _get_git_commit() -> str | None:
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except (FileNotFoundError, OSError, subprocess.SubprocessError):
        return None
    return None


def _append_history(results: BenchmarkResults, output_dir: Path) -> None:
    """Append a summary line to history.jsonl for tracking across runs."""
    history_path = output_dir / "history.jsonl"
    commit = _get_git_commit()

    by_lib: dict[str, dict[str, dict[str, int | None]]] = {}
    for s in results.scores:
        by_lib.setdefault(s.library, {})[s.feature] = {
            "read": s.read_score,
            "write": s.write_score,
        }

    entry = {
        "run_date": results.metadata.run_date.isoformat(),
        "commit": commit,
        "profile": results.metadata.profile,
        "scores": by_lib,
    }

    with open(history_path, "a") as f:
        f.write(json.dumps(entry) + "\n")


def _diagnostic_to_json(diagnostic: Diagnostic) -> dict[str, Any]:
    return {
        "category": diagnostic.category.value,
        "severity": diagnostic.severity.value,
        "location": {
            "feature": diagnostic.location.feature,
            "operation": diagnostic.location.operation.value,
            "test_case_id": diagnostic.location.test_case_id,
            "sheet": diagnostic.location.sheet,
            "cell": diagnostic.location.cell,
        },
        "adapter_message": diagnostic.adapter_message,
        "probable_cause": diagnostic.probable_cause,
    }


def _render_diagnostics_summary(results: BenchmarkResults) -> list[str]:
    diagnostics: list[Diagnostic] = []
    for score in results.scores:
        for tr in score.test_results:
            diagnostics.extend(tr.diagnostics)

    if not diagnostics:
        return ["No diagnostics recorded.", ""]

    by_category: dict[str, int] = {}
    by_severity: dict[str, int] = {}
    for item in diagnostics:
        by_category[item.category.value] = by_category.get(item.category.value, 0) + 1
        by_severity[item.severity.value] = by_severity.get(item.severity.value, 0) + 1

    lines = ["| Group | Value | Count |", "|-------|-------|-------|"]
    for key in sorted(by_category):
        lines.append(f"| category | {key} | {by_category[key]} |")
    for key in sorted(by_severity):
        lines.append(f"| severity | {key} | {by_severity[key]} |")
    lines.append("")
    lines.append("### Diagnostic Details")
    lines.append("")
    lines.append("| Feature | Library | Test Case | Operation | Category | Severity | Message |")
    lines.append("|---------|---------|-----------|-----------|----------|----------|---------|")
    for score in results.scores:
        for tr in score.test_results:
            for item in tr.diagnostics:
                lines.append(
                    "| "
                    f"{score.feature} | {score.library} | {tr.test_case_id} | "
                    f"{tr.operation.value} | {item.category.value} | "
                    f"{item.severity.value} | {item.adapter_message} |"
                )
    lines.append("")
    return lines
