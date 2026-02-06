"""Results rendering to various output formats."""

import json
from pathlib import Path

from excelbench.models import BenchmarkResults, FeatureScore


def render_results(results: BenchmarkResults, output_dir: Path) -> None:
    """Render results to all output formats.

    Args:
        results: The benchmark results.
        output_dir: Directory to write output files.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Write JSON
    render_json(results, output_dir / "results.json")

    # Write markdown summary
    render_markdown(results, output_dir / "README.md")

    # Write CSV
    render_csv(results, output_dir / "matrix.csv")


def render_json(results: BenchmarkResults, path: Path) -> None:
    """Render results to JSON.

    Args:
        results: The benchmark results.
        path: Path to write JSON file.
    """
    data = {
        "metadata": {
            "benchmark_version": results.metadata.benchmark_version,
            "run_date": results.metadata.run_date.isoformat(),
            "excel_version": results.metadata.excel_version,
            "platform": results.metadata.platform,
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
    """Render results to markdown summary.

    Args:
        results: The benchmark results.
        path: Path to write markdown file.
    """
    lines = []

    # Header
    lines.append("# ExcelBench Results")
    lines.append("")
    lines.append(f"*Generated: {results.metadata.run_date.strftime('%Y-%m-%d %H:%M UTC')}*")
    lines.append(f"*Excel Version: {results.metadata.excel_version}*")
    lines.append(f"*Platform: {results.metadata.platform}*")
    lines.append("")

    # Legend
    lines.append("## Score Legend")
    lines.append("")
    lines.append("| Score | Meaning |")
    lines.append("|-------|---------|")
    lines.append("| 🟢 3 | Complete - full fidelity |")
    lines.append("| 🟡 2 | Functional - works for common cases |")
    lines.append("| 🟠 1 | Minimal - basic recognition only |")
    lines.append("| 🔴 0 | Unsupported - errors or data loss |")
    lines.append("| ➖ | Not applicable (library doesn't support this operation) |")
    lines.append("")

    # Summary table
    lines.append("## Summary")
    lines.append("")

    # Build feature x library matrix
    features = sorted(set(s.feature for s in results.scores))
    libraries = sorted(results.libraries.keys())

    # Create header
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

    lines.append(header)
    lines.append(separator)

    # Build lookup for scores
    score_lookup: dict[tuple[str, str], FeatureScore] = {}
    for score in results.scores:
        score_lookup[(score.feature, score.library)] = score

    # Add rows
    for feature in features:
        row = f"| {feature} |"
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

    notes = []
    for score in results.scores:
        if score.notes:
            notes.append(f"- {score.feature}: {score.notes}")
    if notes:
        lines.append("Notes:")
        lines.extend(sorted(notes))
        lines.append("")

    # Library details
    lines.append("## Libraries Tested")
    lines.append("")
    for name, info in sorted(results.libraries.items()):
        caps_str = ", ".join(sorted(info.capabilities))
        lines.append(f"- **{name}** v{info.version} ({info.language}) - {caps_str}")
    lines.append("")

    # Detailed results per feature
    lines.append("## Detailed Results")
    lines.append("")

    for feature in features:
        lines.append(f"### {feature}")
        lines.append("")

        for lib in libraries:
            score = score_lookup.get((feature, lib))
            if not score:
                continue

            lines.append(f"**{lib}**")

            if score.read_score is not None:
                lines.append(f"- Read: {score_emoji(score.read_score)} ({score.read_score}/3)")
            if score.write_score is not None:
                lines.append(f"- Write: {score_emoji(score.write_score)} ({score.write_score}/3)")
            if score.notes:
                lines.append(f"- Notes: {score.notes}")

            # Show failed tests
            failed = [tr for tr in score.test_results if not tr.passed]
            if failed:
                lines.append(f"- Failed tests ({len(failed)}):")
                for tr in failed[:5]:  # Show max 5
                    lines.append(f"  - {tr.test_case_id} ({tr.operation.value})")
                if len(failed) > 5:
                    lines.append(f"  - ... and {len(failed) - 5} more")

            lines.append("")

    # Footer
    lines.append("---")
    lines.append(f"*Benchmark version: {results.metadata.benchmark_version}*")

    with open(path, "w") as f:
        f.write("\n".join(lines))


def render_csv(results: BenchmarkResults, path: Path) -> None:
    """Render results to CSV.

    Args:
        results: The benchmark results.
        path: Path to write CSV file.
    """
    lines = ["library,feature,read_score,write_score"]

    for score in results.scores:
        read = score.read_score if score.read_score is not None else ""
        write = score.write_score if score.write_score is not None else ""
        lines.append(f"{score.library},{score.feature},{read},{write}")

    with open(path, "w") as f:
        f.write("\n".join(lines))


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


def _group_test_cases(test_results: list) -> dict:
    grouped: dict[str, dict] = {}
    for tr in test_results:
        entry = grouped.setdefault(tr.test_case_id, {})
        entry[tr.operation.value] = {
            "passed": tr.passed,
            "expected": tr.expected,
            "actual": tr.actual,
            "notes": tr.notes,
        }
    return grouped
