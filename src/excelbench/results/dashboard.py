"""Combined fidelity + performance dashboard.

Merges data from fidelity results.json and perf/results.json into a single
overview table showing both quality and speed for each library.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def render_dashboard(
    fidelity_json: Path,
    perf_json: Path | None,
    output_path: Path,
) -> None:
    """Generate a combined fidelity × performance markdown dashboard.

    Args:
        fidelity_json: Path to fidelity results.json
        perf_json: Path to perf/results.json (optional — dashboard works without it)
        output_path: Path to write the markdown output
    """
    with open(fidelity_json) as f:
        fidelity_data = json.load(f)

    perf_data: dict[str, Any] | None = None
    if perf_json and perf_json.exists():
        with open(perf_json) as f:
            perf_data = json.load(f)

    lines = _build_dashboard(fidelity_data, perf_data)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w") as f:
        f.write("\n".join(lines))


def _build_dashboard(
    fidelity: dict[str, Any],
    perf: dict[str, Any] | None,
) -> list[str]:
    """Build combined dashboard markdown lines."""
    lines: list[str] = []

    profile = fidelity.get("metadata", {}).get("profile", "xlsx")
    generated = fidelity.get("metadata", {}).get("run_date", "unknown")

    lines.append("# ExcelBench Dashboard")
    lines.append("")
    lines.append(f"*Profile: {profile} | Generated: {generated}*")
    lines.append("")
    lines.append("> Combined fidelity and performance view. Fidelity shows correctness;")
    lines.append("> throughput shows speed. Use this to find the right library for your needs.")
    lines.append("")

    # ── Compute per-library fidelity stats ──
    lib_stats = _compute_fidelity_stats(fidelity)

    # ── Compute per-library throughput ──
    lib_throughput = _compute_throughput(perf) if perf else {}

    # ── Sort by best_green descending ──
    sorted_libs = sorted(lib_stats.keys(), key=lambda x: (-lib_stats[x]["best_green"], x))

    # ── Main table ──
    has_perf = bool(lib_throughput)

    lines.append("## Library Comparison")
    lines.append("")

    if has_perf:
        lines.append(
            "| Library | Caps | Green Features | Pass Rate | "
            "Read (cells/s) | Write (cells/s) | Best For |"
        )
        lines.append(
            "|---------|:----:|:--------------:|:---------:|"
            ":--------------:|:---------------:|----------|"
        )
    else:
        lines.append(
            "| Library | Caps | Green Features | Pass Rate | Best For |"
        )
        lines.append(
            "|---------|:----:|:--------------:|:---------:|----------|"
        )

    for lib in sorted_libs:
        stats = lib_stats[lib]
        throughput = lib_throughput.get(lib, {})

        green_str = f"{stats['best_green']}/{stats['total_scored']}"
        pass_str = f"{stats['pass_rate']:.0f}%"
        best_for = _best_for(lib)

        if has_perf:
            r_rate = _fmt_rate(throughput.get("read_rate"))
            w_rate = _fmt_rate(throughput.get("write_rate"))
            lines.append(
                f"| {lib} | {stats['caps']} | {green_str} | {pass_str} | "
                f"{r_rate} | {w_rate} | {best_for} |"
            )
        else:
            lines.append(
                f"| {lib} | {stats['caps']} | {green_str} | {pass_str} | {best_for} |"
            )

    lines.append("")

    # ── Key Insights ──
    lines.append("## Key Insights")
    lines.append("")
    lines.extend(_generate_insights(lib_stats, lib_throughput))
    lines.append("")

    lines.extend(_best_adapter_by_workload_profile(perf))

    return lines


def _compute_fidelity_stats(data: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Compute per-library fidelity statistics from results JSON."""
    libs_info = data.get("libraries", {})
    results = data.get("results", [])

    # Accumulate per (lib, mode)
    lib_mode_stats: dict[tuple[str, str], dict[str, int]] = {}
    for entry in results:
        lib = entry["library"]
        scores = entry.get("scores", {})
        test_cases = entry.get("test_cases", {})

        for mode in ("read", "write"):
            score = scores.get(mode)
            if score is None:
                continue
            key = (lib, mode)
            defaults = {"scored": 0, "green": 0, "total": 0, "passed": 0}
            stats = lib_mode_stats.setdefault(key, defaults)
            stats["scored"] += 1
            if score == 3:
                stats["green"] += 1
            for _tc_id, tc_data in test_cases.items():
                if isinstance(tc_data, dict) and mode in tc_data:
                    stats["total"] += 1
                    if tc_data[mode].get("passed"):
                        stats["passed"] += 1

    # Aggregate to per-library (best of R/W for green, sum for tests)
    out: dict[str, dict[str, Any]] = {}
    for lib in libs_info:
        caps = set(libs_info[lib].get("capabilities", []))
        has_rw = "read" in caps and "write" in caps
        caps_label = "R+W" if has_rw else ("R" if "read" in caps else "W")

        r_stats = lib_mode_stats.get((lib, "read"), {})
        w_stats = lib_mode_stats.get((lib, "write"), {})

        best_green = max(r_stats.get("green", 0), w_stats.get("green", 0))
        total_scored = max(r_stats.get("scored", 0), w_stats.get("scored", 0))
        total_tests = r_stats.get("total", 0) + w_stats.get("total", 0)
        total_passed = r_stats.get("passed", 0) + w_stats.get("passed", 0)
        pass_rate = (total_passed / total_tests * 100) if total_tests > 0 else 0

        if total_scored == 0:
            continue

        out[lib] = {
            "caps": caps_label,
            "best_green": best_green,
            "total_scored": total_scored,
            "pass_rate": pass_rate,
        }

    return out


def _compute_throughput(data: dict[str, Any]) -> dict[str, dict[str, float | None]]:
    """Extract representative throughput (cells/s) per library from perf results.

    Uses the cell_values_10k scenario as the representative benchmark, falling
    back to cell_values_1k if 10k is unavailable.
    """
    results = data.get("results", [])

    # Build lookup: (feature, lib) -> perf dict
    lookup: dict[tuple[str, str], dict[str, Any]] = {}
    for entry in results:
        lookup[(entry["feature"], entry["library"])] = entry.get("perf", {})

    libs = sorted(set(entry["library"] for entry in results))
    out: dict[str, dict[str, float | None]] = {}

    for lib in libs:
        read_rate: float | None = None
        write_rate: float | None = None

        # Prefer bulk scenarios (representative of streaming), then per-cell
        for scenario in ("cell_values_10k_bulk_read", "cell_values_1k_bulk_read",
                         "cell_values_10k", "cell_values_1k"):
            if read_rate is not None:
                break
            perf = lookup.get((scenario, lib), {})
            read_rate = _extract_rate(perf, "read")

        for scenario in ("cell_values_10k_bulk_write", "cell_values_1k_bulk_write",
                         "cell_values_10k", "cell_values_1k"):
            if write_rate is not None:
                break
            perf = lookup.get((scenario, lib), {})
            write_rate = _extract_rate(perf, "write")

        if read_rate is not None or write_rate is not None:
            out[lib] = {"read_rate": read_rate, "write_rate": write_rate}

    return out


def _extract_rate(perf: dict[str, Any], op: str) -> float | None:
    """Extract cells/s rate from a perf entry for a given operation."""
    if not perf or not isinstance(perf, dict):
        return None
    op_data = perf.get(op)
    if not op_data or not isinstance(op_data, dict):
        return None

    op_count = op_data.get("op_count")
    wall = op_data.get("wall_ms")
    if op_count is None or not isinstance(wall, dict):
        return None

    p50 = wall.get("p50")
    if p50 is None or float(p50) == 0:
        return None

    try:
        return float(op_count) * 1000.0 / float(p50)
    except (TypeError, ValueError, ZeroDivisionError):
        return None


def _fmt_rate(rate: float | None) -> str:
    """Format a cells/s rate for display."""
    if rate is None:
        return "—"
    if rate >= 1_000_000:
        return f"{rate / 1_000_000:.1f}M"
    if rate >= 1_000:
        return f"{rate / 1_000:.0f}K"
    return f"{rate:.0f}"


def _best_for(lib: str) -> str:
    """One-line use-case recommendation per library."""
    recs: dict[str, str] = {
        "openpyxl": "Full-fidelity read + write",
        "xlsxwriter": "High-fidelity write-only workflows",
        "xlsxwriter-constmem": "Large file writes with memory limits",
        "openpyxl-readonly": "Streaming reads when formatting isn't needed",
        "python-calamine": "Fast bulk value reads",
        "pylightxl": "Lightweight value extraction",
        "pyexcel": "Multi-format compatibility layer",
        "pandas": "Data analysis pipelines (accept NaN coercion)",
        "polars": "High-performance DataFrames (values only)",
        "tablib": "Dataset export/import workflows",
        "xlrd": "Legacy .xls file reads",
        "xlwt": "Legacy .xls file writes",
    }
    return recs.get(lib, "General use")



def _best_adapter_by_workload_profile(perf: dict[str, Any] | None) -> list[str]:
    if not perf:
        return []

    results = perf.get("results", [])
    by_size: dict[str, dict[str, tuple[str, float]]] = {
        "small": {},
        "medium": {},
        "large": {},
    }

    for entry in results:
        size = str(entry.get("workload_size") or "small").strip().lower()
        if size not in by_size:
            continue
        perf_entry = entry.get("perf")
        if not isinstance(perf_entry, dict):
            continue

        for op in ("read", "write"):
            rate = _extract_rate(perf_entry, op)
            if rate is None:
                continue
            cur = by_size[size].get(op)
            if cur is None or rate > cur[1]:
                by_size[size][op] = (str(entry.get("library")), rate)

    lines: list[str] = []
    lines.append("## Best Adapter by Workload Profile")
    lines.append("")
    lines.append("| Workload Size | Best Read Adapter | Best Write Adapter |")
    lines.append("|---------------|-------------------|--------------------|")

    has_any = False
    for size in ("small", "medium", "large"):
        read = by_size[size].get("read")
        write = by_size[size].get("write")
        if read or write:
            has_any = True
        read_label = f"{read[0]} ({_fmt_rate(read[1])} cells/s)" if read else "—"
        write_label = f"{write[0]} ({_fmt_rate(write[1])} cells/s)" if write else "—"
        lines.append(f"| {size} | {read_label} | {write_label} |")

    if not has_any:
        lines.append("| small | — | — |")
        lines.append("| medium | — | — |")
        lines.append("| large | — | — |")

    lines.append("")
    return lines

def _generate_insights(
    stats: dict[str, dict[str, Any]],
    throughput: dict[str, dict[str, float | None]],
) -> list[str]:
    """Generate key insights comparing fidelity vs performance."""
    lines: list[str] = []

    # Find fidelity leaders
    by_green = sorted(stats.items(), key=lambda x: -x[1]["best_green"])
    if not by_green:
        return lines
    leaders = [lib for lib, s in by_green if s["best_green"] == by_green[0][1]["best_green"]]
    lines.append(
        f"- **Fidelity leaders**: {', '.join(leaders)} "
        f"({by_green[0][1]['best_green']}/{by_green[0][1]['total_scored']} green features)"
    )

    # Find speed leaders (if perf data available)
    if throughput:
        read_rates = [(lib, t["read_rate"]) for lib, t in throughput.items()
                      if t.get("read_rate") is not None]
        write_rates = [(lib, t["write_rate"]) for lib, t in throughput.items()
                       if t.get("write_rate") is not None]

        if read_rates:
            fastest_reader = max(read_rates, key=lambda x: x[1])  # type: ignore[arg-type, return-value]
            lines.append(
                f"- **Fastest reader**: {fastest_reader[0]} "
                f"({_fmt_rate(fastest_reader[1])} cells/s on cell_values)"
            )

        if write_rates:
            fastest_writer = max(write_rates, key=lambda x: x[1])  # type: ignore[arg-type, return-value]
            lines.append(
                f"- **Fastest writer**: {fastest_writer[0]} "
                f"({_fmt_rate(fastest_writer[1])} cells/s on cell_values)"
            )

    # Abstraction cost callout
    if "pandas" in stats and "openpyxl" in stats:
        pandas_g = stats["pandas"]["best_green"]
        openpyxl_g = stats["openpyxl"]["best_green"]
        if openpyxl_g > pandas_g:
            lines.append(
                f"- **Abstraction cost**: pandas wraps openpyxl but drops from "
                f"{openpyxl_g} to {pandas_g} green features due to DataFrame coercion"
            )

    # Optimization tradeoff callout
    if "xlsxwriter" in stats and "xlsxwriter-constmem" in stats:
        full = stats["xlsxwriter"]["best_green"]
        constmem = stats["xlsxwriter-constmem"]["best_green"]
        if full > constmem:
            lines.append(
                f"- **Optimization cost**: xlsxwriter constant_memory mode loses "
                f"{full - constmem} green features for lower memory usage"
            )

    return lines
