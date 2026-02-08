"""Performance results rendering.

Writes into `<output_root>/perf/`:
- results.json
- README.md
- matrix.csv
- history.jsonl
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from excelbench.perf.runner import PerfResults, perf_results_to_json_dict

_FEATURE_TIERS: dict[str, int] = {
    "cell_values": 0,
    "formulas": 0,
    "multiple_sheets": 0,
    "text_formatting": 1,
    "background_colors": 1,
    "number_formats": 1,
    "alignment": 1,
    "borders": 1,
    "dimensions": 1,
    "merged_cells": 2,
    "conditional_formatting": 2,
    "data_validation": 2,
    "hyperlinks": 2,
    "images": 2,
    "pivot_tables": 2,
    "comments": 2,
    "freeze_panes": 2,
}

_TIER_LABELS = {
    0: "Tier 0 — Basic Values",
    1: "Tier 1 — Formatting",
    2: "Tier 2 — Advanced",
}


def render_perf_results(results: PerfResults, output_root: Path) -> None:
    output_root = Path(output_root)
    perf_dir = output_root / "perf"
    perf_dir.mkdir(parents=True, exist_ok=True)

    render_perf_json(results, perf_dir / "results.json")
    render_perf_markdown(results, perf_dir / "README.md")
    render_perf_csv(results, perf_dir / "matrix.csv")
    append_perf_history(results, perf_dir / "history.jsonl")


def render_perf_json(results: PerfResults, path: Path) -> None:
    with open(path, "w") as f:
        json.dump(perf_results_to_json_dict(results), f, indent=2)


def render_perf_markdown(results: PerfResults, path: Path) -> None:
    data = perf_results_to_json_dict(results)
    libs = sorted(data["libraries"].keys())
    features = sorted({r["feature"] for r in data["results"]})

    lookup: dict[tuple[str, str], dict[str, Any]] = {}
    for r in data["results"]:
        lookup[(r["feature"], r["library"])] = r

    lines: list[str] = []
    lines.append("# ExcelBench Performance Results")
    lines.append("")
    lines.append(f"*Generated: {data['metadata']['run_date']}*")
    lines.append(f"*Profile: {data['metadata']['profile']}*")
    lines.append(f"*Platform: {data['metadata']['platform']}*")
    lines.append(f"*Python: {data['metadata']['python']}*")
    if data["metadata"].get("commit"):
        lines.append(f"*Commit: {data['metadata']['commit']}*")
    cfg = data["metadata"].get("config", {})
    if cfg:
        lines.append(
            "*Config: "
            f"warmup={cfg.get('warmup')} "
            f"iters={cfg.get('iters')} "
            f"breakdown={cfg.get('breakdown')}*"
        )
    lines.append("")

    lines.append("## Notes")
    lines.append("")
    lines.append(
        "These numbers measure only the library under test. "
        "Write timings do NOT include oracle verification."
    )
    lines.append("")

    lines.append("## Summary (p50 wall time)")
    lines.append("")

    header = "| Feature |"
    sep = "|---------|"
    for lib in libs:
        caps = set(data["libraries"][lib].get("capabilities", []))
        if "read" in caps:
            header += f" {lib} (R p50 ms) |"
            sep += "--------------|"
        if "write" in caps:
            header += f" {lib} (W p50 ms) |"
            sep += "--------------|"

    tier_map: dict[int, list[str]] = {0: [], 1: [], 2: []}
    for feat in features:
        tier_map.setdefault(_FEATURE_TIERS.get(feat, 2), []).append(feat)

    for tier in sorted(tier_map.keys()):
        feats = tier_map[tier]
        if not feats:
            continue
        lines.append(f"**{_TIER_LABELS.get(tier, f'Tier {tier}')}**")
        lines.append("")

        lines.append(header)
        lines.append(sep)
        for feat in feats:
            row = f"| {feat} |"
            for lib in libs:
                caps = set(data["libraries"][lib].get("capabilities", []))
                entry = lookup.get((feat, lib))
                perf = entry.get("perf") if entry else None
                if "read" in caps:
                    row += f" {_fmt_p50_ms(perf, 'read')} |"
                if "write" in caps:
                    row += f" {_fmt_p50_ms(perf, 'write')} |"
            lines.append(row)
        lines.append("")

    _append_throughput_section(lines, data, libs, features, lookup)

    issues: list[str] = []
    for r in data["results"]:
        if r.get("notes"):
            issues.append(f"- {r['feature']} / {r['library']}: {r['notes']}")
    if issues:
        lines.append("## Run Issues")
        lines.append("")
        lines.extend(sorted(issues))
        lines.append("")

    with open(path, "w") as f:
        f.write("\n".join(lines))


def _append_throughput_section(
    lines: list[str],
    data: dict[str, Any],
    libs: list[str],
    features: list[str],
    lookup: dict[tuple[str, str], dict[str, Any]],
) -> None:
    # Identify workload-style features that include op_count.
    def _has_op_count(perf: dict[str, Any] | None, op: str) -> bool:
        if not perf or not isinstance(perf, dict):
            return False
        op_data = perf.get(op)
        return isinstance(op_data, dict) and op_data.get("op_count") is not None

    workload_features = []
    for feat in features:
        for lib in libs:
            entry = lookup.get((feat, lib))
            perf = entry.get("perf") if entry else None
            if _has_op_count(perf, "read") or _has_op_count(perf, "write"):
                workload_features.append(feat)
                break

    if not workload_features:
        return

    lines.append("## Throughput (derived from p50)")
    lines.append("")
    lines.append("Computed as: op_count * 1000 / p50_wall_ms")
    lines.append("")

    header = "| Scenario | op_count | op_unit |"
    sep = "|----------|----------|---------|"
    for lib in libs:
        caps = set(data["libraries"][lib].get("capabilities", []))
        if "read" in caps:
            header += f" {lib} (R units/s) |"
            sep += "----------------|"
        if "write" in caps:
            header += f" {lib} (W units/s) |"
            sep += "----------------|"

    lines.append(header)
    lines.append(sep)

    for feat in workload_features:
        base_count, base_unit = _feature_op_meta(libs, lookup, feat)
        row = f"| {feat} | {base_count if base_count is not None else '—'} | {base_unit or '—'} |"
        for lib in libs:
            caps = set(data["libraries"][lib].get("capabilities", []))
            entry = lookup.get((feat, lib))
            perf = entry.get("perf") if entry else None
            if "read" in caps:
                row += f" {_fmt_p50_units_per_sec(perf, 'read')} |"
            if "write" in caps:
                row += f" {_fmt_p50_units_per_sec(perf, 'write')} |"
        lines.append(row)

    lines.append("")


def _feature_op_meta(
    libs: list[str],
    lookup: dict[tuple[str, str], dict[str, Any]],
    feat: str,
) -> tuple[int | None, str | None]:
    for lib in libs:
        entry = lookup.get((feat, lib))
        perf = entry.get("perf") if entry else None
        if not perf or not isinstance(perf, dict):
            continue
        for op in ("read", "write"):
            op_data = perf.get(op)
            if not isinstance(op_data, dict):
                continue
            count = op_data.get("op_count")
            unit = op_data.get("op_unit")
            if count is None:
                continue
            try:
                count_i = int(count)
            except (TypeError, ValueError):
                continue
            return count_i, str(unit) if unit is not None else None
    return None, None


def _fmt_p50_units_per_sec(perf: dict[str, Any] | None, op: str) -> str:
    if not perf or not isinstance(perf, dict):
        return "—"
    op_data = perf.get(op)
    if not op_data or not isinstance(op_data, dict):
        return "—"
    op_count = op_data.get("op_count")
    wall = op_data.get("wall_ms")
    if op_count is None or not isinstance(wall, dict):
        return "—"
    p50 = wall.get("p50")
    if p50 in (None, 0):
        return "—"
    try:
        rate = float(op_count) * 1000.0 / float(p50)
    except (TypeError, ValueError, ZeroDivisionError):
        return "—"
    return _fmt_rate(rate)


def _fmt_rate(rate: float) -> str:
    if rate >= 1_000_000:
        return f"{rate / 1_000_000.0:.2f}M"
    if rate >= 1_000:
        return f"{rate / 1_000.0:.2f}K"
    return f"{rate:.2f}"


def _fmt_p50_ms(perf: dict[str, Any] | None, op: str) -> str:
    if not perf or not isinstance(perf, dict):
        return "—"
    op_data = perf.get(op)
    if not op_data or not isinstance(op_data, dict):
        return "—"
    wall = op_data.get("wall_ms")
    if not wall or not isinstance(wall, dict):
        return "—"
    p50 = wall.get("p50")
    if p50 is None:
        return "—"
    try:
        return f"{float(p50):.2f}"
    except (TypeError, ValueError):
        return "—"


def render_perf_csv(results: PerfResults, path: Path) -> None:
    data = perf_results_to_json_dict(results)
    lines = [
        "library,feature,read_p50_wall_ms,read_p95_wall_ms,read_op_count,read_op_unit,read_p50_units_per_sec,"
        "write_p50_wall_ms,write_p95_wall_ms,write_op_count,write_op_unit,write_p50_units_per_sec",
    ]
    for r in data["results"]:
        perf = r.get("perf") or {}
        read = perf.get("read") or {}
        write = perf.get("write") or {}
        read_wall = (read.get("wall_ms") or {}) if isinstance(read, dict) else {}
        write_wall = (write.get("wall_ms") or {}) if isinstance(write, dict) else {}

        read_count = read.get("op_count") if isinstance(read, dict) else None
        read_unit = read.get("op_unit") if isinstance(read, dict) else None
        write_count = write.get("op_count") if isinstance(write, dict) else None
        write_unit = write.get("op_unit") if isinstance(write, dict) else None

        def _rate(count: Any, p50_ms: Any) -> str:
            try:
                if count is None or p50_ms in (None, 0):
                    return ""
                return str(float(count) * 1000.0 / float(p50_ms))
            except (TypeError, ValueError, ZeroDivisionError):
                return ""

        def _f(v: Any) -> str:
            return "" if v is None else str(v)

        lines.append(
            ",".join(
                [
                    str(r["library"]),
                    str(r["feature"]),
                    _f(read_wall.get("p50")),
                    _f(read_wall.get("p95")),
                    _f(read_count),
                    _f(read_unit),
                    _rate(read_count, read_wall.get("p50")),
                    _f(write_wall.get("p50")),
                    _f(write_wall.get("p95")),
                    _f(write_count),
                    _f(write_unit),
                    _rate(write_count, write_wall.get("p50")),
                ]
            )
        )

    with open(path, "w") as f:
        f.write("\n".join(lines))


def append_perf_history(results: PerfResults, history_path: Path) -> None:
    data = perf_results_to_json_dict(results)

    by_lib: dict[str, dict[str, dict[str, float | None]]] = {}
    for r in data["results"]:
        perf = r.get("perf") or {}
        read_p50 = None
        write_p50 = None
        if isinstance(perf.get("read"), dict):
            read_p50 = (perf["read"].get("wall_ms") or {}).get("p50")
        if isinstance(perf.get("write"), dict):
            write_p50 = (perf["write"].get("wall_ms") or {}).get("p50")
        by_lib.setdefault(r["library"], {})[r["feature"]] = {
            "read_p50": read_p50,
            "write_p50": write_p50,
        }

    entry = {
        "run_date": data["metadata"]["run_date"],
        "commit": data["metadata"].get("commit"),
        "profile": data["metadata"].get("profile"),
        "config": data["metadata"].get("config"),
        "p50_wall_ms": by_lib,
    }

    with open(history_path, "a") as f:
        f.write(json.dumps(entry) + "\n")
