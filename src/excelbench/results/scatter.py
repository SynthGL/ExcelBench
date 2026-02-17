"""Scatter-plot visualizations: fidelity pass-rate vs. throughput.

Each point represents a library.  The Y-axis shows the *test pass rate*
(%), which is continuous and avoids the bunching that discrete 0-3
scores would cause.  Horizontal zone-bands at 50 %/80 %/100 % indicate
the corresponding fidelity-score thresholds.  The X-axis shows
throughput in cells/s on a log scale.

Two figures are produced:
  scatter_tiers   — 1 × 3 grid  (Tier 0, Tier 1, Overall)
  scatter_features — 2 × 3 grid (6 features with perf data)
"""

from __future__ import annotations

import hashlib
import json
import math
from pathlib import Path
from typing import Any

import matplotlib

matplotlib.use("Agg")  # headless-safe, before pyplot import

import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
from matplotlib.axes import Axes  # noqa: E402
from matplotlib.figure import Figure  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402

# ── Feature ↔ perf-workload mapping ────────────────────────────────
# Each fidelity feature maps to a list of perf workloads (prefer first).

_FEATURE_PERF_MAP: dict[str, list[str]] = {
    "cell_values": ["cell_values_10k", "cell_values_1k"],
    "formulas": ["formulas_10k", "formulas_1k"],
    "alignment": ["alignment_1k"],
    "background_colors": ["background_colors_1k"],
    "borders": ["borders_200"],
    "number_formats": ["number_formats_1k"],
}

# ── Tier groupings ─────────────────────────────────────────────────

_TIER_GROUPS: list[tuple[str, list[str]]] = [
    (
        "Tier 0 — Core",
        ["cell_values", "formulas", "multiple_sheets"],
    ),
    (
        "Tier 1 — Formatting",
        [
            "alignment",
            "background_colors",
            "borders",
            "dimensions",
            "number_formats",
            "text_formatting",
        ],
    ),
]

_FEATURE_LABELS: dict[str, str] = {
    "cell_values": "Cell Values",
    "formulas": "Formulas",
    "alignment": "Alignment",
    "background_colors": "Background Colors",
    "borders": "Borders",
    "number_formats": "Number Formats",
}

# ── Consistent library colour palette ──────────────────────────────

_LIB_COLORS: dict[str, str] = {
    "wolfxl": "#f97316",
    "openpyxl": "#2563eb",
    "xlsxwriter": "#7c3aed",
    "python-calamine": "#16a34a",
    "pylightxl": "#9333ea",
    "xlrd": "#78716c",
    "pyexcel": "#475569",
    "pandas": "#db2777",
    "polars": "#0891b2",
    "tablib": "#ca8a04",
    "xlsxwriter-constmem": "#c2410c",
    "openpyxl-readonly": "#4f46e5",
    "rust_xlsxwriter": "#dc2626",
    "fastexcel": "#65a30d",
}
# WolfXL gets a larger marker to stand out
_WOLFXL_MARKER_SCALE = 1.8
_FALLBACK_COLOR = "#6b7280"

# Short display names to reduce label clutter
_LIB_SHORT: dict[str, str] = {
    "python-calamine": "calamine",
    "xlsxwriter-constmem": "xlsx-constmem",
    "openpyxl-readonly": "opx-readonly",
}

# ── Capability markers ─────────────────────────────────────────────

_CAP_MARKERS: dict[str, str] = {"R+W": "o", "R": "s", "W": "D"}
_CAP_LABELS: dict[str, str] = {
    "R+W": "Read + Write",
    "R": "Read-only",
    "W": "Write-only",
}

# ── Score-zone band styling ────────────────────────────────────────

_ZONE_BANDS: list[tuple[float, float, str, str]] = [
    (0, 50, "#fef2f2", "0"),  # red-50
    (50, 80, "#fffbeb", "1"),  # amber-50
    (80, 100, "#f0fdf4", "2–3"),  # green-50
]


# ====================================================================
#  Data extraction
# ====================================================================


def _load_json(path: Path) -> dict[str, Any]:
    with open(path) as f:
        result: dict[str, Any] = json.load(f)
        return result


def _compute_pass_rates(data: dict[str, Any]) -> dict[str, dict[str, float]]:
    """Per-library, per-feature pass rate (%) using best-of(read, write).

    Returns ``{lib: {feature: rate_pct}}``.
    """
    raw: dict[str, dict[str, dict[str, list[bool]]]] = {}

    for entry in data.get("results", []):
        lib = entry["library"]
        feat = entry["feature"]
        for _, tc in entry.get("test_cases", {}).items():
            if not isinstance(tc, dict):
                continue
            for op in ("read", "write"):
                if op not in tc:
                    continue
                raw.setdefault(lib, {}).setdefault(feat, {}).setdefault(op, [])
                raw[lib][feat][op].append(tc[op].get("passed", False))

    rates: dict[str, dict[str, float]] = {}
    for lib, feats in raw.items():
        rates[lib] = {}
        for feat, ops in feats.items():
            best = 0.0
            for _, results in ops.items():
                if results:
                    best = max(best, sum(results) / len(results) * 100)
            rates[lib][feat] = best
    return rates


def _compute_throughputs(data: dict[str, Any]) -> dict[str, dict[str, float]]:
    """Per-library, per-*fidelity*-feature throughput (ops/s).

    Tries scale-suffixed workload names first (legacy), then falls back
    to direct feature names (current format).
    Returns ``{lib: {fidelity_feature: ops_per_sec}}``.
    """
    lookup: dict[tuple[str, str], dict[str, Any]] = {}
    for entry in data.get("results", []):
        lookup[(entry["feature"], entry["library"])] = entry.get("perf", {})

    all_libs = sorted({e["library"] for e in data.get("results", [])})
    out: dict[str, dict[str, float]] = {}

    for lib in all_libs:
        for fidelity_feat, workloads in _FEATURE_PERF_MAP.items():
            # Try scale-suffixed names first, then the plain feature name
            candidates = [*workloads, fidelity_feat]
            for candidate in candidates:
                perf = lookup.get((candidate, lib), {})
                rate = _best_throughput(perf)
                if rate is not None:
                    out.setdefault(lib, {})[fidelity_feat] = rate
                    break
    return out


def _best_throughput(perf: dict[str, Any]) -> float | None:
    """Best-of(read, write) throughput from a single perf entry.

    When ``op_count`` is set, returns cells/s.
    When ``op_count`` is None (fidelity-feature benchmarks), returns
    ops/s = 1000 / wall_ms as a comparable speed metric.
    """
    best: float | None = None
    for op in ("read", "write"):
        op_data = perf.get(op)
        if not op_data or not isinstance(op_data, dict):
            continue
        wall = op_data.get("wall_ms")
        if not isinstance(wall, dict):
            continue
        p50 = wall.get("p50")
        if p50 is None or float(p50) == 0:
            continue
        op_count = op_data.get("op_count")
        if op_count is not None:
            rate = float(op_count) * 1000.0 / float(p50)
        else:
            # No explicit op_count — use ops/s (one op = full feature exercise)
            rate = 1000.0 / float(p50)
        if best is None or rate > best:
            best = rate
    return best


def _compute_representative_throughput(data: dict[str, Any]) -> dict[str, float]:
    """Single representative throughput per library (for the Overall panel).

    Prefers bulk scenarios, then per-cell, then plain feature name.
    """
    lookup: dict[tuple[str, str], dict[str, Any]] = {}
    for entry in data.get("results", []):
        lookup[(entry["feature"], entry["library"])] = entry.get("perf", {})

    all_libs = sorted({e["library"] for e in data.get("results", [])})
    out: dict[str, float] = {}
    scenarios = [
        "cell_values_10k_bulk_read",
        "cell_values_10k_bulk_write",
        "cell_values_10k",
        "cell_values_1k",
        "cell_values",
    ]
    for lib in all_libs:
        for scenario in scenarios:
            perf = lookup.get((scenario, lib), {})
            rate = _best_throughput(perf)
            if rate is not None:
                out[lib] = rate
                break
    return out


def _compute_capabilities(data: dict[str, Any]) -> dict[str, str]:
    """Capability label per library: ``R+W``, ``R``, or ``W``."""
    out: dict[str, str] = {}
    for lib, info in data.get("libraries", {}).items():
        caps = set(info.get("capabilities", []))
        if "read" in caps and "write" in caps:
            out[lib] = "R+W"
        elif "read" in caps:
            out[lib] = "R"
        else:
            out[lib] = "W"
    return out


# ====================================================================
#  Deterministic jitter
# ====================================================================


def _jitter(name: str, scale: float = 1.5) -> float:
    """Small deterministic vertical offset seeded by *name*."""
    h = int(hashlib.sha256(name.encode()).hexdigest()[:8], 16)
    return ((h % 1000) / 500 - 1.0) * scale


# ====================================================================
#  Point aggregation helpers
# ====================================================================

# Each "point" is (lib_name, pass_rate_pct, throughput_cells_s, cap_label).
Point = tuple[str, float, float, str]


def _tier_points(
    tier_features: list[str],
    pass_rates: dict[str, dict[str, float]],
    throughputs: dict[str, dict[str, float]],
    caps: dict[str, str],
) -> list[Point]:
    """Average pass rate and throughput across the features in one tier."""
    points: list[Point] = []
    all_libs = set(pass_rates) | set(throughputs)

    for lib in sorted(all_libs):
        lib_rates = pass_rates.get(lib, {})
        rates = [lib_rates[f] for f in tier_features if f in lib_rates]
        if not rates:
            continue

        lib_tps = throughputs.get(lib, {})
        tps = [lib_tps[f] for f in tier_features if f in lib_tps]
        if not tps:
            continue

        points.append((
            lib,
            sum(rates) / len(rates),
            sum(tps) / len(tps),
            caps.get(lib, "R+W"),
        ))
    return points


def _overall_points(
    pass_rates: dict[str, dict[str, float]],
    rep_throughputs: dict[str, float],
    caps: dict[str, str],
) -> list[Point]:
    """Average pass rate across *all* features, representative throughput."""
    points: list[Point] = []
    for lib in sorted(pass_rates):
        rates = list(pass_rates[lib].values())
        if not rates:
            continue
        tp = rep_throughputs.get(lib)
        if tp is None:
            continue
        points.append((lib, sum(rates) / len(rates), tp, caps.get(lib, "R+W")))
    return points


def _feature_points(
    feature: str,
    pass_rates: dict[str, dict[str, float]],
    throughputs: dict[str, dict[str, float]],
    caps: dict[str, str],
) -> list[Point]:
    """Points for a single feature."""
    points: list[Point] = []
    for lib in sorted(pass_rates):
        rate = pass_rates.get(lib, {}).get(feature)
        tp = throughputs.get(lib, {}).get(feature)
        if rate is None or tp is None:
            continue
        points.append((lib, rate, tp, caps.get(lib, "R+W")))
    return points


# ====================================================================
#  Panel drawing
# ====================================================================


def _draw_panel(ax: Axes, title: str, points: list[Point]) -> None:
    """Render a single scatter panel with zone bands and labelled points."""

    # ── Score-zone background bands ──
    for y_lo, y_hi, colour, _ in _ZONE_BANDS:
        ax.axhspan(y_lo, y_hi, color=colour, alpha=0.7, zorder=0)

    # Threshold lines
    for y_val, ls, c, lw in [
        (50, "--", "#d4d4d8", 0.8),
        (80, "--", "#d4d4d8", 0.8),
        (100, "-", "#86efac", 1.0),
    ]:
        ax.axhline(y_val, color=c, linewidth=lw, linestyle=ls, zorder=1)

    # Zone score annotations on the right edge (very subtle, behind data)
    for y_pos, label, colour in [
        (25, "Score 0", "#ef4444"),
        (65, "Score 1", "#f59e0b"),
        (90, "Score 2", "#22c55e"),
        (102, "Score 3", "#15803d"),
    ]:
        ax.text(
            0.99, y_pos, label,
            transform=ax.get_yaxis_transform(),
            ha="right", va="center", fontsize=5.5, color=colour, alpha=0.3,
            fontstyle="italic", zorder=2,
        )

    # ── Empty-state fallback ──
    if not points:
        ax.text(
            0.5, 0.5, "No data",
            transform=ax.transAxes, ha="center", va="center",
            fontsize=11, color="#9ca3af",
        )
        _style_axes(ax, title)
        return

    # ── Sort points by throughput for consistent label staggering ──
    points_sorted = sorted(points, key=lambda p: p[2])

    # ── Scatter + labels ──
    # Two-pass: draw all markers first, then labels (so labels sit on top)
    jittered: list[tuple[str, float, float, float, str]] = []
    for lib, rate, tp, cap in points_sorted:
        color = _LIB_COLORS.get(lib, _FALLBACK_COLOR)
        marker = _CAP_MARKERS.get(cap, "o")
        y = rate + _jitter(lib, scale=3.0)
        y = float(np.clip(y, -2, 107))

        is_wolfxl = lib == "wolfxl"
        sz = int(80 * _WOLFXL_MARKER_SCALE ** 2) if is_wolfxl else 80
        ec = "#fb923c" if is_wolfxl else _DARK_CARD
        lw = 2.0 if is_wolfxl else 0.8
        zord = 10 if is_wolfxl else 5
        ax.scatter(
            tp, y,
            c=color, marker=marker, s=sz,
            edgecolors=ec, linewidths=lw,
            zorder=zord, alpha=0.95 if is_wolfxl else 0.92,
        )
        jittered.append((lib, y, tp, rate, cap))

    # Place labels with alternating directions to fan out from clusters
    tp_values = [tp for _, _, tp, _, _ in jittered]
    tp_min = min(tp_values) if tp_values else 1
    tp_max = max(tp_values) if tp_values else 2
    # Points in the leftmost 15 % of log-range are forced right (avoid axis clip)
    log_span = math.log10(tp_max / tp_min) if tp_max > tp_min else 1
    left_thresh = tp_min * 10 ** (0.15 * log_span)

    label_ys: list[float] = []
    for i, (lib, y, tp, _rate, cap) in enumerate(jittered):
        color = _LIB_COLORS.get(lib, _FALLBACK_COLOR)
        display_name = _LIB_SHORT.get(lib, lib)

        # Alternate left/right; force right when near left axis edge
        near_left_edge = tp <= left_thresh
        place_right = near_left_edge or (i % 2 == 0)
        x_off = 9 if place_right else -9
        ha = "left" if place_right else "right"
        start_dir = 1 if i % 2 == 0 else -1

        y_off = _stagger_offset(y, label_ys, step=10, start_direction=start_dir)
        label_ys.append(y + y_off)

        use_arrow = abs(y_off) > 12
        is_wolf = lib == "wolfxl"
        ax.annotate(
            display_name, (tp, y),
            xytext=(x_off, y_off), textcoords="offset points",
            fontsize=7.5 if is_wolf else 6,
            color=color,
            fontweight="bold" if is_wolf else "medium",
            ha=ha, va="center",
            bbox=dict(
                boxstyle="round,pad=0.2" if is_wolf else "round,pad=0.15",
                fc="#431407" if is_wolf else _DARK_CARD,
                ec="#f97316" if is_wolf else "none",
                alpha=0.95 if is_wolf else 0.85,
                lw=1.0 if is_wolf else 0,
            ),
            arrowprops=(
                dict(arrowstyle="-", color=color, alpha=0.35, lw=0.6)
                if use_arrow else None
            ),
            zorder=11 if is_wolf else 10,
        )

    _style_axes(ax, title)


def _stagger_offset(
    y: float,
    placed: list[float],
    step: float = 10,
    threshold: float = 8,
    start_direction: int = 1,
) -> float:
    """Pick a vertical label offset that avoids collisions with *placed*.

    *start_direction* alternates between +1 and −1 so adjacent labels
    fan out in opposite vertical directions, preventing one-sided pileup.
    """
    other = -start_direction
    for mult in range(8):
        for sign in (start_direction, other):
            candidate = sign * mult * step
            target = y + candidate
            if all(abs(target - py) >= threshold for py in placed):
                return candidate
    return 0.0


_DARK_BG = "#0b0f19"
_DARK_CARD = "#111827"
_DARK_TEXT = "#e2e8f0"
_DARK_TEXT2 = "#94a3b8"
_DARK_GRID = "#1e293b"


def _style_axes(ax: Axes, title: str) -> None:
    """Apply shared axis styling (dark theme)."""
    ax.set_facecolor(_DARK_CARD)
    ax.set_xscale("log")
    ax.set_ylim(-5, 110)
    ax.set_ylabel("Pass Rate (%)", fontsize=9, labelpad=6, color=_DARK_TEXT)
    ax.set_xlabel("Throughput (ops / s)", fontsize=9, labelpad=6, color=_DARK_TEXT)
    ax.set_title(title, fontsize=11, fontweight="bold", pad=10, color=_DARK_TEXT)
    ax.tick_params(labelsize=8, colors=_DARK_TEXT2)
    ax.grid(True, axis="x", alpha=0.15, color=_DARK_GRID, zorder=0)
    for spine in ax.spines.values():
        spine.set_color(_DARK_GRID)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


# ====================================================================
#  Shared legend
# ====================================================================


def _add_capability_legend(fig: Figure) -> None:
    """Add a small marker-shape legend for capabilities at the bottom."""
    handles = [
        Line2D(
            [0], [0], marker=m, color="none",
            markerfacecolor="#94a3b8", markeredgecolor=_DARK_CARD,
            markersize=8, label=_CAP_LABELS[cap],
        )
        for cap, m in _CAP_MARKERS.items()
    ]
    fig.legend(
        handles=handles, loc="lower center", ncol=3,
        fontsize=8, frameon=False, bbox_to_anchor=(0.5, 0.005),
        labelcolor=_DARK_TEXT2,
    )


# ====================================================================
#  Public render functions
# ====================================================================


def render_scatter_tiers(
    fidelity_json: Path,
    perf_json: Path,
    output_dir: Path,
) -> list[Path]:
    """Generate a 1 × 3 scatter grid grouped by feature tier.

    Panels: [Tier 0 — Core]  [Tier 1 — Formatting]  [Overall]

    Returns list of generated file paths (PNG + SVG).
    """
    fidelity = _load_json(fidelity_json)
    perf = _load_json(perf_json)

    pass_rates = _compute_pass_rates(fidelity)
    throughputs = _compute_throughputs(perf)
    rep_tp = _compute_representative_throughput(perf)
    caps = _compute_capabilities(fidelity)

    output_dir.mkdir(parents=True, exist_ok=True)

    fig, axes = plt.subplots(1, 3, figsize=(19, 7.5))
    fig.set_facecolor(_DARK_BG)
    fig.suptitle(
        "ExcelBench — Fidelity vs. Throughput by Feature Group",
        fontsize=14, fontweight="bold", y=0.98, color=_DARK_TEXT,
    )

    # Tier panels
    for ax, (tier_name, tier_features) in zip(axes[:2], _TIER_GROUPS):
        pts = _tier_points(tier_features, pass_rates, throughputs, caps)
        _draw_panel(ax, tier_name, pts)

    # Overall panel
    overall = _overall_points(pass_rates, rep_tp, caps)
    _draw_panel(axes[2], "Overall", overall)

    _add_capability_legend(fig)

    plt.tight_layout(rect=(0, 0.06, 1, 0.93))

    # Jitter footnote (after tight_layout so positioning is stable)
    fig.text(
        0.5, 0.005,
        "Y-positions include \u00b13 % jitter for visibility  \u00b7  marker shape = capability",
        ha="center", fontsize=7, color=_DARK_TEXT2, fontstyle="italic",
    )

    paths: list[Path] = []
    for ext in ("png", "svg"):
        path = output_dir / f"scatter_tiers.{ext}"
        dpi = 200 if ext == "png" else 150
        fig.savefig(path, dpi=dpi, bbox_inches="tight", facecolor=_DARK_BG)
        paths.append(path)
    plt.close(fig)
    return paths


def render_scatter_features(
    fidelity_json: Path,
    perf_json: Path,
    output_dir: Path,
) -> list[Path]:
    """Generate a 2 × 3 per-feature scatter grid.

    One panel per feature that has matching perf data:
    cell_values, formulas, alignment, background_colors, borders,
    number_formats.

    Returns list of generated file paths (PNG + SVG).
    """
    fidelity = _load_json(fidelity_json)
    perf = _load_json(perf_json)

    pass_rates = _compute_pass_rates(fidelity)
    throughputs = _compute_throughputs(perf)
    caps = _compute_capabilities(fidelity)

    features = list(_FEATURE_PERF_MAP.keys())

    output_dir.mkdir(parents=True, exist_ok=True)

    fig, axes = plt.subplots(2, 3, figsize=(19, 13))
    fig.set_facecolor(_DARK_BG)
    fig.suptitle(
        "ExcelBench — Per-Feature Fidelity vs. Throughput",
        fontsize=14, fontweight="bold", y=0.98, color=_DARK_TEXT,
    )

    for ax, feature in zip(axes.flat, features):
        label = _FEATURE_LABELS.get(feature, feature)
        pts = _feature_points(feature, pass_rates, throughputs, caps)
        _draw_panel(ax, label, pts)

    _add_capability_legend(fig)

    plt.tight_layout(rect=(0, 0.05, 1, 0.93))

    fig.text(
        0.5, 0.005,
        "Y-positions include \u00b13 % jitter for visibility  \u00b7  marker shape = capability",
        ha="center", fontsize=7, color=_DARK_TEXT2, fontstyle="italic",
    )

    paths: list[Path] = []
    for ext in ("png", "svg"):
        path = output_dir / f"scatter_features.{ext}"
        dpi = 200 if ext == "png" else 150
        fig.savefig(path, dpi=dpi, bbox_inches="tight", facecolor=_DARK_BG)
        paths.append(path)
    plt.close(fig)
    return paths
