"""Generate heatmap visualizations of the fidelity score matrix."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import ListedColormap

# Feature display order (tier-grouped)
_FEATURE_ORDER: list[str] = [
    # Tier 0
    "cell_values",
    "formulas",
    "multiple_sheets",
    # Tier 1
    "alignment",
    "background_colors",
    "borders",
    "dimensions",
    "number_formats",
    "text_formatting",
    # Tier 2
    "comments",
    "conditional_formatting",
    "data_validation",
    "freeze_panes",
    "hyperlinks",
    "images",
    "merged_cells",
    # Tier 3
    "named_ranges",
    "tables",
]

_FEATURE_LABELS: dict[str, str] = {
    "cell_values": "Cell Values",
    "formulas": "Formulas",
    "multiple_sheets": "Sheets",
    "text_formatting": "Text Fmt",
    "background_colors": "Bg Colors",
    "number_formats": "Num Fmt",
    "alignment": "Alignment",
    "borders": "Borders",
    "dimensions": "Dimensions",
    "merged_cells": "Merged Cells",
    "conditional_formatting": "Cond. Fmt",
    "data_validation": "Validation",
    "hyperlinks": "Hyperlinks",
    "images": "Images",
    "comments": "Comments",
    "freeze_panes": "Freeze Panes",
    "pivot_tables": "Pivot Tables",
    "named_ranges": "Named Ranges",
    "tables": "Tables",
}

# Tier boundary lines (drawn after these row indices)
_TIER_BOUNDARIES = [2, 8]  # after Sheets, after Text Fmt


def render_heatmap(results_json: Path, output_dir: Path) -> list[Path]:
    """Generate heatmap PNG + SVG from a results.json file.

    Returns list of generated file paths.
    """
    matplotlib.use("Agg")

    with open(results_json) as f:
        data = json.load(f)

    matrix, feature_labels, lib_labels = _build_matrix(data)

    if matrix.size == 0:
        return []

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    paths: list[Path] = []

    for ext in ("png", "svg"):
        path = output_dir / f"heatmap.{ext}"
        _render(matrix, feature_labels, lib_labels, path)
        paths.append(path)

    return paths


def _build_matrix(
    data: dict[str, Any],
) -> tuple[Any, list[str], list[str]]:
    """Build the score matrix from results JSON.

    Returns (np.ndarray[features x libs], feature_labels, lib_labels).
    Each cell is the best-of(read, write) score, or -1 for N/A.
    """
    # Build score lookup: (feature, lib) -> best score
    score_map: dict[tuple[str, str], int] = {}
    all_features: set[str] = set()
    all_libs: set[str] = set()

    for entry in data.get("results", []):
        feat = entry["feature"]
        lib = entry["library"]
        scores = entry.get("scores", {})
        r = scores.get("read")
        w = scores.get("write")
        non_null = [s for s in [r, w] if s is not None]
        best = max(non_null) if non_null else None
        if best is not None:
            score_map[(feat, lib)] = best
            all_features.add(feat)
            all_libs.add(lib)

    # Order features by _FEATURE_ORDER, dropping any not in results
    features = [f for f in _FEATURE_ORDER if f in all_features]
    # Add any features not in our static order
    for f in sorted(all_features):
        if f not in features:
            features.append(f)

    # Sort libraries by total green count descending
    lib_green: dict[str, int] = {}
    for lib in all_libs:
        lib_green[lib] = sum(1 for f in features if score_map.get((f, lib)) == 3)
    libs = sorted(all_libs, key=lambda x: (-lib_green[x], x))

    # Build matrix
    matrix = np.full((len(features), len(libs)), -1, dtype=int)
    for i, feat in enumerate(features):
        for j, lib in enumerate(libs):
            if (feat, lib) in score_map:
                matrix[i, j] = score_map[(feat, lib)]

    feature_labels = [_FEATURE_LABELS.get(f, f) for f in features]
    return matrix, feature_labels, libs


def _render(
    matrix: Any,
    feature_labels: list[str],
    lib_labels: list[str],
    output_path: Path,
) -> None:
    """Render the heatmap to a file."""
    n_features, n_libs = matrix.shape

    # Color map: -1=light gray (N/A), 0=red, 1=orange, 2=yellow, 3=green
    colors = ["#e0e0e0", "#e74c3c", "#e67e22", "#f1c40f", "#2ecc71"]
    cmap = ListedColormap(colors)

    # Shift matrix so -1 maps to index 0, 0->1, 1->2, 2->3, 3->4
    plot_data = matrix + 1

    cell_w = 0.7
    cell_h = 0.45
    fig_w = max(3.5 + n_libs * cell_w, 8)
    fig_h = max(1.5 + n_features * cell_h, 4)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.imshow(plot_data, cmap=cmap, vmin=0, vmax=4, aspect="auto")

    # Annotate cells
    score_labels = {-1: "", 0: "0", 1: "1", 2: "2", 3: "3"}
    for i in range(n_features):
        for j in range(n_libs):
            val = matrix[i, j]
            text = score_labels.get(val, "")
            text_color = "white" if val in (0, 3) else ("black" if val >= 0 else "#999")
            ax.text(j, i, text, ha="center", va="center", fontsize=9, fontweight="bold",
                    color=text_color)

    # Tier boundary lines
    for boundary_row in _TIER_BOUNDARIES:
        if boundary_row < n_features:
            ax.axhline(y=boundary_row + 0.5, color="#555", linewidth=1.5, linestyle="-")

    # Axis labels
    ax.set_xticks(range(n_libs))
    ax.set_xticklabels(lib_labels, rotation=45, ha="right", fontsize=8)
    ax.set_yticks(range(n_features))
    ax.set_yticklabels(feature_labels, fontsize=8)

    # Move x labels to top
    ax.xaxis.tick_top()
    ax.xaxis.set_label_position("top")

    ax.set_title("ExcelBench — Feature Fidelity (best of R/W)", fontsize=12,
                 fontweight="bold", pad=15)

    # Legend in bottom-right
    from matplotlib.patches import Patch
    legend_items = [
        Patch(facecolor="#2ecc71", label="3 — Full"),
        Patch(facecolor="#f1c40f", label="2 — Functional"),
        Patch(facecolor="#e67e22", label="1 — Minimal"),
        Patch(facecolor="#e74c3c", label="0 — Unsupported"),
        Patch(facecolor="#e0e0e0", label="N/A"),
    ]
    ax.legend(handles=legend_items, loc="lower right", fontsize=7,
              bbox_to_anchor=(1.0, -0.15), ncol=5, frameon=False)

    plt.tight_layout()
    dpi = 200 if output_path.suffix == ".png" else 150
    fig.savefig(output_path, dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
