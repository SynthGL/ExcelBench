"""Interactive Plotly scatter-plots: fidelity pass-rate vs. throughput.

Drop-in replacement for the static matplotlib scatter panels.
Returns self-contained HTML fragments that can be embedded in the
single-file HTML dashboard.

Public API:
  render_interactive_scatter_tiers(fidelity_json, perf_json) -> str
  render_interactive_scatter_features(fidelity_json, perf_json) -> str
  render_interactive_scatter_tiers_from_data(fidelity, perf) -> str
  render_interactive_scatter_features_from_data(fidelity, perf) -> str
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ── Reuse data pipeline + constants from the static renderer ──────
from excelbench.results.scatter import (
    _CAP_LABELS,
    _CAP_MARKERS,
    _DARK_BG,
    _DARK_CARD,
    _DARK_GRID,
    _DARK_TEXT,
    _DARK_TEXT2,
    _FALLBACK_COLOR,
    _FEATURE_LABELS,
    _FEATURE_PERF_MAP,
    _LIB_COLORS,
    _LIB_SHORT,
    _TIER_GROUPS,
    _WOLFXL_MARKER_SCALE,
    _ZONE_BANDS,
    Point,
    _compute_capabilities,
    _compute_pass_rates,
    _compute_representative_throughput,
    _compute_throughputs,
    _feature_points,
    _jitter,
    _overall_points,
    _tier_points,
)

# ── Plotly marker-symbol mapping (matplotlib → plotly names) ──────
_PLOTLY_MARKERS: dict[str, str] = {
    "o": "circle",         # R+W
    "s": "square",         # R (read-only)
    "D": "diamond-wide",   # W (write-only)
}

# Plotly to_html config shared across all renders
_PLOTLY_CONFIG: dict[str, Any] = dict(
    displayModeBar=True,
    modeBarButtonsToRemove=["lasso2d", "select2d"],
    displaylogo=False,
    responsive=True,
)


# ====================================================================
#  Subplot axis-id helper
# ====================================================================


def _subplot_axis_suffix(row: int, col: int, ncols: int) -> str:
    """Return the Plotly axis suffix for a given (row, col) in a grid.

    Plotly numbers axes sequentially: first subplot → '', second → '2', etc.
    Index = (row - 1) * ncols + col.
    """
    idx = (row - 1) * ncols + col
    return "" if idx == 1 else str(idx)


# ====================================================================
#  Panel builder
# ====================================================================


def _build_panel(
    fig: go.Figure,
    row: int,
    col: int,
    ncols: int,
    points: list[Point],
) -> None:
    """Populate one subplot with zone bands, threshold lines, and scatter points."""
    suffix = _subplot_axis_suffix(row, col, ncols)
    xref = f"x{suffix}" if suffix else "x"
    yref = f"y{suffix}" if suffix else "y"

    # ── Score-zone background bands ──
    for y_lo, y_hi, colour, _ in _ZONE_BANDS:
        fig.add_shape(
            type="rect",
            xref=f"{xref} domain",
            yref=yref,
            x0=0, x1=1, y0=y_lo, y1=y_hi,
            fillcolor=colour,
            opacity=0.65,
            layer="below",
            line_width=0,
        )

    # ── Threshold lines ──
    for y_val, dash_style, colour, width in [
        (50, "dash", "#444444", 0.8),
        (80, "dash", "#444444", 0.8),
        (100, "solid", "#62c073", 1.0),
    ]:
        fig.add_hline(
            y=y_val, row=row, col=col,
            line=dict(color=colour, width=width, dash=dash_style),
            layer="below",
        )

    # ── Score zone annotations (left edge) ──
    for y_pos, label, colour in [
        (25, "Score 0", "#e4484d"),
        (65, "Score 1", "#e79c12"),
        (90, "Score 2", "#4ead5b"),
        (106, "Score 3", "#62c073"),
    ]:
        fig.add_annotation(
            xref=f"{xref} domain",
            yref=yref,
            x=0.01, y=y_pos,
            text=label,
            showarrow=False,
            font=dict(size=11, color=colour),
            opacity=0.7,
            xanchor="left",
            yanchor="middle",
        )

    if not points:
        fig.add_annotation(
            xref=f"{xref} domain",
            yref=f"{yref} domain",
            x=0.5, y=0.5,
            text="No data",
            showarrow=False,
            font=dict(size=14, color="#a0a0a0"),
        )
        return

    # ── Plot each library as a separate trace for individual hover ──
    for lib, rate, tp, cap in points:
        color = _LIB_COLORS.get(lib, _FALLBACK_COLOR)
        mpl_marker = _CAP_MARKERS.get(cap, "o")
        plotly_symbol = _PLOTLY_MARKERS.get(mpl_marker, "circle")
        display_name = _LIB_SHORT.get(lib, lib)

        y = rate + _jitter(lib, scale=3.0)
        y = max(-2.0, min(107.0, y))

        is_wolfxl = lib == "wolfxl"
        sz = int(16 * _WOLFXL_MARKER_SCALE) if is_wolfxl else 16
        edge_color = "#fb923c" if is_wolfxl else _DARK_CARD
        edge_width = 2.5 if is_wolfxl else 1.0

        fig.add_trace(
            go.Scatter(
                x=[tp],
                y=[y],
                mode="markers+text",
                marker=dict(
                    symbol=plotly_symbol,
                    size=sz,
                    color=color,
                    line=dict(color=edge_color, width=edge_width),
                    opacity=0.95 if is_wolfxl else 0.92,
                ),
                text=[display_name],
                textposition="top center" if is_wolfxl else "bottom center",
                textfont=dict(
                    size=14 if is_wolfxl else 11,
                    color=color,
                    family="system-ui, sans-serif",
                ),
                hovertemplate=(
                    f"<b>{display_name}</b><br>"
                    f"Pass Rate: {rate:.1f}%<br>"
                    f"Throughput: %{{x:,.0f}} ops/s<br>"
                    f"Capability: {_CAP_LABELS.get(cap, cap)}"
                    "<extra></extra>"
                ),
                showlegend=False,
            ),
            row=row, col=col,
        )


# ====================================================================
#  Layout & figure builders
# ====================================================================


def _apply_layout(fig: go.Figure, title: str, nrows: int, ncols: int) -> None:
    """Apply shared dark-theme layout settings to a multi-subplot figure."""
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=20, color=_DARK_TEXT, family="system-ui, sans-serif"),
            x=0.5,
            xanchor="center",
        ),
        paper_bgcolor=_DARK_BG,
        plot_bgcolor=_DARK_CARD,
        font=dict(color=_DARK_TEXT, family="system-ui, sans-serif"),
        margin=dict(l=70, r=40, t=90, b=90),
        showlegend=False,
        height=550 * nrows + 130,
    )

    # Style all axes
    axis_style: dict[str, Any] = dict(
        gridcolor=_DARK_GRID,
        gridwidth=0.5,
        zerolinecolor=_DARK_GRID,
        linecolor=_DARK_GRID,
        tickfont=dict(size=12, color=_DARK_TEXT2),
    )

    n_panels = nrows * ncols
    for i in range(1, n_panels + 1):
        suffix = "" if i == 1 else str(i)
        fig.update_layout(**{
            f"xaxis{suffix}": dict(
                type="log",
                title=dict(text="Throughput (ops/s)", font=dict(size=13, color=_DARK_TEXT)),
                **axis_style,
            ),
            f"yaxis{suffix}": dict(
                range=[-5, 110],
                title=dict(text="Pass Rate (%)", font=dict(size=13, color=_DARK_TEXT)),
                **axis_style,
            ),
        })

    # Modebar customization
    fig.update_layout(
        modebar=dict(
            bgcolor="rgba(0,0,0,0)",
            color=_DARK_TEXT2,
            activecolor="#51a8ff",
        ),
    )


def _build_tiers_figure(
    fidelity: dict[str, Any],
    perf: dict[str, Any],
) -> go.Figure:
    """Build the 1x3 tier scatter grid."""
    pass_rates = _compute_pass_rates(fidelity)
    throughputs = _compute_throughputs(perf)
    rep_tp = _compute_representative_throughput(perf)
    caps = _compute_capabilities(fidelity)

    tier_titles = [name for name, _ in _TIER_GROUPS] + ["Overall"]
    fig = make_subplots(
        rows=1, cols=3,
        subplot_titles=tier_titles,
        horizontal_spacing=0.06,
    )

    # Tier panels
    for col_idx, (_, tier_features) in enumerate(_TIER_GROUPS, start=1):
        pts = _tier_points(tier_features, pass_rates, throughputs, caps)
        _build_panel(fig, 1, col_idx, ncols=3, points=pts)

    # Overall panel
    overall = _overall_points(pass_rates, rep_tp, caps)
    _build_panel(fig, 1, 3, ncols=3, points=overall)

    _apply_layout(fig, "ExcelBench \u2014 Fidelity vs. Throughput by Feature Group", 1, 3)

    # Style subplot titles only (not zone labels added by _build_panel)
    n_subplot_titles = len(tier_titles)
    for ann in fig.layout.annotations[:n_subplot_titles]:
        ann.font = dict(size=16, color=_DARK_TEXT, family="system-ui, sans-serif")

    _add_footer_annotations(fig)
    return fig


def _build_features_figure(
    fidelity: dict[str, Any],
    perf: dict[str, Any],
) -> go.Figure:
    """Build the 2x3 per-feature scatter grid."""
    pass_rates = _compute_pass_rates(fidelity)
    throughputs = _compute_throughputs(perf)
    caps = _compute_capabilities(fidelity)

    features = list(_FEATURE_PERF_MAP.keys())
    feature_titles = [_FEATURE_LABELS.get(f, f) for f in features]

    fig = make_subplots(
        rows=2, cols=3,
        subplot_titles=feature_titles,
        vertical_spacing=0.12,
        horizontal_spacing=0.06,
    )

    for idx, feature in enumerate(features):
        r = idx // 3 + 1
        c = idx % 3 + 1
        pts = _feature_points(feature, pass_rates, throughputs, caps)
        _build_panel(fig, r, c, ncols=3, points=pts)

    _apply_layout(fig, "ExcelBench \u2014 Per-Feature Fidelity vs. Throughput", 2, 3)

    # Style subplot titles only (not zone labels added by _build_panel)
    n_subplot_titles = len(feature_titles)
    for ann in fig.layout.annotations[:n_subplot_titles]:
        ann.font = dict(size=16, color=_DARK_TEXT, family="system-ui, sans-serif")

    _add_footer_annotations(fig)
    return fig


def _add_footer_annotations(fig: go.Figure) -> None:
    """Add capability legend and jitter footnote at the bottom of the figure."""
    legend_parts = []
    for cap, mpl_marker in _CAP_MARKERS.items():
        symbol_char = {"o": "\u25cf", "s": "\u25a0", "D": "\u25c6"}.get(mpl_marker, "\u25cf")
        legend_parts.append(f"{symbol_char} {_CAP_LABELS[cap]}")
    legend_text = "    ".join(legend_parts)

    fig.add_annotation(
        text=f"{legend_text}    |    Y-positions include \u00b13% jitter for visibility",
        xref="paper", yref="paper",
        x=0.5, y=-0.08,
        showarrow=False,
        font=dict(size=12, color=_DARK_TEXT2),
        xanchor="center",
    )


# ====================================================================
#  Public API — returns HTML fragments
# ====================================================================


def render_interactive_scatter_tiers(
    fidelity_json: Path,
    perf_json: Path,
) -> str:
    """Generate an interactive 1x3 tier scatter grid as an HTML fragment.

    The returned string includes the plotly.js library inline.
    """
    fidelity = json.loads(fidelity_json.read_text())
    perf = json.loads(perf_json.read_text())
    fig = _build_tiers_figure(fidelity, perf)
    return str(fig.to_html(full_html=False, include_plotlyjs=True, config=_PLOTLY_CONFIG))


def render_interactive_scatter_features(
    fidelity_json: Path,
    perf_json: Path,
) -> str:
    """Generate an interactive 2x3 per-feature scatter grid as an HTML fragment.

    Assumes plotly.js was already included by the tiers fragment.
    """
    fidelity = json.loads(fidelity_json.read_text())
    perf = json.loads(perf_json.read_text())
    fig = _build_features_figure(fidelity, perf)
    return str(fig.to_html(full_html=False, include_plotlyjs=False, config=_PLOTLY_CONFIG))


# ── Convenience: build from raw dicts (for dashboard integration) ──


def render_interactive_scatter_tiers_from_data(
    fidelity: dict[str, Any],
    perf: dict[str, Any],
) -> str:
    """Like render_interactive_scatter_tiers but accepts pre-loaded dicts."""
    fig = _build_tiers_figure(fidelity, perf)
    return str(fig.to_html(full_html=False, include_plotlyjs=True, config=_PLOTLY_CONFIG))


def render_interactive_scatter_features_from_data(
    fidelity: dict[str, Any],
    perf: dict[str, Any],
) -> str:
    """Like render_interactive_scatter_features but accepts pre-loaded dicts."""
    fig = _build_features_figure(fidelity, perf)
    return str(fig.to_html(full_html=False, include_plotlyjs=False, config=_PLOTLY_CONFIG))
