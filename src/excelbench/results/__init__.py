"""Results rendering and output."""

from excelbench.results.renderer import render_csv, render_json, render_markdown, render_results

__all__ = ["render_csv", "render_json", "render_markdown", "render_results"]

# dashboard and heatmap are imported directly where needed (lazy)
