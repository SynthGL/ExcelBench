"""Smoke tests for the interactive Plotly scatter renderer."""

from __future__ import annotations

from pathlib import Path

from excelbench.results.scatter_interactive import (
    render_interactive_scatter_features,
    render_interactive_scatter_features_from_data,
    render_interactive_scatter_tiers,
    render_interactive_scatter_tiers_from_data,
)

# ── Fixture paths (same as test_results_scatter.py) ──────────────


def _fixture_paths() -> tuple[Path, Path]:
    fidelity = Path("results/xlsx/results.json")
    perf = Path("results/perf/results.json")
    assert fidelity.exists(), "expected repo fixture results/xlsx/results.json"
    assert perf.exists(), "expected repo fixture results/perf/results.json"
    return fidelity, perf


# ── Tiers tests ───────────────────────────────────────────────────


def test_tiers_returns_html_with_plotlyjs() -> None:
    fidelity, perf = _fixture_paths()
    html = render_interactive_scatter_tiers(fidelity, perf)

    assert isinstance(html, str)
    assert len(html) > 10_000, "expected substantial HTML output"
    # Plotly.js must be inlined
    assert "plotly" in html.lower()
    assert "<script" in html.lower()


def test_tiers_contains_dark_theme_colors() -> None:
    fidelity, perf = _fixture_paths()
    html = render_interactive_scatter_tiers(fidelity, perf)

    assert "#0a0a0a" in html, "expected dark background color"
    assert "#191919" in html, "expected dark card color"


def test_tiers_contains_hover_template() -> None:
    fidelity, perf = _fixture_paths()
    html = render_interactive_scatter_tiers(fidelity, perf)

    assert "Pass Rate" in html
    assert "Throughput" in html
    # ops/s appears in hovertemplate; "/" is unicode-escaped as \u002f in Plotly JSON
    assert "ops" in html and ("ops/s" in html or "ops\\u002fs" in html)


def test_tiers_contains_wolfxl_highlighting() -> None:
    fidelity, perf = _fixture_paths()
    html = render_interactive_scatter_tiers(fidelity, perf)

    assert "wolfxl" in html.lower()
    assert "#fb923c" in html, "expected WolfXL orange edge color"


def test_tiers_contains_zone_bands() -> None:
    fidelity, perf = _fixture_paths()
    html = render_interactive_scatter_tiers(fidelity, perf)

    # Zone band colors from _ZONE_BANDS
    assert "#1a0a0a" in html, "expected score-0 zone band color"
    assert "#1a1400" in html, "expected score-1 zone band color"


# ── Features tests ────────────────────────────────────────────────


def test_features_returns_html_without_plotlyjs() -> None:
    fidelity, perf = _fixture_paths()
    html = render_interactive_scatter_features(fidelity, perf)

    assert isinstance(html, str)
    assert len(html) > 1_000
    # Should NOT re-include plotly.js (already included by tiers)
    # Check that there is no massive plotly bundle (the full bundle is ~4MB)
    assert len(html) < 500_000, "features fragment should not re-include plotly.js"


def test_features_contains_feature_labels() -> None:
    fidelity, perf = _fixture_paths()
    html = render_interactive_scatter_features(fidelity, perf)

    # At least some feature labels should appear in subplot titles
    assert "Cell Values" in html
    assert "Formulas" in html


# ── From-data convenience API tests ───────────────────────────────


def test_from_data_tiers_matches_file_api() -> None:
    """The from_data API should produce equivalent output to the file-based API."""
    import json

    fidelity_path, perf_path = _fixture_paths()
    fidelity = json.loads(fidelity_path.read_text())
    perf = json.loads(perf_path.read_text())

    html_file = render_interactive_scatter_tiers(fidelity_path, perf_path)
    html_data = render_interactive_scatter_tiers_from_data(fidelity, perf)

    # Both should include plotly.js and have similar structure
    assert abs(len(html_file) - len(html_data)) < 100, (
        "file-based and dict-based outputs should be nearly identical"
    )


def test_from_data_features_matches_file_api() -> None:
    import json

    fidelity_path, perf_path = _fixture_paths()
    fidelity = json.loads(fidelity_path.read_text())
    perf = json.loads(perf_path.read_text())

    html_file = render_interactive_scatter_features(fidelity_path, perf_path)
    html_data = render_interactive_scatter_features_from_data(fidelity, perf)

    assert abs(len(html_file) - len(html_data)) < 100
