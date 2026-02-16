#!/usr/bin/env python3
# ruff: noqa: E501
"""Generate SVG benchmark charts for the WolfXL README.

Creates dark and light theme variants of a horizontal bar chart
comparing WolfXL vs openpyxl throughput.
"""

from __future__ import annotations

from pathlib import Path

# Benchmark data from large_scale_benchmark.py runs
BENCHMARKS = {
    "Read 10M cells": {"wolfxl": 770_000, "openpyxl": 209_000},
    "Write 10M cells": {"wolfxl": 1_500_000, "openpyxl": 314_000},
    "Read 1M cells": {"wolfxl": 877_000, "openpyxl": 247_000},
    "Write 1M cells": {"wolfxl": 1_170_000, "openpyxl": 342_000},
    "Read 100K cells": {"wolfxl": 911_000, "openpyxl": 239_000},
    "Write 100K cells": {"wolfxl": 1_620_000, "openpyxl": 353_000},
}


def format_throughput(cells_per_sec: int) -> str:
    if cells_per_sec >= 1_000_000:
        return f"{cells_per_sec / 1_000_000:.1f}M/s"
    else:
        return f"{cells_per_sec / 1_000:.0f}K/s"


def generate_benchmark_chart(dark: bool = True) -> str:
    """Generate an SVG horizontal bar chart."""
    # Theme colors
    if dark:
        bg = "#0d1117"
        text_color = "#e6edf3"
        muted_text = "#8b949e"
        wolfxl_color = "#58a6ff"
        openpyxl_color = "#484f58"
        grid_color = "#21262d"
        bar_label_color = "#e6edf3"
        subtitle_color = "#8b949e"
    else:
        bg = "#ffffff"
        text_color = "#1f2328"
        muted_text = "#656d76"
        wolfxl_color = "#0969da"
        openpyxl_color = "#d0d7de"
        grid_color = "#d0d7de"
        bar_label_color = "#1f2328"
        subtitle_color = "#656d76"

    # Chart dimensions
    chart_width = 720
    chart_height = 460
    margin_left = 160
    margin_right = 100
    margin_top = 70
    margin_bottom = 50
    bar_area_width = chart_width - margin_left - margin_right

    # Only show the headline comparisons
    items = [
        ("Write 100K cells", BENCHMARKS["Write 100K cells"]),
        ("Read 100K cells", BENCHMARKS["Read 100K cells"]),
        ("Write 1M cells", BENCHMARKS["Write 1M cells"]),
        ("Read 1M cells", BENCHMARKS["Read 1M cells"]),
        ("Write 10M cells", BENCHMARKS["Write 10M cells"]),
        ("Read 10M cells", BENCHMARKS["Read 10M cells"]),
    ]

    bar_height = 20
    pair_gap = 12
    group_gap = 28
    n_groups = len(items)

    # Calculate total height needed
    content_height = n_groups * (2 * bar_height + pair_gap + group_gap) - group_gap
    chart_height = margin_top + content_height + margin_bottom + 10

    max_val = max(max(d["wolfxl"], d["openpyxl"]) for _, d in items)
    # Round up to nice number
    max_val = ((max_val // 500_000) + 1) * 500_000

    def x_pos(val: int) -> float:
        return margin_left + (val / max_val) * bar_area_width

    svg_parts = []
    svg_parts.append(
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {chart_width} {chart_height}" '
        f"font-family=\"-apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif\">"
    )

    # Background
    svg_parts.append(f'<rect width="{chart_width}" height="{chart_height}" fill="{bg}" rx="8"/>')

    # Title
    svg_parts.append(
        f'<text x="{chart_width // 2}" y="28" text-anchor="middle" '
        f'fill="{text_color}" font-size="16" font-weight="600">'
        f"WolfXL vs openpyxl — Bulk Throughput (cells/sec)</text>"
    )
    svg_parts.append(
        f'<text x="{chart_width // 2}" y="48" text-anchor="middle" '
        f'fill="{subtitle_color}" font-size="11">'
        f"Apple M1 Pro, Python 3.12, median of 3 runs</text>"
    )

    # Grid lines
    grid_steps = [0, 500_000, 1_000_000, 1_500_000]
    if max_val > 1_500_000:
        grid_steps.append(2_000_000)
    for val in grid_steps:
        gx = x_pos(val)
        svg_parts.append(
            f'<line x1="{gx}" y1="{margin_top}" x2="{gx}" y2="{chart_height - margin_bottom}" '
            f'stroke="{grid_color}" stroke-width="1" stroke-dasharray="3,3"/>'
        )
        label = f"{val / 1_000_000:.1f}M" if val >= 1_000_000 else f"{val // 1_000}K"
        if val == 0:
            label = "0"
        svg_parts.append(
            f'<text x="{gx}" y="{chart_height - margin_bottom + 18}" text-anchor="middle" '
            f'fill="{muted_text}" font-size="10">{label}</text>'
        )

    # Bars
    y_cursor = margin_top + 5
    for label, data in items:
        wolfxl_val = data["wolfxl"]
        openpyxl_val = data["openpyxl"]

        # Group label
        svg_parts.append(
            f'<text x="{margin_left - 10}" y="{y_cursor + bar_height + pair_gap // 2 + 3}" '
            f'text-anchor="end" fill="{text_color}" font-size="12">{label}</text>'
        )

        # WolfXL bar
        w = x_pos(wolfxl_val) - margin_left
        svg_parts.append(
            f'<rect x="{margin_left}" y="{y_cursor}" width="{w}" height="{bar_height}" '
            f'fill="{wolfxl_color}" rx="3"/>'
        )
        tp = format_throughput(wolfxl_val)
        svg_parts.append(
            f'<text x="{margin_left + w + 6}" y="{y_cursor + bar_height - 5}" '
            f'fill="{bar_label_color}" font-size="11" font-weight="500">{tp}</text>'
        )

        y_cursor += bar_height + pair_gap

        # openpyxl bar
        w2 = x_pos(openpyxl_val) - margin_left
        svg_parts.append(
            f'<rect x="{margin_left}" y="{y_cursor}" width="{w2}" height="{bar_height}" '
            f'fill="{openpyxl_color}" rx="3"/>'
        )
        tp2 = format_throughput(openpyxl_val)
        svg_parts.append(
            f'<text x="{margin_left + w2 + 6}" y="{y_cursor + bar_height - 5}" '
            f'fill="{muted_text}" font-size="11">{tp2}</text>'
        )

        y_cursor += bar_height + group_gap

    # Legend
    legend_y = chart_height - 18
    legend_x = margin_left
    svg_parts.append(
        f'<rect x="{legend_x}" y="{legend_y - 8}" width="12" height="12" fill="{wolfxl_color}" rx="2"/>'
    )
    svg_parts.append(
        f'<text x="{legend_x + 16}" y="{legend_y + 2}" fill="{text_color}" font-size="11" '
        f'font-weight="500">WolfXL</text>'
    )
    svg_parts.append(
        f'<rect x="{legend_x + 80}" y="{legend_y - 8}" width="12" height="12" fill="{openpyxl_color}" rx="2"/>'
    )
    svg_parts.append(
        f'<text x="{legend_x + 96}" y="{legend_y + 2}" fill="{muted_text}" font-size="11">openpyxl</text>'
    )

    svg_parts.append("</svg>")
    return "\n".join(svg_parts)


def generate_architecture_diagram(dark: bool = True) -> str:
    """Generate an SVG architecture diagram showing the three modes."""
    if dark:
        bg = "#0d1117"
        text_color = "#e6edf3"
        muted = "#8b949e"
        box_bg = "#161b22"
        box_border = "#30363d"
        accent_read = "#3fb950"
        accent_write = "#58a6ff"
        accent_modify = "#d29922"
        arrow_color = "#484f58"
        rust_color = "#dea584"
    else:
        bg = "#ffffff"
        text_color = "#1f2328"
        muted = "#656d76"
        box_bg = "#f6f8fa"
        box_border = "#d0d7de"
        accent_read = "#1a7f37"
        accent_write = "#0969da"
        accent_modify = "#9a6700"
        arrow_color = "#afb8c1"
        rust_color = "#b7410e"

    w, h = 680, 300

    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}"
  font-family="-apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif">
<rect width="{w}" height="{h}" fill="{bg}" rx="8"/>

<!-- Title -->
<text x="{w // 2}" y="28" text-anchor="middle" fill="{text_color}" font-size="14" font-weight="600">Architecture</text>

<!-- Python API layer -->
<rect x="40" y="55" width="600" height="50" fill="{box_bg}" stroke="{box_border}" rx="8"/>
<text x="{w // 2}" y="75" text-anchor="middle" fill="{text_color}" font-size="13" font-weight="600">Python API</text>
<text x="{w // 2}" y="93" text-anchor="middle" fill="{muted}" font-size="11">ws["A1"].value &middot; Font(bold=True) &middot; wb.save() &middot; iter_rows()</text>

<!-- Arrows -->
<line x1="150" y1="105" x2="150" y2="140" stroke="{arrow_color}" stroke-width="2" marker-end="url(#arrow-{("d" if dark else "l")})"/>
<line x1="340" y1="105" x2="340" y2="140" stroke="{arrow_color}" stroke-width="2" marker-end="url(#arrow-{("d" if dark else "l")})"/>
<line x1="530" y1="105" x2="530" y2="140" stroke="{arrow_color}" stroke-width="2" marker-end="url(#arrow-{("d" if dark else "l")})"/>

<!-- Arrow marker -->
<defs>
  <marker id="arrow-{("d" if dark else "l")}" viewBox="0 0 10 10" refX="5" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-auto">
    <path d="M 0 0 L 10 5 L 0 10 z" fill="{arrow_color}"/>
  </marker>
</defs>

<!-- Read mode -->
<rect x="40" y="140" width="220" height="65" fill="{box_bg}" stroke="{accent_read}" stroke-width="2" rx="8"/>
<text x="150" y="163" text-anchor="middle" fill="{accent_read}" font-size="12" font-weight="600">Read Mode</text>
<text x="150" y="180" text-anchor="middle" fill="{muted}" font-size="10">load_workbook(path)</text>
<text x="150" y="195" text-anchor="middle" fill="{muted}" font-size="10">Full style extraction</text>

<!-- Write mode -->
<rect x="230" y="140" width="220" height="65" fill="{box_bg}" stroke="{accent_write}" stroke-width="2" rx="8"/>
<text x="340" y="163" text-anchor="middle" fill="{accent_write}" font-size="12" font-weight="600">Write Mode</text>
<text x="340" y="180" text-anchor="middle" fill="{muted}" font-size="10">Workbook()</text>
<text x="340" y="195" text-anchor="middle" fill="{muted}" font-size="10">New file from scratch</text>

<!-- Modify mode -->
<rect x="420" y="140" width="220" height="65" fill="{box_bg}" stroke="{accent_modify}" stroke-width="2" rx="8"/>
<text x="530" y="163" text-anchor="middle" fill="{accent_modify}" font-size="12" font-weight="600">Modify Mode</text>
<text x="530" y="180" text-anchor="middle" fill="{muted}" font-size="10">load_workbook(modify=True)</text>
<text x="530" y="195" text-anchor="middle" fill="{muted}" font-size="10">Preserves charts/macros</text>

<!-- Arrows to Rust -->
<line x1="150" y1="205" x2="150" y2="235" stroke="{arrow_color}" stroke-width="2" marker-end="url(#arrow-{("d" if dark else "l")})"/>
<line x1="340" y1="205" x2="340" y2="235" stroke="{arrow_color}" stroke-width="2" marker-end="url(#arrow-{("d" if dark else "l")})"/>
<line x1="530" y1="205" x2="530" y2="235" stroke="{arrow_color}" stroke-width="2" marker-end="url(#arrow-{("d" if dark else "l")})"/>

<!-- Rust engines -->
<rect x="40" y="235" width="220" height="45" fill="{box_bg}" stroke="{rust_color}" rx="8"/>
<text x="150" y="258" text-anchor="middle" fill="{rust_color}" font-size="12" font-weight="600">calamine-styles</text>
<text x="150" y="273" text-anchor="middle" fill="{muted}" font-size="10">Rust XLSX parser + styles</text>

<rect x="230" y="235" width="220" height="45" fill="{box_bg}" stroke="{rust_color}" rx="8"/>
<text x="340" y="258" text-anchor="middle" fill="{rust_color}" font-size="12" font-weight="600">rust_xlsxwriter</text>
<text x="340" y="273" text-anchor="middle" fill="{muted}" font-size="10">Rust XLSX writer</text>

<rect x="420" y="235" width="220" height="45" fill="{box_bg}" stroke="{rust_color}" rx="8"/>
<text x="530" y="258" text-anchor="middle" fill="{rust_color}" font-size="12" font-weight="600">XlsxPatcher</text>
<text x="530" y="273" text-anchor="middle" fill="{muted}" font-size="10">Surgical ZIP patcher</text>

</svg>'''
    return svg


def main() -> None:
    out_dir = Path("../wolfxl/assets")  # wolfxl is now a standalone repo
    out_dir.mkdir(parents=True, exist_ok=True)

    # Benchmark charts
    for theme, dark in [("dark", True), ("light", False)]:
        svg = generate_benchmark_chart(dark=dark)
        path = out_dir / f"benchmark-{theme}.svg"
        path.write_text(svg)
        print(f"  Wrote {path}")

    # Architecture diagrams
    for theme, dark in [("dark", True), ("light", False)]:
        svg = generate_architecture_diagram(dark=dark)
        path = out_dir / f"architecture-{theme}.svg"
        path.write_text(svg)
        print(f"  Wrote {path}")


if __name__ == "__main__":
    main()
