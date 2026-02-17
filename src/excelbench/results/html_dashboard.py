"""Single-file interactive HTML dashboard for all ExcelBench results.

Generates one self-contained .html file with:
  - Score matrix (interactive heatmap)
  - Scatter-plot SVGs (embedded inline)
  - Sortable library comparison table
  - Expandable per-feature test-case detail
  - Performance workload tables with phase breakdowns
  - Diagnostics summary
"""

from __future__ import annotations

import html as html_mod
import json
import re
from pathlib import Path
from typing import Any

from excelbench.results.report_policy import (
    filter_memory_rows,
    filter_report_data,
    modify_mode_label,
)

# ── Feature ordering / tier map (shared with other renderers) ──────

_FEATURE_ORDER: list[str] = [
    "cell_values", "formulas", "multiple_sheets",
    "alignment", "background_colors", "borders",
    "dimensions", "number_formats", "text_formatting",
    "comments", "conditional_formatting", "data_validation",
    "freeze_panes", "hyperlinks", "images", "merged_cells",
    "pivot_tables",
    "named_ranges", "tables",
]

_FEATURE_LABELS: dict[str, str] = {
    "cell_values": "Cell Values", "formulas": "Formulas",
    "multiple_sheets": "Multiple Sheets", "alignment": "Alignment",
    "background_colors": "Background Colors", "borders": "Borders",
    "dimensions": "Dimensions", "number_formats": "Number Formats",
    "text_formatting": "Text Formatting", "comments": "Comments",
    "conditional_formatting": "Cond. Formatting",
    "data_validation": "Data Validation", "freeze_panes": "Freeze Panes",
    "hyperlinks": "Hyperlinks", "images": "Images",
    "merged_cells": "Merged Cells", "pivot_tables": "Pivot Tables",
    "named_ranges": "Named Ranges", "tables": "Tables",
}

_TIER_MAP: dict[str, int] = {
    "cell_values": 0, "formulas": 0, "multiple_sheets": 0,
    "alignment": 1, "background_colors": 1, "borders": 1,
    "dimensions": 1, "number_formats": 1, "text_formatting": 1,
    "comments": 2, "conditional_formatting": 2, "data_validation": 2,
    "freeze_panes": 2, "hyperlinks": 2, "images": 2,
    "merged_cells": 2, "pivot_tables": 2,
    "named_ranges": 3, "tables": 3,
}

_TIER_NAMES: dict[int, str] = {
    0: "Tier 0 — Core",
    1: "Tier 1 — Formatting",
    2: "Tier 2 — Advanced",
    3: "Tier 3 — Workbook Metadata",
}

# ── Radar / histogram data helpers ─────────────────────────────────

_DEFAULT_RADAR_LIBS: list[str] = [
    "wolfxl", "openpyxl", "xlsxwriter", "xlsxwriter-constmem",
    "pandas", "python-calamine",
]


def _compute_radar_data(
    fidelity: dict[str, Any],
    perf: dict[str, Any] | None,
    *,
    focus_libs: list[str] | None = None,
) -> list[dict[str, Any]]:
    """Compute 5-axis radar data for selected libraries.

    Axes (all normalised 0–100):
      - Fidelity: green_count / total_scored * 100
      - Read Speed: min(100, lib_ops / openpyxl_ops * 50)
      - Write Speed: same formula for write
      - Feature Coverage: features_with_any_score / total_features * 100
      - Capability Breadth: R-only/W-only=33, R+W=66, R+W+Patch=100
    """
    results = fidelity.get("results", [])
    libs_info = fidelity.get("libraries", {})
    focus = focus_libs or _DEFAULT_RADAR_LIBS
    available = set(libs_info.keys())
    focus = [lib for lib in focus if lib in available]

    all_features = sorted({e["feature"] for e in results})
    total_features = len(all_features)

    # Per-lib aggregates from fidelity
    lib_green: dict[str, int] = {}
    lib_scored: dict[str, int] = {}
    lib_features: dict[str, set[str]] = {}
    for entry in results:
        lib = entry["library"]
        if lib not in available:
            continue
        s = entry.get("scores", {})
        best = max(
            (x for x in [s.get("read"), s.get("write")] if x is not None),
            default=None,
        )
        lib_scored[lib] = lib_scored.get(lib, 0) + (1 if best is not None else 0)
        if best == 3:
            lib_green[lib] = lib_green.get(lib, 0) + 1
        lib_features.setdefault(lib, set())
        if best is not None:
            lib_features[lib].add(entry["feature"])

    # Perf throughput (read/write ops/s via cell_values workloads)
    read_ops: dict[str, float] = {}
    write_ops: dict[str, float] = {}
    if perf:
        for e in perf.get("results", []):
            lib = e["library"]
            p = e.get("perf", {})
            # Prefer bulk workloads for throughput
            for op, store in [("read", read_ops), ("write", write_ops)]:
                if lib in store:
                    continue
                od = p.get(op)
                if not od:
                    continue
                wall = od.get("wall_ms", {})
                p50_raw = wall.get("p50") if isinstance(wall, dict) else None
                try:
                    p50 = float(p50_raw) if p50_raw is not None else None
                except (TypeError, ValueError):
                    p50 = None
                if p50 is None or p50 <= 0:
                    continue

                # New perf schema may omit op_count for feature-level benchmarks.
                # In that case, use ops/s from wall time alone for relative speed.
                oc_raw = od.get("op_count")
                if oc_raw is None:
                    store[lib] = 1000.0 / p50
                    continue

                try:
                    oc = float(oc_raw)
                except (TypeError, ValueError):
                    continue
                if oc <= 0:
                    continue
                store[lib] = oc * 1000.0 / p50

    opx_read = read_ops.get("openpyxl", 1.0)
    opx_write = write_ops.get("openpyxl", 1.0)

    # Capability breadth
    def _breadth(lib: str) -> float:
        caps = set(libs_info.get(lib, {}).get("capabilities", []))
        has_r = "read" in caps
        has_w = "write" in caps
        # wolfxl has patch/modify capability
        has_patch = lib == "wolfxl" or "modify" in caps or "patch" in caps
        if has_r and has_w and has_patch:
            return 100.0
        if has_r and has_w:
            return 66.0
        return 33.0

    out: list[dict[str, Any]] = []
    for lib in focus:
        scored = lib_scored.get(lib, 0)
        fid = (lib_green.get(lib, 0) / scored * 100) if scored else 0.0
        r_speed = min(100.0, (read_ops.get(lib, 0) / opx_read) * 50) if opx_read else 0.0
        w_speed = min(100.0, (write_ops.get(lib, 0) / opx_write) * 50) if opx_write else 0.0
        feat_count = len(lib_features.get(lib, set()))
        coverage = (feat_count / total_features * 100) if total_features else 0.0
        breadth = _breadth(lib)
        out.append({
            "library": lib,
            "axes": ["Fidelity", "Read Speed", "Write Speed", "Feature Coverage", "Capability"],
            "values": [round(fid, 1), round(r_speed, 1), round(w_speed, 1),
                       round(coverage, 1), round(breadth, 1)],
        })
    return out


# ── Helpers ─────────────────────────────────────────────────────────


def _esc(val: Any) -> str:
    return html_mod.escape(str(val)) if val is not None else ""


def _score_cls(score: int | None) -> str:
    return {3: "s3", 2: "s2", 1: "s1", 0: "s0"}.get(score, "sna")  # type: ignore[arg-type]


def _score_label(score: int | None) -> str:
    return str(score) if score is not None else "\u2014"


def _cap_label(caps: set[str] | list[str]) -> str:
    caps_set = set(caps)
    if "read" in caps_set and "write" in caps_set:
        return "R+W"
    return "R" if "read" in caps_set else "W"


def _fmt_val(val: Any) -> str:
    """Render an expected/actual value as short HTML."""
    if val is None:
        return "\u2014"
    if isinstance(val, dict):
        v = val.get("value", val)
        if isinstance(v, dict):
            parts = [f"{k}={vv}" for k, vv in v.items()
                     if vv is not None and vv is not False and vv != ""]
            short = ", ".join(parts[:6])
            if len(parts) > 6:
                short += " \u2026"
            return f"<code class='val'>{_esc(short)}</code>"
        if isinstance(v, list):
            return f"<code class='val'>[{', '.join(_esc(str(x)) for x in v[:8])}]</code>"
        return f"<code class='val'>{_esc(str(v))}</code>"
    return f"<code class='val'>{_esc(str(val))}</code>"


def _fmt_ms(val: float | None) -> str:
    if val is None:
        return "\u2014"
    if val >= 1000:
        return f"{val / 1000:.2f}s"
    return f"{val:.1f}ms"


def _fmt_rate(op_count: int | None, p50_ms: float | None) -> str:
    if p50_ms is None or p50_ms == 0:
        return "\u2014"
    if op_count is not None:
        rate = op_count * 1000.0 / p50_ms
    else:
        rate = 1000.0 / p50_ms
    if rate >= 1_000_000:
        return f"{rate / 1_000_000:.1f}M"
    if rate >= 1_000:
        return f"{rate / 1_000:.0f}K"
    return f"{rate:.0f}"


def _fmt_mb(val: float | None) -> str:
    if val is None:
        return "\u2014"
    return f"{val:.1f}"


def _safe_json(data: Any) -> str:
    """JSON for embedding inside <script>; escapes </script>."""
    return json.dumps(data, ensure_ascii=False).replace("</", r"<\/")


def _namespace_svg_ids(svg: str, prefix: str) -> str:
    """Namespace inline SVG ids so multiple charts can coexist without collisions."""
    id_matches = re.findall(r'id="([^"]+)"', svg)
    if not id_matches:
        return svg

    seen: set[str] = set()
    for old_id in id_matches:
        if old_id in seen:
            continue
        seen.add(old_id)
        new_id = f"{prefix}{old_id}"
        svg = svg.replace(f'id="{old_id}"', f'id="{new_id}"')
        svg = svg.replace(f"url(#{old_id})", f"url(#{new_id})")
        svg = svg.replace(f'href="#{old_id}"', f'href="#{new_id}"')
        svg = svg.replace(f'xlink:href="#{old_id}"', f'xlink:href="#{new_id}"')
    return svg


# ── Public API ──────────────────────────────────────────────────────


def render_html_dashboard(
    fidelity_json: Path,
    perf_json: Path | None,
    output_path: Path,
    scatter_dir: Path | None = None,
    memory_json: Path | None = None,
) -> None:
    """Generate a single self-contained HTML dashboard."""
    fidelity = json.loads(fidelity_json.read_text())
    fidelity = filter_report_data(fidelity)
    perf = None
    if perf_json and perf_json.exists():
        perf = json.loads(perf_json.read_text())
        perf = filter_report_data(perf)

    memory: list[dict[str, Any]] | None = None
    if memory_json and memory_json.exists():
        memory = json.loads(memory_json.read_text())
        memory = filter_memory_rows(memory)

    # Load SVGs from scatter_dir
    heatmap_svg: str | None = None
    scatter_svgs: dict[str, str] = {}
    if scatter_dir:
        for name, prefix in [
            ("heatmap.svg", "heatmap-"),
            ("scatter_tiers.svg", "sctier-"),
            ("scatter_features.svg", "scfeat-"),
        ]:
            svg_path = scatter_dir / name
            if svg_path.exists():
                scatter_svgs[name] = _namespace_svg_ids(svg_path.read_text(), prefix)
        heatmap_svg = scatter_svgs.get("heatmap.svg")

    body_parts = [
        _section_nav(has_memory=memory is not None),
        _section_overview(fidelity, perf),
        _section_radar(fidelity, perf),
        _section_matrix(fidelity),
        _section_scatter(fidelity, perf, heatmap_svg, scatter_svgs),
        _section_comparison(fidelity, perf),
        _section_features(fidelity),
        _section_performance(perf),
        _section_memory(memory),
        _section_diagnostics(fidelity),
        "<footer><p>Generated by ExcelBench</p></footer>",
    ]

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ExcelBench Dashboard</title>
<style>{_CSS}</style>
</head>
<body>
{"".join(body_parts)}
<script>{_JS}</script>
</body>
</html>"""

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html)


# ====================================================================
#  CSS
# ====================================================================

_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700;800&family=Manrope:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
:root{
  --bg:#06090f;
  --bg-panel:#0d121b;
  --bg-panel-2:#111926;
  --bg-elev:#141f2f;
  --border:#223247;
  --border-soft:#1a2739;
  --text:#e9eef8;
  --text-muted:#9baec9;
  --text-dim:#7688a3;
  --accent:#62b7ff;
  --accent-soft:#96d7ff;
  --accent-wolf:#ff8f37;
  --wolf-soft:#2c1707;
  --shadow-soft:0 10px 26px rgba(0,0,0,.24);
  --shadow-hard:0 18px 44px rgba(0,0,0,.34);
  --radius:14px;
  --radius-sm:10px;
  --font-display:'Sora','Avenir Next','Segoe UI',sans-serif;
  --font-body:'Manrope','Segoe UI',sans-serif;
  --font-mono:'IBM Plex Mono','JetBrains Mono',monospace;
  --g3:#0f2f1f;--g3t:#7fdca6;
  --g2:#2c250f;--g2t:#f4d287;
  --g1:#30190f;--g1t:#ffb682;
  --g0:#2d1216;--g0t:#ff838f;
  --na:#162132;--nat:#7f93af;
}
*{box-sizing:border-box;margin:0;padding:0}
body{
  font-family:var(--font-body);
  background:
    radial-gradient(1100px 700px at 3% -10%,rgba(98,183,255,.11),transparent 58%),
    radial-gradient(980px 640px at 96% -20%,rgba(255,143,55,.06),transparent 56%),
    linear-gradient(180deg,#05080e 0%,#060a11 24%,#070c15 100%);
  color:var(--text);
  font-size:14px;
  line-height:1.52;
  letter-spacing:.005em;
}
body::before{
  content:'';
  position:fixed;
  inset:0;
  pointer-events:none;
  z-index:-1;
  background-image:linear-gradient(to bottom,rgba(255,255,255,.032),transparent 62%);
  opacity:.2;
}
body::after{
  content:'';
  position:fixed;
  inset:0;
  pointer-events:none;
  z-index:-1;
  background-image:radial-gradient(circle at 1px 1px,rgba(255,255,255,.22) 1px,transparent 0);
  background-size:4px 4px;
  opacity:.035;
}
a{color:var(--accent);text-decoration:none}
a:hover{color:#b6e2ff}
.container{
  max-width:3200px;
  margin:0 auto;
  padding:1.8rem clamp(1rem,3.2vw,3.6rem);
  animation:section-in .65s ease both;
}
section{margin-bottom:2.15rem}
section:nth-of-type(1){animation-delay:.03s}
section:nth-of-type(2){animation-delay:.06s}
section:nth-of-type(3){animation-delay:.09s}
section:nth-of-type(4){animation-delay:.12s}
section:nth-of-type(5){animation-delay:.15s}
@keyframes section-in{
  from{opacity:0;transform:translateY(10px)}
  to{opacity:1;transform:translateY(0)}
}

/* ── Nav ── */
nav{
  position:sticky;
  top:0;
  z-index:100;
  padding:.78rem clamp(1rem,3.2vw,3.6rem);
  display:flex;
  align-items:center;
  gap:1.2rem;
  backdrop-filter:blur(14px);
  background:linear-gradient(120deg,rgba(8,12,19,.92),rgba(10,16,26,.9));
  box-shadow:0 7px 20px rgba(0,0,0,.3);
  border-bottom:1px solid var(--border);
}
nav .brand{
  font-family:var(--font-display);
  font-weight:800;
  font-size:1.17rem;
  color:var(--text);
  letter-spacing:.01em;
}
nav .links{display:flex;gap:.95rem;flex-wrap:wrap}
nav .links a{
  color:var(--text-muted);
  font-size:.75rem;
  font-weight:600;
  letter-spacing:.04em;
  text-transform:uppercase;
  transition:color .15s ease,opacity .15s ease;
  opacity:.92;
}
nav .links a:hover{
  color:var(--text);
  opacity:1;
}
nav .nav-actions{margin-left:auto;display:flex;align-items:center;gap:.55rem}
nav .wolf-link{
  color:var(--accent-wolf);
  font-size:.78rem;
  font-weight:700;
  display:flex;
  align-items:center;
  gap:.28rem;
}
nav .nav-toggle{display:none}

/* ── Cards ── */
.card{
  background:linear-gradient(180deg,rgba(16,24,37,.93),rgba(13,20,31,.93));
  border-radius:var(--radius);
  box-shadow:var(--shadow-soft);
  padding:1.45rem;
  margin-bottom:1.5rem;
  border:1px solid var(--border-soft);
}
.cards-grid{
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(190px,1fr));
  gap:1.05rem;
  margin:1.1rem 0 .2rem;
}
.stat-card{
  position:relative;
  background:linear-gradient(165deg,rgba(15,23,35,.96),rgba(12,19,30,.92));
  border-radius:12px;
  padding:1.25rem 1.1rem;
  text-align:center;
  border:1px solid var(--border-soft);
  box-shadow:0 8px 18px rgba(0,0,0,.22);
  overflow:hidden;
}
.stat-card::before{
  content:'';
  position:absolute;
  inset:0 auto auto 0;
  width:100%;
  height:2px;
  background:linear-gradient(90deg,var(--accent-soft),transparent 72%);
  opacity:.5;
}
.stat-card .val{
  font-family:var(--font-display);
  font-size:2.2rem;
  font-weight:700;
  color:#6bb6ff;
  letter-spacing:.01em;
}
.stat-card .lbl{
  font-size:.74rem;
  color:var(--text-muted);
  margin-top:.26rem;
  text-transform:uppercase;
  letter-spacing:.075em;
}

/* ── Section headings ── */
h1{
  font-family:var(--font-display);
  font-size:2.14rem;
  font-weight:800;
  margin-bottom:.65rem;
  color:#f2f6ff;
  letter-spacing:.01em;
}
h2{
  font-family:var(--font-display);
  font-size:1.36rem;
  font-weight:750;
  margin-bottom:.86rem;
  color:#e9f1ff;
  padding-bottom:.52rem;
  border-bottom:1px solid var(--border);
  letter-spacing:.01em;
}
h3{
  font-family:var(--font-display);
  font-size:1.05rem;
  font-weight:700;
  margin:1.08rem 0 .55rem;
  color:#e7f0ff;
  letter-spacing:.01em;
}
.meta-bar{
  font-size:.83rem;
  color:var(--text-muted);
  margin-bottom:1rem;
  line-height:1.6;
}

/* ── Tables ── */
table{
  border-collapse:collapse;
  width:100%;
  font-size:.82rem;
  background:transparent;
}
th,td{
  padding:.48rem .64rem;
  border:1px solid var(--border-soft);
  text-align:left;
  vertical-align:top;
}
th{
  background:linear-gradient(180deg,rgba(31,42,60,.98),rgba(24,34,49,.98));
  font-weight:700;
  white-space:nowrap;
  position:sticky;
  top:0;
  z-index:1;
  color:#d7e7ff;
  letter-spacing:.015em;
}
th.sort{cursor:pointer;user-select:none}
th.sort:hover{background:linear-gradient(180deg,rgba(38,52,76,.98),rgba(29,40,60,.98))}
th.sort::after{content:'';margin-left:4px;opacity:.3}
th.asc::after{content:' \\25B2'}
th.desc::after{content:' \\25BC'}
tbody tr:hover{background:rgba(98,183,255,.07)}
.table-scroll{
  overflow-x:auto;
  margin-bottom:1rem;
  border-radius:var(--radius-sm);
  border:1px solid var(--border-soft);
  box-shadow:0 9px 21px rgba(0,0,0,.2);
  background:linear-gradient(180deg,rgba(13,20,31,.95),rgba(11,17,28,.94));
  width:100%;
}
.table-scroll table{border:none}
.table-scroll td,.table-scroll th{border-left:none;border-right:none}

/* ── Score colours (dark) ── */
.s3{background:var(--g3);color:var(--g3t);font-weight:700;text-align:center}
.s2{background:var(--g2);color:var(--g2t);font-weight:700;text-align:center}
.s1{background:var(--g1);color:var(--g1t);font-weight:700;text-align:center}
.s0{background:var(--g0);color:var(--g0t);font-weight:700;text-align:center}
.sna{background:var(--na);color:var(--nat);text-align:center}

/* ── Matrix ── */
.matrix th,.matrix td{
  padding:.36rem .54rem;
  text-align:center;
  min-width:54px;
  font-size:.77rem;
}
.matrix .feat{text-align:left;font-weight:600;min-width:134px;white-space:nowrap}
.matrix .feat a{color:#7bc4ff}
.tier-row td{
  background:linear-gradient(90deg,#101a29,#0f1a2b) !important;
  color:#c9daf4;
  font-weight:700;
  font-size:.75rem;
  padding:.34rem .62rem;
  letter-spacing:.07em;
  text-transform:uppercase;
  border-bottom:1px solid var(--border);
}

/* ── Details ── */
details{
  border:1px solid var(--border-soft);
  border-radius:12px;
  margin-bottom:.7rem;
  background:linear-gradient(180deg,rgba(15,23,35,.96),rgba(12,19,31,.92));
  box-shadow:0 8px 20px rgba(0,0,0,.19);
}
details summary{
  padding:.78rem 1rem;
  cursor:pointer;
  font-weight:650;
  font-size:.86rem;
  list-style:none;
  display:flex;
  align-items:center;
  gap:.5rem;
  color:#d8e6ff;
}
details summary::before{
  content:'\\25B6';
  font-size:.62rem;
  color:var(--text-muted);
  transition:transform .2s ease;
}
details[open] summary::before{transform:rotate(90deg)}
details[open] summary{border-bottom:1px solid var(--border-soft)}
details .content{padding:.8rem 1rem}
.badge{
  display:inline-block;
  padding:.13rem .52rem;
  border-radius:999px;
  font-size:.67rem;
  font-weight:700;
  letter-spacing:.04em;
}
.badge-t0{background:#0f2740;color:#69bdff}
.badge-t1{background:#2b200a;color:#ffb048}
.badge-t2{background:#221136;color:#ca9cff}
.badge-t3{background:#0e2b1d;color:#88ddb0}

/* ── Test case table ── */
.tc-table td,.tc-table th{font-size:.74rem;padding:.3rem .5rem}
.tc-table .pass{color:#76d8a2}
.tc-table .fail{color:#ff7683;font-weight:600}
code.val{
  font-family:var(--font-mono);
  font-size:.71rem;
  background:#172439;
  color:#dbe8ff;
  padding:.1rem .34rem;
  border-radius:4px;
  word-break:break-all;
}

/* ── Perf breakdown bar ── */
.bbar{display:flex;height:18px;border-radius:4px;overflow:hidden;width:100%;min-width:120px}
.bbar span{
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:8px;
  color:#08111d;
  overflow:hidden;
  white-space:nowrap;
  padding:0 3px;
  font-weight:700;
}
.bbar span:nth-child(1){background:#4c97ff}
.bbar span:nth-child(2){background:#9f7cff}
.bbar span:nth-child(3){background:#42c9df}
.bbar span:nth-child(4){background:#45d299}

/* ── SVG container ── */
.svg-wrap{
  overflow-x:auto;
  margin:1rem 0;
  width:100%;
  border:1px solid var(--border-soft);
  border-radius:var(--radius-sm);
  background:linear-gradient(180deg,rgba(13,20,31,.95),rgba(10,16,27,.93));
  box-shadow:0 8px 22px rgba(0,0,0,.18);
  padding:.4rem;
}
.svg-wrap svg{
  display:block;
  width:100%;
  min-width:100%;
  max-width:none;
  height:auto;
}

/* ── Filter ── */
.filter-box{margin-bottom:.75rem}
.filter-box input{
  padding:.46rem .8rem;
  border:1px solid var(--border);
  border-radius:9px;
  font-size:.84rem;
  width:300px;
  outline:none;
  background:#121d2d;
  color:#dce9ff;
}
.filter-box input:focus{
  border-color:#67beff;
  box-shadow:0 0 0 2px rgba(98,183,255,.24);
}
.filter-box input::placeholder{color:#8192aa}
.sr-only{
  position:absolute;
  width:1px;
  height:1px;
  padding:0;
  margin:-1px;
  overflow:hidden;
  clip:rect(0,0,0,0);
  white-space:nowrap;
  border:0;
}

/* ── Memory bars ── */
.mem-group{margin-bottom:1.5rem}
.mem-group h3{font-size:.95rem;margin-bottom:.6rem}
.mem-scale-label{
  font-size:.8rem;
  font-weight:700;
  color:var(--text-muted);
  margin:.8rem 0 .4rem;
  text-transform:uppercase;
  letter-spacing:.09em;
}
.mem-bar-row{display:flex;align-items:center;gap:.5rem;margin-bottom:.35rem;font-size:.78rem}
.mem-bar-label{
  width:130px;
  text-align:right;
  font-weight:600;
  flex-shrink:0;
  overflow:hidden;
  text-overflow:ellipsis;
  white-space:nowrap;
}
.mem-bar-track{
  flex:1;
  height:22px;
  background:#172336;
  border-radius:4px;
  position:relative;
  overflow:hidden;
  min-width:100px;
}
.mem-bar-fill{
  height:100%;
  border-radius:4px;
  display:flex;
  align-items:center;
  padding-left:6px;
  font-size:.7rem;
  font-weight:700;
  color:#f2f6ff;
  white-space:nowrap;
  transition:width .3s ease;
}
.mem-bar-fill.rust{background:linear-gradient(90deg,#f08235,#ffb774)}
.mem-bar-fill.python{background:linear-gradient(90deg,#4f93ff,#80b6ff)}
.mem-bar-val{
  width:80px;
  text-align:right;
  font-family:var(--font-mono);
  font-size:.75rem;
  flex-shrink:0;
}
.mem-bar-tm{
  width:80px;
  text-align:right;
  font-family:var(--font-mono);
  font-size:.72rem;
  color:var(--text-muted);
  flex-shrink:0;
}
.mem-legend{display:flex;gap:1.2rem;margin:.75rem 0 .5rem;font-size:.78rem}
.mem-legend span{display:flex;align-items:center;gap:.3rem}
.mem-legend .dot{width:10px;height:10px;border-radius:2px;display:inline-block}
.mem-legend .dot.rust{background:#f08235}
.mem-legend .dot.python{background:#4f93ff}
.mem-insight{
  background:#221805;
  border:1px solid #50330b;
  border-radius:10px;
  padding:.84rem 1rem;
  font-size:.82rem;
  color:#ffc788;
  margin-bottom:1rem;
}

/* ── WolfXL highlight ── */
.wolfxl-col{background:rgba(255,142,60,.11)}
.wolfxl-hdr{
  background:linear-gradient(145deg,#2e1508,#3b1a08) !important;
  border-bottom:3px solid var(--accent-wolf);
  position:relative;
}
.wolfxl-hdr div:first-child{color:#ffb06a;font-weight:800;font-size:.88rem}
.wolfxl-hdr div:first-child::after{content:' \\1F3C6';font-size:.7rem}
.wolfxl-row{
  border-left:4px solid var(--accent-wolf);
  background:linear-gradient(90deg,rgba(255,142,60,.15),rgba(255,142,60,.05)) !important;
}
.wolfxl-row td:first-child{font-weight:800;color:#ffb06a}
.wolfxl-badge{
  display:inline-block;
  background:linear-gradient(140deg,#ff8e3c,#f4b15e);
  color:#1f1107;
  font-size:.6rem;
  font-weight:800;
  padding:.15rem .45rem;
  border-radius:999px;
  margin-left:.4rem;
  vertical-align:middle;
  letter-spacing:.04em;
  text-transform:uppercase;
}
.wolfxl-hero{
  background:linear-gradient(145deg,rgba(45,20,7,.94),rgba(34,16,6,.92));
  border:1px solid #d0712d;
  box-shadow:0 11px 28px rgba(0,0,0,.24),inset 0 0 0 1px rgba(255,196,118,.15);
  border-radius:14px;
  padding:1.24rem 1.52rem;
  margin-bottom:1.5rem;
  display:flex;
  align-items:center;
  gap:1rem;
}
.wolfxl-hero .wolf-icon{font-size:2rem}
.wolfxl-hero .wolf-text h3{
  margin:0;
  font-family:var(--font-display);
  font-size:1.12rem;
  color:#ffba73;
  font-weight:700;
}
.wolfxl-hero .wolf-text p{
  margin:.18rem 0 0;
  font-size:.82rem;
  color:#ffd0a1;
  line-height:1.48;
}
.wolfxl-hero .wolf-text a{color:#ffb06a;font-weight:700}

/* ── Interactive chart containers ── */
.chart-maximize-wrap{position:relative;margin-bottom:1rem}
.chart-maximize-btn{
  position:absolute;
  top:8px;
  right:8px;
  z-index:10;
  background:rgba(13,22,35,.94);
  border:1px solid var(--border);
  border-radius:8px;
  color:#9db0ca;
  font-size:.96rem;
  padding:.24rem .52rem;
  cursor:pointer;
  transition:all .15s;
  line-height:1;
}
.chart-maximize-btn:hover{
  background:#1a2c42;
  color:#e7f1ff;
  border-color:#71beff;
}
.plotly-chart{
  width:100%;
  overflow-x:auto;
  border:1px solid var(--border-soft);
  border-radius:var(--radius-sm);
  background:linear-gradient(180deg,rgba(12,19,30,.96),rgba(10,16,27,.94));
  box-shadow:0 8px 20px rgba(0,0,0,.18);
}
.plotly-chart .js-plotly-plot{width:100%!important}

/* ── Chart maximize modal ── */
.chart-modal-overlay{
  display:none;
  position:fixed;
  inset:0;
  z-index:9999;
  background:rgba(3,6,12,.92);
  align-items:center;
  justify-content:center;
  padding:2vh 2vw;
}
.chart-modal-overlay.active{display:flex}
.chart-modal-content{
  width:96vw;
  height:92vh;
  background:#0c1421;
  border-radius:12px;
  border:1px solid var(--border);
  position:relative;
  overflow:hidden;
}
.chart-modal-close{
  position:absolute;
  top:10px;
  right:14px;
  z-index:10;
  background:#111d2d;
  border:1px solid var(--border);
  border-radius:8px;
  color:#9db0cd;
  font-size:1.2rem;
  padding:.3rem .7rem;
  cursor:pointer;
  line-height:1;
}
.chart-modal-close:hover{background:#17283d;color:#edf5ff;border-color:#ff7c86}
.chart-modal-body{width:100%;height:100%;overflow:auto}

/* ── Misc ── */
.btn{
  display:inline-block;
  padding:.36rem .88rem;
  border:1px solid var(--border);
  border-radius:9px;
  font-size:.74rem;
  cursor:pointer;
  background:#132032;
  color:#a8bbd4;
  font-weight:600;
  letter-spacing:.045em;
  text-transform:uppercase;
}
.btn:hover{background:#1d3048;color:#ecf4ff}
footer{
  text-align:center;
  padding:2rem;
  color:var(--text-muted);
  font-size:.76rem;
  letter-spacing:.08em;
  text-transform:uppercase;
}
.flex-bar{display:flex;align-items:center;gap:.75rem;flex-wrap:wrap;margin-bottom:.78rem}

@media(max-width:900px){
  body{font-size:15px;line-height:1.58}
  nav{padding:.62rem .95rem;gap:.58rem;flex-wrap:wrap;align-items:flex-start}
  nav .brand{font-size:1.04rem}
  nav .nav-actions{order:2;margin-left:auto;gap:.4rem}
  nav .wolf-link{font-size:.93rem}
  nav .nav-toggle{
    display:inline-flex;
    align-items:center;
    justify-content:center;
    min-width:64px;
    padding:.25rem .55rem;
    border:1px solid var(--border);
    border-radius:8px;
    background:#111d2e;
    color:#d3e0f2;
    font-size:.78rem;
    font-weight:700;
    letter-spacing:.03em;
    cursor:pointer;
  }
  nav .links{display:none;width:100%;order:3}
  nav.menu-open .links{
    display:grid;
    grid-template-columns:repeat(2,minmax(0,1fr));
    gap:.42rem .72rem;
    border-top:1px solid var(--border-soft);
    padding-top:.62rem;
  }
  nav .links a{
    font-size:.84rem;
    line-height:1.25;
    padding:.36rem .18rem;
    border-radius:8px;
    background:rgba(255,255,255,.02);
  }
  .card{padding:1.2rem}
  .cards-grid{grid-template-columns:1fr 1fr}
  .container{padding:1rem .95rem}
  .filter-box input{width:100%}
  .meta-bar{font-size:.82rem}
  h2{font-size:1.26rem}
  h1{font-size:1.84rem}
}
@media(max-width:620px){
  body{font-size:15px;line-height:1.62}
  nav.menu-open .links{grid-template-columns:1fr}
  nav .links a{
    font-size:.9rem;
    letter-spacing:.02em;
    text-transform:none;
    font-weight:700;
  }
  h1{font-size:1.72rem}
  h2{font-size:1.2rem}
  h3{font-size:1rem}
  .cards-grid{grid-template-columns:1fr}
  .stat-card .val{font-size:2.05rem}
  .stat-card .lbl{font-size:.79rem}
  .wolfxl-hero{
    padding:1rem 1.1rem;
    gap:.78rem;
    align-items:flex-start;
  }
  .wolfxl-hero .wolf-icon{font-size:1.7rem}
  .wolfxl-hero .wolf-text h3{font-size:1.04rem}
  .wolfxl-hero .wolf-text p{font-size:.88rem}
  .meta-bar{font-size:.88rem}
  .meta-bar strong{font-weight:800}
  details summary{
    font-size:.9rem;
    line-height:1.35;
    align-items:flex-start;
  }
  .btn{font-size:.78rem}
  .table-scroll{border-radius:9px}
}
@media(min-width:1800px){
  body{font-size:16px}
  .container{padding:2rem 4rem}
  nav{padding:.8rem 4rem;gap:1.8rem}
  nav .brand{font-size:1.3rem}
  nav .links a{font-size:.85rem}
  h1{font-size:2.36rem}
  h2{font-size:1.68rem}
  h3{font-size:1.25rem}
  .stat-card .val{font-size:2.86rem}
  .stat-card .lbl{font-size:.9rem}
  .stat-card{padding:1.6rem}
  table{font-size:.92rem}
  th,td{padding:.56rem .76rem}
  .matrix th,.matrix td{padding:.45rem .65rem;min-width:64px;font-size:.88rem}
  .matrix .feat{min-width:160px}
  details summary{font-size:1rem;padding:.85rem 1.2rem}
  .meta-bar{font-size:.92rem}
  .wolfxl-hero{padding:1.5rem 2rem}
  .wolfxl-hero .wolf-text h3{font-size:1.22rem}
  .wolfxl-hero .wolf-text p{font-size:.95rem}
  .filter-box input{width:360px;font-size:.95rem;padding:.52rem .94rem}
}
@media(min-width:2600px){
  body{font-size:18px}
  .container{padding:2.5rem 5rem}
  nav{padding:.84rem 5rem}
  nav .brand{font-size:1.52rem}
  h1{font-size:2.65rem}
  h2{font-size:1.95rem}
  h3{font-size:1.45rem}
  .stat-card .val{font-size:3.35rem}
  .stat-card .lbl{font-size:1.05rem}
  table{font-size:1rem}
  th,td{padding:.62rem .88rem}
  .matrix th,.matrix td{padding:.5rem .75rem;min-width:76px;font-size:.95rem}
  .matrix .feat{min-width:180px}
  details summary{font-size:1.1rem}
}
"""

# ====================================================================
#  JavaScript
# ====================================================================

_JS = r"""
/* Mobile nav toggle */
const nav=document.querySelector('nav');
const navToggle=nav?.querySelector('.nav-toggle');
const closeNav=()=>{
  if(!nav||!navToggle)return;
  nav.classList.remove('menu-open');
  navToggle.setAttribute('aria-expanded','false');
  navToggle.textContent='Menu';
};
if(nav && navToggle){
  navToggle.addEventListener('click',()=>{
    const open=nav.classList.toggle('menu-open');
    navToggle.setAttribute('aria-expanded',String(open));
    navToggle.textContent=open?'Close':'Menu';
  });
  window.addEventListener('resize',()=>{
    if(window.innerWidth>768){closeNav();}
  });
}
/* Table sorting */
document.querySelectorAll('th.sort').forEach(th=>{
  th.addEventListener('click',()=>{
    const table=th.closest('table'),tbody=table.querySelector('tbody');
    if(!tbody)return;
    const rows=Array.from(tbody.rows),ci=th.cellIndex;
    const num=th.dataset.type==='n';
    const asc=!th.classList.contains('asc');
    th.closest('tr').querySelectorAll('th').forEach(h=>{h.classList.remove('asc','desc')});
    th.classList.add(asc?'asc':'desc');
    rows.sort((a,b)=>{
      let va=a.cells[ci]?.dataset.v||a.cells[ci]?.textContent.trim()||'';
      let vb=b.cells[ci]?.dataset.v||b.cells[ci]?.textContent.trim()||'';
      if(num){va=parseFloat(va)||0;vb=parseFloat(vb)||0}
      return asc?(va>vb?1:va<vb?-1:0):(va<vb?1:va>vb?-1:0);
    });
    rows.forEach(r=>tbody.appendChild(r));
  });
});
/* Filter */
document.querySelectorAll('.filter-input').forEach(inp=>{
  inp.addEventListener('input',()=>{
    const v=inp.value.toLowerCase();
    const tgt=document.getElementById(inp.dataset.target);
    if(!tgt)return;
    tgt.querySelectorAll('tbody tr').forEach(r=>{
      r.style.display=r.textContent.toLowerCase().includes(v)?'':'none';
    });
  });
});
/* Expand / Collapse all */
document.querySelectorAll('.toggle-all').forEach(btn=>{
  btn.addEventListener('click',()=>{
    const sec=btn.closest('section');
    const dets=sec.querySelectorAll('details.expandable');
    const open=Array.from(dets).every(d=>d.open);
    dets.forEach(d=>d.open=!open);
    btn.textContent=open?'Expand All':'Collapse All';
  });
});
/* Chart maximize modal — all content is same-origin Plotly output, no untrusted data */
(function(){
  var overlay=document.createElement('div');
  overlay.className='chart-modal-overlay';
  var content=document.createElement('div');
  content.className='chart-modal-content';
  var closeBtn=document.createElement('button');
  closeBtn.className='chart-modal-close';
  closeBtn.title='Close';
  closeBtn.textContent='\u00d7';
  var modalBody=document.createElement('div');
  modalBody.className='chart-modal-body';
  content.appendChild(closeBtn);
  content.appendChild(modalBody);
  overlay.appendChild(content);
  document.body.appendChild(overlay);
  var closeModal=function(){
    overlay.classList.remove('active');
    while(modalBody.firstChild)modalBody.removeChild(modalBody.firstChild);
  };
  closeBtn.addEventListener('click',closeModal);
  overlay.addEventListener('click',function(e){if(e.target===overlay)closeModal();});
  document.addEventListener('keydown',function(e){if(e.key==='Escape')closeModal();});
  document.querySelectorAll('.chart-maximize-btn').forEach(function(btn){
    btn.addEventListener('click',function(){
      var wrap=btn.closest('.chart-maximize-wrap');
      var chart=wrap?wrap.querySelector('.plotly-chart'):null;
      if(!chart)return;
      while(modalBody.firstChild)modalBody.removeChild(modalBody.firstChild);
      /* Rebuild Plotly chart in modal with full interactivity */
      var srcPlot=chart.querySelector('.js-plotly-plot');
      if(srcPlot&&srcPlot.data&&srcPlot.layout&&window.Plotly){
        var plotDiv=document.createElement('div');
        plotDiv.style.width='100%';
        plotDiv.style.height='100%';
        modalBody.appendChild(plotDiv);
        overlay.classList.add('active');
        var newLayout=Object.assign({},srcPlot.layout,{autosize:true,
          height:content.clientHeight-20,width:content.clientWidth-20});
        window.Plotly.newPlot(plotDiv,srcPlot.data,newLayout,
          {displayModeBar:true,displaylogo:false,responsive:true});
      }else{
        /* Fallback: clone static content */
        modalBody.appendChild(chart.cloneNode(true));
        overlay.classList.add('active');
      }
    });
  });
})();
/* Smooth scroll */
document.querySelectorAll('nav a[href^="#"]').forEach(a=>{
  a.addEventListener('click',e=>{
    const href=a.getAttribute('href');
    if(!href) return;
    const el=document.querySelector(href);
    if(!el) return;
    e.preventDefault();
    el.scrollIntoView({behavior:'smooth',block:'start'});
    if(history && history.pushState){
      history.pushState(null,'',href);
    }else{
      location.hash=href;
    }
    if(window.matchMedia('(max-width: 768px)').matches){
      closeNav();
    }
  });
});
"""

# ====================================================================
#  Section renderers
# ====================================================================


def _section_nav(*, has_memory: bool = False) -> str:
    links = [
        ("#overview", "Overview"),
        ("#radar", "Strength Profiles"),
        ("#matrix", "Score Matrix"),
        ("#scatter", "Scatter Plots"),
        ("#comparison", "Comparison"),
        ("#features", "Features"),
        ("#perf", "Performance"),
    ]
    if has_memory:
        links.append(("#memory", "Memory"))
    links.append(("#diag", "Diagnostics"))
    link_html = "".join(f'<a href="{h}">{t}</a>' for h, t in links)
    return (
        f'<nav>'
        f'<div class="brand">ExcelBench</div>'
        f'<div class="links" id="site-nav-links">{link_html}</div>'
        f'<div class="nav-actions">'
        f'<a class="wolf-link" href="https://github.com/SynthGL/wolfxl" target="_blank" '
        f'rel="noopener noreferrer">\U0001F43A WolfXL</a>'
        f'<button type="button" class="nav-toggle" aria-controls="site-nav-links" '
        f'aria-expanded="false" aria-label="Toggle navigation menu">Menu</button>'
        f'</div></nav>'
    )


def _section_overview(fidelity: dict[str, Any], perf: dict[str, Any] | None) -> str:
    meta = fidelity.get("metadata", {})
    libs = fidelity.get("libraries", {})
    results = fidelity.get("results", [])

    all_features = sorted({e["feature"] for e in results})
    total_libs = len(libs)
    total_feats = len(all_features)

    # Compute overall pass rate
    total_pass = total_tests = 0
    for entry in results:
        for tc in entry.get("test_cases", {}).values():
            if not isinstance(tc, dict):
                continue
            for op in ("read", "write"):
                if op in tc:
                    total_tests += 1
                    if tc[op].get("passed"):
                        total_pass += 1
    pass_rate = (total_pass / total_tests * 100) if total_tests else 0

    # Green count
    green = sum(
        1 for e in results
        if max((s for s in [e["scores"].get("read"), e["scores"].get("write")]
                if s is not None), default=-1) == 3
    )
    total_scored = len(results)

    cards = [
        (str(total_libs), "Libraries Tested"),
        (str(total_feats), "Features Scored"),
        (f"{pass_rate:.0f}%", "Avg Pass Rate"),
        (f"{green}/{total_scored}", "Score\u20033 Results"),
    ]

    if perf:
        perf_meta = perf.get("metadata", {}).get("config", {})
        cards.append((str(perf_meta.get("iters", "?")), "Perf Iterations"))

    cards_html = "".join(
        f'<div class="stat-card"><div class="val">{v}</div>'
        f'<div class="lbl">{lbl}</div></div>'
        for v, lbl in cards
    )

    # WolfXL hero stats
    wolf_green = sum(
        1 for e in results if e["library"] == "wolfxl"
        and max((s for s in [e["scores"].get("read"), e["scores"].get("write")]
                 if s is not None), default=-1) == 3
    )
    wolf_scored = sum(1 for e in results if e["library"] == "wolfxl")
    wolf_hero = ""
    if wolf_scored > 0:
        # Compute throughput headline from perf data if available
        wolf_perf_note = ""
        if perf:
            for e in perf.get("results", []):
                if e["library"] == "wolfxl" and "read" in e.get("perf", {}):
                    p = e["perf"]["read"]
                    oc = p.get("op_count")
                    p50 = (p.get("wall_ms") or {}).get("p50")
                    if p50 and p50 > 0:
                        rate = (oc * 1000.0 / p50) if oc else (1000.0 / p50)
                        if rate > 100_000:
                            wolf_perf_note = (
                                f" &middot; Up to <b>{rate / 1000:.0f}K cells/s</b> read throughput"
                            )
                        break
        wolf_hero = (
            f'<div class="wolfxl-hero">'
            f'<div class="wolf-icon">\U0001F43A</div>'
            f'<div class="wolf-text">'
            f'<h3>WolfXL &mdash; Top-Ranked Across Fidelity &amp; Performance</h3>'
            f'<p><b>{wolf_green}/{wolf_scored}</b> features at full fidelity (score 3)'
            f'{wolf_perf_note}. '
            f'Hybrid Rust+Python architecture for maximum compatibility and speed. '
            f'<a href="https://github.com/SynthGL/wolfxl" '
            f'style="color:#ea580c;font-weight:600">View on GitHub &rarr;</a></p>'
            f'</div></div>'
        )

    return (
        f'<section id="overview" class="container">'
        f'<h1>ExcelBench Dashboard</h1>'
        f'<div class="meta-bar">'
        f'Profile: <b>{_esc(meta.get("profile", "xlsx"))}</b> &middot; '
        f'Platform: {_esc(meta.get("platform", "?"))} &middot; '
        f'Excel: {_esc(meta.get("excel_version", "?"))} &middot; '
        f'Date: {_esc(meta.get("run_date", "?")[:10])}'
        f'</div>'
        f'<div style="margin-top:.45rem;color:#acb4c0;font-size:.82rem">'
        f'Score rubric: 3 = all basic + edge pass; 2 = all basic pass with edge failures; '
        f'1 = at least one basic pass (not all); 0 = no basic pass.'
        f'</div>'
        f'{wolf_hero}'
        f'<div class="cards-grid">{cards_html}</div>'
        f'</section>'
    )


def _section_radar(
    fidelity: dict[str, Any],
    perf: dict[str, Any] | None,
) -> str:
    """Render a 5-axis spider chart comparing top libraries."""
    import plotly.graph_objects as go

    from excelbench.results.scatter import (
        _DARK_BG,
        _DARK_GRID,
        _DARK_TEXT,
        _DARK_TEXT2,
        _FALLBACK_COLOR,
        _LIB_COLORS,
        _LIB_SHORT,
    )
    from excelbench.results.scatter_interactive import _PLOTLY_CONFIG

    data = _compute_radar_data(fidelity, perf)
    if not data:
        return ""

    fig = go.Figure()
    for item in data:
        lib = item["library"]
        axes = item["axes"]
        vals = item["values"]
        # Close the polygon by repeating the first point
        theta = axes + [axes[0]]
        r = vals + [vals[0]]
        color = _LIB_COLORS.get(lib, _FALLBACK_COLOR)
        display_name = _LIB_SHORT.get(lib, lib)
        is_wolfxl = lib == "wolfxl"

        fig.add_trace(go.Scatterpolar(
            r=r,
            theta=theta,
            fill="toself",
            fillcolor=color,
            opacity=0.20 if is_wolfxl else 0.15,
            line=dict(color=color, width=3 if is_wolfxl else 1.5),
            name=display_name,
            hovertemplate=(
                f"<b>{display_name}</b><br>"
                "%{theta}: %{r:.0f}<extra></extra>"
            ),
        ))

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor=_DARK_BG,
        plot_bgcolor=_DARK_BG,
        font=dict(color=_DARK_TEXT, family="system-ui, sans-serif"),
        polar=dict(
            bgcolor=_DARK_BG,
            radialaxis=dict(
                visible=True,
                range=[0, 100],
                gridcolor=_DARK_GRID,
                tickfont=dict(size=10, color=_DARK_TEXT2),
            ),
            angularaxis=dict(
                gridcolor=_DARK_GRID,
                tickfont=dict(size=13, color=_DARK_TEXT),
            ),
        ),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.1,
            xanchor="center",
            x=0.5,
            font=dict(size=13, color=_DARK_TEXT),
        ),
        margin=dict(l=80, r=80, t=40, b=80),
        height=520,
        showlegend=True,
    )

    chart_html = fig.to_html(
        full_html=False, include_plotlyjs=True, config=_PLOTLY_CONFIG,
    )

    return (
        '<section id="radar" class="container">'
        '<h2>Library Strength Profiles</h2>'
        '<p style="font-size:.8rem;color:var(--text2);margin-bottom:.5rem">'
        '5-axis comparison: fidelity, read/write speed, feature coverage, '
        'and capability breadth. All axes normalised 0\u2013100. '
        'Speed baseline: openpyxl\u00a0=\u00a050.</p>'
        '<div class="chart-maximize-wrap">'
        '<button class="chart-maximize-btn" title="Maximize"'
        ' aria-label="Maximize chart">\u26f6</button>'
        f'<div class="plotly-chart">{chart_html}</div></div>'
        '</section>'
    )


def _section_matrix(fidelity: dict[str, Any]) -> str:
    results = fidelity.get("results", [])
    libs_info = fidelity.get("libraries", {})

    # (feat, lib) -> (read, write)
    score_map: dict[tuple[str, str], tuple[int | None, int | None]] = {}
    all_feats: set[str] = set()
    all_libs: set[str] = set()
    for entry in results:
        f, lib = entry["feature"], entry["library"]
        s = entry.get("scores", {})
        score_map[(f, lib)] = (s.get("read"), s.get("write"))
        all_feats.add(f)
        all_libs.add(lib)

    features = [f for f in _FEATURE_ORDER if f in all_feats]
    for f in sorted(all_feats):
        if f not in features:
            features.append(f)

    # Sort libs by green count desc
    def _green(lib: str) -> int:
        return sum(
            1 for feat in features
            if max((x for x in score_map.get((feat, lib), (None, None)) if x is not None),
                   default=-1) == 3
        )
    libs = sorted(all_libs, key=lambda x: (-_green(x), x))

    rows: list[str] = []
    rows.append('<section id="matrix" class="container"><h2>Score Matrix</h2>')
    rows.append('<p style="font-size:.8rem;color:var(--text2);margin-bottom:.5rem">'
                'Best of read/write. Hover for R/W breakdown.</p>')
    rows.append('<div class="table-scroll"><table class="matrix">')

    # Header
    rows.append("<thead><tr><th class='feat'>Feature</th>")
    for lib in libs:
        cap = _cap_label(libs_info.get(lib, {}).get("capabilities", []))
        hdr_cls = " wolfxl-hdr" if lib == "wolfxl" else ""
        rows.append(f"<th class='{hdr_cls.strip()}'><div>{_esc(lib)}</div>"
                    f"<div style='font-size:.65rem;color:var(--text2)'>"
                    f"{cap}</div></th>")
    rows.append("</tr></thead><tbody>")

    current_tier = -1
    for feat in features:
        tier = _TIER_MAP.get(feat, -1)
        if tier != current_tier:
            current_tier = tier
            tname = _TIER_NAMES.get(tier, f"Tier {tier}")
            rows.append(f'<tr class="tier-row"><td colspan="{len(libs) + 1}">{tname}</td></tr>')

        label = _FEATURE_LABELS.get(feat, feat)
        rows.append(f'<tr><td class="feat"><a href="#feat-{feat}">{label}</a></td>')
        for lib in libs:
            rs, ws = score_map.get((feat, lib), (None, None))
            best = max((x for x in [rs, ws] if x is not None), default=None)
            cls = _score_cls(best)
            if lib == "wolfxl":
                cls += " wolfxl-col"
            tip = f"Read: {_score_label(rs)} / Write: {_score_label(ws)}"
            rows.append(f'<td class="{cls}" title="{tip}">{_score_label(best)}</td>')
        rows.append("</tr>")

    rows.append("</tbody></table></div></section>")
    return "\n".join(rows)


def _section_scatter(
    fidelity: dict[str, Any],
    perf: dict[str, Any] | None,
    heatmap_svg: str | None,
    scatter_svgs: dict[str, str] | None = None,
) -> str:
    if not perf and not heatmap_svg and not scatter_svgs:
        return ""

    parts = ['<section id="scatter" class="container"><h2>Fidelity vs. Throughput</h2>']

    # Interactive Plotly scatter charts (preferred when perf data exists)
    if perf:
        from excelbench.results.scatter_interactive import (
            render_interactive_scatter_features_from_data,
            render_interactive_scatter_tiers_from_data,
        )

        tiers_html = render_interactive_scatter_tiers_from_data(fidelity, perf)
        features_html = render_interactive_scatter_features_from_data(fidelity, perf)

        parts.append(
            '<h3>By Feature Group</h3>'
            f'<div class="chart-maximize-wrap">'
            f'<button class="chart-maximize-btn" title="Maximize"'
            f' aria-label="Maximize chart">\u26f6</button>'
            f'<div class="plotly-chart">{tiers_html}</div></div>'
        )
        parts.append(
            '<h3>Per Feature</h3>'
            f'<div class="chart-maximize-wrap">'
            f'<button class="chart-maximize-btn" title="Maximize"'
            f' aria-label="Maximize chart">\u26f6</button>'
            f'<div class="plotly-chart">{features_html}</div></div>'
        )
    else:
        # Fallback: embed pre-rendered static SVGs when perf data is unavailable
        if scatter_svgs:
            tiers_svg = scatter_svgs.get("scatter_tiers.svg")
            feats_svg = scatter_svgs.get("scatter_features.svg")
            if tiers_svg:
                parts.append(f'<h3>By Feature Group</h3>'
                             f'<div class="svg-wrap">{tiers_svg}</div>')
            if feats_svg:
                parts.append(f'<h3>Per Feature</h3>'
                             f'<div class="svg-wrap">{feats_svg}</div>')

    # Static heatmap SVG (always shown if available)
    if heatmap_svg:
        parts.append(f'<h3>Heatmap</h3><div class="svg-wrap">{heatmap_svg}</div>')

    parts.append("</section>")
    return "\n".join(parts)


def _section_comparison(fidelity: dict[str, Any], perf: dict[str, Any] | None) -> str:
    results = fidelity.get("results", [])
    libs_info = fidelity.get("libraries", {})

    # Per-lib stats
    lib_stats: dict[str, dict[str, Any]] = {}
    for lib, info in libs_info.items():
        cap = _cap_label(info.get("capabilities", []))
        lib_stats[lib] = {
            "cap": cap,
            "modify": modify_mode_label(lib, info.get("capabilities", [])),
            "version": info.get("version", "?"),
            "green": 0,
            "scored": 0,
            "passed": 0,
            "total": 0,
        }

    for entry in results:
        lib = entry["library"]
        if lib not in lib_stats:
            continue
        s = entry.get("scores", {})
        best = max((x for x in [s.get("read"), s.get("write")] if x is not None), default=None)
        if best is not None:
            lib_stats[lib]["scored"] += 1
            if best == 3:
                lib_stats[lib]["green"] += 1
        for tc in entry.get("test_cases", {}).values():
            if not isinstance(tc, dict):
                continue
            for op in ("read", "write"):
                if op in tc:
                    lib_stats[lib]["total"] += 1
                    if tc[op].get("passed"):
                        lib_stats[lib]["passed"] += 1

    # Perf throughput
    lib_tp: dict[str, dict[str, str]] = {}
    if perf:
        perf_lookup: dict[tuple[str, str], dict[str, Any]] = {}
        for e in perf.get("results", []):
            perf_lookup[(e["feature"], e["library"])] = e.get("perf", {})
        for lib in libs_info:
            read_rate = write_rate = raw_read_rate = "\u2014"
            for sc in ("cell_values_10k_bulk_read", "cell_values_10k",
                        "cell_values_1k", "cell_values"):
                p = perf_lookup.get((sc, lib), {}).get("read", {})
                if p and p.get("wall_ms", {}).get("p50"):
                    read_rate = _fmt_rate(p.get("op_count"), p["wall_ms"]["p50"])
                    break
            for sc in ("cell_values_10k_bulk_read_raw", "cell_values_1k_bulk_read_raw"):
                p = perf_lookup.get((sc, lib), {}).get("read", {})
                if p and p.get("wall_ms", {}).get("p50"):
                    raw_read_rate = _fmt_rate(p.get("op_count"), p["wall_ms"]["p50"])
                    break
            for sc in ("cell_values_10k_bulk_write", "cell_values_10k",
                        "cell_values_1k", "cell_values"):
                p = perf_lookup.get((sc, lib), {}).get("write", {})
                if p and p.get("wall_ms", {}).get("p50"):
                    write_rate = _fmt_rate(p.get("op_count"), p["wall_ms"]["p50"])
                    break
            lib_tp[lib] = {"read": read_rate, "raw_read": raw_read_rate, "write": write_rate}

    sorted_libs = sorted(lib_stats.keys(), key=lambda x: (-lib_stats[x]["green"], x))
    has_perf = bool(lib_tp)

    rows: list[str] = []
    rows.append('<section id="comparison" class="container"><h2>Library Comparison</h2>')
    rows.append('<div class="flex-bar">'
                '<div class="filter-box">'
                '<label class="sr-only" for="cmp-filter">Filter libraries</label>'
                '<input id="cmp-filter" type="text" class="filter-input" data-target="cmp-table" '
                'aria-label="Filter libraries" '
                'placeholder="Filter libraries\u2026"></div></div>')
    rows.append('<div class="table-scroll"><table id="cmp-table">')
    rows.append("<thead><tr>"
                "<th class='sort'>Library</th>"
                "<th>Caps</th>"
                "<th>Modify</th>"
                "<th>Version</th>"
                "<th class='sort' data-type='n'>Green</th>"
                "<th class='sort' data-type='n'>Pass Rate</th>")
    if has_perf:
        rows.append("<th class='sort'>Read cells/s</th>"
                    "<th class='sort'>Raw Read cells/s</th>"
                    "<th class='sort'>Write cells/s</th>")
    rows.append("</tr></thead><tbody>")

    for lib in sorted_libs:
        st = lib_stats[lib]
        pr = (st["passed"] / st["total"] * 100) if st["total"] else 0
        row_cls = " class='wolfxl-row'" if lib == "wolfxl" else ""
        wolf_badge = '<span class="wolfxl-badge">Recommended</span>' if lib == "wolfxl" else ""
        rows.append(f"<tr{row_cls}><td><b>{_esc(lib)}</b>{wolf_badge}</td>"
                    f"<td>{st['cap']}</td>"
                    f"<td>{st['modify']}</td>"
                    f"<td style='font-family:monospace;font-size:.75rem'>{_esc(st['version'])}</td>"
                    f"<td data-v='{st['green']}'>{st['green']}/{st['scored']}</td>"
                    f"<td data-v='{pr:.1f}'>{pr:.0f}%</td>")
        if has_perf:
            tp = lib_tp.get(lib, {})
            dash = "\u2014"
            r_val = tp.get('read', dash)
            rr_val = tp.get('raw_read', dash)
            w_val = tp.get('write', dash)
            rows.append(f"<td>{r_val}</td><td>{rr_val}</td><td>{w_val}</td>")
        rows.append("</tr>")

    rows.append("</tbody></table></div></section>")
    return "\n".join(rows)


def _section_features(fidelity: dict[str, Any]) -> str:
    results = fidelity.get("results", [])

    # Group by feature
    by_feat: dict[str, list[dict[str, Any]]] = {}
    for entry in results:
        by_feat.setdefault(entry["feature"], []).append(entry)

    features = [f for f in _FEATURE_ORDER if f in by_feat]
    for f in sorted(by_feat):
        if f not in features:
            features.append(f)

    rows: list[str] = []
    rows.append('<section id="features" class="container">')
    rows.append('<div class="flex-bar"><h2 style="border:none;margin:0;padding:0">'
                'Feature Details</h2>'
                '<button class="btn toggle-all">Expand All</button></div>')

    current_tier = -1
    for feat in features:
        tier = _TIER_MAP.get(feat, -1)
        if tier != current_tier:
            current_tier = tier
            rows.append(f'<h3>{_TIER_NAMES.get(tier, f"Tier {tier}")}</h3>')

        entries = by_feat[feat]
        label = _FEATURE_LABELS.get(feat, feat)
        badge_cls = f"badge-t{tier}" if 0 <= tier <= 3 else ""
        n_libs = len(entries)

        # Feature-level pass rate
        fp = ft = 0
        for e in entries:
            for tc in e.get("test_cases", {}).values():
                if not isinstance(tc, dict):
                    continue
                for op in ("read", "write"):
                    if op in tc:
                        ft += 1
                        if tc[op].get("passed"):
                            fp += 1
        fpr = (fp / ft * 100) if ft else 0

        rows.append(
            f'<details class="expandable" id="feat-{feat}">'
            f'<summary><span>{_esc(label)}</span>'
            f'<span class="badge {badge_cls}">Tier {tier}</span>'
            f'<span style="color:var(--text2);font-size:.78rem;margin-left:auto">'
            f'{n_libs} libraries &middot; {fpr:.0f}% pass rate</span></summary>'
            f'<div class="content">'
        )

        # Per-library summary table
        rows.append('<table><thead><tr>'
                    '<th class="sort">Library</th>'
                    '<th>Read</th><th>Write</th>'
                    '<th class="sort" data-type="n">Pass Rate</th>'
                    '<th>Notes</th></tr></thead><tbody>')

        # Sort by best score desc
        entries_sorted = sorted(
            entries,
            key=lambda e: -max((x for x in [e["scores"].get("read"), e["scores"].get("write")]
                                if x is not None), default=-1),
        )

        for entry in entries_sorted:
            lib = entry["library"]
            rs = entry["scores"].get("read")
            ws = entry["scores"].get("write")
            # Pass rate for this lib on this feature
            lp = lt = 0
            for tc in entry.get("test_cases", {}).values():
                if not isinstance(tc, dict):
                    continue
                for op in ("read", "write"):
                    if op in tc:
                        lt += 1
                        if tc[op].get("passed"):
                            lp += 1
            lpr = (lp / lt * 100) if lt else 0
            notes = _esc(entry.get("notes") or "\u2014")
            feat_row_cls = " class='wolfxl-row'" if lib == "wolfxl" else ""
            rows.append(
                f"<tr{feat_row_cls}><td><b>{_esc(lib)}</b></td>"
                f"<td class='{_score_cls(rs)}'>{_score_label(rs)}</td>"
                f"<td class='{_score_cls(ws)}'>{_score_label(ws)}</td>"
                f"<td data-v='{lpr:.1f}'>{lpr:.0f}%</td>"
                f"<td style='font-size:.75rem'>{notes}</td></tr>"
            )
        rows.append("</tbody></table>")

        # Per-library test-case detail (nested details)
        for entry in entries_sorted:
            lib = entry["library"]
            tcs = entry.get("test_cases", {})
            if not tcs:
                continue
            tp = tf = 0
            for tc in tcs.values():
                if not isinstance(tc, dict):
                    continue
                for op in ("read", "write"):
                    if op in tc:
                        tp += tc[op].get("passed", False)
                        tf += 1

            rows.append(
                f'<details style="margin-top:.4rem">'
                f'<summary style="font-size:.8rem">{_esc(lib)} &mdash; '
                f'{tp}/{tf} test cases passed</summary>'
                f'<div class="content"><table class="tc-table"><thead><tr>'
                f'<th>Test</th><th>Op</th><th>Imp.</th><th>Result</th>'
                f'<th>Expected</th><th>Actual</th></tr></thead><tbody>'
            )

            for tc_id, tc in tcs.items():
                if not isinstance(tc, dict):
                    continue
                for op in ("read", "write"):
                    if op not in tc:
                        continue
                    d = tc[op]
                    passed = d.get("passed", False)
                    imp = d.get("importance", "\u2014")
                    lbl = d.get("label") or tc_id
                    exp = _fmt_val(d.get("expected"))
                    act = _fmt_val(d.get("actual"))
                    pcls = "pass" if passed else "fail"
                    psym = "\u2713" if passed else "\u2717"
                    rows.append(
                        f"<tr><td>{_esc(lbl)}</td><td>{op}</td>"
                        f"<td>{_esc(imp)}</td>"
                        f"<td class='{pcls}'>{psym}</td>"
                        f"<td>{exp}</td><td>{act}</td></tr>"
                    )

            rows.append("</tbody></table></div></details>")

        rows.append("</div></details>")

    rows.append("</section>")
    return "\n".join(rows)


def _section_performance(perf: dict[str, Any] | None) -> str:
    if not perf:
        return (
            '<section id="perf" class="container">'
            '<h2>Performance</h2><p>No perf data.</p></section>'
        )

    meta = perf.get("metadata", {})
    config = meta.get("config", {})
    results = perf.get("results", [])

    # Group by workload
    by_wl: dict[str, list[dict[str, Any]]] = {}
    for entry in results:
        by_wl.setdefault(entry["feature"], []).append(entry)

    # Sort workloads logically
    workload_order = [
        "cell_values_10k", "cell_values_1k",
        "cell_values_10k_bulk_read", "cell_values_1k_bulk_read",
        "cell_values_10k_bulk_write", "cell_values_1k_bulk_write",
        "formulas_10k", "formulas_1k",
        "alignment_1k", "background_colors_1k", "borders_200", "number_formats_1k",
    ]
    workloads = [w for w in workload_order if w in by_wl]
    for w in sorted(by_wl):
        if w not in workloads:
            workloads.append(w)

    rows: list[str] = []
    rows.append('<section id="perf" class="container">'
                '<h2>Performance Benchmarks</h2>')
    rows.append(
        f'<div class="meta-bar">Warmup: {config.get("warmup", "?")} &middot; '
        f'Iterations: {config.get("iters", "?")} &middot; '
        f'Iteration Policy: {_esc(config.get("iteration_policy", "fixed"))} &middot; '
        f'Breakdown: {"Yes" if config.get("breakdown") else "No"} &middot; '
        f'Platform: {_esc(meta.get("platform", "?"))} &middot; '
        f'Python: {_esc(meta.get("python", "?"))}</div>'
    )

    for wl in workloads:
        entries = by_wl[wl]
        n_libs = len(entries)

        rows.append(
            f'<details class="expandable">'
            f'<summary>{_esc(wl)} &mdash; {n_libs} libraries</summary>'
            f'<div class="content">'
        )

        # Collect all ops for this workload
        ops_present: set[str] = set()
        for entry in entries:
            p = entry.get("perf", {})
            for op in ("read", "write"):
                if op in p:
                    ops_present.add(op)

        for op in sorted(ops_present):
            rows.append(f"<h3 style='font-size:.9rem'>{op.title()}</h3>")
            rows.append(
                '<div class="table-scroll"><table><thead><tr>'
                "<th class='sort'>Library</th>"
                "<th class='sort' data-type='n'>p50 (ms)</th>"
                "<th class='sort' data-type='n'>p95 (ms)</th>"
                "<th class='sort' data-type='n'>min (ms)</th>"
                "<th class='sort' data-type='n'>CPU p50</th>"
                "<th class='sort' data-type='n'>RSS (MB)</th>"
                "<th class='sort'>Throughput</th>"
                "<th>Phase Breakdown</th>"
                "</tr></thead><tbody>"
            )

            def _sort_key(e: dict[str, Any]) -> float:
                p = e.get("perf") or {}
                o = p.get(op) or {}
                w = o.get("wall_ms") or {}
                return w.get("p50", 9e9) if isinstance(w, dict) else 9e9

            sorted_entries = sorted(entries, key=_sort_key)

            for entry in sorted_entries:
                lib = entry["library"]
                od = entry.get("perf", {}).get(op)
                if not od:
                    continue
                wall = od.get("wall_ms", {})
                cpu = od.get("cpu_ms", {})
                rss = od.get("rss_peak_mb")
                oc = od.get("op_count")
                rate = _fmt_rate(oc, wall.get("p50"))

                # Breakdown bar
                bd = od.get("breakdown_ms", {})
                bar_html = ""
                if bd:
                    total_bd = sum(v for v in bd.values() if v)
                    if total_bd > 0:
                        bar_parts = []
                        for phase, ms in bd.items():
                            if ms and ms > 0:
                                pct = ms / total_bd * 100
                                bar_parts.append(
                                    f'<span style="width:{pct:.0f}%" '
                                    f'title="{phase}: {ms:.1f}ms">'
                                    f'{phase[:4]}</span>'
                                )
                        bar_html = f'<div class="bbar">{"".join(bar_parts)}</div>'

                perf_row_cls = " class='wolfxl-row'" if lib == "wolfxl" else ""
                rows.append(
                    f"<tr{perf_row_cls}><td><b>{_esc(lib)}</b></td>"
                    f"<td data-v='{wall.get('p50', 9e9)}'>{_fmt_ms(wall.get('p50'))}</td>"
                    f"<td data-v='{wall.get('p95', 9e9)}'>{_fmt_ms(wall.get('p95'))}</td>"
                    f"<td data-v='{wall.get('min', 9e9)}'>{_fmt_ms(wall.get('min'))}</td>"
                    f"<td data-v='{cpu.get('p50', 9e9)}'>{_fmt_ms(cpu.get('p50'))}</td>"
                    f"<td data-v='{rss or 9e9}'>{_fmt_mb(rss)}</td>"
                    f"<td>{rate}</td>"
                    f"<td>{bar_html}</td></tr>"
                )

            rows.append("</tbody></table></div>")

        rows.append("</div></details>")

    rows.append("</section>")
    return "\n".join(rows)


def _section_memory(memory: list[dict[str, Any]] | None) -> str:
    if not memory:
        return ""

    # Classify adapters as Rust or Python
    rust_adapters = {"wolfxl", "rust_xlsxwriter", "python-calamine"}

    def _adapter_cls(name: str) -> str:
        return "rust" if name in rust_adapters else "python"

    # Group by (scale, op)
    groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for entry in memory:
        key = (entry["scale"], entry["op"])
        groups.setdefault(key, []).append(entry)

    # Sort scales in descending order (100k first)
    scale_order = {"100k": 0, "10k": 1, "1k": 2}

    rows: list[str] = []
    rows.append('<section id="memory" class="container">'
                '<h2>Memory Profiling</h2>')
    rows.append('<p style="font-size:.82rem;color:var(--text2);margin-bottom:.5rem">'
                'Each adapter measured in an isolated subprocess. RSS delta = memory '
                'allocated during the operation. Tracemalloc = Python heap only '
                '(Rust/C allocations are invisible to Python\'s allocator).</p>')

    # Check for key insight: Rust adapters with near-zero tracemalloc
    rust_write_100k = [e for e in memory
                       if e["scale"] == "100k" and e["op"] == "write"
                       and e["adapter"] in rust_adapters
                       and e.get("tracemalloc_peak_mb", 1) < 0.1]
    if rust_write_100k:
        rows.append(
            '<div class="mem-insight">'
            '<b>Key finding:</b> Rust-backed adapters allocate '
            '&lt;0.1 MB on the Python heap for 100K-cell writes &mdash; '
            'all memory lives in Rust\'s allocator with zero GC pressure.</div>'
        )

    rows.append('<div class="mem-legend">'
                '<span><span class="dot rust"></span> Rust / C-extension</span>'
                '<span><span class="dot python"></span> Pure Python</span></div>')

    # Render bar charts for each (scale, op) group
    for op_label, op_key in [("Bulk Read", "read"), ("Bulk Write", "write")]:
        op_groups = [(s, groups[(s, op_key)])
                     for s in ["100k", "10k", "1k"]
                     if (s, op_key) in groups]
        if not op_groups:
            continue

        rows.append(f'<div class="mem-group"><h3>{op_label} &mdash; RSS Delta (MB)</h3>')

        for scale, entries in op_groups:
            entries_sorted = sorted(entries, key=lambda e: e.get("rss_delta_mb", 0),
                                    reverse=True)
            max_rss = max((e.get("rss_delta_mb", 0) for e in entries_sorted), default=1)
            if max_rss <= 0:
                max_rss = 1

            cells = entries_sorted[0].get("cells", "?") if entries_sorted else "?"
            rows.append(f'<div class="mem-scale-label">{scale} ({cells:,} cells)</div>'
                        if isinstance(cells, int) else
                        f'<div class="mem-scale-label">{scale} ({cells} cells)</div>')

            for entry in entries_sorted:
                adapter = entry["adapter"]
                rss_delta = entry.get("rss_delta_mb", 0)
                tm_peak = entry.get("tracemalloc_peak_mb", 0)
                pct = max(rss_delta / max_rss * 100, 0.5) if rss_delta > 0 else 0.5
                cls = _adapter_cls(adapter)

                # Show tracemalloc as fraction of RSS
                tm_note = f"{tm_peak:.1f} MB" if tm_peak >= 0.1 else "&lt;0.1 MB"

                rows.append(
                    f'<div class="mem-bar-row">'
                    f'<div class="mem-bar-label" title="{_esc(adapter)}">'
                    f'{_esc(adapter)}</div>'
                    f'<div class="mem-bar-track">'
                    f'<div class="mem-bar-fill {cls}" style="width:{pct:.1f}%">'
                    f'{rss_delta:.1f}</div></div>'
                    f'<div class="mem-bar-val">{rss_delta:+.1f} MB</div>'
                    f'<div class="mem-bar-tm" title="Python heap (tracemalloc)">'
                    f'py: {tm_note}</div></div>'
                )

        rows.append("</div>")  # close mem-group

    # Summary table
    rows.append('<details class="expandable">'
                '<summary>Detailed Memory Table</summary>'
                '<div class="content"><div class="table-scroll">'
                '<table id="mem-table"><thead><tr>'
                '<th class="sort">Adapter</th>'
                '<th class="sort">Op</th>'
                '<th class="sort">Scale</th>'
                '<th class="sort" data-type="n">Cells</th>'
                '<th class="sort" data-type="n">RSS Delta (MB)</th>'
                '<th class="sort" data-type="n">Tracemalloc Peak (MB)</th>'
                '<th class="sort" data-type="n">Python Heap %</th>'
                '<th class="sort" data-type="n">RSS Total (MB)</th>'
                '</tr></thead><tbody>')

    for entry in sorted(memory,
                        key=lambda e: (scale_order.get(e["scale"], 9),
                                       e["op"], -e.get("rss_delta_mb", 0))):
        adapter = entry["adapter"]
        rss_delta = entry.get("rss_delta_mb", 0)
        tm_peak = entry.get("tracemalloc_peak_mb", 0)
        rss_total = entry.get("rss_after_mb", 0)
        cells = entry.get("cells", 0)
        py_pct = (tm_peak / rss_delta * 100) if rss_delta > 0 else 0

        rows.append(
            f'<tr><td><b>{_esc(adapter)}</b></td>'
            f'<td>{entry["op"]}</td>'
            f'<td>{entry["scale"]}</td>'
            f'<td data-v="{cells}">{cells:,}</td>'
            f'<td data-v="{rss_delta}">{rss_delta:+.2f}</td>'
            f'<td data-v="{tm_peak}">{tm_peak:.2f}</td>'
            f'<td data-v="{py_pct:.1f}">{py_pct:.0f}%</td>'
            f'<td data-v="{rss_total}">{rss_total:.1f}</td></tr>'
        )

    rows.append("</tbody></table></div></div></details>")
    rows.append("</section>")
    return "\n".join(rows)


def _section_diagnostics(fidelity: dict[str, Any]) -> str:
    results = fidelity.get("results", [])

    diags: list[dict[str, Any]] = []
    for entry in results:
        for tc in entry.get("test_cases", {}).values():
            if not isinstance(tc, dict):
                continue
            for op in ("read", "write"):
                if op in tc:
                    for d in tc[op].get("diagnostics", []):
                        diags.append(d)

    if not diags:
        return ('<section id="diag" class="container"><h2>Diagnostics</h2>'
                '<p style="color:var(--text2)">No diagnostics recorded.</p></section>')

    # Summary by category
    cat_counts: dict[str, int] = {}
    sev_counts: dict[str, int] = {}
    for d in diags:
        cat_counts[d.get("category", "?")] = cat_counts.get(d.get("category", "?"), 0) + 1
        sev_counts[d.get("severity", "?")] = sev_counts.get(d.get("severity", "?"), 0) + 1

    rows: list[str] = []
    rows.append('<section id="diag" class="container"><h2>Diagnostics</h2>')
    rows.append(f'<p style="color:var(--text2);margin-bottom:.75rem">'
                f'{len(diags)} total diagnostics</p>')

    # Summary cards
    rows.append('<div class="cards-grid">')
    for cat, cnt in sorted(cat_counts.items(), key=lambda x: -x[1]):
        rows.append(f'<div class="stat-card"><div class="val">{cnt}</div>'
                    f'<div class="lbl">{_esc(cat)}</div></div>')
    rows.append('</div>')

    # Severity breakdown
    rows.append('<h3>By Severity</h3><div class="cards-grid">')
    for sev, cnt in sorted(sev_counts.items()):
        rows.append(f'<div class="stat-card"><div class="val">{cnt}</div>'
                    f'<div class="lbl">{_esc(sev)}</div></div>')
    rows.append('</div>')

    # Detail table (behind details)
    rows.append(
        '<details class="expandable"><summary>All Diagnostics</summary>'
        '<div class="content"><div class="table-scroll"><table>'
        '<thead><tr><th class="sort">Category</th><th>Severity</th>'
        '<th>Feature</th><th>Op</th><th>Adapter Message</th>'
        '<th>Probable Cause</th></tr></thead><tbody>'
    )

    dash = "\u2014"
    for d in diags[:500]:  # cap at 500 for page size
        loc = d.get("location", {})
        cause = _esc(d.get('probable_cause') or dash)
        rows.append(
            f"<tr><td>{_esc(d.get('category'))}</td>"
            f"<td>{_esc(d.get('severity'))}</td>"
            f"<td>{_esc(loc.get('feature'))}</td>"
            f"<td>{_esc(loc.get('operation'))}</td>"
            f"<td style='font-size:.75rem'>{_esc(d.get('adapter_message'))}</td>"
            f"<td style='font-size:.75rem'>{cause}</td></tr>"
        )

    rows.append("</tbody></table></div></div></details>")
    rows.append("</section>")
    return "\n".join(rows)
