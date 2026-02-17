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
from pathlib import Path
from typing import Any

# ── Feature ordering / tier map (shared with other renderers) ──────

_FEATURE_ORDER: list[str] = [
    "cell_values", "formulas", "multiple_sheets",
    "alignment", "background_colors", "borders",
    "dimensions", "number_formats", "text_formatting",
    "comments", "conditional_formatting", "data_validation",
    "freeze_panes", "hyperlinks", "images", "merged_cells",
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
    perf = None
    if perf_json and perf_json.exists():
        perf = json.loads(perf_json.read_text())

    memory: list[dict[str, Any]] | None = None
    if memory_json and memory_json.exists():
        memory = json.loads(memory_json.read_text())

    svgs: dict[str, str] = {}
    if scatter_dir:
        for name in ("scatter_tiers", "scatter_features", "heatmap"):
            svg_path = scatter_dir / f"{name}.svg"
            if svg_path.exists():
                svgs[name] = svg_path.read_text()

    body_parts = [
        _section_nav(has_memory=memory is not None),
        _section_overview(fidelity, perf),
        _section_matrix(fidelity),
        _section_scatter(svgs),
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
:root{
  --bg:#f8fafc;--card:#fff;--border:#e2e8f0;
  --text:#1e293b;--text2:#64748b;--accent:#2563eb;
  --g3:#dcfce7;--g3t:#166534;--g2:#fef9c3;--g2t:#854d0e;
  --g1:#ffedd5;--g1t:#9a3412;--g0:#fee2e2;--g0t:#991b1b;
  --na:#f1f5f9;--nat:#94a3b8;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
  background:var(--bg);color:var(--text);font-size:14px;line-height:1.5}
a{color:var(--accent);text-decoration:none}
a:hover{text-decoration:underline}
.container{max-width:1440px;margin:0 auto;padding:1.5rem 1.5rem}

/* ── Nav ── */
nav{position:sticky;top:0;z-index:100;background:linear-gradient(135deg,#0f172a,#1e293b);
  padding:.6rem 1.5rem;display:flex;align-items:center;gap:1.5rem;
  box-shadow:0 2px 8px rgba(0,0,0,.15)}
nav .brand{font-weight:700;font-size:1.1rem;color:#e2e8f0;letter-spacing:-.02em}
nav .links{display:flex;gap:1rem;flex-wrap:wrap}
nav .links a{color:#94a3b8;font-size:.82rem;font-weight:500;transition:color .15s}
nav .links a:hover{color:#fff;text-decoration:none}

/* ── Cards ── */
.card{background:var(--card);border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.06);
  padding:1.5rem;margin-bottom:1.5rem;border:1px solid var(--border)}
.cards-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));
  gap:1rem;margin:1rem 0}
.stat-card{background:var(--card);border-radius:10px;padding:1.2rem;text-align:center;
  border:1px solid var(--border)}
.stat-card .val{font-size:2rem;font-weight:700;color:var(--accent)}
.stat-card .lbl{font-size:.78rem;color:var(--text2);margin-top:.2rem}

/* ── Section headings ── */
section{margin-bottom:2rem}
h1{font-size:1.6rem;font-weight:700;margin-bottom:.5rem}
h2{font-size:1.3rem;font-weight:700;margin-bottom:.75rem;
  padding-bottom:.4rem;border-bottom:2px solid var(--border)}
h3{font-size:1.05rem;font-weight:600;margin:1rem 0 .5rem}
.meta-bar{font-size:.8rem;color:var(--text2);margin-bottom:1rem}

/* ── Tables ── */
table{border-collapse:collapse;width:100%;font-size:.82rem}
th,td{padding:.45rem .6rem;border:1px solid var(--border);text-align:left;vertical-align:top}
th{background:#f1f5f9;font-weight:600;white-space:nowrap;position:sticky;top:0;z-index:1}
th.sort{cursor:pointer;user-select:none}
th.sort:hover{background:#e2e8f0}
th.sort::after{content:'';margin-left:4px;opacity:.3}
th.asc::after{content:' \\25B2'}
th.desc::after{content:' \\25BC'}
tbody tr:hover{background:#f8fafc}
.table-scroll{overflow-x:auto;margin-bottom:1rem;border-radius:8px;border:1px solid var(--border)}
.table-scroll table{border:none}
.table-scroll td,.table-scroll th{border-left:none;border-right:none}

/* ── Score colours ── */
.s3{background:var(--g3);color:var(--g3t);font-weight:700;text-align:center}
.s2{background:var(--g2);color:var(--g2t);font-weight:700;text-align:center}
.s1{background:var(--g1);color:var(--g1t);font-weight:700;text-align:center}
.s0{background:var(--g0);color:var(--g0t);font-weight:700;text-align:center}
.sna{background:var(--na);color:var(--nat);text-align:center}

/* ── Matrix ── */
.matrix th,.matrix td{padding:.35rem .5rem;text-align:center;min-width:52px;font-size:.78rem}
.matrix .feat{text-align:left;font-weight:500;min-width:130px;white-space:nowrap}
.tier-row td{background:#0f172a !important;color:#e2e8f0;font-weight:600;
  font-size:.78rem;padding:.3rem .6rem;letter-spacing:.02em}

/* ── Details ── */
details{border:1px solid var(--border);border-radius:8px;margin-bottom:.6rem;background:var(--card)}
details summary{padding:.7rem 1rem;cursor:pointer;font-weight:500;font-size:.88rem;
  list-style:none;display:flex;align-items:center;gap:.5rem}
details summary::before{content:'\\25B6';font-size:.65rem;
  color:var(--text2);transition:transform .2s}
details[open] summary::before{transform:rotate(90deg)}
details[open] summary{border-bottom:1px solid var(--border)}
details .content{padding:.75rem 1rem}
.badge{display:inline-block;padding:.1rem .5rem;border-radius:4px;font-size:.7rem;font-weight:600}
.badge-t0{background:#dbeafe;color:#1e40af}
.badge-t1{background:#fef3c7;color:#92400e}
.badge-t2{background:#ede9fe;color:#5b21b6}
.badge-t3{background:#d1fae5;color:#065f46}

/* ── Test case table ── */
.tc-table td,.tc-table th{font-size:.75rem;padding:.3rem .5rem}
.tc-table .pass{color:#166534}
.tc-table .fail{color:#991b1b;font-weight:600}
code.val{font-family:'JetBrains Mono','Fira Code',monospace;font-size:.72rem;
  background:#f1f5f9;padding:.1rem .3rem;border-radius:3px;word-break:break-all}

/* ── Perf breakdown bar ── */
.bbar{display:flex;height:18px;border-radius:4px;overflow:hidden;width:100%;min-width:120px}
.bbar span{display:flex;align-items:center;justify-content:center;font-size:8px;
  color:#fff;overflow:hidden;white-space:nowrap;padding:0 3px}
.bbar span:nth-child(1){background:#3b82f6}
.bbar span:nth-child(2){background:#8b5cf6}
.bbar span:nth-child(3){background:#06b6d4}
.bbar span:nth-child(4){background:#10b981}

/* ── SVG container ── */
.svg-wrap{overflow-x:auto;margin:1rem 0}
.svg-wrap svg{max-width:100%;height:auto}

/* ── Filter ── */
.filter-box{margin-bottom:.75rem}
.filter-box input{padding:.4rem .75rem;border:1px solid var(--border);border-radius:6px;
  font-size:.85rem;width:280px;outline:none}
.filter-box input:focus{border-color:var(--accent);box-shadow:0 0 0 2px rgba(37,99,235,.15)}

/* ── Memory bars ── */
.mem-group{margin-bottom:1.5rem}
.mem-group h3{font-size:.95rem;margin-bottom:.6rem}
.mem-scale-label{font-size:.82rem;font-weight:600;color:var(--text2);
  margin:.8rem 0 .4rem;text-transform:uppercase;letter-spacing:.04em}
.mem-bar-row{display:flex;align-items:center;gap:.5rem;margin-bottom:.35rem;font-size:.78rem}
.mem-bar-label{width:130px;text-align:right;font-weight:500;flex-shrink:0;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mem-bar-track{flex:1;height:22px;background:#f1f5f9;border-radius:4px;position:relative;
  overflow:hidden;min-width:100px}
.mem-bar-fill{height:100%;border-radius:4px;display:flex;align-items:center;
  padding-left:6px;font-size:.7rem;font-weight:600;color:#fff;white-space:nowrap;
  transition:width .3s ease}
.mem-bar-fill.rust{background:linear-gradient(90deg,#f97316,#fb923c)}
.mem-bar-fill.python{background:linear-gradient(90deg,#3b82f6,#60a5fa)}
.mem-bar-val{width:80px;text-align:right;font-family:'JetBrains Mono','Fira Code',monospace;
  font-size:.75rem;flex-shrink:0}
.mem-bar-tm{width:80px;text-align:right;font-family:'JetBrains Mono','Fira Code',monospace;
  font-size:.72rem;color:var(--text2);flex-shrink:0}
.mem-legend{display:flex;gap:1.2rem;margin:.75rem 0 .5rem;font-size:.78rem}
.mem-legend span{display:flex;align-items:center;gap:.3rem}
.mem-legend .dot{width:10px;height:10px;border-radius:2px;display:inline-block}
.mem-legend .dot.rust{background:#f97316}
.mem-legend .dot.python{background:#3b82f6}
.mem-insight{background:#fffbeb;border:1px solid #fef3c7;border-radius:8px;padding:.8rem 1rem;
  font-size:.82rem;color:#92400e;margin-bottom:1rem}

/* ── Misc ── */
.btn{display:inline-block;padding:.3rem .8rem;border:1px solid var(--border);border-radius:6px;
  font-size:.78rem;cursor:pointer;background:var(--card);color:var(--text2)}
.btn:hover{background:#f1f5f9}
footer{text-align:center;padding:2rem;color:var(--text2);font-size:.78rem}
.flex-bar{display:flex;align-items:center;gap:.75rem;flex-wrap:wrap;margin-bottom:.75rem}
@media(max-width:768px){
  .cards-grid{grid-template-columns:1fr 1fr}
  nav .links{gap:.5rem}
  .container{padding:1rem}
}
"""

# ====================================================================
#  JavaScript
# ====================================================================

_JS = r"""
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
/* Smooth scroll */
document.querySelectorAll('nav a[href^="#"]').forEach(a=>{
  a.addEventListener('click',e=>{
    e.preventDefault();
    const el=document.querySelector(a.getAttribute('href'));
    if(el) el.scrollIntoView({behavior:'smooth',block:'start'});
  });
});
"""

# ====================================================================
#  Section renderers
# ====================================================================


def _section_nav(*, has_memory: bool = False) -> str:
    links = [
        ("#overview", "Overview"),
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
        f'<nav><div class="brand">ExcelBench</div>'
        f'<div class="links">{link_html}</div></nav>'
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

    return (
        f'<section id="overview" class="container">'
        f'<h1>ExcelBench Dashboard</h1>'
        f'<div class="meta-bar">'
        f'Profile: <b>{_esc(meta.get("profile", "xlsx"))}</b> &middot; '
        f'Platform: {_esc(meta.get("platform", "?"))} &middot; '
        f'Excel: {_esc(meta.get("excel_version", "?"))} &middot; '
        f'Date: {_esc(meta.get("run_date", "?")[:10])}'
        f'</div>'
        f'<div class="cards-grid">{cards_html}</div>'
        f'</section>'
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
        rows.append(f"<th><div>{_esc(lib)}</div><div style='font-size:.65rem;color:var(--text2)'>"
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
            tip = f"Read: {_score_label(rs)} / Write: {_score_label(ws)}"
            rows.append(f'<td class="{cls}" title="{tip}">{_score_label(best)}</td>')
        rows.append("</tr>")

    rows.append("</tbody></table></div></section>")
    return "\n".join(rows)


def _section_scatter(svgs: dict[str, str]) -> str:
    if not svgs:
        return ""
    parts = ['<section id="scatter" class="container"><h2>Fidelity vs. Throughput</h2>']
    for key, label in [
        ("scatter_tiers", "By Feature Group"),
        ("scatter_features", "Per Feature"),
        ("heatmap", "Heatmap"),
    ]:
        if key in svgs:
            parts.append(f'<h3>{label}</h3><div class="svg-wrap">{svgs[key]}</div>')
    parts.append("</section>")
    return "\n".join(parts)


def _section_comparison(fidelity: dict[str, Any], perf: dict[str, Any] | None) -> str:
    results = fidelity.get("results", [])
    libs_info = fidelity.get("libraries", {})

    # Per-lib stats
    lib_stats: dict[str, dict[str, Any]] = {}
    for lib, info in libs_info.items():
        cap = _cap_label(info.get("capabilities", []))
        lib_stats[lib] = {"cap": cap, "version": info.get("version", "?"),
                          "green": 0, "scored": 0, "passed": 0, "total": 0}

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
                '<input type="text" class="filter-input" data-target="cmp-table" '
                'placeholder="Filter libraries\u2026"></div></div>')
    rows.append('<div class="table-scroll"><table id="cmp-table">')
    rows.append("<thead><tr>"
                "<th class='sort'>Library</th>"
                "<th>Caps</th>"
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
        rows.append(f"<tr><td><b>{_esc(lib)}</b></td>"
                    f"<td>{st['cap']}</td>"
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
            rows.append(
                f"<tr><td><b>{_esc(lib)}</b></td>"
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

                rows.append(
                    f"<tr><td><b>{_esc(lib)}</b></td>"
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
    rust_adapters = {"wolfxl", "calamine-styled", "rust_xlsxwriter", "python-calamine"}

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
