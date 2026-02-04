# ExcelBench: Comprehensive Excel Library Benchmark Suite

> Design Document - February 4, 2026

## Executive Summary

**ExcelBench** is a comprehensive benchmark suite comparing Excel libraries' **feature parity** across Python and Rust ecosystems. Unlike typical benchmarks focused on speed, ExcelBench measures **fidelity** - how accurately each library can read and write Excel features compared to native Excel.

**Goals:**
- Serve as a **definitive developer reference** for choosing Excel libraries
- Maintain **research-grade rigor** with reproducible methodology
- Provide **objective, nuanced scoring** (not just "supported/unsupported")

**Target Libraries:**

| Python | Rust |
|--------|------|
| openpyxl | calamine |
| xlsxwriter | rust_xlsxwriter |
| xlrd | umya-spreadsheet |
| pylightxl | |
| pyexcel | |

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                        ExcelBench                                │
├─────────────────────────────────────────────────────────────────┤
│  1. TEST FILE GENERATOR (xlwings → Excel)                       │
│     - Creates labeled .xlsx files exercising each feature       │
│     - Self-documenting: description column + test column        │
│     - Source of truth: real Excel writes the files              │
│                                                                 │
│  2. TEST HARNESS                                                │
│     - Loads each library via unified adapter interface          │
│     - Attempts to read/write each feature                       │
│     - Scores fidelity (0-3 scale) using defined rubrics         │
│     - Rust libraries integrated via PyO3 bindings               │
│                                                                 │
│  3. RESULTS ENGINE                                              │
│     - Raw output: results.json (source of truth)                │
│     - Generated views: markdown tables, CSV                     │
│     - Future: interactive website                               │
└─────────────────────────────────────────────────────────────────┘
```

---

## Key Design Decisions

### 1. Test File Generation: xlwings + Real Excel

**Decision:** Use xlwings to drive the actual Excel application on Mac.

**Rationale:**
- Guaranteed correct output - Excel itself writes the test files
- Avoids chicken-and-egg problem of using a library to test libraries
- Python-based, fits our stack

**Tradeoff:** Requires Excel installed on the machine running the generator.

### 2. Fidelity Scoring (0-3 Scale)

**Decision:** Score features on a 4-point fidelity scale with separate Read and Write scores.

| Score | Meaning |
|-------|---------|
| 0 | Unsupported - errors, corruption, or complete data loss |
| 1 | Minimal - basic recognition but significant limitations |
| 2 | Functional - works for common cases, some edge case failures |
| 3 | Complete - full fidelity, indistinguishable from Excel |

**Rationale:**
- Binary (yes/no) loses important nuance
- Separate read/write scores because capabilities often differ
- Detailed rubrics (see `rubrics/fidelity-rubrics.md`) ensure reproducibility

### 3. Tiered Feature Coverage

**Decision:** Organize features into tiers, prioritized by pain-point research.

**Tier 1 - Essential (9 features):**
1. Cell Values
2. Formulas
3. Basic Text Formatting
4. Cell Background Color
5. Number Formats
6. Cell Alignment
7. Borders
8. Column Widths / Row Heights
9. Multiple Sheets

**Tier 2 - Standard (8 features):**
10. Merged Cells
11. Conditional Formatting
12. Data Validation
13. Hyperlinks
14. Images/Embedded Objects
15. Pivot Tables
16. Comments and Notes
17. Freeze Panes / Split Views

**Tier 3 - Advanced (6 features):**
18. Charts
19. Named Ranges
20. Complex Conditional Formatting
21. Tables (Structured References)
22. Print Settings
23. Protection

### 4. Rust Integration via PyO3

**Decision:** Create Python bindings for Rust libraries rather than maintaining two harnesses.

**Rationale:**
- Single test runner (Python) simplifies maintenance
- Side-by-side comparison of Python and Rust libraries
- Avoid duplicating test logic

### 5. Results as JSON Source of Truth

**Decision:** Store raw results in JSON, generate all other formats from it.

**Rationale:**
- Maximum flexibility for future output formats
- Researchers can consume raw data directly
- Version-controlled and diffable

---

## Repository Structure

```
ExcelBench/
├── README.md                    # Project overview, quick results link
├── METHODOLOGY.md               # Research methodology, scoring rubrics
├── pyproject.toml               # Python deps (uv)
├── Cargo.toml                   # Rust workspace for PyO3 bindings
│
├── generator/                   # xlwings-based test file generator
│   ├── __init__.py
│   ├── generate.py              # Main entry point
│   ├── features/                # One module per feature
│   │   ├── borders.py
│   │   ├── formulas.py
│   │   └── ...
│   └── manifest.py              # Generates manifest.json
│
├── harness/                     # Test execution engine
│   ├── __init__.py
│   ├── runner.py                # Orchestrates test runs
│   ├── scoring.py               # Applies fidelity rubrics
│   ├── adapters/                # Library adapters
│   │   ├── base.py              # Protocol definition
│   │   ├── openpyxl_adapter.py
│   │   ├── xlsxwriter_adapter.py
│   │   └── ...
│   └── rust_bindings/           # PyO3 wrappers
│       ├── Cargo.toml
│       └── src/
│           ├── calamine_bind.rs
│           ├── umya_bind.rs
│           └── rust_xlsxwriter_bind.rs
│
├── rubrics/                     # Fidelity scoring definitions
│   └── fidelity-rubrics.md      # GPT 5.2 Pro generated rubrics
│
├── test_files/                  # Generated Excel files (gitignored)
│   ├── tier1/
│   ├── tier2/
│   ├── tier3/
│   └── manifest.json
│
├── results/                     # Benchmark outputs
│   ├── results.json             # Raw data
│   ├── README.md                # Generated summary
│   ├── matrix.csv               # Flat export
│   └── features/                # Per-feature breakdowns
│
├── scripts/
│   ├── generate_test_files.py   # Runs generator
│   ├── run_benchmark.py         # Runs full benchmark
│   └── render_results.py        # JSON → markdown/csv
│
├── prompts/                     # LLM handoff prompts
│   └── gpt-fidelity-rubrics-prompt.md
│
└── docs/
    └── plans/
        └── 2026-02-04-excelbench-design.md  # This document
```

---

## Test File Structure

Each generated Excel file is **self-documenting** with a consistent 3-column pattern:

| Column A (Label) | Column B (Test Cell) | Column C (Expected) |
|------------------|----------------------|---------------------|
| "Bold text" | **Hello** | `bold=True` |
| "14pt Arial" | Hello | `font_size=14, font_name=Arial` |
| "Red background" | [red cell] | `bg_color=#FF0000` |

**Benefits:**
- Human-verifiable: open in Excel to visually confirm
- Machine-parseable: Column C contains structured expected values
- Isolated features: one file per category prevents interaction effects

**manifest.json** describes all test files:
```json
{
  "generated": "2026-02-04T10:30:00Z",
  "excel_version": "16.83",
  "files": [
    {
      "path": "tier1/07_borders.xlsx",
      "feature": "borders",
      "tier": 1,
      "test_cases": ["thin_border", "thick_border", "diagonal_up", "..."]
    }
  ]
}
```

---

## Results Schema

```json
{
  "metadata": {
    "benchmark_version": "1.0.0",
    "run_date": "2026-02-04T10:30:00Z",
    "excel_version": "16.83",
    "platform": "darwin-arm64"
  },
  "libraries": {
    "openpyxl": {
      "version": "3.1.2",
      "language": "python",
      "capabilities": ["read", "write"]
    }
  },
  "results": [
    {
      "feature": "borders",
      "tier": 1,
      "library": "openpyxl",
      "scores": {
        "read": 3,
        "write": 2
      },
      "test_cases": {
        "thin_border": {"read": "pass", "write": "pass"},
        "diagonal_up": {"read": "pass", "write": "fail"},
        "theme_color_border": {"read": "partial", "write": "fail"}
      },
      "notes": "Diagonal borders not supported in write mode"
    }
  ]
}
```

**Generated Outputs:**
1. `results/README.md` - Feature matrix with color-coded scores
2. `results/features/*.md` - Per-feature detailed breakdowns
3. `results/matrix.csv` - Flat export for analysis

---

## Implementation Phases

### Phase 1 - Foundation (MVP)

**Scope:**
- Set up repo structure, pyproject.toml, basic CI
- Implement generator for 3 Tier 1 features: cell values, basic formatting, borders
- Build harness with 2 adapters: openpyxl (read/write), xlsxwriter (write)
- Basic results.json output + simple markdown table

**Deliverable:** Can run benchmark on subset, see comparison table

**Key Tasks:**
- [ ] Initialize Python project with uv
- [ ] Set up xlwings generator framework
- [ ] Implement cell values generator + tests
- [ ] Implement basic formatting generator + tests
- [ ] Implement borders generator + tests
- [ ] Create adapter protocol/interface
- [ ] Implement openpyxl adapter
- [ ] Implement xlsxwriter adapter
- [ ] Build test runner
- [ ] Output results.json
- [ ] Generate markdown summary

---

### Phase 2 - Tier 1 Complete

**Scope:**
- Remaining Tier 1 features: formulas, number formats, alignment, dimensions, multiple sheets
- Add adapters: pylightxl, xlrd, pyexcel
- Integrate fidelity rubrics into scoring logic
- Improve results rendering (per-feature pages)

**Deliverable:** Complete Tier 1 benchmark across all Python libraries

**Key Tasks:**
- [ ] Implement formulas generator (formula text + cached values)
- [ ] Implement number formats generator
- [ ] Implement alignment generator
- [ ] Implement dimensions generator
- [ ] Implement multiple sheets generator
- [ ] Create pylightxl adapter
- [ ] Create xlrd adapter
- [ ] Create pyexcel adapter
- [ ] Parse rubrics into scoring logic
- [ ] Generate per-feature markdown pages

---

### Phase 3 - Rust Integration

**Scope:**
- Set up PyO3 bindings workspace
- Build adapters: calamine, rust_xlsxwriter, umya-spreadsheet
- Run Tier 1 benchmark including Rust libraries

**Deliverable:** Python vs Rust comparison for Tier 1

**Key Tasks:**
- [ ] Set up Cargo workspace with maturin/PyO3
- [ ] Create calamine Python bindings
- [ ] Create rust_xlsxwriter Python bindings
- [ ] Create umya-spreadsheet Python bindings
- [ ] Implement Rust library adapters
- [ ] Run comparative benchmark
- [ ] Update results with Rust libraries

---

### Phase 4 - Tier 2 Expansion

**Scope:**
- Add Tier 2 features: merged cells, conditional formatting, pivot tables, etc.
- Pain-point research: scan GitHub issues to prioritize edge cases
- Expand test cases based on real-world complaints

**Deliverable:** Comprehensive Tier 1+2 benchmark

**Key Tasks:**
- [ ] Research GitHub issues for each library (pain points)
- [ ] Prioritize edge cases within Tier 2 features
- [ ] Implement Tier 2 generators (8 features)
- [ ] Extend adapters for Tier 2 features
- [ ] Update scoring with Tier 2 rubrics
- [ ] Publish updated results

---

### Phase 5 - Polish & Publish

**Scope:**
- Tier 3 features (as time permits)
- Website/interactive viewer (optional)
- Performance benchmarking track (separate from fidelity)
- Write-up: methodology doc suitable for citation

**Deliverable:** Public release, shareable reference

**Key Tasks:**
- [ ] Implement priority Tier 3 features
- [ ] Write METHODOLOGY.md with academic rigor
- [ ] Create project website (optional)
- [ ] Add performance benchmark track
- [ ] Publish to GitHub with proper documentation
- [ ] Announce/share with community

---

## Deferred Items (Future Work)

### Performance Benchmarking Track
- Read speed: time to load N rows × M columns
- Write speed: time to generate files of various sizes
- Memory usage profiling
- Streaming/chunked read capabilities

*Different methodology, different test files. Can be added without changing fidelity architecture.*

### Round-Trip Integrity Testing
- Create → read → write → compare workflow
- Detects "silent corruption" where features degrade through cycles

*Valuable but adds complexity; fidelity scoring covers most cases.*

### Multi-Dimensional Scoring Expansion
- Separate scores for: Read, Write, Round-trip, Edge cases
- Aggregated "overall" score with weighting

*Current approach already separates read/write; expand dimensions later.*

### Interactive Website
- Filterable comparison matrix
- "Find libraries that support X" search
- Per-library detail pages

*JSON output enables this; build after initial results validated.*

### CI Integration for Library Maintainers
- GitHub Actions workflow libraries can adopt
- Automated re-scoring on library releases

*Nice-to-have after benchmark is stable.*

---

## Technical Notes

### xlwings on Mac
- Requires Microsoft Excel installed
- Uses AppleScript under the hood to drive Excel
- Install: `uv add xlwings`
- Ensure Excel has automation permissions in System Preferences

### PyO3 Bindings Strategy
- Use maturin for building Python wheels from Rust
- Each Rust library gets minimal bindings exposing read/write primitives
- Bindings live in `harness/rust_bindings/`

### Scoring Consistency
- Rubrics are defined in `rubrics/fidelity-rubrics.md`
- Two independent scorers should arrive at the same score
- When in doubt, score conservatively (lower)
- Document edge cases in test results notes

---

## Open Questions

1. **xlrd scope**: xlrd has limited xlsx support (primarily xls). Include xlsx tests anyway, or focus on xls format for xlrd specifically?

2. **pyexcel abstraction**: pyexcel wraps other libraries. Test it as its own entity, or note which backend it's using?

3. **Version pinning**: How to handle library version updates? Re-run entire benchmark or incremental updates?

4. **Community contributions**: Accept PRs for additional libraries? What's the bar for inclusion?

---

## References

- [Fidelity Scoring Rubrics](../../rubrics/fidelity-rubrics.md) - Detailed 0-3 scoring definitions per feature
- [GPT Handoff Prompt](../../prompts/gpt-fidelity-rubrics-prompt.md) - Prompt used to generate rubrics

---

*Document authored: February 4, 2026*
*Last updated: February 4, 2026*
