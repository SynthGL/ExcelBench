# CLAUDE.md — ExcelBench

## What This Project Is

ExcelBench is an objective, reproducible benchmark suite that scores Python Excel libraries on **feature fidelity** (correctness) and optionally **performance** (speed/memory). It tests 16 scored features across 12+ adapters using real Excel-generated fixtures as ground truth.

## Quick Reference

```bash
# Install deps
uv sync

# Run tests
uv run pytest

# Run fidelity benchmark (requires pre-built fixtures)
uv run excelbench benchmark --tests fixtures/excel --output results

# Run both xlsx + xls profiles
uv run excelbench benchmark-profiles --output results

# Run performance benchmark
uv run excelbench perf --tests fixtures/excel --output results

# Regenerate reports from existing results.json (no re-run needed)
uv run excelbench report --input results/xlsx/results.json --output results/xlsx

# Generate heatmap PNG + SVG
uv run excelbench heatmap

# Generate combined fidelity + performance dashboard
uv run excelbench dashboard

# Generate single-file interactive HTML dashboard
uv run excelbench html

# Generate fidelity-vs-throughput scatter plots
uv run excelbench scatter

# Lint + typecheck
uv run ruff check
uv run mypy
```

### Rust Adapters (optional)

```bash
uv sync --extra rust
uv run maturin develop --manifest-path rust/excelbench_rust/Cargo.toml \
  --features calamine,rust_xlsxwriter,umya
uv run python -c "import excelbench_rust; print(excelbench_rust.build_info())"
```

**Gotcha**: `uv sync` may uninstall the locally-built extension. Rerun `maturin develop` after syncing.

### Fixture Generation (requires Excel installed)

```bash
uv run excelbench generate --output fixtures/excel      # .xlsx (xlwings + Excel)
uv run excelbench generate-xls --output fixtures/excel_xls  # .xls
```

## Architecture

```
src/excelbench/
  cli.py                    # Typer CLI: generate, benchmark, perf, report
  models.py                 # Core dataclasses: CellValue, CellFormat, BorderInfo, etc.
  generator/
    base.py                 # FeatureGenerator protocol
    generate.py             # xlsx generation orchestrator (xlwings)
    generate_xls.py         # xls generation (xlwt)
    features/               # 17 feature generators (one file per feature)
  harness/
    runner.py               # Benchmark orchestrator: read tests + write verification
    adapters/
      base.py               # ExcelAdapter ABC + ReadOnlyAdapter/WriteOnlyAdapter
      __init__.py            # Registry: get_all_adapters() with optional-import guards
      openpyxl_adapter.py   # Reference adapter (R+W, full fidelity)
      xlsxwriter_adapter.py # Write-only, full fidelity
      calamine_adapter.py   # python-calamine (read-only)
      ...                   # 12 more adapters (see registry)
      rust_*.py / umya_*.py # Rust/PyO3 adapters (require excelbench_rust)
  perf/
    runner.py               # Performance measurement (wall/cpu/rss, phase breakdown)
    renderer.py             # Perf results to markdown/CSV/JSON
  results/
    renderer.py             # Fidelity results to markdown/CSV/JSON
    html_dashboard.py       # Single-file interactive HTML dashboard generator
    scatter.py              # Fidelity-vs-throughput scatter plots (PNG/SVG)

rust/excelbench_rust/       # Separate PyO3 crate (not part of hatchling build)
  src/lib.rs                # Module entry + build_info()
  src/calamine_backend.rs   # calamine read bindings
  src/rust_xlsxwriter_backend.rs  # rust_xlsxwriter write bindings
  src/umya_backend.rs       # umya-spreadsheet R+W bindings

fixtures/
  excel/                    # Canonical .xlsx fixtures (git-tracked, Excel-generated)
  excel_xls/                # Canonical .xls fixtures
  throughput_xlsx/           # Scale fixtures for perf benchmarks

tests/                      # pytest + pytest-cov
```

## Key Patterns

### Adding a New Adapter

1. Create `src/excelbench/harness/adapters/<name>_adapter.py`
2. Subclass `ExcelAdapter` (or `ReadOnlyAdapter`/`WriteOnlyAdapter`)
3. Implement required abstract methods (see `base.py`)
4. Add optional-import guard + registration in `__init__.py` → `get_all_adapters()`
5. Add dependency to `pyproject.toml` if needed

### Adding a New Feature (Tier 3)

1. Create generator in `src/excelbench/generator/features/<name>.py` (subclass `FeatureGenerator`)
2. Register in `features/__init__.py` and `generate.py`
3. Add read/write handling in `harness/runner.py`
4. Add adapter methods to `base.py` protocol + implement in each adapter
5. Regenerate fixtures: `uv run excelbench generate --output fixtures/excel`

### Scoring Model

- 3 = all tests pass (complete fidelity)
- 2 = >=80% pass
- 1 = >=50% pass
- 0 = <50% pass
- Per-test importance weights: `basic` (must-pass) vs `edge` (bonus)

### Oracle Strategy

- **Primary**: Excel via xlwings (write verification)
- **Fallback**: openpyxl (CI/headless)
- Performance mode skips oracle entirely

## Conventions

- **Python 3.11+**, type hints required on function signatures
- **Package manager**: `uv` (lockfile: `uv.lock`)
- **Linter**: `ruff` (line-length 100, rules: E/F/I/N/W/UP)
- **Type checker**: `mypy --strict` (subset of files; expand incrementally)
- **Test runner**: `pytest -v --cov=excelbench`
- **Build backend**: hatchling (pure Python); Rust via maturin (separate crate)
- **Commit style**: conventional commits (`feat`/`fix`/`refactor`/`test`/`docs`/`chore`)

## Current State (as of 2026-02-12)

- **Fidelity**: Tier 0 (3) + Tier 1 (6) + Tier 2 (7 scored + pivot_tables pending) = 16 scored features
- **Adapters**: 12 Python xlsx + 2 xls + 3 Rust/PyO3 = 17 total
- **Performance**: Runner + renderer + throughput dashboard operational
- **Visualizations**: Heatmap (PNG/SVG), combined fidelity+perf dashboard, tier list
- **Rust adapters**: Built locally, not in CI benchmark; cell values only (no formatting)
- **Tier 3 features**: named_ranges + tables have generators/tests, not yet in official scored results
- **CLI commands**: generate, generate-xls, benchmark, benchmark-profiles, perf, report, heatmap, dashboard, html, scatter
- **CI/CD**: GitHub Actions CI (lint/test/benchmark on all pushes) + deploy-dashboard (auto-deploys HTML to Vercel on results/generator changes)

## Trackers & Plans

- `docs/plans/2026-02-04-excelbench-design.md` — Original design doc
- `docs/plans/phase1-implementation.md` — Phase 1 plan (done)
- `docs/plans/phase3-rust-integration.md` — Rust integration plan (mostly done)
- `docs/plans/2026-02-08-performance-benchmarks.md` — Perf track design
- `docs/trackers/library-expansion-tracker.md` — Adapter inventory + scores
- `docs/trackers/performance-benchmarks.md` — Perf implementation tracker
- `docs/trackers/performance-benchmark-runs.md` — Run log

## Deployment

- **Dashboard URL**: https://excelbench.vercel.app
- **Auto-deploy workflow**: `.github/workflows/deploy-dashboard.yml`
  - Triggers on pushes to `master` that change `results/xlsx/results.json`, `results/perf/results.json`, `results/xlsx/*.svg`, or `src/excelbench/results/html_dashboard.py`
  - Also supports manual trigger via `workflow_dispatch`
  - Generates HTML dashboard with `uv run excelbench html --output deploy/index.html`, then deploys to Vercel
  - Requires 3 GitHub secrets: `VERCEL_TOKEN`, `VERCEL_ORG_ID`, `VERCEL_PROJECT_ID`

## Gotchas

- `uv sync` clobbers locally-installed Rust extension — always rerun `maturin develop` after
- xlwings requires Excel installed + macOS accessibility permissions for fixture generation
- Pivot tables need a Windows-generated fixture (macOS Excel limitation)
- `test_files/` is gitignored scratch space; `fixtures/excel/` is the canonical tracked set
- `results_dev_*` directories are ephemeral benchmark outputs (gitignored)
