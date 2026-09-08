# ExcelBench

**Objective, reproducible spreadsheet-library benchmarks with a Python-first decision surface.**

Most Excel library comparisons focus on speed. ExcelBench focuses on the question developers actually have:

> Can this library handle my real spreadsheet without breaking the parts I care about?

ExcelBench models 22 XLSX features across 14 Python adapters plus a
cross-language context lane (Apache POI, Excelize, zavora-xlsx). Tier 4
(sheet protection, page setup, chart anchoring) was added in the
2026-09-08 competitor snapshot.

## Results at a Glance

> Competitor snapshot: 2026-09-08 | WolfXL 2.1.0 wheel | aspose-cells-foss 26.7 | [Fidelity](results-2026-09-08/xlsx/README.md) | [Mutation](results-2026-09-08/mutation/README.md) | [Calc](results-2026-09-08/calc/README.md) | [Zavora cross-language](results-2026-09-08/cross-language/README.md)
>
> In that snapshot: openpyxl 22/22, WolfXL 19/22, aspose-cells-foss 11/22
> green features (WolfXL 2.1.0 gaps: conditional-formatting read,
> named-range and print-title writes).
>
> Mutation suite (two-cell edit on a corporate template): WolfXL patches in
> 1.5s with 100% preservation; openpyxl and aspose rewrite at 60%;
> zavora-xlsx 0.1.2 corrupts package relationships (0%).
>
> Calc tier (133-formula financial DAG, cache-free fixture, LibreOffice
> oracle): LibreOffice 133/133, WolfXL 55/133, aspose-cells-foss 25/133,
> zavora 0/133.
>
> Python release snapshot: 2026-04-29 UTC | wheel-backed WolfXL 2.0 rerun | [Fidelity](results-release-2026-04-28/README.md) | [Perf](results-release-2026-04-28/perf/README.md) | [Dashboard](results-release-2026-04-28/DASHBOARD.md)
>
> In that snapshot, WolfXL reaches `18/18` green features with `100%` pass rate.
>
> Cross-language context snapshot: [Apache POI `18/18`](results-cross-language/README.md) | [Excelize `18/18`](results-cross-language/README.md)
>
> Pivot capability lane: [separate artifact](results-cross-language-pivots/README.md) because the shipped macOS pivot fixture is not scoreable, while `excelize` can still emit pivot-bearing workbooks.
>
> Historical public baseline: 2026-02-17 | Excel 16.105.3 | macOS (Apple Silicon) | [Full results](results/xlsx/README.md)
>
> Newer performance snapshot: 2026-04-20 | [Perf results](results/perf/README.md)
>
> Read [Public Reporting Status](docs/public-reporting.md) before quoting numbers across snapshots.

![ExcelBench Heatmap](results/xlsx/heatmap.png)

**The current story:** ExcelBench now has three useful lanes. The Python release lane answers the migration question. The cross-language lane answers the ecosystem-positioning question. The pivot capability lane captures pivot evidence separately when the scored fixture is not valid on this platform. Keep every claim tied to the exact dated artifact you are citing.

## Comparison Scope

ExcelBench intentionally keeps the main public comparison **Python-first**. That is the decision surface most users care about: `openpyxl`, `xlsxwriter`, `python-calamine`, `pandas`, and adjacent Python options.

Cross-language libraries matter too, but for a different reason: they show how strong WolfXL looks next to mature spreadsheet tooling outside Python. The current checked-in cross-language lane includes:

- `Apache POI`
- `Excelize`
- `zavora-xlsx` (Rust writer, added 2026-09-08)

Pivot tables sit in a separate capability lane. On macOS, the shipped pivot fixture does not currently contain scoreable pivot OOXML, so the pivot story is tracked as a dedicated artifact instead of being mixed into the scored lane.

See [cross-language comparison strategy](docs/trackers/cross-language-comparison-strategy.md).

## Lanes

1. **Python replacement lane**
Use this when the question is: what should a Python team use instead of `openpyxl`?

2. **Cross-language context lane**
Use this when the question is: how does WolfXL compare to serious spreadsheet tooling in Java, Go, and Rust?

3. **Pivot capability lane**
Use this when the question is: can the cross-language helpers detect or emit pivot-bearing workbooks even when the main scored fixture is not valid on macOS?

4. **Template-mutation lane**
Use this when the question is: which engine can surgically edit a corporate template fastest while preserving the package?

5. **Formula-recalculation lane**
Use this when the question is: which engine actually computes a 133-formula financial DAG from scratch (cache-free fixture, LibreOffice oracle)?

### Library Comparison

| Library | Caps | Fidelity | Read Speed | Write Speed | Modify |
|---------|:----:|:--------:|:----------:|:-----------:|:------:|
| **wolfxl** | R+W | 19/22 in 2026-09-08 competitor snapshot (18/18 in 2026-04-29 snapshot) | workload-specific | workload-specific | Patch (100% preservation in mutation suite) |
| openpyxl | R+W | 22/22 in 2026-09-08 competitor snapshot | 1x (baseline per workload) | 1x (baseline per workload) | Rewrite (60% preservation in mutation suite) |
| aspose-cells-foss | R+W | 11/22 in 2026-09-08 competitor snapshot | workload-specific | workload-specific | Rewrite (60% preservation) |
| xlsxwriter | W | 15/18 in 2026-04-29 release snapshot | -- | ~1x | No |
| xlsxwriter-constmem | W | 12/18 in 2026-04-29 release snapshot | -- | ~2x | No |
| python-calamine | R | 1/18 in 2026-04-29 release snapshot | ~1.3x | -- | No |
| pandas | R+W | 3/18 in 2026-04-29 release snapshot | <1x | <1x | Rebuild |
| polars | R | 0/18 in 2026-04-29 release snapshot | ~1x | -- | No |

> Speed numbers are snapshot-specific. Always cite the artifact date, workload, and profile. See [performance results](results/perf/README.md), [METHODOLOGY.md](METHODOLOGY.md), and [Public Reporting Status](docs/public-reporting.md).

### Key Findings
- **High-fidelity libraries are rare**: in the 2026-04-29 release snapshot, only openpyxl and WolfXL reached 18/18 green features; in the 2026-09-08 competitor snapshot (22 features), openpyxl holds 22/22 while WolfXL 2.1.0 drops to 19/22
- **WolfXL 2.1.0 regression signal**: conditional-formatting read scores 0 and named-range / print-title writes score 2 in the 2026-09-08 snapshot; tracked for the WolfXL repo
- **Patch modify is structurally different**: WolfXL's `load_workbook(path, modify=True)` uses surgical ZIP patching; it is the only engine reaching 100% package preservation in the mutation suite
- **The abstraction tax is real**: pandas wraps openpyxl but drops from 16 to 3 green features due to DataFrame coercion (errors become NaN)
- **Speed vs fidelity tradeoff is measurable**: use the perf snapshot together with the fidelity matrix rather than quoting one without the other
- **Optimization modes have clear costs**: openpyxl-readonly loses 13 green features for streaming speed
- **Cross-language context is now strong too**: `Apache POI` and `Excelize` land at `18/18` in the scored write lane; `zavora-xlsx` 0.1.2 writes fast but corrupts hyperlink relationships on mutate and cannot recalculate
- **Calculation is a differentiator**: only LibreOffice computes the full 133-formula DAG; WolfXL 2.1.0 covers 55/133 (power-operator family returns None), aspose-cells-foss 25/133


See the [release snapshot dashboard](results-release-2026-04-28/DASHBOARD.md) for the fresh wheel-backed combined view, or the [historical dashboard](results/DASHBOARD.md) for the older public baseline.

### Score Legend

| Score | Meaning |
|:------|:--------|
| 🟢 3 | **Complete** -- full fidelity, indistinguishable from Excel |
| 🟡 2 | **Functional** -- works for common cases, some edge-case failures |
| 🟠 1 | **Minimal** -- basic recognition but significant limitations |
| 🔴 0 | **Unsupported** -- errors, corruption, or complete data loss |

## Libraries Tested

### XLSX Profile (14 adapters)

| Library | Version | Lang | Caps | Green Features |
|:--------|:--------|:-----|:-----|:--------------:|
| [WolfXL](https://github.com/SynthGL/wolfxl) | 2.1.0 | Python (Rust core) | R+W | 19/22 (2026-09-08 snapshot) |
| [openpyxl](https://openpyxl.readthedocs.io/) | 3.1.5 | Python | R+W | 22/22 (2026-09-08 snapshot) |
| [aspose-cells-foss](https://pypi.org/project/aspose-cells-foss/) | 26.7 | Python (JVM-core FOSS) | R+W | 11/22 (2026-09-08 snapshot) |
| [XlsxWriter](https://xlsxwriter.readthedocs.io/) | 3.2.9 | Python | W | 15/18 |
| [xlsxwriter-constmem](https://xlsxwriter.readthedocs.io/) | 3.2.9 | Python | W | 12/18 |
| [openpyxl-readonly](https://openpyxl.readthedocs.io/) | 3.1.5 | Python | R | 3/18 |
| [pandas](https://pandas.pydata.org/) | 3.0.0 | Python | R+W | 3/18 |
| [pyexcel](https://github.com/pyexcel/pyexcel) | 0.7.4 | Python | R+W | 3/18 |
| [tablib](https://tablib.readthedocs.io/) | 3.9.0 | Python | R+W | 3/18 |
| [pylightxl](https://github.com/PydPiper/pylightxl) | 1.61 | Python | R+W | 2/18 |
| [python-calamine](https://github.com/dimastbk/python-calamine) | 0.6.1 | Rust | R | 1/18 |
| [polars](https://pola.rs/) | 1.38.1 | Rust | R | 0/18 |
| [xlwt](https://github.com/python-excel/xlwt) | 1.3.0 | Python | W | 4/18 |
| [xlrd](https://github.com/python-excel/xlrd) | 2.0.2 | Python | R | .xls only |

> Green-feature counts are per dated snapshot; `x/18` numbers come from the
> 2026-04-29 release snapshot and `x/22` from the 2026-09-08 competitor
> snapshot (Tier 4 added). Never mix counts across snapshots.

### Cross-Language Context (oracle helpers)

| Library | Version | Lang | Caps | Notes |
|:--------|:--------|:-----|:-----|:------|
| Apache POI | 5.x | Java | W | 18/18 scored write lane (2026-04-29) |
| Excelize | 2.x | Go | W | 18/18 scored write lane (2026-04-29) |
| zavora-xlsx | 0.1.2 | Rust | W | Fast writer; 0.1.2 corrupts hyperlink relationships on mutate; no recalc |

### XLS Profile (2 adapters)

| Library | Green Features | Notes |
|:--------|:--------------:|:------|
| xlrd | 4/4 | Full .xls read fidelity |
| python-calamine | 2/4 | Cross-format reader |

### Optional: Rust Backends (PyO3)

Five additional adapters via Rust/PyO3 extension modules:

| Library | Caps | Source | Notes |
|:--------|:-----|:-------|:------|
| [WolfXL](https://github.com/SynthGL/wolfxl) (calamine-styled) | R | PyPI | Full-fidelity Rust reader with style extraction |
| [WolfXL](https://github.com/SynthGL/wolfxl) (rust_xlsxwriter) | W | PyPI | Full-fidelity Rust writer |
| calamine (basic) | R | Local | Direct calamine bindings (data only, no styles) |
| rust_xlsxwriter (direct) | W | Local | Direct rust_xlsxwriter bindings |
| umya-spreadsheet | R+W | Local | Rust read + write |

```bash
# WolfXL adapters (from PyPI — no Rust toolchain needed)
uv sync --extra rust

# Local-only adapters (requires Rust toolchain + maturin)
uv run maturin develop --manifest-path rust/excelbench_rust/Cargo.toml \
  --features calamine,rust_xlsxwriter,umya
```

> `uv sync` may uninstall locally-built extensions; rerun `maturin develop` after.

### Cross-Language Context

ExcelBench now ships a separate cross-language context snapshot for mature non-Python spreadsheet libraries:

- `Apache POI` (Java)
- `Excelize` (Go)
- `zavora-xlsx` (Rust writer, added 2026-09-08; helper in `tools/external-oracles/zavora`)

These are not framed as Python drop-in replacements. They answer a different question: how strong is WolfXL relative to serious spreadsheet tooling outside Python?

Current checked-in cross-language snapshot:

- [`results-cross-language/README.md`](results-cross-language/README.md)
- [`results-cross-language/CONTEXT.md`](results-cross-language/CONTEXT.md)
- [`results-cross-language-pivots/README.md`](results-cross-language-pivots/README.md)
- [`docs/cross-language-context.md`](docs/cross-language-context.md)

Current takeaways from that snapshot:

- `apache-poi`: `18/18` green features in the scored write surfaces of this lane
- `excelize`: `18/18` green features in the scored write surfaces of this lane
- `pivot_tables`: tracked in a separate capability artifact because the shipped macOS fixture does not contain scoreable pivot OOXML, while `excelize` can still emit pivot-bearing workbooks

The concrete rollout plan for the first two candidates is here:

- [Apache POI + Excelize rollout plan](docs/trackers/apache-poi-excelize-rollout-plan.md)
- [Apache POI adapter design](docs/trackers/apache-poi-adapter-design.md)
- [Excelize adapter design](docs/trackers/excelize-adapter-design.md)

Run the dedicated cross-language context snapshot with:

```bash
uv run excelbench cross-language-context --tests fixtures/excel --output results-cross-language
```

Run the dedicated pivot capability artifact with:

```bash
uv run excelbench cross-language-pivot-context --fixture fixtures/excel/tier2/15_pivot_tables.xlsx --output results-cross-language-pivots
```

## How It Works

1. **Generate reference files** -- [xlwings](https://www.xlwings.org/) drives real Excel to produce canonical `.xlsx`/`.xls` test files with known features.
2. **Read tests** -- each library reads the Excel-generated file; extracted values are compared to the expected manifest.
3. **Write tests** -- each library writes a new file from the same spec; the output is verified by a trusted oracle (Excel via xlwings, or openpyxl in CI).
4. **Score** -- pass rates map to the 0-3 fidelity scale per feature.

Full methodology: [METHODOLOGY.md](METHODOLOGY.md)

## Public Reporting Rules

- Treat each `results/` directory as a dated snapshot.
- Do not merge February fidelity claims and April perf claims into one undated headline.
- Cite the artifact date and workload whenever quoting a speedup number.
- Keep WolfXL-specific release claims aligned with the WolfXL repo's evidence page.

See [docs/public-reporting.md](docs/public-reporting.md).

## WolfXL Docs

WolfXL documentation lives in the [wolfxl repository](https://github.com/SynthGL/wolfxl/tree/main/docs).

- [Quickstart](https://github.com/SynthGL/wolfxl/blob/main/docs/getting-started/quickstart.md)
- [Openpyxl migration guide](https://github.com/SynthGL/wolfxl/blob/main/docs/migration/openpyxl-migration.md)
- [Compatibility matrix](https://github.com/SynthGL/wolfxl/blob/main/docs/migration/compatibility-matrix.md)
- [Benchmark methodology](https://github.com/SynthGL/wolfxl/blob/main/docs/performance/methodology.md)
- [Known limitations](https://github.com/SynthGL/wolfxl/blob/main/docs/trust/limitations.md)

## Quick Start

```bash
# Install
uv sync

# Run the benchmark against pre-built fixtures (no Excel required)
uv run excelbench benchmark --tests fixtures/excel --output results

# Template-mutation suite (wall time, RSS, preservation)
uv run excelbench mutation

# Formula-recalculation tier (cache-free fixture, LibreOffice oracle)
uv run excelbench calc

# Generate the heatmap
uv run excelbench heatmap

# Generate the combined fidelity + performance dashboard
uv run excelbench dashboard

# View results
open results/xlsx/README.md  # macOS; use xdg-open on Linux
```

To regenerate canonical fixtures from scratch (requires Excel installed):

```bash
uv run excelbench generate --output fixtures/excel
```

## Feature Coverage

### Tested (22 features)

| Tier | Features | Count |
|:-----|:---------|:-----:|
| **Tier 0** -- Core | Cell values, formulas, multiple sheets | 3 |
| **Tier 1** -- Formatting | Text formatting, background colors, number formats, alignment, borders, dimensions | 6 |
| **Tier 2** -- Advanced | Merged cells, conditional formatting, data validation, hyperlinks, images, comments, freeze panes, pivot tables | 8 |
| **Tier 3** -- Workbook metadata | Named ranges, tables | 2 |
| **Tier 4** -- Production surfaces (2026-09-08) | Sheet protection, page setup, chart anchoring | 3 |

> Pivot tables are tested but score N/A across all adapters in the current macOS run.
> Green-feature denominators: /18 in the 2026-04-29 release snapshot, /22 in the 2026-09-08 competitor snapshot.

### Planned

None currently queued. Tier 4 (charts anchoring, print settings, protection) shipped 2026-09-08.

## Detailed Results

- **[Competitor snapshot fidelity](results-2026-09-08/xlsx/README.md)** -- 22 features x 14 Python adapters, WolfXL 2.1.0 + aspose-cells-foss
- **[Competitor snapshot heatmap](results-2026-09-08/xlsx/heatmap.png)** ([SVG](results-2026-09-08/xlsx/heatmap.svg)) -- 22x14 visual score matrix
- **[Competitor snapshot mutation](results-2026-09-08/mutation/README.md)** -- template mutation: wall time, RSS, preservation
- **[Competitor snapshot calc](results-2026-09-08/calc/README.md)** -- 133-formula financial DAG vs LibreOffice oracle
- **[Competitor snapshot zavora cross-language](results-2026-09-08/cross-language/README.md)** -- zavora-xlsx 0.1.2 Rust writer context
- **[XLSX results](results/xlsx/README.md)** -- per-library, per-test-case breakdowns with tier list
- **[Release snapshot results](results-release-2026-04-28/README.md)** -- fresh wheel-backed WolfXL 2.0 rerun
- **[XLS results](results/xls/README.md)** -- legacy format results
- **[Performance results](results/perf/README.md)** -- throughput benchmarks (cells/s)
- **[Release snapshot perf](results-release-2026-04-28/perf/README.md)** -- matching wheel-backed perf snapshot
- **[Dashboard](results/DASHBOARD.md)** -- combined fidelity + performance comparison
- **[Release snapshot dashboard](results-release-2026-04-28/DASHBOARD.md)** -- combined view for the fresh rerun
- **[Heatmap (PNG)](results/xlsx/heatmap.png)** | **[SVG](results/xlsx/heatmap.svg)** -- visual score matrix

## Project Status

**v0.1.0** -- actively maintained benchmarking harness with dated fidelity and performance snapshots, reproducible methodology, and multi-adapter coverage across Python and Rust-backed spreadsheet libraries.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for setup instructions, how to add features, and how to add library adapters.

## License

MIT
