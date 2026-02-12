# ExcelBench

**Objective, reproducible fidelity scores for Python Excel libraries.**

Most Excel library comparisons focus on speed. ExcelBench answers the question developers actually have: **"Can this library handle my complex spreadsheet?"**

We test 16 features across 12 Python adapters, scoring each on a 0-3 fidelity scale against real Excel-generated reference files.

## Results at a Glance

> Last run: 2026-02-08 | Excel 16.105.3 | macOS (Apple Silicon) | [Full results](results/xlsx/README.md)

![ExcelBench Heatmap](results/xlsx/heatmap.png)

**The story:** openpyxl and xlsxwriter achieve full fidelity across all 16 features. Once you move past basic cell values, nearly every other library drops to zero. Formatting, comments, hyperlinks, images, merged cells, conditional formatting -- all red.

### Library Tiers

| Tier | Library | Green Features | Summary |
|:----:|---------|:--------------:|---------|
| **S** | openpyxl | 16/16 | Reference adapter -- full read + write fidelity |
| **S** | xlsxwriter | 16/16 | Best write-only option -- full formatting support |
| **A** | xlsxwriter-constmem | 13/16 | Memory-optimized write -- loses images, comments, row height |
| **B** | xlwt | 4/16 | Legacy .xls writer -- basic formatting subset |
| **C** | openpyxl-readonly, pandas, pyexcel, tablib | 2-3/16 | Values + basic formatting only |
| **D** | polars | 0/16 | Columnar type coercion drops all fidelity |

### Key Findings

- **The abstraction tax is real**: pandas wraps openpyxl but drops from 16 to 3 green features due to DataFrame coercion (errors become NaN)
- **Speed vs fidelity tradeoff**: xlsxwriter-constmem writes at 4.7M cells/s but loses 3 features; polars reads at 1.3M cells/s but scores 0/16 green
- **Optimization modes have clear costs**: openpyxl-readonly loses 13 green features for streaming speed

See the [full dashboard](results/DASHBOARD.md) for the combined fidelity + performance comparison.

### Score Legend

| Score | Meaning |
|:------|:--------|
| 🟢 3 | **Complete** -- full fidelity, indistinguishable from Excel |
| 🟡 2 | **Functional** -- works for common cases, some edge-case failures |
| 🟠 1 | **Minimal** -- basic recognition but significant limitations |
| 🔴 0 | **Unsupported** -- errors, corruption, or complete data loss |

## Libraries Tested

### XLSX Profile (12 adapters)

| Library | Version | Lang | Caps | Green Features |
|:--------|:--------|:-----|:-----|:--------------:|
| [openpyxl](https://openpyxl.readthedocs.io/) | 3.1.5 | Python | R+W | 16/16 |
| [XlsxWriter](https://xlsxwriter.readthedocs.io/) | 3.2.9 | Python | W | 16/16 |
| [xlsxwriter-constmem](https://xlsxwriter.readthedocs.io/) | 3.2.9 | Python | W | 13/16 |
| [openpyxl-readonly](https://openpyxl.readthedocs.io/) | 3.1.5 | Python | R | 3/16 |
| [pandas](https://pandas.pydata.org/) | 3.0.0 | Python | R+W | 3/16 |
| [pyexcel](https://github.com/pyexcel/pyexcel) | 0.7.4 | Python | R+W | 3/16 |
| [tablib](https://tablib.readthedocs.io/) | 3.9.0 | Python | R+W | 3/16 |
| [pylightxl](https://github.com/PydPiper/pylightxl) | 1.61 | Python | R+W | 2/16 |
| [python-calamine](https://github.com/dimastbk/python-calamine) | 0.6.1 | Rust | R | 1/16 |
| [polars](https://pola.rs/) | 1.38.1 | Rust | R | 0/16 |
| [xlwt](https://github.com/python-excel/xlwt) | 1.3.0 | Python | W | 4/16 |
| [xlrd](https://github.com/python-excel/xlrd) | 2.0.2 | Python | R | .xls only |

### XLS Profile (2 adapters)

| Library | Green Features | Notes |
|:--------|:--------------:|:------|
| xlrd | 4/4 | Full .xls read fidelity |
| python-calamine | 2/4 | Cross-format reader |

### Optional: Rust Backends (PyO3)

Three additional adapters via a local PyO3 extension module:

| Library | Caps | Notes |
|:--------|:-----|:------|
| calamine (Rust) | R | Direct Rust calamine bindings |
| rust_xlsxwriter | W | Rust write bindings |
| umya-spreadsheet | R+W | Rust read + write |

```bash
uv sync --extra rust
uv run maturin develop --manifest-path rust/excelbench_rust/Cargo.toml \
  --features calamine,rust_xlsxwriter,umya
```

> `uv sync` may uninstall the locally-built extension; rerun `maturin develop` after.

## How It Works

1. **Generate reference files** -- [xlwings](https://www.xlwings.org/) drives real Excel to produce canonical `.xlsx`/`.xls` test files with known features.
2. **Read tests** -- each library reads the Excel-generated file; extracted values are compared to the expected manifest.
3. **Write tests** -- each library writes a new file from the same spec; the output is verified by a trusted oracle (Excel via xlwings, or openpyxl in CI).
4. **Score** -- pass rates map to the 0-3 fidelity scale per feature.

Full methodology: [METHODOLOGY.md](METHODOLOGY.md)

## Quick Start

```bash
# Install
uv sync

# Run the benchmark against pre-built fixtures (no Excel required)
uv run excelbench benchmark --tests fixtures/excel --output results

# Generate the heatmap
uv run excelbench heatmap

# Generate the combined fidelity + performance dashboard
uv run excelbench dashboard

# View results
open results/xlsx/README.md
```

To regenerate canonical fixtures from scratch (requires Excel installed):

```bash
uv run excelbench generate --output fixtures/excel
```

## Feature Coverage

### Scored (16 features)

| Tier | Features |
|:-----|:---------|
| **Tier 0** -- Basic Values | Cell values, formulas, multiple sheets |
| **Tier 1** -- Formatting | Text formatting, background colors, number formats, alignment, borders, dimensions |
| **Tier 2** -- Advanced | Merged cells, conditional formatting, data validation, hyperlinks, images, comments, freeze panes |

### In Progress (Tier 3)

Named ranges and tables have generators and tests but are not yet in the official scored results.

### Planned

Charts, print settings, protection.

> Pivot tables have a generator but require a Windows-generated fixture (macOS Excel limitation).

## Detailed Results

- **[XLSX results](results/xlsx/README.md)** -- per-library, per-test-case breakdowns with tier list
- **[XLS results](results/xls/README.md)** -- legacy format results
- **[Performance results](results/perf/README.md)** -- throughput benchmarks (cells/s)
- **[Dashboard](results/DASHBOARD.md)** -- combined fidelity + performance comparison
- **[Heatmap (PNG)](results/xlsx/heatmap.png)** | **[SVG](results/xlsx/heatmap.svg)** -- visual score matrix

## Project Status

**v0.1.0** -- 16 scored features, 12 Python xlsx adapters + 2 xls + 3 Rust/PyO3.

1084 tests passing. Actively maintained.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for setup instructions, how to add features, and how to add library adapters.

## License

MIT
