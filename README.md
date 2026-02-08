# ExcelBench

**Objective, reproducible fidelity scores for Python Excel libraries.**

Most Excel library comparisons focus on speed. ExcelBench answers the question developers actually have: **"Can this library handle my complex spreadsheet?"**

We test 17 features — from cell values to conditional formatting to images — across 5 mainstream Python libraries, scoring each on a 0–3 fidelity scale against real Excel-generated reference files.

## Results at a Glance

> Last run: 2026-02-06 &bull; Excel 16.105.3 &bull; macOS (Apple Silicon)

### XLSX Profile

| Feature | openpyxl | | xlsxwriter | python-calamine | pylightxl | |
|:--------|:--------:|:--------:|:----------:|:---------------:|:---------:|:---------:|
| | **Read** | **Write** | **Write** | **Read** | **Read** | **Write** |
| **Tier 1 — Essential** | | | | | | |
| Cell values | 🟢 3 | 🟢 3 | 🟢 3 | 🟠 1 | 🟢 3 | 🟠 1 |
| Formulas | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🟢 3 |
| Text formatting | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Background colors | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Number formats | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Alignment | 🟢 3 | 🟢 3 | 🟢 3 | 🟠 1 | 🔴 0 | 🟠 1 |
| Borders | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Dimensions | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Multiple sheets | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 | 🟢 3 |
| **Tier 2 — Standard** | | | | | | |
| Merged cells | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Conditional formatting | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Data validation | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Hyperlinks | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Images | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Comments | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Freeze panes | 🟢 3 | 🟢 3 | 🟢 3 | 🔴 0 | 🔴 0 | 🔴 0 |
| Pivot tables | ➖ | ➖ | ➖ | ➖ | ➖ | ➖ |

**xlrd** is omitted from the XLSX table — it only supports the legacy `.xls` format (see [XLS results](#xls-profile) below).

### XLS Profile

| Feature | xlrd (Read) | python-calamine (Read) |
|:--------|:-----------:|:----------------------:|
| Cell values | 🟢 3 | 🟢 3 |
| Alignment | 🟢 3 | 🟠 1 |
| Dimensions | 🟢 3 | 🔴 0 |
| Multiple sheets | 🟢 3 | 🟢 3 |

### Score Legend

| Score | Meaning |
|:------|:--------|
| 🟢 3 | **Complete** — full fidelity, indistinguishable from Excel |
| 🟡 2 | **Functional** — works for common cases, some edge-case failures |
| 🟠 1 | **Minimal** — basic recognition but significant limitations |
| 🔴 0 | **Unsupported** — errors, corruption, or complete data loss |
| ➖ | Not applicable (library doesn't support this format/operation) |

## Libraries Tested

| Library | Version | Lang | Capabilities | Notes |
|:--------|:--------|:-----|:-------------|:------|
| [openpyxl](https://openpyxl.readthedocs.io/) | 3.1.5 | Python | Read + Write | Full-featured, pure Python |
| [XlsxWriter](https://xlsxwriter.readthedocs.io/) | 3.2.9 | Python | Write only | Write-optimized, excellent formatting |
| [python-calamine](https://github.com/dimastbk/python-calamine) | 0.6.1 | Rust | Read only | Fast reads via Rust `calamine` crate |
| [pylightxl](https://github.com/PydPiper/pylightxl) | 1.61 | Python | Read + Write | Zero-dependency, lightweight |
| [xlrd](https://github.com/python-excel/xlrd) | 2.0.2 | Python | Read only | Legacy `.xls` format only |

## How It Works

1. **Generate reference files** — [xlwings](https://www.xlwings.org/) drives real Excel to produce canonical `.xlsx`/`.xls` test files with known features.
2. **Read tests** — each library reads the Excel-generated file; extracted values are compared to the expected manifest.
3. **Write tests** — each library writes a new file from the same spec; the output is verified by re-reading with a trusted oracle (Excel via xlwings, or openpyxl in CI).
4. **Score** — pass rates map to the 0–3 fidelity scale per feature.

## Quick Start

```bash
# Install
uv sync

# Run the benchmark against pre-built fixtures (no Excel required)
uv run excelbench benchmark --tests fixtures/excel --output results

# View results
cat results/xlsx/README.md
```

To regenerate canonical fixtures from scratch (requires Excel installed):

```bash
uv run excelbench generate --output fixtures/excel
```

## Optional: Rust Backends (PyO3)

ExcelBench can optionally load additional adapters backed by Rust libraries via a
local PyO3 extension module (`excelbench_rust`). This is intentionally kept as a
separate crate so the main `excelbench` package remains pure-Python.

Prereqs:
- Rust toolchain (`rustup`, `cargo`)

Build + install the extension into the active venv:

```bash
# Install maturin + other optional deps
uv sync --extra rust

# Build/editable-install the PyO3 module
uv run maturin develop --manifest-path rust/excelbench_rust/Cargo.toml \
  --features calamine,rust_xlsxwriter,umya

# Sanity check
uv run python -c "import excelbench_rust; print(excelbench_rust.build_info())"
```

Notes:
- `uv sync` may uninstall the locally-built extension module; rerun `maturin develop` if Rust adapters disappear.
- You can build subsets (faster iteration):
  - `--features calamine`
  - `--features rust_xlsxwriter`
  - `--features umya`

Once installed, additional adapters may appear in `get_all_adapters()`:
- `calamine` (Rust, read-only)
- `rust_xlsxwriter` (Rust, write-only)
- `umya-spreadsheet` (Rust, read+write)

## Detailed Results

- **[XLSX detailed results](results/xlsx/README.md)** — per-library, per-test-case breakdowns
- **[XLS detailed results](results/xls/README.md)** — legacy format results
- **[CSV export](results/matrix.csv)** — machine-readable flat file
- **[Run history](results/history.jsonl)** — append-only log of scores across runs

## Methodology

- **Real Excel as source of truth** — test fixtures are generated by Excel itself via xlwings, not hand-crafted XML
- **Independent Read/Write scores** — because library capabilities often differ
- **Detailed scoring rubrics** — objective 0–3 criteria for each feature ([rubrics](rubrics/fidelity-rubrics.md))
- **Reproducible** — canonical fixtures are tracked in git; CI runs the full benchmark on every push

Full methodology: [METHODOLOGY.md](METHODOLOGY.md)

## Feature Coverage

### Implemented (Tier 1 + 2)

| Tier | Features |
|:-----|:---------|
| **Tier 1** — Essential | Cell values, formulas, text formatting, background colors, number formats, alignment, borders, dimensions, multiple sheets |
| **Tier 2** — Standard | Merged cells, conditional formatting, data validation, hyperlinks, images, pivot tables*, comments, freeze panes |

\* Pivot tables require a Windows-generated fixture; macOS support is limited.

### Planned (Tier 3)

Charts, named ranges, complex conditional formatting, tables (structured references), print settings, protection.

## Project Status

**v0.1.0** — Tier 1 + Tier 2 complete for 5 Python libraries. CI green. Actively maintained.

Roadmap:
- Rust library integration (rust_xlsxwriter, umya-spreadsheet) via PyO3
- Tier 3 features (charts, named ranges, protection)
- Interactive web viewer for results

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for setup instructions, how to add features, and how to add library adapters.

## License

MIT
