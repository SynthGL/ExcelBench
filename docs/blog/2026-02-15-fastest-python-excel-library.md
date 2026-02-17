# The Fastest Full-Featured Python Excel Library You've Never Heard Of

Most Python Excel tutorials start with `pip install openpyxl`. It's the default, the safe choice, the one every Stack Overflow answer recommends. But what if there's a library that's **3–12x faster** and supports the same 17 out of 18 Excel features?

That library is **WolfXL**.

## The Problem: Speed vs. Features

If you process Excel files in Python, you've hit this wall. The landscape splits into two camps:

**Full-featured but slow.** openpyxl reads and writes fonts, borders, formulas, merged cells, conditional formatting, data validations, hyperlinks, comments, freeze panes, named ranges, and tables. It does everything. But at 100K cells, you're waiting 400ms+ per read — and that's before you touch formatting.

**Fast but stripped-down.** python-calamine and polars rip through cell values at native speed. fastexcel wraps calamine with a nice API. But try reading a font color or writing a border — you can't. These libraries explicitly skip formatting, formulas, and merged cells.

This tradeoff has been the status quo for years. Until now.

## Measuring It: ExcelBench

Before making claims, we built the measurement tool. [ExcelBench](https://excelbench.vercel.app) is an open-source benchmark suite that scores Python Excel libraries on two axes:

- **Feature fidelity**: 18 features tested per library using real Excel-generated fixtures as ground truth. Each feature is scored 0–3 based on pass rate across basic and edge-case tests.
- **Performance**: Wall-clock time (p50) for read and write operations at various scales, from 200 to 100K cells.

The benchmark tests 19 libraries across 12 Python adapters and 5 Rust/PyO3 adapters. Results are verified against Excel's own output (via xlwings) or openpyxl as an oracle.

## The Results

Here are the headline numbers from our benchmark runs on an M3 MacBook Pro:

| Operation | WolfXL | openpyxl | xlsxwriter | Speedup |
|---|---|---|---|---|
| Per-cell read (10K) | 10.0ms | 35.2ms | — | **3.5x** |
| Per-cell write (10K) | 11.9ms | 36.8ms | 28.8ms | **3.1x** |
| Bulk read (100K) | 102.6ms | 362.2ms | — | **3.5x** |
| Bulk write (100K) | 57.8ms | 287.8ms | 172.3ms | **5.0x** |
| Styled read (bg colors 1K) | 1.6ms | 7.6ms | — | **4.8x** |
| Borders write (200 cells) | 0.9ms | 10.7ms | 4.3ms | **12.5x** |
| Formulas read (10K) | 13.5ms | 43.8ms | — | **3.2x** |
| Formulas write (10K) | 8.0ms | 37.1ms | 77.7ms | **4.6x** |

The speedups range from 3x on plain cell I/O to **12.5x on border writes**. At the 100K scale that matters for real workloads, WolfXL's bulk write pushes **1.73 million cells/sec** versus openpyxl's 347K/sec.

And these aren't just values. These are fully-formatted cells with the correct fonts, colors, number formats, and borders preserved.

## Where It Stands Among Rust-backed Libraries

WolfXL isn't the first library to wrap Rust Excel crates in Python. Here's how it compares:

| Library | Read | Write | Styles | Formulas | Merged Cells |
|---|---|---|---|---|---|
| **WolfXL** | calamine | rust_xlsxwriter | Full | Full | Full |
| python-calamine | calamine | — | None | None | None |
| fastexcel | calamine | — | None | None | None |
| fastxlsx | calamine | rust_xlsxwriter | None | None | None |

The key differentiator: **fastexcel** and **python-calamine** are read-only and skip all formatting. **fastxlsx** wraps the same Rust crates as WolfXL for read+write, but explicitly does not support styles, formulas, or merged cells — it's designed for raw data throughput.

WolfXL is the only Rust-backed Python library with full style fidelity.

## Fidelity: 17 out of 18

On ExcelBench's 18-feature fidelity matrix, WolfXL scores green (3/3) on 17 features for both read and write:

- **Cell values** — strings, numbers, booleans, dates, datetimes, blanks, errors
- **Formulas** — preserved on read, written correctly on write (custom XML parser eliminated 86% of FFI overhead)
- **Text formatting** — bold, italic, underline, strikethrough, font name/size/color
- **Background colors** — solid fills with full hex color fidelity
- **Number formats** — currency, percentages, dates, custom patterns
- **Alignment** — horizontal, vertical, wrap text, rotation, indent
- **Borders** — all 14 border styles, per-edge colors
- **Dimensions** — row heights and column widths
- **Multiple sheets** — read all sheets, create new sheets on write
- **Merged cells** — range detection and preservation
- **Conditional formatting** — rules, operators, formulas, formats
- **Data validations** — list, whole number, decimal, date, text length
- **Hyperlinks** — external URLs, internal references, display text, tooltips
- **Comments** — text, author, per-cell notes
- **Freeze panes** — row/column splits
- **Named ranges** and **Tables** — full read and write

The one gap: **images** (feature 14/18). Neither calamine nor rust_xlsxwriter has image support yet.

## How It Works

WolfXL is a hybrid: it uses **calamine** (via calamine-styled, a fork with OOXML style parsing) for reading and **rust_xlsxwriter** for writing. Both are compiled to native code via PyO3/maturin.

The architecture hits a sweet spot:

1. **calamine** is the fastest Excel reader in any language. The styled fork adds full OOXML parsing for format records, border definitions, and shared strings — without sacrificing calamine's read speed.
2. **rust_xlsxwriter** is a pure-Rust implementation of the xlsx writer that produces files Excel opens without complaint. It handles the full OOXML spec for styles, formulas, and structural features.
3. **A custom XML formula parser** replaces the naive per-cell FFI calls with a single-pass extraction from the sheet XML. This eliminated 86% of formula read overhead.
4. **A Python cell cache** prevents redundant FFI round-trips when the same cell is accessed multiple times (common in verification and formatting workflows).

The PyO3 boundary passes plain dicts across the FFI — `{"type": "string", "value": "Hello"}` for values, `{"bold": true, "font_color": "#FF0000"}` for formats. This keeps the API simple and the serialization cost low.

## Getting Started

WolfXL provides an openpyxl-compatible API so you can switch with minimal code changes:

```python
from wolfxl import load_workbook, Workbook, Font, PatternFill

# Reading
wb = load_workbook("report.xlsx")
ws = wb["Sheet1"]
print(ws["A1"].value)           # cell value
print(ws["A1"].font.bold)       # True/False
print(ws["A1"].fill.fgColor)    # hex color string

# Writing
wb = Workbook()
ws = wb.active
ws["A1"] = "Revenue"
ws["A1"].font = Font(bold=True, color="FF0000")
ws["B1"] = 42_000
ws["B1"].number_format = "$#,##0"
wb.save("output.xlsx")
```

The `load_workbook()` function wraps calamine for reading. `Workbook()` wraps rust_xlsxwriter for writing. Both present the familiar `wb['Sheet1']['A1'].value` interface.

## Limitations

Transparency matters more than marketing:

- **No images.** Neither calamine nor rust_xlsxwriter supports embedded images yet.
- **No read-modify-write.** You can read OR write, not open an existing file, modify it, and save. This is a fundamental limitation of the hybrid approach — calamine is a reader, rust_xlsxwriter is a writer.
- **Diagonal borders.** Read support has a gap for diagonal border detection (score 1/3 on that sub-test).
- **Requires prebuilt wheels.** The Rust compilation step means you need maturin or prebuilt wheels. We're working on PyPI distribution.
- **xlsx only.** No .xls (legacy binary format) or .xlsb support.

## Try It

- **Live dashboard**: [excelbench.vercel.app](https://excelbench.vercel.app) — interactive fidelity heatmap + performance charts
- **ExcelBench repo**: [github.com/SynthGL/ExcelBench](https://github.com/SynthGL/ExcelBench) — run the benchmarks yourself
- **Reproduce**: `uv run excelbench benchmark --tests fixtures/excel --output results` then `uv run excelbench html`

The data is open. The methodology is reproducible. If your Python code touches Excel at scale, the numbers speak for themselves.
