# pyumya

**Fast Excel file manipulation for Python, powered by Rust.**

pyumya is a Python wrapper around [umya-spreadsheet](https://github.com/MathNya/umya-spreadsheet),
a pure Rust library for reading and writing `.xlsx` and `.xlsm` files. It provides the speed of
Rust with the ergonomics of Python.

!!! note "Source of truth (current)"
    The current API documented here is the in-repo `wolfxl._rust.UmyaBook` class.
    A standalone PyPI package named `pyumya` is planned, but not the primary source of truth yet.

## Why pyumya?

| Feature | openpyxl | python-calamine | XlsxWriter | **pyumya** |
|---------|:--------:|:---------------:|:----------:|:----------:|
| Read .xlsx | Yes | Yes | No | **Yes** |
| Write .xlsx | Yes | No | Yes | **Yes** |
| Formatting R/W | Yes | No | Write only | **Yes** |
| Rust speed | No | Yes (read) | No | **Yes** |
| Precompiled wheels | N/A | Yes | N/A | **Planned** |

## Installation

```bash
# From ExcelBench (current, requires Rust toolchain)
uv sync --extra rust
uv run maturin develop --manifest-path rust/excelbench_rust/Cargo.toml --features umya

# Standalone PyPI package (planned)
# pip install pyumya
```

## Quick Start

### Reading an Excel file

```python
from wolfxl._rust import UmyaBook

# Open a workbook
book = UmyaBook.open("report.xlsx")

# List sheets
print(book.sheet_names())  # ["Sheet1", "Summary"]

# Read a cell value
value = book.read_cell_value("Sheet1", "A1")
print(value)  # {"type": "string", "value": "Hello"}

# Read cell formatting
fmt = book.read_cell_format("Sheet1", "A1")
print(fmt)  # {"bold": True, "font_size": 14.0}

# Read cell borders
border = book.read_cell_border("Sheet1", "B2")
print(border)  # {"top": {"style": "thin", "color": "#000000"}}
```

### Writing an Excel file

```python
from wolfxl._rust import UmyaBook

# Create a new workbook
book = UmyaBook()
book.add_sheet("Data")

# Write cell values
book.write_cell_value("Data", "A1", {"type": "string", "value": "Name"})
book.write_cell_value("Data", "B1", {"type": "number", "value": 42.0})
book.write_cell_value("Data", "C1", {"type": "boolean", "value": True})
book.write_cell_value("Data", "D1", {"type": "date", "value": "2026-01-15"})

# Apply formatting
book.write_cell_format("Data", "A1", {
    "bold": True,
    "font_size": 14.0,
    "font_color": "#FF0000",
})

# Apply borders
book.write_cell_border("Data", "A1", {
    "top": {"style": "thin", "color": "#000000"},
    "bottom": {"style": "double", "color": "#0000FF"},
})

# Set dimensions
book.set_row_height("Data", 0, 30.0)    # row index is 0-based
book.set_column_width("Data", "A", 25.0)

# Save
book.save("output.xlsx")
```

## Current API Coverage

`wolfxl._rust.UmyaBook` currently exposes **30+ methods** covering:

- **File I/O**: Create, open, save workbooks
- **Sheet management**: List and add sheets
- **Cell values**: Read/write 7 types (string, number, boolean, formula, error, date, datetime)
- **Cell formatting**: Read/write font, fill, alignment, number format properties
- **Borders**: Read/write all edge styles and colors
- **Dimensions**: Read/write row heights and column widths
- **Phase 1 workbook features**: Merged cells, comments, hyperlinks, freeze panes, images (anchor metadata),
  data validation, conditional formatting (limited write support)
- **Phase 2 workbook features (early)**: Named ranges, tables, and worksheet auto filters

See the [API Reference](api-reference.md) for details and the
[Limitations](limitations.md) page for what's not yet supported.

## Roadmap

We're actively expanding the API to match the feature set of the
[Elixir umya_spreadsheet_ex wrapper](https://hexdocs.pm/umya_spreadsheet_ex/).
See the [parity tracker](https://github.com/wolfgangschoenberger/ExcelBench/blob/master/docs/trackers/umya-python-parity-tracker.md)
for progress.

**Phase 1** (in progress): Merged cells, comments, hyperlinks, freeze panes, images,
data validation, conditional formatting.

**Phase 2**: Named ranges, tables, auto filters, rich text, charts, pivot tables.

**Phase 3**: Standalone PyPI package with precompiled wheels.
