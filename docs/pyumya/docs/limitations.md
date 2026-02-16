# Limitations & Compatibility

This page documents known limitations, unsupported features, and compatibility notes.
We aim to be honest about what pyumya can and cannot do.

## File Format Support

| Format | Read | Write | Notes |
|--------|:----:|:-----:|-------|
| `.xlsx` | Yes | Yes | Full support (Office 2007+) |
| `.xlsm` | Yes | Yes | Macro-enabled; macros are preserved but not executable |
| `.xls` | No | No | Legacy format not supported by umya-spreadsheet |
| `.ods` | No | No | OpenDocument not supported |
| `.csv` | No | No | Planned (Tier 2) |

## Feature Support Matrix

### Fully Supported

| Feature | Read | Write | Since |
|---------|:----:|:-----:|-------|
| Cell values (string, number, boolean) | Yes | Yes | v0.1 |
| Formulas | Yes | Yes | v0.1 |
| Error values (#DIV/0!, #N/A, #VALUE!) | Yes | Yes | v0.1 |
| Dates and datetimes | Yes | Yes | v0.1 |
| Text formatting (bold, italic, etc.) | Yes | Yes | v0.1 |
| Background colors | Yes | Yes | v0.1 |
| Borders (all styles) | Yes | Yes | v0.1 |
| Number formats | Yes | Yes | v0.1 |
| Alignment and text rotation | Yes | Yes | v0.1 |
| Row heights and column widths | Yes | Yes | v0.1 |
| Multiple sheets | Yes | Yes | v0.1 |

### Not Yet Implemented (Planned)

These features are supported by the underlying umya-spreadsheet Rust library
but not yet exposed through pyumya's Python bindings.

| Feature | Priority | Tracking |
|---------|----------|----------|
| Rich text | P1 | T1.4 |
| Charts | P1 | T1.6 |
| Pivot tables | P1 | T1.7 |
| Sheet clone/remove/rename | P2 | T2.1 |
| Insert/remove rows/columns | P2 | T2.2 |
| Print settings | P2 | T2.3 |
| Sheet/workbook protection | P2 | T2.4 |
| Document properties | P2 | T2.5 |
| CSV export | P2 | T2.6 |
| Shapes and drawings | P2 | T2.10 |

### Newly Exposed / Partial Support

These features are now exposed in `wolfxl._rust.UmyaBook`, but some have
known limitations or partial read/write coverage.

| Feature | Read | Write | Notes |
|---------|:----:|:-----:|-------|
| Merged cells | Yes | Merge only | Unmerge not exposed |
| Comments (notes) | Yes | Yes | Text extraction is best-effort; threaded comments not supported |
| Hyperlinks | Yes | Yes | Internal links require `internal=True` |
| Freeze panes | Yes | Yes | `mode/top_left_cell` is canonical; row/column aliases supported for convenience |
| Images | Anchor metadata only | By file path | Does not expose image bytes/format on read |
| Data validation | Yes | Yes | Uses `range` + `validation_type` keys (aliases accepted) |
| Conditional formatting | Yes | Limited | Write supports rule_type/operator/formula + basic colors |
| Named ranges | Yes | Yes | Supports workbook- and sheet-scoped names |
| Tables | Yes | Yes | AutoFilter is worksheet-level; table autofilter is emulated via sheet auto filter |
| Auto filters | Yes | Yes | Can set/remove/inspect a sheet's auto filter range |

### Not Supported (umya-spreadsheet limitations)

These features are not supported by the underlying Rust library either.

| Feature | Notes |
|---------|-------|
| Macro execution | Existing macros preserved on read/write, but cannot be created or run |
| External links | References to other workbooks not supported |
| OLE objects | Embedded objects not supported |
| Smart Art | Not supported |
| 3D charts | Only 2D chart types |
| Form controls | Not supported |
| Digital signatures | Not supported |
| Power Query connections | Not supported |
| Custom add-in functions | Not supported |

## Platform Compatibility

| Platform | Architecture | Status |
|----------|:------------:|:------:|
| macOS 11+ | x86_64 | Tested |
| macOS 11+ | ARM64 (Apple Silicon) | Tested |
| Linux (glibc 2.17+) | x86_64 | Expected to work |
| Linux (glibc 2.17+) | aarch64 | Expected to work |
| Windows 10+ | x86_64 | Expected to work |

!!! note "Current installation"
    pyumya currently requires a Rust toolchain for building. Precompiled wheels
    for pip installation are planned for the standalone PyPI release.

## Python Version Support

- Python 3.11+ required (matches ExcelBench's minimum)
- PyO3 0.22 binding

## Known Issues

### Single-save limitation

`UmyaBook.save()` can only be called once per instance. The internal spreadsheet
is consumed on save. To write the same workbook to multiple locations, open or
create separate instances.

### Thread safety

`UmyaBook` is marked as `unsendable` in PyO3 and cannot be shared across Python
threads. Create separate instances per thread.

### Date system

pyumya uses the Excel 1900 date system exclusively. The 1904 date system (used by
some older macOS Excel files) is not explicitly handled — dates from such files
may be off by ~4 years.

### Row indexing

Row indices in `read_row_height`, `set_row_height` are **0-based** (matching Python
convention). This differs from Excel's 1-based row numbers and from the cell
reference format (e.g., "A1" refers to row 1).
