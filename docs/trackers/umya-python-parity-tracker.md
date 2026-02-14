# umya-python Parity Tracker

Created: 02/13/2026 03:07 AM PST
Reference: [umya_spreadsheet_ex (Elixir)](https://github.com/alexiob/umya_spreadsheet_ex) v0.7.0

## Goal

Bring our PyO3 `UmyaBook` bindings (and eventually a standalone `umya-python` package)
to feature parity with the Elixir `umya_spreadsheet_ex` NIF wrapper. This means matching
both **API surface** (every function the Elixir wrapper exposes) and **documentation
quality** (guides, examples, troubleshooting, limitations page).

Quick links:
- Current Rust backend: `rust/excelbench_rust/src/umya/` (split from monolith, 6 modules)
- Python adapter: `src/excelbench/harness/adapters/umya_adapter.py`
- Elixir source: 35 modules in `lib/umya_spreadsheet/`
- Elixir docs: https://hexdocs.pm/umya_spreadsheet_ex/
- **pyumya docs skeleton**: `docs/pyumya/` (mkdocs-material)

---

## Key Decisions

### Decision 1: Scope — Standalone `pyumya` package

**Chose standalone over ExcelBench-only** because:
- The niche is completely empty — no Python wrapper for umya-spreadsheet on PyPI
- python-calamine proves the model works (PyO3 + maturin + precompiled wheels → 1.6M downloads/mo)
- ExcelBench becomes a consumer of pyumya, not its container (cleaner architecture)
- Benchmark scores improve as a side effect

**Strategy**: Develop in ExcelBench repo for now, extract to standalone repo at T3.2 milestone.
Design API as if standalone from day one.

### Decision 2: Phase 1 ordering — confirmed ascending LOC

Order by benchmark-score-per-LOC, clustered by API similarity. See Implementation Order below.

### Decision 3: Doc infra — mkdocs-material skeleton, grow organically

Start with 3 pages (index, api-reference, limitations), add guide pages as Phase 1 features land.
Located at `docs/pyumya/`. By end of Phase 1 → ~10 pages organically.

### Decision 4: Module split — do it BEFORE Phase 1

Split the 795-line monolith before adding ~510 LOC of Phase 1 features. Avoids
1,300-line file, makes each feature a self-contained PR targeting its own module file.

---

## Summary Dashboard

| Category | Elixir functions | Our functions | Coverage | Status |
|----------|:---:|:---:|:---:|:---:|
| File I/O (open/save/new) | 6 | 3 | 50% | Partial |
| Sheet management | 20 | 2 | 10% | Stub |
| Cell values | 6 | 2 | 33% | Partial |
| Font formatting | 10+ | 7 (monolithic) | ~60% | Partial |
| Background/fill | 5 | 2 (monolithic) | 40% | Partial |
| Borders | 6 | 2 | 33% | Done (in scope) |
| Alignment | 8 | 4 (monolithic) | 50% | Partial |
| Number formats | 4 | 2 (monolithic) | 50% | Partial |
| Row/column ops | 12+ | 4 | 33% | Partial |
| Merged cells | 2 | 2 | 100% | Done |
| Comments | 6 | 2 | 33% | Partial |
| Hyperlinks | 6+ | 2 | 33% | Partial |
| Images | 4+ | 2 | 50% | Partial |
| Charts | 2+ | 0 | 0% | Missing |
| Conditional formatting | 6+ | 2 | 33% | Partial |
| Data validation | 6+ | 2 | 33% | Partial |
| Formulas (named/array) | 6+ | 3 (basic + named ranges) | ~45% | Partial |
| Auto filters | 4+ | 4 | 100% | Done |
| Tables (ListObjects) | 8+ | 2 | 25% | Partial |
| Pivot tables | 4+ | 0 | 0% | Missing |
| Rich text | 6+ | 0 | 0% | Missing |
| Print/page setup | 10+ | 0 | 0% | Missing |
| Page breaks | 3+ | 0 | 0% | Missing |
| Sheet views (freeze, grid) | 5+ | 2 | 40% | Partial |
| Protection (sheet/wb) | 4+ | 0 | 0% | Missing |
| Document properties | 4+ | 0 | 0% | Missing |
| CSV export | 3+ | 0 | 0% | Missing |
| Performance (lazy/light) | 3+ | 0 | 0% | Missing |
| Shapes/drawings | 4+ | 0 | 0% | Missing |
| Workbook views | 3+ | 0 | 0% | Missing |
| **Documentation** | 15 guide pages | 3 | 17% | Skeleton |

**Overall API coverage: ~25%** (37 of ~150+ functions)

---

## Tier 0 — ExcelBench Immediate Wins

These gaps directly block ExcelBench Tier 2 scoring. The Rust library (umya-spreadsheet)
already supports them; we just need PyO3 bindings + Python adapter wiring.

### 0.1 Merged Cells R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `sheet_functions.ex` |
| **Elixir functions** | `add_merge_cells/3`, `get_merge_cells/2` |
| **umya-spreadsheet API** | `worksheet.add_merge_cells()`, `worksheet.get_merge_cells()` |
| **Rust effort** | ~30 LOC — iterate merge collection, return `Vec<String>` |
| **Python adapter** | Wire `read_merged_ranges()` and `merge_cells()` (currently return `[]`/no-op) |
| **ExcelBench impact** | Unlocks `merged_cells` feature scoring for umya adapter |
| **Status** | [x] Done |

### 0.2 Comments R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `comment_functions.ex` |
| **Elixir functions** | `add_comment/5`, `get_comment/3`, `update_comment/4-5`, `remove_comment/3`, `has_comments/2`, `get_comments_count/2` |
| **umya-spreadsheet API** | `cell.get_comment()`, `cell.set_comment()`, `Comment::new()` |
| **Rust effort** | ~50 LOC — read comment text/author, write comment with author |
| **Python adapter** | Wire `read_comments()` and `add_comment()` (currently return `[]`/no-op) |
| **ExcelBench impact** | Unlocks `comments` feature scoring |
| **Status** | [x] Done |

### 0.3 Hyperlinks R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `hyperlink.ex` |
| **Elixir functions** | `add_hyperlink/4-5`, `get_hyperlinks/2`, `get_hyperlink/3`, `update_hyperlink/4-5`, `remove_hyperlink/3`, `remove_all_hyperlinks/2` |
| **umya-spreadsheet API** | `worksheet.get_hyperlink_collection()`, `Hyperlink::new()`, `cell.set_hyperlink()` |
| **Rust effort** | ~60 LOC — iterate hyperlink collection, create/attach hyperlinks |
| **Python adapter** | Wire `read_hyperlinks()` and `add_hyperlink()` (currently return `[]`/no-op) |
| **ExcelBench impact** | Unlocks `hyperlinks` feature scoring |
| **Status** | [x] Done |

### 0.4 Freeze Panes R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `sheet_view_functions.ex` |
| **Elixir functions** | `set_freeze_panes/4`, `get_freeze_panes/2`, sheet view configuration |
| **umya-spreadsheet API** | `SheetView::get_pane()`, `Pane::set_*()` |
| **Rust effort** | ~40 LOC — read/write pane split position and frozen state |
| **Python adapter** | Wire `read_freeze_panes()` and `set_freeze_panes()` (currently return `{}`/no-op) |
| **ExcelBench impact** | Unlocks `freeze_panes` feature scoring |
| **Status** | [x] Done |

### 0.5 Images R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `image_functions.ex` |
| **Elixir functions** | `add_image/4-6`, `get_images/2`, `get_image/3`, `remove_image/3` |
| **umya-spreadsheet API** | `worksheet.get_image_collection()`, `Image::new()`, anchoring |
| **Rust effort** | ~80 LOC — image bytes, anchor coordinates, format detection |
| **Python adapter** | Wire `read_images()` and `add_image()` (currently return `[]`/no-op) |
| **ExcelBench impact** | Unlocks `images` feature scoring |
| **Status** | [x] Done |

### 0.6 Data Validation R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `data_validation.ex` |
| **Elixir functions** | `add_data_validation_*` (list, whole, decimal, date, text_length, custom, formula), `get_data_validations/2`, `remove_data_validation/3` |
| **umya-spreadsheet API** | `worksheet.get_data_validations()`, `DataValidation::new()`, various validation types |
| **Rust effort** | ~100 LOC — parse validation type/operator/formula/ranges |
| **Python adapter** | Wire `read_data_validations()` and `add_data_validation()` (currently return `[]`/no-op) |
| **ExcelBench impact** | Unlocks `data_validation` feature scoring |
| **Status** | [x] Done |

### 0.7 Conditional Formatting R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `conditional_formatting_functions.ex` |
| **Elixir functions** | `add_cell_value_rule/7`, `add_color_scale/5-7`, `add_data_bar/5`, `add_top_bottom_rule/6`, `add_text_rule/6`, `get_conditional_formatting/2`, `remove_conditional_formatting/3` |
| **umya-spreadsheet API** | `worksheet.get_conditional_formatting_collection()`, `ConditionalFormatting::new()`, rule types |
| **Rust effort** | ~150 LOC — complex: multiple rule types, operators, color specs |
| **Python adapter** | Wire `read_conditional_formats()` and `add_conditional_format()` (currently return `[]`/no-op) |
| **ExcelBench impact** | Unlocks `conditional_format` feature scoring |
| **Status** | [x] Done |

**Tier 0 total: ~510 LOC Rust + adapter wiring — ALL COMPLETE**
**Impact: umya adapter now supports full Tier 0-2 scoring (7 new features)**

---

## Tier 1 — Feature Parity (Tier 3 + General Use)

These fill out the "power user" API surface that the Elixir wrapper provides.

### 1.1 Named Ranges R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `formula_functions.ex` |
| **Elixir functions** | `add_defined_name/4`, `get_defined_names/1`, `remove_defined_name/2` |
| **umya-spreadsheet API** | `spreadsheet.get_defined_names()`, `Worksheet::add_defined_name()` |
| **Rust module** | `umya/named_ranges.rs` (~137 LOC) |
| **Python adapter** | `read_named_ranges()` and `add_named_range()` wired |
| **ExcelBench score** | Read: 3, Write: 3 |
| **Key insight** | `DefinedName::set_name()` is `pub(crate)` — workaround via `Worksheet::add_defined_name()` which calls it internally. umya disperses workbook-level names to worksheets on read; dual collection needed to avoid duplicates. |
| **Status** | [x] Done |

### 1.2 Tables (ListObjects) R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `table.ex` |
| **Elixir functions** | `add_table/5-6`, `get_tables/2`, `get_table/3`, `remove_table/3`, `add_table_column/4`, `set_table_style/4`, `set_totals_row/4`, `get_table_data_range/3` |
| **umya-spreadsheet API** | `worksheet.get_tables()`, `Table::new()`, `TableColumn::new()`, `TableStyleInfo::new()` |
| **Rust module** | `umya/tables.rs` (~160 LOC) |
| **Python adapter** | `read_tables()` and `add_table()` wired |
| **ExcelBench score** | Read: 2, Write: 3 |
| **Key insight** | Read score is 2 (not 3) because umya doesn't expose autoFilter state on Table — `tbl_autofilter` edge test fails. All public APIs (Table, TableColumn, TableStyleInfo) are fully accessible. |
| **Status** | [x] Done |

### 1.3 Auto Filters R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `auto_filter_functions.ex` |
| **Elixir functions** | `set_auto_filter/3`, `remove_auto_filter/2`, `get_auto_filter/2`, `has_auto_filter/2` |
| **umya-spreadsheet API** | `worksheet.get_auto_filter()`, `ws.set_auto_filter()`, `ws.remove_auto_filter()` |
| **Rust module** | `umya/auto_filter.rs` (~53 LOC) |
| **Python methods** | `get_auto_filter()`, `set_auto_filter()`, `remove_auto_filter()`, `has_auto_filter()` |
| **Key insight** | Worksheet-level auto filter only — umya Table struct doesn't expose table-level autoFilter. Tables write path sets worksheet AF as heuristic. |
| **Status** | [x] Done |

### 1.4 Rich Text R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `rich_text.ex` |
| **Elixir functions** | `set_rich_text/4`, `get_rich_text/3`, `create_rich_text/1`, `add_text_element/3`, `rich_text_from_html/1`, `rich_text_to_html/1` |
| **umya-spreadsheet API** | `RichText`, `TextElement`, HTML conversion |
| **Rust effort** | ~100 LOC |
| **Status** | [ ] Not started |

### 1.5 Array Formulas

| Item | Detail |
|------|--------|
| **Elixir module** | `formula_functions.ex` |
| **Elixir functions** | `set_array_formula/4` |
| **umya-spreadsheet API** | `cell.set_formula()` with array flag |
| **Rust effort** | ~20 LOC (extend existing formula handling) |
| **Status** | [ ] Not started |

### 1.6 Charts R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `chart_functions.ex` |
| **Elixir functions** | `add_chart/9`, `add_chart_with_options/16` |
| **Supported types** | Line, Bar, Pie, Area, Scatter, Doughnut, Radar, OfPie, Bar3D, Line3D, Pie3D |
| **umya-spreadsheet API** | Chart creation with series, axes, legends, 3D views |
| **Rust effort** | ~200 LOC (complex: multiple chart types, series config, positioning) |
| **Status** | [ ] Not started |

### 1.7 Pivot Tables R/W

| Item | Detail |
|------|--------|
| **Elixir module** | `pivot_table.ex` |
| **Elixir functions** | `add_pivot_table/7`, `get_pivot_tables/2`, `refresh_pivot_table/3`, `remove_pivot_table/3` |
| **umya-spreadsheet API** | `PivotTable`, field configuration, positioning |
| **Rust effort** | ~150 LOC |
| **Note** | Elixir docs flag this as "basic support only" |
| **Status** | [ ] Not started |

**Tier 1 progress: 3/7 done (~350 LOC shipped), ~320 LOC remaining**

---

## Tier 2 — General-Purpose Library Completeness

These make the library useful beyond benchmarking — a real replacement for openpyxl in
Rust-speed scenarios.

### 2.1 Sheet Management (full)

| Item | Detail |
|------|--------|
| **Elixir functions** | `clone_sheet/3`, `remove_sheet/2`, `rename_sheet/3`, `set_sheet_state/3`, `get_sheet_state/2`, `get_sheet_count/1`, `get_active_sheet/1`, `set_sheet_protection/4`, `get_sheet_protection/2` |
| **Current** | `add_sheet()`, `sheet_names()` only |
| **Rust effort** | ~80 LOC |
| **Status** | [ ] Not started |

### 2.2 Row/Column Operations (full)

| Item | Detail |
|------|--------|
| **Elixir functions** | `insert_new_row/4`, `remove_row/4`, `insert_new_column/4`, `remove_column/4`, `set_column_auto_width/4`, `get_column_auto_width/3`, `get_row_hidden/3`, `get_column_hidden/3`, `set_row_style/5`, `copy_row_styling/4-6`, `copy_column_styling/4-6` |
| **Current** | `read_row_height`, `read_column_width`, `set_row_height`, `set_column_width` only |
| **Rust effort** | ~120 LOC |
| **Status** | [ ] Not started |

### 2.3 Print Settings / Page Setup

| Item | Detail |
|------|--------|
| **Elixir module** | `print_settings_functions.ex` |
| **Elixir functions** | orientation, paper size, margins, headers/footers, print area, repeating rows/cols, fit-to-page, centering, scaling |
| **Rust effort** | ~100 LOC |
| **Status** | [ ] Not started |

### 2.4 Protection (Workbook + Sheet)

| Item | Detail |
|------|--------|
| **Elixir modules** | `workbook_protection_functions.ex`, `workbook_functions.ex` |
| **Elixir functions** | `set_password/3`, `set_workbook_protection/2`, sheet-level protection |
| **Rust effort** | ~40 LOC |
| **Status** | [ ] Not started |

### 2.5 Document Properties

| Item | Detail |
|------|--------|
| **Elixir module** | `document_properties.ex` |
| **Functions** | title, author, subject, description, keywords, category |
| **Rust effort** | ~40 LOC |
| **Status** | [ ] Not started |

### 2.6 CSV Export

| Item | Detail |
|------|--------|
| **Elixir module** | `csv_functions.ex` |
| **Functions** | `write_csv/3-4` with delimiter/encoding options |
| **Rust effort** | ~40 LOC |
| **Status** | [ ] Not started |

### 2.7 Page Breaks

| Item | Detail |
|------|--------|
| **Elixir module** | `page_breaks.ex` |
| **Functions** | `add_page_break_row/3`, `add_page_break_column/3`, `get_page_breaks/2` |
| **Rust effort** | ~30 LOC |
| **Status** | [ ] Not started |

### 2.8 Workbook/Sheet Views

| Item | Detail |
|------|--------|
| **Elixir modules** | `workbook_view_functions.ex`, `sheet_view_functions.ex` |
| **Functions** | active tab, window position/size, grid lines visibility, cell selection, zoom |
| **Rust effort** | ~60 LOC |
| **Status** | [ ] Not started |

### 2.9 Performance Modes

| Item | Detail |
|------|--------|
| **Elixir module** | `performance_functions.ex`, `csv_writer_option.ex` |
| **Functions** | Lazy reading (deferred sheet loading), light writer (streaming), compression level |
| **umya-spreadsheet API** | `reader::xlsx::lazy_read()`, `writer::xlsx::write_with_option()` |
| **Rust effort** | ~50 LOC |
| **Status** | [ ] Not started |

### 2.10 Shapes / Drawings

| Item | Detail |
|------|--------|
| **Elixir modules** | `drawing.ex`, `vml_drawing.ex` |
| **Functions** | Rectangles, circles, arrows, text boxes, connectors |
| **Rust effort** | ~120 LOC (complex positioning + styling) |
| **Status** | [ ] Not started |

### 2.11 Cell-Level Protection Properties

| Item | Detail |
|------|--------|
| **Elixir functions** | `get_cell_locked/3`, `get_cell_hidden/3`, `set_cell_indent/4` |
| **Rust effort** | ~20 LOC (extend existing format handling) |
| **Status** | [ ] Not started |

### 2.12 Formatted Value Retrieval

| Item | Detail |
|------|--------|
| **Elixir functions** | `get_formatted_value/3` (returns value "as displayed in Excel") |
| **umya-spreadsheet API** | `cell.get_formatted_value()` |
| **Rust effort** | ~30 LOC |
| **Status** | [ ] Not started |

### 2.13 File Format Options

| Item | Detail |
|------|--------|
| **Elixir module** | `file_format_options.ex` |
| **Functions** | Compression level (0-9), encryption, password-protected save, binary format |
| **Rust effort** | ~40 LOC |
| **Status** | [ ] Not started |

**Tier 2 total estimate: ~770 LOC Rust**

---

## Tier 3 — Architecture & Packaging

Structural work to make this a proper standalone library (not just ExcelBench bindings).

### 3.1 Module Splitting

| Item | Detail |
|------|--------|
| **Current** | `umya_backend.rs` was one 795-line monolith |
| **Target** | Split into `umya/` directory with focused modules |
| **Initial split (6 files)** | `mod.rs` (struct + I/O), `cell_values.rs`, `formatting.rs`, `borders.rs`, `dimensions.rs`, `util.rs` |
| **Phase 1 additions** | `merged_cells.rs`, `comments.rs`, `hyperlinks.rs`, `freeze_panes.rs`, `images.rs`, `data_validation.rs`, `conditional_fmt.rs` |
| **Status** | [x] **Done** — 13 modules: mod.rs + 12 feature modules |

### 3.2 Standalone Crate

| Item | Detail |
|------|--------|
| **Current** | `excelbench_rust` crate — tightly coupled to benchmark project |
| **Target** | Separate `umya-python` (or `pyumya`) crate with its own repo, Cargo.toml, CI |
| **Decisions needed** | Package name, repo location, relationship to ExcelBench |
| **Status** | [ ] Not started |

### 3.3 Precompiled Wheels

| Item | Detail |
|------|--------|
| **Current** | Requires local `maturin develop` build |
| **Target** | maturin CI → PyPI with wheels for: linux-x86_64, linux-aarch64, macos-x86_64, macos-arm64, windows-x86_64 |
| **Elixir equivalent** | Rustler precompiled NIF binaries for 6 targets |
| **Status** | [ ] Not started |

### 3.4 Thread Safety

| Item | Detail |
|------|--------|
| **Current** | `#[pyclass(unsendable)]` — single-threaded only |
| **Elixir** | Documented concurrent usage (NIF scheduler handles this) |
| **Target** | Evaluate `Send + Sync` feasibility; at minimum document limitations |
| **Status** | [ ] Not started |

### 3.5 Error Handling Consistency

| Item | Detail |
|------|--------|
| **Current** | Mix of `PyResult<T>`, `PyValueError`, `PyIOError` |
| **Elixir** | Consistent `{:ok, result}` / `{:error, reason}` tuples on every function |
| **Target** | Typed Python exceptions: `UmyaIOError`, `UmyaValueError`, `UmyaSheetNotFound`, etc. |
| **Status** | [ ] Not started |

### 3.6 Type Stubs (.pyi)

| Item | Detail |
|------|--------|
| **Current** | No type stubs — IDE sees `Any` for all PyO3 classes |
| **Target** | `umya_python.pyi` stub file with full type annotations for IDE autocomplete |
| **Status** | [ ] Not started |

---

## Tier D — Documentation Parity

The Elixir wrapper's documentation is a gold standard we want to match. Their doc site
includes **15 guide pages**, **3 API reference modules**, a **limitations page**, and a
**troubleshooting guide** — all with executable code examples.

### Documentation Quality Bar (from Elixir wrapper)

**What they do well:**
- Every function has a code example showing input → output
- Progressive complexity: basic example → advanced options
- Error handling patterns shown (`case` / pattern matching)
- Best practices section per feature (e.g., "keep comments concise, include dates")
- Explains Excel integration behavior (e.g., "hover tooltip shows comment")
- Limitations page is honest about what's NOT supported (3D charts, macros, OLE)
- Troubleshooting guide for common issues
- Consistent formatting: function signature → description → example → notes

**What we need to match:**
- Every function must have a Python code example
- Show both simple and advanced usage
- Include error handling patterns
- Document return types and edge cases
- Honest limitations page

### D.1 Guide Pages (matching Elixir's 15 guides)

| # | Guide Page | Elixir equivalent | Status |
|---|-----------|-------------------|--------|
| 1 | Getting Started / Installation | README | [x] Done — `docs/pyumya/docs/index.md` |
| 2 | Cell Operations | (README inline examples) | [ ] Not started |
| 3 | Styling and Formatting | `styling_formatting` guide | [ ] Not started |
| 4 | Sheet Operations | `sheet_operations` guide | [ ] Not started |
| 5 | Formula Functions | `formula_functions` guide | [ ] Not started |
| 6 | Comments | `comments` guide | [x] Stub — `guides/comments.md` |
| 7 | Hyperlinks | (README section) | [x] Stub — `guides/hyperlinks.md` |
| 8 | Images | `image_handling` guide | [x] Stub — `guides/images.md` |
| 9 | Charts | `charts` guide | [ ] Not started |
| 10 | Data Validation | `data_validation` guide | [x] Stub — `guides/data-validation.md` |
| 11 | Conditional Formatting | (README section) | [x] Stub — `guides/conditional-formatting.md` |
| 12 | Excel Tables | `excel_tables` guide | [ ] Not started |
| 13 | Pivot Tables | `pivot_tables` guide | [ ] Not started |
| 14 | Auto Filters | `auto_filters` guide | [ ] Not started |
| 15 | Print Settings | `print_settings` guide | [ ] Not started |
| 16 | CSV Export & Performance | `csv_performance` guide | [ ] Not started |
| 17 | File Format Options | `file_format_options` guide | [ ] Not started |
| 18 | Window Settings | `window_settings` guide | [ ] Not started |

### D.2 Reference Pages

| # | Page | Elixir equivalent | Status |
|---|------|-------------------|--------|
| 1 | API Reference (all functions) | 3 module pages on HexDocs | [x] Done — `docs/pyumya/docs/api-reference.md` |
| 2 | Limitations & Compatibility | `limitations.html` | [x] Done — `docs/pyumya/docs/limitations.md` |
| 3 | Troubleshooting | `troubleshooting.html` | [ ] Not started |
| 4 | Changelog | Standard | [ ] Not started |

### D.3 Documentation Infrastructure

| # | Item | Target | Status |
|---|------|--------|--------|
| 1 | Doc framework | mkdocs-material | [x] Done — `docs/pyumya/mkdocs.yml` |
| 2 | Auto-generated API reference | From .pyi stubs or docstrings | [ ] Not started |
| 3 | CI doc build | Build docs on every PR | [ ] Not started |
| 4 | Hosted docs | ReadTheDocs or GitHub Pages | [ ] Not started |

### D.4 Documentation Quality Checklist (per function)

Every public function must have:

- [ ] One-line description
- [ ] Parameter documentation with types
- [ ] Return type documentation
- [ ] Basic code example (3-5 lines)
- [ ] Advanced example (if function has optional params)
- [ ] Error handling example (what happens on bad input)
- [ ] Edge case notes (empty strings, max values, etc.)

---

## Effort Summary

| Tier | Scope | Est. Rust LOC | Priority |
|------|-------|:---:|:---:|
| **T0** | ExcelBench scoring (7 features) | ~510 | **P0** — immediate |
| **T1** | Power features (7 features) | ~670 | P1 — near-term |
| **T2** | General completeness (13 items) | ~770 | P2 — medium-term |
| **T3** | Architecture/packaging (6 items) | N/A (infra) | P2 — medium-term |
| **TD** | Documentation (18 guides + 4 ref + infra) | N/A (prose) | P1 — parallel |
| **Total** | | ~1,950 LOC | |

Current: ~1,650 LOC (795 base + ~510 T0 + ~350 T1) → Target: ~2,750 LOC (1.7x remaining growth) + docs + packaging

---

## Implementation Order (Revised per Decisions)

### Phase 0: Foundation (before Phase 1)
1. [x] **T3.1 Module splitting** — split `umya_backend.rs` → `umya/` (6 files). Codex handoff in progress.
2. [x] **D.3 Doc infrastructure** — mkdocs-material skeleton at `docs/pyumya/` (3 pages: index, api-reference, limitations)

### Phase 1: ExcelBench Scoring (T0)
Each feature → new `.rs` file in `umya/` + adapter wiring + guide page in docs.

1. 0.1 Merged cells (easiest, ~30 LOC) → `umya/merged_cells.rs`
2. 0.4 Freeze panes (~40 LOC) → `umya/freeze_panes.rs`
3. 0.2 Comments (~50 LOC) → `umya/comments.rs`
4. 0.3 Hyperlinks (~60 LOC) → `umya/hyperlinks.rs`
5. 0.5 Images (~80 LOC) → `umya/images.rs`
6. 0.6 Data validation (~100 LOC) → `umya/data_validation.rs`
7. 0.7 Conditional formatting (~150 LOC) → `umya/conditional_fmt.rs`

### Phase 2: Power Features (T1)
1. [x] 1.1 Named ranges (~137 LOC) → `umya/named_ranges.rs` — Score: R3/W3
2. [x] 1.3 Auto filters (~53 LOC) → `umya/auto_filter.rs` — Worksheet-level AF + table AF heuristic
3. [ ] 1.5 Array formulas (~20 LOC) → extend `umya/cell_values.rs`
4. [ ] 1.4 Rich text (~100 LOC) → `umya/rich_text.rs`
5. [x] 1.2 Tables (~160 LOC) → `umya/tables.rs` — Score: R2/W3
6. [ ] 1.7 Pivot tables (~150 LOC) → `umya/pivot_tables.rs`
7. [ ] 1.6 Charts (~200 LOC) → `umya/charts.rs`

### Phase 3: General Library (T2) + Packaging (T3)
1. 2.1 Sheet management → extend `umya/mod.rs`
2. 2.2 Row/column ops → extend `umya/dimensions.rs`
3. 2.4 Protection → `umya/protection.rs`
4. 2.5 Document properties → `umya/properties.rs`
5. 2.8 Workbook/sheet views → `umya/views.rs`
6. 2.3 Print settings → `umya/print_setup.rs`
7. 2.6 CSV export → `umya/csv_export.rs`
8. 2.7 Page breaks → `umya/page_breaks.rs`
9. 2.9 Performance modes → extend `umya/mod.rs`
10. 2.12 Formatted values → extend `umya/cell_values.rs`
11. 2.13 File format options → extend `umya/mod.rs`
12. 2.10 Shapes (last — highest effort, lowest priority) → `umya/shapes.rs`
13. 2.11 Cell-level protection → extend `umya/formatting.rs`
14. 3.2-3.6 Standalone crate, wheels, thread safety, error handling, stubs

### Phase D: Documentation (parallel with all phases)
- [x] D.3 infrastructure — mkdocs-material skeleton created
- [x] D.2 limitations page — `docs/pyumya/docs/limitations.md` created (honest from day one)
- [x] D.1 guide stubs — 7 T0 guide pages created (merged cells, comments, hyperlinks, freeze panes, images, data validation, conditional formatting)
- [ ] D.4 quality checklist — enforce from Phase 2 onward

---

## Session Log (append-only)

### 02/13/2026 03:07 AM PST
- Created: Initial tracker from gap analysis vs Elixir umya_spreadsheet_ex v0.7.0
- Inventoried: 35 Elixir modules, ~150+ functions, 15 guide pages
- Our baseline: 15 methods in `umya_backend.rs`, ~15% API coverage
- Key insight: Elixir wrapper treats umya as general-purpose library; we treat it as benchmark adapter
- Decision: Track both API parity AND documentation parity as first-class goals
- Next: Decide scope (standalone package vs ExcelBench-only) → begin Phase 1

### 02/13/2026 03:20 AM PST
- **Decision 1**: Standalone `pyumya` package (empty niche, python-calamine proves model)
- **Decision 2**: Phase 1 ordering confirmed (ascending LOC, clustered by API similarity)
- **Decision 3**: mkdocs-material skeleton — 3 pages created (`index.md`, `api-reference.md`, `limitations.md`)
- **Decision 4**: Module split BEFORE Phase 1 (avoid 1,300-line monolith)
- **Executed**: T3.1 module split handed off to Codex (gpt-5.3-codex, xhigh reasoning)
- **Executed**: D.3 doc infrastructure created at `docs/pyumya/`
- **Created**: `docs/pyumya/docs/index.md` — getting started with quick examples
- **Created**: `docs/pyumya/docs/api-reference.md` — full API reference for current 15 methods
- **Created**: `docs/pyumya/docs/limitations.md` — honest limitations page (matches Elixir's pattern)
- **Updated**: Tracker with all 4 decisions, revised implementation order
- Next: Verify Codex module split → commit → begin Phase 1 (T0.1 merged cells)

### 02/13/2026 04:20 AM PST
- **Executed**: D.3 doc infrastructure completed — `mkdocs-material>=9.5` + `mkdocs-minify-plugin` added to `pyproject.toml[docs]`
- **Executed**: `uv sync --extra docs` — all deps installed, `mkdocs build --strict` passes
- **Created**: 7 guide stubs for T0 features in `docs/pyumya/docs/guides/`:
  - `merged-cells.md`, `comments.md`, `hyperlinks.md`, `freeze-panes.md`
  - `images.md`, `data-validation.md`, `conditional-formatting.md`
- **Updated**: `mkdocs.yml` — full nav with Guides section, minify plugin, emoji extension, social links
- **Updated**: `.gitignore` — added `docs/pyumya/site/`
- **Verified**: Dev server returns 200 on all pages (index, guides, api-reference, limitations)
- Guide stubs contain anticipated API examples — will need updating once T0 Rust methods land
- Next: Update guide pages as T0 features are implemented by parallel agent

### 02/13/2026 ~02:00 PM PST
- **Phase 2 (Tier 3 features) COMPLETE**: Named ranges + Tables implemented end-to-end
- **Rust modules created** (both compile with zero warnings):
  - `named_ranges.rs` (~137 LOC) — read/write with workbook + sheet scoping
  - `tables.rs` (~160 LOC) — read/write with columns, styles, totals row
- **Python adapter wired**: `read_named_ranges`, `add_named_range`, `read_tables`, `add_table`
- **Tier 3 fixtures generated**: `fixtures/excel/tier3/18_named_ranges.xlsx`, `19_tables.xlsx` (openpyxl-generated)
- **Benchmark scores**: named_ranges R3/W3, tables R2/W3
- **Runner fix**: `_normalize_named_range_refers_to()` now strips `$` from addresses for fair cross-adapter comparison
- **Key learnings**:
  - `DefinedName::set_name()` is `pub(crate)` — workaround via `Worksheet::add_defined_name()`
  - umya disperses workbook-level names to worksheets on read; need dual collection
  - Tables autoFilter state not exposed publicly by umya — hardcode `false`, accept score=2
- **Upstream issues filed**: #305 (hyperlink tooltip), #306 (CF serialization), #307 (image read)
- **Committed**: `8c63fbb` (Phase 1) + `c1e87cc` (Phase 2) — pushed to origin
- **Tests**: 1134 passed, 0 failed
- Next: Auto filters (1.3), array formulas (1.5), or standalone pyumya extraction planning

### 02/13/2026 04:05 PM PST
- **Formula read fix**: `cell_values.rs` lines 53-54 — returned raw `formula` instead of `norm` (normalized with `=` prefix). Two-line fix: `formula.to_string()` → `&norm`. Score: formulas read 0 → 3.
- **Auto filter module created**: `umya/auto_filter.rs` (~53 LOC) — 4 methods: `get_auto_filter`, `set_auto_filter`, `remove_auto_filter`, `has_auto_filter`. Matches Elixir parity (1.3).
- **Table autofilter heuristic**: Read path now checks worksheet autoFilter range overlap. Write path sets worksheet-level autoFilter when `autofilter: true`. Tables read score stays at 2 (upstream limitation: table-level autoFilter not in umya's Table struct).
- **Key learning**: `AutoFilter::get_range()` returns `&Range`, then `Range::get_range()` returns `String` — `Range` doesn't implement `Display`.
- **Tests**: 1134 passed, 0 failed
- **Benchmark**: formulas R3/W3, named_ranges R3/W3, tables R2/W3
- Next: Rich text (1.4), array formulas (1.5), or standalone pyumya extraction

### 02/13/2026 ~05:00 AM PST
- **Phase 1 COMPLETE**: All 7 T0 features implemented end-to-end
- **Rust modules created** (all compile with zero warnings):
  - `merged_cells.rs` (~30 LOC) — read/write merge cell ranges
  - `freeze_panes.rs` (~120 LOC) — read/write freeze and split panes via SheetView/Pane
  - `comments.rs` (~50 LOC) — read/write comments (RichText extraction only; Text is pub(crate))
  - `hyperlinks.rs` (~60 LOC) — read/write hyperlinks via cell.get_hyperlink()
  - `images.rs` (~50 LOC) — read/write images via anchor markers
  - `data_validation.rs` (~130 LOC) — read/write data validations with type/operator enum mapping
  - `conditional_fmt.rs` (~130 LOC) — read/write conditional formatting with style extraction
- **Python adapter wired**: All 7 read + 7 write methods now call Rust (was empty stubs)
- **Tests**: 1132 passed, 0 failed
- **Key learnings**:
  - `Text` type is `pub(crate)` — comment plain text inaccessible; use `RichText::get_text()` instead
  - `get_sheets_views()` (extra 's') for immutable SheetView access
  - Empty dict guards needed in adapter — Rust validates strictly, Python may pass `{}`
  - `PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1` required for Python 3.14+ with PyO3 0.22.6
- Next: Run fidelity benchmark to measure actual scores, then commit
