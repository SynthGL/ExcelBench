# Codex Sprint 1 Prompt: calamine Read + rust_xlsxwriter Write — Easy + Medium

> Historical naming note: this prompt predates the module rename. References to
> `excelbench_rust` refer to what is now imported as `wolfxl._rust`. The crate path
> `rust/excelbench_rust/` remains current.

## Context

ExcelBench is a benchmark suite that scores Excel libraries on feature fidelity. We have a Rust/PyO3 extension crate (`rust/excelbench_rust/`) with three backends:
- **calamine** (read-only, fast) — `calamine_styled_backend.rs` currently handles Tier 0 + Tier 1 read (cell values, formatting, borders, dimensions)
- **rust_xlsxwriter** (write-only) — `rust_xlsxwriter_backend.rs` currently handles Tier 0 + Tier 1 write
- **umya** (read+write, slow) — reference implementation in `src/umya/` with ALL features implemented

The calamine backend uses a **forked calamine** with styles PR #538 merged. The fork is at `git = "https://github.com/SynthGL/calamine.git", branch = "styles"` in Cargo.toml.

## Objective

Implement formulas read, move column-width padding to Rust, wire 4 Tier 2 write features in rust_xlsxwriter, and add 4 Tier 2 read features to calamine_styled_backend by parsing raw OOXML XML from the xlsx zip archive. Skip conditional_formatting, data_validation, and images (Hard tier — separate sprint).

## Files to Modify

### 1. `rust/excelbench_rust/src/calamine_styled_backend.rs`

**A. Store file path in CalamineStyledBook struct**

Add `file_path: String` to the struct so we can re-open the xlsx as a zip for Tier 2 parsing:

```rust
pub struct CalamineStyledBook {
    workbook: XlsxReader,
    sheet_names: Vec<String>,
    style_cache: HashMap<String, SheetCache>,
    file_path: String,  // ADD THIS
}
```

Update `open()` to store path.

**B. Move column width padding to Rust**

In `read_column_width()`, subtract Excel's Calibri 11pt font-metric padding BEFORE returning to Python:

```rust
const CALIBRI_WIDTH_PADDING: f64 = 0.83203125;
const ALT_WIDTH_PADDING: f64 = 0.7109375;
const WIDTH_TOLERANCE: f64 = 0.0005;

// In read_column_width(), after getting raw width:
fn strip_excel_padding(raw: f64) -> f64 {
    let frac = raw % 1.0;
    for padding in [CALIBRI_WIDTH_PADDING, ALT_WIDTH_PADDING] {
        if (frac - padding).abs() < WIDTH_TOLERANCE {
            let adjusted = raw - padding;
            if adjusted >= 0.0 {
                return (adjusted * 10000.0).round() / 10000.0;
            }
        }
    }
    (raw * 10000.0).round() / 10000.0
}
```

**C. Add `read_cell_formula()` method**

calamine has `worksheet_formula(sheet_name)` which returns a `Range<String>` of formula strings. Expose it:

```rust
pub fn read_cell_formula(&mut self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<PyObject> {
    let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let range = self.workbook.worksheet_formula(sheet)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Formula error: {e}")))?;
    match range.get_value((row, col)) {
        Some(f) if !f.is_empty() => {
            let d = PyDict::new(py);
            d.set_item("type", "formula")?;
            d.set_item("formula", format!("={f}"))?;
            Ok(d.into())
        }
        _ => Ok(py.None()),
    }
}
```

**D. Add Tier 2 read methods using zip + quick-xml**

Add `zip` as a dependency in Cargo.toml (it's already a transitive dep of calamine — just add `zip = "2"` directly).

Also add `quick-xml = "0.37"` (calamine already depends on it, just pin what calamine uses).

Create a helper that opens the zip and finds the sheet XML path:

```rust
use std::io::Read as IoRead;
use zip::ZipArchive;

/// Get raw XML for a given sheet from the xlsx zip archive.
fn sheet_xml_content(file_path: &str, sheet_name: &str) -> PyResult<String> {
    // Open zip, read workbook.xml to find sheet rId, resolve via _rels/workbook.xml.rels
    // Then read xl/worksheets/sheet{N}.xml
    // ... (implementation details below)
}
```

**D1. `read_merged_ranges(sheet) -> Vec<String>`**

Parse `<mergeCells><mergeCell ref="A1:C3"/></mergeCells>` from sheet XML.
Return: `["A1:C3", "D5:F10"]`

Reference (umya returns same format): `src/umya/merged_cells.rs:8-20`

**D2. `read_hyperlinks(py, sheet) -> PyObject` (list of dicts)**

Parse `<hyperlinks><hyperlink ref="A1" r:id="rId1" display="Click" tooltip="..."/></hyperlinks>` from sheet XML.
Resolve `r:id` to actual URL via `xl/worksheets/_rels/sheet{N}.xml.rels`.
For internal links (no r:id, has `location` attr), set `internal=true`.

Return format (match umya `src/umya/hyperlinks.rs:11-36`):
```python
[{"cell": "A1", "target": "https://example.com", "display": "Click", "tooltip": "Tip", "internal": False}]
```

**D3. `read_comments(py, sheet) -> PyObject` (list of dicts)**

Find the comments file: in `xl/worksheets/_rels/sheet{N}.xml.rels`, look for relationship with type `...comments`. Then parse `xl/comments{N}.xml`:
```xml
<comments>
  <authors><author>Author Name</author></authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text><r><t>Comment text</t></r></text>
    </comment>
  </commentList>
</comments>
```

Return format (match umya `src/umya/comments.rs:25-44`):
```python
[{"cell": "A1", "text": "Comment text", "author": "Author Name", "threaded": False}]
```

**D4. `read_freeze_panes(py, sheet) -> PyObject` (dict)**

Parse `<sheetViews><sheetView ...><pane xSplit="0" ySplit="1" topLeftCell="A2" state="frozen"/></sheetView></sheetViews>` from sheet XML.

Return format (match umya `src/umya/freeze_panes.rs:46-71`):
```python
{"mode": "freeze", "top_left_cell": "A2", "x_split": 0, "y_split": 1}
```

Only return data if `state="frozen"` or `state="frozenSplit"`.

### 2. `rust/excelbench_rust/src/rust_xlsxwriter_backend.rs`

Wire up 4 Tier 2 write methods. The rust_xlsxwriter Rust API already supports all of these — just need to add queuing + apply-on-save.

**A. `merge_cells(sheet, range_str)`**

Queue merge range, apply during `save()` using `worksheet.merge_range(first_row, first_col, last_row, last_col, "", &Format::new())`.

Parse "A1:C3" into (row1, col1, row2, col2) using the existing `a1_to_row_col` helper.

Add: `merge_ranges: Vec<(String, u32, u16, u32, u16)>` to struct (sheet, r1, c1, r2, c2).

**B. `add_hyperlink(sheet, cell_a1, url, display, tooltip)`**

Queue hyperlink, apply during `save()` using `worksheet.write_url(row, col, url)` or `write_url_with_text()` if display text differs.

For tooltip: `let mut url_obj = Url::new(url); url_obj.set_tip(tooltip); worksheet.write_url(row, col, &url_obj)`.

Accept a Python dict: `{"cell": "A1", "target": "https://...", "display": "Click", "tooltip": "Tip"}`.

Add: `hyperlinks: Vec<(String, u32, u16, String, Option<String>, Option<String>)>` to struct.

**C. `add_comment(sheet, comment_dict)`**

Queue comment, apply during `save()` using `worksheet.insert_note(row, col, &note)`.

```rust
use rust_xlsxwriter::Note;
let mut note = Note::new(text);
if let Some(author) = author {
    note = note.set_author(author);
}
worksheet.insert_note(row, col, &note)?;
```

Accept dict: `{"cell": "A1", "text": "...", "author": "..."}`.

Add: `comments: Vec<(String, u32, u16, String, Option<String>)>` to struct.

**D. `set_freeze_panes(sheet, settings_dict)`**

Queue pane settings, apply during `save()` using `worksheet.set_freeze_panes(row, col)`.

Accept dict: `{"top_left_cell": "A2"}` — parse to (row, col).

Add: `freeze_panes: HashMap<String, (u32, u16)>` to struct.

### 3. `rust/excelbench_rust/Cargo.toml`

Add direct dependencies (both are already transitive deps):
```toml
zip = { version = "2", optional = true, default-features = false, features = ["deflate"] }
quick-xml = { version = "0.37", optional = true }
```

Add to calamine feature: `calamine = ["dep:calamine", "dep:zip", "dep:quick-xml", "dep:chrono"]`

### 4. `src/excelbench/harness/adapters/rust_calamine_styled_adapter.py`

**A. Remove Python-side column width padding** — the Rust side now handles it:
```python
def read_column_width(self, workbook, sheet, column):
    return workbook.read_column_width(sheet, column)  # Rust handles padding
```

Remove `_CALIBRI_WIDTH_PADDING`, `_ALT_WIDTH_PADDING`, `_WIDTH_TOLERANCE` constants.

**B. Add formula read support:**

The adapter should check for formulas and return them. Add to `read_cell_value()`:
```python
def read_cell_value(self, workbook, sheet, cell):
    # Check for formula first
    formula = workbook.read_cell_formula(sheet, cell)
    if formula is not None and isinstance(formula, dict):
        return cell_value_from_payload(formula)
    # Fall back to regular value
    payload = workbook.read_cell_value(sheet, cell)
    if not isinstance(payload, dict):
        return CellValue(type=CellType.STRING, value=str(payload))
    return cell_value_from_payload(payload)
```

**C. Implement Tier 2 read passthroughs** (replace the `return []` / `return {}` stubs):
```python
def read_merged_ranges(self, workbook, sheet):
    return workbook.read_merged_ranges(sheet)

def read_hyperlinks(self, workbook, sheet):
    return workbook.read_hyperlinks(sheet)

def read_comments(self, workbook, sheet):
    return workbook.read_comments(sheet)

def read_freeze_panes(self, workbook, sheet):
    return workbook.read_freeze_panes(sheet)
```

### 5. `src/excelbench/harness/adapters/rust_xlsxwriter_adapter.py`

**Replace the `return` stubs** with passthroughs for the 4 wired features:
```python
def merge_cells(self, workbook, sheet, cell_range):
    workbook.merge_cells(sheet, cell_range)

def add_hyperlink(self, workbook, sheet, link):
    workbook.add_hyperlink(sheet, link)

def add_comment(self, workbook, sheet, comment):
    workbook.add_comment(sheet, comment)

def set_freeze_panes(self, workbook, sheet, settings):
    workbook.set_freeze_panes(sheet, settings)
```

## Testing

After making changes:

```bash
# Build Rust extension
cd rust/excelbench_rust
PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1 maturin develop --features calamine,rust_xlsxwriter,umya

# Run existing tests (should all pass — 1155 tests)
cd ../..
uv run pytest

# Run fidelity benchmark for calamine-styled
uv run excelbench benchmark --tests fixtures/excel --output results_dev_calamine_styled --adapter calamine-styled

# Run fidelity benchmark for rust_xlsxwriter
uv run excelbench benchmark --tests fixtures/excel --output results_dev_rxw --adapter rust_xlsxwriter

# Check matrix.csv scores — target: formulas=3, merged_cells=3, hyperlinks=3, comments=3, freeze_panes=3
cat results_dev_calamine_styled/matrix.csv
cat results_dev_rxw/matrix.csv
```

## Expected Outcomes

**calamine-styled read scores** (from 0 → 3):
- formulas: 0 → 3
- merged_cells: 0 → 3
- hyperlinks: 0 → 3
- comments: 0 → 3
- freeze_panes: 0 → 3

**rust_xlsxwriter write scores** (from 0 → 3):
- merged_cells: 0 → 3
- hyperlinks: 0 → 3
- comments: 0 → 3
- freeze_panes: 0 → 3

**Performance constraint**: Open time must stay under 2ms. The Tier 2 read methods use lazy zip parsing — only executed when called, so they add zero overhead to open/format/border reads.

## Reference Files

Study these existing implementations for patterns and dict formats:
- `rust/excelbench_rust/src/umya/merged_cells.rs` — merged cells read/write reference
- `rust/excelbench_rust/src/umya/hyperlinks.rs` — hyperlinks read/write reference
- `rust/excelbench_rust/src/umya/comments.rs` — comments read/write reference
- `rust/excelbench_rust/src/umya/freeze_panes.rs` — freeze panes read/write reference
- `rust/excelbench_rust/src/umya/mod.rs` — how UmyaBook exposes these methods
- `src/excelbench/harness/adapters/pyumya_adapter.py` — Python adapter reference
- `src/excelbench/harness/runner.py` — how the benchmark harness calls adapter methods

## Important Notes

- The calamine fork (SynthGL/calamine branch: styles) uses calamine 0.33.0-ish with styles PR #538
- PyO3 version is 0.24; needs `PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1` env var for Python 3.14
- All Rust methods that return Python objects need `py: Python<'_>` parameter
- Tier 2 read methods that parse zip should cache the parsed data per-sheet in the style_cache or a separate cache
- The `zip` crate's `ZipArchive` needs `Read + Seek` — use `File` directly, not `BufReader`
- Sheet XML paths follow pattern: read `xl/workbook.xml` for `<sheet name="..." sheetId="..." r:id="rId{N}"/>`, then resolve rId via `xl/_rels/workbook.xml.rels` to get `xl/worksheets/sheet{N}.xml`
- Comments rels are in `xl/worksheets/_rels/sheet{N}.xml.rels` with type ending in `/comments`
- Hyperlinks use `r:id` attribute pointing to `xl/worksheets/_rels/sheet{N}.xml.rels` with type ending in `/hyperlink`
