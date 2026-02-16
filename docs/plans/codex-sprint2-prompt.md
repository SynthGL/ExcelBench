# Codex Sprint 2 Prompt: Fix cell_values Regression + Add Tier 2/3 Read & Write

> Historical naming note: this prompt predates the module rename. References to
> `excelbench_rust` refer to what is now imported as `wolfxl._rust`. The crate path
> `rust/excelbench_rust/` remains current.

## Context

This is a follow-up to Sprint 1 which added formulas read, column-width padding in Rust, 4 Tier 2 read features (merged_cells, hyperlinks, comments, freeze_panes) to calamine_styled_backend, and 4 Tier 2 write features to rust_xlsxwriter_backend.

**Sprint 1 current scores (verified):**

calamine-styled read:
- cell_values: **1** (REGRESSION — was 3 before Sprint 1)
- formulas: 3 ✓
- text_formatting: 3, background_colors: 3, number_formats: 3, alignment: 3
- borders: 1 (known limitation — diagonal borders)
- dimensions: 3, multiple_sheets: 3
- merged_cells: 3 ✓, hyperlinks: 3 ✓, comments: 3 ✓, freeze_panes: 3 ✓
- conditional_formatting: 0, data_validation: 0, images: 0, named_ranges: 0, tables: 0

rust_xlsxwriter write:
- merged_cells: 3 ✓, hyperlinks: 3 ✓, comments: 3 ✓, freeze_panes: 3 ✓
- conditional_formatting: 0, data_validation: 0, images: 0, named_ranges: 0, tables: 0

## Part 1: Fix cell_values Regression (CRITICAL)

### Root Cause

The Sprint 1 Python adapter change in `rust_calamine_styled_adapter.py` intercepts formulas in `read_cell_value()`:

```python
def read_cell_value(self, workbook, sheet, cell):
    formula = workbook.read_cell_formula(sheet, cell)
    if formula is not None and isinstance(formula, dict):
        return cell_value_from_payload(formula)
    ...
```

This breaks 3 cell_values tests because **error cells ARE formulas** in Excel:
- `=1/0` produces `#DIV/0!` → expected `{"type": "error", "value": "#DIV/0!"}`, got `{"type": "formula", "value": "=1/0"}`
- `=NA()` produces `#N/A` → same problem
- `="text"+1` produces `#VALUE!` → same problem

The benchmark harness uses DIFFERENT code paths for cell_values vs formulas:
- `cell_values` test → calls `adapter.read_cell_value()` → expects computed value (type=number/error/string)
- `formulas` test → calls `adapter.read_cell_value()` → then checks if type==FORMULA

The formula interception must NOT override the computed error value for error cells.

### Fix (in `rust/excelbench_rust/src/calamine_styled_backend.rs`)

Modify `read_cell_value()` to check for formulas ONLY when the computed value is NOT an error. When calamine returns `Data::Error`, always return the error value — never check for formulas on error cells:

```rust
pub fn read_cell_value(&mut self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<PyObject> {
    let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    self.ensure_sheet_exists(sheet)?;

    let range = self.workbook.worksheet_range(sheet).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("Failed to read sheet {sheet}: {e}"))
    })?;

    let value = match range.get_value((row, col)) {
        None => return cell_blank(py),
        Some(v) => v,
    };

    // Error cells take priority — never return formula for errors
    if matches!(value, Data::Error(_)) {
        let normalized = map_error_value(&format!("{value:?}"));
        let d = PyDict::new(py);
        d.set_item("type", "error")?;
        d.set_item("value", normalized)?;
        return Ok(d.into());
    }

    // For non-error cells, check for formula
    if let Ok(formula_range) = self.workbook.worksheet_formula(sheet) {
        if let Some(f) = formula_range.get_value((row, col)) {
            if !f.is_empty() {
                let formula = if f.starts_with('=') { f.clone() } else { format!("={f}") };
                let d = PyDict::new(py);
                d.set_item("type", "formula")?;
                d.set_item("formula", &formula)?;
                d.set_item("value", &formula)?;
                Ok(d.into())
            } else {
                // No formula — fall through to regular value
                Self::data_to_py(py, value)
            }
        } else {
            Self::data_to_py(py, value)
        }
    } else {
        // worksheet_formula() failed (no formulas in sheet) — return regular value
        Self::data_to_py(py, value)
    }
}
```

This means you need to extract the existing match arms (String, Float, Int, Bool, DateTime, etc.) into a helper method `fn data_to_py(py, value) -> PyResult<PyObject>`. Keep the Error arm in the main method for the early return.

**Also fix the Python adapter** — remove the formula interception from `read_cell_value()` since the Rust side now handles it:

```python
def read_cell_value(self, workbook, sheet, cell):
    payload = workbook.read_cell_value(sheet, cell)
    if not isinstance(payload, dict):
        return CellValue(type=CellType.STRING, value=str(payload))
    return cell_value_from_payload(payload)
```

The `read_cell_formula()` Rust method should remain available (it's already exposed) but the Python adapter should NOT call it — `read_cell_value()` in the Rust backend now handles formula detection internally.

### Fix (in `rust/excelbench_rust/src/calamine_styled_backend.rs` — read_cell_formula)

Make `read_cell_formula()` gracefully handle sheets without formulas:

```rust
pub fn read_cell_formula(&mut self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<PyObject> {
    let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    self.ensure_sheet_exists(sheet)?;

    // Gracefully handle sheets that have no formula data
    let range = match self.workbook.worksheet_formula(sheet) {
        Ok(r) => r,
        Err(_) => return Ok(py.None()),
    };
    // ... rest unchanged
}
```

### Expected outcome after fix

- cell_values: 1 → **3** (all 18 tests pass)
- formulas: 3 (unchanged)

### Verification

```bash
PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1 uv run maturin develop --manifest-path rust/excelbench_rust/Cargo.toml --features calamine,rust_xlsxwriter,umya
uv run excelbench benchmark -t fixtures/excel -o /tmp/test_cv -a calamine-styled -f cell_values -f formulas
# Expect: cell_values=3, formulas=3
```

## Part 2: Fix Hyperlink Internal Detection (Latent Bug)

In `calamine_styled_backend.rs`, the `compute_hyperlinks()` method has a subtle bug:

```rust
// CURRENT (wrong):
let internal = n.location.is_some();
let target = if let Some(loc) = n.location {
    loc
} else if let Some(rid) = n.rid {
    ...
```

A hyperlink can have BOTH `r:id` (external URL) AND `location` (anchor within the page). The current code treats any hyperlink with a `location` as internal, even if it also has an `r:id` pointing to an external URL.

**Fix:**

```rust
// An internal link has location but NO r:id
let internal = n.location.is_some() && n.rid.is_none();
let target = if let Some(rid) = &n.rid {
    // External link — r:id takes priority
    rid_targets.get(rid).cloned().unwrap_or_default()
} else if let Some(loc) = &n.location {
    // Internal link — location only
    loc.clone()
} else {
    String::new()
};
```

This doesn't affect the current score (hyperlinks=3) but prevents future test failures.

## Part 3: Extract Shared OOXML Helpers (Code Quality)

Both `calamine_styled_backend.rs` and `rust_xlsxwriter_backend.rs` have duplicated zip/XML parsing functions (~200 lines each). Extract these to a shared module.

### Create `rust/excelbench_rust/src/ooxml_util.rs`

Move these functions from BOTH backends into the shared module:

```rust
// From both backends (identical implementations):
pub fn normalize_zip_path(path: &str) -> String { ... }
pub fn join_and_normalize(base_dir: &str, target: &str) -> String { ... }
pub fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> { ... }
pub fn parse_workbook_sheet_rids(xml: &str) -> PyResult<Vec<(String, String)>> { ... }
pub fn parse_relationship_targets(xml: &str) -> PyResult<HashMap<String, String>> { ... }
pub fn zip_read_to_string(zip: &mut ZipArchive<File>, name: &str) -> PyResult<String> { ... }
pub fn zip_read_to_string_opt(zip: &mut ZipArchive<File>, name: &str) -> PyResult<Option<String>> { ... }
```

### Update `rust/excelbench_rust/src/lib.rs`

Add `mod ooxml_util;` (NOT behind a feature flag — both calamine and rust_xlsxwriter features need it).

Actually — since `ooxml_util` depends on `zip` and `quick-xml` which are optional, gate the module:

```rust
#[cfg(any(feature = "calamine", feature = "rust_xlsxwriter"))]
mod ooxml_util;
```

### Update both backends

Replace local helper functions with `use crate::ooxml_util::*;` imports.

## Part 4: Sprint 2 — Tier 2/3 Features

### 4A. Calamine Read: conditional_formatting

Parse `<conditionalFormatting>` elements from sheet XML.

OOXML structure:
```xml
<conditionalFormatting sqref="A1:A10">
  <cfRule type="cellIs" operator="greaterThan" priority="1">
    <formula>100</formula>
    <dxf>
      <fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill>
      <font><color rgb="FF0000FF"/></font>
    </dxf>
  </cfRule>
</conditionalFormatting>
```

Note: The `dxf` (differential formatting) element can appear inline or as an index into `xl/styles.xml`. Start by handling inline dxf elements. If the cfRule has a `dxfId` attribute instead, look up the dxf in `xl/styles.xml` (in the `<dxfs>` section, 0-indexed).

Add `read_conditional_formats(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject>` method.

Return format (match umya `src/umya/conditional_fmt.rs:74-147`):
```python
[{
    "range": "A1:A10",
    "rule_type": "cellIs",        # or "expression", "colorScale", "dataBar", etc.
    "operator": "greaterThan",     # None for expression/colorScale/dataBar
    "formula": "100",              # None if not applicable
    "priority": 1,
    "stop_if_true": False,
    "format": {
        "bg_color": "#FFFF00",     # None if not set
        "font_color": "#0000FF"    # None if not set
    }
}]
```

Implementation approach:
1. Parse sheet XML for `<conditionalFormatting>` elements
2. For each `<cfRule>`, extract type, operator, priority, stopIfTrue
3. For `<formula>` child, extract text content
4. For `<dxf>` child (inline), parse fill bgColor and font color
5. For `dxfId` attribute, load `xl/styles.xml` once (cache it), then index into `<dxfs>` list
6. Cache in tier2_cache

### 4B. Calamine Read: data_validation

Parse `<dataValidations>` from sheet XML.

OOXML structure:
```xml
<dataValidations count="1">
  <dataValidation type="whole" operator="between" sqref="B2:B10"
                  allowBlank="1" showInputMessage="1" showErrorMessage="1"
                  promptTitle="Enter value" prompt="Between 1 and 100"
                  errorTitle="Invalid" error="Must be 1-100">
    <formula1>1</formula1>
    <formula2>100</formula2>
  </dataValidation>
</dataValidations>
```

Add `read_data_validations(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject>` method.

Return format (match umya `src/umya/data_validation.rs:65-140`):
```python
[{
    "range": "B2:B10",
    "validation_type": "whole",    # or "decimal", "list", "date", "time", "textLength", "custom"
    "operator": "between",         # or "notBetween", "equal", "greaterThan", etc.
    "formula1": "1",
    "formula2": "100",             # None for single-formula validations
    "allow_blank": True,
    "show_input": True,
    "show_error": True,
    "prompt_title": "Enter value",
    "prompt": "Between 1 and 100",
    "error_title": "Invalid",
    "error": "Must be 1-100"
}]
```

### 4C. Calamine Read: named_ranges

Parse `<definedNames>` from `xl/workbook.xml` (NOT from individual sheet XML).

OOXML structure:
```xml
<definedNames>
  <definedName name="SalesTotal">Sheet1!$A$1:$A$10</definedName>
  <definedName name="LocalRange" localSheetId="0">Sheet1!$B$2</definedName>
</definedNames>
```

Add `read_named_ranges(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject>` method.

**Important**: The `sheet` parameter is used to filter results. Return ALL workbook-scoped names plus sheet-scoped names where the sheet matches. Normalize `refers_to` by stripping `=`, `$`, and surrounding quotes from sheet names.

Return format (match umya `src/umya/named_ranges.rs:30-80`):
```python
[{
    "name": "SalesTotal",
    "scope": "workbook",           # or "sheet"
    "refers_to": "Sheet1!A1:A10"   # normalized: no =, no $, no quotes around sheet name
}]
```

Note: This reads from workbook.xml which you already parse in `ensure_sheet_xml_paths()`. Extend that parsing to also capture `<definedName>` elements.

### 4D. Calamine Read: tables

Tables are stored as separate XML files referenced from each sheet.

1. In `xl/worksheets/_rels/sheet{N}.xml.rels`, find relationships with type ending in `/table`
2. Each relationship target points to `xl/tables/table{N}.xml`:

```xml
<table id="1" name="Table1" displayName="Table1" ref="A1:D5"
       totalsRowShown="0" headerRowCount="1">
  <autoFilter ref="A1:D5"/>
  <tableColumns count="4">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Age"/>
    <tableColumn id="3" name="City"/>
    <tableColumn id="4" name="Score"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2"/>
</table>
```

Add `read_tables(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject>` method.

Return format (match umya `src/umya/tables.rs:29-100`):
```python
[{
    "name": "Table1",
    "ref": "A1:D5",
    "header_row": True,
    "totals_row": False,
    "style": "TableStyleMedium2",   # None if no tableStyleInfo
    "columns": ["Name", "Age", "City", "Score"],
    "autofilter": True               # True if <autoFilter> element exists
}]
```

### 4E. rust_xlsxwriter Write: conditional_formatting

rust_xlsxwriter v0.79+ supports conditional formatting via `Worksheet::add_conditional_format()`.

Add `add_conditional_format(&mut self, sheet: &str, rule_dict: &Bound<'_, PyAny>) -> PyResult<()>` method.

Queue the rule, apply during `save()`:

```rust
use rust_xlsxwriter::{ConditionalFormatCell, ConditionalFormatFormula,
    ConditionalFormatCellRule, ConditionalFormatType};

// In save():
// Parse the range from the rule dict
// Based on rule_type:
//   "cellIs" → ConditionalFormatCell::new()
//              .set_rule(map_operator(...))
//              .set_value(formula_value)
//              .set_format(&dxf_format)
//   "expression" → ConditionalFormatFormula::new()
//                   .set_rule(formula)
//                   .set_format(&dxf_format)
// ws.add_conditional_format(first_row, first_col, last_row, last_col, &cf)?;
```

**Support the wrapper key pattern**: Accept both `{"cf_rule": {...}}` and direct `{...}` dicts (consistent with umya).

Input dict format:
```python
{
    "range": "A1:A10",
    "rule_type": "cellIs",     # or "expression"
    "operator": "greaterThan", # for cellIs rules
    "formula": "100",          # formula value or expression
    "priority": 1,             # optional
    "stop_if_true": False,     # optional
    "format": {
        "bg_color": "#FFFF00",
        "font_color": "#0000FF"
    }
}
```

### 4F. rust_xlsxwriter Write: data_validation

rust_xlsxwriter supports data validation via `Worksheet::add_data_validation()`.

Add `add_data_validation(&mut self, sheet: &str, validation_dict: &Bound<'_, PyAny>) -> PyResult<()>` method.

```rust
use rust_xlsxwriter::{DataValidation, DataValidationType, DataValidationRule};

// Queue, apply during save():
// let mut dv = DataValidation::new();
// dv.set_input_title(prompt_title).set_input_message(prompt);
// dv.set_error_title(error_title).set_error_message(error);
// Based on validation_type:
//   "whole" → dv.set_type(Integer), set_criteria(operator), set_value(formula1)
//   "decimal" → dv.set_type(Decimal), etc.
//   "list" → dv.set_type(List), set_list_source(formula1)
//   "custom" → dv.set_type(Custom), set_formula(formula1)
// ws.add_data_validation(first_row, first_col, last_row, last_col, &dv)?;
```

**Support the wrapper key pattern**: Accept both `{"validation": {...}}` and direct `{...}` dicts.

Input dict format: same as the read output above.

### 4G. rust_xlsxwriter Write: named_ranges

rust_xlsxwriter supports defined names via `Workbook::define_name()`.

Add `add_named_range(&mut self, sheet: &str, nr_dict: &Bound<'_, PyAny>) -> PyResult<()>` method.

Queue the name definition, apply during `save()` on the Workbook (not Worksheet):

```rust
// In save(), before pushing worksheets:
// For each named range:
//   wb.define_name(name, refers_to)?;
// Note: rust_xlsxwriter define_name takes the full reference (e.g., "Sheet1!$A$1:$A$10")
// For sheet-scoped names, the refers_to should include the sheet prefix
```

Input dict format:
```python
{"name": "SalesTotal", "scope": "workbook", "refers_to": "Sheet1!A1:A10"}
```

### 4H. rust_xlsxwriter Write: tables

rust_xlsxwriter supports tables via `Worksheet::add_table()`.

Add `add_table(&mut self, sheet: &str, table_dict: &Bound<'_, PyAny>) -> PyResult<()>` method.

```rust
use rust_xlsxwriter::Table;

// Queue, apply during save():
// let mut table = Table::new();
// table.set_name(name);
// if let Some(style) = style { table.set_style_name(style); }
// table.set_header_row(header_row);
// table.set_total_row(totals_row);
// if let Some(cols) = columns {
//     for (i, col_name) in cols.iter().enumerate() {
//         let col = TableColumn::new().set_header(col_name);
//         table.add_column(&col);
//     }
// }
// if autofilter { table.set_autofilter(true); }
// ws.add_table(first_row, first_col, last_row, last_col, &table)?;
```

**Support the wrapper key pattern**: Accept both `{"table": {...}}` and direct `{...}` dicts.

Input dict format: same as the read output above.

## Part 5: Python Adapter Wiring

### `src/excelbench/harness/adapters/rust_calamine_styled_adapter.py`

Replace the remaining stubs:

```python
def read_conditional_formats(self, workbook, sheet):
    return workbook.read_conditional_formats(sheet)

def read_data_validations(self, workbook, sheet):
    return workbook.read_data_validations(sheet)

def read_named_ranges(self, workbook, sheet):
    return workbook.read_named_ranges(sheet)

def read_tables(self, workbook, sheet):
    return workbook.read_tables(sheet)
```

### `src/excelbench/harness/adapters/rust_xlsxwriter_adapter.py`

Replace the remaining stubs:

```python
def add_conditional_format(self, workbook, sheet, rule):
    workbook.add_conditional_format(sheet, rule)

def add_data_validation(self, workbook, sheet, validation):
    workbook.add_data_validation(sheet, validation)

def add_named_range(self, workbook, sheet, named_range):
    workbook.add_named_range(sheet, named_range)

def add_table(self, workbook, sheet, table):
    workbook.add_table(sheet, table)
```

## Files to Modify

1. `rust/excelbench_rust/src/calamine_styled_backend.rs` — fix cell_values regression, fix hyperlinks internal, add 4 new read methods (CF, DV, named_ranges, tables)
2. `rust/excelbench_rust/src/rust_xlsxwriter_backend.rs` — add 4 new write methods (CF, DV, named_ranges, tables)
3. `rust/excelbench_rust/src/ooxml_util.rs` — NEW: shared zip/XML helpers extracted from both backends
4. `rust/excelbench_rust/src/lib.rs` — add `mod ooxml_util;`
5. `src/excelbench/harness/adapters/rust_calamine_styled_adapter.py` — revert formula interception in read_cell_value, wire 4 new read passthroughs
6. `src/excelbench/harness/adapters/rust_xlsxwriter_adapter.py` — wire 4 new write passthroughs

## Testing

```bash
# Build
PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1 uv run maturin develop --manifest-path rust/excelbench_rust/Cargo.toml --features calamine,rust_xlsxwriter,umya

# Unit tests (should all pass — 1155 tests)
uv run pytest

# Verify cell_values fix
uv run excelbench benchmark -t fixtures/excel -o /tmp/test_fix -a calamine-styled -f cell_values -f formulas
# Expect: cell_values=3, formulas=3

# Full calamine-styled benchmark
uv run excelbench benchmark -t fixtures/excel -o results_dev_calamine_styled -a calamine-styled

# Full rust_xlsxwriter benchmark
uv run excelbench benchmark -t fixtures/excel -o results_dev_rxw -a rust_xlsxwriter

# Check scores
cat results_dev_calamine_styled/matrix.csv
cat results_dev_rxw/matrix.csv
```

## Expected Outcomes

**calamine-styled read** (target scores):
- cell_values: 1 → **3** (fix regression)
- formulas: 3 (unchanged)
- conditional_formatting: 0 → **3**
- data_validation: 0 → **3**
- named_ranges: 0 → **3**
- tables: 0 → **3**

**rust_xlsxwriter write** (target scores):
- conditional_formatting: 0 → **3**
- data_validation: 0 → **3**
- named_ranges: 0 → **3**
- tables: 0 → **3**

## Reference Files

Study these for patterns and dict formats:
- `rust/excelbench_rust/src/umya/conditional_fmt.rs` — CF read/write reference (full implementation)
- `rust/excelbench_rust/src/umya/data_validation.rs` — DV read/write reference (full implementation)
- `rust/excelbench_rust/src/umya/named_ranges.rs` — named ranges read/write reference
- `rust/excelbench_rust/src/umya/tables.rs` — tables read/write reference
- `src/excelbench/harness/adapters/openpyxl_adapter.py` — Python reference adapter (full fidelity)
- `src/excelbench/harness/adapters/pyumya_adapter.py` — Python adapter reference for umya
- `src/excelbench/harness/runner.py` — how the benchmark harness calls each adapter method, dict projection, formula normalization
- `rust/excelbench_rust/src/calamine_styled_backend.rs` — existing Sprint 1 implementation to build on
- `rust/excelbench_rust/src/rust_xlsxwriter_backend.rs` — existing Sprint 1 implementation to build on

## Important Notes

- **CRITICAL**: Fix cell_values FIRST, then verify the fix before proceeding to Sprint 2 features
- The calamine fork is at `git = "https://github.com/wolfiesch/calamine.git", branch = "styles"` in Cargo.toml
- PyO3 version is 0.24; needs `PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1` env var
- All Rust methods returning Python objects need `py: Python<'_>` parameter
- Use the existing `Tier2SheetCache` for caching new read results per-sheet
- For conditional formatting dxf styles: if `dxfId` attribute is present (index into xl/styles.xml), you need to parse the `<dxfs>` section once and cache it — this is a workbook-level cache, not per-sheet
- For named_ranges: parse from `xl/workbook.xml` (already loaded by `ensure_sheet_xml_paths()`)
- For tables: find table rels in sheet .rels file, then parse each `xl/tables/table{N}.xml`
- The runner's `_project_rule()` function only compares keys present in the expected dict, so returning extra keys is fine
- The runner normalizes formulas (strips `=` prefix, handles sheet quoting) — return raw values
- rust_xlsxwriter conditional format API: `ConditionalFormatCell`, `ConditionalFormatFormula`, etc. — check docs.rs/rust_xlsxwriter for exact struct names
- rust_xlsxwriter DataValidation API: `DataValidation::new()` with `.set_minimum()`, `.set_maximum()`, `.set_input_title()`, `.set_error_title()`, etc.
- rust_xlsxwriter Table API: `Table::new()` with `.set_name()`, `.set_style_name()`, `.set_header_row()`, `.set_total_row()`, `.add_column()`
- **Do NOT skip images** for now — that requires binary blob handling and is out of scope
- **Performance constraint**: Open time must stay under 2ms. All new read methods use lazy parsing.
- Accept the wrapper key patterns used by runner.py: `{"cf_rule": {...}}`, `{"validation": {...}}`, `{"table": {...}}`, `{"named_range": {...}}` — unwrap if present, otherwise use the dict directly
