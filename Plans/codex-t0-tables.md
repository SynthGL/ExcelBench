# Codex Handoff: Implement "Tables (Structured References)" Feature (Tier 3)

> **Status: COMPLETED** — This feature was implemented via Codex handoff. This document is retained as a historical reference for the design decisions and implementation pattern.

## Context

ExcelBench is a benchmark suite that scores Python Excel libraries on feature fidelity. This document described adding the second Tier 3 feature: **tables** (Excel ListObjects with structured references).

The codebase follows a strict pattern for adding features. The recently-added `named_ranges` feature is the most up-to-date reference. Follow its patterns exactly.

## Important Lessons from Previous Handoff

1. **Only modify files listed in the "Files to Modify" section.** Do not touch any other file.
2. **Do not change existing logic** in `generate.py` beyond adding the import + registration.
3. **openpyxl DefinedName `attr_text` must start with `=`** — always prefix references.
4. **When no match is found in a `read_*_actual` function, return a dict with the expected keys set to sentinel values** (e.g., `"name": target_name, "ref": "not_found"`), not `{}`.

## What Are Excel Tables?

Excel Tables (ListObject in the OOXML spec) are structured ranges with:
- A name (e.g., `Table1`)
- A cell range (e.g., `$A$1:$D$10`)
- Headers (first row of the range are column names)
- Optional total row
- Optional auto-filter
- A display style (e.g., `TableStyleMedium9`)

In openpyxl, they are `openpyxl.worksheet.table.Table` objects on `worksheet.tables`.

## Step-by-Step Instructions

### Step 1: Create the Generator

**File**: `src/excelbench/generator/features/tables.py` (new)

Follow the exact pattern of `src/excelbench/generator/features/named_ranges.py`:

```python
"""Generator for table (ListObject/structured reference) test cases (Tier 3)."""

from pathlib import Path
import xlwings as xw
from excelbench.generator.base import FeatureGenerator
from excelbench.models import Importance, TestCase


class TablesGenerator(FeatureGenerator):
    feature_name = "tables"
    tier = 3
    filename = "19_tables.xlsx"

    def __init__(self) -> None:
        self._ops: list[dict[str, object]] = []

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)
        test_cases: list[TestCase] = []
        row = 2
        # ... (see test cases below)
        return test_cases

    def post_process(self, output_path: Path) -> None:
        # Use openpyxl to create the actual Table objects
        ...
```

**Test cases to generate:**

For each test case, write data into cells first (via xlwings in `generate`), then create the Table object in `post_process` using openpyxl.

| ID | Label | Table Name | Range | Headers | Style | Importance |
|----|-------|-----------|-------|---------|-------|------------|
| `tbl_basic` | Table: basic 3-col | `SalesData` | `E2:G5` | Name, Qty, Price | `TableStyleMedium9` | basic |
| `tbl_with_totals` | Table: with totals row | `Summary` | `E7:G11` | Item, Count, Total | `TableStyleLight1` | basic |
| `tbl_no_style` | Table: no style | `PlainTable` | `E13:F16` | Key, Value | `None` | basic |
| `tbl_single_col` | Table: single column | `SingleCol` | `E18:E21` | Score | `TableStyleMedium2` | edge |
| `tbl_single_row` | Table: header only (no data rows) | `EmptyTable` | `E23:G23` | A, B, C | `TableStyleMedium9` | edge |
| `tbl_autofilter` | Table: with autoFilter | `Filtered` | `E25:G28` | Region, Sales, Year | `TableStyleMedium9` | edge |

**Generator `generate` method** — for each test case, write the cell data into the sheet. Example for `tbl_basic`:

```python
# Write data for SalesData table (columns E-G to avoid A:C metadata columns)
sheet.range("E2").value = "Name"
sheet.range("F2").value = "Qty"
sheet.range("G2").value = "Price"
sheet.range("E3").value = "Widget"
sheet.range("F3").value = 10
sheet.range("G3").value = 4.99
sheet.range("E4").value = "Gadget"
sheet.range("F4").value = 5
sheet.range("G4").value = 12.50
sheet.range("E5").value = "Gizmo"
sheet.range("F5").value = 8
sheet.range("G5").value = 7.25
```

Store each table definition in `self._ops` as:
```python
self._ops.append({
    "name": "SalesData",
    "ref": "E2:G5",
    "style": "TableStyleMedium9",
    "totals_row": False,
})
```

Expected dict for each test case:
```python
expected = {
    "table": {
        "name": "SalesData",
        "ref": "E2:G5",
        "header_row": True,
        "totals_row": False,
        "style": "TableStyleMedium9",
        "columns": ["Name", "Qty", "Price"],
    }
}
```

For `tbl_with_totals`, set `totals_row: True` and include a `"totals_row_count"` key.
For `tbl_no_style`, set `style: None`.
For `tbl_autofilter`, include `"autofilter": True` in the expected dict.

**Generator `post_process` method** — use openpyxl to create the tables:

```python
def post_process(self, output_path: Path) -> None:
    if not self._ops:
        return

    from openpyxl import load_workbook
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = load_workbook(output_path)
    try:
        ws = wb[self.feature_name]
        for op in self._ops:
            name = str(op["name"])
            ref = str(op["ref"])
            style_name = op.get("style")
            totals = bool(op.get("totals_row", False))

            style = None
            if style_name:
                style = TableStyleInfo(
                    name=str(style_name),
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False,
                )

            table = Table(displayName=name, ref=ref)
            if style:
                table.tableStyleInfo = style
            if totals:
                table.totalsRowCount = 1
            ws.add_table(table)

        wb.save(output_path)
    finally:
        wb.close()
```

### Step 2: Register the Generator

**File**: `src/excelbench/generator/features/__init__.py`

Add import and `__all__` entry for `TablesGenerator`, alphabetically sorted among existing entries.

**File**: `src/excelbench/generator/generate.py`

Add `TablesGenerator()` to the list returned by `get_all_generators()`. Add the import alongside existing imports. Do NOT change any other logic in this file.

### Step 3: Add Adapter Base Methods

**File**: `src/excelbench/harness/adapters/base.py`

Add these methods to the `Tier 3 Operations` section (after `add_named_range`):

```python
def read_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
    """Read table (ListObject) definitions from a sheet.

    Returns a list of dicts with keys:
    - name: table display name
    - ref: cell range (e.g. "A1:D10")
    - header_row: bool
    - totals_row: bool
    - style: style name or None
    - columns: list of column header strings
    """
    return []

def add_table(self, workbook: Any, sheet: str, table: JSONDict) -> None:
    """Add a table (ListObject) to a sheet.

    table dict should include keys: name, ref, style, columns, header_row, totals_row.
    """
    return None
```

These must NOT be abstract — give them default no-op implementations so existing adapters don't break.

### Step 4: Implement in openpyxl Adapter

**File**: `src/excelbench/harness/adapters/openpyxl_adapter.py`

Add `read_tables` near the existing `read_named_ranges` method:

```python
def read_tables(self, workbook: Workbook, sheet: str) -> list[JSONDict]:
    ws = workbook[sheet]
    out: list[JSONDict] = []
    for tbl in ws.tables.values():
        cols: list[str] = []
        for col in tbl.tableColumns:
            cols.append(str(col.name))
        out.append({
            "name": tbl.displayName,
            "ref": tbl.ref,
            "header_row": tbl.headerRowCount != 0,
            "totals_row": (tbl.totalsRowCount or 0) > 0,
            "style": tbl.tableStyleInfo.name if tbl.tableStyleInfo else None,
            "columns": cols,
            "autofilter": tbl.autoFilter is not None,
        })
    return out
```

Add `add_table` near the existing `add_named_range` method:

```python
def add_table(self, workbook: Workbook, sheet: str, table: JSONDict) -> None:
    from openpyxl.worksheet.table import Table, TableStyleInfo

    data = table.get("table", table)
    name = data.get("name")
    ref = data.get("ref")
    if not name or not ref:
        return

    style_name = data.get("style")
    style = None
    if style_name:
        style = TableStyleInfo(
            name=str(style_name),
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )

    tbl = Table(displayName=str(name), ref=str(ref))
    if style:
        tbl.tableStyleInfo = style
    if data.get("totals_row"):
        tbl.totalsRowCount = 1

    ws = workbook[sheet]
    ws.add_table(tbl)
```

### Step 5: Add Runner Integration

**File**: `src/excelbench/harness/runner.py`

**Read dispatch** — in `test_read_case`, add after the `named_ranges` elif:

```python
elif feature == "tables":
    actual = read_tables_actual(adapter, workbook, sheet, expected)
```

**Write dispatch** — in `test_write`, add after the `named_ranges` elif:

```python
elif test_file.feature == "tables":
    _write_table_case(adapter, workbook, target_sheet, tc.expected)
```

**Helper functions** — add near the other `read_*_actual` / `_write_*_case` functions:

```python
def read_tables_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: JSONDict,
) -> JSONDict:
    """Read tables and return the one matching the expected name."""
    expected_table = expected.get("table", expected)
    target_name = str(expected_table.get("name") or "")
    if not target_name:
        return {"table": {"name": "", "ref": "not_found", "header_row": True,
                          "totals_row": False, "style": None, "columns": []}}

    all_tables = adapter.read_tables(workbook, sheet)

    for tbl in all_tables:
        if str(tbl.get("name", "")).lower() != target_name.lower():
            continue

        result: JSONDict = {
            "table": {
                "name": tbl.get("name", target_name),
                "ref": tbl.get("ref", ""),
                "header_row": tbl.get("header_row", True),
                "totals_row": tbl.get("totals_row", False),
                "style": tbl.get("style"),
                "columns": tbl.get("columns", []),
            }
        }
        # Only include optional keys if expected includes them
        if "autofilter" in expected_table:
            result["table"]["autofilter"] = tbl.get("autofilter", False)
        if "totals_row_count" in expected_table:
            result["table"]["totals_row_count"] = 1 if tbl.get("totals_row") else 0
        return result

    # No match found — return diagnostic sentinel
    return {
        "table": {
            "name": target_name,
            "ref": "not_found",
            "header_row": True,
            "totals_row": False,
            "style": None,
            "columns": [],
        }
    }


def _write_table_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    expected: JSONDict,
) -> None:
    table_data = expected.get("table", expected)

    # Write cell data for columns + any data rows if present in expected
    columns = table_data.get("columns", [])
    ref = table_data.get("ref", "")
    if columns and ref:
        from openpyxl.utils import range_boundaries
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        for ci, col_name in enumerate(columns):
            cell = _coord_to_cell(min_row, min_col + ci)
            adapter.write_cell_value(
                workbook, sheet, cell,
                _cell_value_from_raw(col_name),
            )

    adapter.add_table(workbook, sheet, expected)
```

Note: `_coord_to_cell` already exists in `runner.py` — reuse it directly.

### Step 6: Register in Renderer

**File**: `src/excelbench/results/renderer.py`

Add to `_FEATURE_TIERS` dict (after `named_ranges`):

```python
"tables": (3, "Workbook Metadata"),
```

### Step 7: Add Tests

**File**: `tests/test_tables.py` (new)

Follow the pattern of `tests/test_named_ranges.py`. Create a test with:

1. A `_StubAdapter` class that exercises base-class defaults (copy the pattern from `test_named_ranges.py` and add `read_tables` / `add_table` stub)
2. `TestTablesBase` — verify `read_tables` returns `[]` and `add_table` is a no-op
3. `TestOpenpyxlTables` — create a workbook, add a table via `add_table`, save, reopen, verify with `read_tables`
4. Fixture-based test with `@pytest.mark.skipif(not FIXTURE.exists())`

```python
FIXTURE = Path("fixtures/excel/tier3/19_tables.xlsx")

class TestOpenpyxlTables:
    def test_roundtrip_table_in_memory(self, tmp_path: Path) -> None:
        adapter = OpenpyxlAdapter()
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "tables")
        # Write header + data cells
        adapter.write_cell_value(wb, "tables", "A1", CellValue(type=CellType.STRING, value="Name"))
        adapter.write_cell_value(wb, "tables", "B1", CellValue(type=CellType.STRING, value="Qty"))
        adapter.write_cell_value(wb, "tables", "A2", CellValue(type=CellType.STRING, value="X"))
        adapter.write_cell_value(wb, "tables", "B2", CellValue(type=CellType.NUMBER, value=10))

        adapter.add_table(
            wb, "tables",
            {"table": {"name": "TestTable", "ref": "A1:B2", "style": "TableStyleMedium9"}},
        )

        path = tmp_path / "tables.xlsx"
        adapter.save_workbook(wb, path)

        wb2 = adapter.open_workbook(path)
        try:
            tables = adapter.read_tables(wb2, "tables")
            assert isinstance(tables, list)
            assert len(tables) == 1
            assert tables[0]["name"] == "TestTable"
            assert tables[0]["ref"] == "A1:B2"
            assert tables[0]["columns"] == ["Name", "Qty"]
        finally:
            adapter.close_workbook(wb2)
```

## Files to Modify (summary)

| File | Action |
|------|--------|
| `src/excelbench/generator/features/tables.py` | **Create** |
| `src/excelbench/generator/features/__init__.py` | Add import + `__all__` entry |
| `src/excelbench/generator/generate.py` | Add import + `TablesGenerator()` to list (NO other changes) |
| `src/excelbench/harness/adapters/base.py` | Add 2 default methods in Tier 3 section |
| `src/excelbench/harness/adapters/openpyxl_adapter.py` | Override 2 methods |
| `src/excelbench/harness/runner.py` | Add read/write dispatch + 3 helper functions (reuse existing `_coord_to_cell`) |
| `src/excelbench/results/renderer.py` | Add `"tables"` to `_FEATURE_TIERS` |
| `tests/test_tables.py` | **Create** |

## Acceptance Criteria

1. `uv run ruff check` passes with no new errors
2. `uv run pytest` passes (all existing tests + new tests, 0 failures)
3. `TablesGenerator` is registered and produces 6 test cases
4. `read_tables` / `add_table` exist on `ExcelAdapter` base with no-op defaults
5. `OpenpyxlAdapter` overrides both with real implementations
6. Runner dispatches `tables` for both read and write paths
7. `_FEATURE_TIERS` in renderer includes `"tables": (3, "Workbook Metadata")`
8. No changes to any file NOT listed in "Files to Modify"

## Do NOT

- Modify any existing feature generators, adapters, or tests
- Change scoring logic
- Touch Rust code
- Regenerate fixtures
- Add dependencies to pyproject.toml
- Modify `generate.py` beyond the import + list addition
- Return `{}` from any `read_*_actual` function — always return a dict with expected keys
