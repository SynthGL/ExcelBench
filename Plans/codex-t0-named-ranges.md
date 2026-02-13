# Codex Handoff: Implement "Named Ranges" Feature (Tier 3)

> **Status: COMPLETED** — This feature was implemented via Codex handoff. This document is retained as a historical reference for the design decisions and implementation pattern.

## Context

ExcelBench is a benchmark suite that scores Python Excel libraries on feature fidelity. This document described adding the first Tier 3 feature: **named ranges**.

The codebase follows a strict pattern for adding features. Every existing feature was implemented using the same 5-step process. Follow the patterns exactly.

## Objective

Add full "named ranges" support to ExcelBench: generator, runner integration, adapter base methods, openpyxl adapter implementation, and tests.

## Step-by-Step Instructions

### Step 1: Create the Generator

**File**: `src/excelbench/generator/features/named_ranges.py`

Follow the exact pattern of `src/excelbench/generator/features/hyperlinks.py`. Key structure:

```python
from excelbench.generator.base import FeatureGenerator
from excelbench.models import Importance, TestCase

class NamedRangesGenerator(FeatureGenerator):
    feature_name = "named_ranges"
    tier = 3
    filename = "18_named_ranges.xlsx"

    def generate(self, sheet) -> list[TestCase]:
        ...

    def post_process(self, output_path) -> None:
        # Use openpyxl to write named ranges (xlwings named range API is fragile on macOS)
        ...
```

Test cases to generate (each should set values in cells, then define a named range pointing to them):

| ID | Label | Named Range | Scope | Refers To | Expected |
|----|-------|-------------|-------|-----------|----------|
| `nr_simple_cell` | Named range: single cell | `SingleCell` | Workbook | `named_ranges!$B$2` | `{"name": "SingleCell", "scope": "workbook", "refers_to": "named_ranges!$B$2", "value": 42}` |
| `nr_cell_range` | Named range: cell range | `DataRange` | Workbook | `named_ranges!$B$3:$D$3` | `{"name": "DataRange", "scope": "workbook", "refers_to": "named_ranges!$B$3:$D$3"}` |
| `nr_formula_ref` | Named range: used in formula | `TaxRate` | Workbook | `named_ranges!$B$4` | `{"name": "TaxRate", "scope": "workbook", "refers_to": "named_ranges!$B$4", "value": 0.08}` |
| `nr_sheet_scope` | Named range: sheet-scoped | `LocalName` | Sheet | `named_ranges!$B$5` | `{"name": "LocalName", "scope": "sheet", "refers_to": "named_ranges!$B$5", "value": "local"}` |
| `nr_cross_sheet` | Named range: cross-sheet reference | `OtherSheet` | Workbook | `Targets!$A$1` | `{"name": "OtherSheet", "scope": "workbook", "refers_to": "Targets!$A$1"}` |
| `nr_special_chars` | Named range: underscore name | `_my_range` | Workbook | `named_ranges!$B$7` | `{"name": "_my_range", "scope": "workbook", "refers_to": "named_ranges!$B$7"}` |

In `post_process`, use openpyxl to define the named ranges:
```python
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName

wb = load_workbook(output_path)
# Workbook-scoped: wb.defined_names.add(DefinedName(name, attr_text=f"={refers_to}"))
# Sheet-scoped: dn = DefinedName(name, attr_text=f"={refers_to}"); ws = wb[sheet]; ws.defined_names.append(dn)
wb.save(output_path)
```

Mark `nr_sheet_scope`, `nr_cross_sheet`, and `nr_special_chars` with `importance=Importance.EDGE`.

### Step 2: Register the Generator

**File**: `src/excelbench/generator/features/__init__.py`

Add import and export:
```python
from excelbench.generator.features.named_ranges import NamedRangesGenerator
# Add "NamedRangesGenerator" to __all__
```

**File**: `src/excelbench/generator/generate.py`

Find where generators are instantiated (look for the list containing `HyperlinksGenerator()`, `FreezePanesGenerator()`, etc.) and add `NamedRangesGenerator()`.

### Step 3: Add Adapter Base Methods

**File**: `src/excelbench/harness/adapters/base.py`

Add these methods (with default no-op implementations) to the `ExcelAdapter` class. Follow the pattern of existing Tier 2 methods like `read_hyperlinks` / `add_hyperlink`:

```python
# Read
def read_named_ranges(self, workbook: Any, sheet: str) -> list[JSONDict]:
    """Read all named ranges from the workbook. Return list of dicts with keys: name, scope, refers_to."""
    return []

# Write
def add_named_range(self, workbook: Any, sheet: str, named_range: JSONDict) -> None:
    """Add a named range. named_range dict has keys: name, scope, refers_to."""
    return
```

These should NOT be abstract — give them default no-op implementations (return empty list / return None) so existing adapters don't break. This matches how Tier 2 methods were added (check `add_hyperlink`, `add_comment`, etc. for the pattern).

### Step 4: Implement in openpyxl Adapter

**File**: `src/excelbench/harness/adapters/openpyxl_adapter.py`

Override both methods:

```python
def read_named_ranges(self, workbook, sheet):
    results = []
    # Workbook-scoped names
    for dn in workbook.defined_names.definedName:
        results.append({
            "name": dn.name,
            "scope": "sheet" if dn.localSheetId is not None else "workbook",
            "refers_to": dn.attr_text.lstrip("="),
        })
    return results

def add_named_range(self, workbook, sheet, named_range):
    from openpyxl.workbook.defined_name import DefinedName
    name = named_range["name"]
    refers_to = named_range["refers_to"]
    scope = named_range.get("scope", "workbook")
    if scope == "sheet":
        dn = DefinedName(name, attr_text=f"={refers_to}")
        ws = workbook[sheet]
        ws.defined_names.append(dn)
    else:
        workbook.defined_names.add(DefinedName(name, attr_text=f"={refers_to}"))
```

### Step 5: Add Runner Integration

**File**: `src/excelbench/harness/runner.py`

Add read dispatch (in `test_read_case`, find the `elif feature ==` chain):

```python
elif feature == "named_ranges":
    actual = read_named_ranges_actual(adapter, workbook, sheet, expected)
```

Add write dispatch (in `test_write`, find the `elif test_file.feature ==` chain):

```python
elif test_file.feature == "named_ranges":
    _write_named_range_case(adapter, workbook, target_sheet, tc.expected)
```

Implement the helper functions (add near the other `read_*_actual` / `_write_*_case` functions):

```python
def read_named_ranges_actual(
    adapter: ExcelAdapter, workbook: Any, sheet: str, expected: JSONDict
) -> JSONDict:
    """Read named ranges and find the one matching expected['name']."""
    all_names = adapter.read_named_ranges(workbook, sheet)
    target_name = expected.get("name", "")
    for nr in all_names:
        if nr.get("name", "").lower() == target_name.lower():
            result = {
                "name": nr["name"],
                "scope": nr.get("scope", "workbook"),
                "refers_to": nr.get("refers_to", ""),
            }
            # If expected has a value key, read the cell value at the referred location
            if "value" in expected:
                ref_str = nr.get("refers_to", "")
                cell_ref = _parse_named_range_single_cell(ref_str, sheet)
                if cell_ref is not None:
                    ref_sheet, ref_cell = cell_ref
                    cv = adapter.read_cell_value(workbook, ref_sheet, ref_cell)
                    result["value"] = cv.value
            return result
    return {"name": target_name, "scope": "not_found", "refers_to": ""}


def _write_named_range_case(
    adapter: ExcelAdapter, workbook: Any, sheet: str, expected: JSONDict
) -> None:
    adapter.add_named_range(workbook, sheet, expected)
```

### Step 6: Add Tests

**File**: `tests/test_named_ranges.py`

Follow the pattern of `tests/test_runner_feature_reads.py`. Create a minimal test:

```python
"""Tests for named ranges feature (Tier 3)."""
import pytest
from excelbench.harness.adapters import OpenpyxlAdapter

FIXTURE = "fixtures/excel/tier3/18_named_ranges.xlsx"

@pytest.fixture
def adapter():
    return OpenpyxlAdapter()

class TestNamedRangesRead:
    @pytest.mark.skipif(
        not __import__("pathlib").Path(FIXTURE).exists(),
        reason="Named ranges fixture not generated yet",
    )
    def test_read_named_ranges_returns_list(self, adapter):
        wb = adapter.open_workbook(__import__("pathlib").Path(FIXTURE))
        try:
            names = adapter.read_named_ranges(wb, "named_ranges")
            assert isinstance(names, list)
            assert len(names) >= 1
        finally:
            adapter.close_workbook(wb)

class TestNamedRangesBase:
    """Test that base adapter methods exist and have sensible defaults."""

    def test_read_named_ranges_default_returns_empty(self):
        from excelbench.harness.adapters.base import ExcelAdapter
        # Verify the method exists on the base class
        assert hasattr(ExcelAdapter, "read_named_ranges")

    def test_add_named_range_default_is_noop(self):
        from excelbench.harness.adapters.base import ExcelAdapter
        assert hasattr(ExcelAdapter, "add_named_range")
```

## Acceptance Criteria

1. `uv run ruff check` passes with no new errors
2. `uv run pytest` passes (existing 1116 tests + new tests, 0 failures)
3. `NamedRangesGenerator` is registered and would produce test cases when `excelbench generate` runs
4. The `read_named_ranges` / `add_named_range` methods exist on `ExcelAdapter` base class with no-op defaults
5. `OpenpyxlAdapter` overrides both methods with real implementations
6. Runner dispatches to named_ranges for both read and write paths
7. No changes to any existing feature code — this is purely additive

## Files to Modify (summary)

| File | Action |
|------|--------|
| `src/excelbench/generator/features/named_ranges.py` | **Create** |
| `src/excelbench/generator/features/__init__.py` | Add import + `__all__` entry |
| `src/excelbench/generator/generate.py` | Add to generator list |
| `src/excelbench/harness/adapters/base.py` | Add 2 default methods |
| `src/excelbench/harness/adapters/openpyxl_adapter.py` | Override 2 methods |
| `src/excelbench/harness/runner.py` | Add read/write dispatch + 2 helper functions |
| `tests/test_named_ranges.py` | **Create** |

## Do NOT

- Modify any existing feature generators or test files
- Change scoring logic
- Touch Rust code
- Regenerate fixtures (the generator code is sufficient; fixtures require Excel)
- Add dependencies to pyproject.toml
