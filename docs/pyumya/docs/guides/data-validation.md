# Data Validation

Add dropdown lists, input constraints, and validation rules to cells.

## Reading Validations

```python
from excelbench_rust import UmyaBook

book = UmyaBook.open("form.xlsx")
validations = book.read_data_validations("Sheet1")
for v in validations:
    print(f"{v['ranges']}: {v['type']} — {v.get('formula1', '')}")
# ['B2:B100']: list — "Option A,Option B,Option C"
# ['C2:C100']: whole — 1
```

## Writing Validations

### Dropdown List

```python
book = UmyaBook()
book.add_sheet("Form")

# Create a dropdown list
book.add_data_validation("Form", {
    "ranges": ["B2:B100"],
    "type": "list",
    "formula1": "Option A,Option B,Option C",
    "show_dropdown": True,
})

book.save("output.xlsx")
```

### Numeric Constraints

```python
# Whole number between 1 and 100
book.add_data_validation("Form", {
    "ranges": ["C2:C50"],
    "type": "whole",
    "operator": "between",
    "formula1": "1",
    "formula2": "100",
    "error_title": "Invalid Input",
    "error_message": "Enter a number between 1 and 100",
})
```

### Date Range

```python
# Dates in 2026 only
book.add_data_validation("Form", {
    "ranges": ["D2:D50"],
    "type": "date",
    "operator": "between",
    "formula1": "2026-01-01",
    "formula2": "2026-12-31",
})
```

## Validation Types

| Type | Description | Example |
|------|-------------|---------|
| `list` | Dropdown selection | `"Red,Green,Blue"` |
| `whole` | Integer constraint | Between 1 and 100 |
| `decimal` | Float constraint | Greater than 0.0 |
| `date` | Date constraint | After 2026-01-01 |
| `textLength` | String length | Max 50 characters |
| `custom` | Custom formula | `=AND(A1>0, A1<100)` |

## Operators

`between`, `notBetween`, `equal`, `notEqual`, `greaterThan`,
`lessThan`, `greaterThanOrEqual`, `lessThanOrEqual`.
