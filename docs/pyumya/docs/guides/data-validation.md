# Data Validation

Add dropdown lists, input constraints, and validation rules to cells.

## Reading Validations

```python
from wolfxl._rust import UmyaBook

book = UmyaBook.open("form.xlsx")
validations = book.read_data_validations("Sheet1")
for v in validations:
    print(f"{v['range']}: {v['validation_type']} — {v.get('formula1')}")
# B2:B100: list — Option A,Option B,Option C
# C2:C100: whole — 1
```

## Writing Validations

```python
from wolfxl._rust import UmyaBook

book = UmyaBook()
book.add_sheet("Form")

book.add_data_validation("Form", {
    "range": "B2:B100",
    "validation_type": "list",
    "formula1": "Option A,Option B,Option C",
    "allow_blank": True,
})

# Whole number between 1 and 100
book.add_data_validation("Form", {
    "range": "C2:C50",
    "validation_type": "whole",
    "operator": "between",
    "formula1": "1",
    "formula2": "100",
    "error_title": "Invalid Input",
    "error": "Enter a number between 1 and 100",
    "show_error": True,
})

# Dates in 2026 only
book.add_data_validation("Form", {
    "range": "D2:D50",
    "validation_type": "date",
    "operator": "between",
    "formula1": "2026-01-01",
    "formula2": "2026-12-31",
})

book.save("output.xlsx")
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
