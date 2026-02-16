# Conditional Formatting

Apply formatting rules that change cell appearance based on values.

## Reading Conditional Formats

```python
from wolfxl._rust import UmyaBook

book = UmyaBook.open("dashboard.xlsx")
rules = book.read_conditional_formats("Sheet1")
for r in rules:
    print(f"{r['range']}: {r['rule_type']}")
# A1:A100: cellIs
# B1:B100: colorScale
```

## Writing Conditional Formats

```python
from wolfxl._rust import UmyaBook

book = UmyaBook()
book.add_sheet("Sales")

# Highlight cells greater than 1000
book.add_conditional_format("Sales", {
    "range": "B2:B50",
    "rule_type": "cellIs",
    "operator": "greaterThan",
    "formula": "1000",
    "format": {"bg_color": "#C6EFCE", "font_color": "#006100"},  # green
})

# Highlight cells below target
book.add_conditional_format("Sales", {
    "range": "B2:B50",
    "rule_type": "cellIs",
    "operator": "lessThan",
    "formula": "500",
    "format": {"bg_color": "#FFC7CE", "font_color": "#9C0006"},  # red
})

book.save("output.xlsx")
```

## Rule Types

| Type | Description | Notes |
|------|-------------|-----------|
| `cellIs` | Compare cell value | Writing supported (operator/formula + basic colors) |
| `expression` | Formula-based condition | Writing supported (formula + basic colors) |
| `colorScale` | Gradient fill | Currently read-only in the Python API |
| `dataBar` | In-cell bar chart | Currently read-only in the Python API |
| `iconSet` | Icon set | Currently read-only in the Python API |
