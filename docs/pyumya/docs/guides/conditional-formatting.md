# Conditional Formatting

Apply formatting rules that change cell appearance based on values.

## Reading Conditional Formats

```python
from excelbench_rust import UmyaBook

book = UmyaBook.open("dashboard.xlsx")
rules = book.read_conditional_formats("Sheet1")
for r in rules:
    print(f"{r['ranges']}: {r['type']}")
# ['A1:A100']: cellIs
# ['B1:B100']: colorScale
```

## Writing Conditional Formats

### Cell Value Rules

```python
book = UmyaBook()
book.add_sheet("Sales")

# Highlight cells greater than 1000
book.add_conditional_format("Sales", {
    "ranges": ["B2:B50"],
    "type": "cellIs",
    "operator": "greaterThan",
    "formula": "1000",
    "format": {"bg_color": "#C6EFCE", "font_color": "#006100"},  # green
})

# Highlight cells below target
book.add_conditional_format("Sales", {
    "ranges": ["B2:B50"],
    "type": "cellIs",
    "operator": "lessThan",
    "formula": "500",
    "format": {"bg_color": "#FFC7CE", "font_color": "#9C0006"},  # red
})

book.save("output.xlsx")
```

### Color Scales

```python
# 2-color scale (red to green)
book.add_conditional_format("Sales", {
    "ranges": ["C2:C50"],
    "type": "colorScale",
    "color_scale": {
        "min_color": "#FF0000",
        "max_color": "#00FF00",
    },
})

# 3-color scale (red / yellow / green)
book.add_conditional_format("Sales", {
    "ranges": ["D2:D50"],
    "type": "colorScale",
    "color_scale": {
        "min_color": "#FF0000",
        "mid_color": "#FFFF00",
        "max_color": "#00FF00",
    },
})
```

### Data Bars

```python
book.add_conditional_format("Sales", {
    "ranges": ["E2:E50"],
    "type": "dataBar",
    "data_bar": {"color": "#638EC6"},
})
```

## Rule Types

| Type | Description | Key fields |
|------|-------------|-----------|
| `cellIs` | Compare cell value | `operator`, `formula`, `format` |
| `colorScale` | Gradient fill | `color_scale` with 2-3 colors |
| `dataBar` | In-cell bar chart | `data_bar` with color |
| `top10` | Top/bottom N | `rank`, `percent`, `bottom` |
| `containsText` | Text matching | `text`, `format` |

## Rule Priority

!!! info "Evaluation order"
    When multiple rules apply to the same cell, rules are evaluated
    in the order they were added. The first matching rule determines
    the formatting. This matches Excel's "Stop if True" default behavior.
