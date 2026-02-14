# Freeze Panes

Lock rows and columns so they stay visible while scrolling.

## Reading Freeze Pane Settings

```python
from excelbench_rust import UmyaBook

book = UmyaBook.open("dashboard.xlsx")
panes = book.read_freeze_panes("Sheet1")
print(panes)  # {"row": 1, "column": 0}  (top row frozen)
```

## Writing Freeze Panes

```python
book = UmyaBook()
book.add_sheet("Data")

# Freeze top row (headers stay visible)
book.set_freeze_panes("Data", {"row": 1, "column": 0})

# Freeze first column
book.set_freeze_panes("Data", {"row": 0, "column": 1})

# Freeze both (top-left corner stays fixed)
book.set_freeze_panes("Data", {"row": 1, "column": 1})

book.save("output.xlsx")
```

## Common Patterns

| Use case | Settings | Excel equivalent |
|----------|----------|-----------------|
| Header row | `{"row": 1, "column": 0}` | View > Freeze Top Row |
| First column | `{"row": 0, "column": 1}` | View > Freeze First Column |
| Both | `{"row": 1, "column": 1}` | Select B2 > Freeze Panes |
| Multi-row header | `{"row": 3, "column": 0}` | Select A4 > Freeze Panes |

!!! tip
    The `row` and `column` values specify the **split position** — rows above
    and columns to the left of that position are frozen. This matches the
    cell you'd select in Excel before clicking "Freeze Panes."
