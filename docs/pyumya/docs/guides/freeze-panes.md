# Freeze Panes

Lock rows and columns so they stay visible while scrolling.

## Reading Freeze Pane Settings

```python
from wolfxl._rust import UmyaBook

book = UmyaBook.open("dashboard.xlsx")
panes = book.read_freeze_panes("Sheet1")
print(panes)  # {"mode": "freeze", "top_left_cell": "A2"}  (top row frozen)
```

## Writing Freeze Panes

```python
from wolfxl._rust import UmyaBook

book = UmyaBook()
book.add_sheet("Data")

# Freeze panes are configured via:
# - mode: "freeze" | "split"
# - top_left_cell: e.g. "B2" (the first scrollable cell)

# To freeze the top row (headers stay visible):
# book.set_freeze_panes("Data", {"mode": "freeze", "top_left_cell": "A2"})

# To freeze the first column:
# book.set_freeze_panes("Data", {"mode": "freeze", "top_left_cell": "B1"})

# To freeze both top row + first column:
book.set_freeze_panes("Data", {"mode": "freeze", "top_left_cell": "B2"})

book.save("output.xlsx")
```

## Common Patterns

| Use case | Settings | Excel equivalent |
|----------|----------|-----------------|
| Header row | `{"mode": "freeze", "top_left_cell": "A2"}` | View > Freeze Top Row |
| First column | `{"mode": "freeze", "top_left_cell": "B1"}` | View > Freeze First Column |
| Both | `{"mode": "freeze", "top_left_cell": "B2"}` | Select B2 > Freeze Panes |
| Multi-row header | `{"mode": "freeze", "top_left_cell": "A4"}` | Select A4 > Freeze Panes |

!!! tip
    `top_left_cell` is the first scrollable cell. Rows above and columns to
    the left of that cell become frozen. This matches the cell you'd select
    in Excel before clicking "Freeze Panes."
