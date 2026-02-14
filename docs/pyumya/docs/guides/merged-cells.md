# Merged Cells

Merge and unmerge cell ranges in Excel workbooks.

## Reading Merged Ranges

```python
from excelbench_rust import UmyaBook

book = UmyaBook.open("report.xlsx")
merged = book.read_merged_cells("Sheet1")
print(merged)  # ["A1:D1", "B3:B5"]
```

## Writing Merged Ranges

```python
book = UmyaBook()
book.add_sheet("Data")

# Write header spanning 4 columns
book.write_cell_value("Data", "A1", {"type": "string", "value": "Quarterly Report"})
book.merge_cells("Data", "A1:D1")

# Center the merged header
book.write_cell_format("Data", "A1", {"h_align": "center", "bold": True})

book.save("output.xlsx")
```

## How Excel Handles Merged Cells

!!! info "Value placement"
    When cells are merged, only the **top-left cell** retains the value.
    Reading any other cell in the merged range returns a blank value.
    This is an Excel specification behavior, not a pyumya limitation.

## Edge Cases

- Merging a range that overlaps an existing merge will raise an error in Excel
- Single-cell "merges" (e.g., `"A1:A1"`) are valid but have no visual effect
- Merged cells with borders apply the border to the entire merged region
