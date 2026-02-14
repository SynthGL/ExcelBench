# Comments

Add, read, and manage cell comments (notes) in Excel workbooks.

## Reading Comments

```python
from excelbench_rust import UmyaBook

book = UmyaBook.open("annotated.xlsx")
comments = book.read_comments("Sheet1")
for c in comments:
    print(f"{c['cell']}: {c['text']} (by {c['author']})")
# A1: Review this value (by John)
# B3: Updated 2026-01-15 (by Jane)
```

## Writing Comments

```python
book = UmyaBook()
book.add_sheet("Data")

book.write_cell_value("Data", "A1", {"type": "number", "value": 42.0})
book.add_comment("Data", "A1", {
    "text": "This value needs verification",
    "author": "Reviewer",
})

book.save("output.xlsx")
```

## How Comments Appear in Excel

!!! tip
    Comments appear as hover tooltips in Excel. A small red triangle in the
    cell corner indicates a comment is present. In newer versions of Excel,
    these are called "Notes" (threaded "Comments" are a separate feature
    that pyumya does not currently support).

## Best Practices

- Keep comment text concise — long comments are hard to read in the tooltip
- Include dates in comments for audit trails
- Use the `author` field consistently across your organization
