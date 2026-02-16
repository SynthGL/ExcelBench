# Hyperlinks

Add and read hyperlinks in Excel cells.

## Reading Hyperlinks

```python
from wolfxl._rust import UmyaBook

book = UmyaBook.open("links.xlsx")
links = book.read_hyperlinks("Sheet1")
for link in links:
    print(f"{link['cell']}: {link['target']} ({link.get('display', '')})")
# A1: https://example.com (Example Site)
# B2: mailto:user@example.com (Contact Us)
```

## Writing Hyperlinks

```python
from wolfxl._rust import UmyaBook

book = UmyaBook()
book.add_sheet("Links")

# Web URL
book.add_hyperlink("Links", {
    "cell": "A1",
    "target": "https://example.com",
    "display": "Visit Example",
})

# Email link
book.add_hyperlink("Links", {
    "cell": "A2",
    "target": "mailto:support@example.com",
    "display": "Email Support",
})

# Internal reference (another sheet)
book.add_hyperlink("Links", {
    "cell": "A3",
    "target": "#Summary!A1",
    "display": "Go to Summary",
    "internal": True,
})

book.save("output.xlsx")
```

## Hyperlink Types

| Type | Target format | Example |
|------|--------------|---------|
| Web URL | `https://...` | `https://example.com` |
| Email | `mailto:...` | `mailto:user@example.com` |
| Internal | `#SheetName!Cell` | `#Summary!A1` |
| File | Relative or absolute path | `../other.xlsx` |

## Styling

!!! note
    Excel automatically applies blue underline formatting to hyperlinked cells.
    pyumya does not auto-apply this styling — if you want the visual cue,
    apply font formatting separately.
