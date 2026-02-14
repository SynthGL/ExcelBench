# Images

Embed and read images in Excel worksheets.

## Reading Images

```python
from excelbench_rust import UmyaBook

book = UmyaBook.open("report.xlsx")
images = book.read_images("Sheet1")
for img in images:
    print(f"Image at {img['cell']}: anchor={img['anchor']}, offset={img['offset']}")
# Image at A1: anchor=oneCell, offset=[0, 0]
```

## Writing Images

```python
from excelbench_rust import UmyaBook

book = UmyaBook()
book.add_sheet("Report")

# Embed an image by file path
book.add_image("Report", {
    "path": "logo.png",
    "cell": "B2",
})

book.save("output.xlsx")
```

## Formats

The Python API currently exposes image **anchors** (positioning) but does not
expose image bytes or a detected format on read. For writes, provide a file
path (e.g. PNG/JPEG).

## Image Positioning

!!! note "Anchor behavior"
    Images are anchored to a cell position. When rows/columns are
    resized, the image moves with its anchor cell. The image size
    is determined by the original image dimensions — pyumya does
    not currently support explicit width/height overrides.

## Best Practices

- Use PNG for logos and diagrams (lossless, supports transparency)
- Use JPEG for photographs (smaller file size)
- Keep images under 1 MB for reasonable workbook file sizes
- Place images in dedicated "cover" or "chart" sheets to avoid layout issues
