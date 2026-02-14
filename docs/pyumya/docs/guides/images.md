# Images

Embed and read images in Excel worksheets.

## Reading Images

```python
from excelbench_rust import UmyaBook

book = UmyaBook.open("report.xlsx")
images = book.read_images("Sheet1")
for img in images:
    print(f"Cell {img['cell']}: {img['format']} ({len(img['data'])} bytes)")
# Cell A1: png (24576 bytes)
```

## Writing Images

```python
from pathlib import Path
from excelbench_rust import UmyaBook

book = UmyaBook()
book.add_sheet("Report")

# Embed an image from file
book.add_image("Report", "B2", {
    "data": Path("logo.png").read_bytes(),
    "format": "png",
})

book.save("output.xlsx")
```

## Supported Formats

| Format | Extension | Read | Write |
|--------|-----------|:----:|:-----:|
| PNG | `.png` | Yes | Yes |
| JPEG | `.jpg`, `.jpeg` | Yes | Yes |
| GIF | `.gif` | Yes | Yes |
| BMP | `.bmp` | Yes | Yes |
| EMF | `.emf` | Yes | Yes |

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
