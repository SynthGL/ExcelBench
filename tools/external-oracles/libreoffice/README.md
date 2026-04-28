# LibreOffice External Oracle

Optional helper for validating `.xlsx` workbooks through LibreOffice Calc in
headless mode. It is a renderer/open-save oracle, not a normal benchmark
adapter.

## Run

```bash
python libreoffice_oracle.py < request.json
```

The helper locates LibreOffice through `LIBREOFFICE_BIN`, `soffice`,
`libreoffice`, or `/Applications/LibreOffice.app/Contents/MacOS/soffice`.
Missing LibreOffice returns a structured skip.

## Operations

- `open_save_validate`: converts an input workbook back to `.xlsx` using the
  `Calc Office Open XML` filter.
- `render_validate` / `render_pdf`: exports an input workbook to PDF using the
  `calc_pdf_Export` filter.

