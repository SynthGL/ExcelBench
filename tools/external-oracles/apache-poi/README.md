# Apache POI External Oracle

This optional helper generates `.xlsx` workbooks with Apache POI. It uses the
same JSON stdin/stdout contract as the Excelize, LibreOffice, ClosedXML, NPOI,
and ExcelJS helpers.

Bootstrap from this directory:

```bash
./build.sh
```

`fetch_deps.py` downloads a pinned Maven Central dependency set for Apache POI
5.5.1 into `deps/lib/` and verifies SHA-256 checksums before compilation.

Supported operations:

- `write_fixture`: writes the current POI fixture with tables, formulas,
  comments, rich text, hyperlinks, data validation, merged ranges, freeze panes,
  and sheet protection.
- `read_metadata`: inspects package parts for worksheets, shared strings,
  comments, VML drawings, tables, drawings, media, and calc-chain metadata.

Apache POI is intentionally a local pre-release oracle, not a normal ExcelBench
adapter or public benchmark claim.
