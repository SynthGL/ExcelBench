# NPOI External Oracle

This optional helper generates `.xlsx` workbooks with NPOI, the .NET port of
Apache POI. It uses the same JSON stdin/stdout contract as the Excelize,
LibreOffice, and ClosedXML helpers.

Run from the repository root:

```bash
dotnet run --project tools/external-oracles/npoi/npoi-oracle.csproj --configuration Release --no-launch-profile --verbosity quiet --
```

Supported operations:

- `write_fixture`: writes sheets, cells, formulas, rich text, comments, merged
  regions, and sheet protection.
- `read_metadata`: inspects package parts for worksheets, shared strings,
  comments, VML drawings, and calc-chain metadata.

NPOI is intentionally a local pre-release oracle, not a normal ExcelBench
adapter or public benchmark claim.
