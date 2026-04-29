# ClosedXML External Oracle

This optional helper lets ExcelBench generate .NET/ClosedXML workbooks through
the same JSON stdin/stdout contract as the other external oracles.

Run through the catalog from the repository root:

```bash
dotnet run --project tools/external-oracles/closedxml/closedxml-oracle.csproj --configuration Release --no-launch-profile --verbosity quiet --
```

Supported operations:

- `write_fixture`: writes workbook sheets, cells, tables, conditional formats,
  and pivot tables from the JSON request.
- `read_metadata`: inspects workbook package parts for table, pivot table,
  pivot cache, and worksheet counts.

ClosedXML is intentionally kept as an optional pre-release oracle. It is not a
normal ExcelBench adapter and should not become a public benchmark claim until
generated cases pass the manual truth/promote gates in
`docs/trackers/external-oracle-expansion.md`.
