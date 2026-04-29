# Cross-Language Pivot Context

This artifact exists because the main cross-language benchmark lane does not score pivot tables on macOS.

## Fixture Check

- Shipped fixture contains pivot OOXML parts: **No**
- Fixture path: `fixtures/excel/tier2/15_pivot_tables.xlsx`
- Fixture pivot-related parts: none detected

## Helper Detection

| Helper | Available | Detects pivots in shipped fixture | Notes |
|---|---:|---:|---|
| apache-poi | Yes | No | package metadata helper |
| excelize | Yes | No | sheet metadata helper |

## Write Probes

| Tool | Pivot write support | Evidence |
|---|---:|---|
| apache-poi | No | ApachePoiAdapter does not implement pivot table creation yet. |
| excelize | Yes | OOXML parts + openpyxl readback |

### Excelize Probe Details

- Output workbook: `results-cross-language-pivots/excelize-pivot-probe.xlsx`
- Pivot-related OOXML parts:
  - `xl/pivotCache/pivotCacheDefinition1.xml`
  - `xl/pivotTables/_rels/pivotTable1.xml.rels`
  - `xl/pivotTables/pivotTable1.xml`
- Openpyxl readback:
```json
[
  {
    "name": "SalesPivot",
    "source_range": "Data!A1:D5",
    "target_cell": "Pivot!B3:E10"
  }
]
```

## Interpretation

- The main cross-language scorecard remains correct to mark `pivot_tables` as not scored on macOS.
- The shipped fixture currently does not provide scoreable pivot evidence for this lane.
- `excelize` can still emit pivot-bearing workbooks, and this artifact captures that separately from the main scorecard.
