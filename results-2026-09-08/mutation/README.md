# ExcelBench Mutation Results

*Generated: 2026-09-08T06:28:04Z*
*Template SHA-256: da08cc9e607f267b438ec7ce2480d6eb3650d1c19928f283a18296a2b79d808e*
*Repeats: 3*

## Comparison

| Engine | Wall ms | Peak RSS KB | Preservation score | Lost parts | Verdict |
|--------|---------|-------------|--------------------|------------|---------|
| aspose-cells-foss | 690.000 | 170752.000 | 60.000 | 3 | Loss detected |
| openpyxl | 710.000 | 170336.000 | 60.000 | 3 | Loss detected |
| wolfxl | 710.000 | 175488.000 | 100.000 | 0 | Preserved |
| zavora-xlsx | 1270.000 | 169520.000 | 0.000 | 1 | Mutation integrity failed |

## Notes

Wall time and peak RSS are measured in isolated subprocesses with macOS `/usr/bin/time -l`.
Preservation compares each output package against the original template. It does not measure formula recalculation or rendering.
