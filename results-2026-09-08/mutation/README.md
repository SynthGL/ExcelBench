# ExcelBench Mutation Results

*Generated: 2026-09-08T05:42:21Z*
*Template SHA-256: 74de96e1ea2b6ee8d52e9ff13b51db36d9140b371194afa1b70bba7158feca5b*
*Repeats: 3*

## Comparison

| Engine | Wall ms | Peak RSS KB | Preservation score | Lost parts | Verdict |
|--------|---------|-------------|--------------------|------------|---------|
| aspose-cells-foss | 770.000 | 170848.000 | 60.000 | 3 | Loss detected |
| openpyxl | 750.000 | 170336.000 | 60.000 | 3 | Loss detected |
| wolfxl | 780.000 | 175696.000 | 100.000 | 0 | Preserved |
| zavora-xlsx | 1320.000 | 169568.000 | 0.000 | 1 | Mutation integrity failed |

## Notes

Wall time and peak RSS are measured in isolated subprocesses with macOS `/usr/bin/time -l`.
Preservation compares each output package against the original template. It does not measure formula recalculation or rendering.
