# Excelize External Oracle

Optional Go helper for generating and inspecting `.xlsx` fixtures with
Excelize. It is intentionally outside the Python package runtime path: missing
Go or missing helper source should skip external-oracle checks, not fail the
normal ExcelBench suite.

## Run

From this directory:

```bash
go run . < request.json
go test ./...
```

From Python, use:

```python
from pathlib import Path
from excelbench.harness.external_oracles import external_oracle_catalog

repo_root = Path(".../ExcelBench")
tool = external_oracle_catalog(repo_root=repo_root)["excelize"]
```

## Operations

- `write_fixture`: writes an `.xlsx` workbook to `output_path`.
- `read_metadata`: opens `input_path` and reports sheet-level table, pivot,
  slicer, and conditional-formatting counts.

## Write Payload Keys

- `sheets`
- `cells`
- `columns`
- `tables`
- `conditional_formats`
- `charts`
- `pivots`
- `slicers`
- `pictures`

This helper is a fixture generator, not a public benchmark adapter yet. Promote
individual cases into canonical fixtures only after manual truth-passing.

