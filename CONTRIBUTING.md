# Contributing to ExcelBench

## Quick Start
```bash
uv sync --extra dev
```

## Generate Fixtures (Excel required)
```bash
uv run excelbench generate --output fixtures/excel
```

## Run Benchmark
```bash
uv run excelbench benchmark --tests fixtures/excel --output results
```

## Adding a Feature
1. **Generator**: Add a new generator in `src/excelbench/generator/features/`.
2. **Manifest**: Register it in `src/excelbench/generator/features/__init__.py` and `generate.py`.
3. **Harness**: Add read/write logic in `src/excelbench/harness/runner.py`.
4. **Adapters**: Ensure required adapter methods exist for the feature.
5. **Fixtures**: Regenerate `fixtures/excel/`.

## Adding an Adapter
1. Implement the adapter in `src/excelbench/harness/adapters/`.
2. Export it from `src/excelbench/harness/adapters/__init__.py`.
3. Add to `get_all_adapters()` if it should run by default.
4. Verify read/write capability flags.

## Tests
```bash
pytest
```

## Lint / Type Check
```bash
ruff check
mypy
```

`mypy` uses the target set in `pyproject.toml` (`[tool.mypy].files`). CI currently checks a
small, strictly-typed subset; expand it incrementally as typing coverage improves.
