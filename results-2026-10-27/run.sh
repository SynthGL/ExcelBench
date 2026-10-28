#!/bin/bash
# Run from the repository root with the pinned environment already installed.
# EB_VENV names an environment containing ExcelBench and requirements.txt.
set -euo pipefail
OUT=results-2026-10-27
EB="${EB_VENV:-.venv-benchmark-native}/bin/excelbench"
mkdir -p "$OUT/calc"
echo "start_utc $(date -u +%Y-%m-%dT%H:%M:%SZ)"
echo "load_before $(sysctl -n vm.loadavg)"
/usr/bin/time -p "$EB" calc -o "$OUT/calc" > "$OUT/calc-run.txt" 2>&1
echo "calc_exit $?"
echo "perf_start_utc $(date -u +%Y-%m-%dT%H:%M:%SZ)"
echo "perf_load_before $(sysctl -n vm.loadavg)"
/usr/bin/time -p "$EB" perf --tests fixtures/excel --output "$OUT" --profile xlsx \
  --warmup 3 --iters 25 --iteration-policy fixed --memory-mode getrusage \
  --feature cell_values --feature formulas --feature text_formatting --feature background_colors \
  --feature number_formats --feature alignment --feature borders --feature dimensions \
  --feature multiple_sheets --feature merged_cells --feature conditional_formatting \
  --feature data_validation --feature hyperlinks --feature images --feature pivot_tables \
  --feature comments --feature freeze_panes --feature named_ranges --feature tables \
  --adapter openpyxl --adapter xlsxwriter --adapter python-calamine --adapter wolfxl \
  --adapter pylightxl --adapter xlrd --adapter pyexcel --adapter xlwt --adapter pandas \
  --adapter xlsxwriter-constmem --adapter openpyxl-readonly --adapter polars --adapter tablib \
  > "$OUT/perf-run.txt" 2>&1
echo "perf_exit $?"
echo "load_after $(sysctl -n vm.loadavg)"
echo "end_utc $(date -u +%Y-%m-%dT%H:%M:%SZ)"
