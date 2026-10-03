#!/bin/zsh
# Runs the 2026-10-03 competitor snapshot lanes (fidelity, cross-language,
# mutation, calc) from the repository root. Perf is a separate run.
#
# Environment:
#   EB_VENV        venv with ExcelBench installed (-e .) and the pins in pins.txt
#   EB_COMMERCIAL  interpreter with the WolfXL Commercial 2.3.0 wheel installed
set -u
OUT=results-2026-10-03
EB="$EB_VENV/bin/excelbench"

if [[ "$(date -u +%F)" != "2026-10-03" ]]; then
  echo "refusing to run: UTC date is $(date -u +%F)"
  exit 2
fi

echo "start_utc $(date -u +%Y-%m-%dT%H:%M:%SZ)"
echo "load_start $(sysctl -n vm.loadavg)"

# Fidelity and cross-language run side by side (cross-language work runs on the
# remote Docker host or in zavora's own process; no shared LibreOffice profile).
"$EB" benchmark -t fixtures/excel -o "$OUT/xlsx" \
  -a openpyxl -a xlsxwriter -a python-calamine -a aspose-cells-foss -a wolfxl \
  -a pylightxl -a xlrd -a pyexcel -a xlwt -a pandas -a xlsxwriter-constmem \
  -a openpyxl-readonly -a polars -a tablib -a sheetjs -a exceljs -a libreoffice \
  > "$OUT/logs/xlsx.log" 2>&1 &
xlsx_pid=$!
EXCELBENCH_ORACLE_DOCKER_CONTEXT=pc "$EB" cross-language-context \
  -t fixtures/excel -o "$OUT/cross-language" > "$OUT/logs/cross-language.log" 2>&1 &
xl_pid=$!
wait $xlsx_pid; echo "xlsx_exit $?"
wait $xl_pid; echo "cross_language_exit $?"
"$EB" heatmap -i "$OUT/xlsx/results.json" -o "$OUT/xlsx" > "$OUT/logs/heatmap.log" 2>&1
echo "heatmap_exit $?"

# Mutation timings: wait (up to 10 min) for the 1-minute load to drop to 8 or less.
for i in $(seq 1 60); do
  l1=$(sysctl -n vm.loadavg | awk '{print $2}')
  echo "gate_check $(date -u +%H:%M:%SZ) load1=$l1"
  if (( l1 <= 8.0 )); then break; fi
  sleep 10
done
echo "mutation_load_before $(sysctl -n vm.loadavg)"
"$EB" mutation -o "$OUT/mutation" --repeats 3 > "$OUT/logs/mutation.log" 2>&1
echo "mutation_exit $?"
echo "mutation_load_after $(sysctl -n vm.loadavg)"

EXCELBENCH_WOLFXL_COMMERCIAL_PYTHON="$EB_COMMERCIAL" "$EB" calc -o "$OUT/calc" \
  > "$OUT/logs/calc.log" 2>&1
echo "calc_exit $?"
echo "end_utc $(date -u +%Y-%m-%dT%H:%M:%SZ)"
