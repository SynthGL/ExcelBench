#!/bin/zsh
set -u
cd "$HOME/wolf-lane/officelibs-perf-2026-10-03/src"
F=(cell_values formulas text_formatting background_colors number_formats alignment borders dimensions multiple_sheets merged_cells conditional_formatting data_validation hyperlinks images pivot_tables comments freeze_panes named_ranges tables)
A=(openpyxl xlsxwriter python-calamine wolfxl pylightxl xlrd pyexcel xlwt pandas xlsxwriter-constmem openpyxl-readonly polars tablib)
args=()
for f in $F; do args+=(--feature $f); done
for a in $A; do args+=(--adapter $a); done
until [[ "$(date -u +%F)" == "2026-10-03" ]]; do sleep 15; done
echo "date_gate_passed $(date -u +%Y-%m-%dT%H:%M:%SZ)"
for i in $(seq 1 60); do
  l1=$(sysctl -n vm.loadavg | awk "{print \$2}")
  echo "gate_check $(date -u +%H:%M:%SZ) load1=$l1"
  if (( l1 <= 4.0 )); then break; fi
  sleep 10
done
echo "start_utc $(date -u +%Y-%m-%dT%H:%M:%SZ)"
echo "load_before $(sysctl -n vm.loadavg)"
ps -Ao pcpu,comm -r | head -4
echo "CMD: excelbench perf --tests fixtures/excel --output ../out --profile xlsx --warmup 3 --iters 25 --iteration-policy fixed --memory-mode getrusage ${args[*]}"
/usr/bin/time -p ../venv/bin/excelbench perf --tests fixtures/excel --output ../out --profile xlsx --warmup 3 --iters 25 --iteration-policy fixed --memory-mode getrusage "${args[@]}"
rc=$?
echo "exit_code $rc"
echo "load_after $(sysctl -n vm.loadavg)"
echo "end_utc $(date -u +%Y-%m-%dT%H:%M:%SZ)"
