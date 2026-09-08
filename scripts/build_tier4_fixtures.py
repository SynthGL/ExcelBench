#!/usr/bin/env python3
"""Build Tier-4 structural fixtures: sheet_protection, page_setup, chart_anchor.

No Excel application is involved: workbooks are constructed with openpyxl
through OpenpyxlAdapter's Tier-4 methods, and every manifest expectation is
recorded from the post-save readback, so expectations always equal roundtrip
reality.

Outputs under <fixtures_root> (default: fixtures/excel):
  tier4/20_sheet_protection.xlsx
  tier4/21_page_setup.xlsx
  tier4/22_chart_anchor.xlsx
  tier4_manifest_fragment.json   (to be merged into manifest.json)
"""

from __future__ import annotations

import json
import sys
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.models import CellType, CellValue

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_FIXTURES_ROOT = REPO_ROOT / "fixtures" / "excel"
TIER4_DIR = "tier4"
GENERATOR = "openpyxl-structural"
GENERATOR_VERSION = "0.1.0"

JSONDict = dict[str, Any]

PROTECTION_CASES: list[JSONDict] = [
    {
        "sheet": "sheet_protection_basic",
        "id": "prot_basic",
        "label": "Protection: basic (OOXML defaults)",
        "importance": "basic",
        "settings": {"protected": True},
    },
    {
        "sheet": "sheet_protection_password",
        "id": "prot_password",
        "label": "Protection: with password hash",
        "importance": "basic",
        "settings": {"protected": True, "password": "bench-secret"},
    },
    {
        "sheet": "sheet_protection_granular",
        "id": "prot_granular",
        "label": "Protection: granular attribute flags",
        "importance": "edge",
        "settings": {
            "protected": True,
            "format_cells": False,
            "insert_rows": True,
            "select_locked_cells": True,
            "sort": False,
            "auto_filter": True,
        },
    },
]

PAGE_SETUP_CASES: list[JSONDict] = [
    {
        "sheet": "page_setup_landscape",
        "id": "page_setup_landscape",
        "label": "Page setup: landscape + fit-to-width",
        "importance": "basic",
        "settings": {"orientation": "landscape", "fit_to_width": 1},
    },
    {
        "sheet": "page_setup_scale",
        "id": "page_setup_scale",
        "label": "Page setup: portrait at 75% scale",
        "importance": "basic",
        "settings": {"orientation": "portrait", "scale": 75},
    },
    {
        "sheet": "page_setup_titles",
        "id": "page_setup_titles",
        "label": "Page setup: print titles + header/footer",
        "importance": "edge",
        "settings": {
            "print_title_rows": "$1:$2",
            "header_center": "ExcelBench Report",
            "footer_center": "Page &P",
        },
    },
]

CHART_DATA_NUMBERS = [12, 25, 18, 31, 22, 40, 29]  # E2:E8
CHART_DATA_LABELS = ["Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7"]  # F2:F8

CHART_CASES: list[JSONDict] = [
    {
        "id": "chart_bar",
        "label": "Chart: bar with twoCell anchor",
        "importance": "basic",
        "chart": {
            "type": "bar",
            "anchor_type": "twoCell",
            "from": "B2",
            "to": "H12",
            "data_ref": "E2:E8",
            "categories_ref": "F2:F8",
        },
    },
    {
        "id": "chart_line",
        "label": "Chart: line with twoCell anchor",
        "importance": "edge",
        "chart": {
            "type": "line",
            "anchor_type": "twoCell",
            "from": "J2",
            "to": "P12",
            "data_ref": "E2:E8",
            "categories_ref": "F2:F8",
        },
    },
]


def build_sheet_protection(adapter: OpenpyxlAdapter, out_dir: Path) -> JSONDict:
    """Build 20_sheet_protection.xlsx and return its manifest entry."""
    path = out_dir / "20_sheet_protection.xlsx"
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "sheet_protection")
    for case in PROTECTION_CASES:
        adapter.add_sheet(workbook, case["sheet"])
    for case in PROTECTION_CASES:
        adapter.set_sheet_protection(workbook, case["sheet"], case["settings"])
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    try:
        test_cases = []
        for case in PROTECTION_CASES:
            expected = dict(adapter.read_sheet_protection(reopened, case["sheet"]))
            password = case["settings"].get("password")
            if password is not None:
                expected["password"] = password
            test_cases.append(_test_case(case, {"protection": expected}))
    finally:
        adapter.close_workbook(reopened)
    return _manifest_entry(path.name, "sheet_protection", test_cases)


def build_page_setup(adapter: OpenpyxlAdapter, out_dir: Path) -> JSONDict:
    """Build 21_page_setup.xlsx and return its manifest entry."""
    path = out_dir / "21_page_setup.xlsx"
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "page_setup")
    for case in PAGE_SETUP_CASES:
        adapter.add_sheet(workbook, case["sheet"])
    for case in PAGE_SETUP_CASES:
        adapter.set_page_setup(workbook, case["sheet"], case["settings"])
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    try:
        test_cases = [
            _test_case(
                case, {"page_setup": adapter.read_page_setup(reopened, case["sheet"])}
            )
            for case in PAGE_SETUP_CASES
        ]
    finally:
        adapter.close_workbook(reopened)
    return _manifest_entry(path.name, "page_setup", test_cases)


def build_chart_anchor(adapter: OpenpyxlAdapter, out_dir: Path) -> JSONDict:
    """Build 22_chart_anchor.xlsx and return its manifest entry."""
    path = out_dir / "22_chart_anchor.xlsx"
    workbook = adapter.create_workbook()
    adapter.add_sheet(workbook, "chart_anchor")
    for offset, value in enumerate(CHART_DATA_NUMBERS):
        adapter.write_cell_value(
            workbook,
            "chart_anchor",
            f"E{2 + offset}",
            CellValue(type=CellType.NUMBER, value=value),
        )
    for offset, label in enumerate(CHART_DATA_LABELS):
        adapter.write_cell_value(
            workbook,
            "chart_anchor",
            f"F{2 + offset}",
            CellValue(type=CellType.STRING, value=label),
        )
    for case in CHART_CASES:
        adapter.add_chart_with_anchor(workbook, "chart_anchor", case["chart"])
    adapter.save_workbook(workbook, path)

    reopened = adapter.open_workbook(path)
    try:
        charts_by_from = {
            chart["from"]: chart
            for chart in adapter.read_chart_anchors(reopened, "chart_anchor")
        }
    finally:
        adapter.close_workbook(reopened)

    test_cases = []
    for case in CHART_CASES:
        spec = case["chart"]
        readback = charts_by_from.get(spec["from"])
        if readback is None:
            raise RuntimeError(f"chart anchor readback missing for {spec['from']}")
        expected = {
            **readback,
            "data_ref": spec["data_ref"],
            "categories_ref": spec["categories_ref"],
        }
        test_cases.append(_test_case(case, {"charts": [expected]}))
    return _manifest_entry(path.name, "chart_anchor", test_cases)


def _test_case(case: JSONDict, expected: JSONDict) -> JSONDict:
    entry: JSONDict = {
        "id": case["id"],
        "label": case["label"],
        "row": 2,
    }
    if case.get("sheet"):
        entry["sheet"] = case["sheet"]
    entry["expected"] = expected
    entry["importance"] = case["importance"]
    return entry


def _manifest_entry(
    filename: str, feature: str, test_cases: list[JSONDict]
) -> JSONDict:
    return {
        "path": f"{TIER4_DIR}/{filename}",
        "feature": feature,
        "tier": 4,
        "file_format": "xlsx",
        "test_cases": test_cases,
    }


def build_all(fixtures_root: Path = DEFAULT_FIXTURES_ROOT) -> Path:
    """Build all Tier-4 fixtures plus the manifest fragment.

    Returns the path to the written fragment.
    """
    out_dir = fixtures_root / TIER4_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    adapter = OpenpyxlAdapter()
    entries = [
        build_sheet_protection(adapter, out_dir),
        build_page_setup(adapter, out_dir),
        build_chart_anchor(adapter, out_dir),
    ]
    fragment: JSONDict = {
        "generated_at": datetime.now(UTC).isoformat(),
        "excel_version": "none",
        "generator_version": GENERATOR_VERSION,
        "generator": GENERATOR,
        "file_format": "xlsx",
        "files": entries,
    }
    fragment_path = fixtures_root / "tier4_manifest_fragment.json"
    with open(fragment_path, "w") as f:
        json.dump(fragment, f, indent=2)
        f.write("\n")
    return fragment_path


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    fixtures_root = Path(args[0]).resolve() if args else DEFAULT_FIXTURES_ROOT
    fragment_path = build_all(fixtures_root)
    print(f"Tier-4 fixtures written to {fixtures_root / TIER4_DIR}")
    print(f"Manifest fragment: {fragment_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
