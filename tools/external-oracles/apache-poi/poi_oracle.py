#!/usr/bin/env python3
"""JSON stdin/stdout wrapper for the Apache POI helper."""

from __future__ import annotations

import json
import tempfile
import shutil
import subprocess
import sys
import zipfile
from base64 import b64encode
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
CLASSES = ROOT / "build" / "classes"
LIB = ROOT / "deps" / "lib"
JAVA = Path("/opt/homebrew/opt/openjdk/bin/java")

JSONDict = dict[str, Any]


def main() -> int:
    """Run the requested helper operation."""
    request = json.load(sys.stdin)
    operation = request["operation"]
    if operation == "write_fixture":
        return _write_fixture(request)
    if operation == "write_adapter_workbook":
        return _write_adapter_workbook(request)
    if operation == "read_metadata":
        _write_json(_read_metadata(request))
        return 0
    _write_json({"error": "apache_poi_oracle_failed", "message": f"Unsupported {operation}"})
    return 1


def _write_fixture(request: JSONDict) -> int:
    output_path = request.get("output_path")
    if not output_path:
        _write_json(
            {
                "error": "apache_poi_oracle_failed",
                "message": "write_fixture requires output_path",
            }
        )
        return 1
    java = _java_executable()
    classpath = _classpath()
    if java is None or classpath is None:
        _write_json({"skipped": True, "notes": "Apache POI helper is not built"})
        return 0
    completed = subprocess.run(
        [
            java,
            "-cp",
            classpath,
            "PoiOracle",
            "write_fixture",
            request["fixture_id"],
            output_path,
        ],
        text=True,
        capture_output=True,
        check=False,
        timeout=180,
    )
    if completed.returncode != 0:
        _write_json(
            {
                "error": "apache_poi_oracle_failed",
                "message": completed.stderr.strip() or completed.stdout.strip(),
            }
        )
        return completed.returncode
    json_line = next(
        line for line in reversed(completed.stdout.splitlines()) if line.strip().startswith("{")
    )
    _write_json(json.loads(json_line))
    return 0


def _write_adapter_workbook(request: JSONDict) -> int:
    output_path = request.get("output_path")
    payload = request.get("payload") or {}
    if not output_path:
        _write_json(
            {
                "error": "apache_poi_oracle_failed",
                "message": "write_adapter_workbook requires output_path",
            }
        )
        return 1
    java = _java_executable()
    classpath = _classpath()
    if java is None or classpath is None:
        _write_json({"skipped": True, "notes": "Apache POI helper is not built"})
        return 0
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".spec", delete=False) as spec:
        spec_path = Path(spec.name)
        try:
            spec.write(_serialize_adapter_payload(payload))
        finally:
            spec.flush()
    try:
        completed = subprocess.run(
            [
                java,
                "-cp",
                classpath,
                "PoiOracle",
                "write_adapter_workbook",
                request.get("fixture_id", "apache-poi-adapter"),
                output_path,
                str(spec_path),
            ],
            text=True,
            capture_output=True,
            check=False,
            timeout=180,
        )
    finally:
        spec_path.unlink(missing_ok=True)
    if completed.returncode != 0:
        _write_json(
            {
                "error": "apache_poi_oracle_failed",
                "message": completed.stderr.strip() or completed.stdout.strip(),
            }
        )
        return completed.returncode
    json_line = next(
        line for line in reversed(completed.stdout.splitlines()) if line.strip().startswith("{")
    )
    _write_json(json.loads(json_line))
    return 0


def _read_metadata(request: JSONDict) -> JSONDict:
    input_path = request.get("input_path") or request.get("output_path")
    if not input_path:
        return {"error": "apache_poi_oracle_failed", "message": "read_metadata requires input_path"}
    with zipfile.ZipFile(input_path) as workbook:
        names = workbook.namelist()
    return {
        "fixture_id": request["fixture_id"],
        "operation": request["operation"],
        "input_path": input_path,
        "tool": "apache-poi",
        "counts": {
            "worksheets": _count(names, "xl/worksheets/sheet"),
            "shared_strings": _count(names, "xl/sharedStrings"),
            "comments": _count(names, "xl/comments"),
            "vml_drawings": _count(names, "xl/drawings/vmlDrawing"),
            "tables": _count(names, "xl/tables/table"),
            "drawings": _count(names, "xl/drawings/drawing"),
            "media": _count(names, "xl/media/"),
            "calc_chain": _count(names, "xl/calcChain"),
        },
    }


def _classpath() -> str | None:
    jars = sorted(LIB.glob("*.jar"))
    if not jars or not (CLASSES / "PoiOracle.class").exists():
        return None
    return ":".join([str(CLASSES), *(str(jar) for jar in jars)])


def _java_executable() -> str | None:
    if JAVA.exists():
        return str(JAVA)
    return shutil.which("java")


def _count(names: list[str], fragment: str) -> int:
    return sum(1 for name in names if fragment in name)


def _encode(value: str) -> str:
    return b64encode(value.encode("utf-8")).decode("ascii")


def _serialize_adapter_payload(payload: JSONDict) -> str:
    lines: list[str] = []
    for sheet in payload.get("sheets", []):
        lines.append(f"SHEET\t{_encode(str(sheet))}")
    for entry in payload.get("values", []):
        lines.append(
            "\t".join(
                [
                    "VALUE",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry["cell"])),
                    _encode(str(entry["type"])),
                    _encode(str(entry.get("value", ""))),
                    _encode(str(entry.get("formula", ""))),
                ]
            )
        )
    for entry in payload.get("formats", []):
        lines.append(
            "\t".join(
                [
                    "FORMAT",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry["cell"])),
                    "1" if entry.get("bold") else "0",
                    "1" if entry.get("italic") else "0",
                    _encode(str(entry.get("underline") or "")),
                    "1" if entry.get("strikethrough") else "0",
                    _encode(str(entry.get("font_name") or "")),
                    _encode(str(entry.get("font_size") or "")),
                    _encode(str(entry.get("font_color") or "")),
                    _encode(str(entry.get("bg_color") or "")),
                    _encode(str(entry.get("number_format") or "")),
                    _encode(str(entry.get("h_align") or "")),
                    _encode(str(entry.get("v_align") or "")),
                    "1" if entry.get("wrap") else "0",
                    _encode(str(entry.get("rotation") or "")),
                    _encode(str(entry.get("indent") or "")),
                ]
            )
        )
    for entry in payload.get("borders", []):
        lines.append(
            "\t".join(
                [
                    "BORDER",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry["cell"])),
                    _encode(json.dumps(entry["border"], sort_keys=True, separators=(",", ":"))),
                ]
            )
        )
    for entry in payload.get("conditional_formats", []):
        lines.append(
            "\t".join(
                [
                    "CF",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry.get("range") or "")),
                    _encode(str(entry.get("rule_type") or "")),
                    _encode(str(entry.get("operator") or "")),
                    _encode(str(entry.get("formula") or "")),
                    "1" if entry.get("stop_if_true") else "0",
                    _encode(str((entry.get("format") or {}).get("bg_color") or "")),
                ]
            )
        )
    for entry in payload.get("row_heights", []):
        lines.append(
            "\t".join(
                [
                    "ROW_HEIGHT",
                    _encode(str(entry["sheet"])),
                    str(entry["row"]),
                    str(entry["height"]),
                ]
            )
        )
    for entry in payload.get("column_widths", []):
        lines.append(
            "\t".join(
                [
                    "COL_WIDTH",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry["column"])),
                    str(entry["width"]),
                ]
            )
        )
    for entry in payload.get("merges", []):
        lines.append(
            "\t".join(["MERGE", _encode(str(entry["sheet"])), _encode(str(entry["range"]))])
        )
    for entry in payload.get("validations", []):
        lines.append(
            "\t".join(
                [
                    "VALIDATION",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry.get("range") or entry.get("cell") or "")),
                    _encode(str(entry.get("validation_type") or "")),
                    _encode(str(entry.get("operator") or "")),
                    _encode(str(entry.get("formula1") or "")),
                    _encode(str(entry.get("formula2") or "")),
                    "1" if entry.get("allow_blank", True) else "0",
                    _encode(str(entry.get("error_title") or "")),
                    _encode(str(entry.get("error") or "")),
                ]
            )
        )
    for entry in payload.get("hyperlinks", []):
        lines.append(
            "\t".join(
                [
                    "HYPERLINK",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry["cell"])),
                    _encode(str(entry.get("target") or entry.get("url") or "")),
                    _encode(str(entry.get("display") or entry.get("label") or "")),
                    _encode(str(entry.get("tooltip") or "")),
                    "1" if entry.get("internal") else "0",
                ]
            )
        )
    for entry in payload.get("comments", []):
        lines.append(
            "\t".join(
                [
                    "COMMENT",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry["cell"])),
                    _encode(str(entry.get("text") or "")),
                    _encode(str(entry.get("author") or "ExcelBench")),
                ]
            )
        )
    for entry in payload.get("images", []):
        lines.append(
            "\t".join(
                [
                    "IMAGE",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry.get("cell") or "")),
                    _encode(str(entry.get("path") or "")),
                    _encode(str(entry.get("anchor") or "")),
                ]
            )
        )
    for entry in payload.get("named_ranges", []):
        lines.append(
            "\t".join(
                [
                    "NAME",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry.get("name") or "")),
                    _encode(str(entry.get("scope") or "workbook")),
                    _encode(str(entry.get("refers_to") or "")),
                ]
            )
        )
    for entry in payload.get("freeze_panes", []):
        lines.append(
            "\t".join(
                [
                    "FREEZE",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry.get("mode") or "freeze")),
                    str(entry.get("x_split", 0)),
                    str(entry.get("y_split", 0)),
                    _encode(str(entry.get("top_left_cell") or "")),
                ]
            )
        )
    for entry in payload.get("tables", []):
        lines.append(
            "\t".join(
                [
                    "TABLE",
                    _encode(str(entry["sheet"])),
                    _encode(str(entry.get("ref") or entry.get("range") or "")),
                    _encode(str(entry.get("name") or "")),
                    _encode(str(entry.get("style") or "")),
                    "1" if entry.get("totals_row") else "0",
                    "1" if entry.get("autofilter") else "0",
                ]
            )
        )
    return "\n".join(lines) + ("\n" if lines else "")


def _write_json(payload: JSONDict) -> None:
    print(json.dumps(payload, sort_keys=True))


if __name__ == "__main__":
    raise SystemExit(main())
