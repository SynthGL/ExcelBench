#!/usr/bin/env python3
"""JSON stdin/stdout wrapper for the Apache POI helper."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import zipfile
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


def _write_json(payload: JSONDict) -> None:
    print(json.dumps(payload, sort_keys=True))


if __name__ == "__main__":
    raise SystemExit(main())
