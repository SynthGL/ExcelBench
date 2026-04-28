#!/usr/bin/env python3
"""LibreOffice subprocess helper for external oracle validation."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

JSONDict = dict[str, Any]


def main() -> int:
    try:
        request = json.load(sys.stdin)
        response, exit_code = handle_request(request)
    except Exception as exc:  # pragma: no cover - last-resort CLI guard
        response = {"error": "libreoffice_oracle_failed", "message": str(exc)}
        exit_code = 1
    print(json.dumps(response, sort_keys=True))
    return exit_code


def handle_request(request: JSONDict) -> tuple[JSONDict, int]:
    soffice = resolve_soffice()
    if soffice is None:
        return {
            "skipped": True,
            "notes": "LibreOffice executable not found",
            "tool": "libreoffice",
        }, 0

    operation = request.get("operation")
    if operation == "open_save_validate":
        return run_conversion(
            soffice=soffice,
            request=request,
            extension="xlsx",
            filter_name="Calc Office Open XML",
        )
    if operation in {"render_validate", "render_pdf"}:
        return run_conversion(
            soffice=soffice,
            request=request,
            extension="pdf",
            filter_name="calc_pdf_Export",
        )
    return {"error": "unsupported_operation", "message": f"unsupported operation {operation!r}"}, 1


def resolve_soffice() -> str | None:
    """Find a LibreOffice executable without assuming one install shape."""
    candidates = [
        os.environ.get("LIBREOFFICE_BIN"),
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    return None


def run_conversion(
    *,
    soffice: str,
    request: JSONDict,
    extension: str,
    filter_name: str,
) -> tuple[JSONDict, int]:
    """Run LibreOffice headless conversion and return structured diagnostics."""
    input_path = Path(str(request.get("input_path") or request.get("output_path") or ""))
    if not input_path:
        return {"error": "missing_input_path", "message": "input_path is required"}, 1
    if not input_path.exists():
        return {"error": "missing_input_path", "message": f"{input_path} does not exist"}, 1

    requested_output = request.get("output_path")
    with tempfile.TemporaryDirectory(prefix="excelbench-lo-") as tmp:
        tmp_path = Path(tmp)
        out_dir = tmp_path / "out"
        profile_dir = tmp_path / "profile"
        out_dir.mkdir()
        profile_uri = profile_dir.resolve().as_uri()
        convert_to = f"{extension}:{filter_name}"
        command = [
            soffice,
            "--headless",
            "--norestore",
            "--nodefault",
            "--nolockcheck",
            "--nofirststartwizard",
            f"-env:UserInstallation={profile_uri}",
            "--convert-to",
            convert_to,
            "--outdir",
            str(out_dir),
            str(input_path),
        ]
        completed = subprocess.run(
            command,
            text=True,
            capture_output=True,
            check=False,
            timeout=120,
        )
        converted = out_dir / f"{input_path.stem}.{extension}"
        if completed.returncode != 0 or not converted.exists():
            return {
                "error": "libreoffice_conversion_failed",
                "returncode": completed.returncode,
                "stdout": completed.stdout,
                "stderr": completed.stderr,
                "expected_output": str(converted),
            }, 1

        final_output = converted
        if requested_output:
            final_output = Path(str(requested_output))
            final_output.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(converted, final_output)

        return {
            "fixture_id": request.get("fixture_id"),
            "operation": request.get("operation"),
            "tool": "libreoffice",
            "input_path": str(input_path),
            "output_path": str(final_output),
            "bytes": final_output.stat().st_size,
            "stdout": completed.stdout,
            "stderr": completed.stderr,
        }, 0


if __name__ == "__main__":
    raise SystemExit(main())
