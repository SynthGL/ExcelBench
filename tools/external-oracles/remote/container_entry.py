#!/usr/bin/env python3
"""In-container side of the remote Docker oracle transport.

Reads one envelope from stdin, materializes the embedded files inside the
container, runs the baked-in helper command with the rewritten oracle request
on its stdin, and writes one envelope to stdout carrying the helper's exact
stdout/stderr bytes, its return code, and the produced output workbook bytes.

Envelope in::

    {"request": {...}, "files": [{"path": "/work/in/x.xlsx", "base64": "..."}]}

Envelope out::

    {"protocol": 1, "returncode": 0, "stdout_base64": "...",
     "stderr_base64": "...", "output_base64": "..." | null}

Usage (as the image ENTRYPOINT)::

    container_entry.py --cwd /opt/helper [--timeout 600] -- helper-cmd [args...]
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from base64 import b64decode, b64encode
from pathlib import Path

PROTOCOL = 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--cwd", required=True)
    parser.add_argument("--timeout", type=float, default=600.0)
    parser.add_argument("command", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    command = args.command[1:] if args.command[:1] == ["--"] else args.command
    if not command:
        parser.error("missing helper command after --")

    envelope = json.loads(sys.stdin.buffer.read())
    request = envelope["request"]
    for item in envelope.get("files", []):
        path = Path(item["path"])
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b64decode(item["base64"]))
    output_path = request.get("output_path")
    if output_path:
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    try:
        completed = subprocess.run(
            command,
            input=json.dumps(request, sort_keys=True).encode("utf-8"),
            capture_output=True,
            check=False,
            timeout=args.timeout,
            cwd=args.cwd,
        )
        returncode = completed.returncode
        stdout = completed.stdout
        stderr = completed.stderr
    except subprocess.TimeoutExpired as exc:
        returncode = 124
        stdout = exc.stdout or b""
        stderr = (exc.stderr or b"") + (
            f"container helper timed out after {args.timeout:g}s\n".encode()
        )

    output_bytes = None
    if output_path and Path(output_path).is_file():
        output_bytes = Path(output_path).read_bytes()

    sys.stdout.write(
        json.dumps(
            {
                "protocol": PROTOCOL,
                "returncode": returncode,
                "stdout_base64": b64encode(stdout).decode("ascii"),
                "stderr_base64": b64encode(stderr).decode("ascii"),
                "output_base64": (
                    None if output_bytes is None else b64encode(output_bytes).decode("ascii")
                ),
            }
        )
    )
    sys.stdout.flush()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
