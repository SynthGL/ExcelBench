#!/usr/bin/env python3
"""Run an ExcelBench external-oracle helper inside a container on a Docker context.

Drop-in replacement for a local helper command: reads the oracle request JSON
from stdin and prints the helper's stdout unchanged, forwards its stderr, and
exits with its return code. Bind mounts would resolve on the remote host, so
every file crosses the wire inside stdin/stdout envelopes instead:

* ``input_path`` (or an existing ``output_path`` when no ``input_path`` is
  given, mirroring the helpers' read fallback) and every ``payload.images[].path``
  that exists locally are base64-embedded and rewritten to container paths.
* ``output_path`` is rewritten to a container path; the bytes the helper wrote
  there are copied back to the original local ``output_path``.

Usage::

    docker_oracle.py --context pc --image excelbench-poi-oracle:5.5.1 < request.json
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from base64 import b64decode, b64encode
from pathlib import Path
from typing import Any

IN_DIR = "/work/in"
OUT_DIR = "/work/out"
PROTOCOL = 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--context", required=True, help="docker context name, e.g. pc")
    parser.add_argument("--image", required=True)
    parser.add_argument("--docker", default="docker", help="docker CLI executable")
    args = parser.parse_args()

    request: dict[str, Any] = json.loads(sys.stdin.buffer.read())
    files: list[dict[str, str]] = []

    def embed(local: Path, container_path: str) -> str:
        files.append({"path": container_path, "base64": b64encode(local.read_bytes()).decode()})
        return container_path

    local_output: Path | None = None
    output_path = request.get("output_path")
    input_path = request.get("input_path")
    if output_path:
        local_output = Path(output_path)
        container_output = f"{OUT_DIR}/{local_output.name}"
        request["output_path"] = container_output
        if not input_path and local_output.is_file():
            embed(local_output, container_output)
    if input_path and Path(input_path).is_file():
        request["input_path"] = embed(Path(input_path), f"{IN_DIR}/input{Path(input_path).suffix}")

    payload = request.get("payload")
    if isinstance(payload, dict) and isinstance(payload.get("images"), list):
        for index, entry in enumerate(payload["images"]):
            if not isinstance(entry, dict):
                continue
            image_path = entry.get("path")
            if isinstance(image_path, str) and image_path and Path(image_path).is_file():
                entry["path"] = embed(
                    Path(image_path), f"{IN_DIR}/image-{index}{Path(image_path).suffix}"
                )

    command = [
        args.docker,
        "--context",
        args.context,
        "run",
        "--rm",
        "-i",
        "--pull",
        "never",
        "--network",
        "none",
        args.image,
    ]
    envelope_in = json.dumps({"request": request, "files": files}).encode("utf-8")
    completed = subprocess.run(command, input=envelope_in, capture_output=True, check=False)

    envelope = _parse_envelope(completed.stdout)
    if completed.returncode != 0 or envelope is None:
        sys.stderr.buffer.write(completed.stderr)
        sys.stdout.write(
            json.dumps(
                {
                    "error": "docker_oracle_transport_failed",
                    "message": (
                        f"{' '.join(command)} exited {completed.returncode}: "
                        f"{completed.stderr.decode('utf-8', 'replace').strip()}"
                    ),
                },
                sort_keys=True,
            )
            + "\n"
        )
        return completed.returncode or 1

    output_b64 = envelope.get("output_base64")
    if local_output is not None and output_b64 is not None:
        local_output.parent.mkdir(parents=True, exist_ok=True)
        local_output.write_bytes(b64decode(output_b64))
    sys.stderr.buffer.write(completed.stderr)
    sys.stderr.buffer.write(b64decode(envelope["stderr_base64"]))
    sys.stdout.buffer.write(b64decode(envelope["stdout_base64"]))
    sys.stdout.flush()
    return int(envelope["returncode"])


def _parse_envelope(stdout: bytes) -> dict[str, Any] | None:
    try:
        envelope = json.loads(stdout)
    except ValueError:
        return None
    if not isinstance(envelope, dict) or envelope.get("protocol") != PROTOCOL:
        return None
    return envelope


if __name__ == "__main__":
    raise SystemExit(main())
