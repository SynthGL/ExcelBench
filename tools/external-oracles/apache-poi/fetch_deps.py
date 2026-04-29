#!/usr/bin/env python3
"""Fetch pinned Apache POI helper dependencies from Maven Central."""

from __future__ import annotations

import hashlib
import urllib.error
import urllib.request
from pathlib import Path

MAVEN = "https://repo.maven.apache.org/maven2"
ROOT = Path(__file__).resolve().parent
LIB = ROOT / "deps" / "lib"

ARTIFACTS = (
    ("org.apache.poi", "poi-ooxml", "5.5.1"),
    ("org.apache.poi", "poi", "5.5.1"),
    ("org.apache.poi", "poi-ooxml-lite", "5.5.1"),
    ("org.apache.xmlbeans", "xmlbeans", "5.3.0"),
    ("org.apache.commons", "commons-compress", "1.28.0"),
    ("commons-io", "commons-io", "2.21.0"),
    ("org.apache.commons", "commons-lang3", "3.18.0"),
    ("com.github.virtuald", "curvesapi", "1.08"),
    ("org.apache.logging.log4j", "log4j-api", "2.24.3"),
    ("org.apache.commons", "commons-collections4", "4.5.0"),
    ("commons-codec", "commons-codec", "1.20.0"),
    ("org.apache.commons", "commons-math3", "3.6.1"),
    ("com.zaxxer", "SparseBitSet", "1.3"),
)


def main() -> None:
    """Download and verify pinned runtime jars."""
    LIB.mkdir(parents=True, exist_ok=True)
    for group_id, artifact_id, version in ARTIFACTS:
        jar_url = _artifact_url(group_id, artifact_id, version, "jar")
        output_path = LIB / f"{artifact_id}-{version}.jar"
        algorithm, expected_checksum = _download_checksum(jar_url)
        if output_path.exists() and _checksum(output_path, algorithm) == expected_checksum:
            print(f"OK {output_path.name}")
            continue
        print(f"FETCH {output_path.name}")
        output_path.write_bytes(_download_bytes(jar_url))
        actual_checksum = _checksum(output_path, algorithm)
        if actual_checksum != expected_checksum:
            output_path.unlink(missing_ok=True)
            raise RuntimeError(
                f"Checksum mismatch for {output_path.name}: "
                f"expected {expected_checksum}, got {actual_checksum}"
            )


def _artifact_url(group_id: str, artifact_id: str, version: str, packaging: str) -> str:
    group_path = group_id.replace(".", "/")
    filename = f"{artifact_id}-{version}.{packaging}"
    return f"{MAVEN}/{group_path}/{artifact_id}/{version}/{filename}"


def _download_bytes(url: str) -> bytes:
    with urllib.request.urlopen(url, timeout=60) as response:
        data = response.read()
    return bytes(data)


def _download_text(url: str) -> str:
    return _download_bytes(url).decode("utf-8")


def _download_checksum(jar_url: str) -> tuple[str, str]:
    for extension, algorithm in (("sha512", "sha512"), ("sha256", "sha256"), ("sha1", "sha1")):
        try:
            return algorithm, _download_text(f"{jar_url}.{extension}").split()[0]
        except urllib.error.HTTPError as exc:
            if exc.code != 404:
                raise
    raise RuntimeError(f"No checksum sidecar found for {jar_url}")


def _checksum(path: Path, algorithm: str) -> str:
    hasher = hashlib.new(algorithm)
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


if __name__ == "__main__":
    main()
