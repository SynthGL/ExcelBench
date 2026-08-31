"""Deterministic, exact-file manifests for benchmark evidence snapshots."""

from __future__ import annotations

import hashlib
import json
import os
import re
import stat
import tempfile
import unicodedata
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path, PurePosixPath
from typing import Any

SCHEMA_VERSION = 1
DEFAULT_MANIFEST_NAME = "excelbench-evidence.json"
MAX_FILES = 10_000
MAX_FILE_BYTES = 512 * 1024 * 1024
MAX_TOTAL_BYTES = 2 * 1024 * 1024 * 1024
_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_GIT_SHA_RE = re.compile(r"^[0-9a-f]{40}$")
_WINDOWS_RESERVED_BASENAMES = {
    "aux",
    "clock$",
    "con",
    "conin$",
    "conout$",
    "nul",
    "prn",
    *(f"com{suffix}" for suffix in (*range(1, 10), "¹", "²", "³")),
    *(f"lpt{suffix}" for suffix in (*range(1, 10), "¹", "²", "³")),
}
_WINDOWS_FORBIDDEN_CHARACTERS = frozenset('<>:"|?*')


class EvidenceManifestError(ValueError):
    """The evidence snapshot or manifest violates its fail-closed contract."""


@dataclass(frozen=True, order=True)
class EvidenceSubject:
    """An external artifact or source identity bound into a snapshot."""

    name: str
    sha256: str
    version: str | None = None

    def to_dict(self) -> dict[str, str]:
        value = {"name": self.name, "sha256": self.sha256}
        if self.version is not None:
            value["version"] = self.version
        return value


@dataclass(frozen=True, order=True)
class EvidenceArtifact:
    """One immutable regular file in an evidence snapshot."""

    path: str
    sha256: str
    size_bytes: int

    def to_dict(self) -> dict[str, str | int]:
        return {
            "path": self.path,
            "sha256": self.sha256,
            "size_bytes": self.size_bytes,
        }


def canonical_json_bytes(value: object) -> bytes:
    """Return the one wire representation used for hashing and publication."""
    return json.dumps(
        value,
        ensure_ascii=False,
        allow_nan=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def build_evidence_manifest(
    root: Path,
    *,
    snapshot_id: str,
    repository: str,
    source_sha: str,
    observed_at: str,
    subjects: Sequence[EvidenceSubject] = (),
    manifest_name: str = DEFAULT_MANIFEST_NAME,
) -> dict[str, Any]:
    """Inventory ``root`` exactly and return a path-free deterministic manifest.

    ``observed_at`` is explicit rather than read from the clock so rerunning with
    identical inputs produces identical bytes. The output manifest itself is
    excluded when it lives directly below ``root``.
    """
    root = root.resolve(strict=True)
    if not root.is_dir():
        raise EvidenceManifestError("evidence root must be a directory")
    snapshot_id = _required_text(snapshot_id, "snapshot_id")
    repository = _required_text(repository, "repository")
    source_sha = source_sha.strip().lower()
    if _GIT_SHA_RE.fullmatch(source_sha) is None:
        raise EvidenceManifestError("source_sha must be a full lowercase 40-character Git SHA")
    observed_at = _canonical_utc_timestamp(observed_at)
    manifest_name = _safe_manifest_name(manifest_name)
    _read_existing_manifest_destination(root / manifest_name)
    normalized_subjects = _validate_subjects(subjects)
    artifacts = _inventory(root, excluded_root_name=manifest_name)
    artifact_dicts = [artifact.to_dict() for artifact in artifacts]
    artifact_set_sha256 = hashlib.sha256(canonical_json_bytes(artifact_dicts)).hexdigest()
    return {
        "schema": "https://excelbench.dev/schemas/evidence-manifest/v1",
        "schema_version": SCHEMA_VERSION,
        "snapshot_id": snapshot_id,
        "observed_at": observed_at,
        "source": {"repository": repository, "commit": source_sha},
        "subjects": [subject.to_dict() for subject in normalized_subjects],
        "artifacts": artifact_dicts,
        "artifact_count": len(artifacts),
        "total_size_bytes": sum(artifact.size_bytes for artifact in artifacts),
        "artifact_set_sha256": artifact_set_sha256,
    }


def verify_evidence_manifest(
    root: Path,
    manifest: Mapping[str, Any],
    *,
    expected_source_sha: str | None = None,
    manifest_name: str = DEFAULT_MANIFEST_NAME,
) -> None:
    """Fail unless ``manifest`` exactly describes every regular file in ``root``."""
    expected = _validate_manifest_document(manifest)
    source = _mapping(manifest["source"], "source")
    source_sha = _string(source.get("commit"), "source.commit")
    if expected_source_sha is not None and source_sha != expected_source_sha:
        raise EvidenceManifestError("manifest source commit does not match expected source SHA")
    root = root.resolve(strict=True)
    manifest_name = _safe_manifest_name(manifest_name)
    destination_manifest = _read_existing_manifest_destination(root / manifest_name)
    if (
        destination_manifest is not None
        and canonical_json_bytes(destination_manifest) != canonical_json_bytes(dict(manifest))
    ):
        raise EvidenceManifestError("manifest file does not match the manifest being verified")
    actual = _inventory(
        root,
        excluded_root_name=manifest_name,
        require_artifact=False,
    )
    if actual != expected:
        actual_by_path = {artifact.path: artifact for artifact in actual}
        expected_by_path = {artifact.path: artifact for artifact in expected}
        missing = sorted(expected_by_path.keys() - actual_by_path.keys())
        extra = sorted(actual_by_path.keys() - expected_by_path.keys())
        changed = sorted(
            path
            for path in actual_by_path.keys() & expected_by_path.keys()
            if actual_by_path[path] != expected_by_path[path]
        )
        raise EvidenceManifestError(
            f"evidence mismatch: missing={missing!r}, extra={extra!r}, changed={changed!r}"
        )


def read_evidence_manifest(path: Path) -> dict[str, Any]:
    """Read a bounded UTF-8 manifest and reject duplicate JSON keys."""
    if path.is_symlink() or not path.is_file():
        raise EvidenceManifestError("manifest must be a regular non-symlink file")
    if path.stat().st_size > 4 * 1024 * 1024:
        raise EvidenceManifestError("manifest exceeds the 4 MiB size limit")

    def no_duplicates(pairs: Iterable[tuple[str, Any]]) -> dict[str, Any]:
        value: dict[str, Any] = {}
        for key, item in pairs:
            if key in value:
                raise EvidenceManifestError(f"manifest contains duplicate key {key!r}")
            value[key] = item
        return value

    try:
        decoded = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=no_duplicates)
    except EvidenceManifestError:
        raise
    except (OSError, UnicodeDecodeError, ValueError) as exc:
        raise EvidenceManifestError("manifest is not valid UTF-8 JSON") from exc
    if not isinstance(decoded, dict):
        raise EvidenceManifestError("manifest root must be an object")
    return decoded


def write_evidence_manifest(
    path: Path, manifest: Mapping[str, Any], *, replace: bool = False
) -> None:
    """Publish one canonical manifest without exposing a partially written file."""
    if path.is_symlink():
        raise EvidenceManifestError("manifest destination must not be a symlink")
    name = _safe_manifest_name(path.name)
    path.parent.mkdir(parents=True, exist_ok=True)
    path = path.parent.resolve(strict=True) / name
    if path.is_symlink():
        raise EvidenceManifestError("manifest destination must not be a symlink")
    if replace:
        _read_existing_manifest_destination(path)
    _validate_manifest_document(manifest)
    contents = canonical_json_bytes(dict(manifest)) + b"\n"
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as output:
            output.write(contents)
            output.flush()
            os.fsync(output.fileno())
        if replace:
            _read_existing_manifest_destination(path)
            os.replace(temporary, path)
        else:
            try:
                os.link(temporary, path)
            except FileExistsError as exc:
                raise EvidenceManifestError("manifest already exists; pass replace=True") from exc
            temporary.unlink()
        if os.name != "nt":
            directory = os.open(path.parent, os.O_RDONLY)
            try:
                os.fsync(directory)
            finally:
                os.close(directory)
    finally:
        temporary.unlink(missing_ok=True)


def parse_subject(value: str) -> EvidenceSubject:
    """Parse ``NAME=SHA256`` or ``NAME@VERSION=SHA256`` CLI syntax."""
    identity, separator, digest = value.partition("=")
    if not separator:
        raise EvidenceManifestError("subject must use NAME=SHA256 or NAME@VERSION=SHA256")
    name, version_separator, version = identity.partition("@")
    return EvidenceSubject(
        name=_required_text(name, "subject.name"),
        version=_required_text(version, "subject.version") if version_separator else None,
        sha256=digest.strip().lower(),
    )


def _inventory(
    root: Path, *, excluded_root_name: str, require_artifact: bool = True
) -> list[EvidenceArtifact]:
    artifacts: list[EvidenceArtifact] = []
    entries = _inventory_entries(root, excluded_root_name=excluded_root_name)
    hashed_signatures: dict[str, tuple[int, int, int, int, int]] = {}
    total_size = 0
    for relative, path, _initial_signature in entries:
        digest, size, signature = _sha256_stable_file(path)
        hashed_signatures[relative] = signature
        total_size += size
        if total_size > MAX_TOTAL_BYTES:
            raise EvidenceManifestError("evidence snapshot exceeds total size limit")
        artifacts.append(EvidenceArtifact(path=relative, sha256=digest, size_bytes=size))
        if len(artifacts) > MAX_FILES:
            raise EvidenceManifestError("evidence snapshot exceeds file-count limit")

    # Hashing can be long-running. Re-scan the complete namespace so a producer
    # cannot add, remove, replace, or mutate an artifact after the first walk
    # and still receive a successful exact-snapshot manifest.
    final_entries = _inventory_entries(root, excluded_root_name=excluded_root_name)
    final_signatures = {
        relative: signature for relative, _path, signature in final_entries
    }
    if final_signatures != hashed_signatures:
        raise EvidenceManifestError(
            "evidence file set or metadata changed while inventorying"
        )
    if require_artifact and not artifacts:
        raise EvidenceManifestError("evidence snapshot must contain at least one artifact")
    return artifacts


def _inventory_entries(
    root: Path, *, excluded_root_name: str
) -> list[tuple[str, Path, tuple[int, int, int, int, int]]]:
    entries: list[tuple[str, Path, tuple[int, int, int, int, int]]] = []
    # The manifest is excluded from its own inventory, but its portable name is
    # still reserved so a differently-cased artifact cannot alias it on Windows.
    seen_casefolded: set[str] = {_portable_path_key(excluded_root_name)}
    for path in sorted(root.rglob("*"), key=lambda item: item.as_posix()):
        if path.parent == root and path.name == excluded_root_name:
            continue
        try:
            metadata = os.stat(path, follow_symlinks=False)
        except OSError as exc:
            raise EvidenceManifestError(
                f"evidence entry changed while inventorying: {path.relative_to(root)}"
            ) from exc
        if stat.S_ISLNK(metadata.st_mode):
            raise EvidenceManifestError(
                f"symlinks are not allowed in evidence: {path.relative_to(root)}"
            )
        if stat.S_ISDIR(metadata.st_mode):
            continue
        if not stat.S_ISREG(metadata.st_mode):
            raise EvidenceManifestError(f"non-regular evidence entry: {path.relative_to(root)}")
        relative = _canonical_relative_path(path.relative_to(root))
        portable_key = _portable_path_key(relative)
        if portable_key in seen_casefolded:
            raise EvidenceManifestError(f"case-insensitive evidence path collision: {relative}")
        seen_casefolded.add(portable_key)
        entries.append((relative, path, _file_signature(metadata)))
        if len(entries) > MAX_FILES:
            raise EvidenceManifestError("evidence snapshot exceeds file-count limit")
    return entries


def _canonical_relative_path(path: Path) -> str:
    raw = path.as_posix()
    try:
        raw.encode("utf-8", errors="strict")
    except UnicodeEncodeError as exc:
        raise EvidenceManifestError(
            f"evidence path is not valid UTF-8: {raw!r}"
        ) from exc
    if "\\" in raw:
        raise EvidenceManifestError(f"evidence path contains a backslash: {raw!r}")
    normalized = unicodedata.normalize("NFC", raw)
    if raw != normalized:
        raise EvidenceManifestError(f"evidence path is not Unicode NFC-normalized: {raw!r}")
    pure = PurePosixPath(normalized)
    if pure.is_absolute() or ".." in pure.parts or not pure.parts:
        raise EvidenceManifestError(f"unsafe evidence path: {raw!r}")
    canonical = pure.as_posix()
    _portable_path_key(canonical)
    return canonical


def _portable_path_key(path: str) -> str:
    key_parts: list[str] = []
    for segment in PurePosixPath(path).parts:
        if segment.endswith((".", " ")):
            raise EvidenceManifestError(
                f"evidence path is not portable to Windows: {path!r}"
            )
        if any(
            ord(character) < 32 or character in _WINDOWS_FORBIDDEN_CHARACTERS
            for character in segment
        ):
            raise EvidenceManifestError(
                f"evidence path contains a Windows-forbidden character: {path!r}"
            )
        basename = segment.split(".", maxsplit=1)[0].casefold()
        if basename in _WINDOWS_RESERVED_BASENAMES:
            raise EvidenceManifestError(
                f"evidence path uses a Windows-reserved basename: {path!r}"
            )
        key_parts.append(segment.casefold())
    return "/".join(key_parts)


def _file_signature(metadata: os.stat_result) -> tuple[int, int, int, int, int]:
    return (
        metadata.st_dev,
        metadata.st_ino,
        metadata.st_size,
        metadata.st_mtime_ns,
        metadata.st_ctime_ns,
    )


def _sha256_stable_file(
    path: Path,
) -> tuple[str, int, tuple[int, int, int, int, int]]:
    flags = os.O_RDONLY | getattr(os, "O_BINARY", 0)
    flags |= getattr(os, "O_NOFOLLOW", 0)
    try:
        descriptor = os.open(path, flags)
    except OSError as exc:
        raise EvidenceManifestError(f"cannot open evidence file safely: {path}") from exc

    digest = hashlib.sha256()
    with os.fdopen(descriptor, "rb") as source:
        before = os.fstat(source.fileno())
        if not stat.S_ISREG(before.st_mode):
            raise EvidenceManifestError(f"evidence entry is not a regular file: {path}")
        if before.st_size > MAX_FILE_BYTES:
            raise EvidenceManifestError(f"evidence file exceeds size limit: {path}")

        bytes_read = 0
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            bytes_read += len(chunk)
            if bytes_read > MAX_FILE_BYTES:
                raise EvidenceManifestError(f"evidence file exceeds size limit: {path}")
            digest.update(chunk)
        after = os.fstat(source.fileno())

    try:
        current = os.stat(path, follow_symlinks=False)
    except OSError as exc:
        raise EvidenceManifestError(f"evidence file changed while hashing: {path}") from exc
    if (
        stat.S_ISLNK(current.st_mode)
        or bytes_read != before.st_size
        or _file_signature(before) != _file_signature(after)
        or _file_signature(after) != _file_signature(current)
    ):
        raise EvidenceManifestError(f"evidence file changed while hashing: {path}")
    return digest.hexdigest(), before.st_size, _file_signature(current)


def _canonical_utc_timestamp(value: str) -> str:
    text = _required_text(value, "observed_at")
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError as exc:
        raise EvidenceManifestError("observed_at must be an ISO-8601 timestamp") from exc
    if parsed.tzinfo is None:
        raise EvidenceManifestError("observed_at must include a timezone")
    return parsed.astimezone(UTC).isoformat().replace("+00:00", "Z")


def _validate_subjects(subjects: Sequence[EvidenceSubject]) -> list[EvidenceSubject]:
    normalized: list[EvidenceSubject] = []
    identities: set[tuple[str, str | None]] = set()
    for subject in subjects:
        name = _required_text(subject.name, "subject.name")
        version = (
            _required_text(subject.version, "subject.version")
            if subject.version is not None
            else None
        )
        digest = subject.sha256.strip().lower()
        if _SHA256_RE.fullmatch(digest) is None:
            raise EvidenceManifestError("subject.sha256 must be a lowercase SHA-256 digest")
        identity = (name.casefold(), version)
        if identity in identities:
            raise EvidenceManifestError(f"duplicate evidence subject: {name!r}")
        identities.add(identity)
        normalized.append(EvidenceSubject(name=name, version=version, sha256=digest))
    return sorted(normalized, key=lambda item: (item.name.casefold(), item.version or ""))


def _validate_manifest_shape(manifest: Mapping[str, Any]) -> None:
    required = {
        "schema",
        "schema_version",
        "snapshot_id",
        "observed_at",
        "source",
        "subjects",
        "artifacts",
        "artifact_count",
        "total_size_bytes",
        "artifact_set_sha256",
    }
    if set(manifest) != required:
        raise EvidenceManifestError(
            f"manifest fields must be exactly {sorted(required)!r}"
        )
    schema_version = manifest["schema_version"]
    if (
        not isinstance(schema_version, int)
        or isinstance(schema_version, bool)
        or schema_version != SCHEMA_VERSION
    ):
        raise EvidenceManifestError(f"unsupported schema_version: {schema_version!r}")
    if manifest["schema"] != "https://excelbench.dev/schemas/evidence-manifest/v1":
        raise EvidenceManifestError("unsupported manifest schema URI")
    snapshot_id = _string(manifest["snapshot_id"], "snapshot_id")
    if snapshot_id != _required_text(snapshot_id, "snapshot_id"):
        raise EvidenceManifestError("snapshot_id must use its canonical text form")
    observed_at = _string(manifest["observed_at"], "observed_at")
    if observed_at != _canonical_utc_timestamp(observed_at):
        raise EvidenceManifestError("observed_at must use canonical UTC form")
    source = _mapping(manifest["source"], "source")
    if set(source) != {"repository", "commit"}:
        raise EvidenceManifestError("source must contain exactly repository and commit")
    repository = _string(source["repository"], "source.repository")
    if repository != _required_text(repository, "source.repository"):
        raise EvidenceManifestError("source.repository must use its canonical text form")
    commit = _string(source["commit"], "source.commit")
    if _GIT_SHA_RE.fullmatch(commit) is None:
        raise EvidenceManifestError("source.commit must be a full lowercase Git SHA")
    if not isinstance(manifest["subjects"], list) or not isinstance(manifest["artifacts"], list):
        raise EvidenceManifestError("subjects and artifacts must be arrays")
    artifact_count = manifest["artifact_count"]
    if (
        not isinstance(artifact_count, int)
        or isinstance(artifact_count, bool)
        or artifact_count < 1
    ):
        raise EvidenceManifestError("artifact_count must be a positive integer")
    total_size_bytes = manifest["total_size_bytes"]
    if (
        not isinstance(total_size_bytes, int)
        or isinstance(total_size_bytes, bool)
        or total_size_bytes < 0
    ):
        raise EvidenceManifestError("total_size_bytes must be a non-negative integer")
    artifact_set_sha256 = _string(
        manifest["artifact_set_sha256"], "artifact_set_sha256"
    )
    if _SHA256_RE.fullmatch(artifact_set_sha256) is None:
        raise EvidenceManifestError("artifact_set_sha256 must be a lowercase SHA-256 digest")


def _validate_manifest_document(manifest: Mapping[str, Any]) -> list[EvidenceArtifact]:
    _validate_manifest_shape(manifest)
    subjects = [
        _subject_from_mapping(value, index)
        for index, value in enumerate(manifest["subjects"])
    ]
    if subjects != _validate_subjects(subjects):
        raise EvidenceManifestError("subjects must be sorted by canonical identity")
    artifacts = [
        _artifact_from_mapping(value, index)
        for index, value in enumerate(manifest["artifacts"])
    ]
    if artifacts != sorted(artifacts):
        raise EvidenceManifestError("artifacts must be sorted by canonical path")
    if len({_portable_path_key(artifact.path) for artifact in artifacts}) != len(artifacts):
        raise EvidenceManifestError("manifest contains case-insensitive path collisions")
    artifact_dicts = [artifact.to_dict() for artifact in artifacts]
    digest = hashlib.sha256(canonical_json_bytes(artifact_dicts)).hexdigest()
    if manifest["artifact_set_sha256"] != digest:
        raise EvidenceManifestError("artifact_set_sha256 does not match the artifact inventory")
    if manifest["artifact_count"] != len(artifacts):
        raise EvidenceManifestError("artifact_count does not match the artifact inventory")
    if manifest["total_size_bytes"] != sum(
        artifact.size_bytes for artifact in artifacts
    ):
        raise EvidenceManifestError("total_size_bytes does not match the artifact inventory")
    return artifacts


def _read_existing_manifest_destination(path: Path) -> dict[str, Any] | None:
    if path.is_symlink():
        raise EvidenceManifestError("manifest destination must not be a symlink")
    if not path.exists():
        return None
    if not path.is_file():
        raise EvidenceManifestError("manifest destination must be a regular file")
    try:
        manifest = read_evidence_manifest(path)
        _validate_manifest_document(manifest)
    except EvidenceManifestError as exc:
        raise EvidenceManifestError(
            "manifest destination exists but is not a valid evidence manifest"
        ) from exc
    return manifest


def _artifact_from_mapping(value: Any, index: int) -> EvidenceArtifact:
    item = _mapping(value, f"artifacts[{index}]")
    if set(item) != {"path", "sha256", "size_bytes"}:
        raise EvidenceManifestError(f"artifacts[{index}] has unexpected fields")
    path = _string(item["path"], f"artifacts[{index}].path")
    if _canonical_relative_path(Path(path)) != path:
        raise EvidenceManifestError(f"artifacts[{index}].path is not canonical")
    digest = _string(item["sha256"], f"artifacts[{index}].sha256")
    if _SHA256_RE.fullmatch(digest) is None:
        raise EvidenceManifestError(f"artifacts[{index}].sha256 is invalid")
    size = item["size_bytes"]
    if not isinstance(size, int) or isinstance(size, bool) or size < 0 or size > MAX_FILE_BYTES:
        raise EvidenceManifestError(f"artifacts[{index}].size_bytes is invalid")
    return EvidenceArtifact(path=path, sha256=digest, size_bytes=size)


def _subject_from_mapping(value: Any, index: int) -> EvidenceSubject:
    item = _mapping(value, f"subjects[{index}]")
    if set(item) not in ({"name", "sha256"}, {"name", "sha256", "version"}):
        raise EvidenceManifestError(f"subjects[{index}] has unexpected fields")
    return EvidenceSubject(
        name=_string(item["name"], f"subjects[{index}].name"),
        sha256=_string(item["sha256"], f"subjects[{index}].sha256"),
        version=(
            _string(item["version"], f"subjects[{index}].version")
            if "version" in item
            else None
        ),
    )


def _safe_manifest_name(value: str) -> str:
    name = _required_text(value, "manifest_name")
    if Path(name).name != name or name in {".", ".."}:
        raise EvidenceManifestError("manifest_name must be one filename")
    if _canonical_relative_path(Path(name)) != name:
        raise EvidenceManifestError("manifest_name must use its canonical portable form")
    return name


def _required_text(value: str, label: str) -> str:
    stripped = value.strip()
    if not stripped or any(character in stripped for character in "\r\n\x00"):
        raise EvidenceManifestError(f"{label} must be non-empty single-line text")
    return stripped


def _mapping(value: Any, label: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping) or not all(isinstance(key, str) for key in value):
        raise EvidenceManifestError(f"{label} must be an object with string keys")
    return value


def _string(value: Any, label: str) -> str:
    if not isinstance(value, str):
        raise EvidenceManifestError(f"{label} must be a string")
    return value
