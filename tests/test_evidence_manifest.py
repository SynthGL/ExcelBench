import json
import os
from pathlib import Path
from typing import Any

import pytest
from typer.testing import CliRunner

import excelbench.evidence_manifest as evidence_manifest
from excelbench.cli import app
from excelbench.evidence_manifest import (
    EvidenceManifestError,
    EvidenceSubject,
    build_evidence_manifest,
    canonical_json_bytes,
    parse_subject,
    read_evidence_manifest,
    verify_evidence_manifest,
    write_evidence_manifest,
)

SOURCE_SHA = "a" * 40
OBSERVED_AT = "2026-08-31T00:00:00Z"
RUNNER = CliRunner()


def _manifest(root: Path) -> dict[str, Any]:
    return build_evidence_manifest(
        root,
        snapshot_id="wolfxl-2.1-linux-x86_64",
        repository="SynthGL/ExcelBench",
        source_sha=SOURCE_SHA,
        observed_at=OBSERVED_AT,
        subjects=[EvidenceSubject("wolfxl-wheel", "b" * 64, "2.1.0")],
    )


def test_manifest_is_deterministic_path_free_and_exact(tmp_path: Path) -> None:
    root = tmp_path / "results"
    (root / "nested").mkdir(parents=True)
    (root / "nested" / "matrix.csv").write_text("feature,score\ncell,3\n")
    (root / "results.json").write_text('{"passed":true}\n')

    first = _manifest(root)
    second = _manifest(root)

    assert canonical_json_bytes(first) == canonical_json_bytes(second)
    serialized = canonical_json_bytes(first).decode()
    assert str(tmp_path) not in serialized
    assert [item["path"] for item in first["artifacts"]] == [
        "nested/matrix.csv",
        "results.json",
    ]
    assert first["artifact_count"] == 2
    verify_evidence_manifest(root, first, expected_source_sha=SOURCE_SHA)


def test_verification_rejects_missing_extra_and_changed_files(tmp_path: Path) -> None:
    root = tmp_path / "results"
    root.mkdir()
    artifact = root / "results.json"
    artifact.write_text("before")
    manifest = _manifest(root)

    artifact.write_text("after")
    with pytest.raises(EvidenceManifestError, match="changed=.*results.json"):
        verify_evidence_manifest(root, manifest)

    artifact.write_text("before")
    (root / "extra.txt").write_text("unexpected")
    with pytest.raises(EvidenceManifestError, match="extra=.*extra.txt"):
        verify_evidence_manifest(root, manifest)

    (root / "extra.txt").unlink()
    artifact.unlink()
    with pytest.raises(EvidenceManifestError, match="missing=.*results.json"):
        verify_evidence_manifest(root, manifest)


def test_manifest_file_is_excluded_and_atomic_no_clobber_is_default(tmp_path: Path) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "results.json").write_text("{}")
    manifest = _manifest(root)
    path = root / "excelbench-evidence.json"

    write_evidence_manifest(path, manifest)
    assert read_evidence_manifest(path) == manifest
    verify_evidence_manifest(root, read_evidence_manifest(path))

    with pytest.raises(EvidenceManifestError, match="already exists"):
        write_evidence_manifest(path, manifest)


def test_manifest_destination_rejects_symlinks_and_unrelated_artifacts(
    tmp_path: Path,
) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "artifact.txt").write_text("evidence")
    external = tmp_path / "external.json"
    external.write_text("do not replace")
    destination = root / "excelbench-evidence.json"
    try:
        destination.symlink_to(external)
    except OSError:
        pytest.skip("symlinks unavailable")

    with pytest.raises(EvidenceManifestError, match="must not be a symlink"):
        _manifest(root)
    with pytest.raises(EvidenceManifestError, match="must not be a symlink"):
        write_evidence_manifest(destination, {"irrelevant": True}, replace=True)
    assert external.read_text() == "do not replace"

    destination.unlink()
    destination.write_text("benchmark output")
    with pytest.raises(EvidenceManifestError, match="not a valid evidence manifest"):
        _manifest(root)
    with pytest.raises(EvidenceManifestError, match="not a valid evidence manifest"):
        write_evidence_manifest(destination, {"irrelevant": True}, replace=True)
    assert destination.read_text() == "benchmark output"


def test_read_rejects_duplicate_json_keys(tmp_path: Path) -> None:
    path = tmp_path / "manifest.json"
    path.write_text('{"schema_version":1,"schema_version":1}')

    with pytest.raises(EvidenceManifestError, match="duplicate key"):
        read_evidence_manifest(path)


def test_read_wraps_oversized_json_integer_errors(tmp_path: Path) -> None:
    path = tmp_path / "manifest.json"
    path.write_text('{"artifact_count":' + "9" * 5_000 + "}")

    with pytest.raises(EvidenceManifestError, match="not valid UTF-8 JSON"):
        read_evidence_manifest(path)


def test_symlinks_and_case_collisions_fail_closed(tmp_path: Path) -> None:
    root = tmp_path / "results"
    root.mkdir()
    target = root / "target.json"
    target.write_text("{}")
    link = root / "link.json"
    try:
        link.symlink_to(target)
    except OSError:
        pytest.skip("symlinks unavailable")

    with pytest.raises(EvidenceManifestError, match="symlinks are not allowed"):
        _manifest(root)

    link.unlink()
    (root / "A.json").write_text("a")
    (root / "a.json").write_text("b")
    with pytest.raises(EvidenceManifestError, match="case-insensitive"):
        _manifest(root)


def test_windows_aliases_and_reserved_names_fail_closed(tmp_path: Path) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "artifact.json").write_text("{}")
    trailing_dot = root / "report."
    try:
        trailing_dot.write_text("alias")
    except OSError:
        pytest.skip("host filesystem rejects Windows-aliased names")

    with pytest.raises(EvidenceManifestError, match="not portable to Windows"):
        _manifest(root)

    trailing_dot.unlink()
    reserved = root / "CON.txt"
    try:
        reserved.write_text("reserved")
    except OSError:
        pytest.skip("host filesystem rejects Windows-reserved names")
    with pytest.raises(EvidenceManifestError, match="Windows-reserved basename"):
        _manifest(root)


@pytest.mark.parametrize("manifest_name", ["CON.json", "report.", r"nested\\manifest.json"])
def test_manifest_name_must_be_portable(
    tmp_path: Path, manifest_name: str
) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "artifact.json").write_text("{}")

    with pytest.raises(EvidenceManifestError):
        build_evidence_manifest(
            root,
            snapshot_id="snapshot",
            repository="SynthGL/ExcelBench",
            source_sha=SOURCE_SHA,
            observed_at=OBSERVED_AT,
            manifest_name=manifest_name,
        )


def test_manifest_name_portable_key_is_reserved(tmp_path: Path) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "artifact.json").write_text("{}")
    (root / "ExcelBench-Evidence.json").write_text("collision")

    with pytest.raises(EvidenceManifestError, match="case-insensitive"):
        _manifest(root)


def test_artifact_metadata_changes_during_hashing_fail_closed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "artifact.json").write_text("{}")
    real_signature = evidence_manifest._file_signature
    calls = 0

    def changing_signature(metadata: Any) -> tuple[int, int, int, int, int]:
        nonlocal calls
        calls += 1
        signature = real_signature(metadata)
        if calls == 2:
            return (*signature[:-1], signature[-1] + 1)
        return signature

    monkeypatch.setattr(evidence_manifest, "_file_signature", changing_signature)
    with pytest.raises(EvidenceManifestError, match="changed while hashing"):
        _manifest(root)


def test_artifact_created_during_hashing_fails_closed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "artifact.json").write_text("{}")
    real_hash = evidence_manifest._sha256_stable_file
    created = False

    def create_late_artifact(path: Path) -> tuple[str, int, tuple[int, int, int, int, int]]:
        nonlocal created
        result = real_hash(path)
        if not created:
            (root / "late.json").write_text("late")
            created = True
        return result

    monkeypatch.setattr(evidence_manifest, "_sha256_stable_file", create_late_artifact)
    with pytest.raises(EvidenceManifestError, match="changed while inventorying"):
        _manifest(root)


@pytest.mark.skipif(os.name == "nt", reason="requires POSIX byte filenames")
def test_non_utf8_artifact_name_fails_closed(tmp_path: Path) -> None:
    root = tmp_path / "results"
    root.mkdir()
    byte_path = os.path.join(os.fsencode(root), b"invalid-\xff.json")
    descriptor = os.open(byte_path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    try:
        os.write(descriptor, b"{}")
    finally:
        os.close(descriptor)

    with pytest.raises(EvidenceManifestError, match="not valid UTF-8"):
        _manifest(root)


def test_timestamp_source_and_subject_contracts_are_strict(tmp_path: Path) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "results.json").write_text("{}")

    with pytest.raises(EvidenceManifestError, match="40-character Git SHA"):
        build_evidence_manifest(
            root,
            snapshot_id="snapshot",
            repository="SynthGL/ExcelBench",
            source_sha="main",
            observed_at=OBSERVED_AT,
        )
    with pytest.raises(EvidenceManifestError, match="include a timezone"):
        build_evidence_manifest(
            root,
            snapshot_id="snapshot",
            repository="SynthGL/ExcelBench",
            source_sha=SOURCE_SHA,
            observed_at="2026-08-31T00:00:00",
        )
    assert parse_subject(f"wolfxl@2.1.0={'b' * 64}") == EvidenceSubject(
        "wolfxl", "b" * 64, "2.1.0"
    )
    with pytest.raises(EvidenceManifestError, match="lowercase SHA-256"):
        build_evidence_manifest(
            root,
            snapshot_id="snapshot",
            repository="SynthGL/ExcelBench",
            source_sha=SOURCE_SHA,
            observed_at=OBSERVED_AT,
            subjects=[EvidenceSubject("wolfxl", "invalid")],
        )


def test_tampered_aggregate_fields_are_rejected(tmp_path: Path) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "results.json").write_text("{}")
    manifest = _manifest(root)
    tampered = json.loads(json.dumps(manifest))
    tampered["artifact_count"] = 99

    with pytest.raises(EvidenceManifestError, match="artifact_count"):
        verify_evidence_manifest(root, tampered)


def test_boolean_schema_version_is_rejected(tmp_path: Path) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "results.json").write_text("{}")
    manifest = _manifest(root)
    manifest["schema_version"] = True

    with pytest.raises(EvidenceManifestError, match="unsupported schema_version"):
        verify_evidence_manifest(root, manifest)


def test_cli_builds_and_verifies_exact_snapshot(tmp_path: Path) -> None:
    root = tmp_path / "results"
    root.mkdir()
    (root / "results.json").write_text("{}")

    built = RUNNER.invoke(
        app,
        [
            "evidence-manifest",
            "--root",
            str(root),
            "--snapshot-id",
            "release-linux",
            "--source-sha",
            SOURCE_SHA,
            "--observed-at",
            OBSERVED_AT,
            "--subject",
            f"wolfxl@2.1.0={'b' * 64}",
        ],
    )
    assert built.exit_code == 0, built.output
    assert (root / "excelbench-evidence.json").exists()

    verified = RUNNER.invoke(
        app,
        [
            "verify-evidence",
            "--root",
            str(root),
            "--expected-source-sha",
            SOURCE_SHA,
        ],
    )
    assert verified.exit_code == 0, verified.output
