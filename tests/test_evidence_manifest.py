import json
from pathlib import Path
from typing import Any

import pytest
from typer.testing import CliRunner

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


def test_read_rejects_duplicate_json_keys(tmp_path: Path) -> None:
    path = tmp_path / "manifest.json"
    path.write_text('{"schema_version":1,"schema_version":1}')

    with pytest.raises(EvidenceManifestError, match="duplicate key"):
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
