#!/usr/bin/env python3
"""Generate local external-oracle fixtures for pre-release hardening."""

from __future__ import annotations

import argparse
from pathlib import Path

from excelbench.harness.external_fixture_pack import generate_external_fixture_pack


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("results_dev_external/fixtures"),
        help="Output directory for generated fixtures and manifest.",
    )
    parser.add_argument(
        "--no-validators",
        action="store_true",
        help="Skip LibreOffice validation helpers.",
    )
    args = parser.parse_args()

    repo_root = Path(__file__).resolve().parents[1]
    results = generate_external_fixture_pack(
        args.output,
        repo_root=repo_root,
        include_validators=not args.no_validators,
    )
    for result in results:
        status = "PASS" if result.passed else "FAIL"
        print(f"{status} {result.fixture_id}: {result.workbook_path}")
        if result.missing_parts:
            print(f"  missing parts: {', '.join(result.missing_parts)}")
        for validation in result.validations:
            if validation.skipped:
                print(f"  {validation.tool_name}: SKIP {validation.notes or ''}".rstrip())
            else:
                validation_status = "PASS" if validation.passed else "FAIL"
                operation = validation.payload.get("operation")
                print(f"  {validation.tool_name}/{operation}: {validation_status}")
    return 0 if all(result.passed for result in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
