#!/usr/bin/env python3
"""Validate local external-oracle fixtures with WolfXL."""

from __future__ import annotations

import argparse
from pathlib import Path

from excelbench.harness.external_wolfxl_validation import (
    validate_wolfxl_external_fixture_pack,
)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--fixtures",
        type=Path,
        default=Path("results_dev_external/fixtures"),
        help="Directory containing external oracle fixtures and manifest.json.",
    )
    args = parser.parse_args()

    results = validate_wolfxl_external_fixture_pack(args.fixtures)
    for result in results:
        status = "PASS" if result.passed else "FAIL"
        print(f"{status} {result.fixture_id}: {result.modified_workbook}")
        if result.missing_parts_after_save:
            print(f"  missing parts: {', '.join(result.missing_parts_after_save)}")
        if result.readback_failures:
            print(f"  readback failures: {'; '.join(result.readback_failures)}")
        if result.error:
            print(f"  error: {result.error}")
    return 0 if all(result.passed for result in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
