#!/usr/bin/env python3
"""Assert dual-read migration policy knobs reject invalid values."""

from __future__ import annotations

import argparse
import pathlib
import subprocess
import sys


DEFAULT_ARTIFACTS = (
    "snapshot,dispatch,ack_retention,dashboard_pack,policy,escalation,adapter_exports"
)
DEFAULT_FAULT_SCENARIOS = "malformed_fallback_schema,partial_wave_rollback"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Run CI dry-run checks that intentionally pass invalid migration-drill "
            "policy values and assert expected command failures."
        )
    )
    parser.add_argument(
        "--migration-script",
        default="./python/validate_batch_dual_read_migration.py",
        help="Path to dual-read migration drill script.",
    )
    parser.add_argument(
        "--artifacts",
        default=DEFAULT_ARTIFACTS,
        help="Comma-separated artifact policy value to use for dry-run checks.",
    )
    parser.add_argument(
        "--fault-scenarios",
        default=DEFAULT_FAULT_SCENARIOS,
        help="Comma-separated fault-scenario policy value to use for dry-run checks.",
    )
    return parser.parse_args()


def _parse_csv(raw: str, flag: str) -> list[str]:
    values = [token.strip() for token in raw.split(",") if token.strip()]
    if not values:
        raise ValueError(f"{flag} must include at least one value")
    deduped: list[str] = []
    seen: set[str] = set()
    for item in values:
        if item in seen:
            continue
        deduped.append(item)
        seen.add(item)
    return deduped


def _run_expected_failure(
    *,
    case_id: str,
    command: list[str],
    expected_token: str,
) -> None:
    run = subprocess.run(  # noqa: S603
        command,
        capture_output=True,
        text=True,
        check=False,
    )
    output = (run.stdout or "") + (run.stderr or "")
    if run.returncode == 0:
        raise AssertionError(
            f"dry-run case {case_id!r} unexpectedly passed\ncommand: {' '.join(command)}"
        )
    if expected_token not in output:
        raise AssertionError(
            f"dry-run case {case_id!r} missing expected token {expected_token!r}\n"
            f"command: {' '.join(command)}\noutput:\n{output}"
        )
    print(f"Dry-run policy check passed ({case_id}): expected failure observed")


def main() -> int:
    args = parse_args()
    migration_path = pathlib.Path(args.migration_script)
    if not migration_path.exists():
        raise FileNotFoundError(f"dual-read migration script not found: {migration_path}")

    artifacts = _parse_csv(args.artifacts, "--artifacts")
    fault_scenarios = _parse_csv(args.fault_scenarios, "--fault-scenarios")
    first_artifact = artifacts[0]
    invalid_fault_scenarios = ",".join([*fault_scenarios, "unsupported_fault_key"])

    cases = [
        {
            "id": "invalid_wave_spec_empty_segments",
            "command": [
                sys.executable,
                str(migration_path),
                "--artifacts",
                args.artifacts,
                "--wave-spec",
                ";;",
            ],
            "expected_token": "--wave-spec did not include any non-empty wave segments",
        },
        {
            "id": "invalid_wave_spec_duplicate_across_waves",
            "command": [
                sys.executable,
                str(migration_path),
                "--artifacts",
                args.artifacts,
                "--wave-spec",
                f"{first_artifact};{first_artifact}",
            ],
            "expected_token": "--wave-spec contains duplicate artifact",
        },
        {
            "id": "unsupported_fault_scenario_key",
            "command": [
                sys.executable,
                str(migration_path),
                "--artifacts",
                args.artifacts,
                "--fault-injection",
                "--fault-scenarios",
                invalid_fault_scenarios,
            ],
            "expected_token": "--fault-scenarios includes unsupported keys",
        },
    ]

    for case in cases:
        _run_expected_failure(
            case_id=str(case["id"]),
            command=list(case["command"]),
            expected_token=str(case["expected_token"]),
        )

    print(f"Migration policy dry-run checks passed: {len(cases)} cases asserted")
    return 0


if __name__ == "__main__":
    sys.exit(main())
