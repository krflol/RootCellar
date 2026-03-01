#!/usr/bin/env python3
"""Run schema-drift canary checks against nightly batch artifact contracts."""

from __future__ import annotations

import argparse
import copy
import json
import pathlib
import subprocess
import sys
import tempfile


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Validate schema-drift canary expectations by mutating nightly "
            "artifact payloads and asserting validator pass/fail behavior."
        )
    )
    parser.add_argument(
        "--validator-script",
        default="./python/validate_batch_adapter_contracts.py",
        help="Path to batch artifact contract validator script.",
    )
    parser.add_argument(
        "--snapshot",
        default="./ci-batch-throughput-snapshot.json",
        help="Path to throughput snapshot artifact JSON.",
    )
    parser.add_argument(
        "--dispatch",
        default="./ci-batch-alert-dispatch.json",
        help="Path to dispatch report artifact JSON.",
    )
    parser.add_argument(
        "--ack-retention",
        default="./ci-batch-ack-retention-index.json",
        help="Path to acknowledgement retention index artifact JSON.",
    )
    parser.add_argument(
        "--dashboard-pack",
        default="./ci-batch-dashboard-pack.json",
        help="Path to dashboard-pack artifact JSON.",
    )
    parser.add_argument(
        "--policy",
        default="./ci-batch-alert-policy.json",
        help="Path to alert-policy artifact JSON.",
    )
    parser.add_argument(
        "--escalation",
        default="./ci-batch-policy-escalation.json",
        help="Path to policy escalation artifact JSON.",
    )
    parser.add_argument(
        "--adapter-exports",
        default="./ci-batch-dashboard-adapter-exports.json",
        help="Path to adapter exports artifact JSON.",
    )
    parser.add_argument(
        "--schema-snapshot",
        default="./schemas/artifacts/v1/batch-throughput-snapshot.schema.json",
        help="Path to throughput snapshot schema JSON.",
    )
    parser.add_argument(
        "--schema-dispatch",
        default="./schemas/artifacts/v1/batch-alert-dispatch.schema.json",
        help="Path to dispatch report schema JSON.",
    )
    parser.add_argument(
        "--schema-ack-retention",
        default="./schemas/artifacts/v1/batch-ack-retention-index.schema.json",
        help="Path to acknowledgement retention index schema JSON.",
    )
    parser.add_argument(
        "--schema-dashboard-pack",
        default="./schemas/artifacts/v1/batch-dashboard-pack.schema.json",
        help="Path to dashboard-pack schema JSON.",
    )
    parser.add_argument(
        "--schema-policy",
        default="./schemas/artifacts/v1/batch-alert-policy.schema.json",
        help="Path to alert-policy schema JSON.",
    )
    parser.add_argument(
        "--schema-escalation",
        default="./schemas/artifacts/v1/batch-policy-escalation.schema.json",
        help="Path to policy escalation schema JSON.",
    )
    parser.add_argument(
        "--schema-adapter-exports",
        default="./schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json",
        help="Path to adapter exports schema JSON.",
    )
    return parser.parse_args()


def _read_json(path: pathlib.Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        payload = json.load(fh)
    if not isinstance(payload, dict):
        raise ValueError(f"expected JSON object at {path}")
    return payload


def _write_json(path: pathlib.Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _validator_command(
    args: argparse.Namespace,
    payload_paths: dict[str, pathlib.Path],
) -> list[str]:
    return [
        sys.executable,
        str(pathlib.Path(args.validator_script)),
        "--full-family",
        "--snapshot",
        str(payload_paths["snapshot"]),
        "--dispatch",
        str(payload_paths["dispatch"]),
        "--ack-retention",
        str(payload_paths["ack_retention"]),
        "--dashboard-pack",
        str(payload_paths["dashboard_pack"]),
        "--policy",
        str(payload_paths["policy"]),
        "--escalation",
        str(payload_paths["escalation"]),
        "--adapter-exports",
        str(payload_paths["adapter_exports"]),
        "--schema-snapshot",
        args.schema_snapshot,
        "--schema-dispatch",
        args.schema_dispatch,
        "--schema-ack-retention",
        args.schema_ack_retention,
        "--schema-dashboard-pack",
        args.schema_dashboard_pack,
        "--schema-policy",
        args.schema_policy,
        "--schema-escalation",
        args.schema_escalation,
        "--schema-adapter-exports",
        args.schema_adapter_exports,
    ]


def _run_validator(
    args: argparse.Namespace,
    payloads: dict[str, dict],
) -> tuple[int, str]:
    with tempfile.TemporaryDirectory(prefix="rc-schema-canary-") as tmp:
        tmp_root = pathlib.Path(tmp)
        payload_paths = {
            "snapshot": tmp_root / "snapshot.json",
            "dispatch": tmp_root / "dispatch.json",
            "ack_retention": tmp_root / "ack-retention.json",
            "dashboard_pack": tmp_root / "dashboard-pack.json",
            "policy": tmp_root / "policy.json",
            "escalation": tmp_root / "escalation.json",
            "adapter_exports": tmp_root / "adapter-exports.json",
        }
        for key, path in payload_paths.items():
            _write_json(path, payloads[key])

        cmd = _validator_command(args, payload_paths)
        run = subprocess.run(  # noqa: S603
            cmd,
            capture_output=True,
            text=True,
            check=False,
        )
        output = (run.stdout or "") + (run.stderr or "")
        return run.returncode, output


def _require_substrings(output: str, substrings: list[str], case_id: str) -> None:
    for token in substrings:
        if token not in output:
            raise AssertionError(
                f"canary {case_id!r} missing expected output token {token!r}\noutput:\n{output}"
            )


def main() -> int:
    args = parse_args()

    validator_path = pathlib.Path(args.validator_script)
    if not validator_path.exists():
        raise FileNotFoundError(f"validator script not found: {validator_path}")

    base_payloads = {
        "snapshot": _read_json(pathlib.Path(args.snapshot)),
        "dispatch": _read_json(pathlib.Path(args.dispatch)),
        "ack_retention": _read_json(pathlib.Path(args.ack_retention)),
        "dashboard_pack": _read_json(pathlib.Path(args.dashboard_pack)),
        "policy": _read_json(pathlib.Path(args.policy)),
        "escalation": _read_json(pathlib.Path(args.escalation)),
        "adapter_exports": _read_json(pathlib.Path(args.adapter_exports)),
    }

    # Baseline: current artifacts must pass.
    code, output = _run_validator(args, copy.deepcopy(base_payloads))
    if code != 0:
        raise AssertionError(
            "baseline full-family validation failed in canary harness\n"
            f"validator output:\n{output}"
        )
    print("Canary baseline passed: canonical artifacts validated")

    cases = [
        {
            "id": "missing_snapshot_contract",
            "mutate": lambda p: p["snapshot"].pop("artifact_contract", None),
            "expect_tokens": [
                "throughput_snapshot: payload missing artifact_contract object",
            ],
        },
        {
            "id": "dispatch_schema_id_mismatch",
            "mutate": lambda p: p["dispatch"]["artifact_contract"].__setitem__(
                "schema_id",
                "https://rootcellar.dev/schemas/artifacts/v1/invalid-dispatch.schema.json",
            ),
            "expect_tokens": [
                "alert_dispatch: artifact_contract.schema_id",
            ],
        },
        {
            "id": "ack_retention_major_semver_mismatch",
            "mutate": lambda p: p["ack_retention"]["artifact_contract"].__setitem__(
                "schema_version",
                "2.0.0",
            ),
            "expect_tokens": [
                "ack_retention_index: incompatible schema major version",
            ],
        },
        {
            "id": "dashboard_pack_payload_version_mismatch",
            "mutate": lambda p: p["dashboard_pack"].__setitem__(
                "dashboard_pack_version",
                2,
            ),
            "expect_tokens": [
                "dashboard_pack: payload version mismatch for 'dashboard_pack_version'",
            ],
        },
        {
            "id": "policy_compatibility_mode_mismatch",
            "mutate": lambda p: p["policy"]["artifact_contract"].__setitem__(
                "compatibility",
                "breaking-change",
            ),
            "expect_tokens": [
                "alert_policy: compatibility mode mismatch",
            ],
        },
        {
            "id": "adapter_exports_payload_version_mismatch",
            "mutate": lambda p: p["adapter_exports"].__setitem__(
                "adapter_exports_version",
                2,
            ),
            "expect_tokens": [
                "adapter_exports: payload version mismatch for 'adapter_exports_version'",
            ],
        },
    ]

    failed_cases = 0
    for case in cases:
        payloads = copy.deepcopy(base_payloads)
        case["mutate"](payloads)
        code, output = _run_validator(args, payloads)
        if code == 0:
            failed_cases += 1
            print(f"Canary FAILED ({case['id']}): validator unexpectedly passed")
            continue
        try:
            _require_substrings(output, case["expect_tokens"], str(case["id"]))
        except AssertionError as exc:
            failed_cases += 1
            print(f"Canary FAILED ({case['id']}): {exc}")
            continue
        print(f"Canary passed ({case['id']}): expected contract failure observed")

    if failed_cases:
        print(f"Schema canary validation failed: {failed_cases} case(s) did not behave as expected")
        return 1

    print(f"Schema canary validation passed: {len(cases)} drift scenarios asserted")
    return 0


if __name__ == "__main__":
    sys.exit(main())
