#!/usr/bin/env python3
"""Run dual-read migration drills for nightly batch artifact contracts."""

from __future__ import annotations

import argparse
import copy
import json
import pathlib
import subprocess
import sys
import tempfile


ARTIFACT_KEYS = [
    "snapshot",
    "dispatch",
    "ack_retention",
    "dashboard_pack",
    "policy",
    "escalation",
    "adapter_exports",
]

VALIDATOR_LABELS = {
    "snapshot": "throughput_snapshot",
    "dispatch": "alert_dispatch",
    "ack_retention": "ack_retention_index",
    "dashboard_pack": "dashboard_pack",
    "policy": "alert_policy",
    "escalation": "policy_escalation",
    "adapter_exports": "adapter_exports",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Exercise producer/consumer overlap and rollback behavior for artifact "
            "schema major-version migrations."
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
    parser.add_argument(
        "--artifacts",
        default="snapshot,dispatch,ack_retention,dashboard_pack,policy,escalation,adapter_exports",
        help=(
            "Comma-separated artifact keys to include in dual-read matrix. "
            "Allowed values: snapshot,dispatch,ack_retention,dashboard_pack,policy,escalation,adapter_exports."
        ),
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
    schema_paths: dict[str, pathlib.Path],
    fallback_schema_paths: dict[str, pathlib.Path],
) -> list[str]:
    command = [
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
        str(schema_paths["snapshot"]),
        "--schema-dispatch",
        str(schema_paths["dispatch"]),
        "--schema-ack-retention",
        str(schema_paths["ack_retention"]),
        "--schema-dashboard-pack",
        str(schema_paths["dashboard_pack"]),
        "--schema-policy",
        str(schema_paths["policy"]),
        "--schema-escalation",
        str(schema_paths["escalation"]),
        "--schema-adapter-exports",
        str(schema_paths["adapter_exports"]),
    ]
    fallback_arg_map = {
        "snapshot": "--schema-snapshot-fallback",
        "dispatch": "--schema-dispatch-fallback",
        "ack_retention": "--schema-ack-retention-fallback",
        "dashboard_pack": "--schema-dashboard-pack-fallback",
        "policy": "--schema-policy-fallback",
        "escalation": "--schema-escalation-fallback",
        "adapter_exports": "--schema-adapter-exports-fallback",
    }
    for key, path in fallback_schema_paths.items():
        command.extend([fallback_arg_map[key], str(path)])
    return command


def _run_validator(
    args: argparse.Namespace,
    payloads: dict[str, dict],
    schema_overrides: dict[str, dict] | None = None,
    fallback_overrides: dict[str, dict] | None = None,
) -> tuple[int, str]:
    with tempfile.TemporaryDirectory(prefix="rc-dual-read-drill-") as tmp:
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

        schema_paths = {
            "snapshot": pathlib.Path(args.schema_snapshot),
            "dispatch": pathlib.Path(args.schema_dispatch),
            "ack_retention": pathlib.Path(args.schema_ack_retention),
            "dashboard_pack": pathlib.Path(args.schema_dashboard_pack),
            "policy": pathlib.Path(args.schema_policy),
            "escalation": pathlib.Path(args.schema_escalation),
            "adapter_exports": pathlib.Path(args.schema_adapter_exports),
        }
        if schema_overrides:
            for key, payload in schema_overrides.items():
                schema_path = tmp_root / f"{key}-schema-override.json"
                _write_json(schema_path, payload)
                schema_paths[key] = schema_path

        fallback_paths: dict[str, pathlib.Path] = {}
        if fallback_overrides:
            for key, payload in fallback_overrides.items():
                fallback_path = tmp_root / f"{key}-schema-fallback.json"
                _write_json(fallback_path, payload)
                fallback_paths[key] = fallback_path

        command = _validator_command(args, payload_paths, schema_paths, fallback_paths)
        run = subprocess.run(  # noqa: S603
            command,
            capture_output=True,
            text=True,
            check=False,
        )
        output = (run.stdout or "") + (run.stderr or "")
        return run.returncode, output


def _build_synthetic_v2_schema(v1_schema: dict) -> tuple[dict, str]:
    schema_v2 = copy.deepcopy(v1_schema)
    old_id = str(schema_v2.get("$id", ""))
    if "/v1/" in old_id:
        new_id = old_id.replace("/v1/", "/v2/")
    else:
        new_id = f"{old_id}.v2"
    schema_v2["$id"] = new_id

    title = str(schema_v2.get("title", "RootCellar Artifact Schema"))
    if "v2" not in title.lower():
        schema_v2["title"] = f"{title} v2 Synthetic"

    artifact_contract = (
        schema_v2.get("properties", {}).get("artifact_contract", {}).get("properties", {})
    )
    if isinstance(artifact_contract, dict):
        schema_id_prop = artifact_contract.get("schema_id")
        if isinstance(schema_id_prop, dict):
            schema_id_prop["const"] = new_id

    contract_meta = schema_v2.get("x-rootcellar-contract")
    if isinstance(contract_meta, dict):
        contract_meta["schema_version"] = "2.0.0"

    return schema_v2, new_id


def _require_contains(output: str, token: str, phase: str) -> None:
    if token not in output:
        raise AssertionError(
            f"dual-read drill phase {phase!r} missing expected token {token!r}\noutput:\n{output}"
        )


def _parse_artifacts(raw: str) -> list[str]:
    values = [token.strip() for token in raw.split(",") if token.strip()]
    if not values:
        raise ValueError("--artifacts must include at least one artifact key")
    unknown = sorted({item for item in values if item not in ARTIFACT_KEYS})
    if unknown:
        raise ValueError(f"--artifacts includes unsupported keys: {', '.join(unknown)}")
    deduped: list[str] = []
    seen: set[str] = set()
    for item in values:
        if item in seen:
            continue
        deduped.append(item)
        seen.add(item)
    return deduped


def _schema_paths_by_key(args: argparse.Namespace) -> dict[str, pathlib.Path]:
    return {
        "snapshot": pathlib.Path(args.schema_snapshot),
        "dispatch": pathlib.Path(args.schema_dispatch),
        "ack_retention": pathlib.Path(args.schema_ack_retention),
        "dashboard_pack": pathlib.Path(args.schema_dashboard_pack),
        "policy": pathlib.Path(args.schema_policy),
        "escalation": pathlib.Path(args.schema_escalation),
        "adapter_exports": pathlib.Path(args.schema_adapter_exports),
    }


def _run_dual_read_case(
    args: argparse.Namespace,
    base_payloads: dict[str, dict],
    artifact_key: str,
    schema_v2: dict,
    schema_v2_id: str,
) -> None:
    label = VALIDATOR_LABELS[artifact_key]

    upgraded_payloads = copy.deepcopy(base_payloads)
    upgraded_payloads[artifact_key]["artifact_contract"]["schema_id"] = schema_v2_id
    upgraded_payloads[artifact_key]["artifact_contract"]["schema_version"] = "2.0.0"

    code, output = _run_validator(args, copy.deepcopy(upgraded_payloads))
    phase_fail = f"{artifact_key}_producer_v2_consumer_v1"
    if code == 0:
        raise AssertionError(
            f"phase {phase_fail} unexpectedly passed; rollback detection broken"
        )
    _require_contains(output, f"{label}: artifact_contract.schema_id", phase_fail)
    print(f"Dual-read drill passed ({phase_fail} expected fail)")

    code, output = _run_validator(
        args,
        copy.deepcopy(upgraded_payloads),
        fallback_overrides={artifact_key: schema_v2},
    )
    phase_overlap = f"{artifact_key}_producer_v2_consumer_dual_read"
    if code != 0:
        raise AssertionError(
            f"phase {phase_overlap} failed unexpectedly\nvalidator output:\n{output}"
        )
    print(f"Dual-read drill passed ({phase_overlap})")

    code, output = _run_validator(
        args,
        copy.deepcopy(base_payloads),
        fallback_overrides={artifact_key: schema_v2},
    )
    phase_prod_rollback = f"{artifact_key}_producer_v1_consumer_dual_read"
    if code != 0:
        raise AssertionError(
            f"phase {phase_prod_rollback} failed unexpectedly\nvalidator output:\n{output}"
        )
    print(f"Dual-read drill passed ({phase_prod_rollback})")

    code, output = _run_validator(args, copy.deepcopy(base_payloads))
    phase_consumer_rollback = f"{artifact_key}_producer_v1_consumer_v1_post_rollback"
    if code != 0:
        raise AssertionError(
            f"phase {phase_consumer_rollback} failed unexpectedly\nvalidator output:\n{output}"
        )
    print(f"Dual-read drill passed ({phase_consumer_rollback})")


def main() -> int:
    args = parse_args()
    selected_artifacts = _parse_artifacts(args.artifacts)

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
    schemas_v1 = {
        key: _read_json(path)
        for key, path in _schema_paths_by_key(args).items()
    }
    schemas_v2 = {
        key: _build_synthetic_v2_schema(schema_v1)
        for key, schema_v1 in schemas_v1.items()
    }

    # Phase 1: baseline with producer v1 and consumer v1.
    code, output = _run_validator(args, copy.deepcopy(base_payloads))
    if code != 0:
        raise AssertionError(
            "phase baseline_v1_to_v1 failed unexpectedly\nvalidator output:\n" + output
        )
    print("Dual-read drill passed (baseline_v1_to_v1)")

    for artifact_key in selected_artifacts:
        schema_v2, schema_v2_id = schemas_v2[artifact_key]
        _run_dual_read_case(args, base_payloads, artifact_key, schema_v2, schema_v2_id)

    print(
        "Dual-read migration drill passed: producer/consumer overlap and rollback "
        f"verified for artifacts={','.join(selected_artifacts)}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
