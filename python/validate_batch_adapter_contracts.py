#!/usr/bin/env python3
"""Validate nightly batch artifacts against schema and compatibility contracts."""

from __future__ import annotations

import argparse
import json
import pathlib
import re
import sys
from typing import Any


SEMVER_RE = re.compile(r"^(\d+)\.(\d+)\.(\d+)$")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Validate nightly batch artifacts against JSON schema files and "
            "compatibility-version contracts."
        )
    )
    parser.add_argument(
        "--full-family",
        action="store_true",
        help="Validate snapshot/dispatch/ack-index/dashboard-pack/policy artifacts in addition to adapter exports.",
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


def _resolve_ref(schema_root: dict, ref: str) -> dict:
    if not ref.startswith("#/"):
        raise ValueError(f"unsupported $ref format: {ref}")
    node: Any = schema_root
    for part in ref[2:].split("/"):
        if not isinstance(node, dict) or part not in node:
            raise KeyError(f"unable to resolve $ref segment '{part}' in {ref}")
        node = node[part]
    if not isinstance(node, dict):
        raise TypeError(f"resolved $ref is not an object schema: {ref}")
    return node


def _matches_type(value: Any, expected_type: str) -> bool:
    if expected_type == "object":
        return isinstance(value, dict)
    if expected_type == "array":
        return isinstance(value, list)
    if expected_type == "string":
        return isinstance(value, str)
    if expected_type == "integer":
        return isinstance(value, int) and not isinstance(value, bool)
    if expected_type == "number":
        return (isinstance(value, int) and not isinstance(value, bool)) or isinstance(value, float)
    if expected_type == "boolean":
        return isinstance(value, bool)
    if expected_type == "null":
        return value is None
    return True


def _validate_node(value: Any, schema: dict, schema_root: dict, path: str, errors: list[str]) -> None:
    if "$ref" in schema:
        ref_schema = _resolve_ref(schema_root, str(schema["$ref"]))
        _validate_node(value, ref_schema, schema_root, path, errors)
        return

    if "type" in schema:
        expected = schema["type"]
        if isinstance(expected, list):
            if not any(_matches_type(value, str(item)) for item in expected):
                errors.append(f"{path}: expected type one of {expected}, got {type(value).__name__}")
                return
        else:
            expected_type = str(expected)
            if not _matches_type(value, expected_type):
                errors.append(f"{path}: expected type {expected_type}, got {type(value).__name__}")
                return

    if "enum" in schema:
        enum_values = schema["enum"]
        if value not in enum_values:
            errors.append(f"{path}: value {value!r} not in enum {enum_values!r}")

    if "const" in schema:
        if value != schema["const"]:
            errors.append(f"{path}: expected const {schema['const']!r}, got {value!r}")

    if isinstance(value, str):
        min_length = schema.get("minLength")
        if isinstance(min_length, int) and len(value) < min_length:
            errors.append(f"{path}: expected minLength {min_length}, got {len(value)}")
        pattern = schema.get("pattern")
        if isinstance(pattern, str) and not re.match(pattern, value):
            errors.append(f"{path}: value {value!r} does not match pattern {pattern!r}")

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        minimum = schema.get("minimum")
        if isinstance(minimum, (int, float)) and value < minimum:
            errors.append(f"{path}: expected minimum {minimum}, got {value}")

    if isinstance(value, dict):
        required = schema.get("required", [])
        if isinstance(required, list):
            for key in required:
                if key not in value:
                    errors.append(f"{path}: missing required key {key!r}")

        properties = schema.get("properties", {})
        if isinstance(properties, dict):
            for key, child_schema in properties.items():
                if key not in value:
                    continue
                if not isinstance(child_schema, dict):
                    continue
                _validate_node(value[key], child_schema, schema_root, f"{path}.{key}", errors)

        additional = schema.get("additionalProperties", True)
        if additional is False and isinstance(properties, dict):
            for key in value:
                if key not in properties:
                    errors.append(f"{path}: additional property {key!r} not allowed")

    if isinstance(value, list):
        items = schema.get("items")
        if isinstance(items, dict):
            for idx, item in enumerate(value):
                _validate_node(item, items, schema_root, f"{path}[{idx}]", errors)


def _parse_semver(raw: str) -> tuple[int, int, int] | None:
    match = SEMVER_RE.match(raw)
    if not match:
        return None
    return int(match.group(1)), int(match.group(2)), int(match.group(3))


def _validate_contract(payload: dict, schema: dict, artifact_label: str, errors: list[str]) -> None:
    contract_meta = schema.get("x-rootcellar-contract")
    if not isinstance(contract_meta, dict):
        errors.append(f"{artifact_label}: schema missing x-rootcellar-contract metadata")
        return

    artifact_contract = payload.get("artifact_contract")
    if not isinstance(artifact_contract, dict):
        errors.append(f"{artifact_label}: payload missing artifact_contract object")
        return

    schema_id = schema.get("$id")
    payload_schema_id = artifact_contract.get("schema_id")
    if schema_id != payload_schema_id:
        errors.append(
            f"{artifact_label}: artifact_contract.schema_id {payload_schema_id!r} "
            f"does not match schema $id {schema_id!r}"
        )

    schema_semver = str(contract_meta.get("schema_version", ""))
    payload_semver = str(artifact_contract.get("schema_version", ""))
    schema_parsed = _parse_semver(schema_semver)
    payload_parsed = _parse_semver(payload_semver)
    if schema_parsed is None:
        errors.append(f"{artifact_label}: schema contract version is not semver: {schema_semver!r}")
    if payload_parsed is None:
        errors.append(
            f"{artifact_label}: artifact_contract.schema_version is not semver: {payload_semver!r}"
        )
    if schema_parsed and payload_parsed and schema_parsed[0] != payload_parsed[0]:
        errors.append(
            f"{artifact_label}: incompatible schema major version "
            f"(schema={schema_semver}, payload={payload_semver})"
        )

    expected_compatibility = str(contract_meta.get("compatibility_mode", ""))
    actual_compatibility = str(artifact_contract.get("compatibility", ""))
    if expected_compatibility and expected_compatibility != actual_compatibility:
        errors.append(
            f"{artifact_label}: compatibility mode mismatch "
            f"(schema={expected_compatibility!r}, payload={actual_compatibility!r})"
        )

    payload_version_field = str(contract_meta.get("payload_version_field", "")).strip()
    expected_payload_version = contract_meta.get("payload_version")
    if not payload_version_field:
        errors.append(f"{artifact_label}: schema contract missing payload_version_field")
    elif payload_version_field not in payload:
        errors.append(
            f"{artifact_label}: payload missing version field {payload_version_field!r}"
        )
    else:
        actual_payload_version = payload.get(payload_version_field)
        if expected_payload_version != actual_payload_version:
            errors.append(
                f"{artifact_label}: payload version mismatch for {payload_version_field!r} "
                f"(schema={expected_payload_version!r}, payload={actual_payload_version!r})"
            )


def _validate_artifact(
    payload_path: pathlib.Path,
    schema_path: pathlib.Path,
    artifact_label: str,
) -> list[str]:
    payload = _read_json(payload_path)
    schema = _read_json(schema_path)

    errors: list[str] = []
    _validate_node(payload, schema, schema, "$", errors)
    _validate_contract(payload, schema, artifact_label, errors)
    return errors


def _artifact_pairs(args: argparse.Namespace) -> list[tuple[pathlib.Path, pathlib.Path, str]]:
    pairs = [
        (
            pathlib.Path(args.escalation),
            pathlib.Path(args.schema_escalation),
            "policy_escalation",
        ),
        (
            pathlib.Path(args.adapter_exports),
            pathlib.Path(args.schema_adapter_exports),
            "adapter_exports",
        ),
    ]
    if not args.full_family:
        return pairs

    extended_pairs = [
        (
            pathlib.Path(args.snapshot),
            pathlib.Path(args.schema_snapshot),
            "throughput_snapshot",
        ),
        (
            pathlib.Path(args.dispatch),
            pathlib.Path(args.schema_dispatch),
            "alert_dispatch",
        ),
        (
            pathlib.Path(args.ack_retention),
            pathlib.Path(args.schema_ack_retention),
            "ack_retention_index",
        ),
        (
            pathlib.Path(args.dashboard_pack),
            pathlib.Path(args.schema_dashboard_pack),
            "dashboard_pack",
        ),
        (
            pathlib.Path(args.policy),
            pathlib.Path(args.schema_policy),
            "alert_policy",
        ),
    ]
    return extended_pairs + pairs


def main() -> int:
    args = parse_args()
    pairs = _artifact_pairs(args)

    all_errors: list[str] = []
    for payload_path, schema_path, label in pairs:
        if not payload_path.exists():
            all_errors.append(f"{label}: payload file not found: {payload_path}")
            continue
        if not schema_path.exists():
            all_errors.append(f"{label}: schema file not found: {schema_path}")
            continue

        errors = _validate_artifact(payload_path, schema_path, label)
        if errors:
            all_errors.extend(errors)
        else:
            print(
                f"Validated {label}: payload={payload_path} schema={schema_path}"
            )

    if all_errors:
        print("Batch artifact contract validation failed with errors:")
        for err in all_errors:
            print(f" - {err}")
        return 1

    if args.full_family:
        print("Batch artifact contract validation passed (full family)")
    else:
        print("Batch artifact contract validation passed (adapter mode)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
