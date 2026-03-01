#!/usr/bin/env python3
"""Run dual-read migration drills for nightly batch artifact contracts."""

from __future__ import annotations

import argparse
import copy
from datetime import UTC, datetime
import json
import pathlib
import subprocess
import sys
import tempfile
import time


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

FAULT_SCENARIO_KEYS = [
    "malformed_fallback_schema",
    "partial_wave_rollback",
]


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
    parser.add_argument(
        "--wave-spec",
        default="",
        help=(
            "Optional semicolon-separated staged rollout waves where each wave is a "
            "comma-separated artifact key list. Example: "
            "snapshot,dispatch;ack_retention,dashboard_pack;policy,escalation,adapter_exports. "
            "When set, waves must cover exactly the --artifacts set with no duplicates."
        ),
    )
    parser.add_argument(
        "--report",
        default="",
        help=(
            "Optional path to write structured dual-read drill diagnostics JSON "
            "(phase-level status, timing, and validator output excerpts)."
        ),
    )
    parser.add_argument(
        "--report-output-max-chars",
        type=int,
        default=2000,
        help=(
            "Maximum number of validator output characters retained per phase in the "
            "structured report."
        ),
    )
    parser.add_argument(
        "--fault-injection",
        action="store_true",
        help=(
            "Enable additional fault-injection drills (malformed fallback schema and "
            "partial-wave rollback rehearsal)."
        ),
    )
    parser.add_argument(
        "--fault-scenarios",
        default="malformed_fallback_schema,partial_wave_rollback",
        help=(
            "Comma-separated fault-injection scenario keys. Allowed values: "
            "malformed_fallback_schema,partial_wave_rollback."
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


def _parse_wave_spec(raw: str, selected_artifacts: list[str]) -> list[list[str]]:
    if not raw.strip():
        return [list(selected_artifacts)]
    waves_raw = [segment.strip() for segment in raw.split(";") if segment.strip()]
    if not waves_raw:
        raise ValueError("--wave-spec did not include any non-empty wave segments")

    waves: list[list[str]] = []
    seen: set[str] = set()
    for idx, segment in enumerate(waves_raw, start=1):
        wave = _parse_artifacts(segment)
        for artifact in wave:
            if artifact in seen:
                raise ValueError(
                    f"--wave-spec contains duplicate artifact {artifact!r} across waves"
                )
            seen.add(artifact)
        waves.append(wave)
        if not wave:
            raise ValueError(f"--wave-spec wave #{idx} is empty")

    selected_set = set(selected_artifacts)
    seen_set = set(seen)
    missing = sorted(selected_set - seen_set)
    extra = sorted(seen_set - selected_set)
    if missing or extra:
        details: list[str] = []
        if missing:
            details.append(f"missing: {','.join(missing)}")
        if extra:
            details.append(f"outside --artifacts: {','.join(extra)}")
        joined = "; ".join(details)
        raise ValueError(
            "--wave-spec must cover exactly the --artifacts set " f"({joined})"
        )
    return waves


def _parse_fault_scenarios(raw: str) -> list[str]:
    values = [token.strip() for token in raw.split(",") if token.strip()]
    if not values:
        return []
    unknown = sorted({item for item in values if item not in FAULT_SCENARIO_KEYS})
    if unknown:
        raise ValueError(
            f"--fault-scenarios includes unsupported keys: {', '.join(unknown)}"
        )
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


def _utc_now_iso() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _clip_output(output: str, max_chars: int) -> str:
    if max_chars <= 0:
        return ""
    text = output.strip()
    if len(text) <= max_chars:
        return text
    keep = max_chars - 3
    if keep <= 0:
        return text[:max_chars]
    return text[:keep] + "..."


def _evaluate_phase(
    *,
    args: argparse.Namespace,
    phase_name: str,
    artifact_key: str | None,
    expectation: str,
    payloads: dict[str, dict],
    fallback_overrides: dict[str, dict] | None = None,
    expected_token: str | None = None,
) -> dict:
    start_perf = time.perf_counter()
    started_at = _utc_now_iso()
    code, output = _run_validator(
        args,
        copy.deepcopy(payloads),
        fallback_overrides=fallback_overrides,
    )
    duration_ms = round((time.perf_counter() - start_perf) * 1000, 3)
    token_found = expected_token in output if expected_token else None
    status = "pass"
    failure_message = ""

    if expectation == "must_pass":
        if code != 0:
            status = "fail"
            failure_message = (
                f"phase {phase_name} failed unexpectedly\nvalidator output:\n{output}"
            )
    elif expectation == "must_fail_with_token":
        if code == 0:
            status = "fail"
            failure_message = (
                f"phase {phase_name} unexpectedly passed; rollback detection broken"
            )
        elif expected_token and not token_found:
            status = "fail"
            failure_message = (
                f"dual-read drill phase {phase_name!r} missing expected token "
                f"{expected_token!r}\noutput:\n{output}"
            )
    else:
        raise ValueError(f"unsupported expectation: {expectation}")

    return {
        "phase": phase_name,
        "artifact": artifact_key,
        "expectation": expectation,
        "status": status,
        "started_at": started_at,
        "duration_ms": duration_ms,
        "validator_exit_code": code,
        "expected_token": expected_token,
        "expected_token_found": token_found,
        "fallback_schema_enabled": bool(fallback_overrides),
        "validator_output_excerpt": _clip_output(
            output, args.report_output_max_chars
        ),
        "failure_message": failure_message,
    }


def _append_phase_or_raise(phase_records: list[dict], record: dict) -> None:
    phase_records.append(record)
    if record["status"] != "pass":
        message = str(record.get("failure_message") or "dual-read drill phase failed")
        raise AssertionError(message)


def _build_upgraded_payloads(
    base_payloads: dict[str, dict],
    schema_v2_ids: dict[str, str],
    artifacts: list[str],
) -> dict[str, dict]:
    payloads = copy.deepcopy(base_payloads)
    for artifact_key in artifacts:
        payloads[artifact_key]["artifact_contract"]["schema_id"] = schema_v2_ids[
            artifact_key
        ]
        payloads[artifact_key]["artifact_contract"]["schema_version"] = "2.0.0"
    return payloads


def _build_malformed_fallback_schema(schema_v2: dict) -> dict:
    malformed = copy.deepcopy(schema_v2)
    malformed.pop("x-rootcellar-contract", None)
    malformed["title"] = str(malformed.get("title", "RootCellar Artifact Schema")) + (
        " malformed fallback"
    )
    return malformed


def _select_partial_wave_artifacts(
    waves: list[list[str]], selected_artifacts: list[str]
) -> list[str] | None:
    for wave in waves:
        if len(wave) >= 2:
            return [wave[0], wave[1]]
    if len(selected_artifacts) >= 2:
        return [selected_artifacts[0], selected_artifacts[1]]
    return None


def _run_fault_malformed_fallback_schema(
    *,
    args: argparse.Namespace,
    phase_records: list[dict],
    base_payloads: dict[str, dict],
    schema_v2_by_key: dict[str, dict],
    schema_v2_id_by_key: dict[str, str],
    selected_artifacts: list[str],
) -> bool:
    if not selected_artifacts:
        return False
    for artifact_key in selected_artifacts:
        upgraded_payloads = _build_upgraded_payloads(
            base_payloads, schema_v2_id_by_key, [artifact_key]
        )
        malformed_fallback = _build_malformed_fallback_schema(
            schema_v2_by_key[artifact_key]
        )
        phase_name = f"fault_malformed_fallback_{artifact_key}_producer_v2_consumer_dual_read"
        record = _evaluate_phase(
            args=args,
            phase_name=phase_name,
            artifact_key=artifact_key,
            expectation="must_fail_with_token",
            payloads=upgraded_payloads,
            fallback_overrides={artifact_key: malformed_fallback},
            expected_token=f"{VALIDATOR_LABELS[artifact_key]}: schema missing x-rootcellar-contract metadata",
        )
        _append_phase_or_raise(phase_records, record)
        print(f"Dual-read drill passed ({phase_name} expected fail)")
    return True


def _run_fault_partial_wave_rollback(
    *,
    args: argparse.Namespace,
    phase_records: list[dict],
    base_payloads: dict[str, dict],
    schema_v2_by_key: dict[str, dict],
    schema_v2_id_by_key: dict[str, str],
    waves: list[list[str]],
    selected_artifacts: list[str],
) -> bool:
    pair = _select_partial_wave_artifacts(waves, selected_artifacts)
    if pair is None:
        print(
            "Dual-read drill skipped (fault_partial_wave_rollback not applicable: "
            "fewer than two artifacts selected)"
        )
        return False

    primary_key, secondary_key = pair[0], pair[1]
    dual_read_fallbacks = {
        key: schema_v2_by_key[key]
        for key in pair
    }
    all_upgraded_payloads = _build_upgraded_payloads(base_payloads, schema_v2_id_by_key, pair)

    phase_all_v2 = f"fault_partial_wave_{primary_key}_{secondary_key}_all_v2_consumer_dual_read"
    all_v2_record = _evaluate_phase(
        args=args,
        phase_name=phase_all_v2,
        artifact_key=None,
        expectation="must_pass",
        payloads=all_upgraded_payloads,
        fallback_overrides=dual_read_fallbacks,
    )
    _append_phase_or_raise(phase_records, all_v2_record)
    print(f"Dual-read drill passed ({phase_all_v2})")

    partial_payloads = copy.deepcopy(all_upgraded_payloads)
    partial_payloads[primary_key]["artifact_contract"]["schema_id"] = base_payloads[
        primary_key
    ]["artifact_contract"]["schema_id"]
    partial_payloads[primary_key]["artifact_contract"]["schema_version"] = base_payloads[
        primary_key
    ]["artifact_contract"]["schema_version"]

    phase_partial_dual = (
        f"fault_partial_wave_{primary_key}_{secondary_key}_partial_rollback_consumer_dual_read"
    )
    partial_dual_record = _evaluate_phase(
        args=args,
        phase_name=phase_partial_dual,
        artifact_key=None,
        expectation="must_pass",
        payloads=partial_payloads,
        fallback_overrides=dual_read_fallbacks,
    )
    _append_phase_or_raise(phase_records, partial_dual_record)
    print(f"Dual-read drill passed ({phase_partial_dual})")

    phase_partial_strict = (
        f"fault_partial_wave_{primary_key}_{secondary_key}_partial_rollback_consumer_v1"
    )
    partial_strict_record = _evaluate_phase(
        args=args,
        phase_name=phase_partial_strict,
        artifact_key=secondary_key,
        expectation="must_fail_with_token",
        payloads=partial_payloads,
        expected_token=f"{VALIDATOR_LABELS[secondary_key]}: artifact_contract.schema_id",
    )
    _append_phase_or_raise(phase_records, partial_strict_record)
    print(f"Dual-read drill passed ({phase_partial_strict} expected fail)")

    phase_full_rollback = (
        f"fault_partial_wave_{primary_key}_{secondary_key}_full_rollback_consumer_v1"
    )
    full_rollback_record = _evaluate_phase(
        args=args,
        phase_name=phase_full_rollback,
        artifact_key=None,
        expectation="must_pass",
        payloads=base_payloads,
    )
    _append_phase_or_raise(phase_records, full_rollback_record)
    print(f"Dual-read drill passed ({phase_full_rollback})")
    return True


def _run_dual_read_case(
    *,
    args: argparse.Namespace,
    phase_records: list[dict],
    base_payloads: dict[str, dict],
    artifact_key: str,
    schema_v2: dict,
    schema_v2_id: str,
    phase_prefix: str = "",
) -> None:
    label = VALIDATOR_LABELS[artifact_key]

    def _phase(name: str) -> str:
        if not phase_prefix:
            return name
        return f"{phase_prefix}_{name}"

    upgraded_payloads = copy.deepcopy(base_payloads)
    upgraded_payloads[artifact_key]["artifact_contract"]["schema_id"] = schema_v2_id
    upgraded_payloads[artifact_key]["artifact_contract"]["schema_version"] = "2.0.0"

    phase_fail = _phase(f"{artifact_key}_producer_v2_consumer_v1")
    fail_record = _evaluate_phase(
        args=args,
        phase_name=phase_fail,
        artifact_key=artifact_key,
        expectation="must_fail_with_token",
        payloads=upgraded_payloads,
        expected_token=f"{label}: artifact_contract.schema_id",
    )
    _append_phase_or_raise(phase_records, fail_record)
    print(f"Dual-read drill passed ({phase_fail} expected fail)")

    phase_overlap = _phase(f"{artifact_key}_producer_v2_consumer_dual_read")
    overlap_record = _evaluate_phase(
        args=args,
        phase_name=phase_overlap,
        artifact_key=artifact_key,
        expectation="must_pass",
        payloads=upgraded_payloads,
        fallback_overrides={artifact_key: schema_v2},
    )
    _append_phase_or_raise(phase_records, overlap_record)
    print(f"Dual-read drill passed ({phase_overlap})")

    phase_prod_rollback = _phase(f"{artifact_key}_producer_v1_consumer_dual_read")
    prod_rollback_record = _evaluate_phase(
        args=args,
        phase_name=phase_prod_rollback,
        artifact_key=artifact_key,
        expectation="must_pass",
        payloads=base_payloads,
        fallback_overrides={artifact_key: schema_v2},
    )
    _append_phase_or_raise(phase_records, prod_rollback_record)
    print(f"Dual-read drill passed ({phase_prod_rollback})")

    phase_consumer_rollback = _phase(
        f"{artifact_key}_producer_v1_consumer_v1_post_rollback"
    )
    consumer_rollback_record = _evaluate_phase(
        args=args,
        phase_name=phase_consumer_rollback,
        artifact_key=artifact_key,
        expectation="must_pass",
        payloads=base_payloads,
    )
    _append_phase_or_raise(phase_records, consumer_rollback_record)
    print(f"Dual-read drill passed ({phase_consumer_rollback})")


def _build_report(
    *,
    args: argparse.Namespace,
    profile: str,
    selected_artifacts: list[str],
    waves: list[list[str]],
    phase_records: list[dict],
    status: str,
    error_message: str | None,
    fault_injection_enabled: bool,
    fault_scenarios: list[str],
    executed_fault_scenarios: list[str],
) -> dict:
    pass_count = sum(1 for record in phase_records if record["status"] == "pass")
    fail_count = len(phase_records) - pass_count
    return {
        "artifact_contract": {
            "schema_id": "urn:rootcellar:artifacts:batch-schema-migration-drill-report:v1",
            "schema_version": "1.0.0",
            "compatibility": "backward-additive",
        },
        "generated_at": _utc_now_iso(),
        "status": status,
        "profile": profile,
        "selected_artifacts": selected_artifacts,
        "waves": waves,
        "summary": {
            "phase_count": len(phase_records),
            "phase_pass_count": pass_count,
            "phase_fail_count": fail_count,
            "wave_count": len(waves),
            "artifact_count": len(selected_artifacts),
        },
        "policy": {
            "wave_spec": args.wave_spec,
            "artifacts": args.artifacts,
            "report_output_max_chars": args.report_output_max_chars,
        },
        "fault_injection": {
            "enabled": fault_injection_enabled,
            "requested_scenarios": fault_scenarios,
            "executed_scenarios": executed_fault_scenarios,
        },
        "validator": {
            "script": str(pathlib.Path(args.validator_script)),
            "schema_snapshot": args.schema_snapshot,
            "schema_dispatch": args.schema_dispatch,
            "schema_ack_retention": args.schema_ack_retention,
            "schema_dashboard_pack": args.schema_dashboard_pack,
            "schema_policy": args.schema_policy,
            "schema_escalation": args.schema_escalation,
            "schema_adapter_exports": args.schema_adapter_exports,
        },
        "error_message": error_message,
        "phases": phase_records,
    }


def main() -> int:
    args = parse_args()
    selected_artifacts = _parse_artifacts(args.artifacts)
    waves = _parse_wave_spec(args.wave_spec, selected_artifacts)
    fault_scenarios = _parse_fault_scenarios(args.fault_scenarios)
    profile = "staged_wave_matrix" if args.wave_spec.strip() else "single_wave_matrix"

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
    schema_v2_by_key = {
        key: value[0]
        for key, value in schemas_v2.items()
    }
    schema_v2_id_by_key = {
        key: value[1]
        for key, value in schemas_v2.items()
    }

    phase_records: list[dict] = []
    executed_fault_scenarios: list[str] = []
    status = "pass"
    error_message: str | None = None
    had_error = False

    try:
        for wave_idx, wave_artifacts in enumerate(waves, start=1):
            use_wave_prefix = bool(args.wave_spec.strip())
            wave_prefix = f"wave{wave_idx}" if use_wave_prefix else ""
            baseline_phase = (
                "baseline_v1_to_v1"
                if not use_wave_prefix
                else f"{wave_prefix}_baseline_v1_to_v1"
            )
            baseline_record = _evaluate_phase(
                args=args,
                phase_name=baseline_phase,
                artifact_key=None,
                expectation="must_pass",
                payloads=base_payloads,
            )
            _append_phase_or_raise(phase_records, baseline_record)
            print(f"Dual-read drill passed ({baseline_phase})")

            for artifact_key in wave_artifacts:
                schema_v2 = schema_v2_by_key[artifact_key]
                schema_v2_id = schema_v2_id_by_key[artifact_key]
                _run_dual_read_case(
                    args=args,
                    phase_records=phase_records,
                    base_payloads=base_payloads,
                    artifact_key=artifact_key,
                    schema_v2=schema_v2,
                    schema_v2_id=schema_v2_id,
                    phase_prefix=wave_prefix,
                )

        if args.fault_injection and "malformed_fallback_schema" in fault_scenarios:
            executed = _run_fault_malformed_fallback_schema(
                args=args,
                phase_records=phase_records,
                base_payloads=base_payloads,
                schema_v2_by_key=schema_v2_by_key,
                schema_v2_id_by_key=schema_v2_id_by_key,
                selected_artifacts=selected_artifacts,
            )
            if executed:
                executed_fault_scenarios.append("malformed_fallback_schema")

        if args.fault_injection and "partial_wave_rollback" in fault_scenarios:
            executed = _run_fault_partial_wave_rollback(
                args=args,
                phase_records=phase_records,
                base_payloads=base_payloads,
                schema_v2_by_key=schema_v2_by_key,
                schema_v2_id_by_key=schema_v2_id_by_key,
                waves=waves,
                selected_artifacts=selected_artifacts,
            )
            if executed:
                executed_fault_scenarios.append("partial_wave_rollback")

        print(
            "Dual-read migration drill passed: producer/consumer overlap and rollback "
            "verified for "
            f"profile={profile}, artifacts={','.join(selected_artifacts)}, "
            f"fault_injection={args.fault_injection}, "
            f"fault_scenarios={','.join(executed_fault_scenarios) if executed_fault_scenarios else 'none'}"
        )
    except Exception as exc:
        status = "fail"
        error_message = str(exc)
        had_error = True
    finally:
        if args.report:
            report_payload = _build_report(
                args=args,
                profile=profile,
                selected_artifacts=selected_artifacts,
                waves=waves,
                phase_records=phase_records,
                status=status,
                error_message=error_message,
                fault_injection_enabled=args.fault_injection,
                fault_scenarios=fault_scenarios,
                executed_fault_scenarios=executed_fault_scenarios,
            )
            _write_json(pathlib.Path(args.report), report_payload)
            print(f"Wrote dual-read drill report: {args.report}")

    if had_error:
        raise AssertionError(error_message or "dual-read migration drill failed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
