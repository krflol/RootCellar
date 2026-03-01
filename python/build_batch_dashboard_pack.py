#!/usr/bin/env python3
"""Build nightly batch dashboard-pack and alert-policy artifacts."""

from __future__ import annotations

import argparse
import json
import pathlib
import sys
from datetime import datetime, timezone


SEVERITY_ORDER = {
    "info": 0,
    "p3": 1,
    "p2": 2,
    "p1": 3,
}

POLICY_SCHEMA_ID = "https://rootcellar.dev/schemas/artifacts/v1/batch-alert-policy.schema.json"
DASHBOARD_PACK_SCHEMA_ID = (
    "https://rootcellar.dev/schemas/artifacts/v1/batch-dashboard-pack.schema.json"
)
ARTIFACT_SCHEMA_VERSION = "1.0.0"
ARTIFACT_COMPATIBILITY_MODE = "backward-additive"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Build dashboard-pack and alert-policy artifacts from nightly batch "
            "snapshot/dispatch/ack-retention reports."
        )
    )
    parser.add_argument(
        "--snapshot",
        required=True,
        help="Path to throughput snapshot JSON.",
    )
    parser.add_argument(
        "--dispatch-report",
        required=True,
        help="Path to dispatch report JSON.",
    )
    parser.add_argument(
        "--ack-retention-index",
        required=True,
        help="Path to acknowledgement retention index JSON.",
    )
    parser.add_argument(
        "--dashboard-pack",
        default="./ci-batch-dashboard-pack.json",
        help="Output path for dashboard pack JSON.",
    )
    parser.add_argument(
        "--policy",
        default="./ci-batch-alert-policy.json",
        help="Output path for alert policy evaluation JSON.",
    )
    parser.add_argument(
        "--max-dispatch-failed-routes",
        type=int,
        default=0,
        help="Maximum allowed failed routes in dispatch report.",
    )
    parser.add_argument(
        "--max-ack-missing-routes",
        type=int,
        default=0,
        help="Maximum allowed ack-missing routes in dispatch report.",
    )
    parser.add_argument(
        "--max-correlation-mismatch-routes",
        type=int,
        default=0,
        help="Maximum allowed correlation-mismatch routes in dispatch report.",
    )
    parser.add_argument(
        "--require-replay-metadata",
        action="store_true",
        help="Treat missing replay metadata on delivered routes as policy breach.",
    )
    parser.add_argument(
        "--require-ack-retention-coverage",
        action="store_true",
        help="Treat ack-retention lookup coverage gaps as policy breach.",
    )
    parser.add_argument(
        "--allow-missing-inputs",
        action="store_true",
        help="Emit degraded artifacts when one or more inputs are missing.",
    )
    parser.add_argument(
        "--fail-on-policy-breach",
        action="store_true",
        help="Exit non-zero when computed policy status is breach.",
    )
    return parser.parse_args()


def _read_json(path: pathlib.Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def _write_json(path: pathlib.Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _load_input(path_str: str, allow_missing: bool, issues: list[dict], input_kind: str) -> dict | None:
    path = pathlib.Path(path_str)
    if not path.exists():
        if allow_missing:
            issues.append(
                {
                    "code": "input_missing",
                    "input": input_kind,
                    "path": str(path).replace("\\", "/"),
                    "message": "input file missing; using degraded defaults",
                }
            )
            return None
        raise FileNotFoundError(f"{input_kind} file not found: {path_str}")
    return _read_json(path)


def _status_check(  # noqa: PLR0913
    check_id: str,
    title: str,
    severity: str,
    observed: object,
    expected: object,
    breach: bool,
    message: str,
    source: str,
) -> dict:
    return {
        "id": check_id,
        "title": title,
        "severity": severity,
        "status": "breach" if breach else "pass",
        "observed": observed,
        "expected": expected,
        "source": source,
        "message": message,
    }


def _highest_severity(checks: list[dict]) -> str:
    highest = "info"
    for check in checks:
        if check.get("status") != "breach":
            continue
        severity = str(check.get("severity", "info")).lower()
        if SEVERITY_ORDER.get(severity, -1) > SEVERITY_ORDER.get(highest, -1):
            highest = severity
    return highest


def _count_missing_replay_fields(dispatch: dict) -> tuple[int, list[dict]]:
    missing = 0
    route_details: list[dict] = []
    routes = dispatch.get("routes", [])
    if not isinstance(routes, list):
        return 0, []
    for route in routes:
        if not isinstance(route, dict):
            continue
        if route.get("status") != "delivered":
            continue
        delivery = route.get("delivery", {})
        if not isinstance(delivery, dict):
            delivery = {}
        final = delivery.get("final", {})
        if not isinstance(final, dict):
            final = {}
        replay = final.get("replay", {})
        if not isinstance(replay, dict):
            replay = {}
        timestamp = replay.get("timestamp")
        nonce = replay.get("nonce")
        window_sec = replay.get("window_sec")
        is_missing = not timestamp or not nonce or not window_sec
        if is_missing:
            missing += 1
        route_details.append(
            {
                "route": route.get("route"),
                "status": route.get("status"),
                "final_http_status": final.get("http_status"),
                "attempt_count": delivery.get("attempt_count"),
                "replay_timestamp": timestamp,
                "replay_nonce": nonce,
                "replay_window_sec": window_sec,
                "replay_missing": is_missing,
            }
        )
    return missing, route_details


def _safe_int(value: object, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _safe_float(value: object, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _route_rows(dispatch: dict) -> list[dict]:
    rows: list[dict] = []
    routes = dispatch.get("routes", [])
    if not isinstance(routes, list):
        return rows
    for route in routes:
        if not isinstance(route, dict):
            continue
        delivery = route.get("delivery", {})
        if not isinstance(delivery, dict):
            delivery = {}
        final = delivery.get("final", {})
        if not isinstance(final, dict):
            final = {}
        replay = final.get("replay", {})
        if not isinstance(replay, dict):
            replay = {}
        ack = route.get("ack", {})
        if not isinstance(ack, dict):
            ack = {}
        correlation = route.get("correlation", {})
        if not isinstance(correlation, dict):
            correlation = {}
        identifiers = route.get("identifiers", {})
        if not isinstance(identifiers, dict):
            identifiers = {}
        rows.append(
            {
                "route": route.get("route"),
                "status": route.get("status"),
                "endpoint": route.get("endpoint"),
                "attempt_count": _safe_int(delivery.get("attempt_count")),
                "final_http_status": final.get("http_status"),
                "ack_received": bool(ack.get("received")),
                "correlation_matches": bool(correlation.get("matches")),
                "idempotency_key": identifiers.get("idempotency_key"),
                "correlation_id": identifiers.get("correlation_id"),
                "replay_timestamp": replay.get("timestamp"),
                "replay_nonce": replay.get("nonce"),
                "replay_window_sec": replay.get("window_sec"),
            }
        )
    return rows


def _build_policy(  # noqa: PLR0913
    snapshot: dict | None,
    dispatch: dict | None,
    ack_index: dict | None,
    args: argparse.Namespace,
    issues: list[dict],
) -> dict:
    snapshot_status = str((snapshot or {}).get("status", "missing"))
    dispatch_status = str((dispatch or {}).get("status", "missing"))
    ack_index_status = str((ack_index or {}).get("status", "missing"))

    dispatch_failed_routes = _safe_int((dispatch or {}).get("failed_route_count"))
    ack_missing_routes = _safe_int((dispatch or {}).get("ack_missing_route_count"))
    correlation_mismatch_routes = _safe_int(
        (dispatch or {}).get("correlation_mismatch_route_count")
    )
    ack_received_routes = _safe_int((dispatch or {}).get("ack_received_route_count"))
    ack_lookup_count = _safe_int(
        ((ack_index or {}).get("lookup_count", {}) if isinstance((ack_index or {}).get("lookup_count"), dict) else {}).get(
            "ack_id"
        )
    )
    replay_missing_count, replay_route_details = _count_missing_replay_fields(dispatch or {})

    checks: list[dict] = []
    checks.append(
        _status_check(
            "snapshot.status.pass",
            "Snapshot status is pass",
            "p3",
            snapshot_status,
            "pass",
            snapshot_status != "pass",
            "Throughput snapshot must remain within thresholds.",
            "snapshot",
        )
    )
    checks.append(
        _status_check(
            "dispatch.failed_routes.max",
            "Dispatch failed routes within threshold",
            "p2",
            dispatch_failed_routes,
            {"max": args.max_dispatch_failed_routes},
            dispatch_failed_routes > args.max_dispatch_failed_routes,
            "Failed dispatch routes indicate downstream delivery reliability risk.",
            "dispatch",
        )
    )
    checks.append(
        _status_check(
            "dispatch.ack_missing.max",
            "Ack-missing routes within threshold",
            "p2",
            ack_missing_routes,
            {"max": args.max_ack_missing_routes},
            ack_missing_routes > args.max_ack_missing_routes,
            "Missing required acknowledgements reduce delivery traceability confidence.",
            "dispatch",
        )
    )
    checks.append(
        _status_check(
            "dispatch.correlation_mismatch.max",
            "Correlation mismatch routes within threshold",
            "p2",
            correlation_mismatch_routes,
            {"max": args.max_correlation_mismatch_routes},
            correlation_mismatch_routes > args.max_correlation_mismatch_routes,
            "Correlation mismatches prevent deterministic cross-system trace joins.",
            "dispatch",
        )
    )
    if args.require_replay_metadata:
        checks.append(
            _status_check(
                "dispatch.replay_metadata.required",
                "Replay metadata present on delivered routes",
                "p2",
                replay_missing_count,
                {"missing_allowed": 0},
                replay_missing_count > 0,
                "Replay timestamp/nonce/window fields must be present for delivered routes.",
                "dispatch",
            )
        )
    if args.require_ack_retention_coverage:
        checks.append(
            _status_check(
                "ack_retention.lookup_coverage.required",
                "Ack retention lookup covers acknowledged routes",
                "p2",
                {
                    "ack_received_route_count": ack_received_routes,
                    "ack_lookup_count": ack_lookup_count,
                },
                "ack_lookup_count >= ack_received_route_count",
                ack_lookup_count < ack_received_routes,
                "Retention lookup count should cover all acknowledged routes.",
                "ack_retention_index",
            )
        )
    if snapshot is None:
        checks.append(
            _status_check(
                "snapshot.file.present",
                "Snapshot file is present",
                "p2",
                False,
                True,
                True,
                "Snapshot artifact missing in current run.",
                "snapshot",
            )
        )
    if dispatch is None:
        checks.append(
            _status_check(
                "dispatch.file.present",
                "Dispatch report file is present",
                "p2",
                False,
                True,
                True,
                "Dispatch artifact missing in current run.",
                "dispatch",
            )
        )
    if ack_index is None:
        checks.append(
            _status_check(
                "ack_retention.file.present",
                "Ack retention index file is present",
                "p2",
                False,
                True,
                True,
                "Ack retention index artifact missing in current run.",
                "ack_retention_index",
            )
        )

    # Only alert on degraded index status when dispatch generated routes that should be indexed.
    if ack_index is not None:
        index_breach = (
            ack_index_status == "missing_dispatch_report"
            or (ack_index_status == "no_routes" and dispatch_status not in {"no_routes", "missing"})
        )
        checks.append(
            _status_check(
                "ack_retention.status.valid",
                "Ack retention index status valid",
                "p2",
                ack_index_status,
                "ok|no_ack_records|no_routes",
                index_breach,
                "Ack retention index should not report degraded dispatch availability when dispatch ran.",
                "ack_retention_index",
            )
        )

    highest_severity = _highest_severity(checks)
    breach_count = sum(1 for check in checks if check.get("status") == "breach")
    status = "breach" if breach_count > 0 else "pass"

    return {
        "artifact_contract": {
            "schema_id": POLICY_SCHEMA_ID,
            "schema_version": ARTIFACT_SCHEMA_VERSION,
            "compatibility": ARTIFACT_COMPATIBILITY_MODE,
        },
        "policy_version": 1,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "status": status,
        "highest_severity": highest_severity,
        "breach_count": breach_count,
        "check_count": len(checks),
        "policy_config": {
            "max_dispatch_failed_routes": args.max_dispatch_failed_routes,
            "max_ack_missing_routes": args.max_ack_missing_routes,
            "max_correlation_mismatch_routes": args.max_correlation_mismatch_routes,
            "require_replay_metadata": args.require_replay_metadata,
            "require_ack_retention_coverage": args.require_ack_retention_coverage,
        },
        "source_status": {
            "snapshot_status": snapshot_status,
            "dispatch_status": dispatch_status,
            "ack_retention_index_status": ack_index_status,
        },
        "checks": checks,
        "replay_route_details": replay_route_details,
        "issues": issues,
    }


def _build_dashboard_pack(
    snapshot: dict | None,
    dispatch: dict | None,
    ack_index: dict | None,
    policy: dict,
    args: argparse.Namespace,
) -> dict:
    snapshot_metrics = (snapshot or {}).get("metrics", {})
    if not isinstance(snapshot_metrics, dict):
        snapshot_metrics = {}
    dispatch_policy = (dispatch or {}).get("dispatch_policy", {})
    if not isinstance(dispatch_policy, dict):
        dispatch_policy = {}
    ack_lookup = (ack_index or {}).get("lookup_count", {})
    if not isinstance(ack_lookup, dict):
        ack_lookup = {}
    retention_policy = (ack_index or {}).get("retention_policy", {})
    if not isinstance(retention_policy, dict):
        retention_policy = {}

    route_rows = _route_rows(dispatch or {})
    failed_routes = [row for row in route_rows if row.get("status") == "failed"]

    return {
        "artifact_contract": {
            "schema_id": DASHBOARD_PACK_SCHEMA_ID,
            "schema_version": ARTIFACT_SCHEMA_VERSION,
            "compatibility": ARTIFACT_COMPATIBILITY_MODE,
        },
        "dashboard_pack_version": 1,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "status": policy.get("status"),
        "highest_severity": policy.get("highest_severity"),
        "source_files": {
            "snapshot": str(pathlib.Path(args.snapshot)).replace("\\", "/"),
            "dispatch_report": str(pathlib.Path(args.dispatch_report)).replace("\\", "/"),
            "ack_retention_index": str(pathlib.Path(args.ack_retention_index)).replace("\\", "/"),
            "policy": str(pathlib.Path(args.policy)).replace("\\", "/"),
        },
        "headline_metrics": {
            "processed_files": _safe_int(snapshot_metrics.get("processed_files")),
            "failure_count": _safe_int(snapshot_metrics.get("failure_count")),
            "throughput_files_per_sec": _safe_float(
                snapshot_metrics.get("throughput_files_per_sec")
            ),
            "dispatch_delivered_route_count": _safe_int(
                (dispatch or {}).get("delivered_route_count")
            ),
            "dispatch_failed_route_count": _safe_int((dispatch or {}).get("failed_route_count")),
            "ack_lookup_count": _safe_int(ack_lookup.get("ack_id")),
            "correlation_lookup_count": _safe_int(ack_lookup.get("correlation_id")),
            "policy_breach_count": _safe_int(policy.get("breach_count")),
        },
        "panels": {
            "throughput": {
                "snapshot_status": (snapshot or {}).get("status", "missing"),
                "thresholds": (snapshot or {}).get("thresholds", {}),
                "metrics": snapshot_metrics,
                "breaches": (snapshot or {}).get("breaches", []),
            },
            "route_delivery": {
                "dispatch_status": (dispatch or {}).get("status", "missing"),
                "dispatch_counts": {
                    "route_count": _safe_int((dispatch or {}).get("route_count")),
                    "configured_route_count": _safe_int(
                        (dispatch or {}).get("configured_route_count")
                    ),
                    "delivered_route_count": _safe_int(
                        (dispatch or {}).get("delivered_route_count")
                    ),
                    "failed_route_count": _safe_int((dispatch or {}).get("failed_route_count")),
                    "ack_missing_route_count": _safe_int(
                        (dispatch or {}).get("ack_missing_route_count")
                    ),
                    "correlation_mismatch_route_count": _safe_int(
                        (dispatch or {}).get("correlation_mismatch_route_count")
                    ),
                },
                "routes": route_rows,
                "failed_routes": failed_routes,
            },
            "forensics": {
                "ack_index_status": (ack_index or {}).get("status", "missing"),
                "retention_policy": retention_policy,
                "lookup_count": ack_lookup,
                "record_count": _safe_int((ack_index or {}).get("record_count")),
            },
            "policy": {
                "status": policy.get("status"),
                "highest_severity": policy.get("highest_severity"),
                "breach_count": policy.get("breach_count"),
                "checks": policy.get("checks", []),
            },
        },
        "drilldowns": {
            "dispatch_policy": dispatch_policy,
            "replay_route_details": policy.get("replay_route_details", []),
            "policy_config": policy.get("policy_config", {}),
        },
    }


def main() -> int:
    args = parse_args()
    if args.max_dispatch_failed_routes < 0:
        raise ValueError("--max-dispatch-failed-routes must be >= 0")
    if args.max_ack_missing_routes < 0:
        raise ValueError("--max-ack-missing-routes must be >= 0")
    if args.max_correlation_mismatch_routes < 0:
        raise ValueError("--max-correlation-mismatch-routes must be >= 0")

    issues: list[dict] = []
    snapshot = _load_input(args.snapshot, args.allow_missing_inputs, issues, "snapshot")
    dispatch = _load_input(args.dispatch_report, args.allow_missing_inputs, issues, "dispatch_report")
    ack_index = _load_input(
        args.ack_retention_index, args.allow_missing_inputs, issues, "ack_retention_index"
    )

    policy = _build_policy(snapshot, dispatch, ack_index, args, issues)
    dashboard_pack = _build_dashboard_pack(snapshot, dispatch, ack_index, policy, args)

    _write_json(pathlib.Path(args.policy), policy)
    _write_json(pathlib.Path(args.dashboard_pack), dashboard_pack)

    print(f"Wrote dashboard pack: {args.dashboard_pack}")
    print(f"Wrote alert policy: {args.policy}")
    print(
        "Policy status:",
        policy["status"],
        "| highest_severity=",
        policy["highest_severity"],
        "| breaches=",
        policy["breach_count"],
    )

    if args.fail_on_policy_breach and policy["status"] == "breach":
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
