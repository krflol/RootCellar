#!/usr/bin/env python3
"""Build policy-owner escalation metadata and downstream adapter exports."""

from __future__ import annotations

import argparse
import json
import pathlib
import sys
from collections import Counter
from datetime import datetime, timezone


SEVERITY_ORDER = {
    "info": 0,
    "p3": 1,
    "p2": 2,
    "p1": 3,
}

ESCALATION_SCHEMA_ID = (
    "https://rootcellar.dev/schemas/artifacts/v1/batch-policy-escalation.schema.json"
)
ADAPTER_EXPORTS_SCHEMA_ID = (
    "https://rootcellar.dev/schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json"
)
ARTIFACT_SCHEMA_VERSION = "1.0.0"
ARTIFACT_COMPATIBILITY_MODE = "backward-additive"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Build policy-owner escalation metadata and dashboard/incident adapter "
            "exports from nightly policy/dashboard artifacts."
        )
    )
    parser.add_argument(
        "--policy",
        required=True,
        help="Path to nightly alert policy artifact JSON.",
    )
    parser.add_argument(
        "--dashboard-pack",
        required=True,
        help="Path to nightly dashboard-pack artifact JSON.",
    )
    parser.add_argument(
        "--escalation",
        default="./ci-batch-policy-escalation.json",
        help="Output path for policy-to-owner escalation metadata JSON.",
    )
    parser.add_argument(
        "--adapter-exports",
        default="./ci-batch-dashboard-adapter-exports.json",
        help="Output path for downstream adapter export JSON.",
    )
    parser.add_argument(
        "--owner-team-default",
        default="rootcellar-core",
        help="Default owner team when source-specific routing is unavailable.",
    )
    parser.add_argument(
        "--owner-team-snapshot",
        default="rootcellar-performance",
        help="Owner team for snapshot-origin policy checks.",
    )
    parser.add_argument(
        "--owner-team-dispatch",
        default="rootcellar-reliability",
        help="Owner team for dispatch-origin policy checks.",
    )
    parser.add_argument(
        "--owner-team-ack-retention",
        default="rootcellar-observability",
        help="Owner team for ack-retention-origin policy checks.",
    )
    parser.add_argument(
        "--owner-team-policy",
        default="rootcellar-observability",
        help="Owner team for policy self-check origins.",
    )
    parser.add_argument(
        "--owner-contact-channel",
        default="#rootcellar-oncall",
        help="Default owner contact channel for escalation records.",
    )
    parser.add_argument(
        "--escalation-target-p1",
        default="pagerduty://rootcellar-critical",
        help="Escalation target for p1 policy breaches.",
    )
    parser.add_argument(
        "--escalation-target-p2",
        default="pagerduty://rootcellar-high",
        help="Escalation target for p2 policy breaches.",
    )
    parser.add_argument(
        "--escalation-target-p3",
        default="slack://rootcellar-alerts",
        help="Escalation target for p3 policy breaches.",
    )
    parser.add_argument(
        "--escalation-target-info",
        default="slack://rootcellar-observability",
        help="Escalation target for info-level policy checks.",
    )
    parser.add_argument(
        "--escalation-sla-minutes-p1",
        type=int,
        default=15,
        help="Response SLA in minutes for p1 breaches.",
    )
    parser.add_argument(
        "--escalation-sla-minutes-p2",
        type=int,
        default=60,
        help="Response SLA in minutes for p2 breaches.",
    )
    parser.add_argument(
        "--escalation-sla-minutes-p3",
        type=int,
        default=240,
        help="Response SLA in minutes for p3 breaches.",
    )
    parser.add_argument(
        "--escalation-sla-minutes-info",
        type=int,
        default=1440,
        help="Response SLA in minutes for info checks.",
    )
    parser.add_argument(
        "--allow-missing-inputs",
        action="store_true",
        help="Emit degraded outputs when policy/dashboard inputs are missing.",
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
                    "message": "input file missing; generated degraded adapter metadata",
                }
            )
            return None
        raise FileNotFoundError(f"{input_kind} file not found: {path_str}")
    return _read_json(path)


def _severity_rank(severity: str) -> int:
    return SEVERITY_ORDER.get(severity.lower(), -1)


def _highest_severity(checks: list[dict]) -> str:
    highest = "info"
    for check in checks:
        severity = str(check.get("severity", "info")).lower()
        if _severity_rank(severity) > _severity_rank(highest):
            highest = severity
    return highest


def _owner_team_for_source(source: str, args: argparse.Namespace) -> str:
    normalized = source.strip().lower()
    if normalized == "snapshot":
        return args.owner_team_snapshot
    if normalized == "dispatch":
        return args.owner_team_dispatch
    if normalized == "ack_retention_index":
        return args.owner_team_ack_retention
    if normalized == "policy":
        return args.owner_team_policy
    return args.owner_team_default


def _target_for_severity(severity: str, args: argparse.Namespace) -> tuple[str, int]:
    normalized = severity.strip().lower()
    if normalized == "p1":
        return args.escalation_target_p1, args.escalation_sla_minutes_p1
    if normalized == "p2":
        return args.escalation_target_p2, args.escalation_sla_minutes_p2
    if normalized == "p3":
        return args.escalation_target_p3, args.escalation_sla_minutes_p3
    return args.escalation_target_info, args.escalation_sla_minutes_info


def _owner_record(source: str, args: argparse.Namespace) -> dict:
    team = _owner_team_for_source(source, args)
    queue = team.replace(" ", "-").lower()
    return {
        "team": team,
        "queue": f"{queue}-oncall",
        "contact_channel": args.owner_contact_channel,
    }


def _policy_checks(policy: dict | None) -> list[dict]:
    checks = (policy or {}).get("checks", [])
    if not isinstance(checks, list):
        return []
    return [check for check in checks if isinstance(check, dict)]


def build_escalation(policy: dict | None, dashboard_pack: dict | None, args: argparse.Namespace, issues: list[dict]) -> dict:
    checks = _policy_checks(policy)
    now = datetime.now(timezone.utc).isoformat()

    assignments: list[dict] = []
    breach_assignments: list[dict] = []
    team_counter: Counter[str] = Counter()

    for check in checks:
        source = str(check.get("source", "unknown"))
        severity = str(check.get("severity", "info")).lower()
        status = str(check.get("status", "pass")).lower()
        owner = _owner_record(source, args)
        target, sla_minutes = _target_for_severity(severity, args)
        assignment = {
            "check_id": check.get("id"),
            "title": check.get("title"),
            "status": status,
            "severity": severity,
            "source": source,
            "message": check.get("message"),
            "observed": check.get("observed"),
            "expected": check.get("expected"),
            "owner": owner,
            "escalation": {
                "target": target,
                "response_sla_minutes": sla_minutes,
                "priority": severity,
            },
        }
        assignments.append(assignment)
        if status == "breach":
            breach_assignments.append(assignment)
            team_counter[owner["team"]] += 1

    if policy is None:
        issues.append(
            {
                "code": "policy_missing",
                "message": "policy input missing; escalation assignments are empty",
            }
        )
    if dashboard_pack is None:
        issues.append(
            {
                "code": "dashboard_pack_missing",
                "message": "dashboard pack input missing; escalation references may be incomplete",
            }
        )

    unique_teams = sorted({a["owner"]["team"] for a in assignments})
    highest_severity = (
        _highest_severity(breach_assignments) if breach_assignments else str((policy or {}).get("highest_severity", "info"))
    )

    return {
        "artifact_contract": {
            "schema_id": ESCALATION_SCHEMA_ID,
            "schema_version": ARTIFACT_SCHEMA_VERSION,
            "compatibility": ARTIFACT_COMPATIBILITY_MODE,
        },
        "escalation_version": 1,
        "generated_at": now,
        "policy_status": str((policy or {}).get("status", "missing")),
        "highest_severity": highest_severity,
        "routing_key": "rootcellar.batch.throughput.policy",
        "owner_policy": {
            "owner_team_default": args.owner_team_default,
            "owner_team_snapshot": args.owner_team_snapshot,
            "owner_team_dispatch": args.owner_team_dispatch,
            "owner_team_ack_retention": args.owner_team_ack_retention,
            "owner_team_policy": args.owner_team_policy,
            "owner_contact_channel": args.owner_contact_channel,
            "target_p1": args.escalation_target_p1,
            "target_p2": args.escalation_target_p2,
            "target_p3": args.escalation_target_p3,
            "target_info": args.escalation_target_info,
        },
        "summary": {
            "check_count": len(assignments),
            "breach_count": len(breach_assignments),
            "unique_owner_team_count": len(unique_teams),
            "unique_owner_teams": unique_teams,
            "breach_count_by_team": dict(sorted(team_counter.items())),
            "assigned_breach_count": len(breach_assignments),
            "unassigned_breach_count": 0,
        },
        "check_assignments": assignments,
        "breach_assignments": breach_assignments,
        "artifact_refs": {
            "policy": str(pathlib.Path(args.policy)).replace("\\", "/"),
            "dashboard_pack": str(pathlib.Path(args.dashboard_pack)).replace("\\", "/"),
        },
        "issues": issues,
    }


def _headline_metrics(dashboard_pack: dict | None) -> dict:
    headline = (dashboard_pack or {}).get("headline_metrics", {})
    if not isinstance(headline, dict):
        return {}
    return headline


def _policy_breaches(policy: dict | None) -> list[dict]:
    checks = _policy_checks(policy)
    breaches: list[dict] = []
    for check in checks:
        if str(check.get("status", "")).lower() != "breach":
            continue
        breaches.append(
            {
                "check_id": check.get("id"),
                "title": check.get("title"),
                "severity": check.get("severity"),
                "source": check.get("source"),
                "message": check.get("message"),
                "observed": check.get("observed"),
                "expected": check.get("expected"),
            }
        )
    return breaches


def build_adapter_exports(
    policy: dict | None,
    dashboard_pack: dict | None,
    escalation: dict,
    args: argparse.Namespace,
    issues: list[dict],
) -> dict:
    now = datetime.now(timezone.utc).isoformat()
    headline = _headline_metrics(dashboard_pack)
    breaches = _policy_breaches(policy)
    breach_assignments = escalation.get("breach_assignments", [])
    if not isinstance(breach_assignments, list):
        breach_assignments = []

    owner_targets = sorted(
        {
            str(assignment.get("escalation", {}).get("target"))
            for assignment in breach_assignments
            if isinstance(assignment, dict)
        }
    )
    owner_teams = sorted(
        {
            str(assignment.get("owner", {}).get("team"))
            for assignment in breach_assignments
            if isinstance(assignment, dict)
        }
    )
    policy_status = str((policy or {}).get("status", "missing"))
    highest_severity = str((policy or {}).get("highest_severity", "info"))
    breach_count = int((policy or {}).get("breach_count", len(breaches)))
    summary = (
        "Nightly batch policy breach detected"
        if policy_status == "breach"
        else "Nightly batch policy checks passing"
    )

    dashboard_points = [
        {
            "metric": "processed_files",
            "value": headline.get("processed_files"),
        },
        {
            "metric": "failure_count",
            "value": headline.get("failure_count"),
        },
        {
            "metric": "throughput_files_per_sec",
            "value": headline.get("throughput_files_per_sec"),
        },
        {
            "metric": "dispatch_failed_route_count",
            "value": headline.get("dispatch_failed_route_count"),
        },
        {
            "metric": "policy_breach_count",
            "value": headline.get("policy_breach_count"),
        },
    ]

    return {
        "artifact_contract": {
            "schema_id": ADAPTER_EXPORTS_SCHEMA_ID,
            "schema_version": ARTIFACT_SCHEMA_VERSION,
            "compatibility": ARTIFACT_COMPATIBILITY_MODE,
        },
        "adapter_exports_version": 1,
        "generated_at": now,
        "status": policy_status,
        "highest_severity": highest_severity,
        "incident_adapter": {
            "event_type": "rootcellar.batch.alert_policy",
            "routing_key": "rootcellar.batch.throughput.policy",
            "status": policy_status,
            "severity": highest_severity,
            "summary": summary,
            "breach_count": breach_count,
            "owner_teams": owner_teams,
            "owner_targets": owner_targets,
            "breaches": breaches,
            "escalations": breach_assignments,
            "artifact_refs": {
                "policy": str(pathlib.Path(args.policy)).replace("\\", "/"),
                "dashboard_pack": str(pathlib.Path(args.dashboard_pack)).replace("\\", "/"),
                "escalation": str(pathlib.Path(args.escalation)).replace("\\", "/"),
            },
        },
        "dashboard_adapter": {
            "dataset": "rootcellar.batch.nightly",
            "status": policy_status,
            "severity": highest_severity,
            "headline_metrics": headline,
            "policy_checks": _policy_checks(policy),
            "metric_points": dashboard_points,
            "dimensions": {
                "routing_key": "rootcellar.batch.throughput.policy",
                "owner_team_default": args.owner_team_default,
            },
            "artifact_refs": {
                "dashboard_pack": str(pathlib.Path(args.dashboard_pack)).replace("\\", "/"),
                "policy": str(pathlib.Path(args.policy)).replace("\\", "/"),
            },
        },
        "issues": issues,
    }


def main() -> int:
    args = parse_args()
    sla_values = [
        args.escalation_sla_minutes_p1,
        args.escalation_sla_minutes_p2,
        args.escalation_sla_minutes_p3,
        args.escalation_sla_minutes_info,
    ]
    if any(value <= 0 for value in sla_values):
        raise ValueError("all escalation SLA minute values must be greater than zero")

    issues: list[dict] = []
    policy = _load_input(args.policy, args.allow_missing_inputs, issues, "policy")
    dashboard_pack = _load_input(
        args.dashboard_pack, args.allow_missing_inputs, issues, "dashboard_pack"
    )

    escalation = build_escalation(policy, dashboard_pack, args, list(issues))
    adapter_exports = build_adapter_exports(policy, dashboard_pack, escalation, args, list(issues))

    _write_json(pathlib.Path(args.escalation), escalation)
    _write_json(pathlib.Path(args.adapter_exports), adapter_exports)

    print(f"Wrote escalation metadata: {args.escalation}")
    print(f"Wrote adapter exports: {args.adapter_exports}")
    print(
        "Escalation status:",
        escalation.get("policy_status"),
        "| breaches=",
        escalation.get("summary", {}).get("breach_count"),
        "| owner_teams=",
        escalation.get("summary", {}).get("unique_owner_team_count"),
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
