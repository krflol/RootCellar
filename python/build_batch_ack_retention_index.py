#!/usr/bin/env python3
"""Build acknowledgement-retention lookup index from dispatch report artifacts."""

from __future__ import annotations

import argparse
import hashlib
import json
import pathlib
import sys
from datetime import datetime, timedelta, timezone


ACK_RETENTION_SCHEMA_ID = (
    "https://rootcellar.dev/schemas/artifacts/v1/batch-ack-retention-index.schema.json"
)
ARTIFACT_SCHEMA_VERSION = "1.0.0"
ARTIFACT_COMPATIBILITY_MODE = "backward-additive"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Derive a forensic acknowledgement-retention index from a nightly "
            "batch alert dispatch report."
        )
    )
    parser.add_argument(
        "--dispatch-report",
        required=True,
        help="Path to alert dispatch report JSON.",
    )
    parser.add_argument(
        "--index",
        default="./ci-batch-ack-retention-index.json",
        help="Output path for acknowledgement-retention index JSON.",
    )
    parser.add_argument(
        "--retention-days",
        type=int,
        default=30,
        help="Retention window in days for acknowledgement lookup records.",
    )
    parser.add_argument(
        "--allow-missing-dispatch-report",
        action="store_true",
        help="Emit a degraded index artifact when dispatch report is missing.",
    )
    return parser.parse_args()


def _read_json(path: pathlib.Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def _write_json(path: pathlib.Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _parse_iso8601(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)
    except ValueError:
        return None


def _sha256_text(value: str | None) -> str | None:
    if value is None:
        return None
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _retention_timestamp(anchor: datetime, retention_days: int) -> str:
    return (anchor + timedelta(days=retention_days)).astimezone(timezone.utc).isoformat()


def _build_missing_report_index(args: argparse.Namespace) -> dict:
    now = datetime.now(timezone.utc)
    return {
        "artifact_contract": {
            "schema_id": ACK_RETENTION_SCHEMA_ID,
            "schema_version": ARTIFACT_SCHEMA_VERSION,
            "compatibility": ARTIFACT_COMPATIBILITY_MODE,
        },
        "index_version": 1,
        "generated_at": now.isoformat(),
        "status": "missing_dispatch_report",
        "source_dispatch_report": str(pathlib.Path(args.dispatch_report)).replace("\\", "/"),
        "retention_policy": {
            "retention_days": args.retention_days,
            "anchor": "index_generated_at",
        },
        "dispatch_summary": {
            "dispatch_status": "missing",
            "route_count": 0,
            "configured_route_count": 0,
            "delivered_route_count": 0,
            "failed_route_count": 0,
            "ack_required_route_count": 0,
            "ack_received_route_count": 0,
            "ack_missing_route_count": 0,
        },
        "record_count": 0,
        "lookup_count": {
            "ack_id": 0,
            "ack_id_sha256": 0,
            "idempotency_key": 0,
            "correlation_id": 0,
        },
        "records": [],
        "lookups": {
            "by_ack_id": [],
            "by_ack_id_sha256": [],
            "by_idempotency_key": [],
            "by_correlation_id": [],
        },
        "issues": [
            {
                "code": "dispatch_report_missing",
                "message": "dispatch report not found; produced degraded retention index",
            }
        ],
    }


def build_index(dispatch_report: dict, args: argparse.Namespace) -> dict:
    generated_at = datetime.now(timezone.utc)
    dispatch_generated_at = _parse_iso8601(str(dispatch_report.get("generated_at", "")))
    anchor = dispatch_generated_at or generated_at

    records: list[dict] = []
    lookup_by_ack_id: list[dict] = []
    lookup_by_ack_id_sha256: list[dict] = []
    lookup_by_idempotency_key: list[dict] = []
    lookup_by_correlation_id: list[dict] = []

    routes = dispatch_report.get("routes", [])
    if not isinstance(routes, list):
        routes = []

    for route in routes:
        if not isinstance(route, dict):
            continue
        identifiers = route.get("identifiers", {})
        if not isinstance(identifiers, dict):
            identifiers = {}
        ack = route.get("ack", {})
        if not isinstance(ack, dict):
            ack = {}
        correlation = route.get("correlation", {})
        if not isinstance(correlation, dict):
            correlation = {}
        delivery = route.get("delivery", {})
        if not isinstance(delivery, dict):
            delivery = {}
        final = delivery.get("final", {})
        if not isinstance(final, dict):
            final = {}
        replay = final.get("replay", {})
        if not isinstance(replay, dict):
            replay = {}

        ack_id = ack.get("ack_id")
        ack_id_text = str(ack_id) if ack_id is not None else None
        ack_id_sha256 = _sha256_text(ack_id_text)
        idempotency_key = identifiers.get("idempotency_key")
        correlation_id = identifiers.get("correlation_id")
        retention_expires_at = _retention_timestamp(anchor, args.retention_days)

        record = {
            "route": route.get("route"),
            "route_status": route.get("status"),
            "configured": route.get("configured"),
            "endpoint": route.get("endpoint"),
            "dispatch_generated_at": anchor.isoformat(),
            "retention_expires_at": retention_expires_at,
            "ack": {
                "field": ack.get("field"),
                "required": bool(ack.get("required")),
                "received": bool(ack.get("received")),
                "ack_id": ack_id_text,
                "ack_id_sha256": ack_id_sha256,
                "parse_error": ack.get("parse_error"),
            },
            "identifiers": {
                "idempotency_key": idempotency_key,
                "correlation_id": correlation_id,
                "digest_sha256": identifiers.get("digest_sha256"),
            },
            "correlation": {
                "field": correlation.get("field"),
                "required": bool(correlation.get("required")),
                "received": bool(correlation.get("received")),
                "matches": bool(correlation.get("matches")),
                "expected": correlation.get("expected"),
                "actual": correlation.get("actual"),
            },
            "delivery": {
                "attempt_count": int(delivery.get("attempt_count", 0)),
                "final_http_status": final.get("http_status"),
                "final_duration_ms": final.get("duration_ms"),
                "final_replay_timestamp": replay.get("timestamp"),
                "final_replay_nonce": replay.get("nonce"),
                "final_replay_window_sec": replay.get("window_sec"),
            },
        }
        record_id = len(records)
        records.append(record)

        if ack_id_text:
            lookup_by_ack_id.append(
                {
                    "value": ack_id_text,
                    "record_id": record_id,
                    "route": record["route"],
                    "route_status": record["route_status"],
                }
            )
        if ack_id_sha256:
            lookup_by_ack_id_sha256.append(
                {
                    "value": ack_id_sha256,
                    "record_id": record_id,
                    "route": record["route"],
                    "route_status": record["route_status"],
                }
            )
        if idempotency_key:
            lookup_by_idempotency_key.append(
                {
                    "value": idempotency_key,
                    "record_id": record_id,
                    "route": record["route"],
                    "route_status": record["route_status"],
                }
            )
        if correlation_id:
            lookup_by_correlation_id.append(
                {
                    "value": correlation_id,
                    "record_id": record_id,
                    "route": record["route"],
                    "route_status": record["route_status"],
                }
            )

    unique_ack_id = {entry["value"] for entry in lookup_by_ack_id}
    unique_ack_id_sha256 = {entry["value"] for entry in lookup_by_ack_id_sha256}
    unique_idempotency = {entry["value"] for entry in lookup_by_idempotency_key}
    unique_correlation = {entry["value"] for entry in lookup_by_correlation_id}

    status = "ok"
    dispatch_status = str(dispatch_report.get("status", ""))
    if dispatch_status == "no_routes":
        status = "no_routes"
    elif not records:
        status = "no_routes"
    elif not lookup_by_ack_id:
        status = "no_ack_records"

    return {
        "artifact_contract": {
            "schema_id": ACK_RETENTION_SCHEMA_ID,
            "schema_version": ARTIFACT_SCHEMA_VERSION,
            "compatibility": ARTIFACT_COMPATIBILITY_MODE,
        },
        "index_version": 1,
        "generated_at": generated_at.isoformat(),
        "status": status,
        "source_dispatch_report": str(pathlib.Path(args.dispatch_report)).replace("\\", "/"),
        "retention_policy": {
            "retention_days": args.retention_days,
            "anchor": "dispatch_generated_at",
            "anchor_timestamp": anchor.isoformat(),
        },
        "dispatch_summary": {
            "dispatch_status": dispatch_status,
            "route_count": int(dispatch_report.get("route_count", 0)),
            "configured_route_count": int(dispatch_report.get("configured_route_count", 0)),
            "delivered_route_count": int(dispatch_report.get("delivered_route_count", 0)),
            "failed_route_count": int(dispatch_report.get("failed_route_count", 0)),
            "ack_required_route_count": int(dispatch_report.get("ack_required_route_count", 0)),
            "ack_received_route_count": int(dispatch_report.get("ack_received_route_count", 0)),
            "ack_missing_route_count": int(dispatch_report.get("ack_missing_route_count", 0)),
            "correlation_required_route_count": int(
                dispatch_report.get("correlation_required_route_count", 0)
            ),
            "correlation_matched_route_count": int(
                dispatch_report.get("correlation_matched_route_count", 0)
            ),
            "correlation_mismatch_route_count": int(
                dispatch_report.get("correlation_mismatch_route_count", 0)
            ),
        },
        "record_count": len(records),
        "lookup_count": {
            "ack_id": len(unique_ack_id),
            "ack_id_sha256": len(unique_ack_id_sha256),
            "idempotency_key": len(unique_idempotency),
            "correlation_id": len(unique_correlation),
        },
        "records": records,
        "lookups": {
            "by_ack_id": lookup_by_ack_id,
            "by_ack_id_sha256": lookup_by_ack_id_sha256,
            "by_idempotency_key": lookup_by_idempotency_key,
            "by_correlation_id": lookup_by_correlation_id,
        },
        "issues": [],
    }


def main() -> int:
    args = parse_args()
    if args.retention_days <= 0:
        raise ValueError("--retention-days must be greater than zero")

    dispatch_path = pathlib.Path(args.dispatch_report)
    if not dispatch_path.exists():
        if not args.allow_missing_dispatch_report:
            raise FileNotFoundError(f"dispatch report file not found: {args.dispatch_report}")
        index = _build_missing_report_index(args)
    else:
        dispatch_report = _read_json(dispatch_path)
        index = build_index(dispatch_report, args)

    _write_json(pathlib.Path(args.index), index)
    print(f"Wrote acknowledgement retention index: {args.index}")
    print(
        "Index status:",
        index["status"],
        "| records=",
        index["record_count"],
        "| ack_lookup_count=",
        index["lookup_count"]["ack_id"],
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
