#!/usr/bin/env python3
"""Build nightly batch throughput trend snapshots and alert-hook payloads."""

from __future__ import annotations

import argparse
import json
import os
import pathlib
import sys
from datetime import datetime, timezone


SNAPSHOT_SCHEMA_ID = (
    "https://rootcellar.dev/schemas/artifacts/v1/batch-throughput-snapshot.schema.json"
)
ARTIFACT_SCHEMA_VERSION = "1.0.0"
ARTIFACT_COMPATIBILITY_MODE = "backward-additive"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Derive a nightly batch trend snapshot and alert-hook payload from a "
            "batch recalc report."
        )
    )
    parser.add_argument(
        "--report",
        required=True,
        help="Path to batch recalc report JSON.",
    )
    parser.add_argument(
        "--snapshot",
        default="./ci-batch-throughput-snapshot.json",
        help="Output path for trend snapshot JSON.",
    )
    parser.add_argument(
        "--alert",
        default="./ci-batch-alert-hook.json",
        help="Output path for alert-hook payload JSON.",
    )
    parser.add_argument(
        "--min-processed-files",
        type=int,
        required=True,
        help="Minimum processed files threshold.",
    )
    parser.add_argument(
        "--min-throughput-files-per-sec",
        type=float,
        required=True,
        help="Minimum throughput files/sec threshold.",
    )
    parser.add_argument(
        "--fail-on-breach",
        action="store_true",
        help="Exit non-zero when any threshold breach is detected.",
    )
    parser.add_argument(
        "--source",
        default="rootcellar.batch-nightly",
        help="Alert source identifier.",
    )
    parser.add_argument(
        "--allow-missing-report",
        action="store_true",
        help="If report is missing, emit a breach snapshot instead of failing immediately.",
    )
    return parser.parse_args()


def _read_json(path: pathlib.Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def _write_json(path: pathlib.Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _env(key: str) -> str | None:
    return os.environ.get(key)


def build_snapshot(report: dict, args: argparse.Namespace) -> dict:
    summary = report.get("summary", {})
    processed = int(summary.get("processed_files", 0))
    failures = int(summary.get("failure_count", 0))
    throughput = float(summary.get("throughput_files_per_sec", 0.0))
    effective_threads = int(summary.get("effective_threads", 0))
    wall_clock_ms = int(summary.get("wall_clock_duration_ms", 0))
    ratio = float(summary.get("aggregate_file_time_ratio", 0.0))

    breaches: list[dict] = []
    if processed < args.min_processed_files:
        breaches.append(
            {
                "metric": "processed_files",
                "actual": processed,
                "threshold": args.min_processed_files,
                "operator": ">=",
                "reason": "batch corpus coverage dropped below minimum threshold",
            }
        )
    if failures > 0:
        breaches.append(
            {
                "metric": "failure_count",
                "actual": failures,
                "threshold": 0,
                "operator": "==",
                "reason": "batch recalc reported workbook failures",
            }
        )
    if throughput < args.min_throughput_files_per_sec:
        breaches.append(
            {
                "metric": "throughput_files_per_sec",
                "actual": throughput,
                "threshold": args.min_throughput_files_per_sec,
                "operator": ">=",
                "reason": "batch throughput below regression threshold",
            }
        )

    status = "pass" if not breaches else "breach"
    now = datetime.now(timezone.utc).isoformat()

    return {
        "artifact_contract": {
            "schema_id": SNAPSHOT_SCHEMA_ID,
            "schema_version": ARTIFACT_SCHEMA_VERSION,
            "compatibility": ARTIFACT_COMPATIBILITY_MODE,
        },
        "snapshot_version": 1,
        "generated_at": now,
        "status": status,
        "source_report": str(pathlib.Path(args.report)).replace("\\", "/"),
        "workflow": {
            "name": _env("GITHUB_WORKFLOW"),
            "repository": _env("GITHUB_REPOSITORY"),
            "ref": _env("GITHUB_REF"),
            "sha": _env("GITHUB_SHA"),
            "run_id": _env("GITHUB_RUN_ID"),
            "run_attempt": _env("GITHUB_RUN_ATTEMPT"),
        },
        "thresholds": {
            "min_processed_files": args.min_processed_files,
            "min_throughput_files_per_sec": args.min_throughput_files_per_sec,
            "max_failure_count": 0,
        },
        "metrics": {
            "processed_files": processed,
            "failure_count": failures,
            "throughput_files_per_sec": throughput,
            "effective_threads": effective_threads,
            "wall_clock_duration_ms": wall_clock_ms,
            "aggregate_file_time_ratio": ratio,
        },
        "breaches": breaches,
    }


def build_missing_report_snapshot(args: argparse.Namespace) -> dict:
    now = datetime.now(timezone.utc).isoformat()
    breach = {
        "metric": "report_available",
        "actual": 0,
        "threshold": 1,
        "operator": "==",
        "reason": "batch recalc report missing; nightly batch run did not produce expected artifact",
    }
    return {
        "artifact_contract": {
            "schema_id": SNAPSHOT_SCHEMA_ID,
            "schema_version": ARTIFACT_SCHEMA_VERSION,
            "compatibility": ARTIFACT_COMPATIBILITY_MODE,
        },
        "snapshot_version": 1,
        "generated_at": now,
        "status": "breach",
        "source_report": str(pathlib.Path(args.report)).replace("\\", "/"),
        "workflow": {
            "name": _env("GITHUB_WORKFLOW"),
            "repository": _env("GITHUB_REPOSITORY"),
            "ref": _env("GITHUB_REF"),
            "sha": _env("GITHUB_SHA"),
            "run_id": _env("GITHUB_RUN_ID"),
            "run_attempt": _env("GITHUB_RUN_ATTEMPT"),
        },
        "thresholds": {
            "min_processed_files": args.min_processed_files,
            "min_throughput_files_per_sec": args.min_throughput_files_per_sec,
            "max_failure_count": 0,
        },
        "metrics": {
            "processed_files": 0,
            "failure_count": 1,
            "throughput_files_per_sec": 0.0,
            "effective_threads": 0,
            "wall_clock_duration_ms": 0,
            "aggregate_file_time_ratio": 0.0,
        },
        "breaches": [breach],
    }


def build_alert_payload(snapshot: dict, args: argparse.Namespace) -> dict:
    status = snapshot["status"]
    breaches = snapshot.get("breaches", [])
    severity = "p3" if status == "breach" else "info"
    summary = (
        "Nightly batch throughput regression detected"
        if status == "breach"
        else "Nightly batch throughput within thresholds"
    )
    return {
        "alert_version": 1,
        "generated_at": snapshot["generated_at"],
        "source": args.source,
        "status": status,
        "severity": severity,
        "summary": summary,
        "routing_key": "rootcellar.batch.throughput",
        "workflow": snapshot.get("workflow", {}),
        "metrics": snapshot.get("metrics", {}),
        "thresholds": snapshot.get("thresholds", {}),
        "breach_count": len(breaches),
        "breaches": breaches,
    }


def main() -> int:
    args = parse_args()
    if args.min_processed_files <= 0:
        raise ValueError("--min-processed-files must be greater than zero")
    if args.min_throughput_files_per_sec <= 0:
        raise ValueError("--min-throughput-files-per-sec must be greater than zero")

    report_path = pathlib.Path(args.report)
    if not report_path.exists():
        if not args.allow_missing_report:
            raise FileNotFoundError(f"report file not found: {args.report}")
        snapshot = build_missing_report_snapshot(args)
    else:
        report = _read_json(report_path)
        snapshot = build_snapshot(report, args)
    alert = build_alert_payload(snapshot, args)
    _write_json(pathlib.Path(args.snapshot), snapshot)
    _write_json(pathlib.Path(args.alert), alert)

    print(f"Wrote trend snapshot: {args.snapshot}")
    print(f"Wrote alert payload: {args.alert}")
    print(
        "Snapshot status:",
        snapshot["status"],
        "| processed_files=",
        snapshot["metrics"]["processed_files"],
        "| throughput_files_per_sec=",
        snapshot["metrics"]["throughput_files_per_sec"],
    )
    if snapshot["breaches"]:
        for breach in snapshot["breaches"]:
            print(
                " - breach:",
                breach["metric"],
                breach["actual"],
                breach["operator"],
                breach["threshold"],
            )

    if args.fail_on_breach and snapshot["status"] == "breach":
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
