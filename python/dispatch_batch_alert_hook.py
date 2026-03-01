#!/usr/bin/env python3
"""Dispatch nightly batch alert payloads to incident/dashboard ingestion endpoints."""

from __future__ import annotations

import argparse
import copy
import hashlib
import hmac
import json
import os
import pathlib
import secrets
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from typing import Callable


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Dispatch a nightly batch alert payload to configured incident and "
            "dashboard ingestion endpoints."
        )
    )
    parser.add_argument(
        "--alert",
        required=True,
        help="Path to alert-hook payload JSON.",
    )
    parser.add_argument(
        "--snapshot",
        default=None,
        help="Optional path to throughput snapshot JSON for dashboard payload enrichment.",
    )
    parser.add_argument(
        "--incident-url",
        default=None,
        help="Incident ingestion webhook URL. Falls back to ROOTCELLAR_INCIDENT_WEBHOOK_URL.",
    )
    parser.add_argument(
        "--dashboard-url",
        default=None,
        help="Dashboard ingestion URL. Falls back to ROOTCELLAR_DASHBOARD_INGEST_URL.",
    )
    parser.add_argument(
        "--incident-auth-token",
        default=None,
        help="Incident bearer token. Falls back to ROOTCELLAR_INCIDENT_WEBHOOK_TOKEN.",
    )
    parser.add_argument(
        "--dashboard-auth-token",
        default=None,
        help="Dashboard bearer token. Falls back to ROOTCELLAR_DASHBOARD_INGEST_TOKEN.",
    )
    parser.add_argument(
        "--auth-scheme",
        default="Bearer",
        help="Authorization scheme prefix for token auth headers.",
    )
    parser.add_argument(
        "--signing-secret",
        default=None,
        help="Optional HMAC signing secret. Falls back to ROOTCELLAR_ALERT_SIGNING_SECRET.",
    )
    parser.add_argument(
        "--signature-header",
        default="X-RootCellar-Signature",
        help="Header name for optional HMAC signature.",
    )
    parser.add_argument(
        "--idempotency-header",
        default="Idempotency-Key",
        help="Header name for idempotency key propagation.",
    )
    parser.add_argument(
        "--idempotency-key-prefix",
        default="rootcellar-batch-alert",
        help="Prefix used when generating deterministic route idempotency keys.",
    )
    parser.add_argument(
        "--correlation-header",
        default="X-Correlation-Id",
        help="Header name for correlation-id propagation.",
    )
    parser.add_argument(
        "--correlation-field",
        default="correlation_id",
        help="JSON field used to read correlation id from downstream response payload.",
    )
    parser.add_argument(
        "--replay-timestamp-header",
        default="X-RootCellar-Timestamp",
        help="Header name for replay-protection timestamp propagation.",
    )
    parser.add_argument(
        "--replay-nonce-header",
        default="X-RootCellar-Nonce",
        help="Header name for replay-protection nonce propagation.",
    )
    parser.add_argument(
        "--replay-window-header",
        default="X-RootCellar-Replay-Window-Sec",
        help="Header name for replay window policy propagation.",
    )
    parser.add_argument(
        "--replay-window-sec",
        type=int,
        default=300,
        help="Replay-protection acceptance window in seconds.",
    )
    parser.add_argument(
        "--ack-field",
        default="ack_id",
        help="JSON field name used for acknowledgement extraction from route responses.",
    )
    parser.add_argument(
        "--require-ack-on-incident",
        action="store_true",
        help="Mark incident route delivery as failed when ack field is absent.",
    )
    parser.add_argument(
        "--require-ack-on-dashboard",
        action="store_true",
        help="Mark dashboard route delivery as failed when ack field is absent.",
    )
    parser.add_argument(
        "--require-correlation-on-incident",
        action="store_true",
        help="Mark incident route delivery as failed when response correlation id is missing or mismatched.",
    )
    parser.add_argument(
        "--require-correlation-on-dashboard",
        action="store_true",
        help="Mark dashboard route delivery as failed when response correlation id is missing or mismatched.",
    )
    parser.add_argument(
        "--dispatch-report",
        default="./ci-batch-alert-dispatch.json",
        help="Output path for dispatch status report JSON.",
    )
    parser.add_argument(
        "--timeout-sec",
        type=float,
        default=10.0,
        help="HTTP request timeout in seconds.",
    )
    parser.add_argument(
        "--max-attempts",
        type=int,
        default=3,
        help="Maximum delivery attempts per configured route.",
    )
    parser.add_argument(
        "--initial-backoff-sec",
        type=float,
        default=0.5,
        help="Initial retry backoff in seconds.",
    )
    parser.add_argument(
        "--backoff-multiplier",
        type=float,
        default=2.0,
        help="Retry backoff multiplier.",
    )
    parser.add_argument(
        "--max-backoff-sec",
        type=float,
        default=5.0,
        help="Maximum retry backoff in seconds.",
    )
    parser.add_argument(
        "--retry-status-codes",
        default="408,425,429,500,502,503,504",
        help="Comma-separated HTTP status codes considered retryable.",
    )
    parser.add_argument(
        "--fail-on-route-error",
        action="store_true",
        help="Exit non-zero if any configured route fails.",
    )
    parser.add_argument(
        "--incident-on-pass",
        action="store_true",
        help="Dispatch incident events even when alert status is pass.",
    )
    return parser.parse_args()


def _read_json(path: pathlib.Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def _write_json(path: pathlib.Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _sanitize_url(url: str) -> str:
    parsed = urllib.parse.urlparse(url)
    if not parsed.scheme:
        return "<invalid-url>"
    host = parsed.hostname or ""
    port = f":{parsed.port}" if parsed.port else ""
    path = parsed.path or "/"
    return f"{parsed.scheme}://{host}{port}{path}"


def _canonical_json_bytes(payload: dict) -> bytes:
    return json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")


def _parse_retry_status_codes(raw: str) -> set[int]:
    out: set[int] = set()
    for token in raw.split(","):
        token = token.strip()
        if not token:
            continue
        out.add(int(token))
    return out


def _build_headers(
    token: str | None,
    auth_scheme: str,
    signing_secret: str | None,
    signature_header: str,
    body: bytes,
    idempotency_header: str,
    idempotency_key: str,
    correlation_header: str,
    correlation_id: str,
    replay_timestamp_header: str,
    replay_timestamp: str,
    replay_nonce_header: str,
    replay_nonce: str,
    replay_window_header: str,
    replay_window_sec: int,
) -> dict[str, str]:
    headers = {
        "Content-Type": "application/json",
    }
    headers[idempotency_header] = idempotency_key
    headers[correlation_header] = correlation_id
    headers[replay_timestamp_header] = replay_timestamp
    headers[replay_nonce_header] = replay_nonce
    headers[replay_window_header] = str(replay_window_sec)
    if token:
        scheme = auth_scheme.strip() or "Bearer"
        headers["Authorization"] = f"{scheme} {token}"
    if signing_secret:
        signing_material = (
            replay_timestamp.encode("utf-8")
            + b"\n"
            + replay_nonce.encode("utf-8")
            + b"\n"
            + str(replay_window_sec).encode("utf-8")
            + b"\n"
            + body
        )
        digest = hmac.new(
            signing_secret.encode("utf-8"), signing_material, hashlib.sha256
        ).hexdigest()
        headers[signature_header] = digest
        headers["X-RootCellar-Signature-Alg"] = "hmac-sha256"
    return headers


def _extract_ack(response_body: str, ack_field: str) -> tuple[bool, str | None, str | None]:
    if not response_body:
        return False, None, None
    try:
        payload = json.loads(response_body)
    except json.JSONDecodeError as exc:
        return False, None, str(exc)
    if isinstance(payload, dict) and ack_field in payload:
        value = payload.get(ack_field)
        if value is None:
            return False, None, None
        return True, str(value), None
    return False, None, None


def _extract_json_field(response_body: str, field: str) -> tuple[bool, str | None, str | None]:
    if not response_body:
        return False, None, None
    try:
        payload = json.loads(response_body)
    except json.JSONDecodeError as exc:
        return False, None, str(exc)
    if isinstance(payload, dict) and field in payload:
        value = payload.get(field)
        if value is None:
            return False, None, None
        return True, str(value), None
    return False, None, None


def _build_route_identifiers(
    alert: dict,
    route: str,
    idempotency_key_prefix: str,
) -> dict[str, str]:
    workflow = alert.get("workflow", {}) if isinstance(alert, dict) else {}
    run_id = workflow.get("run_id")
    run_attempt = workflow.get("run_attempt")
    base = {
        "route": route,
        "routing_key": alert.get("routing_key"),
        "alert_status": alert.get("status"),
        "alert_generated_at": alert.get("generated_at"),
        "workflow_run_id": run_id,
        "workflow_run_attempt": run_attempt,
    }
    digest = hashlib.sha256(_canonical_json_bytes(base)).hexdigest()
    correlation_id = f"rc-batch-{digest[:20]}"
    idempotency_key = f"{idempotency_key_prefix}-{route}-{digest[:24]}"
    return {
        "correlation_id": correlation_id,
        "idempotency_key": idempotency_key,
        "digest_sha256": digest,
    }


def _with_dispatch_envelope(
    payload: dict,
    route: str,
    correlation_id: str,
    idempotency_key: str,
) -> dict:
    enriched = copy.deepcopy(payload)
    enriched["_rootcellar_dispatch"] = {
        "route": route,
        "correlation_id": correlation_id,
        "idempotency_key": idempotency_key,
        "sent_at": datetime.now(timezone.utc).isoformat(),
    }
    return enriched


def _post_json_attempt(
    url: str,
    payload: dict,
    timeout_sec: float,
    headers: dict[str, str],
    retry_status_codes: set[int],
) -> tuple[bool, bool, dict]:
    data = _canonical_json_bytes(payload)
    request_headers = dict(headers)
    req = urllib.request.Request(
        url=url,
        data=data,
        method="POST",
        headers=request_headers,
    )
    started = time.perf_counter()
    try:
        with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
            body = resp.read(4096).decode("utf-8", errors="replace")
            duration_ms = int((time.perf_counter() - started) * 1000.0)
            return True, False, {
                "http_status": resp.status,
                "response_excerpt": body[:512],
                "response_body": body,
                "duration_ms": duration_ms,
                "retriable": False,
            }
    except urllib.error.HTTPError as exc:
        body = exc.read(4096).decode("utf-8", errors="replace")
        duration_ms = int((time.perf_counter() - started) * 1000.0)
        retriable = exc.code in retry_status_codes
        return False, retriable, {
            "http_status": exc.code,
            "error": str(exc),
            "response_excerpt": body[:512],
            "response_body": body,
            "duration_ms": duration_ms,
            "retriable": retriable,
        }
    except Exception as exc:  # noqa: BLE001
        duration_ms = int((time.perf_counter() - started) * 1000.0)
        return False, True, {
            "error": str(exc),
            "duration_ms": duration_ms,
            "retriable": True,
        }


def _post_json_with_retries(
    url: str,
    payload: dict,
    timeout_sec: float,
    header_factory: Callable[[int], tuple[dict[str, str], dict]],
    max_attempts: int,
    initial_backoff_sec: float,
    backoff_multiplier: float,
    max_backoff_sec: float,
    retry_status_codes: set[int],
) -> tuple[bool, dict]:
    attempts: list[dict] = []
    backoff_sec = initial_backoff_sec
    for attempt in range(1, max_attempts + 1):
        headers, replay_meta = header_factory(attempt)
        ok, retriable, detail = _post_json_attempt(
            url, payload, timeout_sec, headers, retry_status_codes
        )
        detail["replay"] = replay_meta
        detail["attempt"] = attempt
        detail["will_retry"] = bool(retriable and attempt < max_attempts and not ok)
        attempts.append(detail)
        if ok:
            break
        if not retriable or attempt >= max_attempts:
            break
        sleep_sec = max(0.0, min(backoff_sec, max_backoff_sec))
        detail["retry_sleep_sec"] = sleep_sec
        time.sleep(sleep_sec)
        backoff_sec = backoff_sec * backoff_multiplier

    final = attempts[-1] if attempts else {}
    return bool(final.get("http_status", 0) >= 200 and final.get("http_status", 0) < 300), {
        "attempt_count": len(attempts),
        "attempts": attempts,
        "final": final,
    }


def _route_record(  # noqa: PLR0913
    alert: dict,
    route: str,
    configured_url: str | None,
    enabled: bool,
    payload: dict,
    timeout_sec: float,
    auth_token: str | None,
    auth_scheme: str,
    signing_secret: str | None,
    signature_header: str,
    max_attempts: int,
    initial_backoff_sec: float,
    backoff_multiplier: float,
    max_backoff_sec: float,
    retry_status_codes: set[int],
    idempotency_header: str,
    idempotency_key_prefix: str,
    correlation_header: str,
    correlation_field: str,
    replay_timestamp_header: str,
    replay_nonce_header: str,
    replay_window_header: str,
    replay_window_sec: int,
    ack_field: str,
    ack_required: bool,
    correlation_required: bool,
) -> dict:
    identifiers = _build_route_identifiers(alert, route, idempotency_key_prefix)
    if not configured_url:
        return {
            "identifiers": identifiers,
            "route": route,
            "status": "skipped",
            "reason": "endpoint not configured",
            "configured": False,
        }
    if not enabled:
        return {
            "identifiers": identifiers,
            "route": route,
            "status": "skipped",
            "reason": "route disabled by alert state policy",
            "endpoint": _sanitize_url(configured_url),
            "configured": True,
        }

    enriched_payload = _with_dispatch_envelope(
        payload,
        route,
        identifiers["correlation_id"],
        identifiers["idempotency_key"],
    )
    request_body = _canonical_json_bytes(enriched_payload)

    def header_factory(attempt: int) -> tuple[dict[str, str], dict]:
        replay_timestamp = str(int(time.time()))
        nonce_seed = (
            f"{route}|{identifiers['idempotency_key']}|{identifiers['correlation_id']}|"
            f"{attempt}|{time.time_ns()}|{secrets.token_hex(8)}"
        )
        replay_nonce = hashlib.sha256(nonce_seed.encode("utf-8")).hexdigest()[:32]
        headers = _build_headers(
            auth_token,
            auth_scheme,
            signing_secret,
            signature_header,
            request_body,
            idempotency_header,
            identifiers["idempotency_key"],
            correlation_header,
            identifiers["correlation_id"],
            replay_timestamp_header,
            replay_timestamp,
            replay_nonce_header,
            replay_nonce,
            replay_window_header,
            replay_window_sec,
        )
        replay_meta = {
            "timestamp_header": replay_timestamp_header,
            "timestamp": replay_timestamp,
            "nonce_header": replay_nonce_header,
            "nonce": replay_nonce,
            "window_header": replay_window_header,
            "window_sec": replay_window_sec,
        }
        return headers, replay_meta

    ok, delivery = _post_json_with_retries(
        configured_url,
        enriched_payload,
        timeout_sec,
        header_factory,
        max_attempts,
        initial_backoff_sec,
        backoff_multiplier,
        max_backoff_sec,
        retry_status_codes,
    )
    final_detail = delivery.get("final", {})
    ack_received, ack_id, ack_parse_error = _extract_ack(
        str(final_detail.get("response_body", "")),
        ack_field,
    )
    ack = {
        "field": ack_field,
        "required": ack_required,
        "received": ack_received,
        "ack_id": ack_id,
        "parse_error": ack_parse_error,
    }
    corr_received, corr_id, corr_parse_error = _extract_json_field(
        str(final_detail.get("response_body", "")),
        correlation_field,
    )
    corr_matches = corr_received and corr_id == identifiers["correlation_id"]
    correlation = {
        "field": correlation_field,
        "required": correlation_required,
        "received": corr_received,
        "matches": corr_matches,
        "expected": identifiers["correlation_id"],
        "actual": corr_id,
        "parse_error": corr_parse_error,
    }
    status = "delivered" if ok else "failed"
    if ok and ack_required and not ack_received:
        status = "failed"
        ack["reason"] = "required acknowledgement missing"
    if ok and correlation_required and not corr_matches:
        status = "failed"
        correlation["reason"] = (
            "required correlation id missing or mismatched in downstream acknowledgement"
        )

    return {
        "identifiers": identifiers,
        "route": route,
        "status": status,
        "configured": True,
        "endpoint": _sanitize_url(configured_url),
        "auth": {
            "token_configured": bool(auth_token),
            "auth_scheme": auth_scheme.strip() or "Bearer",
            "signing_enabled": bool(signing_secret),
            "signature_header": signature_header if signing_secret else None,
        },
        "retry_policy": {
            "max_attempts": max_attempts,
            "initial_backoff_sec": initial_backoff_sec,
            "backoff_multiplier": backoff_multiplier,
            "max_backoff_sec": max_backoff_sec,
            "retry_status_codes": sorted(retry_status_codes),
        },
        "replay_policy": {
            "timestamp_header": replay_timestamp_header,
            "nonce_header": replay_nonce_header,
            "window_header": replay_window_header,
            "window_sec": replay_window_sec,
        },
        "delivery": delivery,
        "ack": ack,
        "correlation": correlation,
    }


def main() -> int:
    args = parse_args()
    if args.max_attempts <= 0:
        raise ValueError("--max-attempts must be greater than zero")
    if args.initial_backoff_sec < 0:
        raise ValueError("--initial-backoff-sec must be >= 0")
    if args.backoff_multiplier < 1.0:
        raise ValueError("--backoff-multiplier must be >= 1.0")
    if args.max_backoff_sec < 0:
        raise ValueError("--max-backoff-sec must be >= 0")
    if args.replay_window_sec <= 0:
        raise ValueError("--replay-window-sec must be greater than zero")
    retry_status_codes = _parse_retry_status_codes(args.retry_status_codes)

    alert = _read_json(pathlib.Path(args.alert))
    snapshot = None
    if args.snapshot:
        snapshot_path = pathlib.Path(args.snapshot)
        if snapshot_path.exists():
            snapshot = _read_json(snapshot_path)

    incident_url = args.incident_url or os.environ.get("ROOTCELLAR_INCIDENT_WEBHOOK_URL")
    dashboard_url = args.dashboard_url or os.environ.get("ROOTCELLAR_DASHBOARD_INGEST_URL")
    incident_auth_token = args.incident_auth_token or os.environ.get(
        "ROOTCELLAR_INCIDENT_WEBHOOK_TOKEN"
    )
    dashboard_auth_token = args.dashboard_auth_token or os.environ.get(
        "ROOTCELLAR_DASHBOARD_INGEST_TOKEN"
    )
    signing_secret = args.signing_secret or os.environ.get("ROOTCELLAR_ALERT_SIGNING_SECRET")
    alert_status = str(alert.get("status", "unknown")).lower()
    incident_enabled = args.incident_on_pass or alert_status == "breach"

    incident_payload = {
        "event_type": "rootcellar.batch.alert",
        "alert": alert,
    }
    dashboard_payload = {
        "event_type": "rootcellar.batch.throughput",
        "alert": alert,
        "snapshot": snapshot,
    }

    routes = [
        _route_record(
            alert,
            "incident",
            incident_url,
            incident_enabled,
            incident_payload,
            args.timeout_sec,
            incident_auth_token,
            args.auth_scheme,
            signing_secret,
            args.signature_header,
            args.max_attempts,
            args.initial_backoff_sec,
            args.backoff_multiplier,
            args.max_backoff_sec,
            retry_status_codes,
            args.idempotency_header,
            args.idempotency_key_prefix,
            args.correlation_header,
            args.correlation_field,
            args.replay_timestamp_header,
            args.replay_nonce_header,
            args.replay_window_header,
            args.replay_window_sec,
            args.ack_field,
            args.require_ack_on_incident,
            args.require_correlation_on_incident,
        ),
        _route_record(
            alert,
            "dashboard",
            dashboard_url,
            True,
            dashboard_payload,
            args.timeout_sec,
            dashboard_auth_token,
            args.auth_scheme,
            signing_secret,
            args.signature_header,
            args.max_attempts,
            args.initial_backoff_sec,
            args.backoff_multiplier,
            args.max_backoff_sec,
            retry_status_codes,
            args.idempotency_header,
            args.idempotency_key_prefix,
            args.correlation_header,
            args.correlation_field,
            args.replay_timestamp_header,
            args.replay_nonce_header,
            args.replay_window_header,
            args.replay_window_sec,
            args.ack_field,
            args.require_ack_on_dashboard,
            args.require_correlation_on_dashboard,
        ),
    ]

    configured_routes = [r for r in routes if r.get("configured")]
    failed_routes = [r for r in routes if r.get("status") == "failed"]
    delivered_routes = [r for r in routes if r.get("status") == "delivered"]
    ack_required_routes = [r for r in routes if r.get("ack", {}).get("required")]
    ack_received_routes = [r for r in routes if r.get("ack", {}).get("received")]
    ack_missing_routes = [
        r
        for r in routes
        if r.get("ack", {}).get("required") and not r.get("ack", {}).get("received")
    ]
    correlation_required_routes = [r for r in routes if r.get("correlation", {}).get("required")]
    correlation_matched_routes = [r for r in routes if r.get("correlation", {}).get("matches")]
    correlation_mismatched_routes = [
        r
        for r in routes
        if r.get("correlation", {}).get("required")
        and not r.get("correlation", {}).get("matches")
    ]
    if failed_routes:
        overall_status = "partial_failure"
    elif delivered_routes:
        overall_status = "delivered"
    else:
        overall_status = "no_routes"

    report = {
        "dispatch_version": 1,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "status": overall_status,
        "alert_status": alert_status,
        "alert_file": str(pathlib.Path(args.alert)).replace("\\", "/"),
        "snapshot_file": str(pathlib.Path(args.snapshot)).replace("\\", "/")
        if args.snapshot
        else None,
        "route_count": len(routes),
        "configured_route_count": len(configured_routes),
        "delivered_route_count": len(delivered_routes),
        "failed_route_count": len(failed_routes),
        "ack_required_route_count": len(ack_required_routes),
        "ack_received_route_count": len(ack_received_routes),
        "ack_missing_route_count": len(ack_missing_routes),
        "correlation_required_route_count": len(correlation_required_routes),
        "correlation_matched_route_count": len(correlation_matched_routes),
        "correlation_mismatch_route_count": len(correlation_mismatched_routes),
        "dispatch_policy": {
            "timeout_sec": args.timeout_sec,
            "max_attempts": args.max_attempts,
            "initial_backoff_sec": args.initial_backoff_sec,
            "backoff_multiplier": args.backoff_multiplier,
            "max_backoff_sec": args.max_backoff_sec,
            "retry_status_codes": sorted(retry_status_codes),
            "ack_field": args.ack_field,
            "require_ack_on_incident": args.require_ack_on_incident,
            "require_ack_on_dashboard": args.require_ack_on_dashboard,
            "correlation_field": args.correlation_field,
            "correlation_header": args.correlation_header,
            "require_correlation_on_incident": args.require_correlation_on_incident,
            "require_correlation_on_dashboard": args.require_correlation_on_dashboard,
            "idempotency_header": args.idempotency_header,
            "idempotency_key_prefix": args.idempotency_key_prefix,
            "replay_timestamp_header": args.replay_timestamp_header,
            "replay_nonce_header": args.replay_nonce_header,
            "replay_window_header": args.replay_window_header,
            "replay_window_sec": args.replay_window_sec,
            "auth_scheme": args.auth_scheme.strip() or "Bearer",
            "incident_auth_configured": bool(incident_auth_token),
            "dashboard_auth_configured": bool(dashboard_auth_token),
            "signing_enabled": bool(signing_secret),
            "signature_header": args.signature_header if signing_secret else None,
        },
        "routes": routes,
    }
    _write_json(pathlib.Path(args.dispatch_report), report)

    print(f"Wrote dispatch report: {args.dispatch_report}")
    print(
        "Dispatch status:",
        report["status"],
        "| alert_status=",
        report["alert_status"],
        "| delivered=",
        report["delivered_route_count"],
        "| failed=",
        report["failed_route_count"],
    )
    for route in routes:
        print(
            " - route:",
            route["route"],
            "| status=",
            route["status"],
            "| endpoint=",
            route.get("endpoint"),
            "| attempts=",
            route.get("delivery", {}).get("attempt_count"),
            "| ack_received=",
            route.get("ack", {}).get("received"),
            "| correlation_match=",
            route.get("correlation", {}).get("matches"),
        )

    if args.fail_on_route_error and failed_routes:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
