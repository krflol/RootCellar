# Incident Response Playbook

Parent: [[Environment Matrix]]
Related: [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]]

## Severity Model
- Sev 1: data loss/corruption risk, security bypass, widespread outage.
- Sev 2: major workflow degradation or significant SLO breach.
- Sev 3: localized issues with workaround.

## Triage Steps
1. Capture trace_id/workbook_id/command context.
2. Pull related dashboards and alert details.
3. Retrieve artifact bundle and audit records.
4. Classify scope and user impact.
5. Execute mitigation and communicate status.

## Nightly Batch Regression Signal
- Alert route key: `rootcellar.batch.throughput`.
- Primary payload artifacts:
  - `ci-batch-alert-hook.json` (breach details and threshold comparisons).
  - `ci-batch-alert-dispatch.json` (delivery outcome by incident/dashboard route).
  - `ci-batch-ack-retention-index.json` (ack lookup keys + retention expiry metadata for forensics).
  - `ci-batch-alert-policy.json` (policy check outcomes + severity ranking used by nightly gate).
  - `ci-batch-policy-escalation.json` (policy-check owner mappings and escalation targets/SLA windows).
  - `ci-batch-dashboard-adapter-exports.json` (incident + dashboard adapter payloads for downstream ingestion systems).
  - `ci-batch-dashboard-pack.json` (panel-ready joined payload for throughput/delivery/forensics drill-down).
  - `ci-batch-throughput-snapshot.json` (trend metrics and gating status).
- Dispatch triage checklist:
  1. Confirm whether incident/dashboard routes were configured.
  2. Review per-route retry attempts and final HTTP status.
  3. Verify auth/signing and replay-protection headers (`timestamp`, `nonce`, `window`) in dispatch report attempts.
  4. Verify acknowledgement fields (`ack_id`) and required-ack status in dispatch report.
  5. Compare expected vs actual `correlation_id` values and track route `idempotency_key` for duplicate-delivery triage.
  6. Use ack-retention index lookups (`ack_id`, `ack_id_sha256`, `idempotency_key`, `correlation_id`) to pivot across incident-system ingestion logs.
  7. Review alert-policy breaches (`ci-batch-alert-policy.json`) to prioritize severity and owner escalation path.
  8. Use policy-escalation metadata (`ci-batch-policy-escalation.json`) to route to owner queues/channels and confirm adapter-export payloads match incident-system expectations.
  9. Confirm nightly artifacts passed full-family schema/compatibility validation (`python/validate_batch_adapter_contracts.py --full-family`) against `schemas/artifacts/v1/*`.

## Communication Cadence
- Sev 1 updates every 30 minutes.
- Sev 2 updates every 60 minutes.

## Postmortem Requirements
- Root cause and contributing factors.
- Detection and response timeline.
- Corrective actions with owner and due date.
- Observability gaps identified and tracked.
