# Dashboards SLOs and Alerts

Parent: [[Observability Charter]]

## Dashboard Packs
- Compatibility dashboard: no-repair rate, feature-gap trend, preserved-only counts.
- Calc dashboard: recalc latency percentiles, cycle rate, mismatch rate.
- UI dashboard: frame time, edit latency, command failure rate.
- Script security dashboard: permission prompts, deny rates, policy violations.
- CLI/batch dashboard: throughput, failure by command, repro mismatch rate.
  - Nightly batch trend artifact source: `ci-batch-throughput-snapshot.json` from `batch-recalc-nightly` workflow.
  - Route delivery status source: `ci-batch-alert-dispatch.json` for incident/dashboard dispatch observability.
  - Ack-retention forensic source: `ci-batch-ack-retention-index.json` (lookup keys for `ack_id`, `ack_id_sha256`, `idempotency_key`, and `correlation_id` with retention expiry windows).
  - Dashboard-pack source: `ci-batch-dashboard-pack.json` (panel-ready throughput, delivery, forensics, and policy drill-down payload).
  - Policy source: `ci-batch-alert-policy.json` (normalized check set with severity ranking and breach counts).
  - Escalation metadata source: `ci-batch-policy-escalation.json` (policy-check owner mappings, escalation targets, and SLA metadata).
  - Dashboard adapter export source: `ci-batch-dashboard-adapter-exports.json` (downstream incident/dashboard ingestion payload contracts).
  - Contract schemas: `schemas/artifacts/v1/batch-policy-escalation.schema.json` and `schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json` (CI validation + compatibility contracts).
  - Route dispatch drill-down: retry attempt counts, auth/signing configuration flags, idempotency/correlation IDs, replay timestamp/nonce/window metadata, and ack/correlation-required vs received/matched counters.

## Initial SLOs
- SLO-COMP-01: corpus no-repair pass >= 92% by RC.
- SLO-CALC-01: p95 recalc <= 350 ms benchmark medium set.
- SLO-UI-01: p95 edit commit <= 120 ms.
- SLO-SEC-01: zero unresolved critical sandbox/policy findings in release branch.
- SLO-OBS-01: trace completeness >= 95% for critical flows.

## Alert Policies
- P1: any sandbox bypass or policy bypass event.
- P1: corpus no-repair drop >= 5 points day-over-day.
- P2: trace completeness < 90% for 2 consecutive hours in staging.
- P2: recalc p95 exceeds budget by > 30% for 24h.
- P3: rising unsupported feature detections in top workflows.
- P3: nightly batch throughput regression breach (`ci-batch-alert-hook.json` with `routing_key=rootcellar.batch.throughput`).
- Nightly policy artifact (`ci-batch-alert-policy.json`) normalizes these checks for CI gate enforcement and downstream incident-routing adapters.
- Escalation routing artifact (`ci-batch-policy-escalation.json`) maps breach checks to owner teams/channels/targets for incident handoff.

## Runbooks
Runbooks are in [[docs/RootCellar/06-Operations/Incident Response Playbook]].
Each alert maps to owner team and triage checklist.
