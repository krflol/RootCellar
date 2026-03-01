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
  10. Confirm schema-drift canary gate passed (`python/validate_batch_schema_canaries.py`) and review canary-case output for unexpected validator behavior.
  11. Confirm dual-read migration drill gate passed (`python/validate_batch_dual_read_migration.py`) across the selected artifact matrix/staged waves/fault scenarios and verify no rollback regression for schema major transitions (`ALERT_POLICY_SCHEMA_MIGRATION_DRILL_ARTIFACTS`, `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_WAVE_SPEC`, `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_FAULT_INJECTION_ENABLED`, `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_FAULT_SCENARIOS`, `alert_policy_schema_migration_drill_artifacts`, `alert_policy_schema_migration_drill_wave_spec`, `alert_policy_schema_migration_drill_fault_injection_enabled`, `alert_policy_schema_migration_drill_fault_scenarios`).
  12. Confirm migration policy dry-run gate passed (`python/validate_batch_migration_policy_dry_run.py`) so invalid staged-wave specs and unsupported fault-scenario keys are rejected by policy parsing (`ALERT_POLICY_SCHEMA_MIGRATION_DRY_RUN_POLICY_VALIDATION_ENABLED`, `alert_policy_schema_migration_dry_run_policy_validation_enabled`).
  13. Inspect `ci-batch-schema-migration-drill.json` for failing phase diagnostics (`phase`, `expectation`, `validator_exit_code`, `expected_token_found`, `validator_output_excerpt`) and executed fault scenario list before escalating schema rollback decisions.

## Communication Cadence
- Sev 1 updates every 30 minutes.
- Sev 2 updates every 60 minutes.

## Postmortem Requirements
- Root cause and contributing factors.
- Detection and response timeline.
- Corrective actions with owner and due date.
- Observability gaps identified and tracked.
