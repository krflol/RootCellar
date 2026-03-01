# Sprint 06 - CLI Batch and Repro Mode

Parent: [[Sprint Cadence and Capacity]]
Dates: May 25, 2026 to June 7, 2026

## Sprint Goal
Enable reliable headless automation workflows with reproducibility records.

## Execution Status (March 1, 2026)
- Batch baseline delivered:
  - `batch recalc` command with bounded Rayon parallelism, deterministic ordering, fail-on-error gating, and detail-level artifact output.
  - Nightly CI execution/publishing flow: `.github/workflows/batch-recalc-nightly.yml`.
  - Nightly corpus coverage expanded to deterministic compatibility slices (`python/build_batch_nightly_corpus.py`, `target-files=32`) with throughput regression threshold checks.
  - Nightly synthetic benchmark gate delivered: optional `bench recalc-synthetic` execution with CI policy knobs for workload/threshold control (`BATCH_BENCH_*`) and benchmark artifact publication (`ci-batch-bench-recalc-synthetic.json`, `ci-batch-bench-events.jsonl`).
  - Nightly trend snapshot + alert-hook artifacts delivered (`python/build_batch_trend_snapshot.py` producing `ci-batch-throughput-snapshot.json` and `ci-batch-alert-hook.json`).
  - Nightly alert-route dispatch utility delivered (`python/dispatch_batch_alert_hook.py` producing `ci-batch-alert-dispatch.json` for route status introspection).
  - Dispatch hardening delivered: token auth, optional signing, retry/backoff controls, and ack-required tracking for endpoint integrations.
  - Dispatch traceability IDs delivered: deterministic per-route idempotency keys and correlation IDs with optional downstream correlation-match enforcement.
  - Dispatch replay-protection policy delivered: timestamp/nonce/window headers with per-attempt replay metadata in dispatch artifacts.
  - Ack-retention forensic indexing delivered: `python/build_batch_ack_retention_index.py` producing `ci-batch-ack-retention-index.json` with hashed ack lookup keys and expiry windows.
  - Dashboard-pack and alert-policy artifacts delivered: `python/build_batch_dashboard_pack.py` producing `ci-batch-dashboard-pack.json` and `ci-batch-alert-policy.json` with policy-gated CI enforcement.
  - Policy-owner escalation and adapter exports delivered: `python/build_batch_policy_adapters.py` producing `ci-batch-policy-escalation.json` and `ci-batch-dashboard-adapter-exports.json`.
  - Adapter schema/compatibility validation delivered: `python/validate_batch_adapter_contracts.py` with versioned schemas (`schemas/artifacts/v1/*`) and nightly CI contract enforcement.
  - Full artifact-family schema/compatibility validation delivered: `python/validate_batch_adapter_contracts.py --full-family` now gates snapshot/dispatch/ack-retention/dashboard-pack/policy plus escalation/adapter artifacts in nightly CI.
  - Schema-drift canary fixture validation delivered: `python/validate_batch_schema_canaries.py` now asserts expected contract failures for representative drift scenarios in nightly CI.
  - Dual-read migration drills delivered: `python/validate_batch_dual_read_migration.py` now verifies producer/consumer overlap and rollback behavior for schema major-version transitions across snapshot/dispatch/ack-retention/dashboard-pack/policy/escalation/adapter artifacts, with workflow subset control via `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_ARTIFACTS`, staged-wave control via `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_WAVE_SPEC`, fault-injection control via `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_FAULT_INJECTION_ENABLED`/`ALERT_POLICY_SCHEMA_MIGRATION_DRILL_FAULT_SCENARIOS`, and diagnostics artifact output (`ci-batch-schema-migration-drill.json`).
  - Migration policy dry-run checks delivered: `python/validate_batch_migration_policy_dry_run.py` now asserts invalid staged-wave specs and unsupported fault-scenario policy values fail as expected in nightly CI (`ALERT_POLICY_SCHEMA_MIGRATION_DRY_RUN_POLICY_VALIDATION_ENABLED`).
- Repro baseline delivered:
  - `repro record`, `repro check`, and `repro diff` command paths.
  - CI execution/publishing flow: `.github/workflows/repro-bundle.yml`.

## Commitments
- Epic 05 primary.
- Epic 02 deterministic replay checks.
- Epic 07 artifact registry enrichment.

## Stories
1. Implement `batch` command with bounded Rayon parallelism.
2. Implement `repro record` and `repro check` command paths.
3. Emit JSONL report with standardized event schema.
4. Integrate macro execution into CLI with policy enforcement.
5. Add CI nightly batch benchmark and deterministic replay job.

## Acceptance Criteria
- Batch runs produce consistent artifacts with stable naming in deterministic mode.
- Repro check detects drift in calc or output bytes.
- JSONL output validates against schema in CI.

## Exit Signals
- Headless mode used by at least one internal pipeline.
