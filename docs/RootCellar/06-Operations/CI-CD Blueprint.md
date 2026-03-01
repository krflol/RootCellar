# CI-CD Blueprint

Parent: [[Environment Matrix]]

## Pipeline Stages
1. Build and static checks.
2. Unit/integration tests.
3. Schema validation for telemetry/artifacts.
4. Corpus and golden suites.
5. Benchmark and deterministic replay checks.
6. Security validation subset.
7. Package signing and provenance metadata.

## Promotion Rules
- PR -> main requires all PR gates.
- Main -> staging requires nightly gate green in last 24h.
- Staging -> release requires release gate checklist approval.

## Artifact Policy
- Every pipeline run uploads structured artifact bundle.
- Release artifacts include checksums, signatures, and manifest provenance.
- Standardized CI artifact naming: `rootcellar-<workflow>-<run_id>-<run_attempt>`.
- Standardized CI artifact retention: `21` days for nightly/PR workflow artifacts.
- Bundle manifests include run metadata and `retention_policy_days`.

## Current Implemented CI Slice (March 1, 2026)
- Workflow: `.github/workflows/corpus-part-graph.yml`.
  - Triggers: `pull_request`, `push` to `main`, nightly `schedule`, and `workflow_dispatch`.
  - Steps:
    1. Install fixture-generation/interoperability Python dependencies (`python/requirements-interop.txt`).
    2. Generate deterministic workbook fixtures (`python/generate_corpus_fixtures.py`).
    3. Run workspace tests.
    4. Run `part-graph-corpus --fail-on-errors`.
    5. Upload corpus report + JSONL + generated fixtures as run artifacts.
    6. Assemble standardized artifact bundle directory + manifest prior to upload.
- Workflow: `.github/workflows/repro-bundle.yml`.
  - Triggers: `pull_request`, `push` to `main`, nightly `schedule`, and `workflow_dispatch`.
  - Steps:
    1. Install fixture-generation/interoperability Python dependencies (`python/requirements-interop.txt`).
    2. Generate deterministic workbook fixtures for repro validation.
    3. Run workspace tests.
    4. Run `repro record/check/diff` against baseline and mutated candidates.
    5. Assert mismatch detection for mutated candidate.
    6. Assemble standardized artifact bundle directory + manifest prior to upload.
    7. Upload bundle + diff outputs + JSONL traces as run artifacts.
- Workflow: `.github/workflows/excel-interop.yml`.
  - Triggers: `pull_request`, `push` to `main`, nightly `schedule`, and `workflow_dispatch`.
  - Policy env knobs:
    - `EXCEL_INTEROP_MIN_EXCEL_AUTHORED_SAMPLES` (minimum curated real Excel-authored samples required in assembled interop corpus).
    - `EXCEL_INTEROP_REQUIRED_CURATED_FEATURES` (comma-separated curated feature tags required in assembled corpus).
    - Current default in workflow: `5`.
    - Current required curated feature defaults: `formulas,styles,comments,charts,defined_names`.
  - Steps:
    1. Install interoperability Python dependencies (`python/requirements-interop.txt`).
    2. Assemble deterministic interop corpus (`python/assemble_excel_interop_corpus.py`) from generated fixtures plus curated `corpus/excel-authored/manifest.json` samples, including legal-clearance metadata and minimum-sample policy enforcement.
    3. Run bidirectional interop harness (`python/verify_excel_interop.py`) with corpus sweep, manifest capture, and required-fixture assertions (`--require-corpus-fixture`).
    4. Upload interop report + workdir + assembled corpus as run artifacts.
    5. Assemble standardized artifact bundle directory + manifest prior to upload.
- Workflow: `.github/workflows/batch-recalc-nightly.yml`.
  - Triggers: nightly `schedule` and `workflow_dispatch`.
  - Benchmark env knobs:
    - `BATCH_BENCH_RECALC_SYNTHETIC_ENABLED`
    - `BATCH_BENCH_CHAINS`, `BATCH_BENCH_CHAIN_LENGTH`, `BATCH_BENCH_ITERATIONS`, `BATCH_BENCH_CHANGED_CHAIN`
    - `BATCH_BENCH_MIN_DURATION_SPEEDUP_RATIO`, `BATCH_BENCH_MAX_EVALUATED_CELLS_RATIO`
  - Optional route config secrets:
    - `ROOTCELLAR_INCIDENT_WEBHOOK_URL`
    - `ROOTCELLAR_DASHBOARD_INGEST_URL`
    - `ROOTCELLAR_INCIDENT_WEBHOOK_TOKEN`
    - `ROOTCELLAR_DASHBOARD_INGEST_TOKEN`
    - `ROOTCELLAR_ALERT_SIGNING_SECRET`
  - Dispatch policy env knobs:
    - `ALERT_DISPATCH_MAX_ATTEMPTS`, `ALERT_DISPATCH_INITIAL_BACKOFF_SEC`, `ALERT_DISPATCH_BACKOFF_MULTIPLIER`, `ALERT_DISPATCH_MAX_BACKOFF_SEC`
    - `ALERT_DISPATCH_REPLAY_WINDOW_SEC`, `ALERT_DISPATCH_REPLAY_TIMESTAMP_HEADER`, `ALERT_DISPATCH_REPLAY_NONCE_HEADER`, `ALERT_DISPATCH_REPLAY_WINDOW_HEADER`
    - `ALERT_DISPATCH_REQUIRE_ACK_ON_INCIDENT`, `ALERT_DISPATCH_REQUIRE_ACK_ON_DASHBOARD`
    - `ALERT_DISPATCH_REQUIRE_CORRELATION_ON_INCIDENT`, `ALERT_DISPATCH_REQUIRE_CORRELATION_ON_DASHBOARD`
    - `ALERT_ACK_RETENTION_DAYS`
  - Dashboard/policy env knobs:
    - `ALERT_POLICY_MAX_DISPATCH_FAILED_ROUTES`, `ALERT_POLICY_MAX_ACK_MISSING_ROUTES`, `ALERT_POLICY_MAX_CORRELATION_MISMATCH_ROUTES`
    - `ALERT_POLICY_REQUIRE_REPLAY_METADATA`, `ALERT_POLICY_REQUIRE_ACK_RETENTION_COVERAGE`, `ALERT_POLICY_FAIL_ON_BREACH`
    - `ALERT_POLICY_OWNER_TEAM_DEFAULT`, `ALERT_POLICY_OWNER_TEAM_SNAPSHOT`, `ALERT_POLICY_OWNER_TEAM_DISPATCH`, `ALERT_POLICY_OWNER_TEAM_ACK_RETENTION`, `ALERT_POLICY_OWNER_TEAM_POLICY`
    - `ALERT_POLICY_OWNER_CONTACT_CHANNEL`
    - `ALERT_POLICY_ESCALATION_TARGET_P1`, `ALERT_POLICY_ESCALATION_TARGET_P2`, `ALERT_POLICY_ESCALATION_TARGET_P3`, `ALERT_POLICY_ESCALATION_TARGET_INFO`
    - `ALERT_POLICY_ESCALATION_SLA_MINUTES_P1`, `ALERT_POLICY_ESCALATION_SLA_MINUTES_P2`, `ALERT_POLICY_ESCALATION_SLA_MINUTES_P3`, `ALERT_POLICY_ESCALATION_SLA_MINUTES_INFO`
    - `ALERT_POLICY_SCHEMA_VALIDATION_ENABLED`, `ALERT_POLICY_SCHEMA_CANARY_VALIDATION_ENABLED`, `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_VALIDATION_ENABLED`, `ALERT_POLICY_SCHEMA_MIGRATION_DRY_RUN_POLICY_VALIDATION_ENABLED`
    - `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_ARTIFACTS`
    - `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_WAVE_SPEC`
    - `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_FAULT_INJECTION_ENABLED`, `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_FAULT_SCENARIOS`
    - `ALERT_POLICY_SCHEMA_SNAPSHOT_PATH`, `ALERT_POLICY_SCHEMA_DISPATCH_PATH`, `ALERT_POLICY_SCHEMA_ACK_RETENTION_PATH`
    - `ALERT_POLICY_SCHEMA_DASHBOARD_PACK_PATH`, `ALERT_POLICY_SCHEMA_POLICY_PATH`
    - `ALERT_POLICY_SCHEMA_ESCALATION_PATH`, `ALERT_POLICY_SCHEMA_ADAPTER_EXPORTS_PATH`
  - Bundle manifest records migration-drill policy state, artifact matrix, staged-wave policy, fault-injection policy, dry-run policy state, and drill-report publication flags (`alert_policy_schema_migration_drill_validation_enabled`, `alert_policy_schema_migration_dry_run_policy_validation_enabled`, `alert_policy_schema_migration_drill_artifacts`, `alert_policy_schema_migration_drill_wave_spec`, `alert_policy_schema_migration_drill_fault_injection_enabled`, `alert_policy_schema_migration_drill_fault_scenarios`, `schema_migration_drill_report_generated`).
  - Steps:
    1. Install fixture-generation/interoperability Python dependencies (`python/requirements-interop.txt`).
    2. Assemble deterministic nightly compatibility corpus slice (`python/build_batch_nightly_corpus.py`) using generated fixtures + curated workbook samples.
    3. Run workspace tests.
    4. Run `batch recalc` with bounded concurrency and diagnostic detail output over the expanded corpus slice.
    5. Optionally run synthetic recalc benchmark (`bench recalc-synthetic`) and publish benchmark report/events artifacts when enabled by policy.
    6. Build throughput trend snapshot + alert-hook payload (`python/build_batch_trend_snapshot.py`) with threshold metadata.
    7. Dispatch alert payload to configured incident/dashboard ingestion routes (`python/dispatch_batch_alert_hook.py`) and emit dispatch report artifact with auth/retry/ack/idempotency/correlation/replay metadata.
    8. Build acknowledgement-retention lookup index (`python/build_batch_ack_retention_index.py`) from dispatch output for incident forensics.
    9. Build dashboard-pack and alert-policy artifacts (`python/build_batch_dashboard_pack.py`) from snapshot + dispatch + ack-retention outputs.
    10. Build policy-owner escalation metadata + downstream adapter exports (`python/build_batch_policy_adapters.py`) from alert-policy and dashboard-pack artifacts.
    11. Validate full nightly artifact family against versioned schemas and compatibility contracts (`python/validate_batch_adapter_contracts.py --full-family`).
    12. Run schema-drift canary checks (`python/validate_batch_schema_canaries.py`) to assert deterministic fail behavior for representative compatibility regressions.
    13. Run dual-read migration drills (`python/validate_batch_dual_read_migration.py`) to verify producer/consumer overlap and rollback behavior for major-version schema transitions across snapshot/dispatch/ack-retention/dashboard-pack/policy/escalation/adapter artifact families, including optional staged-wave scenarios, fault-injection cases (malformed fallback schemas + partial-wave rollback rehearsal), and structured diagnostics output (`ci-batch-schema-migration-drill.json`).
    14. Run migration-policy dry-run checks (`python/validate_batch_migration_policy_dry_run.py`) to assert invalid staged-wave specs and unsupported fault-scenario keys are rejected by policy parsing.
    15. Enforce nightly gate from snapshot + policy status plus optional synthetic benchmark thresholds.
    16. Assemble standardized artifact bundle directory + manifest prior to upload.
    17. Upload batch report + batch/benchmark JSONL + trend snapshot + alert payload + dispatch report + ack-retention index + dashboard-pack + alert-policy + policy-escalation + adapter-exports + schema-migration-drill diagnostics + benchmark report + assembled corpus manifest/files as run artifacts.

## Failure Handling
- Auto-create incident ticket for failing nightly gates with regression labels.
- Nightly batch gate now fails on either throughput snapshot breach or alert-policy breach status.
- Nightly batch gate also fails when synthetic benchmark thresholds are enabled and violated (`duration_speedup_ratio` below minimum or `evaluated_cells_reduction_ratio` above maximum).
- Incident/dash adapters consume `ci-batch-policy-escalation.json` + `ci-batch-dashboard-adapter-exports.json` for owner-targeted escalation and dashboard sync.
- Nightly artifact publication is blocked when schema/compatibility validation fails for snapshot/dispatch/ack-retention/dashboard-pack/policy/escalation/adapter payloads.
- Nightly artifact publication is also blocked when schema-drift canary assertions fail (unexpected validator pass/fail behavior).
- Nightly artifact publication is also blocked when dual-read migration drill assertions fail (producer/consumer overlap or rollback regression).
- Nightly artifact publication is also blocked when migration-policy dry-run assertions fail (invalid staged-wave spec or unsupported fault-scenario policy acceptance regression).
- Block release branch merge on unresolved P1/P2 alerts.
