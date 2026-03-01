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
    1. Generate deterministic workbook fixtures (`python/generate_corpus_fixtures.py`).
    2. Run workspace tests.
    3. Run `part-graph-corpus --fail-on-errors`.
    4. Upload corpus report + JSONL + generated fixtures as run artifacts.
    5. Assemble standardized artifact bundle directory + manifest prior to upload.
- Workflow: `.github/workflows/repro-bundle.yml`.
  - Triggers: `pull_request`, `push` to `main`, nightly `schedule`, and `workflow_dispatch`.
  - Steps:
    1. Generate deterministic workbook fixtures for repro validation.
    2. Run workspace tests.
    3. Run `repro record/check/diff` against baseline and mutated candidates.
    4. Assert mismatch detection for mutated candidate.
    5. Assemble standardized artifact bundle directory + manifest prior to upload.
    6. Upload bundle + diff outputs + JSONL traces as run artifacts.
- Workflow: `.github/workflows/batch-recalc-nightly.yml`.
  - Triggers: nightly `schedule` and `workflow_dispatch`.
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
    - `ALERT_POLICY_SCHEMA_VALIDATION_ENABLED`, `ALERT_POLICY_SCHEMA_CANARY_VALIDATION_ENABLED`
    - `ALERT_POLICY_SCHEMA_SNAPSHOT_PATH`, `ALERT_POLICY_SCHEMA_DISPATCH_PATH`, `ALERT_POLICY_SCHEMA_ACK_RETENTION_PATH`
    - `ALERT_POLICY_SCHEMA_DASHBOARD_PACK_PATH`, `ALERT_POLICY_SCHEMA_POLICY_PATH`
    - `ALERT_POLICY_SCHEMA_ESCALATION_PATH`, `ALERT_POLICY_SCHEMA_ADAPTER_EXPORTS_PATH`
  - Steps:
    1. Assemble deterministic nightly compatibility corpus slice (`python/build_batch_nightly_corpus.py`) using generated fixtures + curated workbook samples.
    2. Run workspace tests.
    3. Run `batch recalc` with bounded concurrency and diagnostic detail output over the expanded corpus slice.
    4. Build throughput trend snapshot + alert-hook payload (`python/build_batch_trend_snapshot.py`) with threshold metadata.
    5. Dispatch alert payload to configured incident/dashboard ingestion routes (`python/dispatch_batch_alert_hook.py`) and emit dispatch report artifact with auth/retry/ack/idempotency/correlation/replay metadata.
    6. Build acknowledgement-retention lookup index (`python/build_batch_ack_retention_index.py`) from dispatch output for incident forensics.
    7. Build dashboard-pack and alert-policy artifacts (`python/build_batch_dashboard_pack.py`) from snapshot + dispatch + ack-retention outputs.
    8. Build policy-owner escalation metadata + downstream adapter exports (`python/build_batch_policy_adapters.py`) from alert-policy and dashboard-pack artifacts.
    9. Validate full nightly artifact family against versioned schemas and compatibility contracts (`python/validate_batch_adapter_contracts.py --full-family`).
    10. Run schema-drift canary checks (`python/validate_batch_schema_canaries.py`) to assert deterministic fail behavior for representative compatibility regressions.
    11. Enforce nightly gate from snapshot + policy status after dispatch routing.
    12. Assemble standardized artifact bundle directory + manifest prior to upload.
    13. Upload batch report + JSONL + trend snapshot + alert payload + dispatch report + ack-retention index + dashboard-pack + alert-policy + policy-escalation + adapter-exports + assembled corpus manifest/files as run artifacts.

## Failure Handling
- Auto-create incident ticket for failing nightly gates with regression labels.
- Nightly batch gate now fails on either throughput snapshot breach or alert-policy breach status.
- Incident/dash adapters consume `ci-batch-policy-escalation.json` + `ci-batch-dashboard-adapter-exports.json` for owner-targeted escalation and dashboard sync.
- Nightly artifact publication is blocked when schema/compatibility validation fails for snapshot/dispatch/ack-retention/dashboard-pack/policy/escalation/adapter payloads.
- Nightly artifact publication is also blocked when schema-drift canary assertions fail (unexpected validator pass/fail behavior).
- Block release branch merge on unresolved P1/P2 alerts.
