# Epic 07 - Radical Observability and Introspection

Parent: [[docs/RootCellar/00-Program/RootCellar Master Plan]]
Specs: [[docs/RootCellar/04-Observability/Observability Charter]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/04-Observability/AI Introspection Workflow]]

## Objective
Make every critical product path inspectable by humans and AI through structured telemetry, traces, and artifact bundles.

## Execution Status
- Status: In progress (telemetry and bundle baseline delivered).
- Tracking links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]
- Completed slice:
1. Event envelope schema contract and JSONL sink.
2. Trace context propagation through CLI/core flows.
3. Artifact bundle workflows (`repro record/check/diff`) with explicit telemetry events.
4. Recalc introspection artifacts (`--dep-graph-report`, `--dag-timing-report`) with telemetry output events.
5. Function-level recalc observability via dependency graph `function_call_count` metrics.
6. AST-level recalc observability via `formula_ast_ids`, `ast_node_count`, and `ast_unique_node_count`.
7. DAG-level recalc observability via critical-path and fan-in/fan-out timing metrics.
8. Configurable slow-node threshold introspection (`--dag-slow-threshold-us`) propagated into DAG artifacts and telemetry payload context.
9. XLSX part-graph observability via `interop.xlsx.part_graph.built` and graph-aware save payload flags (`relationships_preserved`, `unknown_parts_preserved`, strategy).
10. Corpus-level part-graph observability via aggregate run events (`artifact.part_graph.corpus.start|end`) and machine-readable corpus artifact output.
11. CI artifact publication baseline for corpus observability bundles (`corpus-part-graph` workflow uploads report + JSONL + fixture corpus).
12. CI artifact publication baseline for reproducibility observability bundles (`repro-bundle` workflow uploads bundle, diff outputs, and JSONL traces).
13. CI artifact naming/retention policy aligned across workflow families, with per-run manifest metadata including retention fields.
14. Batch recalc observability delivered with aggregate run lifecycle + per-file telemetry (`artifact.batch.recalc.start|file|end`) and inspectable batch report artifacts.
15. Nightly batch artifact publication baseline (`batch-recalc-nightly` workflow) with standardized bundle naming, retention, and manifest metadata.
16. Nightly batch observability expanded with throughput regression gates (`throughput_files_per_sec`) and deterministic corpus-slice manifests (`python/build_batch_nightly_corpus.py` + corpus manifest artifact).
17. Nightly batch trend and alert payload artifacts delivered (`python/build_batch_trend_snapshot.py`) for dashboard and incident-ingestion integration.
18. Nightly alert-route dispatch integration delivered (`python/dispatch_batch_alert_hook.py`) with per-route delivery status artifacts and explicit post-dispatch gate enforcement.
19. Alert-route dispatch hardening delivered (token auth, optional HMAC signing, retry/backoff policy, and ack tracking) for reliable cross-system observability.
20. Cross-system traceability identifiers delivered for alert dispatch paths (deterministic idempotency keys + correlation IDs with optional downstream correlation-match enforcement).
21. Alert replay-protection + ack-retention forensics delivered (timestamp/nonce/window dispatch metadata and `ci-batch-ack-retention-index.json` lookup artifact publication).
22. Nightly dashboard-pack + alert-policy artifacts delivered (`ci-batch-dashboard-pack.json`, `ci-batch-alert-policy.json`) with policy gate enforcement in CI.
23. Policy-to-owner escalation metadata + adapter exports delivered (`ci-batch-policy-escalation.json`, `ci-batch-dashboard-adapter-exports.json`) for downstream incident/dashboard ingestion systems.
24. Adapter export schema validation + compatibility contracts delivered in CI (`schemas/artifacts/v1/*`, `python/validate_batch_adapter_contracts.py`) to reject incompatible artifact versions.
25. Full artifact-family schema validation + compatibility contracts delivered in CI (`python/validate_batch_adapter_contracts.py --full-family`) covering snapshot/dispatch/ack-retention/dashboard-pack/policy plus escalation/adapter artifacts.
26. Schema-drift canary fixture gate + migration playbook delivered (`python/validate_batch_schema_canaries.py`, `Artifact Schema Migration Playbook`) for compatibility-regression detection and major-version rollout discipline.
27. Dual-read migration drill gate delivered (`python/validate_batch_dual_read_migration.py`) for producer/consumer overlap and rollback verification in schema major-version transitions.
28. Multi-artifact dual-read matrix expansion delivered for snapshot/dispatch/ack-retention/dashboard-pack/policy/escalation/adapter families, including artifact-subset and staged-wave policy knobs plus manifest introspection (`ALERT_POLICY_SCHEMA_MIGRATION_DRILL_ARTIFACTS`, `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_WAVE_SPEC`, `alert_policy_schema_migration_drill_artifacts`, `alert_policy_schema_migration_drill_wave_spec`).
29. Dual-read drill diagnostics artifact publication delivered (`ci-batch-schema-migration-drill.json`) with phase-level expected-fail token checks, validator output excerpts, and timing metadata for forensic triage.
30. Dual-read fault-injection drill scenarios delivered (malformed fallback schema + partial-wave rollback rehearsal) with explicit policy knobs and manifest exposure (`ALERT_POLICY_SCHEMA_MIGRATION_DRILL_FAULT_INJECTION_ENABLED`, `ALERT_POLICY_SCHEMA_MIGRATION_DRILL_FAULT_SCENARIOS`, `alert_policy_schema_migration_drill_fault_injection_enabled`, `alert_policy_schema_migration_drill_fault_scenarios`).
31. Migration-drill policy dry-run checks delivered (`python/validate_batch_migration_policy_dry_run.py`) to enforce expected parser rejection of invalid staged-wave specs and unsupported fault-scenario keys in nightly CI (`ALERT_POLICY_SCHEMA_MIGRATION_DRY_RUN_POLICY_VALIDATION_ENABLED`, `alert_policy_schema_migration_dry_run_policy_validation_enabled`).
- Remaining:
1. Cross-surface UI->engine->script trace bridge completion.
2. AI introspection query surfaces over stored artifacts.

## Scope
- Unified telemetry taxonomy across UI/engine/script/CLI.
- End-to-end trace correlation.
- Inspectable artifact bundle schema.
- Dashboards, SLOs, alerting, and forensic workflows.

## Deliverables
- Event SDKs and schema validation in CI.
- Trace context propagation across process boundaries.
- Artifact registry with query tooling.
- Operational dashboards and alert policies.

## Stories
1. Define event taxonomy and semantic conventions.
2. Implement trace propagation in all surfaces.
3. Define and implement artifact bundle writer/reader.
4. Build dashboard packs and alert runbooks.
5. Add AI introspection interfaces over artifacts.

## Acceptance Criteria
- >= 95% of critical workflows produce complete end-to-end traces.
- Every release candidate build has artifact completeness report.
- Incident triage can reconstruct user action -> engine outcome -> script actions.

## Dependencies
- None; this epic is foundational and parallelized with all others.

## Observability Requirements
This epic defines them for all epics; see linked specs.
