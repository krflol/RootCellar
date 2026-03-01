# Execution Status

Parent: [[docs/RootCellar/00-Program/RootCellar Master Plan]]
Last updated: March 1, 2026

## Current Execution Slice
- Slice: Sprint 02 calc internals + Sprint 06 batch CLI/observability hardening on top of completed Sprint 00/01 baselines.
- Status: In progress.

## Execution Plan Linkage
- Canonical plan-to-delivery board: [[Execution Plan Board]]
- Completed plan items linked in docs:
1. [[docs/RootCellar/02-Sprints/Sprint 00 - Foundation and Telemetry Bootstrap#Execution Status]]
2. [[docs/RootCellar/02-Sprints/Sprint 01 - Workbook and XLSX Skeleton#Execution Status]]
3. [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
- In-progress plan items linked in docs:
1. [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph#Execution Status]]
2. [[docs/RootCellar/00-Program/Milestone Roadmap#Current Milestone Status (February 28, 2026)]]

## Completed In Code
1. Rust workspace scaffold created.
2. `rootcellar-core` implemented with:
   - Workbook transaction model (`begin_txn`, mutations, commit digest).
   - Structured telemetry envelope + JSONL sink.
   - XLSX inspector producing compatibility report artifacts.
   - Minimal recalc engine with A1 refs, arithmetic eval, and cycle detection.
   - Parser scaffold with arithmetic precedence/parentheses and sheet dependency-graph analysis (topological order + cycle/parse diagnostics).
   - Built-in function coverage increment for recalc (`SUM`, `MIN`, `MAX`, `IF`, `AVERAGE`/`AVG`, `ABS`, `AND`, `OR`, `NOT`).
   - AST interning scaffold for parser structures (deduplicated AST node IDs, per-formula AST IDs, and intern previews).
   - Incremental recalc path (`recalc_sheet_from_roots`) for dependency-scoped recompute.
   - Recalc DAG timing artifact support (`recalc_sheet_with_dag_timing`, `recalc_sheet_from_roots_with_dag_timing`).
   - DAG analysis metrics in timing artifacts (critical path, fan-in/fan-out, slow-node threshold), including deterministic tie-break behavior and threshold override support.
   - XLSX workbook model loader (workbook/sheet XML + shared strings + formula/value projection).
   - XLSX part-graph reconstruction baseline with `_rels` edge extraction, dangling-target detection, and known/unknown part classification.
3. `rootcellar-cli` implemented with:
    - `open` command (inspect workbook + report JSON output).
    - `tx-demo` command (transaction + telemetry demo).
    - `tx-save` command (apply one or more value/formula transaction edits and persist with preserve/normalize mode), including repeated `--set` and `--setf` mutations and range expansion (`A1:B3=...`).
    - `tx-save` post-mutation recalc now uses incremental dependency-scoped recompute.
    - `recalc` command (load workbook model, recalc one/all sheets, emit report JSON).
    - `recalc --dep-graph-report` output for inspectable dependency graph artifacts.
    - `recalc --dag-timing-report` output for inspectable per-node DAG timing artifacts.
    - `recalc --dag-slow-threshold-us` to override DAG slow-node threshold when generating timing artifacts.
    - `part-graph-corpus` command to scan directories recursively for `.xlsx` files, emit aggregate part-graph artifacts, and optionally fail on inspection errors (`--fail-on-errors`).
    - `save` command (load workbook model and write normalized/preserve-mode baseline `.xlsx` output).
    - Save/open summaries now include part-graph diagnostics (node/edge counts, dangling-edge counts, and preserve/normalize graph flags).
    - `repro` command group (`record`, `check`, and `diff`) for reproducibility bundles, including `--against` workbook comparison and `repro diff --format text|json --output <path>` artifact export.
4. Telemetry schema contract added:
   - `schemas/events/v1/envelope.schema.json`.
5. Artifact telemetry now includes diff output artifact events:
   - `artifact.bundle.diff.output`.
6. Calc telemetry includes dependency graph build events:
   - `calc.dependency_graph.built`.
   - Includes `function_call_count` metric in dependency graph analytics.
   - Includes AST interning metrics (`ast_node_count`, `ast_unique_node_count`) and payload introspection (`formula_ast_ids`, `ast_intern_preview`).
7. Recalc artifact telemetry includes DAG/dependency output events:
   - `artifact.recalc.dep_graph.output`, `artifact.recalc.dag_timing.output`.
8. Recalc telemetry includes per-node timing summary event:
   - `calc.recalc.dag_timing`.
   - Includes `critical_path_duration_us`, `max_fan_in`, `max_fan_out`, and slow-node metrics.
   - Includes `slow_nodes_threshold_override_us` payload context when configured.
9. Interop telemetry includes part-graph build/save diagnostics:
   - `interop.xlsx.part_graph.built`.
   - `interop.xlsx.load.end` and `interop.xlsx.save.end` now include part-graph metrics and save graph flags in payload.
10. Corpus-level part-graph telemetry includes aggregate run events:
   - `artifact.part_graph.corpus.start`, `artifact.part_graph.corpus.end`.
11. CI publication baseline added for corpus part-graph validation:
   - `.github/workflows/corpus-part-graph.yml` (PR/push/schedule/workflow_dispatch).
   - `python/generate_corpus_fixtures.py` deterministic corpus fixture generator for CI execution.
12. CI publication baseline added for reproducibility bundle validation:
   - `.github/workflows/repro-bundle.yml` (PR/push/schedule/workflow_dispatch).
   - Includes `repro record/check/diff` artifact publication and explicit mismatch assertion on mutated candidates.
13. CI artifact policy alignment completed across corpus + repro workflows:
   - Standardized artifact names: `rootcellar-<workflow>-<run_id>-<run_attempt>`.
   - Standardized retention window: `21` days.
   - Standardized manifest metadata (`retention_policy_days`, run/repo/ref/sha fields) in assembled artifact bundles.
14. Batch recalc CLI delivery with bounded parallel scheduling:
   - `batch recalc` command group for recursive `.xlsx` directory execution using bounded Rayon threadpool sizing (`--threads`).
   - Deterministic file ordering with aggregate batch artifact output (`--report`) and fail-fast gate (`--fail-on-errors`).
   - Detail-level artifact control for per-file recalc payload introspection (`--detail-level minimal|diagnostic|forensic`).
   - Batch telemetry events: `artifact.batch.recalc.start`, `artifact.batch.recalc.file`, `artifact.batch.recalc.end`.
15. Nightly batch CI publication baseline added:
   - `.github/workflows/batch-recalc-nightly.yml` (nightly schedule + manual dispatch).
   - Includes fixture generation, workspace tests, `batch recalc` execution, artifact sanity assertions, and standardized artifact/manifest upload policy.
16. Nightly batch CI coverage and throughput gating expanded:
   - `python/build_batch_nightly_corpus.py` assembles deterministic nightly corpus slices from generated fixtures plus curated workbook samples.
   - Nightly workflow now runs `batch recalc` against expanded corpus (`target-files=32`) and enforces minimum processed-file and throughput thresholds.
   - Batch report summary now includes throughput metrics (`throughput_files_per_sec`, `aggregate_file_time_ratio`) for CI/SLO introspection.
17. Nightly batch trend snapshots and alert-hook payloads delivered:
   - `python/build_batch_trend_snapshot.py` generates per-run throughput snapshots from batch reports with threshold and breach metadata.
   - Produces alert-hook payload artifact (`ci-batch-alert-hook.json`) with routing key, severity, and breach details for downstream incident tooling ingestion.
   - Nightly batch workflow now publishes snapshot + alert payload artifacts alongside report/event/corpus outputs and enforces threshold breaches via explicit gate step.
18. Nightly batch alert-route dispatch integration delivered:
   - `python/dispatch_batch_alert_hook.py` routes alert payloads to incident and dashboard endpoints.
   - Supports route policy (`incident on breach` by default), per-route status accounting, and dispatch report artifact output (`ci-batch-alert-dispatch.json`).
   - Nightly workflow dispatches payloads before gate enforcement and records route configuration metadata in manifest output.
19. Alert dispatch hardening completed:
   - Route auth support via token headers (`ROOTCELLAR_INCIDENT_WEBHOOK_TOKEN`, `ROOTCELLAR_DASHBOARD_INGEST_TOKEN`) and optional HMAC request signing (`ROOTCELLAR_ALERT_SIGNING_SECRET`).
   - Retry/backoff policy with configurable attempt limits and retryable status-code handling.
   - Per-route acknowledgement tracking (`ack_id` extraction, required-ack mode, and ack counters in dispatch report).
20. Cross-system alert traceability keys delivered:
   - Deterministic per-route idempotency keys propagated via `Idempotency-Key` request header and dispatch payload metadata.
   - Correlation IDs propagated via `X-Correlation-Id` header and payload metadata, with optional required downstream correlation-match checks.
   - Dispatch artifacts now expose correlation and idempotency metadata/counters for route-level forensic tracing.
21. Alert replay-protection and acknowledgement retention indexing delivered:
   - Dispatch requests now carry replay-protection policy fields (`X-RootCellar-Timestamp`, `X-RootCellar-Nonce`, `X-RootCellar-Replay-Window-Sec`) with per-attempt replay metadata persisted in dispatch route artifacts.
   - Nightly workflow now passes configurable replay policy knobs (`ALERT_DISPATCH_REPLAY_*`) into dispatch routing and captures policy values in artifact manifest metadata.
   - `python/build_batch_ack_retention_index.py` added to derive `ci-batch-ack-retention-index.json` from dispatch reports with retention expiry windows and lookup keys (`ack_id`, `ack_id_sha256`, `idempotency_key`, `correlation_id`) for incident forensics.
22. Nightly dashboard-pack and alert-policy wiring delivered:
   - `python/build_batch_dashboard_pack.py` added to derive `ci-batch-dashboard-pack.json` and `ci-batch-alert-policy.json` from snapshot + dispatch + ack-retention artifacts.
   - Policy checks now evaluate snapshot threshold status, dispatch failed/ack-missing/correlation-mismatch counts, replay metadata completeness, and ack-retention lookup coverage.
   - Nightly workflow now enforces policy gate status and publishes dashboard/policy artifacts plus manifest policy knobs for downstream dashboard/incident integration.
23. Policy-to-owner escalation metadata and dashboard adapter exports delivered:
   - `python/build_batch_policy_adapters.py` added to derive `ci-batch-policy-escalation.json` and `ci-batch-dashboard-adapter-exports.json` from alert-policy + dashboard-pack artifacts.
   - Escalation metadata now maps policy checks to owner teams/queues/channels with severity-targeted escalation SLA/targets for downstream incident-routing systems.
   - Adapter exports now include incident-adapter and dashboard-adapter payloads with breach/owner context and metric points for downstream ingestion.
   - Nightly workflow now publishes escalation/adapter artifacts and captures owner/escalation policy configuration in manifest metadata.
24. Adapter export schema validation and compatibility-version contracts delivered:
   - Added versioned artifact schemas:
     - `schemas/artifacts/v1/batch-policy-escalation.schema.json`
     - `schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json`
   - Added contract validator utility: `python/validate_batch_adapter_contracts.py` (schema-shape validation + compatibility-version checks on `artifact_contract` and payload version fields).
   - `python/build_batch_policy_adapters.py` now emits `artifact_contract` metadata (`schema_id`, `schema_version`, `compatibility`) in both adapter outputs.
   - Nightly workflow now runs adapter schema/contract validation and captures schema validation config in manifest metadata.
25. Full artifact-family schema validation and compatibility-version contracts delivered:
   - Added versioned artifact schemas:
     - `schemas/artifacts/v1/batch-throughput-snapshot.schema.json`
     - `schemas/artifacts/v1/batch-alert-dispatch.schema.json`
     - `schemas/artifacts/v1/batch-ack-retention-index.schema.json`
     - `schemas/artifacts/v1/batch-dashboard-pack.schema.json`
     - `schemas/artifacts/v1/batch-alert-policy.schema.json`
   - `python/build_batch_trend_snapshot.py`, `python/dispatch_batch_alert_hook.py`, `python/build_batch_ack_retention_index.py`, and `python/build_batch_dashboard_pack.py` now emit `artifact_contract` metadata for their artifacts.
   - `python/validate_batch_adapter_contracts.py` now supports `--full-family` mode to validate snapshot/dispatch/ack-retention/dashboard-pack/policy plus escalation/adapter artifacts in one compatibility gate.
   - Nightly workflow now validates the full artifact family (`Validate batch artifact schemas and version contracts`) and records all schema-path policy knobs in manifest metadata.
26. Schema-drift canary fixtures and migration playbook delivered:
   - Added canary harness utility: `python/validate_batch_schema_canaries.py` (mutates canonical nightly artifacts and asserts expected validator failures for schema-id mismatch, missing contract fields, semver-major mismatch, compatibility-mode mismatch, and payload-version mismatch).
   - Nightly workflow now runs explicit canary gate step (`Run schema drift canary checks`) with policy knob `ALERT_POLICY_SCHEMA_CANARY_VALIDATION_ENABLED`.
   - Manifest metadata now records canary gate policy state (`alert_policy_schema_canary_validation_enabled`).
   - Added compatibility rollout reference: `docs/RootCellar/04-Observability/Artifact Schema Migration Playbook.md`.
27. Baseline test suite passing in offline mode.

## Verification
- `cargo fmt --all`: pass.
- `cargo test --workspace --offline`: pass.
- `cargo run -p rootcellar-cli --offline -- --help`: pass.
- `cargo run -p rootcellar-cli --offline -- tx-demo --jsonl ./tmp-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- open ./sample.xlsx --jsonl ./sample-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- open ./sample-formula.xlsx --jsonl ./sample-open-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- recalc ./sample-formula.xlsx --jsonl ./sample-recalc-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- save ./sample-formula.xlsx ./normalized.xlsx --mode normalize --jsonl ./save-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- open ./normalized.xlsx --jsonl ./normalized-open-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- recalc ./normalized.xlsx --jsonl ./normalized-recalc-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- save ./preserve-source.xlsx ./preserved-output.xlsx --mode preserve --jsonl ./preserve-save-events.jsonl`: pass (unknown part retained).
- `cargo run -p rootcellar-cli --offline -- tx-save ./tx-source.xlsx ./tx-out-preserve.xlsx --sheet Sheet1 --cell A1 --value 77 --mode preserve --jsonl ./tx-save-events.jsonl`: pass (sheet updated + unknown part retained).
- `cargo run -p rootcellar-cli --offline -- tx-save ./tx-source.xlsx ./tx-out-normalize.xlsx --sheet Sheet1 --cell A1 --value 88 --mode normalize --jsonl ./tx-save-normalize-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- repro record ./sample-formula.xlsx --bundle ./repro-bundle --jsonl ./repro-record-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- repro check ./repro-bundle --jsonl ./repro-check-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- tx-save ./sample-formula.xlsx ./sample-formula-mutated.xlsx --sheet Sheet1 --set A1=111 --set A2=9 --set B1=world --mode preserve --jsonl ./tx-multiset-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- tx-save ./sample-formula.xlsx ./sample-formula-formulaedit.xlsx --sheet Sheet1 --set A1=50 --set A2=4 --setf A3==A1+A2 --mode preserve --jsonl ./tx-formula-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- repro check ./repro-bundle --against ./sample-formula.xlsx --jsonl ./repro-check-against-same-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- repro check ./repro-bundle --against ./sample-formula-mutated.xlsx --jsonl ./repro-check-against-mutated-events.jsonl`: expected fail (mismatch detected).
- `cargo run -p rootcellar-cli --offline -- repro diff ./repro-bundle --against ./sample-formula.xlsx --jsonl ./repro-diff-same-events.jsonl`: pass (no diffs).
- `cargo run -p rootcellar-cli --offline -- repro diff ./repro-bundle --against ./sample-formula-mutated.xlsx --jsonl ./repro-diff-mutated-events.jsonl`: pass (cell-level diffs reported).
- `cargo run -p rootcellar-cli --offline -- repro diff --help`: pass (`--output` option available).
- `cargo run -p rootcellar-cli --offline -- repro diff ./repro-bundle --against ./sample-formula-mutated.xlsx --format json --output ./repro-diff-output.json --jsonl ./repro-diff-output-events.jsonl`: pass (JSON diff artifact + output telemetry event).
- `cargo run -p rootcellar-cli --offline -- repro diff ./repro-bundle --against ./sample-formula-mutated.xlsx --output ./repro-diff-output.txt --jsonl ./repro-diff-output-text-events.jsonl`: pass (text diff artifact output).
- `cargo run -p rootcellar-cli --offline -- recalc ./sample-formula.xlsx --report ./sample-recalc-with-graph.json --dep-graph-report ./sample-dep-graph-report.json --jsonl ./sample-dep-graph-events.jsonl`: pass (recalc + dependency graph artifact output).
- `cargo run -p rootcellar-cli --offline -- repro check ./repro-bundle --jsonl ./repro-check-post-graph-events.jsonl`: pass (hash compatibility preserved after parser/graph internals update).
- `cargo run -p rootcellar-cli --offline -- tx-save ./sample-formula.xlsx ./incremental-seed.xlsx --sheet Sheet1 --set A1=10 --set B1=5 --setf C1==A1+B1 --setf D1==C1*2 --setf E1==B1*3 --mode preserve --jsonl ./incremental-seed-events.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- tx-save ./incremental-seed.xlsx ./incremental-mutated.xlsx --sheet Sheet1 --set A1=20 --mode preserve --jsonl ./incremental-mutated-events.jsonl`: pass (`calc.recalc.incremental.*` emitted; evaluated 3/4 formulas based on changed roots).
- `cargo run -p rootcellar-cli --offline -- recalc --help`: pass (`--dag-timing-report` and `--dag-slow-threshold-us` options available).
- `cargo run -p rootcellar-cli --offline -- recalc ./sample-formula.xlsx --report ./sample-recalc-with-dag.json --dep-graph-report ./sample-dep-graph-report-v2.json --dag-timing-report ./sample-dag-timing-report.json --jsonl ./sample-dag-timing-events.jsonl`: pass (recalc + dependency graph + DAG timing artifacts output).
- `cargo run -p rootcellar-cli --offline -- tx-save ./sample-formula.xlsx ./sample-formula-incremental-check.xlsx --sheet Sheet1 --set A1=33 --mode preserve --jsonl ./sample-formula-incremental-check-events.jsonl`: pass (incremental recalc path remains active after DAG timing changes).
- `cargo run -p rootcellar-cli --offline -- tx-save ./sample-formula.xlsx ./sample-function-coverage.xlsx --sheet Sheet1 --set A1=8 --set B1=3 --setf "C1==SUM(A1,B1,2)" --setf "D1==MIN(A1,B1,2)" --setf "E1==MAX(A1,B1,2)" --setf "F1==IF(A1-B1,10,20)" --mode preserve --jsonl ./sample-function-coverage-events.jsonl`: pass (function baseline formulas evaluated without parse errors).
- `cargo run -p rootcellar-cli --offline -- recalc ./sample-function-coverage.xlsx --report ./sample-function-recalc-report.json --dep-graph-report ./sample-function-dep-graph-report.json --dag-timing-report ./sample-function-dag-timing-report.json --jsonl ./sample-function-recalc-events.jsonl`: pass (`function_call_count=4` observed in dependency graph telemetry + report).
- `cargo run -p rootcellar-cli --offline -- recalc ./sample-function-coverage.xlsx --report ./sample-function-recalc-report-v2.json --dep-graph-report ./sample-function-dep-graph-report-v2.json --dag-timing-report ./sample-function-dag-timing-report-v2.json --jsonl ./sample-function-recalc-events-v2.jsonl`: pass (AST interning metrics/IDs observed: `ast_node_count=21`, `ast_unique_node_count=12`, `formula_ast_ids` present).
- `cargo run -p rootcellar-cli --offline -- recalc ./sample-function-coverage.xlsx --report ./sample-function-recalc-report-v3.json --dep-graph-report ./sample-function-dep-graph-report-v3.json --dag-timing-report ./sample-function-dag-timing-report-v3.json --jsonl ./sample-function-recalc-events-v3.jsonl`: pass (DAG analysis fields present in timing artifacts/events).
- `cargo run -p rootcellar-cli --offline -- recalc ./incremental-seed.xlsx --report ./incremental-seed-recalc-report-v2.json --dep-graph-report ./incremental-seed-dep-graph-report-v2.json --dag-timing-report ./incremental-seed-dag-timing-report-v2.json --jsonl ./incremental-seed-recalc-events-v2.jsonl`: pass (critical path and fan-in/fan-out confirmed: `critical_path=[C1,D1]`, `max_fan_in=1`, `max_fan_out=1`).
- `cargo run -p rootcellar-cli --offline -- tx-save ./sample-formula.xlsx ./sample-function-coverage-v2.xlsx --sheet Sheet1 --set A1=8 --set B1=3 --setf "C1==SUM(A1,B1,2)" --setf "D1==MIN(A1,B1,2)" --setf "E1==MAX(A1,B1,2)" --setf "F1==IF(A1-B1,10,20)" --setf "G1==AVERAGE(A1,B1,1)" --setf "H1==ABS(B1-A1)" --setf "I1==AND(A1,B1,1)" --setf "J1==OR(0,0,B1-A1)" --setf "K1==NOT(B1-A1)" --setf "L1==NOT(0)" --mode preserve --jsonl ./sample-function-coverage-v2-events.jsonl`: pass (expanded built-ins evaluate with no parse errors).
- `cargo run -p rootcellar-cli --offline -- recalc ./sample-function-coverage-v2.xlsx --report ./sample-function-recalc-report-v4.json --dep-graph-report ./sample-function-dep-graph-report-v4.json --dag-timing-report ./sample-function-dag-timing-report-v4.json --dag-slow-threshold-us 10 --jsonl ./sample-function-recalc-events-v4.jsonl`: pass (`function_call_count=10`, `slow_nodes_threshold_us=10`, threshold override surfaced in telemetry).
- `cargo run -p rootcellar-cli --offline -- recalc ./sample-function-coverage-v2.xlsx --dag-slow-threshold-us 5`: expected fail (`--dag-slow-threshold-us requires --dag-timing-report`).
- `cargo run -p rootcellar-cli --offline -- open ./sample-formula.xlsx --report ./sample-formula.rootcellar-report-v2.json --jsonl ./sample-open-events-v2.jsonl`: pass (inspection report now includes `part_graph`, and telemetry emits `interop.xlsx.part_graph.built`).
- `cargo run -p rootcellar-cli --offline -- save ./sample-formula.xlsx ./normalized-v2.xlsx --mode normalize --jsonl ./save-events-v2.jsonl`: pass (save report includes normalized graph flags and graph metrics).
- `cargo run -p rootcellar-cli --offline -- tx-save ./tx-source.xlsx ./tx-out-preserve-v2.xlsx --sheet Sheet1 --cell A1 --value 77 --mode preserve --jsonl ./tx-save-events-v2.jsonl`: pass (save report includes preserve graph flags and graph metrics).
- `cargo run -p rootcellar-cli --offline -- --help`: pass (`part-graph-corpus` command available).
- `cargo run -p rootcellar-cli --offline -- part-graph-corpus . --max-files 8 --report ./part-graph-corpus-report-v1.json --jsonl ./part-graph-corpus-events-v1.jsonl`: pass (aggregate corpus report + telemetry with `discovered_files=20`, `processed_files=8`, `failure_count=0`).
- `cargo run -p rootcellar-cli --offline -- part-graph-corpus ./tmp-corpus-invalid --report ./tmp-corpus-invalid-report.json --fail-on-errors --jsonl ./tmp-corpus-invalid-events.jsonl`: expected fail (non-zero exit for malformed workbook; `failure_count=1` captured in artifact/events).
- `python python/generate_corpus_fixtures.py --output-dir ./.ci/corpus-fixtures`: pass (deterministic fixture corpus generated: 4 files).
- `cargo run -p rootcellar-cli --offline -- part-graph-corpus ./.ci/corpus-fixtures --fail-on-errors --report ./ci-part-graph-corpus-report-v2.json --jsonl ./ci-part-graph-corpus-events-v2.jsonl`: pass (`failure_count=0`, `total_dangling_edges=1`, `total_unknown_part_count=1`).
- `cargo run -p rootcellar-cli --offline -- part-graph-corpus --help`: pass (`--fail-on-errors`, `--max-files`, and `--report` options available).
- `python python/generate_corpus_fixtures.py --output-dir ./.ci/repro-fixtures`: pass (repro fixture set generated).
- `cargo run -p rootcellar-cli --offline -- repro record ./.ci/repro-fixtures/formula.xlsx --bundle ./ci-repro-bundle-v1 --jsonl ./ci-repro-record-events-v1.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- repro check ./ci-repro-bundle-v1 --jsonl ./ci-repro-check-events-v1.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- repro diff ./ci-repro-bundle-v1 --against ./.ci/repro-fixtures/formula.xlsx --output ./ci-repro-diff-same-v1.txt --jsonl ./ci-repro-diff-same-events-v1.jsonl`: pass (no diffs).
- `cargo run -p rootcellar-cli --offline -- tx-save ./.ci/repro-fixtures/formula.xlsx ./.ci/repro-fixtures/formula-mutated.xlsx --sheet Sheet1 --cell A1 --value 42 --mode preserve --jsonl ./ci-repro-tx-save-events-v1.jsonl`: pass.
- `cargo run -p rootcellar-cli --offline -- repro diff ./ci-repro-bundle-v1 --against ./.ci/repro-fixtures/formula-mutated.xlsx --format json --output ./ci-repro-diff-mutated-v1.json --jsonl ./ci-repro-diff-mutated-events-v1.jsonl`: pass (changed cells reported).
- `cargo run -p rootcellar-cli --offline -- repro check ./ci-repro-bundle-v1 --against ./.ci/repro-fixtures/formula-mutated.xlsx --jsonl ./ci-repro-check-mutated-events-v1.jsonl`: expected fail (mismatch detected for mutated workbook).
- `rg -n "retention-days|retention_policy_days|rootcellar-corpus-part-graph|rootcellar-repro-bundle" .github/workflows/corpus-part-graph.yml .github/workflows/repro-bundle.yml`: pass (policy alignment markers present in both workflows).
- `cargo run -p rootcellar-cli --offline -- batch --help`: pass (`recalc` subcommand available).
- `cargo run -p rootcellar-cli --offline -- batch recalc --help`: pass (`--threads`, `--detail-level`, and `--fail-on-errors` options available).
- `python python/generate_corpus_fixtures.py --output-dir ./.ci/batch-fixtures`: pass (batch fixture set generated).
- `cargo run -p rootcellar-cli --offline -- batch recalc ./.ci/batch-fixtures --threads 2 --detail-level diagnostic --fail-on-errors --report ./ci-batch-recalc-report-v1.json --jsonl ./ci-batch-recalc-events-v1.jsonl`: pass (processed=4, failures=0, diagnostic per-file recalc payloads captured).
- `cargo run -p rootcellar-cli --offline -- batch recalc ./tmp-corpus-invalid --threads 2 --fail-on-errors --report ./tmp-batch-invalid-report-v1.json --jsonl ./tmp-batch-invalid-events-v1.jsonl`: expected fail (malformed workbook detected; failure surfaced in report/events and non-zero exit asserted).
- `rg -n "retention-days|retention_policy_days|rootcellar-batch-recalc|artifact_kind|batch-recalc-nightly" .github/workflows/batch-recalc-nightly.yml`: pass (nightly batch artifact policy markers present).
- `python python/build_batch_nightly_corpus.py --output-dir ./.ci/batch-nightly-corpus --target-files 32`: pass (expanded deterministic corpus assembled from 7 seeds into 32 files).
- `cargo run -p rootcellar-cli --offline -- batch recalc ./.ci/batch-nightly-corpus/files --threads 4 --detail-level diagnostic --fail-on-errors --report ./ci-batch-recalc-report-v2.json --jsonl ./ci-batch-recalc-events-v2.jsonl`: pass (`processed=32`, `failures=0`, `throughput_files_per_sec=320.0`).
- `python -c "import json; s=json.load(open('./ci-batch-recalc-report-v2.json','r',encoding='utf-8'))['summary']; assert s['processed_files'] >= 24; assert s['failure_count'] == 0; assert s['throughput_files_per_sec'] >= 5.0"`: pass (nightly threshold policy locally validated).
- `rg -n "BATCH_MIN_PROCESSED_FILES|BATCH_MIN_THROUGHPUT_FILES_PER_SEC|build_batch_nightly_corpus|batch-nightly-corpus/files" .github/workflows/batch-recalc-nightly.yml`: pass (expanded corpus + throughput threshold gates present).
- `python python/build_batch_trend_snapshot.py --report ./ci-batch-recalc-report-v2.json --snapshot ./ci-batch-throughput-snapshot-v1.json --alert ./ci-batch-alert-hook-v1.json --min-processed-files 24 --min-throughput-files-per-sec 5.0 --fail-on-breach`: pass (snapshot status `pass`, alert severity `info`).
- `python python/build_batch_trend_snapshot.py --report ./ci-batch-recalc-report-v2.json --snapshot ./ci-batch-throughput-snapshot-breach-v1.json --alert ./ci-batch-alert-hook-breach-v1.json --min-processed-files 128 --min-throughput-files-per-sec 500.0 --fail-on-breach`: expected fail (breach mode exercised; alert payload includes threshold violations).
- `rg -n "build_batch_trend_snapshot|ci-batch-throughput-snapshot|ci-batch-alert-hook|fail-on-breach" .github/workflows/batch-recalc-nightly.yml`: pass (trend snapshot + alert-hook workflow integration markers present).
- `python python/build_batch_trend_snapshot.py --report ./does-not-exist-report.json --snapshot ./ci-batch-throughput-snapshot-missing-v1.json --alert ./ci-batch-alert-hook-missing-v1.json --min-processed-files 24 --min-throughput-files-per-sec 5.0 --allow-missing-report`: pass (missing report degraded into breach snapshot for downstream routing).
- `python python/dispatch_batch_alert_hook.py --alert ./ci-batch-alert-hook-v1.json --snapshot ./ci-batch-throughput-snapshot-v1.json --dispatch-report ./ci-batch-alert-dispatch-v1.json --fail-on-route-error`: pass (no routes configured; dispatch status `no_routes`).
- `python python/dispatch_batch_alert_hook.py --alert ./ci-batch-alert-hook-breach-v1.json --snapshot ./ci-batch-throughput-snapshot-breach-v1.json --incident-url http://127.0.0.1:8765/incident --dashboard-url http://127.0.0.1:8765/dashboard --dispatch-report ./ci-batch-alert-dispatch-local-v1.json --fail-on-route-error`: pass (local harness route simulation; 2 delivered routes).
- `python python/dispatch_batch_alert_hook.py --alert ./ci-batch-alert-hook-breach-v1.json --snapshot ./ci-batch-throughput-snapshot-breach-v1.json --incident-url http://127.0.0.1:9988/incident --dashboard-url http://127.0.0.1:9988/dashboard --dispatch-report ./ci-batch-alert-dispatch-fail-v1.json --fail-on-route-error`: expected fail (route delivery failures propagated for CI gating).
- `rg -n "dispatch_batch_alert_hook|ci-batch-alert-dispatch|ROOTCELLAR_INCIDENT_WEBHOOK_URL|ROOTCELLAR_DASHBOARD_INGEST_URL|allow-missing-report|Enforce nightly batch gate" .github/workflows/batch-recalc-nightly.yml`: pass (endpoint-routing and explicit gate integration markers present).
- `python python/dispatch_batch_alert_hook.py --alert ./ci-batch-alert-hook-breach-v1.json --snapshot ./ci-batch-throughput-snapshot-breach-v1.json --incident-url http://127.0.0.1:8770/incident --dashboard-url http://127.0.0.1:8770/dashboard --incident-auth-token incident-token --dashboard-auth-token dashboard-token --signing-secret super-secret-signing-key --max-attempts 3 --initial-backoff-sec 0.1 --backoff-multiplier 2.0 --max-backoff-sec 0.2 --require-ack-on-incident --require-ack-on-dashboard --dispatch-report ./ci-batch-alert-dispatch-auth-retry-ack-v1.json --fail-on-route-error`: pass (local harness confirmed auth headers, HMAC signature, retry-on-503, and ack-required success).
- `python python/dispatch_batch_alert_hook.py --alert ./ci-batch-alert-hook-breach-v1.json --snapshot ./ci-batch-throughput-snapshot-breach-v1.json --incident-url http://127.0.0.1:8771/incident --dashboard-url http://127.0.0.1:8771/dashboard --require-ack-on-dashboard --dispatch-report ./ci-batch-alert-dispatch-ack-missing-v1.json --fail-on-route-error`: expected fail (ack-required violation produces failed route and non-zero exit).
- `python python/dispatch_batch_alert_hook.py --alert ./ci-batch-alert-hook-breach-v1.json --snapshot ./ci-batch-throughput-snapshot-breach-v1.json --incident-url http://127.0.0.1:8772/incident --dashboard-url http://127.0.0.1:8772/dashboard --incident-auth-token incident-token --dashboard-auth-token dashboard-token --signing-secret super-secret-signing-key --max-attempts 3 --initial-backoff-sec 0.1 --backoff-multiplier 2.0 --max-backoff-sec 0.2 --require-ack-on-incident --require-ack-on-dashboard --require-correlation-on-incident --require-correlation-on-dashboard --dispatch-report ./ci-batch-alert-dispatch-auth-retry-corr-v1.json --fail-on-route-error`: pass (local harness confirmed idempotency/correlation headers + retry + ack/correlation matches).
- `python python/dispatch_batch_alert_hook.py --alert ./ci-batch-alert-hook-breach-v1.json --snapshot ./ci-batch-throughput-snapshot-breach-v1.json --incident-url http://127.0.0.1:8773/incident --dashboard-url http://127.0.0.1:8773/dashboard --require-correlation-on-incident --require-correlation-on-dashboard --dispatch-report ./ci-batch-alert-dispatch-corr-mismatch-v1.json --fail-on-route-error`: expected fail (correlation mismatch enforcement produces failed routes and non-zero exit).
- `python python/build_batch_ack_retention_index.py --dispatch-report ./ci-batch-alert-dispatch-auth-retry-corr-v1.json --index ./ci-batch-ack-retention-index-v1.json --retention-days 30`: pass (`record_count=2`, lookup keys emitted for `ack_id`/idempotency/correlation).
- `python python/build_batch_ack_retention_index.py --dispatch-report ./does-not-exist-dispatch.json --index ./ci-batch-ack-retention-index-missing-v1.json --retention-days 30 --allow-missing-dispatch-report`: pass (degraded missing-dispatch index emitted for always-on workflow artifact continuity).
- `python python/dispatch_batch_alert_hook.py --alert ./ci-batch-alert-hook-breach-v1.json --snapshot ./ci-batch-throughput-snapshot-breach-v1.json --incident-url http://127.0.0.1:9989/incident --dashboard-url http://127.0.0.1:9989/dashboard --dispatch-report ./ci-batch-alert-dispatch-fail-replay-v1.json --max-attempts 2 --initial-backoff-sec 0.1 --max-backoff-sec 0.1 --fail-on-route-error ; python python/build_batch_ack_retention_index.py --dispatch-report ./ci-batch-alert-dispatch-fail-replay-v1.json --index ./ci-batch-ack-retention-index-fail-route-v1.json --retention-days 30`: pass (index generation remains available for forensic triage after dispatch partial-failure paths).
- Inline local Python HTTP harness (port `8780`) + `python/dispatch_batch_alert_hook.py`: pass (replay headers present on all requests and replay nonce remained unique across dashboard retry attempts).
- `rg -n "ALERT_DISPATCH_REPLAY_|ALERT_ACK_RETENTION_DAYS|build_batch_ack_retention_index|ci-batch-ack-retention-index|replay-window" .github/workflows/batch-recalc-nightly.yml python/dispatch_batch_alert_hook.py python/build_batch_ack_retention_index.py`: pass (replay-policy and ack-retention indexing integration markers present).
- `python python/build_batch_dashboard_pack.py --snapshot ./ci-batch-throughput-snapshot-v1.json --dispatch-report ./ci-batch-alert-dispatch-v1.json --ack-retention-index ./ci-batch-ack-retention-index-no-routes-v2.json --dashboard-pack ./ci-batch-dashboard-pack-no-routes-v1.json --policy ./ci-batch-alert-policy-no-routes-v1.json --max-dispatch-failed-routes 0 --max-ack-missing-routes 0 --max-correlation-mismatch-routes 0 --require-replay-metadata --require-ack-retention-coverage --allow-missing-inputs --fail-on-policy-breach`: pass (policy status `pass` for no-route baseline).
- `python python/build_batch_dashboard_pack.py --snapshot ./ci-batch-throughput-snapshot-v1.json --dispatch-report ./ci-batch-alert-dispatch-replay-v1.json --ack-retention-index ./ci-batch-ack-retention-index-replay-v1.json --dashboard-pack ./ci-batch-dashboard-pack-pass-v1.json --policy ./ci-batch-alert-policy-pass-v1.json --max-dispatch-failed-routes 0 --max-ack-missing-routes 0 --max-correlation-mismatch-routes 0 --require-replay-metadata --require-ack-retention-coverage --allow-missing-inputs --fail-on-policy-breach`: pass (healthy delivered-route path with replay metadata and ack-retention coverage).
- `python python/build_batch_dashboard_pack.py --snapshot ./ci-batch-throughput-snapshot-breach-v1.json --dispatch-report ./ci-batch-alert-dispatch-fail-replay-v1.json --ack-retention-index ./ci-batch-ack-retention-index-fail-route-v1.json --dashboard-pack ./ci-batch-dashboard-pack-breach-fail-v1.json --policy ./ci-batch-alert-policy-breach-fail-v1.json --max-dispatch-failed-routes 0 --max-ack-missing-routes 0 --max-correlation-mismatch-routes 0 --require-replay-metadata --require-ack-retention-coverage --allow-missing-inputs --fail-on-policy-breach`: expected fail (policy breach returns non-zero while still writing introspection artifacts).
- `python python/build_batch_dashboard_pack.py --snapshot ./missing-snapshot.json --dispatch-report ./missing-dispatch.json --ack-retention-index ./missing-ack-index.json --dashboard-pack ./ci-batch-dashboard-pack-missing-v1.json --policy ./ci-batch-alert-policy-missing-v1.json --max-dispatch-failed-routes 0 --max-ack-missing-routes 0 --max-correlation-mismatch-routes 0 --allow-missing-inputs`: pass (degraded policy/dashboard artifacts produced for missing-input forensic continuity).
- `rg -n "ALERT_POLICY_|build_batch_dashboard_pack|ci-batch-dashboard-pack|ci-batch-alert-policy|policy gate" .github/workflows/batch-recalc-nightly.yml python/build_batch_dashboard_pack.py`: pass (dashboard-pack and policy-gate workflow integration markers present).
- `python -m py_compile python/build_batch_nightly_corpus.py python/build_batch_trend_snapshot.py python/dispatch_batch_alert_hook.py python/build_batch_ack_retention_index.py python/build_batch_dashboard_pack.py python/build_batch_policy_adapters.py python/validate_batch_adapter_contracts.py python/validate_batch_schema_canaries.py`: pass (script syntax checks with full-family validator and canary harness included).
- `python python/build_batch_policy_adapters.py --policy ./ci-batch-alert-policy-pass-v1.json --dashboard-pack ./ci-batch-dashboard-pack-pass-v1.json --escalation ./ci-batch-policy-escalation-pass-v2.json --adapter-exports ./ci-batch-dashboard-adapter-exports-pass-v2.json`: pass (adapter outputs regenerated with artifact-contract metadata).
- `python python/build_batch_policy_adapters.py --policy ./ci-batch-alert-policy-breach-v1.json --dashboard-pack ./ci-batch-dashboard-pack-breach-v1.json --escalation ./ci-batch-policy-escalation-breach-v2.json --adapter-exports ./ci-batch-dashboard-adapter-exports-breach-v2.json`: pass.
- `python python/build_batch_policy_adapters.py --policy ./missing-policy.json --dashboard-pack ./missing-dashboard-pack.json --escalation ./ci-batch-policy-escalation-missing-v2.json --adapter-exports ./ci-batch-dashboard-adapter-exports-missing-v2.json --allow-missing-inputs`: pass.
- `python python/build_batch_policy_adapters.py --policy ./ci-batch-alert-policy-pass-v1.json --dashboard-pack ./ci-batch-dashboard-pack-pass-v1.json --escalation ./ci-batch-policy-escalation.json --adapter-exports ./ci-batch-dashboard-adapter-exports.json ; python python/validate_batch_adapter_contracts.py`: pass (canonical nightly artifact names validate against default schema paths).
- `python python/validate_batch_adapter_contracts.py --escalation ./ci-batch-policy-escalation-pass-v2.json --adapter-exports ./ci-batch-dashboard-adapter-exports-pass-v2.json --schema-escalation ./schemas/artifacts/v1/batch-policy-escalation.schema.json --schema-adapter-exports ./schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json`: pass.
- `python python/validate_batch_adapter_contracts.py --escalation ./ci-batch-policy-escalation-breach-v2.json --adapter-exports ./ci-batch-dashboard-adapter-exports-breach-v2.json --schema-escalation ./schemas/artifacts/v1/batch-policy-escalation.schema.json --schema-adapter-exports ./schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json`: pass.
- `python python/validate_batch_adapter_contracts.py --escalation ./ci-batch-policy-escalation-missing-v2.json --adapter-exports ./ci-batch-dashboard-adapter-exports-missing-v2.json --schema-escalation ./schemas/artifacts/v1/batch-policy-escalation.schema.json --schema-adapter-exports ./schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json`: pass.
- `python python/validate_batch_adapter_contracts.py --escalation ./ci-batch-policy-escalation-pass-v1.json --adapter-exports ./ci-batch-dashboard-adapter-exports-pass-v1.json --schema-escalation ./schemas/artifacts/v1/batch-policy-escalation.schema.json --schema-adapter-exports ./schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json`: expected fail (legacy artifacts without `artifact_contract` rejected by compatibility contract checks).
- `rg -n "validate_batch_adapter_contracts|ALERT_POLICY_SCHEMA_|batch-policy-escalation.schema.json|batch-dashboard-adapter-exports.schema.json" .github/workflows/batch-recalc-nightly.yml python/validate_batch_adapter_contracts.py schemas/artifacts/v1/batch-policy-escalation.schema.json schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json`: pass (adapter schema-validation integration markers present).
- `python python/build_batch_trend_snapshot.py --report ./ci-batch-recalc-report-v2.json --snapshot ./ci-batch-throughput-snapshot.json --alert ./ci-batch-alert-hook.json --min-processed-files 24 --min-throughput-files-per-sec 5.0 --allow-missing-report ; python python/dispatch_batch_alert_hook.py --alert ./ci-batch-alert-hook.json --snapshot ./ci-batch-throughput-snapshot.json --dispatch-report ./ci-batch-alert-dispatch.json --fail-on-route-error ; python python/build_batch_ack_retention_index.py --dispatch-report ./ci-batch-alert-dispatch.json --index ./ci-batch-ack-retention-index.json --retention-days 30 --allow-missing-dispatch-report ; python python/build_batch_dashboard_pack.py --snapshot ./ci-batch-throughput-snapshot.json --dispatch-report ./ci-batch-alert-dispatch.json --ack-retention-index ./ci-batch-ack-retention-index.json --dashboard-pack ./ci-batch-dashboard-pack.json --policy ./ci-batch-alert-policy.json --max-dispatch-failed-routes 0 --max-ack-missing-routes 0 --max-correlation-mismatch-routes 0 --require-replay-metadata --require-ack-retention-coverage --allow-missing-inputs --fail-on-policy-breach ; python python/build_batch_policy_adapters.py --policy ./ci-batch-alert-policy.json --dashboard-pack ./ci-batch-dashboard-pack.json --escalation ./ci-batch-policy-escalation.json --adapter-exports ./ci-batch-dashboard-adapter-exports.json --allow-missing-inputs ; python python/validate_batch_adapter_contracts.py --full-family --snapshot ./ci-batch-throughput-snapshot.json --dispatch ./ci-batch-alert-dispatch.json --ack-retention ./ci-batch-ack-retention-index.json --dashboard-pack ./ci-batch-dashboard-pack.json --policy ./ci-batch-alert-policy.json --escalation ./ci-batch-policy-escalation.json --adapter-exports ./ci-batch-dashboard-adapter-exports.json --schema-snapshot ./schemas/artifacts/v1/batch-throughput-snapshot.schema.json --schema-dispatch ./schemas/artifacts/v1/batch-alert-dispatch.schema.json --schema-ack-retention ./schemas/artifacts/v1/batch-ack-retention-index.schema.json --schema-dashboard-pack ./schemas/artifacts/v1/batch-dashboard-pack.schema.json --schema-policy ./schemas/artifacts/v1/batch-alert-policy.schema.json --schema-escalation ./schemas/artifacts/v1/batch-policy-escalation.schema.json --schema-adapter-exports ./schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json`: pass (full artifact-family schema/contract gate validated on canonical nightly artifact names).
- `python python/validate_batch_adapter_contracts.py --full-family --snapshot ./ci-batch-throughput-snapshot-v1.json --dispatch ./ci-batch-alert-dispatch-v1.json --ack-retention ./ci-batch-ack-retention-index-v1.json --dashboard-pack ./ci-batch-dashboard-pack-pass-v1.json --policy ./ci-batch-alert-policy-pass-v1.json --escalation ./ci-batch-policy-escalation-pass-v1.json --adapter-exports ./ci-batch-dashboard-adapter-exports-pass-v1.json --schema-snapshot ./schemas/artifacts/v1/batch-throughput-snapshot.schema.json --schema-dispatch ./schemas/artifacts/v1/batch-alert-dispatch.schema.json --schema-ack-retention ./schemas/artifacts/v1/batch-ack-retention-index.schema.json --schema-dashboard-pack ./schemas/artifacts/v1/batch-dashboard-pack.schema.json --schema-policy ./schemas/artifacts/v1/batch-alert-policy.schema.json --schema-escalation ./schemas/artifacts/v1/batch-policy-escalation.schema.json --schema-adapter-exports ./schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json`: expected fail (legacy artifacts rejected for missing `artifact_contract` and missing full-family required fields).
- `rg -n "validate_batch_adapter_contracts|--full-family|ALERT_POLICY_SCHEMA_SNAPSHOT_PATH|ALERT_POLICY_SCHEMA_DISPATCH_PATH|ALERT_POLICY_SCHEMA_ACK_RETENTION_PATH|ALERT_POLICY_SCHEMA_DASHBOARD_PACK_PATH|ALERT_POLICY_SCHEMA_POLICY_PATH|Validate batch artifact schemas and version contracts" .github/workflows/batch-recalc-nightly.yml python/validate_batch_adapter_contracts.py`: pass (full-family schema-validation integration markers present).
- `python python/validate_batch_schema_canaries.py --validator-script ./python/validate_batch_adapter_contracts.py --snapshot ./ci-batch-throughput-snapshot.json --dispatch ./ci-batch-alert-dispatch.json --ack-retention ./ci-batch-ack-retention-index.json --dashboard-pack ./ci-batch-dashboard-pack.json --policy ./ci-batch-alert-policy.json --escalation ./ci-batch-policy-escalation.json --adapter-exports ./ci-batch-dashboard-adapter-exports.json --schema-snapshot ./schemas/artifacts/v1/batch-throughput-snapshot.schema.json --schema-dispatch ./schemas/artifacts/v1/batch-alert-dispatch.schema.json --schema-ack-retention ./schemas/artifacts/v1/batch-ack-retention-index.schema.json --schema-dashboard-pack ./schemas/artifacts/v1/batch-dashboard-pack.schema.json --schema-policy ./schemas/artifacts/v1/batch-alert-policy.schema.json --schema-escalation ./schemas/artifacts/v1/batch-policy-escalation.schema.json --schema-adapter-exports ./schemas/artifacts/v1/batch-dashboard-adapter-exports.schema.json`: pass (baseline pass + six deterministic drift scenarios failed as expected).
- `rg -n "ALERT_POLICY_SCHEMA_CANARY_VALIDATION_ENABLED|Run schema drift canary checks|alert_policy_schema_canary_validation_enabled|validate_batch_schema_canaries.py" .github/workflows/batch-recalc-nightly.yml python/validate_batch_schema_canaries.py`: pass (canary gate integration markers present).

## Next Execution Slice
1. Continue function parity expansion beyond current starter set (lookup/text/date families) with compatibility-focused semantics.
2. Harden incremental scheduler/perf for larger dependency graphs and continue DAG introspection tuning.
3. Add dual-read migration drills for future artifact major-version rollouts (producer/consumer overlap and rollback verification).
