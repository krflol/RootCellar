# RootCellar

Execution has started with a Rust workspace baseline focused on Sprint 00 through Sprint 02 foundations plus Sprint 06 headless batch/repro flows.

## Workspace
- `crates/rootcellar-core`: workbook transaction model, telemetry envelopes/sinks, XLSX inspection/report generation.
- `crates/rootcellar-cli`: command-line interface for inspection/reporting and transaction demo flow.
- `schemas/events/v1`: JSON schema contract for event envelopes.
- `schemas/artifacts/v1`: JSON schema contracts for nightly batch artifact family outputs.

## Quick Start
```bash
cargo test
cargo run -p rootcellar-cli -- --help
cargo run -p rootcellar-cli -- open ./example.xlsx --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- part-graph-corpus ./corpus --report ./part-graph-corpus-report.json --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- batch recalc ./corpus --threads 4 --report ./batch-recalc-report.json --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- batch recalc ./corpus --threads 4 --detail-level diagnostic --fail-on-errors --report ./batch-recalc-report-diagnostic.json --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- recalc ./example.xlsx --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- recalc ./example.xlsx --dep-graph-report ./dep-graph.json --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- recalc ./example.xlsx --dag-timing-report ./dag-timing.json --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- recalc ./example.xlsx --dag-timing-report ./dag-timing.json --dag-slow-threshold-us 10 --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- save ./example.xlsx ./normalized.xlsx --mode normalize --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- repro record ./example.xlsx --bundle ./repro-bundle --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- repro check ./repro-bundle --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- repro check ./repro-bundle --against ./candidate.xlsx --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- repro diff ./repro-bundle --against ./candidate.xlsx --limit 25 --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- repro diff ./repro-bundle --against ./candidate.xlsx --format json --output ./diff.json --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- tx-save ./example.xlsx ./edited.xlsx --sheet Sheet1 --set A1=123 --set B1=true --mode preserve --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- tx-save ./example.xlsx ./edited.xlsx --sheet Sheet1 --setf C1==A1+B1 --mode preserve --jsonl ./events.jsonl
python python/build_batch_ack_retention_index.py --dispatch-report ./ci-batch-alert-dispatch.json --index ./ci-batch-ack-retention-index.json --retention-days 30
python python/build_batch_dashboard_pack.py --snapshot ./ci-batch-throughput-snapshot.json --dispatch-report ./ci-batch-alert-dispatch.json --ack-retention-index ./ci-batch-ack-retention-index.json --dashboard-pack ./ci-batch-dashboard-pack.json --policy ./ci-batch-alert-policy.json --require-replay-metadata --require-ack-retention-coverage
python python/build_batch_policy_adapters.py --policy ./ci-batch-alert-policy.json --dashboard-pack ./ci-batch-dashboard-pack.json --escalation ./ci-batch-policy-escalation.json --adapter-exports ./ci-batch-dashboard-adapter-exports.json
python python/validate_batch_adapter_contracts.py --full-family
python python/validate_batch_schema_canaries.py
python python/validate_batch_dual_read_migration.py
python python/validate_batch_migration_policy_dry_run.py
```

## Current Status
- Foundation telemetry and trace context primitives are in place.
- Engine model supports transactional mutations with stable commit digest ordering.
- XLSX inspector reports core required parts and compatibility findings.
- XLSX inspection reports now include an inspectable workbook part graph (nodes/edges + dangling relationship detection).
- CLI includes `part-graph-corpus` for aggregate part-graph validation over workbook directories.
- CI baseline includes `.github/workflows/corpus-part-graph.yml` to publish corpus part-graph artifacts on PR/main/nightly runs.
- CI baseline includes `.github/workflows/repro-bundle.yml` to publish reproducibility bundle artifacts on PR/main/nightly runs.
- CI baseline includes `.github/workflows/batch-recalc-nightly.yml` to publish nightly bounded-parallel batch recalc artifacts.
- CI workflows now use aligned artifact naming/retention policy and include manifest metadata.
- Batch recalc command surface (`batch recalc`) now supports bounded Rayon threadpool sizing, deterministic file ordering, and detail-level report payload control.
- Nightly batch corpus assembly utility (`python/build_batch_nightly_corpus.py`) builds deterministic compatibility slices for broader CI coverage.
- Batch recalc artifacts now include throughput summaries (`throughput_files_per_sec`, `aggregate_file_time_ratio`) used by nightly regression thresholds.
- Nightly throughput trend + alert utility (`python/build_batch_trend_snapshot.py`) emits dashboard-ready snapshot and alert-hook payload artifacts.
- Nightly alert dispatch utility (`python/dispatch_batch_alert_hook.py`) routes alert payloads to incident/dashboard endpoints and emits dispatch-status artifacts.
- Optional CI route secrets for dispatch: `ROOTCELLAR_INCIDENT_WEBHOOK_URL`, `ROOTCELLAR_DASHBOARD_INGEST_URL`.
- Optional dispatch auth/signing secrets: `ROOTCELLAR_INCIDENT_WEBHOOK_TOKEN`, `ROOTCELLAR_DASHBOARD_INGEST_TOKEN`, `ROOTCELLAR_ALERT_SIGNING_SECRET`.
- Dispatch utility supports retry/backoff controls and ack tracking for route-level observability.
- Dispatch utility propagates deterministic `Idempotency-Key` and `X-Correlation-Id` values for cross-system alert traceability.
- Dispatch utility propagates replay-protection metadata (`X-RootCellar-Timestamp`, `X-RootCellar-Nonce`, `X-RootCellar-Replay-Window-Sec`) per delivery attempt.
- Nightly workflow publishes `ci-batch-ack-retention-index.json` for ack-id/idempotency/correlation forensic lookups with retention-expiry metadata.
- Nightly dashboard/policy utility (`python/build_batch_dashboard_pack.py`) publishes `ci-batch-dashboard-pack.json` and `ci-batch-alert-policy.json` from snapshot/dispatch/forensic artifacts.
- Nightly escalation/adapter utility (`python/build_batch_policy_adapters.py`) publishes `ci-batch-policy-escalation.json` and `ci-batch-dashboard-adapter-exports.json` for downstream incident/dashboard ingestion integration.
- Nightly batch artifact schema validator (`python/validate_batch_adapter_contracts.py --full-family`) enforces schema shape + compatibility version contracts before artifact publication.
- Nightly schema-drift canary utility (`python/validate_batch_schema_canaries.py`) asserts expected validator failures for representative compatibility-regression scenarios.
- Nightly dual-read migration drill utility (`python/validate_batch_dual_read_migration.py`) verifies producer/consumer overlap and rollback behavior across snapshot/dispatch/ack-retention/dashboard-pack/policy/escalation/adapter artifacts, with optional subset targeting via `--artifacts`, staged rollout waves via `--wave-spec`, per-phase diagnostics via `--report`, and fault-injection scenarios via `--fault-injection --fault-scenarios`.
- Nightly migration-policy dry-run harness (`python/validate_batch_migration_policy_dry_run.py`) asserts invalid staged-wave specs and unsupported fault-scenario keys fail fast in CI policy validation.
- Nightly gate now enforces both throughput snapshot status and alert-policy status for route-delivery/forensic policy checks.
- Minimal calculation engine supports A1 references, arithmetic formulas, and cycle detection.
- Formula parser scaffold now supports precedence and parentheses for arithmetic recalc.
- Built-in function baseline now supports `SUM`, `MIN`, `MAX`, `IF`, `AVERAGE`/`AVG`, `ABS`, `AND`, `OR`, `NOT`, `LEN`, `CHOOSE`, `MATCH`, `EXACT`, `FIND`, `SEARCH`, `CODE`, `N`, `VALUE`, `DATEVALUE`, `TIMEVALUE`, `ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISLOGICAL`, `ISERROR`, `COUNT`, `COUNTA`, `COUNTBLANK`, `DATE`, `YEAR`, `MONTH`, `DAY`, `DAYS`, `TIME`, `HOUR`, `MINUTE`, `SECOND`, `EDATE`, `EOMONTH`, `WEEKDAY`, `WEEKNUM`, and `ISOWEEKNUM`.
- AST interning scaffold now exposes deduplicated formula-node IDs for parser introspection.
- Incremental recalc from changed roots is available in core and used by `tx-save` post-mutation workflows.
- Incremental recalc now reuses cached reverse-dependency indexes during impacted-cell discovery and DAG degree analysis to reduce repeated graph traversal overhead on larger formula sets.
- XLSX workbook loader projects worksheet values/formulas into the in-memory model for recalc workflows.
- XLSX saver writes workbook model back to `.xlsx` with deterministic sheet/cell ordering baseline.
- Preserve mode now uses passthrough copy semantics to retain unknown XML parts exactly.
- Transactional preserve save rewrites only changed worksheet parts while retaining untouched/unknown parts.
- Save artifacts expose graph-aware flags for `normalize` vs `preserve` strategies.
- Repro bundle workflows can record and check deterministic recalc artifacts, including `--against` comparisons for external workbook candidates.
- `repro diff` provides deterministic cell-level deltas (changed/added/removed) against bundle baselines and can write text/JSON artifacts via `--output`.
- `tx-save` supports repeated value (`--set`) and formula (`--setf`) mutations in one transaction.
- `recalc` can emit inspectable dependency graph artifacts via `--dep-graph-report`.
- `recalc` can emit per-node timing artifacts via `--dag-timing-report` and telemetry event `calc.recalc.dag_timing`.
- `recalc` supports configurable slow-node threshold overrides via `--dag-slow-threshold-us` when emitting DAG timing artifacts.
- Dependency graph observability includes `function_call_count` metrics for formula introspection.
- Dependency graph observability includes AST metrics (`ast_node_count`, `ast_unique_node_count`) and `formula_ast_ids`.
- DAG timing observability includes `critical_path`, `max_fan_in`, `max_fan_out`, and slow-node threshold diagnostics.

## Next Build Slice
- Continue function parity expansion beyond current starter set.
- Continue parser/evaluator and scheduler optimization work on top of the AST interning scaffold.
- Start desktop shell initialization and bridge UI->engine trace context propagation.
