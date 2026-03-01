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
cargo run -p rootcellar-cli -- bench recalc-synthetic --chains 16 --chain-length 256 --iterations 5 --changed-chain 1 --report ./bench-recalc-report.json --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- save ./example.xlsx ./preserved.xlsx --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- save ./example.xlsx ./normalized.xlsx --mode normalize --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- repro record ./example.xlsx --bundle ./repro-bundle --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- repro check ./repro-bundle --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- repro check ./repro-bundle --against ./candidate.xlsx --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- repro diff ./repro-bundle --against ./candidate.xlsx --limit 25 --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- repro diff ./repro-bundle --against ./candidate.xlsx --format json --output ./diff.json --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- tx-save ./example.xlsx ./edited.xlsx --sheet Sheet1 --set A1=123 --set B1=true --mode preserve --jsonl ./events.jsonl
cargo run -p rootcellar-cli -- tx-save ./example.xlsx ./edited.xlsx --sheet Sheet1 --setf C1==A1+B1 --mode preserve --jsonl ./events.jsonl
python -m pip install -r ./python/requirements-interop.txt
python python/assemble_excel_interop_corpus.py --repo-root . --output-dir ./target/excel-interop-corpus --excel-authored-manifest ./corpus/excel-authored/manifest.json --min-excel-authored-samples 5 --required-curated-feature formulas --required-curated-feature styles --required-curated-feature comments --required-curated-feature charts --required-curated-feature defined_names
python python/verify_excel_interop.py --workspace . --workdir ./target/excel-interop-gate --corpus-dir ./target/excel-interop-corpus --corpus-manifest ./target/excel-interop-corpus/manifest.json --max-corpus-files 32 --require-corpus-fixture styles.xlsx --require-corpus-fixture comments.xlsx --require-corpus-fixture chart.xlsx --require-corpus-fixture defined-names.xlsx --report ./target/excel-interop-gate-report.json
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
- CI baseline includes `.github/workflows/excel-interop.yml` to enforce bidirectional Excel/openpyxl interoperability on PR/main/nightly runs.
- CI baseline includes `.github/workflows/batch-recalc-nightly.yml` to publish nightly bounded-parallel batch recalc artifacts.
- CI workflows now use aligned artifact naming/retention policy and include manifest metadata.
- Batch recalc command surface (`batch recalc`) now supports bounded Rayon threadpool sizing, deterministic file ordering, and detail-level report payload control.
- Synthetic recalc benchmark command surface (`bench recalc-synthetic`) emits repeatable full-vs-incremental performance reports on generated dependency workloads.
- Nightly batch CI now supports optional synthetic benchmark execution/gating via `BATCH_BENCH_*` policy knobs and publishes benchmark artifacts (`ci-batch-bench-recalc-synthetic.json`, `ci-batch-bench-events.jsonl`) in the bundle.
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
- Formula parser scaffold now supports precedence and parentheses for arithmetic recalc, plus quoted text literals and boolean constants (`TRUE`/`FALSE`).
- Built-in function baseline now supports `SUM`, `SUMSQ`, `PRODUCT`, `MIN`, `MAX`, `MEDIAN`, `SMALL`, `LARGE`, `GEOMEAN`, `HARMEAN`, `VARP`, `VAR`/`VARS`, `STDEVP`, `STDEV`/`STDEVS`, `IF`, `IFERROR`, `IFS`, `SWITCH`, `AVERAGE`/`AVG`, `ABS`, `INT`, `FACT`, `FACTDOUBLE`, `COMBIN`, `PERMUT`, `GCD`, `LCM`, `QUOTIENT`, `MOD`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `MROUND`, `POWER`, `SQRT`, `SIGN`, `EVEN`, `ODD`, `ISEVEN`, `ISODD`, `CEILING`, `FLOOR`, `PI`, `EXP`, `LN`, `LOG`, `LOG10`, `SIN`, `COS`, `TAN`, `SINH`, `COSH`, `TANH`, `ASINH`, `ACOSH`, `ATANH`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `RADIANS`, `DEGREES`, `PV`, `FV`, `NPV`, `PMT`, `BITAND`, `BITOR`, `BITXOR`, `BITLSHIFT`, `BITRSHIFT`, `AND`, `OR`, `XOR`, `NOT`, `LEN`, `LOWER`, `UPPER`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `SUBSTITUTE`, `REPLACE`, `CONCAT`, `TEXTJOIN`, `CHOOSE`, `MATCH`, `EXACT`, `FIND`, `SEARCH`, `CODE`, `N`, `VALUE`, `DATEVALUE`, `TIMEVALUE`, `ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISLOGICAL`, `ISERROR`, `COUNT`, `COUNTA`, `COUNTBLANK`, `DATE`, `YEAR`, `MONTH`, `DAY`, `DAYS`, `TIME`, `HOUR`, `MINUTE`, `SECOND`, `EDATE`, `EOMONTH`, `WEEKDAY`, `WEEKNUM`, and `ISOWEEKNUM`.
- Conditional/selector evaluation (`IF`, `IFERROR`, `IFS`, `SWITCH`, `CHOOSE`, `INDEX`) now preserves typed branch results (text/bool/number) instead of forcing numeric-only outputs.
- `SWITCH` case matching now avoids cross-type text-to-zero false matches by comparing text cases as text (with numeric coercion retained for non-text scalar cases).
- `IF`/`IFS` condition coercion now supports `TRUE`/`FALSE` text and numeric text while treating invalid non-numeric/non-boolean text conditions as parse errors.
- Logical aggregators (`AND`, `OR`, `XOR`, `NOT`) now share the same condition coercion semantics as `IF`/`IFS` for text/boolean/numeric-text inputs.
- Arithmetic operators now coerce numeric text (including trimmed numeric text) while treating invalid non-numeric text operands as parse errors.
- Selector index functions (`CHOOSE`, `INDEX`) now coerce numeric text indexes while rejecting invalid non-numeric text indexes with parse errors.
- Lookup index functions (`MATCH`, `XMATCH`) now coerce numeric text needles/candidates/match modes while rejecting invalid non-numeric text inputs with parse errors.
- Full-sheet recalc scheduling now reuses cached dependency-analysis ordering (`topo + cycle tail`) to avoid rebuilding/sorting an all-formula target set each run.
- Dependency analysis now derives cyclic-formula tails directly from Kahn indegree state, avoiding an extra topo-set allocation and set-difference pass per recalc.
- Dependency/topology traversal now reuses precomputed `dependents_by_ref` adjacency vectors instead of rebuilding set-backed adjacency from `dependency_refs` during ordering.
- Recalc evaluation loop now reuses a single recursion stack across target cells and only materializes per-cell duration maps when DAG timing capture is enabled.
- Incremental recalc now propagates changed-root slices directly through internal scheduling (instead of allocating/sorting an intermediate root set) while preserving impacted-cell traversal semantics.
- Incremental impacted traversal now uses `Vec + HashSet` accumulation (with deterministic topo/cycle ordering applied afterward) to reduce tree-set insertion overhead during dependency-scoped recompute.
- Dependency-graph telemetry now skips expensive payload assembly for sinks that opt out (for example `NoopEventSink`), reducing recalc-time JSON allocation churn in non-logging runs.
- AST interning scaffold now exposes deduplicated formula-node IDs for parser introspection.
- Incremental recalc from changed roots is available in core and used by `tx-save` post-mutation workflows.
- Incremental recalc now reuses cached reverse-dependency indexes during impacted-cell discovery and DAG degree analysis to reduce repeated graph traversal overhead on larger formula sets.
- Incremental formula ordering now reuses cached topological-position indexes, reducing repeated full-topo scans for dependency-scoped recompute and DAG topological targeting.
- XLSX workbook loader projects worksheet values/formulas into the in-memory model for recalc workflows.
- XLSX saver writes workbook model back to `.xlsx` with deterministic sheet/cell ordering baseline.
- `save` now defaults to `preserve` mode for compatibility-first workbook output (`normalize` remains available via `--mode normalize`).
- Preserve mode now uses passthrough copy semantics to retain unknown XML parts exactly.
- Transactional preserve save rewrites only changed worksheet parts while retaining untouched/unknown parts.
- Interop corpus assembly utility (`python/assemble_excel_interop_corpus.py`) merges deterministic generated fixtures with optional curated real Excel-authored samples from `corpus/excel-authored/manifest.json`, including legal-clearance metadata and minimum-sample policy checks.
- Interop CI gate now enforces curated-sample coverage (`EXCEL_INTEROP_MIN_EXCEL_AUTHORED_SAMPLES=5`) while still requiring the deterministic generated fixture set.
- Interop CI gate enforces baseline curated feature coverage (`formulas`, `styles`, `comments`, `charts`, `defined_names`) so curated corpus runs exercise broader compatibility paths.
- Corpus fixture generator (`python/generate_corpus_fixtures.py`) now emits deterministic rich-feature fixtures (`styles.xlsx`, `comments.xlsx`, `chart.xlsx`, `defined-names.xlsx`) in addition to baseline structure fixtures and writes a fixture manifest (`manifest.json`).
- Excel interop verification harness (`python/verify_excel_interop.py`) validates bidirectional workbook opening/round-trip flows across openpyxl and `rootcellar-cli`, including deterministic corpus sweeps, manifest capture, and required-fixture assertions (`--require-corpus-fixture`).
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
- Replace/expand the seeded curated sample set with verified Microsoft Excel-authored workbooks under `corpus/excel-authored/files` while keeping `EXCEL_INTEROP_MIN_EXCEL_AUTHORED_SAMPLES >= 1` green.
- Continue parser/evaluator and scheduler optimization work on top of the AST interning scaffold with benchmark-backed validation.
- Start desktop shell initialization and bridge UI->engine trace context propagation.
- Add a minimal UI smoke check in CI (startup + one engine command round-trip) to guard M0 readiness.
