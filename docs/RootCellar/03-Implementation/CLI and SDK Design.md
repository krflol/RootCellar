# CLI and SDK Design

Parent: [[Architecture Overview]]
Related epic: [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]]

## CLI Goals
- Execute workbook transformations non-interactively.
- Run macros/add-ins with policy enforcement.
- Recalculate and export outputs at scale.
- Emit machine-readable reports and reproducibility artifacts.

## Command Surface v1
- `rootcellar open <file> --report <path>`
- `rootcellar recalc <file> [--mode preserve|normalize]`
- `rootcellar run-macro <file> --macro <name>`
- `rootcellar batch recalc <dir> [--threads N] [--detail-level minimal|diagnostic|forensic]`
- `rootcellar repro record|check <bundle>`

## Current Implemented Baseline (February 28, 2026)
- Implemented: `open`, `save`, `recalc`, `tx-demo`, `tx-save`, `batch recalc`, `repro record`, `repro check`, `repro diff`.
- Implemented corpus validator: `part-graph-corpus <dir> [--max-files N] [--fail-on-errors]`.
- Implemented diff artifact output: `repro diff --format text|json --output <path>`.
- Implemented dependency graph artifact output: `recalc --dep-graph-report <path>`.
- Implemented DAG timing artifact output: `recalc --dag-timing-report <path>`.
- Implemented DAG threshold override: `recalc --dag-slow-threshold-us <microseconds>` (requires `--dag-timing-report`).
- Implemented incremental post-mutation recalc in `tx-save` using changed-root invalidation.
- Implemented parser introspection metrics in dependency artifacts (`function_call_count`, `ast_node_count`, `ast_unique_node_count`, `formula_ast_ids`).
- Implemented DAG analysis metrics in timing artifacts (`critical_path`, `max_fan_in`, `max_fan_out`, slow-node threshold).
- Implemented function evaluator starter set: `SUM`, `SUMSQ`, `PRODUCT`, `MIN`, `MAX`, `MEDIAN`, `SMALL`, `LARGE`, `GEOMEAN`, `HARMEAN`, `VARP`, `VAR`/`VARS`, `STDEVP`, `STDEV`/`STDEVS`, `IF`, `IFERROR`, `AVERAGE`/`AVG`, `ABS`, `INT`, `FACT`, `FACTDOUBLE`, `COMBIN`, `PERMUT`, `GCD`, `LCM`, `QUOTIENT`, `MOD`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `MROUND`, `POWER`, `SQRT`, `SIGN`, `EVEN`, `ODD`, `ISEVEN`, `ISODD`, `CEILING`, `FLOOR`, `PI`, `EXP`, `LN`, `LOG`, `LOG10`, `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `RADIANS`, `DEGREES`, `PV`, `FV`, `NPV`, `PMT`, `AND`, `OR`, `NOT`, `LEN`, `CHOOSE`, `MATCH`, `EXACT`, `FIND`, `SEARCH`, `CODE`, `N`, `VALUE`, `DATEVALUE`, `TIMEVALUE`, `ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISLOGICAL`, `ISERROR`, `COUNT`, `COUNTA`, `COUNTBLANK`, `DATE`, `YEAR`, `MONTH`, `DAY`, `DAYS`, `TIME`, `HOUR`, `MINUTE`, `SECOND`, `EDATE`, `EOMONTH`, `WEEKDAY`, `WEEKNUM`, `ISOWEEKNUM`.
- Implemented incremental scheduler/perf optimization via reverse-dependency index reuse for impacted-formula selection and DAG degree/adjacency derivation.
- Implemented workbook part-graph artifacts in `open` reports and graph-aware preserve/normalize save flags in `save`/`tx-save` outputs.
- Implemented CI workflow publication for corpus validator artifacts (`.github/workflows/corpus-part-graph.yml`).
- Implemented CI workflow publication for repro bundle artifacts (`.github/workflows/repro-bundle.yml`).
- Implemented nightly CI workflow publication for batch recalc artifacts (`.github/workflows/batch-recalc-nightly.yml`).
- Implemented nightly batch corpus assembly utility (`python/build_batch_nightly_corpus.py`) for broader deterministic compatibility slices.
- Implemented batch throughput summary metrics in report artifacts (`throughput_files_per_sec`, `aggregate_file_time_ratio`) for regression gating.
- Implemented nightly batch trend snapshot + alert-hook builder (`python/build_batch_trend_snapshot.py`) for dashboard/SLO ingestion (`ci-batch-throughput-snapshot.json`, `ci-batch-alert-hook.json`).
- Implemented nightly alert-route dispatch utility (`python/dispatch_batch_alert_hook.py`) for incident/dashboard endpoint delivery and route-status artifacts (`ci-batch-alert-dispatch.json`).
- Implemented dispatch hardening features: token auth, optional HMAC signing, retry/backoff policy, and acknowledgement tracking in route artifacts.
- Implemented deterministic route idempotency/correlation propagation and correlation-match enforcement options for cross-system incident traceability.
- Implemented dispatch replay-protection policy controls (timestamp/nonce/window headers + per-attempt replay metadata) and nightly ack-retention index generation (`python/build_batch_ack_retention_index.py`).
- Implemented nightly dashboard-pack/policy artifact builder (`python/build_batch_dashboard_pack.py`) and policy-gated CI wiring over snapshot + dispatch + ack-retention artifacts.
- Implemented policy-owner escalation and adapter export builder (`python/build_batch_policy_adapters.py`) for downstream incident/dashboard ingestion payloads.
- Implemented full artifact-family schema + compatibility contract validation (`schemas/artifacts/v1/*`, `python/validate_batch_adapter_contracts.py --full-family`) and nightly CI enforcement.
- Implemented schema-drift canary harness (`python/validate_batch_schema_canaries.py`) and nightly canary gate for contract-regression assertions.
- Implemented multi-artifact dual-read migration drill harness (`python/validate_batch_dual_read_migration.py`) and nightly migration gate for producer/consumer overlap and rollback assertions, including artifact-subset targeting via `--artifacts`, staged-wave scenarios via `--wave-spec`, fault-injection scenarios via `--fault-injection --fault-scenarios`, and per-phase diagnostics artifacts via `--report`.
- Implemented migration policy dry-run harness (`python/validate_batch_migration_policy_dry_run.py`) and nightly policy gate to assert invalid staged-wave specs and unsupported fault-scenario keys fail fast.
- Implemented aligned CI artifact policy (name pattern + retention + manifest metadata) across corpus and repro workflows.
- Tracking links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]

## JSONL Reporting Contract
Each run emits JSONL lines for:
- lifecycle events
- compatibility findings
- performance timings
- mutation summaries
- errors with stack and correlation IDs

## SDK Model
- Rust crate exposes engine/session APIs.
- Python bindings expose safe automation APIs mirroring script model.
- Versioned API contracts; semantic versioning for breaking changes.

## Batch Parallelism
- Rayon-based directory job execution with bounded concurrency.
- Deterministic mode enforces stable work ordering and output naming.

## Observability
- CLI includes trace headers and artifact IDs in stdout summary and JSONL.
- Headless runs export artifact bundles for offline introspection.

## Exit Criteria
- CI pipeline can run nightly corpus in headless mode with reproducibility and bounded-parallel batch checks.
