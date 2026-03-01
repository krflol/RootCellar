# Execution Plan Board

Parent: [[RootCellar Master Plan]]
Evidence ledger: [[Execution Status]]
Last updated: March 1, 2026

## Completed Items
- [x] Sprint 00 foundation baseline delivered (workspace scaffold, telemetry envelope + JSONL sink, CLI skeleton commands).
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 00 - Foundation and Telemetry Bootstrap]], [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Sprint 01 interop baseline delivered (XLSX load/save model projection, preserve passthrough, selective sheet overrides).
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 01 - Workbook and XLSX Skeleton]], [[docs/RootCellar/01-Epics/Epic 01 - XLSX Fidelity and Workbook Model]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Repro artifact workflows delivered for headless runs (`repro record`, `repro check`, `repro diff`).
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]], [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Diff artifact output path delivered (`repro diff --output`) for CI/introspection consumption.
  Plan links: [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Sprint 02 parser/dependency scaffold increment delivered (AST precedence parser, dependency graph analysis, graph telemetry, and `recalc --dep-graph-report` artifact output).
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Incremental invalidation baseline delivered for changed-root recompute and wired into `tx-save`.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Recalc DAG timing artifacts delivered (`recalc --dag-timing-report`) with telemetry summaries.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Function coverage baseline delivered for parser/evaluator (`SUM`, `MIN`, `MAX`, `IF`) with function-call observability metrics.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] AST interning scaffold delivered with inspectable AST IDs and dedup metrics in dependency artifacts.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Rich DAG analysis baseline delivered in timing artifacts (`critical_path`, `max_fan_in`, `max_fan_out`, slow-node threshold metrics).
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Function coverage and DAG threshold configuration increment delivered (`AVERAGE`/`AVG`, `ABS`, `AND`, `OR`, `NOT`; `recalc --dag-slow-threshold-us`).
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] XLSX workbook part-graph baseline delivered in interop artifacts (`open` report graph nodes/edges + save graph flags for preserve/normalize).
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 01 - Workbook and XLSX Skeleton]], [[docs/RootCellar/01-Epics/Epic 01 - XLSX Fidelity and Workbook Model]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/XLSX Import Export Pipeline]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Corpus part-graph validation command delivered (`part-graph-corpus`) with aggregate artifact output, telemetry, and fail-on-error mode for CI gating.
  Plan links: [[docs/RootCellar/01-Epics/Epic 01 - XLSX Fidelity and Workbook Model]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] CI workflow publication delivered for corpus part-graph validation artifacts (`.github/workflows/corpus-part-graph.yml`) with scheduled and PR/main triggers.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/01-Epics/Epic 01 - XLSX Fidelity and Workbook Model]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] CI workflow publication delivered for reproducibility bundle checks/artifacts (`.github/workflows/repro-bundle.yml`) with scheduled and PR/main triggers.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] CI artifact naming/retention policy aligned across corpus + repro workflows (standardized artifact names, retention window, and manifest metadata).
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Batch directory recalc execution delivered (`batch recalc`) with bounded Rayon parallelism and nightly CI artifact publication (`.github/workflows/batch-recalc-nightly.yml`).
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]], [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Nightly batch coverage expanded to broader deterministic corpus slices with throughput regression thresholds.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]], [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/05-Quality/Performance and Benchmarking]], [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Nightly batch throughput trend snapshots and alert-hook payload artifacts delivered for dashboard/SLO integration.
  Plan links: [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/05-Quality/Performance and Benchmarking]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Nightly batch alert-hook routing integrated for incident and dashboard ingestion endpoints with dispatch status artifacts.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/Incident Response Playbook]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Alert-route dispatch hardened with authenticated delivery, retry/backoff policy, and acknowledgement tracking for route-level introspection.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/Incident Response Playbook]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Dispatch idempotency and correlation traceability delivered (per-route idempotency keys, correlation IDs, correlation-match gating, and dispatch artifact counters).
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/Incident Response Playbook]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Dispatch replay-protection policy and acknowledgement-retention indexing delivered (timestamp/nonce/window headers, per-attempt replay metadata, and retention lookup artifact publication).
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/Incident Response Playbook]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Nightly dashboard-pack and alert-policy wiring delivered from snapshot/dispatch/ack-retention artifacts, including policy-gated CI enforcement and published introspection bundle outputs.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/Incident Response Playbook]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Policy-to-owner escalation metadata and dashboard adapter exports delivered for downstream incident/dashboard ingestion systems (`ci-batch-policy-escalation.json`, `ci-batch-dashboard-adapter-exports.json`).
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/Incident Response Playbook]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Adapter export schema validation and compatibility-version contracts enforced in CI (`schemas/artifacts/v1/*`, `python/validate_batch_adapter_contracts.py`, nightly workflow validation gate).
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Full artifact-family schema validation and compatibility-version contracts enforced in CI for snapshot/dispatch/ack-retention/dashboard-pack/policy plus escalation/adapter outputs.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Schema-drift canary fixture gate and artifact schema migration playbook delivered for compatibility-regression detection and major-version rollout discipline.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Artifact Schema Migration Playbook]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Dual-read migration drill gate delivered for schema major-version producer/consumer overlap and rollback verification.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Artifact Schema Migration Playbook]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Multi-artifact dual-read migration matrices delivered for snapshot/dispatch/ack-retention/dashboard-pack/policy/escalation/adapter families, with CI artifact-subset policy knob and manifest introspection fields.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Artifact Schema Migration Playbook]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]], [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Migration-drill forensic report artifact and staged wave controls delivered (`ci-batch-schema-migration-drill.json`, `--wave-spec`, per-phase diagnostics and timings, manifest wave policy fields).
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Artifact Schema Migration Playbook]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/Incident Response Playbook]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Migration-drill fault-injection scenarios delivered for malformed fallback schemas and partial-wave rollback rehearsal, with CI policy knobs and manifest introspection fields.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Artifact Schema Migration Playbook]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/Incident Response Playbook]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Migration-drill negative-policy dry-run checks delivered for invalid staged-wave specs and unsupported fault-scenario keys, with explicit CI gate and manifest policy-state exposure.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Artifact Schema Migration Playbook]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/Incident Response Playbook]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]

## In Progress Items
- [ ] Sprint 02 parser/dependency-graph core is in progress.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]]
  Current state: parser, dependency graph, incremental invalidation, DAG timing + analysis artifacts, expanded starter function set (including lookup/text/date/time + counting/stat/math/finance increment: `LEN`, `CHOOSE`, `MATCH`, `INDEX`, `XMATCH`, `EXACT`, `FIND`, `SEARCH`, `CODE`, `N`, `VALUE`, `DATEVALUE`, `TIMEVALUE`, `ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISLOGICAL`, `ISERROR`, `COUNT`, `COUNTA`, `COUNTBLANK`, `SUM`, `SUMSQ`, `PRODUCT`, `MIN`, `MAX`, `MEDIAN`, `SMALL`, `LARGE`, `GEOMEAN`, `HARMEAN`, `VARP`, `VAR`/`VARS`, `STDEVP`, `STDEV`/`STDEVS`, `IF`, `IFERROR`, `AVERAGE`/`AVG`, `ABS`, `INT`, `FACT`, `FACTDOUBLE`, `COMBIN`, `PERMUT`, `GCD`, `LCM`, `QUOTIENT`, `MOD`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `MROUND`, `POWER`, `SQRT`, `SIGN`, `EVEN`, `ODD`, `ISEVEN`, `ISODD`, `CEILING`, `FLOOR`, `PI`, `EXP`, `LN`, `LOG`, `LOG10`, `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `RADIANS`, `DEGREES`, `PV`, `FV`, `NPV`, `PMT`, `AND`, `OR`, `NOT`, `DATE`, `YEAR`, `MONTH`, `DAY`, `DAYS`, `TIME`, `HOUR`, `MINUTE`, `SECOND`, `EDATE`, `EOMONTH`, `WEEKDAY`, `WEEKNUM`, `ISOWEEKNUM`), reverse-dependency index reuse for incremental scheduler/perf hardening, AST interning scaffold, and configurable slow-node thresholds are implemented; deeper function parity expansion and further scheduler/perf hardening are pending.
  Evidence: [[Execution Status#Current Execution Slice]], [[Execution Status#Next Execution Slice]]
- [ ] Milestone M0 remains in progress pending UI shell startup and desktop->engine trace bridge completion.
  Plan links: [[docs/RootCellar/00-Program/Milestone Roadmap#Milestone M0 Foundation Ready (March 15, 2026)]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Next Execution Slice]]

## Next Planned Items
1. Continue function parity expansion (deeper lookup/text/date families) and further optimize parser/intern scheduler hot paths.
2. Start Tauri shell initialization and bridge trace context into UI->engine command paths.
