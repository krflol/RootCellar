# Execution Plan Board

Parent: [[RootCellar Master Plan]]
Evidence ledger: [[Execution Status]]
Last updated: March 4, 2026

## Execution Traceability Atlas
- Traceability navigator for this execution board: [[docs/RootCellar/04-Observability/Execution Traceability Atlas]]
- Active traceability runbook for human/AI inspection: [[docs/RootCellar/04-Observability/AI Introspection Runbook]]
- Story-to-plan crosswalk: [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]

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
- [x] Nightly synthetic recalc benchmark gate wired into batch CI with policy knobs, threshold enforcement, and benchmark artifact publication.
  Plan links: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/05-Quality/Performance and Benchmarking]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Desktop UI observability slice for capture QA and trace visibility delivered.
  Plan links: [[docs/RootCellar/00-Program/Execution Status]], [[docs/RootCellar/06-Operations/CI-CD Blueprint]]
  Evidence: [[Execution Status#Execution Plan Linkage]], [[apps/desktop/src/main.ts]], [[apps/desktop/scripts/capture-ui.mjs]]
- [x] PRD decomposition and traceability planning suite delivered for implementation handoff.
  Plan links: [[docs/RootCellar/RootCellar Master Plan]], [[docs/RootCellar PRD]], [[docs/RootCellar/00-Program/PRD Decomposition Map]], [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[docs/RootCellar/04-Observability/Traceability Spine]], [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]
  Evidence: [[Execution Status]], [[Execution Status#Completed In Code]], [[docs/RootCellar/04-Observability/AI Introspection Workflow]], [[docs/RootCellar/04-Observability/Trace Correlation Model]], [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]
- [x] Desktop command context continuity now delivered from action entry to trace metadata output.
  Plan links: [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[docs/RootCellar/04-Observability/Traceability Spine]], [[apps/desktop/src/main.ts]], [[apps/desktop/src-tauri/src/main.rs]], [[crates/rootcellar-core/src/telemetry.rs]], [[Execution Status#Completed In Code]]
  Evidence: [[Execution Status#Current Execution Slice]], [[Execution Status#Completed In Code]]
- [x] Desktop trace output now includes deterministic artifact references for command-level introspection (status, duration, and linked artifact IDs).
  Plan links: [[docs/RootCellar/04-Observability/Traceability Spine]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[docs/RootCellar/04-Observability/Trace Correlation Model]], [[docs/RootCellar/00-Program/PRD Decomposition Map]], [[apps/desktop/src-tauri/src/main.rs]], [[apps/desktop/src/main.ts]]
  Evidence: [[Execution Status#Current Execution Slice]], [[Execution Status#Completed In Code]]
- [x] Desktop command-chain continuity for open/edit/preview/save/recalc/round-trip is now smoke-asserted in backend tests.
  Plan links: [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[docs/RootCellar/04-Observability/Execution Traceability Atlas]], [[docs/RootCellar/04-Observability/AI Introspection Runbook]], [[apps/desktop/src-tauri/src/main.rs]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Desktop artifact-index output is now emitted per command and wired into trace metadata with trace id/command metadata for local debugging and forensic replay.
  Plan links: [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[docs/RootCellar/04-Observability/Traceability Spine]], [[apps/desktop/src-tauri/src/main.rs]], [[apps/desktop/src/main.ts]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Desktop command-output trace schema now has a machine-checkable contract.
  Plan links: [[docs/RootCellar/04-Observability/Traceability Spine]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[apps/desktop/src/desktopTraceOutput.ts]]
  Evidence: [[Execution Status#Completed In Code]], [[schemas/desktop/v1/command-output-trace.schema.json]], [[apps/desktop/src/desktopTraceOutput.test.ts]]
- [x] Desktop linked artifact IDs now resolve to deterministic artifact-index records for inspection joins.
  Plan links: [[docs/RootCellar/04-Observability/Traceability Spine]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[apps/desktop/scripts/resolve-desktop-trace-artifacts.ts]], [[apps/desktop/src/desktopTraceJoin.ts]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Desktop continuity smoke now validates command output artifacts against manifest index records.
  Plan links: [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[docs/RootCellar/04-Observability/Traceability Spine]], [[apps/desktop/src-tauri/src/main.rs]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]

- [x] Sprint 04 Python macro foundation is now wired in desktop runtime (`interop_run_macro`).
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 04 - Python Macros Alpha]], [[docs/RootCellar/01-Epics/Epic 04 - Python Automation Platform]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[apps/desktop/src-tauri/src/main.rs]], [[apps/desktop/src-tauri/src/script.rs]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Desktop macro policy persistence and consent UX is now implemented for per-script permission control.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 04 - Python Macros Alpha]], [[docs/RootCellar/01-Epics/Epic 04 - Python Automation Platform]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[apps/desktop/src/main.ts]]
  Evidence: [[Execution Status#Current Execution Slice]], [[Execution Status#Completed In Code]], [[Execution Status#Verification]]

- [x] Sprint 03 grid loop foundations are shipped to a working UI shape.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 03 - Grid Editing Loop]], [[docs/RootCellar/01-Epics/Epic 03 - Desktop Grid UX and Productivity]], [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[apps/desktop/src/main.ts]]
  Evidence: [[Execution Status#Completed In Code]], [[apps/desktop/src/previewInteractions.test.ts]], [[apps/desktop/src/editRangePresets.test.ts]], [[apps/desktop/src/presetReuse.ts]], [[apps/desktop/src/presetReuse.test.ts]], [[apps/desktop/src/presetReuseView.ts]], [[apps/desktop/src/recalcFreshness.ts]], [[apps/desktop/src/recalcFreshness.test.ts]]

- [x] Sprint 03 interaction acceptance now includes clipboard paste + undo/redo history baseline.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 03 - Grid Editing Loop]], [[docs/RootCellar/01-Epics/Epic 03 - Desktop Grid UX and Productivity]], [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[apps/desktop/src-tauri/src/main.rs]], [[apps/desktop/src/main.ts]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]
- [x] Sprint 03 accessibility and edit-lifecycle telemetry hardening is complete.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 03 - Grid Editing Loop]], [[docs/RootCellar/01-Epics/Epic 03 - Desktop Grid UX and Productivity]], [[docs/RootCellar/03-Implementation/UI Grid and Interaction Design]], [[apps/desktop/src/main.ts]], [[apps/desktop/src/main.accessibility.test.ts]]
  Evidence: [[Execution Status#Completed In Code]], [[Execution Status#Verification]]

## In Progress Items
- [ ] Sprint 02 parser/dependency-graph core is in progress.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]]
  Current state: parser, dependency graph, incremental invalidation, DAG timing + analysis artifacts, expanded starter function set (including lookup/text/date/time + counting/stat/math/finance increment: `LEN`, `LOWER`, `UPPER`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `SUBSTITUTE`, `REPLACE`, `CONCAT`, `TEXTJOIN`, `CHOOSE`, `MATCH`, `INDEX`, `XMATCH`, `EXACT`, `FIND`, `SEARCH`, `CODE`, `N`, `VALUE`, `DATEVALUE`, `TIMEVALUE`, `ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISLOGICAL`, `ISERROR`, `COUNT`, `COUNTA`, `COUNTBLANK`, `SUM`, `SUMSQ`, `PRODUCT`, `MIN`, `MAX`, `MEDIAN`, `SMALL`, `LARGE`, `GEOMEAN`, `HARMEAN`, `VARP`, `VAR`/`VARS`, `STDEVP`, `STDEV`/`STDEVS`, `IF`, `IFERROR`, `IFS`, `SWITCH`, `AVERAGE`/`AVG`, `ABS`, `INT`, `FACT`, `FACTDOUBLE`, `COMBIN`, `PERMUT`, `GCD`, `LCM`, `QUOTIENT`, `MOD`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `MROUND`, `POWER`, `SQRT`, `SIGN`, `EVEN`, `ODD`, `ISEVEN`, `ISODD`, `CEILING`, `FLOOR`, `PI`, `EXP`, `LN`, `LOG`, `LOG10`, `SIN`, `COS`, `TAN`, `SINH`, `COSH`, `TANH`, `ASINH`, `ACOSH`, `ATANH`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `RADIANS`, `DEGREES`, `PV`, `FV`, `NPV`, `PMT`, `BITAND`, `BITOR`, `BITXOR`, `BITLSHIFT`, `BITRSHIFT`, `AND`, `OR`, `XOR`, `NOT`, `DATE`, `YEAR`, `MONTH`, `DAY`, `DAYS`, `TIME`, `HOUR`, `MINUTE`, `SECOND`, `EDATE`, `EOMONTH`, `WEEKDAY`, `WEEKNUM`, `ISOWEEKNUM`), reverse-dependency index reuse and topological-position index reuse for incremental scheduler/perf hardening, AST interning scaffold, configurable slow-node thresholds, text-returning formula evaluation, typed branch-preserving conditional/selector evaluation (`IF`, `IFERROR`, `IFS`, `SWITCH`, `CHOOSE`, `INDEX`), and formula-language literals (quoted text + `TRUE`/`FALSE`) are implemented; deeper function parity expansion and further scheduler/perf hardening are pending.
  Evidence: [[Execution Status#Current Execution Slice]], [[Execution Status#Next Execution Slice]]
- [ ] Sprint 04 Python macro platform remains in-progress for signed package trust policy and provenance attestation.
  Plan links: [[docs/RootCellar/02-Sprints/Sprint 04 - Python Macros Alpha]], [[docs/RootCellar/01-Epics/Epic 04 - Python Automation Platform]], [[docs/RootCellar/03-Implementation/Python Scripting Host Design]], [[apps/desktop/src-tauri/src/main.rs]], [[apps/desktop/src-tauri/src/script.rs]]
  Current state: core macro execution, policy UX, event surfaces, and trust provenance plumbing are implemented; remaining work is signed macro package policy and verifier metadata lineage.
  Evidence: [[Execution Status#Current Execution Slice]], [[Execution Status#Next Execution Slice]]
  Next testable deliverable: signed package verifier mode, certificate pinning controls, and offline trust manifest export.

- [ ] Cross-surface event lookup assertions for desktop trace continuity are now validated in tests and CI command invocation.
  Plan links: [[docs/RootCellar/04-Observability/Execution Traceability Atlas]], [[docs/RootCellar/04-Observability/AI Introspection Runbook]], [[apps/desktop/scripts/capture-ui.mjs]], [[apps/desktop/src-tauri/src/main.rs]]
  Current state: command-level continuity and response-level metadata assertions are now covered in `desktop_trace_continuity_smoke_open_edit_save_recalc` across open/edit/preview/save/recalc/round-trip, plus event-log inclusion checks for that flow.

## Next Planned Items
 1. Stabilize next-phase calc edge-case compatibility in Sprint 02 while extending function parity in compatibility-stressed fixtures.
2. Advance `Sprint 04 - Python Macros Alpha` into signed package policy enforcement, verifier metadata artifacts, and trust distribution model hardening.
3. Expand curated Excel-authored corpus with scenario-manager and layout-rich variants, while preserving policy gates (`EXCEL_INTEROP_MIN_EXCEL_AUTHORED_SAMPLES=17` and `EXCEL_INTEROP_MIN_VERIFIED_EXCEL_SAMPLES=17`).


