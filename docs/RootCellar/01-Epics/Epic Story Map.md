# Epic Story Map

Parent: [[docs/RootCellar/00-Program/Execution Plan Board]]
Tracking: [[docs/RootCellar/00-Program/Execution Status]]
Story matrix: [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]

## Epic 01 - XLSX Fidelity and Workbook Model
- **E01-S1: Preserve-First Load/Save Path**
  - Intent: load workbook, maintain unknown parts, preserve default on save.
  - Done by: `interoperability pipeline` in core + desktop open/save commands.
  - Acceptance: no data-loss on unsupported parts in preserve mode.
  - Artifacts/traceability: `interop.xlsx.load/end`, `interop.xlsx.save/*`, `part-graph` and compatibility issue summaries.
  - Evidence links: [[docs/RootCellar/03-Implementation/XLSX Import Export Pipeline]], [[docs/RootCellar/03-Implementation/Engine Workbook Model Spec]], [[docs/RootCellar/00-Program/Execution Plan Board]]
- **E01-S2: Part-Graph Validation and Drift Detection**
  - Intent: verify relationship integrity before and after transformations.
  - Done by: `part-graph-corpus` workflow and dangling-target reporting.
  - Acceptance: generated report includes edge/node counts and dangling edge inventory.
  - Evidence links: [[apps/desktop/src-tauri/src/main.rs]], [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]
- **E01-S3: Compatibility Report Semantics**
  - Intent: surface parse/save risk classifications and unknown inventory per workbook.
  - Done by: `CompatibilityStatus` render path + `open` response contract.
  - Acceptance: every unsupported/unknown item maps to visible issue class.
  - Evidence links: [[docs/RootCellar/03-Implementation/Compatibility Panel Design]], [[docs/RootCellar/03-Implementation/XLSX Import Export Pipeline]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]]

## Epic 02 - Calculation Engine and Determinism
- **E02-S1: Parser + AST + Coercion Baseline**
  - Intent: parse Excel-like formulas and evaluate deterministic baselines.
  - Done by: parser/evaluator scaffold in `rootcellar-core`.
  - Acceptance: boolean/text/numeric coercion behavior is explicit and test-backed.
  - Evidence links: [[docs/RootCellar/03-Implementation/Calculation Engine Design]], [[docs/RootCellar/05-Quality/Test Strategy]]
- **E02-S2: Dependency Graph + Incremental Recalc**
  - Intent: rerun only impacted nodes after edits.
  - Done by: dependency graph with changed-root traversal.
  - Acceptance: impacted-set regression vs full recalc tracked in benchmarks.
  - Evidence links: [[docs/RootCellar/03-Implementation/Calculation Engine Design]], [[docs/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]
- **E02-S3: Reproducible Scheduler Performance**
  - Intent: keep deterministic behavior while scaling to larger DAGs.
  - Done by: cached topo-position and reverse-dependency reuse.
  - Acceptance: nightly benchmark gates remain green with speedup thresholds.
  - Evidence links: [[docs/RootCellar/05-Quality/Performance and Benchmarking]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]]

## Epic 03 - Desktop Grid UX and Productivity
- **E03-S1: Interaction Baseline Shell**
  - Intent: make open/edit/preview path observable and stable.
  - Done by: current `main.ts` action flows and capture sections.
  - Acceptance: open/edit/save/recalc commands execute end-to-end in shell.
  - Evidence links: [[docs/RootCellar/03-Implementation/UI Grid and Interaction Design]], [[docs/RootCellar/02-Sprints/Sprint 03 - Grid Editing Loop]], [[docs/RootCellar/00-Program/Execution Status]]
- **E03-S2: Edit Lifecycle + Recalc Refresh**
  - Intent: apply edits through transactions and show dirty state.
  - Done by: `applyCellEdit`, `recalcLoadedWorkbook`, dirty-sheet state.
  - Acceptance: edit events create mutation artifacts and recalc freshness updates.
  - Evidence links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/04-Observability/UI-to-Engine Trace Bridge]]
- **E03-S3: Accessibility & Keyboard Pathing**
  - Intent: baseline parity for navigation/editing through keyboard and focus semantics.
  - Done by: interaction work planned for later sprint slice.
  - Acceptance: planned interaction matrix covers arrows/home/end/page/edit flow.
  - Evidence links: [[docs/RootCellar/03-Implementation/UI Grid and Interaction Design]]

## Epic 04 - Python Automation Platform
- **E04-S1: Worker Lifecycle & Permission Boundary**
  - Intent: script execution runs through isolated worker process.
  - Done by: CLI/engine scaffolding and host design.
  - Acceptance: deny-by-default for privileged permissions with audit log path.
  - Evidence links: [[docs/RootCellar/03-Implementation/Python Scripting Host Design]], [[docs/RootCellar/03-Implementation/Security and Permission Model]]
- **E04-S2: Macro/Event Entry Surfaces**
  - Intent: support imperative and event-driven calls.
  - Done by: command/RPC contract design + manifest model.
  - Acceptance: macro invocation yields reproducible trace and permission record.
  - Evidence links: [[docs/RootCellar/03-Implementation/Python Scripting Host Design]], [[docs/RootCellar/04-Observability/Trace Correlation Model]]
- **E04-S3: UDF Bridge + Object Model v1**
  - Intent: deterministic UDF API and cache semantics.
  - Acceptance: stable cache key and side-effect policy enforced.
  - Evidence links: [[docs/RootCellar/03-Implementation/Python Scripting Host Design]]

## Epic 05 - Headless CLI and SDK
- **E05-S1: Command Contract and JSONL Baseline**
  - Intent: stable CLI command suite with trace-capable outputs.
  - Done by: CLI command matrix and artifact logs.
  - Acceptance: open/save/recalc/batch commands emit consistent artifact envelopes.
  - Evidence links: [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/04-Observability/Telemetry Taxonomy and Event Schema]]
- **E05-S2: Batch Throughput and Reproducibility**
  - Intent: reliable corpus runs with deterministic ordering and policy gates.
  - Done by: `batch recalc`, workflows, benchmark suite.
  - Acceptance: nightly throughput and repro checks pass gates.
  - Evidence links: [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]]
- **E05-S3: CI Schema/Contract Validation**
  - Intent: artifact contracts rejected on drift.
  - Done by: schema validators in nightly workflows.
  - Acceptance: migration, canary, and dual-read gates remain enforced.
  - Evidence links: [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/05-Quality/Release Gates]]

## Epic 06 - Compatibility and Migration Tooling
- **E06-S1: Compatibility Taxonomy and Panel Schema**
  - Intent: keep parser/calc findings visible and actionable.
  - Acceptance: every parser or runtime finding maps to supported/partial/preserved/not-supported.
  - Evidence links: [[docs/RootCellar/03-Implementation/Compatibility Panel Design]]
- **E06-S2: Migration Scan + Reports**
  - Intent: scan workbook complexity and produce migration recommendations.
  - Done by: planned in replacement-beta stream.
  - Acceptance: migration output includes remediations and VBA complexity class.
  - Evidence links: [[docs/RootCellar/01-Epics/Epic 06 - Compatibility and Migration Tooling]], [[docs/RootCellar/01-Epics/Epic 03 - Desktop Grid UX and Productivity]]

## Epic 07 - Radical Observability and Introspection
- **E07-S1: End-to-End Trace Correlation**
  - Intent: link action -> command -> engine event -> artifact path.
  - Done by: `ui_command_id` propagation in desktop/backend path.
  - Acceptance: one desktop command has the same UI command id in event context.
  - Evidence links: [[docs/RootCellar/04-Observability/Trace Correlation Model]], [[docs/RootCellar/04-Observability/UI-to-Engine Trace Bridge]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]]
- **E07-S2: Artifact Catalog and Query**
  - Intent: expose queries for trace-scoped evidence retrieval.
  - Acceptance: users/AI can reconstruct a user action from output + trace.
  - Evidence links: [[docs/RootCellar/04-Observability/AI Introspection Workflow]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]]
- **E07-S3: Alert and Dashboard Completeness**
  - Intent: maintain completeness and incident readiness from traces.
  - Acceptance: SLO policy has explicit trace coverage checks.
  - Evidence links: [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/06-Operations/Incident Response Playbook]]

## Epic 08 - Enterprise Trust and Distribution
- **E08-S1: Signature and Trust Modes**
  - Intent: enforce policy-driven script/add-in trust.
  - Acceptance: unsigned artifacts blocked in signed-only mode.
  - Evidence links: [[docs/RootCellar/03-Implementation/Security and Permission Model]], [[docs/RootCellar/01-Epics/Epic 04 - Python Automation Platform]]
- **E08-S2: Policy Decision Forensics**
  - Intent: record signed policy decisions and escalation targets.
  - Acceptance: every policy decision has trace-linked reason payload.
  - Evidence links: [[docs/RootCellar/04-Observability/Trace Correlation Model]], [[docs/RootCellar/06-Operations/Incident Response Playbook]]
