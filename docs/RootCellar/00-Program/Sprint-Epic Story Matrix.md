# Sprint-Epic Story Matrix

Parent: [[Execution Plan Board]], [[docs/RootCellar/00-Program/PRD Decomposition Map]]
Last updated: March 4, 2026

## Purpose
This note expands the PRD into a concrete traceable backlog by story, status, and execution artifact.

## Completed Stories
| Story | Epic | PRD Coverage | Sprint | Status | Evidence |
|---|---|---|---|---|---|
| S00-CORE-01 | Epic 07 | Trace event contract baseline | Sprint 00 | [x] Done | [[docs/RootCellar/04-Observability/Telemetry Taxonomy and Event Schema]], `schemas/events/v1/envelope.schema.json`, [[docs/RootCellar/00-Program/Execution Status#Completed In Code]], [[apps/desktop/src-tauri/src/main.rs]], [[crates/rootcellar-core/src/telemetry.rs]] |
| S00-CORE-02 | Epic 07 / 05 | CLI + JSONL lifecycle instrumentation | Sprint 00 | [x] Done | [[docs/RootCellar/00-Program/Execution Status#Completed In Code]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[crates/rootcellar-cli/src/main.rs]], [[apps/desktop/src/main.ts]] |
| S01-IO-01 | Epic 01 | XLSX open/edit/save preserve-first | Sprint 01 | [x] Done | [[docs/RootCellar/03-Implementation/XLSX Import Export Pipeline]], [[docs/RootCellar/03-Implementation/Engine Workbook Model Spec]], [[apps/desktop/src-tauri/src/main.rs]], [[docs/RootCellar/00-Program/Execution Status#Completed In Code]] |
| S01-IO-02 | Epic 01 | Part-graph validation and drift checks | Sprint 01, Sprint 06 | [x] Done | [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]], [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[apps/desktop/src-tauri/src/main.rs]], [[crates/rootcellar-cli/src/main.rs]] |
| S02-CALC-01 | Epic 02 | Parser scaffold and AST foundation | Sprint 02 | [x] Done | [[docs/RootCellar/03-Implementation/Calculation Engine Design]], [[docs/RootCellar/00-Program/Execution Status#Completed In Code]], [[crates/rootcellar-core/src/calc.rs]] |
| S02-CALC-02 | Epic 02 | Dependency graph and cycle diagnostics | Sprint 02 | [x] Done | [[crates/rootcellar-core/src/calc.rs]], [[docs/RootCellar/03-Implementation/Calculation Engine Design]], [[apps/desktop/src/main.ts]] |
| S02-CALC-03 | Epic 02 | Incremental recalc + `tx-save` integration | Sprint 02 | [x] Done | [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[crates/rootcellar-core/src/calc.rs]], [[docs/RootCellar/00-Program/Execution Status#Completed In Code]] |
| S02-CALC-04 | Epic 02 | DAG timing artifacts and scheduler metrics | Sprint 02 | [x] Done | [[docs/RootCellar/04-Observability/Telemetry Taxonomy and Event Schema]], [[apps/desktop/scripts/capture-ui.mjs]], [[docs/RootCellar/03-Implementation/Calculation Engine Design]] |
| S02-CALC-05 | Epic 02 | AST interning introspection metrics | Sprint 02 | [x] Done | [[crates/rootcellar-core/src/calc.rs]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[docs/RootCellar/00-Program/Execution Status#Completed In Code]] |
| S05-CLI-01 | Epic 05 | CLI command baseline with trace-capable outputs | Sprint 02 | [x] Done | [[crates/rootcellar-cli/src/main.rs]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/00-Program/Execution Status#Completed In Code]] |
| S05-CLI-02 | Epic 05 | Repro record/check/diff workflows | Sprint 06 | [x] Done | [[crates/rootcellar-cli/src/main.rs]], [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/00-Program/Execution Status#Completed In Code]] |
| S05-CLI-03 | Epic 05 | Batch throughput + artifact publishing | Sprint 06 | [x] Done | [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]], [[.github/workflows/batch-recalc-nightly.yml]], [[apps/desktop/src/main.ts]] |
| S07-TRACE-01 | Epic 07 | Desktop command context continuity to trace metadata | Sprint 00, Sprint 02, Sprint 06 | [x] Done | [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[docs/RootCellar/04-Observability/Traceability Spine]], [[apps/desktop/src/main.ts]], [[apps/desktop/src-tauri/src/main.rs]], [[crates/rootcellar-core/src/telemetry.rs]] |

## In Progress Stories
| Story | Epic | PRD Coverage | Sprint | Status | Next Link |
|---|---|---|---|---|---|
| S02-CALC-06 | Epic 02 | Formula registry expansion + parity backlog | Sprint 02 / 03 | [~] In progress | [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/03-Implementation/Calculation Engine Design]], [[docs/RootCellar/04-Observability/Telemetry Taxonomy and Event Schema]] |
| S02-CALC-07 | Epic 02 | Scheduler/perf hardening at scale | Sprint 02 / 08 | [~] In progress | [[docs/RootCellar/05-Quality/Performance and Benchmarking]], [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]], [[docs/RootCellar/RootCellar Master Plan]] |
| S03-UI-01 | Epic 03 | Editing loop and command lifecycle | Sprint 03 | [~] In progress | [[docs/RootCellar/03-Implementation/UI Grid and Interaction Design]], [[docs/RootCellar/02-Sprints/Sprint 03 - Grid Editing Loop]], [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]] |
| S07-TRACE-02 | Epic 07 | Full command-to-artifact lineage and queryability | Sprint 02 / 03 | [~] In progress | [[docs/RootCellar/04-Observability/Traceability Spine]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[docs/RootCellar/00-Program/Execution Status]] |
| S03-UI-02 | Epic 03 | Accessibility baseline for core input/edit flows | Sprint 03 | [~] In progress | [[docs/RootCellar/03-Implementation/UI Grid and Interaction Design]], [[docs/RootCellar/02-Sprints/Sprint 03 - Grid Editing Loop]] |

## Planned Stories
| Story | Epic | PRD Coverage | Sprint | Target | Execution Link |
|---|---|---|---|---|---|
| S03-UI-03 | Epic 03 | Sort/filter/freeze panes/find-replace baseline | Sprint 03 / 04 | Planned | [[docs/RootCellar/03-Implementation/UI Grid and Interaction Design]], [[docs/RootCellar/02-Sprints/Sprint 03 - Grid Editing Loop]] |
| S04-SCRIPT-01 | Epic 04 | Python worker lifecycle + sandboxing | Sprint 04 / 07 | Planned | [[docs/RootCellar/03-Implementation/Python Scripting Host Design]], [[docs/RootCellar/03-Implementation/Security and Permission Model]], [[docs/RootCellar/02-Sprints/Sprint 04 - Python Macros Alpha]] |
| S04-SCRIPT-02 | Epic 04 | Macro/event invocation and audit logs | Sprint 04 / 07 | Planned | [[docs/RootCellar/03-Implementation/Python Scripting Host Design]], [[docs/RootCellar/01-Epics/Epic 04 - Python Automation Platform]] |
| S06-COMP-01 | Epic 06 | Compatibility issue taxonomy in UI/CLI | Sprint 05 | Planned | [[docs/RootCellar/03-Implementation/Compatibility Panel Design]], [[docs/RootCellar/03-Implementation/XLSX Import Export Pipeline]], [[docs/RootCellar/02-Sprints/Sprint 05 - Compatibility Panel and Corpus Tests]] |
| S06-COMP-02 | Epic 06 | VBA migration assistant and remediation guidance | Sprint 05 / 08 | Planned | [[docs/RootCellar/03-Implementation/Compatibility Panel Design]], [[docs/RootCellar/01-Epics/Epic 06 - Compatibility and Migration Tooling]] |
| S08-TRUST-01 | Epic 08 | Signed-only mode + verification policy | Sprint 07 / 08 | Planned | [[docs/RootCellar/03-Implementation/Security and Permission Model]], [[docs/RootCellar/01-Epics/Epic 08 - Enterprise Trust and Distribution]], [[docs/RootCellar/02-Sprints/Sprint 07 - Security Hardening and Add-in Signing]] |

## Story-to-Execution Readings
- Current execution slice: `[[docs/RootCellar/00-Program/Execution Status#Current Execution Slice]]`, `[[docs/RootCellar/00-Program/Execution Status#Next Execution Slice]]`
- Traceability plane: `[[docs/RootCellar/04-Observability/Traceability Spine]]`
- Risk and dependency context: `[[docs/RootCellar/00-Program/Dependency Map]]`, `[[docs/RootCellar/00-Program/Risk Register]]`
- Execution atlas and runbook: `[[docs/RootCellar/04-Observability/Execution Traceability Atlas]]`, `[[docs/RootCellar/04-Observability/AI Introspection Runbook]]`
