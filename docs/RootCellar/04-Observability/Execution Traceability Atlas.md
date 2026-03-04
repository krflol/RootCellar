# Execution Traceability Atlas

Parent: [[docs/RootCellar/00-Program/Execution Status]]
Execution control: [[Execution Plan Board]], [[docs/RootCellar/00-Program/PRD Decomposition Map]]

## Purpose
Keep a single entry point for implementation tracing from PRD intent, through epics and sprints, into completed and active artifacts.

## Program Slice Index
1. [[docs/RootCellar/00-Program/Execution Status#Current Execution Slice]]
2. [[docs/RootCellar/00-Program/Execution Status#Next Execution Slice]]
3. [[docs/RootCellar/00-Program/Execution Plan Board#Completed Items]]
4. [[docs/RootCellar/00-Program/Execution Plan Board#In Progress Items]]
5. [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]
6. [[docs/RootCellar/04-Observability/Traceability Spine]]
7. [[docs/RootCellar/04-Observability/AI Introspection Runbook]]

## Completed Delivery Links
- UI-to-engine trace context propagation (`ui_command_id`/`ui_command_name`) and baseline observability slices: [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[docs/RootCellar/04-Observability/Traceability Spine]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[Execution Plan Board#Completed Items]], [[Execution Status#Completed In Code]]
- Desktop open/edit/preview/save/recalc/round-trip continuity assertions in command metadata: [[apps/desktop/src-tauri/src/main.rs]], [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[docs/RootCellar/04-Observability/Traceability Spine]], [[docs/RootCellar/04-Observability/AI Introspection Runbook]]
- Desktop event-stream assertions for open/edit/save/recalc continuity: [[apps/desktop/src-tauri/src/main.rs]], [[docs/RootCellar/00-Program/Execution Plan Board#In Progress Items]]
- Recalc/dependency/perf hardening and DAG timing artifacts: [[docs/RootCellar/00-Program/RootCellar Master Plan]], [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]], [[docs/RootCellar/04-Observability/Trace Correlation Model]]
- Repro and batch artifact governance: [[docs/RootCellar/06-Operations/CI-CD Blueprint]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/00-Program/Execution Status#Completed In Code]]
- Schema validation and migration policy instrumentation: [[docs/RootCellar/04-Observability/Artifact Schema Migration Playbook]], [[docs/RootCellar/00-Program/Execution Status#Completed In Code]]

## In Progress Delivery Links
- Parser expansion and scheduler/perf continuation: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]], [[docs/RootCellar/02-Sprints/Sprint 08 - Scalability and Hardening]], [[docs/RootCellar/03-Implementation/Calculation Engine Design]], [[docs/RootCellar/05-Quality/Performance and Benchmarking]]
- Desktop-wide trace continuity closure (command trace context + artifact index continuity assertions in command-path): [[docs/RootCellar/00-Program/Milestone Roadmap#Milestone M0 Foundation Ready (March 15, 2026)]], [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[docs/RootCellar/04-Observability/AI Introspection Runbook]], [[.github/workflows/desktop-ui-smoke.yml]], [[apps/desktop/src-tauri/src/main.rs]], [[apps/desktop/src/main.ts]]
- Command-line to desktop observability harmonization: [[docs/RootCellar/00-Program/Execution Status]], [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]], [[docs/RootCellar/04-Observability/Traceability Spine]]

## Story-Level Readings
- completed: [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix#Completed Stories]]
- in progress: [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix#In Progress Stories]]
- planned: [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix#Planned Stories]]

## Workflow Trace Cards
- If you need to inspect UI behavior, start with:
  - [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]]
  - [[docs/RootCellar/04-Observability/AI Introspection Runbook]]
- If you need core engine behavior, start with:
  - [[docs/RootCellar/03-Implementation/Calculation Engine Design]]
  - [[docs/RootCellar/04-Observability/Traceability Spine]]
  - [[docs/RootCellar/04-Observability/Trace Correlation Model]]
- If you need operational evidence, start with:
  - [[docs/RootCellar/06-Operations/CI-CD Blueprint]]
  - [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]]
  - [[docs/RootCellar/04-Observability/Audit and Forensics]]
