# RootCellar Planning Hub

## Purpose
This vault expands the PRD into an implementation-ready program plan.
Source contract: [[docs/RootCellar PRD]].

## How To Navigate
- Program control: [[docs/RootCellar/00-Program/RootCellar Master Plan]]
- Execution board (plan <-> delivery): [[docs/RootCellar/00-Program/Execution Plan Board]]
- Execution tracking: [[docs/RootCellar/00-Program/Execution Status]]
- Milestones and sequencing: [[docs/RootCellar/00-Program/Milestone Roadmap]]
- PRD decomposition and planning: [[docs/RootCellar/00-Program/PRD Decomposition Map]]
- Dependency tracking: [[docs/RootCellar/00-Program/Dependency Map]]
- Risks and mitigations: [[docs/RootCellar/00-Program/Risk Register]]
- Decision log: [[docs/RootCellar/00-Program/Decision Register]]

## Current Delivery Snapshot (March 4, 2026)
- Completed: Sprint 00 foundation and Sprint 01 interop baseline are delivered and verified.
- In progress: PRD decomposition and traceability planning are now documented in linked notes, with desktop traceability bridge implementation next.
- Active traceability index: [[docs/RootCellar/04-Observability/Traceability Spine]], [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]
- Live status detail: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]

## Delivery Workstreams
- Epic set: [[docs/RootCellar/01-Epics/Epic 01 - XLSX Fidelity and Workbook Model]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/01-Epics/Epic 03 - Desktop Grid UX and Productivity]], [[docs/RootCellar/01-Epics/Epic 04 - Python Automation Platform]], [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]], [[docs/RootCellar/01-Epics/Epic 06 - Compatibility and Migration Tooling]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/01-Epics/Epic 08 - Enterprise Trust and Distribution]]
- Sprint execution: [[docs/RootCellar/02-Sprints/Sprint Cadence and Capacity]]
- Implementation specs: [[docs/RootCellar/03-Implementation/Architecture Overview]]
- Observability package: [[docs/RootCellar/04-Observability/Observability Charter]]
- Introspection playbooks: [[docs/RootCellar/04-Observability/AI Introspection Runbook]], [[docs/RootCellar/04-Observability/Execution Traceability Atlas]]
- Quality package: [[docs/RootCellar/05-Quality/Test Strategy]]
- Ops package: [[docs/RootCellar/06-Operations/Environment Matrix]]
- Architecture decisions: [[docs/RootCellar/07-ADRs/ADR-0001 Tauri + Web Grid + Rust Engine]]
- Delivery catalog: [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]

## Product Guardrails
- XLSX round-trip safety is a release blocker.
- Excel-like UX for top workflows is mandatory before broad rollout.
- Python automation ships only with enforced capability permissions and auditability.
- Deterministic mode is a strategic differentiator and must be first-class in engine and CLI.
- Every major subsystem must emit inspectable artifacts for human and AI introspection.

## Observability-First Rule
No feature is "done" unless it produces:
1. Structured events.
2. Trace links to user intent and artifact outputs.
3. Queryable diagnostics for failure and performance.
4. A reproducible artifact bundle where relevant.

## Canonical Vocabulary
- Preserve Mode: prioritize compatibility/passthrough.
- Normalize Mode: prioritize deterministic canonical output.
- Repro Record: immutable bundle of inputs, config, versions, outputs, checksums.
- Introspection Artifact: structured snapshot exposing state transitions or UI/engine/script decisions.
