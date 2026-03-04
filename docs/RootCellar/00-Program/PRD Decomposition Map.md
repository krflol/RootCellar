# PRD Decomposition Map

Parent: [[docs/RootCellar PRD]]
Execution control: [[docs/RootCellar/00-Program/RootCellar Master Plan]], [[Execution Plan Board]], [[Execution Status]]
Traceability execution map: [[docs/RootCellar/04-Observability/Execution Traceability Atlas]], [[docs/RootCellar/04-Observability/AI Introspection Runbook]]

## Purpose
This note turns the PRD into execution-ready slices aligned to existing epics, implementation specs, and sprint sequencing.

## PRD Area Coverage Matrix
- Vision and replacement contract:
  - Primary: [[docs/RootCellar/01-Epics/Epic 01 - XLSX Fidelity and Workbook Model]], [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/01-Epics/Epic 03 - Desktop Grid UX and Productivity]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
  - Current execution linkage: [[docs/RootCellar/00-Program/Execution Plan Board]]
- Core stack and technology:
  - Primary: [[docs/RootCellar/03-Implementation/Architecture Overview]], [[docs/RootCellar/03-Implementation/UI Grid and Interaction Design]], [[docs/RootCellar/06-Operations/CI-CD Blueprint]]
  - Current execution linkage: [[docs/RootCellar/03-Implementation/Architecture Overview#Boundary Rules]]
- Workbook model and partial-fidelity preservation:
  - Primary: [[docs/RootCellar/01-Epics/Epic 01 - XLSX Fidelity and Workbook Model]], [[docs/RootCellar/03-Implementation/Engine Workbook Model Spec]], [[docs/RootCellar/03-Implementation/XLSX Import Export Pipeline]]
  - Current sprint anchor: [[docs/RootCellar/02-Sprints/Sprint 01 - Workbook and XLSX Skeleton]], [[docs/RootCellar/02-Sprints/Sprint 05 - Compatibility Panel and Corpus Tests]]
- Calculation engine and determinism:
  - Primary: [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]], [[docs/RootCellar/03-Implementation/Calculation Engine Design]], [[docs/RootCellar/00-Program/Dependency Map]]
  - Current execution linkage: [[docs/RootCellar/02-Sprints/Sprint 02 - Calc Baseline and Dependency Graph]]
- Grid UX and editing workflow:
  - Primary: [[docs/RootCellar/01-Epics/Epic 03 - Desktop Grid UX and Productivity]], [[docs/RootCellar/03-Implementation/UI Grid and Interaction Design]]
  - Current execution linkage: [[docs/RootCellar/02-Sprints/Sprint 03 - Grid Editing Loop]]
- Scripting platform and sandbox:
  - Primary: [[docs/RootCellar/01-Epics/Epic 04 - Python Automation Platform]], [[docs/RootCellar/03-Implementation/Python Scripting Host Design]]
  - Current execution linkage: [[docs/RootCellar/02-Sprints/Sprint 04 - Python Macros Alpha]]
- Headless and automation:
  - Primary: [[docs/RootCellar/01-Epics/Epic 05 - Headless CLI and SDK]], [[docs/RootCellar/03-Implementation/CLI and SDK Design]], [[docs/RootCellar/03-Implementation/Compatibility Panel Design]]
  - Current execution linkage: [[docs/RootCellar/02-Sprints/Sprint 06 - CLI Batch and Repro Mode]]
- Compatibility reporting and migration:
  - Primary: [[docs/RootCellar/01-Epics/Epic 06 - Compatibility and Migration Tooling]], [[docs/RootCellar/04-Observability/Artifact Schema Migration Playbook]], [[docs/RootCellar/03-Implementation/Compatibility Panel Design]]
  - Current execution linkage: [[docs/RootCellar/02-Sprints/Sprint 05 - Compatibility Panel and Corpus Tests]]
- Observability and introspection:
  - Primary: [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]], [[docs/RootCellar/04-Observability/Observability Charter]], [[docs/RootCellar/04-Observability/Telemetry Taxonomy and Event Schema]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/04-Observability/Trace Correlation Model]]
  - Current execution linkage: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]
- Story and traceability registry:
  - Primary: [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]
  - Playbooks: [[docs/RootCellar/04-Observability/Execution Traceability Atlas]], [[docs/RootCellar/04-Observability/AI Introspection Runbook]]
  - Current execution linkage: [[docs/RootCellar/04-Observability/Traceability Spine]]
- Enterprise trust, signing, and policy:
  - Primary: [[docs/RootCellar/01-Epics/Epic 08 - Enterprise Trust and Distribution]], [[docs/RootCellar/06-Operations/Incident Response Playbook]], [[docs/RootCellar/03-Implementation/Security and Permission Model]]
  - Current execution linkage: [[docs/RootCellar/02-Sprints/Sprint 07 - Security Hardening and Add-in Signing]]

## PRD->Epic Mapping to Current Delivery Slices
1. "XLSX open/edit/save without breakage" -> Epic 01 and Epic 06.
2. "Excel-like UX for common workflows" -> Epic 03 plus Epic 07 instrumentation.
3. "Formula correctness and parser parity" -> Epic 02 plus Epic 05 CLI reproducibility surfaces.
4. "Python macros/events/UDFs with sandbox" -> Epic 04 and Epic 08 security completion.
5. "Headless automation and reproducibility" -> Epic 05 with Epic 07 traceability.
6. "Radical observability end-to-end" -> Epic 07 + 04 and cross-surface traceability implementation.

## Delivery Dependencies
- M0 core readiness requires:
  - Foundation telemetry contract stability.
  - Desktop startup and shell command routing.
  - Trace linkage from UI command entry to engine transaction and artifact output.
- Dependency reference set:
  - [[docs/RootCellar/00-Program/Dependency Map]]
  - [[docs/RootCellar/00-Program/Risk Register]]
  - [[docs/RootCellar/00-Program/Decision Register]]

## Next Decomposition Targets
1. Convert each PRD functional requirement bucket into ticketized user stories under active sprint docs.
2. Link each PRD bucket to one or more rows in `[[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]`.
3. Add explicit artifact ID and trace-id expectation per requirement bucket in [[docs/RootCellar/04-Observability/Desktop Inspectable Artifact Map]].
4. Extend `[[docs/RootCellar/04-Observability/Traceability Spine]]` with one-click query recipes per workflow, and add CI-gated acceptance checks.
5. Add a per-bucket AI-readable introspection path in [[docs/RootCellar/04-Observability/AI Introspection Runbook]].
