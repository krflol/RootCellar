# Epic 01 - XLSX Fidelity and Workbook Model

Parent: [[docs/RootCellar/00-Program/RootCellar Master Plan]]
Specs: [[docs/RootCellar/03-Implementation/Engine Workbook Model Spec]], [[docs/RootCellar/03-Implementation/XLSX Import Export Pipeline]]
Story registry: [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]

## Objective
Implement workbook model and XLSX import/export pipeline with high round-trip fidelity and preserve-first behavior.

## Execution Status
- Status: In progress (baseline delivered).
- Tracking links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]
- Completed slice:
1. Workbook model + XLSX load/save projection.
2. Preserve passthrough and selective worksheet override mode.
3. Compatibility report artifact from CLI `open`.
4. Part-graph baseline in interop reports (nodes/edges, dangling edge detection, known/unknown part classification).
5. Graph-aware save flags for preserve/normalize outputs in save artifacts and telemetry.
6. Corpus part-graph validation CLI command (`part-graph-corpus`) with aggregate reporting and fail-on-error mode.
7. CI workflow baseline for corpus part-graph artifact publication with deterministic generated fixtures.
- Remaining:
1. Broader corpus no-repair campaign with part-graph validation gates.
2. Deeper relationship validation (cross-part invariants and dangling-edge repair guidance).

## Scope
- Workbook entities and sparse cell store.
- XLSX parser/writer for core parts.
- Unknown-part passthrough registry.
- Preserve and normalize modes.

## Out Of Scope
- Full editing for all preserved-only features.
- Legacy XLS formats.

## Deliverables
- Engine model crate v1.
- XLSX load/save service with compatibility report output.
- Corpus no-repair nightly run in CI.

## Stories
1. Implement part graph reconstruction and validation.
2. Add shared strings/styles integration with preserve fallback.
3. Implement save pipeline with deterministic ordering option.
4. Emit compatibility findings per workbook.

## Acceptance Criteria
- >= 80% no-repair on initial corpus by Sprint 05.
- Unknown parts are preserved byte-stable in preserve mode where feasible.
- Save/open cycle does not drop known workbook metadata.

## Dependencies
- [[Epic 07 - Radical Observability and Introspection]] for trace/artifact contracts.

## Observability Requirements
- Parse/save stage timing metrics.
- Unknown part inventory artifact.
- Per-file compatibility report.
