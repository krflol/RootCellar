# Sprint 01 - Workbook and XLSX Skeleton

Parent: [[Sprint Cadence and Capacity]]
Dates: March 16, 2026 to March 29, 2026

## Sprint Goal
Load and save simple XLSX files with preserve-mode scaffolding and workbook model persistence.

## Execution Status
- Status: Completed (interop baseline).
- Tracking links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]
- Delivered stories:
1. Parse workbook/worksheet/shared strings into workbook model projection.
2. Unknown part preserve pipeline for passthrough and selective worksheet overrides.
3. CLI `open` compatibility summary and `save` preserve/normalize modes.
4. Parser/save telemetry and verification command runs documented.
5. Part-graph baseline emitted in compatibility reports and save outputs (with preserve/normalize graph flags).
6. Corpus part-graph validator command delivered for recursive XLSX scans and aggregate artifact generation.
7. CI workflow baseline delivered for corpus validator artifact publication.
- Remaining beyond baseline:
1. Corpus-scale part graph validation and relationship-invariant hardening.

## Commitments
- Epic 01 primary.
- Epic 07 instrumentation hardening.
- Epic 05 initial CLI open/report command.

## Stories
1. Implement zip ingest and part graph reconstruction.
2. Parse workbook/worksheet basics and shared strings.
3. Implement unknown part capture and re-emit pipeline.
4. Add CLI `open --report` compatibility summary.
5. Add parser telemetry for part counts and parse durations.

## Acceptance Criteria
- Can round-trip simple multi-sheet workbook without repair prompt.
- Unknown parts counted and listed in report artifact.
- Save path includes deterministic ordering toggle flag.

## Exit Signals
- Nightly corpus smoke run created with at least 30 seed files.
- No critical parser panics in fuzz smoke.
