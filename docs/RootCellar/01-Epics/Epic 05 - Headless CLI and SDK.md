# Epic 05 - Headless CLI and SDK

Parent: [[docs/RootCellar/00-Program/RootCellar Master Plan]]
Specs: [[docs/RootCellar/03-Implementation/CLI and SDK Design]]

## Objective
Make RootCellar first-class in automation pipelines through CLI and SDK surfaces.

## Execution Status
- Status: In progress (CLI baseline delivered).
- Tracking links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]
- Completed slice:
1. CLI command lifecycle (`open`, `save`, `recalc`, `tx-demo`, `tx-save`).
2. Repro command family (`record`, `check`, `diff`) with deterministic comparisons.
3. JSONL event emission with trace/span IDs and artifact events.
4. CI workflow baseline for reproducibility bundle artifact publication (`.github/workflows/repro-bundle.yml`).
5. Repro/corpus CI artifact naming and retention policy alignment implemented for publishable bundles.
6. Batch directory recalc delivery with bounded Rayon parallel scheduling (`batch recalc` with deterministic ordering, `--threads`, `--fail-on-errors`, `--detail-level`).
7. Nightly batch artifact publication baseline (`.github/workflows/batch-recalc-nightly.yml`) with standardized naming/retention manifest policy.
8. Nightly batch corpus coverage expanded to deterministic 32-file compatibility slice with explicit throughput regression thresholds.
9. Nightly batch trend snapshot and alert-hook artifacts integrated for downstream dashboard/incident pipelines.
10. Nightly alert payload dispatch integrated for configured incident/dashboard endpoints with dispatch report artifacts.
11. Dispatch reliability/security baseline added with token auth, retry/backoff policy, and acknowledgement tracking.
12. Dispatch traceability baseline added with deterministic idempotency keys and correlation IDs across endpoint routes.
13. Dispatch replay-protection policy and acknowledgement-retention index publication integrated into nightly batch workflow.
14. Nightly dashboard-pack and alert-policy artifact generation integrated with CI gate enforcement.
15. Policy-to-owner escalation metadata and downstream dashboard/incident adapter export artifacts integrated into nightly workflow.
16. Adapter export schema/compatibility validation integrated into nightly workflow for contract enforcement before publication.
17. Full artifact-family schema/compatibility validation integrated into nightly workflow for snapshot/dispatch/ack-retention/dashboard-pack/policy plus adapter/escalation contracts.
18. Schema-drift canary fixture gate integrated into nightly workflow (`python/validate_batch_schema_canaries.py`) with policy knob control and manifest exposure.
19. Dual-read migration drill gate integrated into nightly workflow (`python/validate_batch_dual_read_migration.py`) with full artifact-family matrix coverage, artifact-subset + staged-wave + fault-injection policy knob control, structured diagnostics report output, and manifest exposure.
- Remaining:
1. SDK public API stabilization and Python bindings plan implementation.

## Scope
- CLI commands for load/recalc/macro/export.
- Batch processing with parallel scheduling.
- JSONL reporting and repro record/check workflows.
- Rust SDK baseline and Python bindings plan.

## Deliverables
- CLI v1 command set.
- Artifact bundle generation.
- Nightly corpus and batch performance jobs in CI.

## Stories
1. Build command parser and session lifecycle.
2. Implement headless mode integration with engine and script host.
3. Add JSONL emitter and schema validation.
4. Implement repro record/check command family.

## Acceptance Criteria
- Batch jobs operate reliably across directories with bounded parallelism.
- JSONL output schema stable and documented.
- Repro record/check catches deterministic mismatches.

## Dependencies
- [[Epic 01 - XLSX Fidelity and Workbook Model]]
- [[Epic 02 - Calculation Engine and Determinism]]
- [[Epic 04 - Python Automation Platform]]

## Observability Requirements
- Headless run trace coverage.
- Throughput and failure dashboards.
- Artifact registry integration.
