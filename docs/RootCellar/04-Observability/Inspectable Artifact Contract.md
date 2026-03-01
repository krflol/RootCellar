# Inspectable Artifact Contract

Parent: [[Observability Charter]]

## Artifact Bundle Overview
Every critical run (desktop save, macro run, CLI batch task) can emit a bundle with manifest and typed artifacts.

## Bundle Structure
- `manifest.json`
- `events.jsonl`
- `trace.json`
- `state/` snapshots (workbook region, dependency graph slice)
- `interop/` compatibility report, part inventory, transform log
- `script/` script metadata, permission usage, stdout/stderr capture
- `ui/` command timeline, viewport diagnostics
- `perf/` benchmark samples and percentile summaries
- `checksums/` SHA-256 map for all files

## Manifest Required Fields
- `bundle_id`, `created_at`, `product_version`, `git_sha`, `mode`, `source_inputs`.
- `trace_roots[]`, `artifact_index[]`, `retention_policy`.

## Introspection Levels
- Level 1 Minimal: events + summary metrics.
- Level 2 Diagnostic: includes key snapshots and traces.
- Level 3 Forensic: full snapshots, logs, policy/audit details.

## Storage and Retention
- Local default for developer runs.
- Configurable remote sink for CI/staging/prod.
- Retention tiers by severity and compliance policy.

## Contract Tests
- Bundle validator checks required files/fields.
- Checksum verifier ensures artifact integrity.
- Schema validator ensures compatibility across versions.
- Batch artifact contract validator (`python/validate_batch_adapter_contracts.py --full-family`) enforces the nightly artifact-family schemas in `schemas/artifacts/v1/`:
  - `batch-throughput-snapshot.schema.json`
  - `batch-alert-dispatch.schema.json`
  - `batch-ack-retention-index.schema.json`
  - `batch-dashboard-pack.schema.json`
  - `batch-alert-policy.schema.json`
  - `batch-policy-escalation.schema.json`
  - `batch-dashboard-adapter-exports.schema.json`
- Schema-drift canary harness (`python/validate_batch_schema_canaries.py`) mutates canonical artifacts to assert expected validator failures for compatibility regressions.
- Migration process and semver/compatibility policy are tracked in [[Artifact Schema Migration Playbook]].
