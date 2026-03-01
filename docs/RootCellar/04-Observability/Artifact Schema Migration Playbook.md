# Artifact Schema Migration Playbook

Parent: [[Observability Charter]]
Related: [[Inspectable Artifact Contract]], [[docs/RootCellar/06-Operations/CI-CD Blueprint]]

## Purpose
Define how RootCellar evolves artifact schemas without breaking humans, AI agents, or downstream ingestion systems.

## Versioning Contract
- Schema files are versioned under `schemas/artifacts/v1/*`.
- Each artifact payload must include `artifact_contract`:
  - `schema_id`
  - `schema_version` (semver)
  - `compatibility`
- Each schema includes `x-rootcellar-contract`:
  - `schema_version`
  - `payload_version_field`
  - `payload_version`
  - `compatibility_mode`
- Compatibility rule:
  - `backward-additive` requires matching major semver between payload and schema.

## Change Classes
- Additive change:
  - New optional fields only.
  - No major version bump.
  - Existing canary drift cases must still fail as expected.
- Breaking change:
  - Required field removals/renames or type narrowing.
  - Major semver bump.
  - New schema path/version family and dual-read migration window required.

## Migration Workflow
1. Propose schema change with impacted artifacts and downstream consumers.
2. Update schema JSON and generator payload(s).
3. Update validator contract behavior if compatibility policy changes.
4. Update canary drift suite in `python/validate_batch_schema_canaries.py`.
5. Run local validation:
   - `python python/validate_batch_adapter_contracts.py --full-family`
   - `python python/validate_batch_schema_canaries.py`
   - `python python/validate_batch_dual_read_migration.py`
6. Update CI knobs/workflow if new schema files are introduced.
7. Publish migration note in execution docs and incident runbook references.

## Canary Policy
- Canary harness must assert:
  - canonical artifact family passes
  - representative schema-drift scenarios fail with explicit error signals
- Current canary suite location: `python/validate_batch_schema_canaries.py`.
- Nightly gate location: `.github/workflows/batch-recalc-nightly.yml` (`Run schema drift canary checks`).

## Dual-Read Drill Policy
- Dual-read drill must assert five phases:
  1. producer `v1` -> consumer `v1` pass.
  2. producer `v2` -> consumer `v1` fail (rollback detection).
  3. producer `v2` -> consumer dual-read (`v1` primary + `v2` fallback) pass.
  4. producer rollback to `v1` -> consumer dual-read pass.
  5. producer `v1` -> consumer rollback to strict `v1` pass.
- Current drill suite location: `python/validate_batch_dual_read_migration.py`.
- Nightly gate location: `.github/workflows/batch-recalc-nightly.yml` (`Run dual-read migration drills`).

## Rollback Policy
- If canary or full-family schema validation fails in CI:
  - block artifact publication and release promotion.
  - revert schema/generator mismatch or disable new producer behavior behind a temporary compatibility flag.
  - retain last known-good schema paths in workflow env until compatibility is restored.
- If dual-read migration drills fail:
  - block schema major-version rollout.
  - restore consumer fallback coverage or roll producers back to previous major.
  - re-run canary + dual-read suites before reopening rollout.

## Migration Checklist
- [ ] Schema + payload contracts updated.
- [ ] Full-family validator passes locally.
- [ ] Schema-drift canaries pass locally.
- [ ] Dual-read migration drills pass locally.
- [ ] Nightly workflow knobs/manifest fields updated.
- [ ] Execution board/status evidence links updated.
- [ ] Incident playbook references reviewed.
