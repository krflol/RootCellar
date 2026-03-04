# AI Introspection Runbook

Parent: [[docs/RootCellar/04-Observability/Observability Charter]]
Execution control: [[Execution Traceability Atlas]], [[docs/RootCellar/00-Program/Execution Status]]

## Purpose
Give humans and AI a reliable path from an observed behavior to root cause using only repository artifacts and trace IDs.

## Fast Triage Playbooks
1. Confirm trace continuity for a suspicious command (`open`, `edit`, `save`, `recalc`).
   - Inspect latest command output for `Latest trace:` block.
   - Extract `trace_id`, `ui_command_id`, `ui_command_name`, and any `artifact_id`.
   - Cross-check those IDs in `events.jsonl` and artifact manifests.
2. Resolve a failing workbook save/recalc path.
   - Verify `interop.xlsx.save.end` or `calc.recalc.*` events for matching `trace_id`.
   - Pull associated manifest and compare `workbook_id` / `surface_session_id` continuity.
   - Correlate with run command output and capture image (`desktop-save-recalc.png`, etc.).
3. Diagnose formula or recalc regression.
   - Compare `artifact.recalc.dep_graph.output` / `artifact.recalc.dag_timing.output`.
   - Review `calc.recalc.*` timing spans and impacted-cell ordering.
   - Compare against prior stable report in `python/` benchmark outputs and CI thresholds.
4. Investigate schema or migration drift.
   - Validate artifact against `schemas/artifacts/v1/*`.
   - Inspect migration drill/fault-injection reports for version and compatibility mismatches.

## Query Recipes
- event stream: `rg "trace_id\":\"<id>\"` events.jsonl`
- artifact manifest: `rg "artifact_id|trace_id|run_id|command_id" -g "*.json" target .`
- command output: search for `ui_command_id` and `command_status` in desktop capture transcripts.
- deterministic desktop join: `cd apps/desktop && npm run trace:join -- --trace-output <file> --artifact-index <file> --trace-id <id>`
- decision history: check linked epic/story in [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]].

## Root-Cause Workflow
- Step 1: identify user-visible symptom in status line, capture, or output.
- Step 2: collect `trace_id` and `ui_command_id` from the same command surface.
- Step 3: map trace to event stream and event-level artifacts.
- Step 4: map trace to persisted artifacts (`artifact.*.output`, reports, migration/check bundles), then run the deterministic join utility when `linked_artifact_ids` is populated to locate exact artifact-index records.
- Step 5: map trace to decision docs (epic → sprint → plan item).
- Step 6: determine whether to file a bugfix ticket, hardening task, or doc correction.

## AI Readable Evidence Bundle
- command output block
- engine trace events
- artifact report + schema id
- mutation/txn outputs (if mutation path)
- compatibility report
- test/CI artifacts that prove regression coverage

## Artifact Expectations by Surface
- Desktop QA output: `Latest trace`, `ui_command_id`, `command_status`, `duration_ms`.
- Engine events: `interop.*`, `calc.*`, `artifact.*`, `dependency.*` when enabled.
- Repro bundles: `repro record`, `repro check`, `repro diff`.
- Batch governance: throughput snapshot, dashboard pack, alert/policy artifacts.

## Escalation Rules
- Blocked UI evidence flow (no trace header): stop and fix before any feature merge.
- Missing artifact linkage for any mutation command: move to blocker list immediately.
- Schema mismatch in output artifacts: trigger migration drift review before merging.

## Completion Checklist
- `ui_command_id` and `trace_id` exist in UI command output.
- At least one matching engine event exists for same trace.
- At least one deterministic artifact link exists for the same trace.
- Deterministic manifest or CI report is available for replay.
- The investigation path is recorded in linked execution notes.
