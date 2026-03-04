# Desktop Inspectable Artifact Map

Parent: [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]]
Related: [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]

## Goal
Make each desktop workflow reproducible from UI artifact to backend artifact chain, with stable human and AI inspection paths.

## Artifact Surfaces
- `desktop-capture`:
  - `ui` capture images in `apps/desktop/artifacts/ui-captures`.
  - Deterministic scenarios include `fresh`, `stale`, `pending`, `edit-cell`, `save-recalc`, and `mobile-stale`.
- `compatibility-output`:
  - command output panel lines with `Latest trace:` header from `openWorkbook`, `loadSheetPreview`, and related commands.
  - compatibility JSON preview payload with `trace` context.
- `edit-output`:
  - command payload + trace metadata in `applyCellEdit` output.
- `save-output`:
  - artifact reference line for save events and trace lineage in save and round-trip commands.
- `recalc-output`:
  - recalc duration and trace lineage lines for full and incremental recalc paths.

## Traceable Relationships
1. `ui_command_id` -> `trace_id` (request context).
2. `trace_id` -> command output block.
3. `trace_id` -> engine events (`ui.*`, `engine.*`, `calc.*`, `interop.*` where available).
4. `trace_id` -> artifact manifest entries (`artifact.*` or offline report files).
5. `trace_id` -> UI screenshot capture set for human review.

## Artifact Index by Workflow
- Open workbook workflow:
  - Trace start: desktop event invocation.
  - Primary outputs: command output (`Latest trace`), compatibility JSON snippet, optional `interop.xlsx.open` event in engine logs.
- Edit workflow:
  - Trace start: command invocation.
  - Primary outputs: edit preview output, dependency recalc event evidence, command output trace header.
- Save + recalc workflow:
  - Trace start: save command invocation.
  - Primary outputs: compatibility output from open/save cycle, recalc trace output, artifact linkage line, screenshot capture `desktop-save-recalc.png`.
- Macro/scripting workflow (in progress):
  - CLI trace start: `run-macro` command path currently emits macro/session/permission artifacts in JSONL.
  - Current outputs: script span trace (`script.session.*`), permission events (`script.permission.granted|denied`), and macro completion trace (`script.macro.run`).
  - Planned desktop outputs: same artifact continuity in `main.ts` command orchestration and user policy prompt telemetry.

## Query Cookbook
- Start with latest trace: search event store for `event_name: ...` and `trace_id: ...`.
- Next step: map trace_id to manifest index for artifact discovery.
- Next step: read UI output block and screenshot pair for UX-level context.
- Next step: compare `txn_id` between UI and engine to isolate mutation correctness issues.
- Deterministic artifact join (desktop local debug):
  - `cd apps/desktop`
  - `npm run trace:join -- --trace-output <path-to-capture-output> --artifact-index <path-to-desktop-artifact-index.jsonl>`
  - Optional: `--trace-id <trace_id|trace_root_id>` to filter a single command chain.

## CI Visibility Checks
- `desktop-ui-capture` workflow publishes capture and manifest metadata for each seeded state.
- Trace-continuity checks currently assert:
   - one trace id appears in open/edit/preview/save/recalc/round-trip command outputs and remains stable across command chain
  - trace id appears on at least one artifact output file path
  - command-level artifact index records are emitted to `ROOTCELLAR_DESKTOP_ARTIFACT_INDEX` target when configured
  - Traceability spine references: [[docs/RootCellar/04-Observability/Traceability Spine]]
  - Unit assertion point: [[apps/desktop/src-tauri/src/main.rs]]

## Open Risks
- None newly introduced in desktop trace continuity after command-output schema closure.
- Long-term risk remains: command-output line formatting can drift if formatting fields are changed outside the desktop trace-output helper.
- Mitigation: keep `desktopTraceOutput` as single formatter and validate all header lines through command-output schema tests.

## Next Implementation Actions
1. [x] Add explicit `trace_context` in every invoke to backend command boundary.
   - Completed via `toTraceInput` in `apps/desktop/src/main.ts` and per-command trace propagation into `loadSheetPreview`.
2. [x] Add artifact-index emission from desktop session mode for local debugging. (wire up index sink via `ROOTCELLAR_DESKTOP_ARTIFACT_INDEX`, command-level `artifact_refs`, and trace output path)
3. [x] Add smoke test for trace presence across open, edit, preview, save, recalc, and round-trip commands.
4. [x] Add command-output schema snapshot for `ui_command_id` and artifact linkage fields.
   - Completed via `schemas/desktop/v1/command-output-trace.schema.json`, `apps/desktop/src/desktopTraceOutput.ts`, and `apps/desktop/src/desktopTraceOutput.test.ts`.
5. [x] Execute closure checks from [[docs/RootCellar/04-Observability/AI Introspection Runbook]] in each command-path validation.
   - Deterministic artifact joins are now available via `apps/desktop/scripts/resolve-desktop-trace-artifacts.ts`.
