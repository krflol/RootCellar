# UI to Engine Trace Bridge

Parent: [[docs/RootCellar/03-Implementation/Architecture Overview]]
Specs: [[docs/RootCellar/04-Observability/Trace Correlation Model]], [[docs/RootCellar/04-Observability/Telemetry Taxonomy and Event Schema]], [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]], [[docs/RootCellar/01-Epics/Epic 07 - Radical Observability and Introspection]]
Operational playbooks: [[docs/RootCellar/04-Observability/Execution Traceability Atlas]], [[docs/RootCellar/04-Observability/AI Introspection Runbook]]

## Purpose
Define the exact contract that makes one desktop action traceable from intent, through engine mutation, to persisted artifacts.

## Problem Definition
Desktop currently emits command outputs and trace headers for QA, but trace lineage is not consistently surfaced as:
- stable UI action identity
- engine transaction identity
- downstream artifact identifiers
- script/external command correlation

This bridge fills that gap.

## Current UI Command Set
- `openWorkbook`
- `loadSheetPreview`
- `applyCellEdit`
- `saveWorkbook`
- `recalcLoadedWorkbook`
- `runEngineRoundTrip`
- future: script launch, batch run, macro execution

## Current Delivery State
- Delivered: `ui_command_id` and `ui_command_name` now move from desktop command entry into Tauri command wrappers and then into shared `TraceContext` (`trace`, `ui_command_id`, `ui_command_name`).
- Delivered: CI-level event-lookup assertions for open/edit/preview/save/recalc/round-trip continuity are implemented in
  `desktop_trace_continuity_smoke_open_edit_save_recalc` and wired to event logs via `event_log_path`.
- Trace/path continuity coverage now also validates `interop_sheet_preview` and `engine_round_trip` command-output artifact-index parity.
- Evidence: [[apps/desktop/src-tauri/src/main.rs]], [[docs/RootCellar/04-Observability/Traceability Spine]], [[docs/RootCellar/04-Observability/Execution Traceability Atlas]], [[apps/desktop/src/main.ts]], [[apps/desktop/src-tauri/src/main.rs]], [[crates/rootcellar-core/src/telemetry.rs]]

## Mandatory Trace Context
All desktop command invocations should include:
1. `surface`: fixed to `desktop`
2. `surface_session_id`: random UUID per app runtime
3. `ui_command_id`: random UUID per command button/menu action
4. `trace_id`: propagated to all backend calls and logs
5. `span_id`: current UI operation span
6. `workbook_id`: stable workbook identity for the active file/session
7. `command_payload_digest`: hash of normalized command payload for post-incident replay

## Ingress Contract (Frontend to Engine/CLI Surface)
- Frontend posts all command invocations with `{trace_context, request_payload}` wrapper.
- Engine command handlers must:
  - echo `trace_id`, `trace_context`, and `ui_command_id` in successful response metadata
  - echo command metadata (`ui_command_name`, `command_payload_digest`) where available
  - attach IDs to any produced artifacts (compatibility report, session preview snapshot, recalc report, save output)
  - preserve a deterministic mapping from `ui_command_id` to `txn_id` where mutation occurs

## Egress Contract (Back to UI)
- UI logs an observable payload per command including:
  - `trace_id`
  - `"trace_root_id"`
  - `"linked_artifact_ids"`
  - `command_status`
  - `duration_ms`
  - `ui_error` fields on failure

## Artifact Expectations by Command
### openWorkbook
- `workbook_open` event in trace graph.
- `interop.xlsx.load.end` includes part-graph and compatibility context.
  - UI compatibility preview text should include `trace_id` lines for traceability.

### applyCellEdit
- `engine.txn.open` and `engine.txn.commit` with mutation digest.
- `calc.recalc.start` and `calc.recalc.end` for changed cell and impacted set.
- UI should map edit output to returned transaction id.

### saveWorkbook
- `interop.xlsx.save.end` includes `trace_id`, `artifact_id`, and save strategy.
- UI compatibility report should reference save artifact id.

### recalcLoadedWorkbook
- `calc.recalc.start|end` with `recalc_scope`.
- If DAG artifact enabled, `artifact.recalc.dag_timing.output` should carry `trace_id`.

## Implementation Tickets
1. [x] Add a single trace context builder in desktop app state.
2. [x] Stamp commands at click/keyboard invocation before async work starts.
3. [x] Persist current trace context in each visible command section for QA capture.
4. [x] Thread IDs through preview/demo state and output renderers so artifact output can include trace lineage.
5. [ ] Add invariant tests:
   - [x] one command produces matching `ui_command_id`/`trace_id` through command output (unit smoke in `desktop_trace_continuity_smoke_open_edit_save_recalc`).
    - [x] artifact id appears in both command output and bundle index record when enabled.
6. [x] Add CI assertions for desktop trace continuity across open/edit/preview/save/recalc/round-trip with matching event/index lookup.

## Acceptance Criteria
- Any command emitted from desktop UI has non-null `trace_id`.
- UI-visible output for mutation commands includes:
  - linked `trace_id`
  - linked `txn_id` when mutation occurred
  - linked artifact reference if output artifact exists
- Command round-trip and trace continuity are smoke-tested in unit tests, with CI workflow assertions now implemented.

## Related Execution Tracking
- Current execution status: [[Execution Status]]
- Plan board update expected for this slice: [[Execution Plan Board]]
- Desktop implementation touchpoints:
  - [[apps/desktop/src/main.ts]]
  - [[docs/RootCellar/04-Observability/AI Introspection Runbook]]
