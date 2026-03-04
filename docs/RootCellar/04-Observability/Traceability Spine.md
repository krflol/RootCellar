# Traceability Spine

Parent: [[docs/RootCellar/03-Implementation/UI-to-Engine Trace Bridge]]
Execution control: [[Execution Plan Board]], [[Execution Status]]
Execution atlas: [[Execution Traceability Atlas]], [[AI Introspection Runbook]]

## Why this exists
This note defines one deterministic path from user intent to persisted evidence so humans and AI can reconstruct behavior across:
- UI interaction
- Engine events and spans
- CLI-style artifacts
- Repro or CI outputs

## Canonical Trace Chain
- Intent command: `open`, `edit`, `save`, `recalc`, and `runEngineRoundTrip` in `apps/desktop/src/main.ts`.
- Trace root: `trace_id` in `TraceContext` from `crates/rootcellar-core/src/telemetry.rs`.
- UI action binding: `ui_command_id` / `ui_command_name` now thread through `apps/desktop/src-tauri/src/main.rs`.
- Event contract: command metadata now appears in response context via `TraceEcho` and UI trace header formatting.
- Artifact mapping: save/recalc reports and compatibility outputs include trace context where available, then can be indexed through manifest/inspection artifacts.

## Workflow Nodes
| Workflow | Request Surface | Trace IDs Added | Engine-Side Evidence | Desktop Evidence |
|---|---|---|---|---|
| Open workbook | `openWorkbook` | `trace_id`, `surface`, `surface_session_id`, `ui_command_id` | `interop.xlsx.load.end`, `interop.xlsx.part_graph.built` (when enabled), `artifact_id` if emitted | `Latest trace:` + `ui_command_id` fields in command output |
| Preview sheet | `loadSheetPreview` | same as open + request payload digest | preview pipeline events in CLI/engine layer | trace header section for preview |
| Apply cell edit | `applyCellEdit` | same + command name/id + parent span | `engine.txn.*`, `calc.recalc.*`, dependency/artifact timing when enabled | edit output + linked trace fields |
| Save workbook | `saveWorkbook` | same + workbook id if available | `interop.xlsx.save.end`, save graph flags (`relationships_preserved`, `unknown_parts_preserved`) | save output trace header and compatibility summary |
| Recalc sheet | `recalcLoadedWorkbook` | same | `calc.recalc.start/end`, `artifact.recalc.dep_graph.output` / `artifact.recalc.dag_timing.output` | recalc output and capture metadata |
| Round trip | `runEngineRoundTrip` | same + command payload digest | open + edit + save/recalc event sequence | capture artifacts for full command chain |

## Required Artifacts for Full Spine Closure
- Command outputs must carry:
  - `trace_id`
  - `trace_root_id` when available
  - `command_status`
  - `duration_ms`
- Engine response payload should preserve:
  - `trace` block with trace IDs
  - `ui_command_id` and `ui_command_name`
  - `artifact_id` / manifest link when produced
- For AI/human queries:
  - filter `events.jsonl` by trace_id
  - open corresponding command output block
  - map to compatible artifact index and report path

## Implementation Status
- Completed
  - `ui_command_id` and `ui_command_name` added to desktop command wrappers.
  - Command context now propagates through Tauri boundary into core `TraceContext`.
  - `TraceEcho` carries command metadata to response payload and command-output formatting.
- `desktop_trace_continuity_smoke_open_edit_save_recalc` validates open/edit/preview/save/recalc/round-trip chain integrity for `trace_id`, command metadata, status, duration, and linked artifact IDs in backend responses.
- `desktop_trace_continuity_smoke_open_edit_save_recalc` now asserts event-stream inclusion for the same trace root during open/edit/save/recalc and artifact-index parity for preview/round-trip.
  - `ROOTCELLAR_DESKTOP_EVENT_JSONL` and per-command `event_log_path` make event sinks deterministic for tests.
- In progress closure status is tracked from the same matrix used by program delivery: [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix#In Progress Stories]]

## Query Recipes
- Start with command output header:
  - Find `Latest trace:` line in `apps/desktop/src/main.ts` trace rendering path.
- Then walk event stream:
  - Search for same `trace_id` in `events.jsonl` and `trace.json` when present.
- Finally reconcile artifacts:
  - Compare generated report/manifest files named with run and trace suffixes in local or CI artifacts.
- AI-friendly reconstruction:
  - For each trace, reconstruct command intent -> event timeline -> artifact list -> user-visible UI output block.

## Open Validation Tasks
1. Standardize `artifact_id` naming and trace back references in desktop trace headers.
2. [x] Complete CI smoke with desktop open/edit/preview/save/recalc/round-trip run and assert common trace IDs exist in both command output and engine events.
   - Implemented in `desktop_trace_continuity_smoke_open_edit_save_recalc` with deterministic event-log assertions.
3. [x] Add a machine-readable UI trace schema for command metadata (schema test in CI).
   - Implemented in `schemas/desktop/v1/command-output-trace.schema.json` and `apps/desktop/src/desktopTraceOutput.test.ts`.
   - Production formatter and parser are centralized in `apps/desktop/src/desktopTraceOutput.ts`.
4. [x] Resolve linked artifact IDs to persistent artifact records for deterministic joins.
   - Implemented in `apps/desktop/src/desktopTraceJoin.ts` and CLI runner `apps/desktop/scripts/resolve-desktop-trace-artifacts.ts` (`npm run trace:join`).
