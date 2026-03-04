# Sprint 03 - Grid Editing Loop

Parent: [[Sprint Cadence and Capacity]]
Dates: April 13, 2026 to April 26, 2026

## Sprint Goal
Deliver usable desktop editing loop with selection, cell edit commit, clipboard baseline, and undo/redo.

## Execution Status
- Status: In progress.
- Tracking links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]], [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]

## Commitments
- Epic 03 primary.
- Epic 02 integration for formula edits.
- Epic 07 UI interaction instrumentation.

## Completed within Sprint
- Desktop editing-flow backbone: selection-driven preview rendering, formula-bar entry/apply action, range presets, and recalc freshness indicators wired end-to-end in desktop shell (`apps/desktop/src/main.ts`).
- Command continuity hardening: trace-linked `open/edit/preview/save/recalc/round-trip` path assertions and artifact-index joins are now smoke-asserted (`apps/desktop/src-tauri/src/main.rs`).
- Traceability tooling: command output schema + join helper tests and artifact-index join scripts added (`apps/desktop/src/desktopTraceOutput.ts`, `apps/desktop/src/desktopTraceJoin.ts`, `apps/desktop/scripts/resolve-desktop-trace-artifacts.ts`, `apps/desktop/src-tauri/Cargo.toml` tests).
- Preview/UI test coverage for navigation formulas and range preset behavior (`apps/desktop/src/previewInteractions.test.ts`, `apps/desktop/src/editRangePresets.test.ts`, `apps/desktop/src/presetReuse.test.ts`).

## Stories
1. Complete selection state machine and keyboard navigation paths (in progress).
2. Complete formula bar + engine transaction commit consistency (in progress).
3. Add clipboard copy/paste baseline for values and formulas (in progress).
4. Implement undo/redo stack using transaction digests (not started).
5. Add UI latency and error telemetry panels for edit loops (not started).

## Current Blockers
- Undo/redo history design depends on persisted transaction digest snapshots from `rootcellar-core`.
- Clipboard parity requires deterministic value/formula copy semantics and optional paste handling policy.
- Accessibility and focus management require a dedicated interaction review pass once selection/keyboard flows stabilize.

## Acceptance Criteria
- Common navigation and edit flows pass checklist.
- Undo/redo is consistent for value/formula/multi-cell edits.
- UI emits trace-linked events for edit lifecycle.

## Exit Signals
- Scroll and selection benchmarks running in CI.
- Accessibility semantic mirror prototype in place.
