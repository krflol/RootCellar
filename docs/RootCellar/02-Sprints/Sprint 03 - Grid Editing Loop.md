# Sprint 03 - Grid Editing Loop

Parent: [[Sprint Cadence and Capacity]]
Dates: April 13, 2026 to April 26, 2026

## Sprint Goal
Deliver usable desktop editing loop with selection, cell edit commit, clipboard baseline, and undo/redo.

## Commitments
- Epic 03 primary.
- Epic 02 integration for formula edits.
- Epic 07 UI interaction instrumentation.

## Stories
1. Implement selection state machine and keyboard navigation paths.
2. Wire formula bar to engine transaction commits.
3. Add clipboard copy/paste for values and basic formatting.
4. Implement undo/redo stack using transaction digests.
5. Add UI latency and error telemetry panels.

## Acceptance Criteria
- Common navigation and edit flows pass checklist.
- Undo/redo is consistent for value/formula/multi-cell edits.
- UI emits trace-linked events for edit lifecycle.

## Exit Signals
- Scroll and selection benchmarks running in CI.
- Accessibility semantic mirror prototype in place.