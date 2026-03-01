# UI Grid and Interaction Design

Parent: [[Architecture Overview]]
Related epic: [[docs/RootCellar/01-Epics/Epic 03 - Desktop Grid UX and Productivity]]

## Rendering Architecture
- Canvas/WebGL layer for visible cell rendering and scroll performance.
- DOM overlay for editor, formula bar, context menus, accessibility elements.
- Viewport manager computes virtualized window plus prefetch buffer.

## Interaction Model
- Selection states: active cell, range selection, multi-range selection.
- Keyboard navigation parity for arrows/tab/enter/page/home/end and modifier variants.
- Edit lifecycle: enter edit, commit, cancel, formula autocomplete, IME path.

## Clipboard and Interop
- Internal rich clipboard payload plus plain text fallback.
- Excel interop path for values + baseline formatting.
- Paste preview warnings for unsupported features.

## Undo/Redo
- Command stack references engine transactions, not raw UI state.
- Grouped operations for fill/autofill and multi-cell edits.

## Accessibility Strategy
- Semantic mirror for focused region and navigation announcements.
- Keyboard-first command coverage map.
- Screen reader event hooks for focus/value/formula changes.

## Performance Budgets
- Scroll render target: <= 16 ms frame budget in common workload.
- Selection move target: <= 50 ms response.
- Clipboard paste target: <= 200 ms for medium ranges.

## Observability Hooks
- `ui.grid.scroll`, `ui.grid.render.frame`
- `ui.edit.begin|commit|cancel`
- `ui.selection.change`
- `ui.command.invoked`