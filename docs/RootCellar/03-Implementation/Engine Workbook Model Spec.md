# Engine Workbook Model Spec

Parent: [[Architecture Overview]]
Related epic: [[docs/RootCellar/01-Epics/Epic 01 - XLSX Fidelity and Workbook Model]]

## Scope
Define engine-side structures and APIs that support high-fidelity XLSX round-trip and efficient mutation/recalc.

## Required Modules
- `model_workbook`: workbook, sheet, table, names, metadata.
- `model_cells`: sparse cell store, formula/value payloads, style refs.
- `model_layout`: row/column props, freeze panes, hidden/grouping state.
- `model_features`: conditional format, validation, merges, hyperlinks.
- `model_parts`: unknown part preservation and relationship registry.

## Sparse Cell Store Design
- Chunked sparse map keyed by row block + column offset.
- Compression strategy for empty runs.
- Fast range iteration path for aggregation and export.

## Style Handling
- Maintain style table references, avoid eager style expansion.
- Preserve unmapped style records through pass-through registry.
- Provide style projection for UI subset editing.

## Partial Fidelity Contract
Feature not yet editable is still preserved if loaded.
- Unknown elements stored in part registry with byte-preserving payload.
- Save pipeline reinserts payload in original relationship graph position where possible.

## APIs
- `Workbook::begin_txn()`
- `Txn::apply(op)`
- `Txn::validate()`
- `Txn::commit()` -> `CommitResult { txn_id, changed_cells, invalidated_nodes, events }`
- `Workbook::snapshot(region, detail_level)` for introspection.

## Failure Modes
- Corrupt XML: load with recoverable warnings where possible.
- Unsupported feature edit attempt: reject with capability error and compatibility note.
- Oversized workbook memory pressure: trigger spill and warning telemetry.

## Observability Hooks
- `engine.txn.begin`, `engine.txn.commit`, `engine.txn.rollback`
- `engine.model.preserve_part_attached`
- `engine.model.invariant_violation`