# Trace Correlation Model

Parent: [[Observability Charter]]

## Goal
Allow one-click traversal from user action to all downstream engine/script/interop effects.

## Correlation IDs
- `trace_id`: one user command or CLI invocation.
- `span_id`: operation segment within trace.
- `parent_span_id`: hierarchical relation.
- `txn_id`: engine mutation transaction key.
- `artifact_id`: output artifact checksum key.
- `script_run_id`: unique script invocation key.

## Propagation Rules
- UI command starts trace and span.
- Engine calls inherit trace, create child spans per major operation.
- Script RPC includes trace context headers and creates nested spans in worker.
- CLI commands create trace root at command start.

## Critical Trace Paths
1. Cell edit -> engine commit -> recalc -> UI repaint -> artifact emission.
2. File open -> parse pipeline -> compatibility analysis -> panel render.
3. Macro run -> permission checks -> RPC operations -> workbook mutations -> audit log.
4. Batch run -> per-file worker spans -> summary report.

## Completeness KPI
- Critical path traces with no missing span links >= 95% in staging.
- Any drop below threshold triggers release gate warning.

## Debugging Workflow
- Start from error event.
- Traverse to parent trace and sibling spans.
- Load related artifact manifest.
- Compare expected vs actual state snapshots.