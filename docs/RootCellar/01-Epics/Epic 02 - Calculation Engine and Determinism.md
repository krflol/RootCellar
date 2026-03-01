# Epic 02 - Calculation Engine and Determinism

Parent: [[docs/RootCellar/00-Program/RootCellar Master Plan]]
Specs: [[docs/RootCellar/03-Implementation/Calculation Engine Design]]

## Objective
Deliver a reliable Excel-like calculation engine with incremental recalc and deterministic mode.

## Execution Status
- Status: In progress.
- Tracking links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]
- Completed slice:
1. Arithmetic recalc baseline with cycle detection.
2. Parser scaffold for operator precedence and parentheses.
3. Dependency graph analysis scaffold with topological ordering and diagnostics.
4. Incremental invalidation baseline (`recalc_sheet_from_roots`) integrated into `tx-save`.
5. DAG timing artifact baseline (`recalc --dag-timing-report`) with per-node timing summaries.
6. Function coverage increment (`SUM`, `MIN`, `MAX`, `IF`, `AVERAGE`/`AVG`, `ABS`, `AND`, `OR`, `NOT`) with parser/evaluator integration.
7. AST interning scaffold with dedup metrics and stable per-formula AST IDs for introspection.
8. DAG analysis baseline (critical path, fan-in/fan-out, slow-node thresholds) emitted in timing artifacts.
9. Configurable DAG slow-node threshold and deterministic tie-break behavior in DAG introspection paths.
- Remaining:
1. Broader function registry parity beyond current starter set and deterministic scheduler guarantees.
2. Additional DAG analysis refinement and performance tuning at larger scales.

## Scope
- Formula parser and AST.
- Dependency graph and invalidation.
- Evaluator and function registry.
- Deterministic scheduler and repro checks.

## Deliverables
- Function parity matrix and rollout plan.
- Incremental recalc engine with cycle reporting.
- Deterministic mode replay validation command.

## Stories
1. Build parser + AST store with span metadata.
2. Build dependency graph and cycle detection.
3. Add evaluator for baseline function families.
4. Integrate Rayon scheduler with stable tie-breakers.
5. Implement recalc trace artifact export.

## Acceptance Criteria
- Golden suite threshold >= 99.5% for in-scope functions.
- Deterministic replay hash stable across OS in supported benchmark set.
- Cycle errors include clear path diagnostics.

## Dependencies
- [[Epic 01 - XLSX Fidelity and Workbook Model]]
- [[Epic 07 - Radical Observability and Introspection]]

## Observability Requirements
- Node-level timing traces.
- Invalidated subgraph size metrics.
- Determinism mismatch alerts.
