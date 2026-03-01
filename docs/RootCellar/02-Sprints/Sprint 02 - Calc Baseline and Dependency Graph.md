# Sprint 02 - Calc Baseline and Dependency Graph

Parent: [[Sprint Cadence and Capacity]]
Dates: March 30, 2026 to April 12, 2026

## Sprint Goal
Establish formula parser, dependency graph, and incremental recalc baseline for key function families.

## Execution Status
- Status: In progress.
- Tracking links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]
- Delivered subset:
1. Arithmetic parser scaffold with precedence and parentheses support.
2. Dependency graph analysis scaffold (edge inventory, topological ordering, cycle/parse diagnostics).
3. `calc.dependency_graph.built` telemetry event with inspectable dependency refs.
4. Recalc reporting artifacts and deterministic value fingerprints.
5. Optional dependency graph artifact output from CLI `recalc --dep-graph-report`.
6. Incremental invalidation baseline (`recalc_sheet_from_roots`) and integration into `tx-save`.
7. Optional DAG timing artifact output from CLI `recalc --dag-timing-report`.
8. Initial function coverage baseline (`SUM`, `MIN`, `MAX`, `IF`) with dependency graph function-call metrics.
9. AST interning scaffold with dedup metrics and per-formula AST IDs in dependency reports.
10. DAG analysis baseline with critical path, fan-in/fan-out, and slow-node threshold outputs.
11. Function coverage increment (`AVERAGE`/`AVG`, `ABS`, `AND`, `OR`, `NOT`) with parser/evaluator test coverage.
12. Configurable DAG slow-node threshold output via CLI (`recalc --dag-slow-threshold-us`) and core recalc options API.
- In progress / pending:
1. Broader function parity beyond current starter coverage with compatibility-focused behavior.
2. Incremental scheduler hardening for larger dependency graphs.
3. Additional DAG tuning/perf validation against benchmark workloads.

## Commitments
- Epic 02 primary.
- Epic 01 integration for formula cells and cached values.
- Epic 07 recalc trace artifact.

## Stories
1. Implement lexer/parser and AST intern pool.
2. Build dependency graph construction from formula references.
3. Implement invalidation and topological recompute baseline.
4. Add basic function registry (math/logical/text starter set).
5. Export recalc trace DAG artifact with node timings.

## Acceptance Criteria
- Recalc updates only dependents for edited inputs.
- Cycles detected and surfaced with human-readable path.
- Golden suite baseline established and automated in CI.

## Exit Signals
- p95 recalc latency tracked for benchmark set.
- Determinism mismatch detector wired but allowed-warning stage.
