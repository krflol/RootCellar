# Calculation Engine Design

Parent: [[Architecture Overview]]
Related epic: [[docs/RootCellar/01-Epics/Epic 02 - Calculation Engine and Determinism]]

## Components
- Lexer/parser for Excel-like grammar.
- AST store with interned nodes.
- Dependency graph (cell/formula/name nodes).
- Evaluator with function registry and error model.
- Recalc scheduler using Rayon for independent subgraphs.

## Parser and AST
- Parse formulas into normalized AST with source span mapping.
- Support sheet-qualified refs, names, structured refs, range expressions.
- Store parse warnings for compatibility panel and formula editor hints.

## Dependency Graph
- Directed graph where edges represent data dependencies.
- On mutation, invalidate impacted nodes and recompute reachable dependents.
- Cycle detection returns explicit cycle path and classification (direct/indirect).

## Evaluation Strategy
- Scalar-first baseline with vectorized range operations where safe.
- Cache policy keyed by node version and dependency fingerprints.
- Function registry tags: deterministic, volatile, external-effect.

## Determinism Controls
- Stable topological tie-break by sheet_id then cell address.
- Deterministic reduction order for parallel aggregations.
- Configurable numeric precision strategy documented with caveats.

## Array Semantics
- Support classic CSE arrays in Phase A.
- Dynamic arrays staged in Phase B with spill range conflict detection.

## Telemetry and Artifacts
- `calc.recalc.start|end`
- `calc.node.evaluate`
- `calc.cycle.detected`
- Artifact: recalc trace DAG with timings and dependency edges.

## Exit Criteria
- Golden parity suite thresholds met.
- Cross-platform deterministic replay checks green in deterministic mode.