# XLSX Import Export Pipeline

Parent: [[Architecture Overview]]
Related epic: [[docs/RootCellar/01-Epics/Epic 01 - XLSX Fidelity and Workbook Model]]

## Pipeline Stages
1. Zip ingest and central directory validation.
2. Part graph reconstruction from `_rels`.
3. Typed parsing for supported parts (workbook, worksheet, sharedStrings, styles, calc metadata).
4. Opaque capture for unsupported/unknown parts.
5. Normalized in-memory model projection.
6. Save writer with mode-specific behavior.

## Preserve Mode
- Default mode for desktop editing.
- Preserve unknown parts and element ordering where safe.
- Minimize rewrite footprint to reduce corruption risk.

## Normalize Mode
- Canonical ordering of XML attributes/elements where feasible.
- Drops non-preservable noise with explicit compatibility notes.
- Produces stable output for reproducibility workflows.

## Relationship Handling
- Maintain bidirectional map of part IDs and relationship targets.
- Detect dangling relationships and surface warnings.
- Preserve non-understood relationship types as opaque links.

## Validation
- Schema-level checks for required workbook/sheet structures.
- Cross-part consistency checks (shared string references, style IDs, defined names).
- Post-write reopen sanity check in CI.

## Compatibility Outputs
- Per-workbook compatibility report artifact with:
  - feature status matrix
  - unknown part inventory
  - part graph nodes/edges (including dangling targets)
  - transformations performed
  - potential risk hints
- Corpus-level compatibility artifact:
  - aggregate part graph counts and failure inventory via `part-graph-corpus`

## Telemetry
- `interop.xlsx.load.start|end`
- `interop.xlsx.part_graph.built`
- `interop.xlsx.save.start|end`
- `interop.xlsx.unknown_part.count`
- `interop.xlsx.repair_risk.detected`
