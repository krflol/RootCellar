# Data Model and IDs

Parent: [[Architecture Overview]]

## ID Strategy
- `workbook_id`: UUID v7, immutable per loaded workbook session.
- `sheet_id`: monotonic u64 assigned at load/create, persisted in sidecar metadata where needed.
- `cell_addr`: canonical A1 reference plus sheet_id.
- `txn_id`: UUID v7 per transaction.
- `trace_id` and `span_id`: W3C trace context compatible.
- `artifact_id`: content-addressed SHA-256 for generated artifacts.

## Workbook Core Schema
- Workbook:
  - metadata, date_system, locale, sheets[], names[], style_table, part_registry.
- Worksheet:
  - dimensions, row_props map, col_props map, cell_store sparse structure, merges, validations, conditional_formats.
- CellRecord:
  - raw_value, formula_text, parsed_ast_id, cached_value, style_ref, dependency_tags.
- PartRegistryEntry:
  - part_path, content_type, relationship graph edges, preserve_mode policy.

## Mutation Protocol
- Mutations expressed as typed operations (`SetCellValue`, `SetFormula`, `InsertRows`, `ApplyStyle`, etc.).
- Transaction holds ordered operation list and conflict policy.
- Commit returns mutation digest with changed entities and calc invalidation set.

## Serialization Rules
- Internal snapshots serialized as versioned JSON for introspection only.
- Runtime persistence remains XLSX-focused; snapshots are diagnostics artifacts, not canonical storage.

## Invariants
1. No orphan relationship entries in part registry.
2. Every formula cell must carry either parsed AST ref or parse error payload.
3. Every committed transaction yields deterministic mutation digest ordering in deterministic mode.
4. Undo stack entries reference immutable pre/post state deltas by entity ID.