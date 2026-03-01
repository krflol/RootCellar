# Telemetry Taxonomy and Event Schema

Parent: [[Observability Charter]]

## Event Envelope (v1)
```json
{
  "event_name": "calc.recalc.end",
  "event_version": "1.0.0",
  "timestamp": "2026-03-30T15:11:05.123Z",
  "severity": "info",
  "trace_id": "...",
  "span_id": "...",
  "session_id": "...",
  "workbook_id": "...",
  "txn_id": "...",
  "actor": { "type": "user", "id": "..." },
  "context": { "surface": "desktop", "mode": "preserve" },
  "metrics": { "duration_ms": 42.7 },
  "payload": {}
}
```

## Event Domains
- `ui.*` UI interactions and render performance.
- `engine.*` model transactions and invariant checks.
- `calc.*` parser/evaluator/recalc behavior.
- `interop.*` file import/export and compatibility transforms.
- `script.*` automation runtime and permission events.
- `policy.*` trust and authorization decisions.
- `cli.*` headless command lifecycle.
- `artifact.*` bundle creation, validation, retention.

## Required Fields By Domain
- `ui.*`: `command_id`, `interaction_latency_ms`, `viewport`.
- `engine.*`: `mutation_count`, `entity_counts`, `commit_status`.
- `calc.*`: `invalidated_nodes`, `evaluated_nodes`, `cycle_count`.
- `interop.*`: `part_count`, `unknown_part_count`, `rewrite_count`.
- `script.*`: `script_hash`, `permission_set`, `sandbox_profile`.
- `policy.*`: `policy_mode`, `decision`, `decision_reason_codes`.

## Error Taxonomy
- `E_PARSE_*` parse/format failures.
- `E_CALC_*` formula and recalc failures.
- `E_SCRIPT_*` scripting execution and policy failures.
- `E_POLICY_*` trust/policy failures.
- `E_IO_*` filesystem/network/export failures.

## Schema Lifecycle
- Backward compatible additions allowed in minor versions.
- Removals/renames require major version and migration guide.
- CI validates example fixtures and producers against schema.