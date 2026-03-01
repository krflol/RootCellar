# Dependency Map

Parent: [[RootCellar Master Plan]]

## Critical Path Dependencies
1. Workbook model stable IDs and mutation API -> blocks calc graph, scripting API, undo/redo, observability links.
2. XLSX passthrough import/export -> blocks compatibility confidence and corpus progression.
3. Dependency graph correctness -> blocks incremental recalc performance and UDF invalidation semantics.
4. Grid virtualization + edit lifecycle -> blocks usable desktop alpha.
5. Permission gate + process isolation -> blocks any production Python automation release.
6. Trace correlation IDs everywhere -> blocks radical observability objectives.

## Dependency Matrix
| Upstream | Downstream | Risk if delayed | Mitigation |
|---|---|---|---|
| Workbook entity IDs | Audit logs, traces, undo history | Untraceable edits | Freeze ID format in Sprint 00 |
| XLSX parser part graph | Preserve mode save | Data loss on save | Preserve unknown parts as opaque blobs |
| Formula parser AST | Calc evaluator, formula editor | Divergent behavior UI vs engine | Single AST shared crate |
| Permission policy engine | Script runner, add-in manager | Unsafe macro execution | Ship deny-all default; policy tests |
| Structured event SDK | Dashboards, alerts, forensic bundle | Blind operations | Add events as DoD requirement |

## External Dependencies
- Signing infrastructure (code/add-ins).
- OS sandbox APIs per platform.
- Corpus acquisition and legal review for sample workbook usage.

## Active Dependency Risks
- UI edit semantics may drift from engine transaction semantics.
- Dynamic array behavior may force parser/evaluator refactor.
- Cross-platform sandbox parity can lag Windows-first implementation.

See also [[Risk Register]].