# Decision Register

Parent: [[RootCellar Master Plan]]

## Active ADRs
- [[docs/RootCellar/07-ADRs/ADR-0001 Tauri + Web Grid + Rust Engine]]
- [[docs/RootCellar/07-ADRs/ADR-0002 Process-Isolated Python Sandbox]]
- [[docs/RootCellar/07-ADRs/ADR-0003 Deterministic Mode and Repro Records]]

## Pending Decisions
| ID | Decision | Needed By | Options | Decision Driver |
|---|---|---|---|---|
| D-01 | Grid rendering backend baseline | Sprint 02 | Canvas 2D vs WebGL-first | Perf on million-cell workloads |
| D-02 | Formula function rollout order | Sprint 02 | By usage frequency vs by family completeness | User value and parity optics |
| D-03 | Add-in package format details | Sprint 04 | Zip manifest v1 vs OCI-like bundle | Signing and enterprise distribution |
| D-04 | Telemetry storage backend for local inspect mode | Sprint 05 | Embedded DB vs file-based JSONL + index | Simplicity and query speed |
| D-05 | Collaboration protocol foundation | Phase C planning | OT vs CRDT | Determinism + merge UX |

## Decision SLA
- Architectural and security decisions: <= 5 business days.
- UX interaction decisions: <= 3 business days.
- If SLA breached, default to low-regret implementation with ADR note.