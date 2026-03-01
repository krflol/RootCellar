# Sprint 00 - Foundation and Telemetry Bootstrap

Parent: [[Sprint Cadence and Capacity]]
Dates: March 2, 2026 to March 15, 2026

## Sprint Goal
Establish skeleton architecture and mandatory observability plumbing so subsequent feature work is traceable.

## Execution Status
- Status: Completed.
- Tracking links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]]
- Delivered stories:
1. Rust workspace scaffold and crate boundaries.
2. Transaction API baseline with IDs and commit digest shape.
3. Event envelope + JSONL sink in Rust core and CLI integration.
- Deferred from this sprint plan:
1. UI boot flow placeholder and viewport prototype (moved forward; still pending).

## Commitments
- Epic 01: workbook model scaffolding and part registry interfaces.
- Epic 03: desktop shell startup and grid viewport prototype.
- Epic 07: telemetry SDK, event schema v0, trace context propagation base.
- Epic 05: CLI skeleton with command routing and logging.

## Stories
1. Create Rust workspace crate boundaries and shared DTO package.
2. Implement transaction API stubs with IDs and commit digest shape.
3. Implement UI boot flow and viewport render placeholder.
4. Build event emitter SDK for Rust and TypeScript.
5. Add CI checks for event schema linting.

## Definition Of Done
- Trace IDs visible in UI action logs and engine transaction logs.
- Artifact bundle writer can emit minimal session manifest.
- CI pipeline runs unit tests and publishes build metadata.

## Risks
- Overbuilding framework before vertical slice delivery.
- Mitigation: enforce sprint demo with one end-to-end edit trace.
