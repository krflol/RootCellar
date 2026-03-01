# RootCellar Master Plan

Parent: [[docs/RootCellar/RootCellar Planning Hub]]
Source: [[docs/RootCellar PRD]]

## Program Objective
Deliver RootCellar as a credible Excel replacement for common business workbooks, with secure Python automation and a headless engine surface for reproducible workflows.

## Outcome Targets
- Compatibility: >= 92% of target corpus round-trips without Excel repair prompt by Beta.
- Correctness: >= 99.5% pass rate on golden formula suite for in-scope functions.
- Performance: p95 recalc latency <= 350 ms for benchmark medium models; linear-ish scale in batch recalc with Rayon.
- Security: zero known sandbox escape paths in release candidate.
- Operability: 100% of top-level user actions traceable across UI -> engine -> scripting -> artifact outputs.

## Workstream Model
- Workstream A: Workbook model + XLSX import/export fidelity.
- Workstream B: Calculation engine + determinism.
- Workstream C: Desktop UX + productivity workflows.
- Workstream D: Python scripting + sandbox + add-ins.
- Workstream E: CLI/SDK + batch automation.
- Workstream F: Compatibility reporting + migration tooling.
- Workstream G: Observability and introspection plane.
- Workstream H: Enterprise trust (signing, policies, managed distribution).

## Governance Rhythm
- Weekly architecture review: API boundaries, determinism tradeoffs, compatibility exceptions.
- Weekly security review: permission model, sandbox, threat regressions.
- Bi-weekly sprint review: epic progress and corpus pass delta.
- Monthly readiness gate: release criteria from [[docs/RootCellar/05-Quality/Release Gates]].

## Definition Of Done (Program Level)
A slice is done only if all are true:
1. Functional acceptance criteria are met.
2. Regression tests and benchmarks pass.
3. Observability contract events/artifacts are emitted and documented.
4. Security posture for the slice is documented and tested.
5. User-facing compatibility implications are captured in the compatibility panel rules.

## Cross-Cutting Artifacts
- Architecture: [[docs/RootCellar/03-Implementation/Architecture Overview]]
- Event and trace model: [[docs/RootCellar/04-Observability/Telemetry Taxonomy and Event Schema]], [[docs/RootCellar/04-Observability/Trace Correlation Model]]
- Artifact contract: [[docs/RootCellar/04-Observability/Inspectable Artifact Contract]]
- Quality and release: [[docs/RootCellar/05-Quality/Test Strategy]], [[docs/RootCellar/05-Quality/Release Gates]]

## Execution Linkage
- Plan-to-delivery board: [[Execution Plan Board]]
- Evidence ledger: [[Execution Status]]
- Current snapshot: Sprint 00 foundation complete, Sprint 01 interop baseline complete, Sprint 02 parser/dependency graph in progress.

## Open Program Risks
Tracked in [[Dependency Map]] and [[Risk Register]].
