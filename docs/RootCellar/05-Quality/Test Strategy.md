# Test Strategy

Parent: [[docs/RootCellar/RootCellar Planning Hub]]

## Test Pyramid
- Unit tests for parser/model/evaluator/policy components.
- Integration tests for UI-engine-script workflows.
- Corpus tests for XLSX round-trip compatibility.
- End-to-end tests for desktop and CLI critical flows.

## Priority Suites
1. XLSX no-repair suite.
2. Golden formula correctness suite.
3. Script sandbox and permission suite.
4. UI interaction parity suite.
5. Deterministic replay suite.

## CI Gates
- PR gate: fast unit + schema checks + smoke integration.
- Nightly gate: full corpus, full golden formula, benchmark, fuzz smoke.
- Release gate: security suite, deterministic cross-platform replay, artifact completeness.

## Test Data Governance
See [[Corpus Governance]].

## Observability Tie-In
All failing tests attach artifact bundles for triage.