# Sprint 04 - Python Macros Alpha

Parent: [[Sprint Cadence and Capacity]]
Dates: April 27, 2026 to May 10, 2026

## Sprint Goal
Ship alpha macro execution with explicit permissions and audit logging in desktop.

## Execution Status
- Status: In progress.
- Tracking links: [[docs/RootCellar/00-Program/Execution Plan Board]], [[docs/RootCellar/00-Program/Execution Status]], [[docs/RootCellar/00-Program/Sprint-Epic Story Matrix]]

## Commitments
- Epic 04 primary.
- Epic 08 policy hooks baseline.
- Epic 07 script trace correlation.

## Stories
1. Launch isolated Python worker process and basic lifecycle controls.
2. Implement RPC command set for workbook/sheet/range reads/writes.
3. Add capability prompt UX and deny-by-default enforcement.
4. Implement audit log entries for permission usage.
5. Run sandbox escape smoke tests in CI.

## Completed in Current Alpha
- [x] Story 1: CLI command and process isolation path with request/result protocol in `crates/rootcellar-cli/src/main.rs` and `crates/rootcellar-cli/src/script.rs` (`macro.run` request + JSON response + worker invocation).
- [x] Story 2: Mutation protocol supports cell and range value/formula operations through typed mutation enums in `crates/rootcellar-cli/src/script.rs` and Python worker (`python/worker_stub.py`) with deterministic serialization.
- [x] Story 4: Permission decision events emitted as part of run-macro telemetry (`script.permission.granted`, `script.permission.denied`) and RPC failure tracing (`script.rpc.error`) in `crates/rootcellar-cli/src/main.rs`.
- [x] Story 5 (in-progression): Core execution path and test scaffolding is present; full sandbox-escape CI hardening remains for broader policy surfaces.

## Acceptance Criteria
- Macro can run on open workbook and mutate ranges via transaction replay path.
- Permission signals can deny reads/writes unless explicitly granted.
- Macro runs emit trace and audit artifacts for permission decisions and lifecycle events.

## Exit Signals
- Security review approves alpha constraints for limited internal use.
