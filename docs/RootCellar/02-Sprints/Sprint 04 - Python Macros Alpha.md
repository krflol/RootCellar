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
- [x] Story 3: Desktop command integration for macro execution (`interop_run_macro` in `apps/desktop/src-tauri/src/main.rs`) with permission telemetry mapping and transaction-backed mutation replay.
- [x] Story 4: Permission decision and macro lifecycle telemetry is emitted for both CLI and desktop execution paths, with permission event + artifact correlation in backend traces.
- [x] Story 5: Desktop macro permission policy persistence and trust prompt UX are now implemented per script path with explicit consent for elevated permission profiles.
- [ ] Story 6 (in-progression): Signed package verification and policy provenance artifacts for macro runtime trust are now in progress (beyond this initial alpha slice).

## Acceptance Criteria
- Macro can run on open workbook and mutate ranges via transaction replay path on both CLI and desktop.
- Permission signals can deny reads/writes unless explicitly granted, including deny cases that do not mutate workbook state.
- Macro runs emit trace and audit artifacts for permission decisions and lifecycle events from both execution surfaces.
- Desktop UX can persist and reapply per-script permission policy with explicit trust prompts when policy is newly approved.

## Exit Signals
- Security review approves alpha constraints for limited internal use.
