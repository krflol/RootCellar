# Sprint 04 - Python Macros Alpha

Parent: [[Sprint Cadence and Capacity]]
Dates: April 27, 2026 to May 10, 2026

## Sprint Goal
Ship alpha macro execution with explicit permissions and audit logging in desktop.

## Execution Status
- Status: Planned.
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

## Acceptance Criteria
- Macro can run on open workbook and mutate ranges via transactions.
- File/network access denied unless explicitly granted.
- Every macro run emits trace and audit artifacts.

## Exit Signals
- Security review approves alpha constraints for limited internal use.
