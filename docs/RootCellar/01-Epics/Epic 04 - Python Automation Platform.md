# Epic 04 - Python Automation Platform

Parent: [[docs/RootCellar/00-Program/RootCellar Master Plan]]
Specs: [[docs/RootCellar/03-Implementation/Python Scripting Host Design]], [[docs/RootCellar/03-Implementation/Security and Permission Model]]

## Objective
Replace VBA workflows with secure, versioned, and observable Python automation.

## Scope
- Macro runner.
- Event hooks (`on_open`, `on_change`, `on_save`, `on_close`).
- Python UDF framework baseline.
- Permission system and script audit logs.
- Add-in packaging and signing baseline.

## Deliverables
- Script host process + RPC boundary.
- Permission prompt and policy enforcement layer.
- API docs for object model v1.
- Add-in install/verify flow.

## Stories
1. Implement script worker lifecycle and quotas.
2. Build capability token checks at RPC boundary.
3. Implement macro/event invocation wiring.
4. Add UDF runtime with cache invalidation policies.
5. Add signed package verification.

## Acceptance Criteria
- No script execution path bypasses capability checks.
- Audit logs capture all privileged script operations.
- Macro/event/UDF examples run in desktop and CLI modes.

## Dependencies
- [[Epic 01 - XLSX Fidelity and Workbook Model]]
- [[Epic 05 - Headless CLI and SDK]]
- [[Epic 07 - Radical Observability and Introspection]]
- [[Epic 08 - Enterprise Trust and Distribution]]

## Observability Requirements
- Script invocation traces and permission decision logs.
- Security anomaly events.
- Script artifact hashes and provenance capture.