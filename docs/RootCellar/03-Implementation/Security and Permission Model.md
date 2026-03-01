# Security and Permission Model

Parent: [[Architecture Overview]]
Related epic: [[docs/RootCellar/01-Epics/Epic 04 - Python Automation Platform]] and [[docs/RootCellar/01-Epics/Epic 08 - Enterprise Trust and Distribution]]

## Security Principles
1. Deny by default for script capabilities.
2. Explicit user/admin consent for privileged operations.
3. Process and OS isolation boundaries are mandatory.
4. All privileged actions are auditable with immutable logs.

## Capability Set v1
- `fs.read` and `fs.write` scoped to approved paths.
- `net.http` scoped to approved domains and methods.
- `clipboard` for explicit user actions.
- `process.exec` behind admin policy only.

## Enforcement Layers
- Policy engine validates manifest-declared and runtime-requested permissions.
- RPC gateway checks capability token per call.
- Worker sandbox restricts system resources and process behavior.
- Add-in signature trust chain verifies publisher identity.

## Trust Modes
- Personal: interactive prompts allowed, unsigned add-ins optional.
- Team: allowlist policies, warnings for unsigned packages.
- Enterprise: signed-only, centrally managed policies, trusted locations enforced.

## Audit Log Requirements
Each privileged action logs:
- actor identity (user/service)
- workbook and script identifiers
- permission used
- timestamp and host context
- trace and transaction correlation IDs
- result and failure details

## Threat Model Hotspots
- Sandbox escapes via native extensions.
- Confused deputy via UI automation APIs.
- Malicious workbook-embedded add-ins.
- Data exfiltration through network permissions.

## Validation Plan
See [[docs/RootCellar/05-Quality/Security Validation]].