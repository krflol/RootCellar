# Audit and Forensics

Parent: [[Observability Charter]]
Related: [[docs/RootCellar/03-Implementation/Security and Permission Model]]

## Audit Log Scope
Audit logs cover all security-sensitive actions:
- Script permission requests and outcomes.
- Add-in install/upgrade/remove events.
- Policy mode changes.
- Trusted location updates.
- Macro and add-in executions with hashes.

## Immutable Log Requirements
- Append-only storage semantics.
- Cryptographic hash chain across records.
- Tamper-evident validation command in CLI.

## Forensic Workflow
1. Collect trace_id or workbook_id from incident report.
2. Pull related audit records and artifact bundle.
3. Verify hash chain and checksums.
4. Reconstruct timeline of actions and policy decisions.
5. Produce incident summary with root cause and containment status.

## PII and Compliance
- Audit payloads avoid direct cell content unless forensic mode explicitly enabled.
- Redaction policy applied before external sharing.
- Enterprise mode defaults to minimal sensitive payload capture.

## Tooling
- `rootcellar audit verify <log_path>`
- `rootcellar artifact inspect <bundle_id>`
- `rootcellar trace show <trace_id>`