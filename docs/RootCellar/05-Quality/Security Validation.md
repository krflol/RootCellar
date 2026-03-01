# Security Validation

Parent: [[Test Strategy]]

## Scope
- Capability enforcement correctness.
- Sandbox hardening and escape testing.
- Add-in signature and trust chain validation.
- Audit log tamper-evidence checks.

## Test Layers
- Unit: permission evaluator and policy parser.
- Integration: script RPC with allowed/denied permutations.
- Dynamic: fuzzing and hostile workbook/add-in inputs.
- Red team: periodic manual adversarial testing.

## Release Criteria
- Zero open critical security findings.
- All high findings have documented mitigation and accepted risk signoff.
- Signed-only mode verified across desktop and CLI.

## Evidence
Security test runs must include artifact bundles and verified audit logs.