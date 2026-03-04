# AI Introspection Workflow

Parent: [[Observability Charter]]
Operational runbook: [[AI Introspection Runbook]]
Execution atlas: [[Execution Traceability Atlas]]

## Purpose
Enable AI assistants and engineers to inspect product behavior using stable artifacts instead of fragile logs alone.

## Inputs For AI/Human Analysis
- Event stream filtered by trace_id/workbook_id.
- Trace graph and span timings.
- Engine mutation digests.
- Recalc DAG artifact.
- Compatibility report and transform logs.
- Script permission and audit records.

## Query Patterns
- "Why did this workbook save produce different bytes?"
- "Which formulas changed output after upgrade?"
- "Why was this macro denied network access?"
- "What caused recalc latency regression in this sprint?"

## Introspection API Contract
- Deterministic JSON schema for each artifact type.
- Stable identifiers across tools.
- Provenance metadata includes product version and commit hash.

## Guardrails
- AI suggestions are advisory; no direct privileged execution.
- Sensitive fields redacted unless policy allows.
- Artifact integrity must verify before automated diagnosis.

## Success Criteria
- Mean time to root cause reduced sprint-over-sprint.
- Debug sessions can replay from artifact bundle alone.
- AI-generated incident summaries align with human postmortems.
