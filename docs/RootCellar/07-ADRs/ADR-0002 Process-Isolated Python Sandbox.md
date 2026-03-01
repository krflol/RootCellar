# ADR-0002 Process-Isolated Python Sandbox

Status: Accepted
Date: February 28, 2026
Related PRD: [[docs/RootCellar PRD]]

## Context
Python automation is required but introduces security risk if embedded without isolation.

## Decision
Run Python in a separate worker process with capability-based RPC and OS-level sandbox controls.

## Consequences
- Positive: stronger security boundary and enterprise credibility.
- Positive: explicit policy and audit controls become enforceable.
- Tradeoff: RPC overhead and more complex debugging unless observability is strong.

## Follow-ups
- Implement policy engine and permission prompts.
- Add sandbox validation suite per OS.