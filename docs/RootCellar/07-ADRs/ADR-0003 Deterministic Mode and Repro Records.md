# ADR-0003 Deterministic Mode and Repro Records

Status: Accepted
Date: February 28, 2026
Related PRD: [[docs/RootCellar PRD]]

## Context
Users need reproducible automation and CI-safe outputs while maintaining compatibility workflows.

## Decision
Ship dual modes: preserve (compatibility-first) and deterministic normalize workflows with repro record/check tooling.

## Consequences
- Positive: clear operational mode for CI and forensic comparisons.
- Positive: deterministic behavior becomes measurable and enforceable.
- Tradeoff: additional complexity in save/eval pipelines and user education.

## Follow-ups
- Implement canonical ordering and reproducibility artifact schema.
- Add mismatch detection in nightly CI.