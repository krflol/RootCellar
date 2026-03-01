# ADR-0001 Tauri + Web Grid + Rust Engine

Status: Accepted
Date: February 28, 2026
Related PRD: [[docs/RootCellar PRD]]

## Context
RootCellar needs Excel-class UX and strong engine performance with cross-platform delivery.

## Decision
Use Tauri desktop shell with React/TypeScript frontend and custom canvas/grid overlay, backed by Rust engine.

## Consequences
- Positive: fast UX development velocity and mature desktop distribution path.
- Positive: Rust engine remains core for performance and determinism.
- Tradeoff: two-language boundary requires robust contracts and tooling.

## Follow-ups
- Define UI-engine transaction schema.
- Build rendering performance benchmark harness.