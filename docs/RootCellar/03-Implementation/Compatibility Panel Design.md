# Compatibility Panel Design

Parent: [[Architecture Overview]]
Related epic: [[docs/RootCellar/01-Epics/Epic 06 - Compatibility and Migration Tooling]]

## Purpose
Make compatibility status explicit so users can trust what will round-trip, calculate, or require alternatives.

## Status Taxonomy
- Supported: feature fully understood and editable.
- Partially Supported: feature works with known limits.
- Preserved Only: feature retained in save path but not editable/fully executable.
- Not Supported: feature may degrade; user action needed.

## Data Sources
- XLSX parser feature detector.
- Calc function capability map.
- Scripting and add-in policy checks.
- Runtime warnings and fallback paths.

## UI Components
- Workbook-level summary score.
- Sheet-level issue breakdown.
- Drill-down panel with affected ranges/features.
- Suggested remediations and migration helpers.

## Artifacts
- Compatibility report JSON attached to save/run events.
- Optional export to markdown/html for governance reviews.

## Telemetry
- `compat.panel.open`
- `compat.issue.detected`
- `compat.remediation.accepted`

## Acceptance Criteria
- Every unsupported or preserved-only feature found by parser has a visible panel entry.
- Panel output and CLI report use same source schema.