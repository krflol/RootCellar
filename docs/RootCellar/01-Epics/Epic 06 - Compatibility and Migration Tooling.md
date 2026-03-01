# Epic 06 - Compatibility and Migration Tooling

Parent: [[docs/RootCellar/00-Program/RootCellar Master Plan]]
Specs: [[docs/RootCellar/03-Implementation/Compatibility Panel Design]]

## Objective
Provide transparent compatibility reporting and practical migration workflows from VBA-heavy workbooks.

## Scope
- Compatibility panel in desktop.
- Compatibility report output in CLI.
- VBA scanner and classifier.
- Migration assistant skeleton generator.

## Deliverables
- Unified compatibility schema consumed by UI and CLI.
- VBA complexity classification rules.
- Python stub generation for common macro patterns.

## Stories
1. Implement feature detector registry tied to XLSX parser and calc capability map.
2. Build panel UI and issue drill-down.
3. Build VBA scanner and module classifier.
4. Generate migration stubs and actionable notes.

## Acceptance Criteria
- Compatibility findings are consistent across UI and CLI.
- Migration report covers module inventory, complexity, and suggested approach.
- Panel includes remediation guidance for all unsupported findings.

## Dependencies
- [[Epic 01 - XLSX Fidelity and Workbook Model]]
- [[Epic 02 - Calculation Engine and Determinism]]
- [[Epic 04 - Python Automation Platform]]

## Observability Requirements
- Feature-gap telemetry by workbook segment.
- Migration assistant adoption and completion metrics.