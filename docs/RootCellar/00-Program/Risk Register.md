# Risk Register

Parent: [[RootCellar Master Plan]]

## Scoring
- Probability: Low/Medium/High.
- Impact: Medium/High/Critical.

## Risks
| ID | Risk | Probability | Impact | Owner | Mitigation | Trigger Signal |
|---|---|---|---|---|---|---|
| R-01 | XLSX edge-case corruption in save path | Medium | Critical | Interop Lead | Opaque passthrough, corpus diff tests, binary diff checks | Repair prompt rate > 5% in nightly corpus |
| R-02 | Formula parity gaps erode trust | High | High | Calc Lead | Targeted function family roadmap, explicit compatibility panel statuses | High-frequency unsupported function telemetry |
| R-03 | Python sandbox bypass | Low | Critical | Security Lead | Process isolation, OS sandbox, capability RPC checks, red-team tests | Any policy bypass in fuzz or pen test |
| R-04 | Grid performance degrades on large sheets | Medium | High | UI Lead | Canvas virtualization budgets, perf CI thresholds | p95 scroll frame > 16ms sustained |
| R-05 | Deterministic mode unstable across OS/CPU | Medium | High | Engine Lead | Canonical ordering, numeric policy docs, repro checks | Hash mismatch in cross-platform golden runs |
| R-06 | Team overloaded by broad PRD scope | High | High | PM | Phase gates, strict slice sequencing, backlog curation | Carryover > 35% for 2 sprints |
| R-07 | Accessibility lags until late | Medium | High | UX Lead | Semantic mirror early, keyboard parity checks each sprint | A11y smoke failures or blocked workflows |

## Escalation Policy
- Critical impact risks escalate same day to program review.
- Any security risk with plausible exploit path blocks release branches.

## Linked Plans
- Quality gates: [[docs/RootCellar/05-Quality/Release Gates]]
- Incident process: [[docs/RootCellar/06-Operations/Incident Response Playbook]]
- Security validation: [[docs/RootCellar/05-Quality/Security Validation]]