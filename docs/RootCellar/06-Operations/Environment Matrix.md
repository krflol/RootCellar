# Environment Matrix

Parent: [[docs/RootCellar/RootCellar Planning Hub]]

## Environments
- Local dev: fast iteration, optional reduced telemetry retention.
- CI ephemeral: deterministic builds and test artifacts.
- Staging: production-like configs and full observability checks.
- Production: policy-controlled telemetry, enterprise trust defaults.

## Configuration Controls
- Mode flags: preserve/normalize/deterministic.
- Telemetry level: minimal/diagnostic/forensic.
- Script trust mode: personal/team/enterprise.
- Sandbox profile per OS.

## Secrets and Keys
- Signing keys managed outside repo in secure key vault.
- Environment-specific credentials rotated on schedule.

## Drift Detection
- Config snapshots emitted as artifacts each deploy.
- Drift alert when staging/prod diverge from approved baseline.