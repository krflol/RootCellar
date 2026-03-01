# Python Scripting Host Design

Parent: [[Architecture Overview]]
Related epic: [[docs/RootCellar/01-Epics/Epic 04 - Python Automation Platform]]

## Process Model
- Desktop app launches script worker process on demand.
- Worker process hosts CPython runtime with constrained boot environment.
- Engine and worker communicate via RPC over local transport.

## Script Types
- Macros: user-invoked procedures.
- Events: lifecycle and worksheet mutation hooks.
- UDFs: callable from cells with cache and invalidation integration.

## RPC Boundary
- Typed request/response schemas only.
- Capability token attached to each request.
- Strict timeout and resource quotas per invocation.

## API Surface v1
- `rc.workbook`, `rc.sheet(name)`, `range.values`, `range.formulas`, limited `range.format`.
- Controlled `rc.io` helpers.
- Limited `rc.ui` notifications/dialogs.

## UDF Runtime Controls
- Determinism flag for allowed function categories.
- Side-effect prohibition in UDF context by policy.
- Caching keyed by input hash + script version + permission context.

## Add-in Packaging
- Manifest with `name`, `version`, `publisher`, `permissions`, `api_min_version`, `signature`.
- Package archive with modules/resources and optional UI contributions.
- Install workflow validates signature and policy before activation.

## Telemetry
- `script.session.start|end`
- `script.macro.run`
- `script.udf.invoke`
- `script.permission.request|granted|denied`
- `script.rpc.error`