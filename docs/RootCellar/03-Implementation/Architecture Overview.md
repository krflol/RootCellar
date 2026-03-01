# Architecture Overview

Parent: [[docs/RootCellar/RootCellar Planning Hub]]
PRD Source: [[docs/RootCellar PRD]]

## System Topology
- Desktop shell: Tauri host process.
- Frontend: React + TypeScript UI with virtualized grid renderer.
- Engine: Rust core crate set for workbook model, calc, import/export, transaction manager.
- Script host: separate Python worker process with RPC boundary and capability enforcement.
- Headless surface: Rust CLI and SDK wrappers over shared engine crates.

## Data/Control Flow
1. User action or CLI command creates a transaction request.
2. Engine transaction manager validates and applies workbook mutations.
3. Calc graph invalidates and recomputes impacted nodes.
4. UI updates focused viewport and formula/editor overlays from engine state snapshots.
5. Event bus emits structured telemetry and introspection artifacts with shared trace IDs.

## Boundary Rules
- UI never mutates workbook state directly; it issues transaction commands.
- Script host never receives raw memory references, only RPC DTOs with capabilities.
- XLSX passthrough layer preserves unsupported parts as opaque units with provenance metadata.
- Deterministic mode toggles stable ordering and reproducibility policies in save/eval pipelines.

## Core Entities
- Workbook, Worksheet, Table, Range, Cell, StyleRef, Name, Pivot, ChartPart, ArtifactBundle.
- IDs are stable within workbook lifecycle; see [[Data Model and IDs]].

## Non-Functional Control Points
- Perf budgets defined in [[docs/RootCellar/05-Quality/Performance and Benchmarking]].
- Security controls defined in [[Security and Permission Model]].
- Observability controls defined in [[docs/RootCellar/04-Observability/Observability Charter]].

## Build Order
1. Workbook model + transaction API.
2. XLSX import/export preserve pipeline.
3. Calc parser + dependency graph + evaluator baseline.
4. Grid editing loop and undo/redo.
5. Script host + permission gate.
6. CLI and reproducibility workflows.