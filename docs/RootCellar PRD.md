

**Product:** RootCellar  
**Type:** Cross-platform spreadsheet application + headless engine  
**Core stack:** Rust (engine) + Python (scripting) + Tauri (desktop shell) + Web UI (grid)

## 1) Vision

RootCellar is a modern, fast, secure spreadsheet platform that is a **practical drop-in replacement for Microsoft Excel** for most real-world business workbooks, while replacing VBA with **first-class Python automation** (macros, events, UDFs, add-ins), and leveraging Rust + Rayon for **performance, determinism, and scalable batch workflows**.

### What “drop-in replacement” means (explicit product contract)

RootCellar prioritizes, in order:

1. **XLSX open/edit/save without breaking files** (round-trip safety)
    
2. **Excel-like user experience** for common workflows
    
3. **Formula correctness** for the most used functions + predictable differences surfaced clearly
    
4. **Scriptability** that exceeds VBA in safety, portability, and developer experience
    

---

## 2) Goals and Non-Goals

### Goals

- Open and save `.xlsx` with high fidelity; avoid “Excel repaired this workbook” scenarios.
    
- Excel-familiar UX: grid, formula bar, named ranges, tables, filters, charts, pivots (full, eventually).
    
- Fast recalc (incremental, multi-threaded).
    
- Python scripting integrated as:
    
    - Macros (imperative automation)
        
    - Events (on_open, on_change, etc.)
        
    - Python UDFs callable from cells
        
    - Add-ins with packaging, signing, permissions, distribution
        
- Secure by default: scripts are sandboxed and permission-gated with auditing.
    
- “Headless RootCellar”: CLI + library APIs for automation and server/batch use.
    

### Non-Goals (even long-term)

- Running VBA as-is inside RootCellar (we provide migration tooling, not execution parity).
    
- Mimicking every undocumented Excel quirk; instead we provide compatibility modes, reports, and deterministic semantics.
    

---

## 3) Target Users

1. **Finance/Ops Analysts:** heavy template use; need XLSX fidelity and familiar UX; rely on macros.
    
2. **Power Users / Model Builders:** complex formulas, charts, pivots, named ranges; performance sensitive.
    
3. **Engineering/Data Teams:** want scripting and reproducible pipelines; prefer Python; want CI automation.
    
4. **IT/Security Admins:** want control over scripts; signing; policy enforcement; audit logs.
    

---

## 4) Product Surfaces

RootCellar is one product with multiple first-class surfaces:

1. **Desktop App (primary)**
    
2. **Headless CLI (first-class, not an afterthought)**
    
3. **Embeddable Engine SDK (Rust, plus Python bindings)**
    
4. **Optional later:** Web client (WASM engine or remote engine)
    

---

## 5) UI / Frontend Technology Choice

### Recommendation: **Tauri (desktop shell) + Web UI grid**

**Why Tauri is the right default**

- Tight Rust integration without Electron weight
    
- Mature cross-platform packaging/signing story
    
- Lets you build an Excel-class UI fastest by leveraging web rendering + accessibility primitives
    
- Keeps the core engine in Rust where it belongs
    

**UI implementation detail that matters**  
Spreadsheets need **high-performance grid virtualization**. In practice, you will implement the grid as:

- a **canvas/WebGL layer** (for cell rendering + scrolling performance)
    
- plus a **DOM overlay** (for editors, menus, formula bar, tooltips, accessibility hooks)
    

**Recommended UI stack**

- **Tauri + React + TypeScript**
    
- Grid: custom canvas-based grid (recommended) or a permissive licensed component if acceptable
    
- Command palette, ribbon-like toolbar, panels, dialogs in standard web UI
    

**Why not pure Rust UI (egui/iced)**  
You _can_, but Excel-like UX, IME/text, clipboard, accessibility, and deep input quirks are significantly harder. Web tech wins on “polish per engineering hour.”

---

## 6) Core Functional Requirements

## 6.1 Workbook Model

RootCellar’s internal workbook model must represent:

- Workbook metadata, sheets, names, defined tables
    
- Rows/columns with widths/heights, hidden state, grouping/outlines
    
- Cells with:
    
    - value types: number, string, bool, error, empty
        
    - formula string + parsed AST + cached result
        
    - style reference
        
    - rich text (eventual), comments/notes (eventual), hyperlinks
        
- Merged cells, conditional formats, data validation rules
    
- Embedded objects: charts, images, shapes (view + preserve early; edit later)
    

**Requirement:** Model supports **partial fidelity**: even if the engine doesn’t _understand_ every feature, it must preserve it during save.

---

## 6.2 XLSX Compatibility and Round-Trip Safety

### Required capabilities

- Read `.xlsx` with:
    
    - workbook structure
        
    - worksheet XML
        
    - sharedStrings
        
    - styles (enough to preserve + map)
        
    - calcChain handling (don’t rely on it; preserve when needed)
        
    - relationships and embedded parts
        
- Write `.xlsx`:
    
    - preserve unknown XML parts where feasible (“passthrough strategy”)
        
    - stable ordering for deterministic output modes
        
    - no Excel “repair” prompts on target corpus
        

### Compatibility modes

- **Preserve Mode (default):** prefer passthrough and minimal mutation of unknown parts.
    
- **Normalize Mode:** rewrite structure into a canonical form for determinism/cleanliness (may drop some unknown artifacts, but is reproducible and CI-friendly).
    

### Compatibility reporting

- Built-in “Compatibility Panel” in UI:
    
    - Supported / Partially Supported / Preserved-only / Not Supported
        
    - actionable notes and suggested alternatives
        

---

## 6.3 Calculation Engine

### Required

- Parser -> AST with Excel-like grammar
    
- Dependency graph with:
    
    - incremental recalc
        
    - cycle detection and cycle reporting
        
- Function library:
    
    - broad coverage with explicit target parity categories (math, stats, financial, text, date/time, lookup/reference, logical)
        
- Error model identical to Excel where documented
    
- Locale handling for:
    
    - decimal/list separators
        
    - date systems (1900/1904)
        
- Named ranges, structured references (tables), cross-sheet references
    
- Array behavior:
    
    - support classic CSE arrays
        
    - support dynamic arrays (eventual full parity)
        

### Performance + Rayon requirements

- Rayon must be used for:
    
    - parallel evaluation of independent dependency subgraphs
        
    - parallel range operations (large SUM/COUNT/aggregation)
        
    - parallel import/export tasks where safe
        
- Must not compromise determinism:
    
    - define stable tie-breaking and ordering rules
        
    - avoid floating-point nondeterminism where possible (document if not)
        

---

## 6.4 Grid UX (Excel Familiarity)

### Required (desktop)

- Cell selection model identical “enough” to Excel:
    
    - multi-range selection
        
    - keyboard navigation
        
    - shift/ctrl behaviors
        
- Fill handle and autofill series
    
- Copy/paste:
    
    - within RootCellar
        
    - between RootCellar and Excel (values + basic formatting)
        
- Undo/redo across edits and structure changes
    
- Find/replace; go-to; name box
    
- Freeze panes, split view
    
- Sort/filter (range and tables)
    
- Format painter; basic style tools; number formats
    
- Basic charts: view and edit common chart types (eventual full)
    
- Pivot tables: view/refresh existing pivots early, full pivot editor later (but PRD includes full scope)
    

### UX principle

Be Excel-familiar where it reduces friction, but improve where safe:

- command palette
    
- safe scripting
    
- explicit compatibility panel
    
- reproducible/exportable build artifacts
    

---

## 6.5 Python Scripting System (VBA replacement)

### 6.5.1 Script Types

1. **Macros** (imperative)
    
    - run from UI: ribbon/menu, hotkey, button
        
    - run from CLI/headless
        
2. **Events**
    
    - `on_open(workbook)`
        
    - `on_change(sheet, changed_range)`
        
    - `on_save(workbook)`
        
    - `on_close(workbook)`
        
    - optional later: timers, background tasks with safe rules
        
3. **Python UDFs**
    
    - register functions as worksheet functions
        
    - accept scalars and ranges
        
    - return scalars or 2D arrays
        
    - caching and invalidation rules defined
        

### 6.5.2 Scripting API (RootCellar Object Model)

Expose a stable, versioned API designed for productivity and safety:

- `rc.app` (limited UI interactions)
    
- `rc.workbook` (sheets, names, events, metadata)
    
- `rc.sheet(name)` -> sheet
    
- `sheet.range("A1:D10")`
    
- `range.values`, `range.formulas`
    
- `range.format` (subset)
    
- `sheet.tables`, `sheet.filters`, `sheet.pivots` (when implemented)
    
- `rc.io` helpers (import/export CSV/JSON)
    
- `rc.ui` (dialogs/notifications) with constraints
    

**Requirement:** API is documented and versioned; add-ins declare minimum API version.

### 6.5.3 Sandbox & Permissions (non-negotiable)

Python is powerful; “just embed CPython” is not a security model.

**RootCellar must implement a capability-based system:**

- Scripts run in a restricted environment with explicit permissions:
    
    - `fs.read`, `fs.write` (scoped paths)
        
    - `net.http` (scoped domains)
        
    - `clipboard` (if needed)
        
    - `process.exec` (admin-only)
        
- Default: no file/network/exec.
    
- UI prompts for user-granted permissions (or admin policy pre-approval).
    
- Full audit log: who ran what, what permissions were used, and hashes of scripts.
    

**Isolation approach (recommended)**

- Run Python scripts in a **separate process** with RPC to the Rust engine.
    
- Apply OS-level sandboxing:
    
    - Windows: AppContainer / Job objects + constrained token
        
    - macOS: sandbox profiles
        
    - Linux: seccomp + namespaces where available  
        This is what makes “Python macros” something IT will allow.
        

### 6.5.4 Add-ins (“Casks” if you want theme, but optional)

Add-ins are packaged units:

- manifest (name, version, publisher, permissions, API min version)
    
- python modules + resources
    
- optional UI contributions (menu items, buttons)
    
- optional compiled extensions (policy controlled)
    

**Signing**

- RootCellar supports signed add-ins and optional “signed only” policy.
    
- UI shows “Verified Publisher” and trust chain.
    

**Distribution**

- Local install (per-user)
    
- Org-managed repo (later)
    
- Workbook-embedded add-ins (policy-controlled; off by default in enterprise)
    

### 6.5.5 Script Migration Tooling

RootCellar includes:

- workbook scanner: detect VBA presence, enumerate modules, classify complexity
    
- migration assistant:
    
    - produce Python skeleton stubs
        
    - map common patterns (sheet loops, range ops, exports)
        
- compatibility report surface
    

---

## 6.6 Import / Export / Interop

### Required formats

- `.xlsx` (primary)
    
- `.csv` (robust import/export wizard)
    
- `.tsv`, `.txt` delimited variants
    
- optional later: `.ods`, `.xlsb` view/import
    

### Exports

- export to `.xlsx` and `.csv`
    
- export to PDF (eventual, full control)
    
- export snapshots/images for reporting
    

### Headless / batch

- CLI must support:
    
    - run macro/add-in on a workbook
        
    - patch inputs, recalc, export
        
    - run across directories with Rayon parallelism
        
    - emit JSONL report artifacts
        

---

## 6.7 Collaboration (Full scope in PRD)

RootCellar should eventually support:

- Shared workbook sessions
    
- Conflict handling
    
- Presence + comments + history
    
- “Enterprise-friendly” on-prem option
    

**Engineering note:** plan for CRDT/OT architecture but don’t contaminate v1 engine design; keep the workbook mutation API clean and serializable.

---

## 7) Non-Functional Requirements

### Performance

- Smooth scroll and selection on large sheets (virtualized rendering).
    
- Incremental recalc: only recompute dependents.
    
- Rayon threadpool controls for server/batch contexts.
    
- Memory: handle large shared string tables and sparse cell storage efficiently.
    

### Determinism (optional but core differentiator)

- Deterministic mode:
    
    - stable output bytes for saved workbooks where feasible
        
    - stable calc results and ordering
        
- Normal mode:
    
    - prioritize compatibility and performance
        
- Provide “repro record/check” workflow in CLI for CI.
    

### Reliability

- Autosave + crash recovery
    
- Corrupt workbook handling: best-effort open with clear diagnostics
    
- Extensive corpus-based regression tests
    

### Accessibility

- Keyboard-first operation
    
- Screen reader support
    
- High contrast / zoom / font scaling
    
- Grid accessibility strategy:
    
    - DOM “semantic mirror” for focused region + navigation announcements
        
    - not every cell must be in the DOM at once
        

### Security

- Macro sandbox and permission system (above)
    
- Signed add-ins and optional “signed-only”
    
- Workspace-level trust: “trusted locations” concept
    
- Telemetry opt-in and scrubbed; enterprise off by default
    

### Internationalization

- Locale-aware separators and date formats
    
- Right-to-left support eventually
    

---

## 8) High-Level Architecture

### 8.1 Core components

1. **Rust Engine**
    
    - workbook model
        
    - calc engine
        
    - import/export pipeline
        
    - mutation API + eventing
        
2. **UI Frontend**
    
    - grid renderer (canvas/webgl + DOM overlay)
        
    - formula editor
        
    - ribbon/toolbar, panels
        
3. **Scripting Host**
    
    - Python runner (separate process)
        
    - permission gate + sandbox enforcement
        
    - UDF bridge
        
4. **Interop Layer**
    
    - XLSX part preservation strategy
        
    - compatibility diagnostics
        
5. **Headless Surface**
    
    - CLI
        
    - programmatic SDK
        

### 8.2 Key APIs

- Engine exposes a “transactional edit” API:
    
    - begin transaction
        
    - apply mutations
        
    - commit -> triggers dependency updates + events
        
- Scripting talks to engine over a constrained RPC protocol:
    
    - no raw pointers, no arbitrary eval access
        
    - capability checks at RPC boundary
        

---

## 9) Technology Stack (Specified)

- **Desktop shell:** Tauri
    
- **Frontend:** React + TypeScript
    
- **Grid:** custom canvas/WebGL virtualization + DOM overlay editors
    
- **Engine:** Rust
    
- **Parallelism:** Rayon
    
- **Scripting:** CPython in a sandboxed worker process (PyO3 optional only for bridge tooling; not as the security boundary)
    
- **File I/O:** Rust OpenXML `.xlsx` reader/writer with passthrough preservation strategy
    
- **CLI:** Rust (same engine), structured JSON output
    
- **Packaging:** platform-native installers + code signing; add-in signing support
    

---

## 10) Testing & Validation Strategy (must be part of the product)

- **Corpus-driven XLSX round-trip tests**
    
    - curated “real world” templates (finance, ops, HR)
        
    - ensure “no repair prompt”
        
- **Golden calc correctness**
    
    - compare outputs vs known results
        
    - allow compatibility mode diffs with explicit explanations
        
- **Fuzzing**
    
    - zip/xml parsing, formula parsing
        
- **Performance regression**
    
    - recalc benchmarks
        
    - UI scroll benchmarks
        
    - batch throughput benchmarks
        
- **Security tests**
    
    - sandbox escape tests
        
    - permission enforcement tests
        

---

## 11) Product Milestones (Project-level, not “MVP-only”)

This is a full project; teams can pick slices to ship incrementally.

- **Phase A: Core viability**
    
    - XLSX fidelity + grid UX + calc baseline + Python macros
        
- **Phase B: Excel replacement credibility**
    
    - broad function parity, pivots/charts parity, migration tooling, admin controls
        
- **Phase C: Enterprise platform**
    
    - signing/policy, collaboration options, managed add-ins, observability, long-term support
        

---

## 12) Success Metrics

- **Compatibility:** % of target corpus round-trips with no repair prompts and no major layout loss
    
- **Adoption:** daily active usage, workbooks opened, retention
    
- **Automation:** # of Python macros/add-ins deployed; time-to-migrate from VBA
    
- **Performance:** recalc latency distribution; batch throughput scaling
    
- **Security:** policy compliance; no silent permission escalations
    

---

## 13) Risks (explicit)

- **XLSX edge cases:** mitigated by passthrough preservation + corpus tests
    
- **Python sandboxing:** must be a process + OS sandbox; “import hooks” alone are not sufficient
    
- **Grid accessibility:** requires deliberate design (semantic mirror)
    
- **Pivot parity:** expensive; must prioritize refresh/view early and build toward full authoring
    
- **Expectation management:** solved via compatibility panel + modes + reports

---

## Planning Suite
- Detailed implementation and delivery notes: [[docs/RootCellar/RootCellar Planning Hub]]
