# RootCellar Desktop

Tauri + TypeScript desktop shell with a compatibility-first interop workflow.

## Current UI Scope
- Native desktop shell boots and invokes Rust commands.
- Open + inspect Excel-authored `.xlsx` workbooks.
- Show compatibility summary (feature score, issues, unknown parts, part-graph stats).
- Render sheet preview table for loaded workbook cells.
- Apply in-app value/formula edits to single cells or A1 ranges (for example `A1` or `A1:B3`).
- Formula bar supports selected-cell edits directly from preview (`Apply From Bar` or `Enter`).
- Jump-to-last-edited helper and visual highlight in preview.
- Copy selected preview cell A1/value/formula to clipboard.
- Arrow-key navigation in preview; `Enter` applies the current edit input to the selected cell.
- Recalculate loaded workbook sheets through `rootcellar-core`.
- Save loaded workbook in either:
  - `preserve` mode (passthrough when clean, sheet-overrides when dirty; interop-first)
  - `normalize` mode (model-driven rewrite)
- Optional post-save source promotion (continue preserve workflow from saved output path).
- Keep existing engine round-trip smoke command for frontend/backend wiring.

## Structure
- `src-tauri/`: Tauri Rust backend.
- `src/`: TypeScript frontend.
- `index.html`: app entry.

## Backend Commands
- `app_status`
- `engine_round_trip`
- `interop_session_status`
- `interop_open_workbook`
- `interop_sheet_preview`
- `interop_apply_cell_edit`
- `interop_recalc_loaded`
- `interop_save_workbook`

## Run
```bash
cd apps/desktop
npm install
npm run tauri dev
```

## Targeted Regression Test
```bash
cargo test --manifest-path apps/desktop/src-tauri/Cargo.toml --locked apply_cell_edit_range
```

## Notes
- Native open/save dialogs are wired through Tauri dialog plugin.
- `.xlsx` extension checks are case-insensitive (`.xlsx` / `.XLSX`).
- `preserve` mode is the default interop path and should be used first when validating Excel round-trips.
