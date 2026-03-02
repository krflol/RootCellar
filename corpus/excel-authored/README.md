# Excel-Authored Corpus

This directory holds curated real Excel-authored workbook samples used by the
compatibility-first interop gate.

## Layout
- `manifest.json`: metadata inventory for curated sample files.
- `files/`: place `.xlsx` samples referenced by `manifest.json`.

## Manifest Entry Fields
- `id`: stable sample identifier (unique within the manifest).
- `path`: path to workbook file relative to this manifest (for example `files/sample.xlsx`).
- `authoring_app`: Excel version/source (for example `Microsoft Excel 365`).
- `source_category`: origin class (for example `internal_template`, `customer_redacted`).
- `legal_clearance`: one of:
  - `approved`
  - `restricted_internal`
  - `restricted_partner`
- `provenance`: one of:
  - `interim_openpyxl`
  - `verified_excel` (requires workbook metadata detection to report `Microsoft Excel`).
- `features`: non-empty list of feature tags (for example `styles`, `charts`, `comments`, `defined_names`).
- `notes`: optional context for reviewers.

## Policy
- Only include workbooks with explicit legal clearance.
- Prefer redacted/minimal samples with deterministic behavior.
- Keep feature tags precise so CI gating and corpus reporting stay actionable.
- CI now enforces expanded curated baseline coverage:
  - minimum curated samples: `17`
  - required curated feature tags: `formulas`, `styles`, `comments`, `charts`, `defined_names`, `tables`, `merged_cells`, `data_validation`, `conditional_formatting`, `external_links`, `pivot_tables`, `query_connections`, `sheet_protection`, `hyperlinks`, `workbook_protection`, `print_settings`, `calc_chain`
- CI verified-Excel floor is currently `17` (`EXCEL_INTEROP_MIN_VERIFIED_EXCEL_SAMPLES`).

## Regeneration
- Verified baseline samples can be regenerated on Windows hosts with Microsoft Excel installed:
  - `powershell -NoProfile -ExecutionPolicy Bypass -File ./python/generate_excel_authored_curated_samples.ps1`
- After regeneration, run corpus assembly and interop gate checks before updating `manifest.json` metadata.
