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
- `features`: non-empty list of feature tags (for example `styles`, `charts`, `comments`, `defined_names`).
- `notes`: optional context for reviewers.

## Policy
- Only include workbooks with explicit legal clearance.
- Prefer redacted/minimal samples with deterministic behavior.
- Keep feature tags precise so CI gating and corpus reporting stay actionable.
- The current seed sample is an internal baseline workbook; migrate to verified
  Microsoft Excel-authored samples as approvals land.
- CI currently enforces at least one curated sample and baseline curated feature
  coverage for `formulas`.
- CI now enforces expanded curated baseline coverage:
  - minimum curated samples: `5`
  - required curated feature tags: `formulas`, `styles`, `comments`, `charts`, `defined_names`
