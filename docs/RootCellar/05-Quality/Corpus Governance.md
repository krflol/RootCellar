# Corpus Governance

Parent: [[Test Strategy]]

## Corpus Segments
- Finance templates.
- Operations trackers.
- HR and planning workbooks.
- Formula stress workbooks.
- Formatting and chart-heavy workbooks.

## Metadata Per Corpus File
- Source category and legal clearance.
- Feature inventory (tables, pivots, formulas, charts, macros).
- Baseline expected behavior and compatibility status.
- Deterministic fixture manifest entry (fixture name, feature tags, generator provenance).

## Deterministic Fixture Baseline
- Script: `python/generate_corpus_fixtures.py`.
- Baseline fixtures:
  - `minimal.xlsx`, `formula.xlsx`, `dangling-edge.xlsx`, `unknown-part.xlsx`.
  - `styles.xlsx`, `comments.xlsx`, `chart.xlsx`, `defined-names.xlsx`.
- Every generated corpus directory includes `manifest.json` for downstream gate/report attribution.

## Curated Excel-Authored Segment
- Location: `corpus/excel-authored/`.
- Metadata source: `corpus/excel-authored/manifest.json`.
- Current seeded sample count: `17` (`internal-formula-baseline`, `internal-styles-baseline`, `internal-comments-baseline`, `internal-chart-baseline`, `internal-defined-names-baseline`, `internal-table-baseline`, `internal-merged-cells-baseline`, `internal-data-validation-baseline`, `internal-conditional-formatting-baseline`, `internal-external-links-baseline`, `internal-pivot-table-baseline`, `internal-query-connection-baseline`, `internal-sheet-protection-baseline`, `internal-hyperlinks-baseline`, `internal-workbook-protection-baseline`, `internal-print-settings-baseline`, `internal-calc-chain-baseline`).
- Current verified provenance count: `17` (`provenance=verified_excel` for all seeded samples).
- Required manifest fields per sample:
  - `id`, `path`, `authoring_app`, `source_category`, `legal_clearance`, `provenance`, `features`.
- Allowed legal-clearance values:
  - `approved`, `restricted_internal`, `restricted_partner`.
- Allowed provenance values:
  - `interim_openpyxl`, `verified_excel`.
- CI assembly utility:
  - `python/assemble_excel_interop_corpus.py` merges generated fixtures with curated samples and enforces `--min-excel-authored-samples` plus `--min-verified-excel-samples` policy.
  - Curated feature coverage can be enforced with repeated `--required-curated-feature` flags (current CI baseline: `formulas`, `styles`, `comments`, `charts`, `defined_names`, `tables`, `merged_cells`, `data_validation`, `conditional_formatting`, `external_links`, `pivot_tables`, `query_connections`, `sheet_protection`, `hyperlinks`, `workbook_protection`, `print_settings`, `calc_chain`).
  - Windows regeneration utility for verified baseline files: `python/generate_excel_authored_curated_samples.ps1` (Excel COM automation).

## Update Process
- New files require metadata + risk review.
- Rebaseline changes require approval from compatibility owner.
- Deprecated files archived with reason and timestamp.
- Any regenerated verified sample set must be followed by assembled-corpus + interop-gate verification and manifest provenance review before merge.

## Metrics
- No-repair rate.
- Layout diff severity.
- Formula output diff rate.
- Save/load performance deltas.
