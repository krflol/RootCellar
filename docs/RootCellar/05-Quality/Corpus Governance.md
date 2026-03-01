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
- Current seeded sample count: `5` (`internal-formula-baseline`, `internal-styles-baseline`, `internal-comments-baseline`, `internal-chart-baseline`, `internal-defined-names-baseline`).
- Required manifest fields per sample:
  - `id`, `path`, `authoring_app`, `source_category`, `legal_clearance`, `features`.
- Allowed legal-clearance values:
  - `approved`, `restricted_internal`, `restricted_partner`.
- CI assembly utility:
  - `python/assemble_excel_interop_corpus.py` merges generated fixtures with curated samples and enforces `--min-excel-authored-samples` policy.
  - Curated feature coverage can be enforced with repeated `--required-curated-feature` flags (current CI baseline: `formulas`, `styles`, `comments`, `charts`, `defined_names`).

## Update Process
- New files require metadata + risk review.
- Rebaseline changes require approval from compatibility owner.
- Deprecated files archived with reason and timestamp.

## Metrics
- No-repair rate.
- Layout diff severity.
- Formula output diff rate.
- Save/load performance deltas.
