# Performance and Benchmarking

Parent: [[Test Strategy]]

## Benchmark Families
- UI scrolling and selection responsiveness.
- Recalc latency across workbook sizes.
- Import/export throughput and memory footprint.
- CLI batch throughput scaling with thread counts.

## Benchmark Rules
- Fixed hardware profiles for baseline comparisons.
- Versioned benchmark datasets.
- Percentile reporting (p50/p95/p99) and confidence intervals.

## Regression Policy
- >10% regression in key SLO metric blocks merge unless waiver approved.
- Waivers require mitigation issue and target sprint.
- Nightly batch gate includes minimum throughput threshold checks on deterministic corpus slices (`throughput_files_per_sec`).
- Nightly batch gate can also enforce synthetic recalc benchmark thresholds (`BATCH_BENCH_MIN_DURATION_SPEEDUP_RATIO`, `BATCH_BENCH_MAX_EVALUATED_CELLS_RATIO`) from `bench recalc-synthetic` summary metrics.

## Artifact Output
- Perf samples stored as artifact bundle section `perf/`.
- Trend dashboards in [[docs/RootCellar/04-Observability/Dashboards SLOs and Alerts]].
