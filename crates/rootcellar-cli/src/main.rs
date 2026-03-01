use clap::{Parser, Subcommand, ValueEnum};
use rayon::prelude::*;
use rootcellar_core::model::CellRef;
use rootcellar_core::{
    analyze_sheet_dependencies, inspect_xlsx, load_workbook_model, preserve_xlsx_passthrough,
    preserve_xlsx_with_sheet_overrides, recalc_sheet, recalc_sheet_from_roots,
    recalc_sheet_with_dag_timing_options, save_workbook_model, CalcError, CellValue, EventEnvelope,
    EventSink, JsonlEventSink, ModelError, Mutation, NoopEventSink, RecalcDagTimingOptions,
    RecalcReport, SaveMode, TelemetryError, TraceContext, Workbook, XlsxInspectionReport,
};
use serde::{Deserialize, Serialize};
use serde_json::{json, to_string_pretty};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::time::Instant;
use thiserror::Error;

#[derive(Debug, Parser)]
#[command(name = "rootcellar")]
#[command(about = "RootCellar headless execution baseline", long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Debug, Subcommand)]
enum Commands {
    /// Inspect an XLSX file and write a compatibility report.
    Open {
        /// Path to .xlsx workbook
        file: PathBuf,
        /// Optional path to write report JSON. Defaults next to workbook.
        #[arg(long)]
        report: Option<PathBuf>,
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
    /// Run XLSX part-graph validation over a directory corpus.
    PartGraphCorpus {
        /// Directory containing .xlsx files (recursive scan)
        dir: PathBuf,
        /// Optional path to write aggregate corpus report JSON.
        #[arg(long)]
        report: Option<PathBuf>,
        /// Optional max number of files to process after deterministic sorting.
        #[arg(long = "max-files")]
        max_files: Option<usize>,
        /// Return non-zero if any workbook fails inspection.
        #[arg(long = "fail-on-errors", default_value_t = false)]
        fail_on_errors: bool,
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
    /// Run bounded-parallel recalc over an .xlsx directory corpus.
    Batch {
        #[command(subcommand)]
        command: BatchCommands,
    },
    /// Run synthetic benchmark workloads.
    Bench {
        #[command(subcommand)]
        command: BenchCommands,
    },
    /// Run an in-memory transaction demo and print snapshot output.
    TxDemo {
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
    /// Apply one or more value/formula edits via transaction and save workbook.
    TxSave {
        /// Source .xlsx workbook
        input: PathBuf,
        /// Destination .xlsx workbook
        output: PathBuf,
        /// Target sheet name
        #[arg(long, default_value = "Sheet1")]
        sheet: String,
        /// Target cell reference in A1 notation (single mutation mode)
        #[arg(long)]
        cell: Option<String>,
        /// New value (number, true/false, or text) for `--cell`
        #[arg(long)]
        value: Option<String>,
        /// Repeated cell assignments, e.g. `--set A1=10 --set B1=true`
        /// Supports ranges, e.g. `--set A1:B3=0`
        #[arg(long = "set")]
        sets: Vec<String>,
        /// Repeated formula assignments, e.g. `--setf C1==A1+B1`
        #[arg(long = "setf")]
        set_formulas: Vec<String>,
        /// Save mode after mutation
        #[arg(long, value_enum, default_value_t = CliSaveMode::Preserve)]
        mode: CliSaveMode,
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
    /// Load workbook and write a normalized XLSX output.
    Save {
        /// Source .xlsx workbook
        input: PathBuf,
        /// Destination .xlsx workbook
        output: PathBuf,
        /// Save mode (`preserve` uses passthrough copy; `normalize` rewrites from model)
        #[arg(long, value_enum, default_value_t = CliSaveMode::Normalize)]
        mode: CliSaveMode,
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
    /// Recalculate formulas from an XLSX workbook model projection.
    Recalc {
        /// Path to .xlsx workbook
        file: PathBuf,
        /// Optional target sheet; defaults to all sheets.
        #[arg(long)]
        sheet: Option<String>,
        /// Optional path to write recalc report JSON.
        #[arg(long)]
        report: Option<PathBuf>,
        /// Optional path to write dependency-graph report JSON.
        #[arg(long = "dep-graph-report")]
        dep_graph_report: Option<PathBuf>,
        /// Optional path to write recalc DAG timing report JSON.
        #[arg(long = "dag-timing-report")]
        dag_timing_report: Option<PathBuf>,
        /// Optional absolute slow-node threshold in microseconds for DAG timing analysis.
        /// Requires `--dag-timing-report`.
        #[arg(long = "dag-slow-threshold-us")]
        dag_slow_threshold_us: Option<u64>,
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
    /// Reproducibility bundle workflows.
    Repro {
        #[command(subcommand)]
        command: ReproCommands,
    },
}

#[derive(Debug, Subcommand)]
enum BatchCommands {
    /// Recalculate all workbooks in a directory corpus with bounded parallelism.
    Recalc {
        /// Directory containing .xlsx files (recursive scan)
        dir: PathBuf,
        /// Optional path to write aggregate batch report JSON.
        #[arg(long)]
        report: Option<PathBuf>,
        /// Optional target sheet; defaults to all sheets.
        #[arg(long)]
        sheet: Option<String>,
        /// Optional max number of files to process after deterministic sorting.
        #[arg(long = "max-files")]
        max_files: Option<usize>,
        /// Maximum Rayon worker threads used for processing.
        #[arg(long)]
        threads: Option<usize>,
        /// Return non-zero if any workbook fails recalc.
        #[arg(long = "fail-on-errors", default_value_t = false)]
        fail_on_errors: bool,
        /// Artifact detail level for per-file payloads.
        #[arg(long = "detail-level", value_enum, default_value_t = CliBatchDetailLevel::Minimal)]
        detail_level: CliBatchDetailLevel,
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
}

#[derive(Debug, Subcommand)]
enum BenchCommands {
    /// Benchmark full vs incremental recalc on a generated dependency workload.
    RecalcSynthetic {
        /// Optional path to write benchmark report JSON.
        #[arg(long)]
        report: Option<PathBuf>,
        /// Number of independent dependency chains (rows).
        #[arg(long, default_value_t = 16)]
        chains: usize,
        /// Number of formula cells per chain.
        #[arg(long = "chain-length", default_value_t = 256)]
        chain_length: usize,
        /// Number of benchmark iterations.
        #[arg(long, default_value_t = 5)]
        iterations: usize,
        /// 1-based chain index whose root value is changed per iteration.
        #[arg(long = "changed-chain", default_value_t = 1)]
        changed_chain: usize,
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
}

#[derive(Debug, Subcommand)]
enum ReproCommands {
    /// Record a reproducibility bundle from a workbook recalc run.
    Record {
        /// Path to .xlsx workbook
        file: PathBuf,
        /// Target bundle directory
        #[arg(long)]
        bundle: PathBuf,
        /// Optional target sheet; defaults to all sheets.
        #[arg(long)]
        sheet: Option<String>,
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
    /// Verify an existing reproducibility bundle.
    Check {
        /// Bundle directory produced by `repro record`
        bundle: PathBuf,
        /// Optional workbook path to compare against the recorded bundle.
        #[arg(long)]
        against: Option<PathBuf>,
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
    /// Print deterministic cell-level diff between bundle baseline and a workbook.
    Diff {
        /// Bundle directory produced by `repro record`
        bundle: PathBuf,
        /// Workbook path to compare against the bundle baseline
        #[arg(long)]
        against: PathBuf,
        /// Output format
        #[arg(long, value_enum, default_value_t = CliOutputFormat::Text)]
        format: CliOutputFormat,
        /// Optional path to write diff output artifact.
        #[arg(long)]
        output: Option<PathBuf>,
        /// Maximum changed cells to print
        #[arg(long, default_value_t = 50)]
        limit: usize,
        /// Optional telemetry JSONL output path.
        #[arg(long)]
        jsonl: Option<PathBuf>,
    },
}

#[derive(Debug, Clone, Copy, ValueEnum)]
enum CliSaveMode {
    Preserve,
    Normalize,
}

#[derive(Debug, Clone, Copy, ValueEnum, PartialEq, Eq)]
enum CliOutputFormat {
    Text,
    Json,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
enum CliBatchDetailLevel {
    Minimal,
    Diagnostic,
    Forensic,
}

impl From<CliSaveMode> for SaveMode {
    fn from(value: CliSaveMode) -> Self {
        match value {
            CliSaveMode::Preserve => SaveMode::Preserve,
            CliSaveMode::Normalize => SaveMode::Normalize,
        }
    }
}

impl CliOutputFormat {
    fn as_str(self) -> &'static str {
        match self {
            CliOutputFormat::Text => "text",
            CliOutputFormat::Json => "json",
        }
    }
}

impl CliBatchDetailLevel {
    fn as_str(self) -> &'static str {
        match self {
            CliBatchDetailLevel::Minimal => "minimal",
            CliBatchDetailLevel::Diagnostic => "diagnostic",
            CliBatchDetailLevel::Forensic => "forensic",
        }
    }

    fn include_recalc_payload(self) -> bool {
        matches!(
            self,
            CliBatchDetailLevel::Diagnostic | CliBatchDetailLevel::Forensic
        )
    }
}

#[derive(Debug, Error)]
enum CliError {
    #[error("calc error: {0}")]
    Calc(#[from] CalcError),
    #[error("interop error: {0}")]
    Interop(#[from] rootcellar_core::interop::InteropError),
    #[error("model error: {0}")]
    Model(#[from] ModelError),
    #[error("telemetry error: {0}")]
    Telemetry(#[from] TelemetryError),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("serialization error: {0}")]
    Serde(#[from] serde_json::Error),
    #[error("requested sheet not found: {0}")]
    SheetNotFound(String),
    #[error("workbook contains no sheets")]
    NoSheets,
    #[error("invalid cell reference: {0}")]
    InvalidCellReference(String),
    #[error("invalid mutation assignment: {0}")]
    InvalidMutationAssignment(String),
    #[error("invalid tx-save arguments: {0}")]
    InvalidTxSaveArgs(String),
    #[error("invalid recalc arguments: {0}")]
    InvalidRecalcArgs(String),
    #[error("invalid corpus arguments: {0}")]
    InvalidCorpusArgs(String),
    #[error("part-graph corpus validation failed: {0}")]
    CorpusValidationFailed(String),
    #[error("invalid batch arguments: {0}")]
    InvalidBatchArgs(String),
    #[error("batch recalc failed: {0}")]
    BatchRecalcFailed(String),
    #[error("invalid benchmark arguments: {0}")]
    InvalidBenchArgs(String),
    #[error("bundle missing required file: {0}")]
    MissingBundleFile(String),
    #[error("repro check failed: {0}")]
    ReproMismatch(String),
}

#[derive(Debug, Serialize, Deserialize)]
struct ReproManifest {
    bundle_version: String,
    bundle_id: String,
    created_at: chrono::DateTime<chrono::Utc>,
    input_file: String,
    recalc_report_file: String,
    sheet: Option<String>,
    hashes: ReproHashes,
}

#[derive(Debug, Serialize, Deserialize)]
struct ReproHashes {
    input_fnv64: String,
    recalc_report_fnv64: String,
}

#[derive(Debug)]
struct RecalcState {
    payload: serde_json::Value,
    cell_map: BTreeMap<String, String>,
}

#[derive(Debug, Serialize, Deserialize)]
struct PartGraphCorpusFileReport {
    path: String,
    issue_count: usize,
    unknown_part_count: usize,
    part_graph: rootcellar_core::WorkbookPartGraphSummary,
}

#[derive(Debug, Serialize, Deserialize)]
struct PartGraphCorpusFailure {
    path: String,
    error: String,
}

#[derive(Debug, Serialize, Deserialize)]
struct PartGraphCorpusSummary {
    discovered_files: usize,
    processed_files: usize,
    success_count: usize,
    failure_count: usize,
    total_issue_count: usize,
    total_unknown_part_count: usize,
    total_nodes: usize,
    total_edges: usize,
    total_dangling_edges: usize,
    total_external_edges: usize,
    files_with_dangling_edges: usize,
    files_with_unknown_parts: usize,
}

#[derive(Debug, Serialize, Deserialize)]
struct PartGraphCorpusReport {
    generated_at: chrono::DateTime<chrono::Utc>,
    input_dir: String,
    max_files: Option<usize>,
    fail_on_errors: bool,
    summary: PartGraphCorpusSummary,
    files: Vec<PartGraphCorpusFileReport>,
    failures: Vec<PartGraphCorpusFailure>,
}

#[derive(Debug, Serialize, Deserialize)]
struct BatchRecalcFileReport {
    path: String,
    duration_ms: u128,
    sheet_count: usize,
    evaluated_cells: usize,
    cycle_count: usize,
    parse_error_count: usize,
    value_fingerprint_fnv64: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    recalc_report: Option<serde_json::Value>,
}

#[derive(Debug, Serialize, Deserialize)]
struct BatchRecalcFailure {
    path: String,
    duration_ms: u128,
    error: String,
}

#[derive(Debug, Serialize, Deserialize)]
struct BatchRecalcSummary {
    discovered_files: usize,
    processed_files: usize,
    success_count: usize,
    failure_count: usize,
    configured_threads: Option<usize>,
    effective_threads: usize,
    total_sheet_count: usize,
    total_evaluated_cells: usize,
    total_cycle_count: usize,
    total_parse_error_count: usize,
    total_file_duration_ms: u128,
    max_file_duration_ms: u128,
    wall_clock_duration_ms: u128,
    throughput_files_per_sec: f64,
    aggregate_file_time_ratio: f64,
}

#[derive(Debug, Serialize, Deserialize)]
struct BatchRecalcReport {
    generated_at: chrono::DateTime<chrono::Utc>,
    input_dir: String,
    report_detail_level: String,
    sheet: Option<String>,
    max_files: Option<usize>,
    fail_on_errors: bool,
    summary: BatchRecalcSummary,
    files: Vec<BatchRecalcFileReport>,
    failures: Vec<BatchRecalcFailure>,
}

#[derive(Debug, Serialize, Deserialize)]
struct BenchRecalcSyntheticIteration {
    iteration: usize,
    full_duration_us: u128,
    incremental_duration_us: u128,
    full_evaluated_cells: usize,
    incremental_evaluated_cells: usize,
    duration_speedup_ratio: f64,
    evaluated_cells_reduction_ratio: f64,
}

#[derive(Debug, Serialize, Deserialize)]
struct BenchRecalcSyntheticSummary {
    chains: usize,
    chain_length: usize,
    iterations: usize,
    changed_chain: usize,
    total_formula_cells: usize,
    average_full_duration_us: f64,
    average_incremental_duration_us: f64,
    duration_speedup_ratio: f64,
    average_full_evaluated_cells: f64,
    average_incremental_evaluated_cells: f64,
    evaluated_cells_reduction_ratio: f64,
}

#[derive(Debug, Serialize, Deserialize)]
struct BenchRecalcSyntheticReport {
    generated_at: chrono::DateTime<chrono::Utc>,
    benchmark: String,
    summary: BenchRecalcSyntheticSummary,
    iterations: Vec<BenchRecalcSyntheticIteration>,
}

fn main() {
    if let Err(err) = run() {
        eprintln!("error: {err}");
        std::process::exit(1);
    }
}

fn run() -> Result<(), CliError> {
    let cli = Cli::parse();

    match cli.command {
        Commands::Open {
            file,
            report,
            jsonl,
        } => run_open(file.as_path(), report.as_ref(), jsonl.as_ref()),
        Commands::PartGraphCorpus {
            dir,
            report,
            max_files,
            fail_on_errors,
            jsonl,
        } => run_part_graph_corpus(
            dir.as_path(),
            report.as_ref(),
            max_files,
            fail_on_errors,
            jsonl.as_ref(),
        ),
        Commands::Batch { command } => match command {
            BatchCommands::Recalc {
                dir,
                report,
                sheet,
                max_files,
                threads,
                fail_on_errors,
                detail_level,
                jsonl,
            } => run_batch_recalc(
                dir.as_path(),
                report.as_ref(),
                sheet.as_ref(),
                max_files,
                threads,
                fail_on_errors,
                detail_level,
                jsonl.as_ref(),
            ),
        },
        Commands::Bench { command } => match command {
            BenchCommands::RecalcSynthetic {
                report,
                chains,
                chain_length,
                iterations,
                changed_chain,
                jsonl,
            } => run_bench_recalc_synthetic(
                report.as_ref(),
                chains,
                chain_length,
                iterations,
                changed_chain,
                jsonl.as_ref(),
            ),
        },
        Commands::TxDemo { jsonl } => run_tx_demo(jsonl.as_ref()),
        Commands::TxSave {
            input,
            output,
            sheet,
            cell,
            value,
            sets,
            set_formulas,
            mode,
            jsonl,
        } => run_tx_save(
            input.as_path(),
            output.as_path(),
            &sheet,
            cell.as_ref(),
            value.as_ref(),
            &sets,
            &set_formulas,
            mode,
            jsonl.as_ref(),
        ),
        Commands::Save {
            input,
            output,
            mode,
            jsonl,
        } => run_save(input.as_path(), output.as_path(), mode, jsonl.as_ref()),
        Commands::Recalc {
            file,
            sheet,
            report,
            dep_graph_report,
            dag_timing_report,
            dag_slow_threshold_us,
            jsonl,
        } => run_recalc(
            file.as_path(),
            sheet.as_ref(),
            report.as_ref(),
            dep_graph_report.as_ref(),
            dag_timing_report.as_ref(),
            dag_slow_threshold_us,
            jsonl.as_ref(),
        ),
        Commands::Repro { command } => match command {
            ReproCommands::Record {
                file,
                bundle,
                sheet,
                jsonl,
            } => run_repro_record(
                file.as_path(),
                bundle.as_path(),
                sheet.as_ref(),
                jsonl.as_ref(),
            ),
            ReproCommands::Check {
                bundle,
                against,
                jsonl,
            } => run_repro_check(bundle.as_path(), against.as_ref(), jsonl.as_ref()),
            ReproCommands::Diff {
                bundle,
                against,
                format,
                output,
                limit,
                jsonl,
            } => run_repro_diff(
                bundle.as_path(),
                against.as_path(),
                format,
                output.as_ref(),
                limit,
                jsonl.as_ref(),
            ),
        },
    }
}

fn run_open(
    file: &Path,
    report_override: Option<&PathBuf>,
    jsonl_path: Option<&PathBuf>,
) -> Result<(), CliError> {
    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;

    let report = inspect_xlsx(file, sink.as_mut(), &trace)?;
    let report_path = match report_override {
        Some(path) => path.to_path_buf(),
        None => default_report_path(file),
    };

    fs::write(&report_path, to_string_pretty(&report)?)?;

    print_open_summary(&report, &report_path);

    Ok(())
}

fn run_part_graph_corpus(
    dir: &Path,
    report_override: Option<&PathBuf>,
    max_files: Option<usize>,
    fail_on_errors: bool,
    jsonl_path: Option<&PathBuf>,
) -> Result<(), CliError> {
    let input_dir_label = normalize_path_for_report(dir);
    if !dir.exists() {
        return Err(CliError::InvalidCorpusArgs(format!(
            "directory does not exist: {}",
            dir.display()
        )));
    }
    if !dir.is_dir() {
        return Err(CliError::InvalidCorpusArgs(format!(
            "path is not a directory: {}",
            dir.display()
        )));
    }
    if matches!(max_files, Some(0)) {
        return Err(CliError::InvalidCorpusArgs(
            "--max-files must be greater than zero when provided".to_string(),
        ));
    }

    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;
    let mut discovered = collect_xlsx_files_recursive(dir)?;
    let discovered_count = discovered.len();
    if let Some(limit) = max_files {
        if discovered.len() > limit {
            discovered.truncate(limit);
        }
    }
    let processed_count = discovered.len();

    sink.emit(
        EventEnvelope::info("artifact.part_graph.corpus.start", &trace)
            .with_context(json!({
                "input_dir": input_dir_label,
                "report": report_override.map(|p| normalize_path_for_report(p.as_path())),
                "max_files": max_files,
                "fail_on_errors": fail_on_errors,
            }))
            .with_metrics(json!({
                "discovered_files": discovered_count,
                "processed_files": processed_count,
            })),
    )?;

    let mut files = Vec::<PartGraphCorpusFileReport>::new();
    let mut failures = Vec::<PartGraphCorpusFailure>::new();
    let mut total_issue_count = 0usize;
    let mut total_unknown_part_count = 0usize;
    let mut total_nodes = 0usize;
    let mut total_edges = 0usize;
    let mut total_dangling_edges = 0usize;
    let mut total_external_edges = 0usize;
    let mut files_with_dangling_edges = 0usize;
    let mut files_with_unknown_parts = 0usize;

    for file in &discovered {
        match inspect_xlsx(file, sink.as_mut(), &trace) {
            Ok(report) => {
                total_issue_count += report.summary.issue_count;
                total_unknown_part_count += report.summary.unknown_part_count;
                total_nodes += report.part_graph.node_count;
                total_edges += report.part_graph.edge_count;
                total_dangling_edges += report.part_graph.dangling_edge_count;
                total_external_edges += report.part_graph.external_edge_count;
                if report.part_graph.dangling_edge_count > 0 {
                    files_with_dangling_edges += 1;
                }
                if report.summary.unknown_part_count > 0 {
                    files_with_unknown_parts += 1;
                }
                files.push(PartGraphCorpusFileReport {
                    path: normalize_path_for_report(file.as_path()),
                    issue_count: report.summary.issue_count,
                    unknown_part_count: report.summary.unknown_part_count,
                    part_graph: report.part_graph.summary(),
                });
            }
            Err(err) => {
                failures.push(PartGraphCorpusFailure {
                    path: normalize_path_for_report(file.as_path()),
                    error: err.to_string(),
                });
            }
        }
    }

    files.sort_by(|a, b| a.path.cmp(&b.path));
    failures.sort_by(|a, b| a.path.cmp(&b.path));

    let summary = PartGraphCorpusSummary {
        discovered_files: discovered_count,
        processed_files: processed_count,
        success_count: files.len(),
        failure_count: failures.len(),
        total_issue_count,
        total_unknown_part_count,
        total_nodes,
        total_edges,
        total_dangling_edges,
        total_external_edges,
        files_with_dangling_edges,
        files_with_unknown_parts,
    };

    let report = PartGraphCorpusReport {
        generated_at: chrono::Utc::now(),
        input_dir: input_dir_label.clone(),
        max_files,
        fail_on_errors,
        summary,
        files,
        failures,
    };

    let report_path = report_override
        .cloned()
        .unwrap_or_else(|| default_part_graph_corpus_report_path(dir));
    let report_json = to_string_pretty(&report)?;
    fs::write(&report_path, report_json.as_bytes())?;

    sink.emit(
        EventEnvelope::info("artifact.part_graph.corpus.end", &trace)
            .with_context(json!({
                "input_dir": input_dir_label,
                "report": normalize_path_for_report(report_path.as_path()),
                "max_files": max_files,
                "fail_on_errors": fail_on_errors,
            }))
            .with_metrics(json!({
                "discovered_files": report.summary.discovered_files,
                "processed_files": report.summary.processed_files,
                "success_count": report.summary.success_count,
                "failure_count": report.summary.failure_count,
                "total_issue_count": report.summary.total_issue_count,
                "total_unknown_part_count": report.summary.total_unknown_part_count,
                "total_nodes": report.summary.total_nodes,
                "total_edges": report.summary.total_edges,
                "total_dangling_edges": report.summary.total_dangling_edges,
                "total_external_edges": report.summary.total_external_edges,
                "output_bytes": report_json.len(),
            })),
    )?;

    println!("Corpus directory: {}", dir.display());
    println!("Corpus report: {}", report_path.display());
    println!(
        "Corpus summary: discovered={}, processed={}, success={}, failures={}",
        report.summary.discovered_files,
        report.summary.processed_files,
        report.summary.success_count,
        report.summary.failure_count
    );
    println!(
        "Part graph totals: nodes={}, edges={}, dangling_edges={}, external_edges={}, unknown_parts={}",
        report.summary.total_nodes,
        report.summary.total_edges,
        report.summary.total_dangling_edges,
        report.summary.total_external_edges,
        report.summary.total_unknown_part_count
    );
    if report.summary.failure_count > 0 {
        println!("Failed files (up to 10):");
        for failure in report.failures.iter().take(10) {
            println!("  - {} :: {}", failure.path, failure.error);
        }
        if report.failures.len() > 10 {
            println!("  - ... {} additional failures", report.failures.len() - 10);
        }
    }

    if fail_on_errors && report.summary.failure_count > 0 {
        return Err(CliError::CorpusValidationFailed(format!(
            "{} workbook(s) failed inspection in corpus run",
            report.summary.failure_count
        )));
    }

    Ok(())
}

fn run_batch_recalc(
    dir: &Path,
    report_override: Option<&PathBuf>,
    sheet: Option<&String>,
    max_files: Option<usize>,
    threads: Option<usize>,
    fail_on_errors: bool,
    detail_level: CliBatchDetailLevel,
    jsonl_path: Option<&PathBuf>,
) -> Result<(), CliError> {
    let input_dir_label = normalize_path_for_report(dir);
    if !dir.exists() {
        return Err(CliError::InvalidBatchArgs(format!(
            "directory does not exist: {}",
            dir.display()
        )));
    }
    if !dir.is_dir() {
        return Err(CliError::InvalidBatchArgs(format!(
            "path is not a directory: {}",
            dir.display()
        )));
    }
    if matches!(max_files, Some(0)) {
        return Err(CliError::InvalidBatchArgs(
            "--max-files must be greater than zero when provided".to_string(),
        ));
    }
    if matches!(threads, Some(0)) {
        return Err(CliError::InvalidBatchArgs(
            "--threads must be greater than zero when provided".to_string(),
        ));
    }

    let requested_threads = threads.unwrap_or_else(default_batch_threads);
    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;
    let mut discovered = collect_xlsx_files_recursive(dir)?;
    let discovered_count = discovered.len();
    if let Some(limit) = max_files {
        if discovered.len() > limit {
            discovered.truncate(limit);
        }
    }
    let processed_count = discovered.len();
    let effective_threads = requested_threads.max(1).min(processed_count.max(1));

    sink.emit(
        EventEnvelope::info("artifact.batch.recalc.start", &trace)
            .with_context(json!({
                "input_dir": input_dir_label,
                "report": report_override.map(|p| normalize_path_for_report(p.as_path())),
                "sheet": sheet,
                "max_files": max_files,
                "fail_on_errors": fail_on_errors,
                "configured_threads": threads,
                "effective_threads": effective_threads,
                "detail_level": detail_level.as_str(),
            }))
            .with_metrics(json!({
                "discovered_files": discovered_count,
                "processed_files": processed_count,
            })),
    )?;

    #[derive(Debug)]
    enum BatchRecalcOutcome {
        Success(BatchRecalcFileReport),
        Failure(BatchRecalcFailure),
    }

    let wall_start = Instant::now();
    let outcomes = if discovered.is_empty() {
        Vec::<BatchRecalcOutcome>::new()
    } else {
        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(effective_threads)
            .build()
            .map_err(|err| {
                CliError::InvalidBatchArgs(format!("failed to initialize Rayon thread pool: {err}"))
            })?;
        pool.install(|| {
            discovered
                .par_iter()
                .map(|file| {
                    let started = Instant::now();
                    let mut worker_sink = NoopEventSink;
                    let worker_trace = trace.child();
                    match build_recalc_state(file.as_path(), sheet, &mut worker_sink, &worker_trace)
                    {
                        Ok(state) => {
                            let duration_ms = started.elapsed().as_millis();
                            let sheet_count =
                                state.payload["sheet_count"].as_u64().unwrap_or(0) as usize;
                            let evaluated_cells = state.payload["totals"]["evaluated_cells"]
                                .as_u64()
                                .unwrap_or(0)
                                as usize;
                            let cycle_count =
                                state.payload["totals"]["cycle_count"].as_u64().unwrap_or(0)
                                    as usize;
                            let parse_error_count = state.payload["totals"]["parse_error_count"]
                                .as_u64()
                                .unwrap_or(0)
                                as usize;
                            let value_fingerprint_fnv64 = state.payload["value_fingerprint_fnv64"]
                                .as_str()
                                .unwrap_or("unknown")
                                .to_string();
                            BatchRecalcOutcome::Success(BatchRecalcFileReport {
                                path: normalize_path_for_report(file.as_path()),
                                duration_ms,
                                sheet_count,
                                evaluated_cells,
                                cycle_count,
                                parse_error_count,
                                value_fingerprint_fnv64,
                                recalc_report: if detail_level.include_recalc_payload() {
                                    Some(state.payload)
                                } else {
                                    None
                                },
                            })
                        }
                        Err(err) => BatchRecalcOutcome::Failure(BatchRecalcFailure {
                            path: normalize_path_for_report(file.as_path()),
                            duration_ms: started.elapsed().as_millis(),
                            error: err.to_string(),
                        }),
                    }
                })
                .collect::<Vec<_>>()
        })
    };
    let wall_clock_duration_ms = wall_start.elapsed().as_millis();

    let mut files = Vec::<BatchRecalcFileReport>::new();
    let mut failures = Vec::<BatchRecalcFailure>::new();
    let mut total_sheet_count = 0usize;
    let mut total_evaluated_cells = 0usize;
    let mut total_cycle_count = 0usize;
    let mut total_parse_error_count = 0usize;
    let mut total_file_duration_ms = 0u128;
    let mut max_file_duration_ms = 0u128;
    for outcome in outcomes {
        match outcome {
            BatchRecalcOutcome::Success(file_report) => {
                total_sheet_count += file_report.sheet_count;
                total_evaluated_cells += file_report.evaluated_cells;
                total_cycle_count += file_report.cycle_count;
                total_parse_error_count += file_report.parse_error_count;
                total_file_duration_ms += file_report.duration_ms;
                max_file_duration_ms = max_file_duration_ms.max(file_report.duration_ms);
                sink.emit(
                    EventEnvelope::info("artifact.batch.recalc.file", &trace)
                        .with_context(json!({
                            "path": file_report.path,
                            "status": "ok",
                            "detail_level": detail_level.as_str(),
                        }))
                        .with_metrics(json!({
                            "duration_ms": file_report.duration_ms,
                            "sheet_count": file_report.sheet_count,
                            "evaluated_cells": file_report.evaluated_cells,
                            "cycle_count": file_report.cycle_count,
                            "parse_error_count": file_report.parse_error_count,
                        })),
                )?;
                files.push(file_report);
            }
            BatchRecalcOutcome::Failure(failure) => {
                total_file_duration_ms += failure.duration_ms;
                max_file_duration_ms = max_file_duration_ms.max(failure.duration_ms);
                sink.emit(
                    EventEnvelope::info("artifact.batch.recalc.file", &trace)
                        .with_context(json!({
                            "path": failure.path,
                            "status": "error",
                            "error": failure.error,
                            "detail_level": detail_level.as_str(),
                        }))
                        .with_metrics(json!({
                            "duration_ms": failure.duration_ms,
                        })),
                )?;
                failures.push(failure);
            }
        }
    }

    files.sort_by(|a, b| a.path.cmp(&b.path));
    failures.sort_by(|a, b| a.path.cmp(&b.path));

    let summary = BatchRecalcSummary {
        discovered_files: discovered_count,
        processed_files: processed_count,
        success_count: files.len(),
        failure_count: failures.len(),
        configured_threads: threads,
        effective_threads,
        total_sheet_count,
        total_evaluated_cells,
        total_cycle_count,
        total_parse_error_count,
        total_file_duration_ms,
        max_file_duration_ms,
        wall_clock_duration_ms,
        throughput_files_per_sec: compute_files_per_second(processed_count, wall_clock_duration_ms),
        aggregate_file_time_ratio: compute_duration_ratio(
            total_file_duration_ms,
            wall_clock_duration_ms,
        ),
    };

    let report = BatchRecalcReport {
        generated_at: chrono::Utc::now(),
        input_dir: input_dir_label.clone(),
        report_detail_level: detail_level.as_str().to_string(),
        sheet: sheet.cloned(),
        max_files,
        fail_on_errors,
        summary,
        files,
        failures,
    };

    let report_path = report_override
        .cloned()
        .unwrap_or_else(|| default_batch_recalc_report_path(dir));
    let report_json = to_string_pretty(&report)?;
    fs::write(&report_path, report_json.as_bytes())?;

    sink.emit(
        EventEnvelope::info("artifact.batch.recalc.end", &trace)
            .with_context(json!({
                "input_dir": input_dir_label,
                "report": normalize_path_for_report(report_path.as_path()),
                "sheet": sheet,
                "max_files": max_files,
                "fail_on_errors": fail_on_errors,
                "configured_threads": threads,
                "effective_threads": effective_threads,
                "detail_level": detail_level.as_str(),
            }))
            .with_metrics(json!({
                "discovered_files": report.summary.discovered_files,
                "processed_files": report.summary.processed_files,
                "success_count": report.summary.success_count,
                "failure_count": report.summary.failure_count,
                "total_sheet_count": report.summary.total_sheet_count,
                "total_evaluated_cells": report.summary.total_evaluated_cells,
                "total_cycle_count": report.summary.total_cycle_count,
                "total_parse_error_count": report.summary.total_parse_error_count,
                "total_file_duration_ms": report.summary.total_file_duration_ms,
                "max_file_duration_ms": report.summary.max_file_duration_ms,
                "wall_clock_duration_ms": report.summary.wall_clock_duration_ms,
                "throughput_files_per_sec": report.summary.throughput_files_per_sec,
                "aggregate_file_time_ratio": report.summary.aggregate_file_time_ratio,
                "output_bytes": report_json.len(),
            })),
    )?;

    println!("Batch corpus directory: {}", dir.display());
    println!("Batch report: {}", report_path.display());
    println!(
        "Batch summary: discovered={}, processed={}, success={}, failures={}",
        report.summary.discovered_files,
        report.summary.processed_files,
        report.summary.success_count,
        report.summary.failure_count
    );
    println!(
        "Batch metrics: sheets={}, evaluated_cells={}, cycles={}, parse_errors={}",
        report.summary.total_sheet_count,
        report.summary.total_evaluated_cells,
        report.summary.total_cycle_count,
        report.summary.total_parse_error_count
    );
    println!(
        "Batch runtime: configured_threads={:?}, effective_threads={}, wall_clock_ms={}, total_file_ms={}",
        report.summary.configured_threads,
        report.summary.effective_threads,
        report.summary.wall_clock_duration_ms,
        report.summary.total_file_duration_ms
    );
    println!(
        "Batch throughput: files_per_sec={:.2}, aggregate_file_time_ratio={:.2}",
        report.summary.throughput_files_per_sec, report.summary.aggregate_file_time_ratio
    );
    if report.summary.failure_count > 0 {
        println!("Failed files (up to 10):");
        for failure in report.failures.iter().take(10) {
            println!("  - {} :: {}", failure.path, failure.error);
        }
        if report.failures.len() > 10 {
            println!("  - ... {} additional failures", report.failures.len() - 10);
        }
    }

    if fail_on_errors && report.summary.failure_count > 0 {
        return Err(CliError::BatchRecalcFailed(format!(
            "{} workbook(s) failed recalc in batch run",
            report.summary.failure_count
        )));
    }

    Ok(())
}

fn run_bench_recalc_synthetic(
    report_override: Option<&PathBuf>,
    chains: usize,
    chain_length: usize,
    iterations: usize,
    changed_chain: usize,
    jsonl_path: Option<&PathBuf>,
) -> Result<(), CliError> {
    if chains == 0 {
        return Err(CliError::InvalidBenchArgs(
            "--chains must be greater than zero".to_string(),
        ));
    }
    if chain_length == 0 {
        return Err(CliError::InvalidBenchArgs(
            "--chain-length must be greater than zero".to_string(),
        ));
    }
    if iterations == 0 {
        return Err(CliError::InvalidBenchArgs(
            "--iterations must be greater than zero".to_string(),
        ));
    }
    if changed_chain == 0 || changed_chain > chains {
        return Err(CliError::InvalidBenchArgs(format!(
            "--changed-chain must be within 1..={chains}"
        )));
    }
    if chains > u32::MAX as usize {
        return Err(CliError::InvalidBenchArgs(format!(
            "--chains exceeds supported row range: {chains}"
        )));
    }
    if chain_length > (u32::MAX as usize).saturating_sub(1) {
        return Err(CliError::InvalidBenchArgs(format!(
            "--chain-length exceeds supported column range: {chain_length}"
        )));
    }

    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;
    sink.emit(
        EventEnvelope::info("benchmark.recalc.synthetic.start", &trace).with_context(json!({
            "chains": chains,
            "chain_length": chain_length,
            "iterations": iterations,
            "changed_chain": changed_chain,
            "report": report_override
                .map(|p| normalize_path_for_report(p.as_path()))
                .unwrap_or_else(|| normalize_path_for_report(default_bench_recalc_report_path().as_path())),
        })),
    )?;

    let template = build_synthetic_recalc_benchmark_workbook(chains, chain_length)?;
    let changed_row = changed_chain as u32;
    let changed_root = CellRef {
        row: changed_row,
        col: 1,
    };

    let mut iteration_rows = Vec::<BenchRecalcSyntheticIteration>::with_capacity(iterations);
    let mut full_duration_total_us = 0f64;
    let mut incremental_duration_total_us = 0f64;
    let mut full_eval_total = 0f64;
    let mut incremental_eval_total = 0f64;

    for iteration in 0..iterations {
        let next_value = (iteration as f64) + 1.0;

        let mut full_wb = template.clone();
        let mut no_op_sink = NoopEventSink;
        let mut full_txn = full_wb.begin_txn(&mut no_op_sink, &trace).expect("begin");
        full_txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: changed_row,
            col: 1,
            value: CellValue::Number(next_value),
        });
        full_txn
            .commit(&mut full_wb, &mut no_op_sink, &trace)
            .expect("commit");

        let full_started = Instant::now();
        let full_report = recalc_sheet(&mut full_wb, "Sheet1", &mut no_op_sink, &trace)?;
        let full_duration_us = full_started.elapsed().as_micros();

        let mut incremental_wb = template.clone();
        let mut incremental_txn = incremental_wb
            .begin_txn(&mut no_op_sink, &trace)
            .expect("begin");
        incremental_txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: changed_row,
            col: 1,
            value: CellValue::Number(next_value),
        });
        incremental_txn
            .commit(&mut incremental_wb, &mut no_op_sink, &trace)
            .expect("commit");

        let incremental_started = Instant::now();
        let incremental_report = recalc_sheet_from_roots(
            &mut incremental_wb,
            "Sheet1",
            &[changed_root],
            &mut no_op_sink,
            &trace,
        )?;
        let incremental_duration_us = incremental_started.elapsed().as_micros();

        let duration_speedup_ratio = if incremental_duration_us == 0 {
            0.0
        } else {
            full_duration_us as f64 / incremental_duration_us as f64
        };
        let evaluated_cells_reduction_ratio = if full_report.evaluated_cells == 0 {
            0.0
        } else {
            incremental_report.evaluated_cells as f64 / full_report.evaluated_cells as f64
        };

        full_duration_total_us += full_duration_us as f64;
        incremental_duration_total_us += incremental_duration_us as f64;
        full_eval_total += full_report.evaluated_cells as f64;
        incremental_eval_total += incremental_report.evaluated_cells as f64;

        iteration_rows.push(BenchRecalcSyntheticIteration {
            iteration: iteration + 1,
            full_duration_us,
            incremental_duration_us,
            full_evaluated_cells: full_report.evaluated_cells,
            incremental_evaluated_cells: incremental_report.evaluated_cells,
            duration_speedup_ratio,
            evaluated_cells_reduction_ratio,
        });
    }

    let count = iterations as f64;
    let average_full_duration_us = full_duration_total_us / count;
    let average_incremental_duration_us = incremental_duration_total_us / count;
    let duration_speedup_ratio = if average_incremental_duration_us == 0.0 {
        0.0
    } else {
        average_full_duration_us / average_incremental_duration_us
    };
    let average_full_evaluated_cells = full_eval_total / count;
    let average_incremental_evaluated_cells = incremental_eval_total / count;
    let evaluated_cells_reduction_ratio = if average_full_evaluated_cells == 0.0 {
        0.0
    } else {
        average_incremental_evaluated_cells / average_full_evaluated_cells
    };

    let report = BenchRecalcSyntheticReport {
        generated_at: chrono::Utc::now(),
        benchmark: "recalc_synthetic".to_string(),
        summary: BenchRecalcSyntheticSummary {
            chains,
            chain_length,
            iterations,
            changed_chain,
            total_formula_cells: chains.saturating_mul(chain_length),
            average_full_duration_us,
            average_incremental_duration_us,
            duration_speedup_ratio,
            average_full_evaluated_cells,
            average_incremental_evaluated_cells,
            evaluated_cells_reduction_ratio,
        },
        iterations: iteration_rows,
    };

    let report_path = report_override
        .cloned()
        .unwrap_or_else(default_bench_recalc_report_path);
    let report_json = to_string_pretty(&report)?;
    fs::write(&report_path, report_json.as_bytes())?;

    sink.emit(
        EventEnvelope::info("benchmark.recalc.synthetic.end", &trace)
            .with_context(json!({
                "report": normalize_path_for_report(report_path.as_path()),
            }))
            .with_metrics(json!({
                "chains": chains,
                "chain_length": chain_length,
                "iterations": iterations,
                "changed_chain": changed_chain,
                "total_formula_cells": report.summary.total_formula_cells,
                "average_full_duration_us": report.summary.average_full_duration_us,
                "average_incremental_duration_us": report.summary.average_incremental_duration_us,
                "duration_speedup_ratio": report.summary.duration_speedup_ratio,
                "average_full_evaluated_cells": report.summary.average_full_evaluated_cells,
                "average_incremental_evaluated_cells": report.summary.average_incremental_evaluated_cells,
                "evaluated_cells_reduction_ratio": report.summary.evaluated_cells_reduction_ratio,
                "output_bytes": report_json.len(),
            })),
    )?;

    println!(
        "Synthetic recalc benchmark report: {}",
        report_path.display()
    );
    println!(
        "Workload: chains={}, chain_length={}, iterations={}, changed_chain={}",
        chains, chain_length, iterations, changed_chain
    );
    println!(
        "Average duration (us): full={:.2}, incremental={:.2}, speedup={:.2}x",
        report.summary.average_full_duration_us,
        report.summary.average_incremental_duration_us,
        report.summary.duration_speedup_ratio
    );
    println!(
        "Average evaluated cells: full={:.2}, incremental={:.2}, ratio={:.4}",
        report.summary.average_full_evaluated_cells,
        report.summary.average_incremental_evaluated_cells,
        report.summary.evaluated_cells_reduction_ratio
    );

    Ok(())
}

fn build_synthetic_recalc_benchmark_workbook(
    chains: usize,
    chain_length: usize,
) -> Result<Workbook, CliError> {
    let mut workbook = Workbook::new();
    let trace = TraceContext::root();
    let mut sink = NoopEventSink;
    let mut txn = workbook.begin_txn(&mut sink, &trace).expect("begin");

    for chain_idx in 0..chains {
        let row = (chain_idx + 1) as u32;
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row,
            col: 1,
            value: CellValue::Number(chain_idx as f64),
        });
        for step in 0..chain_length {
            let col = (step + 2) as u32;
            let prior_ref = to_a1(row, col - 1);
            txn.apply(Mutation::SetCellFormula {
                sheet: "Sheet1".to_string(),
                row,
                col,
                formula: format!("={prior_ref}+1"),
                cached_value: CellValue::Empty,
            });
        }
    }

    txn.commit(&mut workbook, &mut sink, &trace)
        .expect("commit");
    Ok(workbook)
}

fn run_save(
    input: &Path,
    output: &Path,
    mode: CliSaveMode,
    jsonl_path: Option<&PathBuf>,
) -> Result<(), CliError> {
    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;
    let workbook = load_workbook_model(input, sink.as_mut(), &trace)?;
    let report = match mode {
        CliSaveMode::Preserve => {
            preserve_xlsx_passthrough(input, output, &workbook, sink.as_mut(), &trace)?
        }
        CliSaveMode::Normalize => {
            save_workbook_model(&workbook, output, mode.into(), sink.as_mut(), &trace)?
        }
    };

    println!("Input: {}", input.display());
    println!("Output: {}", report.output_path.display());
    println!(
        "Saved workbook: mode={:?}, sheets={}, cells={}, copied_bytes={}",
        report.mode, report.sheet_count, report.cell_count, report.copied_bytes
    );
    println!(
        "Part graph: nodes={}, edges={}, dangling_edges={}, unknown_parts={}",
        report.part_graph.node_count,
        report.part_graph.edge_count,
        report.part_graph.dangling_edge_count,
        report.part_graph.unknown_part_count
    );
    println!(
        "Part graph flags: strategy={}, source_graph_reused={}, relationships_preserved={}, unknown_parts_preserved={}",
        report.part_graph_flags.strategy,
        report.part_graph_flags.source_graph_reused,
        report.part_graph_flags.relationships_preserved,
        report.part_graph_flags.unknown_parts_preserved
    );

    Ok(())
}

fn run_tx_demo(jsonl_path: Option<&PathBuf>) -> Result<(), CliError> {
    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;

    let mut workbook = Workbook::new();
    let mut txn = workbook.begin_txn(sink.as_mut(), &trace)?;

    txn.apply(Mutation::SetCellValue {
        sheet: "Sheet1".to_string(),
        row: 1,
        col: 1,
        value: CellValue::Number(40.0),
    });
    txn.apply(Mutation::SetCellValue {
        sheet: "Sheet1".to_string(),
        row: 2,
        col: 1,
        value: CellValue::Number(2.0),
    });
    txn.apply(Mutation::SetCellFormula {
        sheet: "Sheet1".to_string(),
        row: 3,
        col: 1,
        formula: "=A1+A2".to_string(),
        cached_value: CellValue::Empty,
    });

    let result = txn.commit(&mut workbook, sink.as_mut(), &trace)?;
    let recalc = recalc_sheet(&mut workbook, "Sheet1", sink.as_mut(), &trace)?;
    println!(
        "Committed transaction {} with {} mutations across {} sheets",
        result.txn_id,
        result.mutation_count,
        result.changed_cells.len()
    );
    println!(
        "Recalc report: sheet={}, evaluated_cells={}, cycle_count={}, parse_errors={}",
        recalc.sheet, recalc.evaluated_cells, recalc.cycle_count, recalc.parse_error_count
    );
    println!("Workbook snapshot: {}", workbook.snapshot_json());

    Ok(())
}

fn run_tx_save(
    input: &Path,
    output: &Path,
    sheet: &str,
    cell: Option<&String>,
    value: Option<&String>,
    sets: &[String],
    set_formulas: &[String],
    mode: CliSaveMode,
    jsonl_path: Option<&PathBuf>,
) -> Result<(), CliError> {
    #[derive(Debug, Clone)]
    enum TxAssignment {
        Value {
            cell_ref: String,
            row: u32,
            col: u32,
            value: CellValue,
        },
        Formula {
            cell_ref: String,
            row: u32,
            col: u32,
            formula: String,
        },
    }

    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;
    let mut workbook = load_workbook_model(input, sink.as_mut(), &trace)?;
    if !workbook.sheets.contains_key(sheet) {
        return Err(CliError::SheetNotFound(sheet.to_string()));
    }

    let mut assignments = Vec::<TxAssignment>::new();
    if let (Some(cell), Some(value)) = (cell, value) {
        let (row, col) = parse_a1_cell_ref(cell)
            .ok_or_else(|| CliError::InvalidCellReference(cell.to_string()))?;
        assignments.push(TxAssignment::Value {
            cell_ref: cell.to_ascii_uppercase(),
            row,
            col,
            value: parse_cli_cell_value(value),
        });
    } else if cell.is_some() || value.is_some() {
        return Err(CliError::InvalidTxSaveArgs(
            "tx-save requires both --cell and --value when using single mutation mode".to_string(),
        ));
    }

    for set in sets {
        let (cell_ref, cell_value) = parse_set_assignment(set)?;
        let parsed_value = parse_cli_cell_value(&cell_value);
        if let Some(((start_row, start_col), (end_row, end_col))) = parse_range_ref(&cell_ref) {
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    assignments.push(TxAssignment::Value {
                        cell_ref: to_a1(row, col),
                        row,
                        col,
                        value: parsed_value.clone(),
                    });
                }
            }
        } else {
            let (row, col) = parse_a1_cell_ref(&cell_ref)
                .ok_or_else(|| CliError::InvalidCellReference(cell_ref.clone()))?;
            assignments.push(TxAssignment::Value {
                cell_ref,
                row,
                col,
                value: parsed_value,
            });
        }
    }

    for setf in set_formulas {
        let (cell_ref, raw_formula) = parse_set_assignment(setf)?;
        if parse_range_ref(&cell_ref).is_some() {
            return Err(CliError::InvalidTxSaveArgs(
                "--setf does not support range targets; use single-cell formula assignments"
                    .to_string(),
            ));
        }
        let (row, col) = parse_a1_cell_ref(&cell_ref)
            .ok_or_else(|| CliError::InvalidCellReference(cell_ref.clone()))?;
        let formula = if raw_formula.trim_start().starts_with('=') {
            raw_formula
        } else {
            format!("={raw_formula}")
        };
        assignments.push(TxAssignment::Formula {
            cell_ref,
            row,
            col,
            formula,
        });
    }

    if assignments.is_empty() {
        return Err(CliError::InvalidTxSaveArgs(
            "tx-save requires at least one mutation via --cell/--value, --set, or --setf"
                .to_string(),
        ));
    }

    let mut txn = workbook.begin_txn(sink.as_mut(), &trace)?;
    for assignment in &assignments {
        match assignment {
            TxAssignment::Value {
                row, col, value, ..
            } => txn.apply(Mutation::SetCellValue {
                sheet: sheet.to_string(),
                row: *row,
                col: *col,
                value: value.clone(),
            }),
            TxAssignment::Formula {
                row, col, formula, ..
            } => txn.apply(Mutation::SetCellFormula {
                sheet: sheet.to_string(),
                row: *row,
                col: *col,
                formula: formula.clone(),
                cached_value: CellValue::Empty,
            }),
        }
    }
    let commit = txn.commit(&mut workbook, sink.as_mut(), &trace)?;
    let changed_sheets = commit.changed_cells.keys().cloned().collect::<Vec<_>>();
    let mut recalc_totals = RecalcReport {
        sheet: "all_changed_sheets".to_string(),
        evaluated_cells: 0,
        cycle_count: 0,
        parse_error_count: 0,
    };
    for changed_sheet in &changed_sheets {
        let changed_roots = commit
            .changed_cells
            .get(changed_sheet)
            .map(|cells| cells.as_slice())
            .unwrap_or(&[]);
        let report = recalc_sheet_from_roots(
            &mut workbook,
            changed_sheet,
            changed_roots,
            sink.as_mut(),
            &trace,
        )?;
        recalc_totals.evaluated_cells += report.evaluated_cells;
        recalc_totals.cycle_count += report.cycle_count;
        recalc_totals.parse_error_count += report.parse_error_count;
    }

    let report = match mode {
        CliSaveMode::Preserve => preserve_xlsx_with_sheet_overrides(
            input,
            output,
            &workbook,
            &changed_sheets,
            sink.as_mut(),
            &trace,
        )?,
        CliSaveMode::Normalize => save_workbook_model(
            &workbook,
            output,
            SaveMode::Normalize,
            sink.as_mut(),
            &trace,
        )?,
    };

    println!("Input: {}", input.display());
    println!("Output: {}", report.output_path.display());
    println!("Applied mutations: {}", assignments.len());
    for assignment in assignments.iter().take(5) {
        match assignment {
            TxAssignment::Value {
                cell_ref, value, ..
            } => println!("  - {}!{} = {:?}", sheet, cell_ref, value),
            TxAssignment::Formula {
                cell_ref, formula, ..
            } => println!("  - {}!{} = {}", sheet, cell_ref, formula),
        }
    }
    if assignments.len() > 5 {
        println!("  - ... {} additional mutations", assignments.len() - 5);
    }
    println!(
        "Transaction: {} (changed sheets: {})",
        commit.txn_id,
        changed_sheets.len()
    );
    println!(
        "Saved workbook: mode={:?}, sheets={}, cells={}, copied_bytes={}",
        report.mode, report.sheet_count, report.cell_count, report.copied_bytes
    );
    println!(
        "Part graph: nodes={}, edges={}, dangling_edges={}, unknown_parts={}",
        report.part_graph.node_count,
        report.part_graph.edge_count,
        report.part_graph.dangling_edge_count,
        report.part_graph.unknown_part_count
    );
    println!(
        "Part graph flags: strategy={}, source_graph_reused={}, relationships_preserved={}, unknown_parts_preserved={}",
        report.part_graph_flags.strategy,
        report.part_graph_flags.source_graph_reused,
        report.part_graph_flags.relationships_preserved,
        report.part_graph_flags.unknown_parts_preserved
    );
    println!(
        "Post-mutation recalc: evaluated_cells={}, cycles={}, parse_errors={}",
        recalc_totals.evaluated_cells, recalc_totals.cycle_count, recalc_totals.parse_error_count
    );
    Ok(())
}

fn run_recalc(
    file: &Path,
    sheet: Option<&String>,
    report_override: Option<&PathBuf>,
    dep_graph_report_override: Option<&PathBuf>,
    dag_timing_report_override: Option<&PathBuf>,
    dag_slow_threshold_us: Option<u64>,
    jsonl_path: Option<&PathBuf>,
) -> Result<(), CliError> {
    if dag_slow_threshold_us.is_some() && dag_timing_report_override.is_none() {
        return Err(CliError::InvalidRecalcArgs(
            "--dag-slow-threshold-us requires --dag-timing-report".to_string(),
        ));
    }

    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;
    let report_payload = build_recalc_report_payload(file, sheet, sink.as_mut(), &trace)?;

    let report_path = report_override
        .cloned()
        .unwrap_or_else(|| default_recalc_report_path(file));
    let output_payload = json!({
        "generated_at": chrono::Utc::now(),
        "report": report_payload
    });
    fs::write(&report_path, to_string_pretty(&output_payload)?)?;

    let mut dep_graph_path_written = None::<PathBuf>;
    if let Some(dep_graph_path) = dep_graph_report_override {
        let graph_payload = build_dependency_graph_payload(file, sheet, sink.as_mut(), &trace)?;
        let output_graph_payload = json!({
            "generated_at": chrono::Utc::now(),
            "dependency_graph_report": graph_payload
        });
        let output_graph_json = to_string_pretty(&output_graph_payload)?;
        fs::write(dep_graph_path, output_graph_json.as_bytes())?;
        dep_graph_path_written = Some(dep_graph_path.to_path_buf());
        sink.emit(
            EventEnvelope::info("artifact.recalc.dep_graph.output", &trace)
                .with_context(json!({
                    "workbook": file.display().to_string(),
                    "output": dep_graph_path.display().to_string(),
                }))
                .with_metrics(json!({
                    "output_bytes": output_graph_json.len(),
                })),
        )?;
    }

    let mut dag_timing_path_written = None::<PathBuf>;
    if let Some(dag_timing_path) = dag_timing_report_override {
        let dag_payload =
            build_dag_timing_payload(file, sheet, dag_slow_threshold_us, sink.as_mut(), &trace)?;
        let output_dag_payload = json!({
            "generated_at": chrono::Utc::now(),
            "recalc_dag_timing_report": dag_payload
        });
        let output_dag_json = to_string_pretty(&output_dag_payload)?;
        fs::write(dag_timing_path, output_dag_json.as_bytes())?;
        dag_timing_path_written = Some(dag_timing_path.to_path_buf());
        sink.emit(
            EventEnvelope::info("artifact.recalc.dag_timing.output", &trace)
                .with_context(json!({
                    "workbook": file.display().to_string(),
                    "output": dag_timing_path.display().to_string(),
                    "slow_nodes_threshold_us": dag_slow_threshold_us,
                }))
                .with_metrics(json!({
                    "output_bytes": output_dag_json.len(),
                })),
        )?;
    }

    let sheet_count = report_payload["sheet_count"].as_u64().unwrap_or(0);
    let total_evaluated = report_payload["totals"]["evaluated_cells"]
        .as_u64()
        .unwrap_or(0);
    let total_cycles = report_payload["totals"]["cycle_count"]
        .as_u64()
        .unwrap_or(0);
    let total_parse_errors = report_payload["totals"]["parse_error_count"]
        .as_u64()
        .unwrap_or(0);

    println!("Workbook: {}", file.display());
    println!("Report: {}", report_path.display());
    if let Some(dep_graph_path) = dep_graph_path_written {
        println!("Dependency graph report: {}", dep_graph_path.display());
    }
    if let Some(dag_timing_path) = dag_timing_path_written {
        println!("DAG timing report: {}", dag_timing_path.display());
        if let Some(threshold) = dag_slow_threshold_us {
            println!("DAG slow-node threshold: {}us", threshold);
        }
    }
    println!("Sheets recalculated: {}", sheet_count);
    println!(
        "Totals: evaluated_cells={}, cycle_count={}, parse_errors={}",
        total_evaluated, total_cycles, total_parse_errors
    );

    Ok(())
}

fn run_repro_record(
    file: &Path,
    bundle_dir: &Path,
    sheet: Option<&String>,
    jsonl_path: Option<&PathBuf>,
) -> Result<(), CliError> {
    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;
    let bundle_id = format!(
        "bundle-{}",
        chrono::Utc::now().timestamp_nanos_opt().unwrap_or_default()
    );

    fs::create_dir_all(bundle_dir)?;
    let input_copy = bundle_dir.join("input.xlsx");
    let report_file = bundle_dir.join("recalc-report.json");
    let manifest_file = bundle_dir.join("manifest.json");

    sink.emit(
        EventEnvelope::info("artifact.bundle.record.start", &trace).with_context(json!({
            "bundle_id": bundle_id,
            "bundle_dir": bundle_dir.display().to_string(),
        })),
    )?;

    fs::copy(file, &input_copy)?;
    let report_payload = build_recalc_report_payload(&input_copy, sheet, sink.as_mut(), &trace)?;
    let report_bytes = serde_json::to_vec_pretty(&report_payload)?;
    fs::write(&report_file, &report_bytes)?;

    let input_bytes = fs::read(&input_copy)?;
    let manifest = ReproManifest {
        bundle_version: "1".to_string(),
        bundle_id: bundle_id.clone(),
        created_at: chrono::Utc::now(),
        input_file: "input.xlsx".to_string(),
        recalc_report_file: "recalc-report.json".to_string(),
        sheet: sheet.cloned(),
        hashes: ReproHashes {
            input_fnv64: fnv1a_64_hex(&input_bytes),
            recalc_report_fnv64: fnv1a_64_hex(&report_bytes),
        },
    };
    fs::write(&manifest_file, serde_json::to_vec_pretty(&manifest)?)?;

    sink.emit(
        EventEnvelope::info("artifact.bundle.record.end", &trace).with_context(json!({
            "bundle_id": bundle_id,
            "bundle_dir": bundle_dir.display().to_string(),
        })),
    )?;

    println!("Bundle recorded: {}", bundle_dir.display());
    println!("Manifest: {}", manifest_file.display());
    println!("Input hash (fnv64): {}", manifest.hashes.input_fnv64);
    println!(
        "Recalc report hash (fnv64): {}",
        manifest.hashes.recalc_report_fnv64
    );

    Ok(())
}

fn run_repro_check(
    bundle_dir: &Path,
    against: Option<&PathBuf>,
    jsonl_path: Option<&PathBuf>,
) -> Result<(), CliError> {
    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;
    let manifest_file = bundle_dir.join("manifest.json");
    if !manifest_file.exists() {
        return Err(CliError::MissingBundleFile(
            manifest_file.display().to_string(),
        ));
    }

    sink.emit(
        EventEnvelope::info("artifact.bundle.check.start", &trace).with_context(json!({
            "bundle_dir": bundle_dir.display().to_string(),
            "against": against.map(|p| p.display().to_string()),
        })),
    )?;

    let manifest: ReproManifest = serde_json::from_slice(&fs::read(&manifest_file)?)?;
    let input_file = bundle_dir.join(&manifest.input_file);
    let report_file = bundle_dir.join(&manifest.recalc_report_file);

    if !input_file.exists() {
        return Err(CliError::MissingBundleFile(
            input_file.display().to_string(),
        ));
    }
    if !report_file.exists() {
        return Err(CliError::MissingBundleFile(
            report_file.display().to_string(),
        ));
    }

    let input_bytes = fs::read(&input_file)?;
    let report_bytes = fs::read(&report_file)?;
    let input_hash = fnv1a_64_hex(&input_bytes);
    let report_hash = fnv1a_64_hex(&report_bytes);
    if input_hash != manifest.hashes.input_fnv64 {
        return Err(CliError::ReproMismatch(format!(
            "input hash mismatch (expected {}, got {})",
            manifest.hashes.input_fnv64, input_hash
        )));
    }
    if report_hash != manifest.hashes.recalc_report_fnv64 {
        return Err(CliError::ReproMismatch(format!(
            "report hash mismatch (expected {}, got {})",
            manifest.hashes.recalc_report_fnv64, report_hash
        )));
    }

    let compare_file = against.cloned().unwrap_or(input_file.clone());
    if !compare_file.exists() {
        return Err(CliError::MissingBundleFile(
            compare_file.display().to_string(),
        ));
    }
    let recomputed = build_recalc_report_payload(
        &compare_file,
        manifest.sheet.as_ref(),
        sink.as_mut(),
        &trace,
    )?;
    let recomputed_bytes = serde_json::to_vec_pretty(&recomputed)?;
    let recomputed_hash = fnv1a_64_hex(&recomputed_bytes);
    if recomputed_hash != manifest.hashes.recalc_report_fnv64 {
        return Err(CliError::ReproMismatch(format!(
            "recomputed report hash mismatch for '{}' (expected {}, got {})",
            compare_file.display(),
            manifest.hashes.recalc_report_fnv64,
            recomputed_hash
        )));
    }

    sink.emit(
        EventEnvelope::info("artifact.bundle.check.end", &trace).with_context(json!({
            "bundle_id": manifest.bundle_id,
            "bundle_dir": bundle_dir.display().to_string(),
            "status": "ok",
        })),
    )?;

    println!("Bundle check: OK");
    println!("Bundle: {}", bundle_dir.display());
    println!("Bundle input hash (fnv64): {}", input_hash);
    println!("Bundle report hash (fnv64): {}", report_hash);
    if let Some(path) = against {
        let against_bytes = fs::read(path)?;
        println!(
            "Against workbook hash (fnv64): {}",
            fnv1a_64_hex(&against_bytes)
        );
        println!("Against workbook: {}", path.display());
    }

    Ok(())
}

fn run_repro_diff(
    bundle_dir: &Path,
    against: &Path,
    format: CliOutputFormat,
    output_path: Option<&PathBuf>,
    limit: usize,
    jsonl_path: Option<&PathBuf>,
) -> Result<(), CliError> {
    let trace = TraceContext::root();
    let mut sink = make_sink(jsonl_path)?;
    let manifest_file = bundle_dir.join("manifest.json");
    if !manifest_file.exists() {
        return Err(CliError::MissingBundleFile(
            manifest_file.display().to_string(),
        ));
    }
    if !against.exists() {
        return Err(CliError::MissingBundleFile(against.display().to_string()));
    }

    let manifest: ReproManifest = serde_json::from_slice(&fs::read(&manifest_file)?)?;
    let baseline_input = bundle_dir.join(&manifest.input_file);
    if !baseline_input.exists() {
        return Err(CliError::MissingBundleFile(
            baseline_input.display().to_string(),
        ));
    }

    sink.emit(
        EventEnvelope::info("artifact.bundle.diff.start", &trace).with_context(json!({
            "bundle_dir": bundle_dir.display().to_string(),
            "against": against.display().to_string(),
            "format": format.as_str(),
            "output": output_path.map(|path| path.display().to_string()),
            "limit": limit,
        })),
    )?;

    let baseline = build_recalc_state(
        &baseline_input,
        manifest.sheet.as_ref(),
        sink.as_mut(),
        &trace,
    )?;
    let candidate = build_recalc_state(against, manifest.sheet.as_ref(), sink.as_mut(), &trace)?;

    let baseline_keys = baseline.cell_map.keys().cloned().collect::<BTreeSet<_>>();
    let candidate_keys = candidate.cell_map.keys().cloned().collect::<BTreeSet<_>>();
    let all_keys = baseline_keys
        .union(&candidate_keys)
        .cloned()
        .collect::<Vec<_>>();

    let mut changed = Vec::<(String, String, String)>::new();
    let mut added = Vec::<(String, String)>::new();
    let mut removed = Vec::<(String, String)>::new();
    for key in &all_keys {
        match (baseline.cell_map.get(key), candidate.cell_map.get(key)) {
            (Some(left), Some(right)) => {
                if left != right {
                    changed.push((key.clone(), left.clone(), right.clone()));
                }
            }
            (None, Some(right)) => added.push((key.clone(), right.clone())),
            (Some(left), None) => removed.push((key.clone(), left.clone())),
            (None, None) => {}
        }
    }

    let max = limit.max(1);
    let baseline_fp = baseline.payload["value_fingerprint_fnv64"]
        .as_str()
        .unwrap_or("unknown");
    let against_fp = candidate.payload["value_fingerprint_fnv64"]
        .as_str()
        .unwrap_or("unknown");

    let rendered = match format {
        CliOutputFormat::Json => serde_json::to_string_pretty(&build_repro_diff_json_payload(
            bundle_dir,
            against,
            baseline_fp,
            against_fp,
            &changed,
            &added,
            &removed,
            max,
        ))?,
        CliOutputFormat::Text => render_repro_diff_text(
            bundle_dir,
            against,
            baseline_fp,
            against_fp,
            &changed,
            &added,
            &removed,
            max,
        ),
    };

    let mut output_bytes = None::<usize>;
    if let Some(path) = output_path {
        fs::write(path, rendered.as_bytes())?;
        output_bytes = Some(rendered.len());
        println!("Diff artifact: {}", path.display());
        println!(
            "Diff summary: changed={}, added={}, removed={}",
            changed.len(),
            added.len(),
            removed.len()
        );
        sink.emit(
            EventEnvelope::info("artifact.bundle.diff.output", &trace)
                .with_context(json!({
                    "bundle_dir": bundle_dir.display().to_string(),
                    "against": against.display().to_string(),
                    "format": format.as_str(),
                    "output": path.display().to_string(),
                }))
                .with_metrics(json!({
                    "output_bytes": rendered.len(),
                })),
        )?;
    } else {
        println!("{rendered}");
    }

    sink.emit(
        EventEnvelope::info("artifact.bundle.diff.end", &trace)
            .with_context(json!({
                "bundle_dir": bundle_dir.display().to_string(),
                "against": against.display().to_string(),
                "format": format.as_str(),
                "output": output_path.map(|path| path.display().to_string()),
            }))
            .with_metrics(json!({
                "changed_cells": changed.len(),
                "added_cells": added.len(),
                "removed_cells": removed.len(),
                "output_bytes": output_bytes.unwrap_or(0),
            })),
    )?;

    Ok(())
}

fn build_repro_diff_json_payload(
    bundle_dir: &Path,
    against: &Path,
    baseline_fp: &str,
    against_fp: &str,
    changed: &[(String, String, String)],
    added: &[(String, String)],
    removed: &[(String, String)],
    max: usize,
) -> serde_json::Value {
    let changed_items = changed
        .iter()
        .take(max)
        .map(|(cell, baseline, against_value)| {
            json!({
                "cell": cell,
                "baseline": baseline,
                "against": against_value,
            })
        })
        .collect::<Vec<_>>();
    let added_items = added
        .iter()
        .take(max)
        .map(|(cell, against_value)| {
            json!({
                "cell": cell,
                "against": against_value,
            })
        })
        .collect::<Vec<_>>();
    let removed_items = removed
        .iter()
        .take(max)
        .map(|(cell, baseline)| {
            json!({
                "cell": cell,
                "baseline": baseline,
            })
        })
        .collect::<Vec<_>>();

    json!({
        "bundle": bundle_dir.display().to_string(),
        "against": against.display().to_string(),
        "baseline_value_fingerprint_fnv64": baseline_fp,
        "against_value_fingerprint_fnv64": against_fp,
        "summary": {
            "changed": changed.len(),
            "added": added.len(),
            "removed": removed.len(),
            "limit": max
        },
        "changed": changed_items,
        "added": added_items,
        "removed": removed_items,
        "truncated": {
            "changed": changed.len() > max,
            "added": added.len() > max,
            "removed": removed.len() > max
        }
    })
}

fn render_repro_diff_text(
    bundle_dir: &Path,
    against: &Path,
    baseline_fp: &str,
    against_fp: &str,
    changed: &[(String, String, String)],
    added: &[(String, String)],
    removed: &[(String, String)],
    max: usize,
) -> String {
    let mut lines = Vec::<String>::new();
    lines.push(format!("Bundle: {}", bundle_dir.display()));
    lines.push(format!("Against: {}", against.display()));
    lines.push(format!("Baseline value fingerprint: {baseline_fp}"));
    lines.push(format!("Against value fingerprint: {against_fp}"));
    lines.push(format!(
        "Diff summary: changed={}, added={}, removed={}",
        changed.len(),
        added.len(),
        removed.len()
    ));

    if !changed.is_empty() {
        lines.push(format!("Changed cells (up to {}):", max));
        for (cell, left, right) in changed.iter().take(max) {
            lines.push(format!("  - {}: {} -> {}", cell, left, right));
        }
    }
    if !added.is_empty() {
        lines.push(format!("Added cells (up to {}):", max));
        for (cell, right) in added.iter().take(max) {
            lines.push(format!("  - {}: {}", cell, right));
        }
    }
    if !removed.is_empty() {
        lines.push(format!("Removed cells (up to {}):", max));
        for (cell, left) in removed.iter().take(max) {
            lines.push(format!("  - {}: {}", cell, left));
        }
    }
    if changed.is_empty() && added.is_empty() && removed.is_empty() {
        lines.push("No cell-level differences detected.".to_string());
    }

    lines.join("\n")
}

fn build_recalc_report_payload(
    file: &Path,
    sheet: Option<&String>,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<serde_json::Value, CliError> {
    Ok(build_recalc_state(file, sheet, sink, trace)?.payload)
}

fn build_dependency_graph_payload(
    file: &Path,
    sheet: Option<&String>,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<serde_json::Value, CliError> {
    let workbook = load_workbook_model(file, sink, trace)?;
    let target_sheets = select_target_sheets(&workbook, sheet)?;
    let mut graphs = Vec::new();
    for sheet_name in &target_sheets {
        graphs.push(analyze_sheet_dependencies(
            &workbook, sheet_name, sink, trace,
        )?);
    }
    Ok(json!({
        "sheet_count": target_sheets.len(),
        "graphs": graphs,
    }))
}

fn build_dag_timing_payload(
    file: &Path,
    sheet: Option<&String>,
    slow_nodes_threshold_us: Option<u64>,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<serde_json::Value, CliError> {
    let mut workbook = load_workbook_model(file, sink, trace)?;
    let target_sheets = select_target_sheets(&workbook, sheet)?;
    let options = RecalcDagTimingOptions {
        slow_nodes_threshold_us,
    };
    let mut timings = Vec::new();
    for sheet_name in &target_sheets {
        let (_report, dag) =
            recalc_sheet_with_dag_timing_options(&mut workbook, sheet_name, options, sink, trace)?;
        timings.push(dag);
    }
    Ok(json!({
        "sheet_count": target_sheets.len(),
        "slow_nodes_threshold_us": slow_nodes_threshold_us,
        "dag_timings": timings,
    }))
}

fn build_recalc_state(
    file: &Path,
    sheet: Option<&String>,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<RecalcState, CliError> {
    let mut workbook = load_workbook_model(file, sink, trace)?;
    let target_sheets = select_target_sheets(&workbook, sheet)?;
    let mut reports = Vec::<RecalcReport>::new();
    for sheet_name in &target_sheets {
        reports.push(recalc_sheet(&mut workbook, sheet_name, sink, trace)?);
    }

    let total_evaluated = reports.iter().map(|r| r.evaluated_cells).sum::<usize>();
    let total_cycles = reports.iter().map(|r| r.cycle_count).sum::<usize>();
    let total_parse_errors = reports.iter().map(|r| r.parse_error_count).sum::<usize>();
    let mut snapshot = workbook.snapshot_json();
    if let Some(object) = snapshot.as_object_mut() {
        object.remove("workbook_id");
    }
    let value_fingerprint_fnv64 = workbook_value_fingerprint_fnv64(&workbook);

    let payload = json!({
        "sheet_count": target_sheets.len(),
        "totals": {
            "evaluated_cells": total_evaluated,
            "cycle_count": total_cycles,
            "parse_error_count": total_parse_errors,
        },
        "value_fingerprint_fnv64": value_fingerprint_fnv64,
        "reports": reports,
        "workbook_snapshot": snapshot,
    });
    let cell_map = workbook_cell_map(&workbook, &target_sheets);

    Ok(RecalcState { payload, cell_map })
}

fn select_target_sheets(
    workbook: &Workbook,
    sheet: Option<&String>,
) -> Result<Vec<String>, CliError> {
    let target_sheets = if let Some(sheet_name) = sheet {
        if !workbook.sheets.contains_key(sheet_name) {
            return Err(CliError::SheetNotFound(sheet_name.clone()));
        }
        vec![sheet_name.clone()]
    } else {
        let mut names = workbook.sheets.keys().cloned().collect::<Vec<_>>();
        names.sort();
        names
    };

    if target_sheets.is_empty() {
        return Err(CliError::NoSheets);
    }

    Ok(target_sheets)
}

fn parse_cli_cell_value(value: &str) -> CellValue {
    let trimmed = value.trim();
    if let Ok(n) = trimmed.parse::<f64>() {
        return CellValue::Number(n);
    }

    if trimmed.eq_ignore_ascii_case("true") {
        return CellValue::Bool(true);
    }
    if trimmed.eq_ignore_ascii_case("false") {
        return CellValue::Bool(false);
    }

    CellValue::Text(value.to_string())
}

fn parse_set_assignment(input: &str) -> Result<(String, String), CliError> {
    let (cell, value) = input
        .split_once('=')
        .ok_or_else(|| CliError::InvalidMutationAssignment(input.to_string()))?;
    let cell = cell.trim().to_ascii_uppercase();
    if cell.is_empty() {
        return Err(CliError::InvalidMutationAssignment(input.to_string()));
    }
    Ok((cell, value.to_string()))
}

fn parse_range_ref(input: &str) -> Option<((u32, u32), (u32, u32))> {
    let (start, end) = input.split_once(':')?;
    let (start_row, start_col) = parse_a1_cell_ref(start.trim())?;
    let (end_row, end_col) = parse_a1_cell_ref(end.trim())?;
    let row_min = start_row.min(end_row);
    let row_max = start_row.max(end_row);
    let col_min = start_col.min(end_col);
    let col_max = start_col.max(end_col);
    Some(((row_min, col_min), (row_max, col_max)))
}

fn parse_a1_cell_ref(input: &str) -> Option<(u32, u32)> {
    let mut col_part = String::new();
    let mut row_part = String::new();
    for ch in input.chars() {
        if ch.is_ascii_alphabetic() {
            if !row_part.is_empty() {
                return None;
            }
            col_part.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            row_part.push(ch);
        } else {
            return None;
        }
    }

    if col_part.is_empty() || row_part.is_empty() {
        return None;
    }

    let mut col: u32 = 0;
    for ch in col_part.chars() {
        let next = (ch as u32).checked_sub('A' as u32)? + 1;
        col = col.checked_mul(26)?.checked_add(next)?;
    }
    let row = row_part.parse::<u32>().ok()?;
    if row == 0 || col == 0 {
        return None;
    }
    Some((row, col))
}

fn workbook_value_fingerprint_fnv64(workbook: &Workbook) -> String {
    let mut canonical = String::new();
    for (sheet_name, sheet) in &workbook.sheets {
        canonical.push_str(sheet_name);
        canonical.push('\n');
        for (cell_ref, cell) in &sheet.cells {
            canonical.push_str(&format!("R{}C{}|", cell_ref.row, cell_ref.col));
            if let Some(formula) = &cell.formula {
                canonical.push_str("F:");
                canonical.push_str(formula);
                canonical.push('|');
            }
            canonical.push_str("V:");
            canonical.push_str(&canonical_cell_value(&cell.value));
            canonical.push('\n');
        }
    }
    fnv1a_64_hex(canonical.as_bytes())
}

fn canonical_cell_value(value: &CellValue) -> String {
    match value {
        CellValue::Number(n) => format!("number:{n:.15}"),
        CellValue::Text(t) => format!("text:{t}"),
        CellValue::Bool(b) => format!("bool:{}", if *b { "1" } else { "0" }),
        CellValue::Error(e) => format!("error:{e}"),
        CellValue::Empty => "empty".to_string(),
    }
}

fn workbook_cell_map(workbook: &Workbook, target_sheets: &[String]) -> BTreeMap<String, String> {
    let mut out = BTreeMap::new();
    for sheet_name in target_sheets {
        let Some(sheet) = workbook.sheets.get(sheet_name) else {
            continue;
        };
        for (cell_ref, cell) in &sheet.cells {
            let a1 = to_a1(cell_ref.row, cell_ref.col);
            let key = format!("{}!{}", sheet_name, a1);
            let value = if let Some(formula) = &cell.formula {
                format!("{}|{}", formula, canonical_cell_value(&cell.value))
            } else {
                canonical_cell_value(&cell.value)
            };
            out.insert(key, value);
        }
    }
    out
}

fn to_a1(row: u32, col: u32) -> String {
    let mut col_num = col;
    let mut letters = Vec::new();
    while col_num > 0 {
        let rem = ((col_num - 1) % 26) as u8;
        letters.push((b'A' + rem) as char);
        col_num = (col_num - 1) / 26;
    }
    letters.reverse();
    format!("{}{}", letters.into_iter().collect::<String>(), row)
}

fn fnv1a_64_hex(bytes: &[u8]) -> String {
    let mut hash: u64 = 0xcbf29ce484222325;
    for b in bytes {
        hash ^= *b as u64;
        hash = hash.wrapping_mul(0x100000001b3);
    }
    format!("{hash:016x}")
}

fn make_sink(jsonl_path: Option<&PathBuf>) -> Result<Box<dyn EventSink>, CliError> {
    if let Some(path) = jsonl_path {
        return Ok(Box::new(JsonlEventSink::new(path)?));
    }
    Ok(Box::new(NoopEventSink))
}

fn collect_xlsx_files_recursive(dir: &Path) -> Result<Vec<PathBuf>, CliError> {
    let mut files = Vec::<PathBuf>::new();
    let mut stack = vec![dir.to_path_buf()];
    while let Some(current) = stack.pop() {
        for entry in fs::read_dir(&current)? {
            let entry = entry?;
            let path = entry.path();
            let file_type = entry.file_type()?;
            if file_type.is_dir() {
                stack.push(path);
                continue;
            }
            if !file_type.is_file() {
                continue;
            }
            let is_xlsx = path
                .extension()
                .and_then(|x| x.to_str())
                .map(|ext| ext.eq_ignore_ascii_case("xlsx"))
                .unwrap_or(false);
            if is_xlsx {
                files.push(path);
            }
        }
    }
    files.sort_by(|a, b| a.to_string_lossy().cmp(&b.to_string_lossy()));
    Ok(files)
}

fn default_report_path(file: &Path) -> PathBuf {
    let mut report_path = file.to_path_buf();
    let ext = report_path
        .extension()
        .and_then(|x| x.to_str())
        .unwrap_or_default();
    if !ext.is_empty() {
        report_path.set_extension(format!("{ext}.rootcellar-report.json"));
    } else {
        report_path.set_extension("rootcellar-report.json");
    }
    report_path
}

fn default_part_graph_corpus_report_path(dir: &Path) -> PathBuf {
    dir.join("rootcellar-part-graph-corpus-report.json")
}

fn default_batch_recalc_report_path(dir: &Path) -> PathBuf {
    dir.join("rootcellar-batch-recalc-report.json")
}

fn default_bench_recalc_report_path() -> PathBuf {
    PathBuf::from("rootcellar-bench-recalc-report.json")
}

fn default_batch_threads() -> usize {
    std::thread::available_parallelism()
        .map(|v| v.get())
        .unwrap_or(1)
}

fn compute_files_per_second(processed_files: usize, wall_clock_duration_ms: u128) -> f64 {
    if processed_files == 0 {
        return 0.0;
    }
    if wall_clock_duration_ms == 0 {
        return processed_files as f64;
    }
    let wall_seconds = wall_clock_duration_ms as f64 / 1000.0;
    processed_files as f64 / wall_seconds
}

fn compute_duration_ratio(total_duration_ms: u128, wall_clock_duration_ms: u128) -> f64 {
    if wall_clock_duration_ms == 0 {
        return 0.0;
    }
    total_duration_ms as f64 / wall_clock_duration_ms as f64
}

fn normalize_path_for_report(path: &Path) -> String {
    path.to_string_lossy().replace('\\', "/")
}

fn default_recalc_report_path(file: &Path) -> PathBuf {
    let mut report_path = file.to_path_buf();
    let ext = report_path
        .extension()
        .and_then(|x| x.to_str())
        .unwrap_or_default();
    if !ext.is_empty() {
        report_path.set_extension(format!("{ext}.recalc-report.json"));
    } else {
        report_path.set_extension("recalc-report.json");
    }
    report_path
}

fn print_open_summary(report: &XlsxInspectionReport, report_path: &Path) {
    println!("Workbook: {}", report.workbook_path.display());
    println!("Report: {}", report_path.display());
    println!("Feature Score: {}", report.summary.workbook_feature_score);
    println!(
        "Required Parts: content_types={}, workbook_xml={}, worksheets={}, styles={}, shared_strings={}",
        report.required_parts.content_types,
        report.required_parts.workbook_xml,
        report.required_parts.worksheet_count,
        report.required_parts.styles_xml,
        report.required_parts.shared_strings_xml,
    );

    if report.issues.is_empty() {
        println!("Compatibility Issues: none");
    } else {
        println!("Compatibility Issues: {}", report.issues.len());
        for issue in &report.issues {
            println!("- [{}] {} ({:?})", issue.code, issue.title, issue.status);
        }
    }

    println!(
        "Unknown parts detected: {}",
        report.summary.unknown_part_count
    );
    println!(
        "Part graph: nodes={}, edges={}, dangling_edges={}, external_edges={}",
        report.part_graph.node_count,
        report.part_graph.edge_count,
        report.part_graph.dangling_edge_count,
        report.part_graph.external_edge_count
    );
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs::{self, File};
    use std::time::{SystemTime, UNIX_EPOCH};

    #[test]
    fn parses_a1_cell_reference() {
        assert_eq!(parse_a1_cell_ref("A1"), Some((1, 1)));
        assert_eq!(parse_a1_cell_ref("c7"), Some((7, 3)));
        assert_eq!(parse_a1_cell_ref("AA10"), Some((10, 27)));
        assert_eq!(parse_a1_cell_ref("1A"), None);
        assert_eq!(parse_a1_cell_ref("A0"), None);
    }

    #[test]
    fn parses_set_assignment() {
        let parsed = parse_set_assignment("B2=hello").expect("parse");
        assert_eq!(parsed.0, "B2");
        assert_eq!(parsed.1, "hello");
        let formula = parse_set_assignment("C3==A1+B1").expect("formula parse");
        assert_eq!(formula.0, "C3");
        assert_eq!(formula.1, "=A1+B1");
        let range = parse_set_assignment("A1:B2=0").expect("range parse");
        assert_eq!(range.0, "A1:B2");
        assert_eq!(range.1, "0");
        assert!(parse_set_assignment("BAD").is_err());
    }

    #[test]
    fn parses_range_reference() {
        assert_eq!(parse_range_ref("A1:B2"), Some(((1, 1), (2, 2))));
        assert_eq!(parse_range_ref("B2:A1"), Some(((1, 1), (2, 2))));
        assert_eq!(parse_range_ref("AA10:AB12"), Some(((10, 27), (12, 28))));
        assert_eq!(parse_range_ref("A1"), None);
        assert_eq!(parse_range_ref("A1:B0"), None);
    }

    #[test]
    fn value_fingerprint_changes_on_cell_mutation() {
        let mut wb1 = Workbook::new();
        let mut wb2 = Workbook::new();
        wb1.sheets.insert(
            "Sheet1".to_string(),
            rootcellar_core::model::Sheet {
                name: "Sheet1".to_string(),
                cells: std::collections::BTreeMap::from([(
                    rootcellar_core::model::CellRef { row: 1, col: 1 },
                    rootcellar_core::model::CellRecord {
                        value: CellValue::Number(1.0),
                        formula: Some("=1".to_string()),
                    },
                )]),
            },
        );
        wb2.sheets.insert(
            "Sheet1".to_string(),
            rootcellar_core::model::Sheet {
                name: "Sheet1".to_string(),
                cells: std::collections::BTreeMap::from([(
                    rootcellar_core::model::CellRef { row: 1, col: 1 },
                    rootcellar_core::model::CellRecord {
                        value: CellValue::Number(2.0),
                        formula: Some("=2".to_string()),
                    },
                )]),
            },
        );

        let f1 = workbook_value_fingerprint_fnv64(&wb1);
        let f2 = workbook_value_fingerprint_fnv64(&wb2);
        assert_ne!(f1, f2);
    }

    #[test]
    fn repro_diff_json_payload_respects_limit_and_truncation() {
        let changed = vec![
            (
                "Sheet1!A1".to_string(),
                "number:1.000000000000000".to_string(),
                "number:2.000000000000000".to_string(),
            ),
            (
                "Sheet1!A2".to_string(),
                "number:3.000000000000000".to_string(),
                "number:4.000000000000000".to_string(),
            ),
        ];
        let added = vec![("Sheet1!B1".to_string(), "text:hello".to_string())];
        let removed = vec![("Sheet1!C1".to_string(), "bool:1".to_string())];
        let payload = build_repro_diff_json_payload(
            Path::new("./bundle"),
            Path::new("./candidate.xlsx"),
            "abc123",
            "def456",
            &changed,
            &added,
            &removed,
            1,
        );

        assert_eq!(payload["summary"]["changed"].as_u64(), Some(2));
        assert_eq!(payload["changed"].as_array().map(|a| a.len()), Some(1));
        assert_eq!(payload["added"].as_array().map(|a| a.len()), Some(1));
        assert_eq!(payload["removed"].as_array().map(|a| a.len()), Some(1));
        assert_eq!(payload["truncated"]["changed"].as_bool(), Some(true));
        assert_eq!(payload["truncated"]["added"].as_bool(), Some(false));
    }

    #[test]
    fn repro_diff_text_output_handles_no_changes() {
        let rendered = render_repro_diff_text(
            Path::new("./bundle"),
            Path::new("./candidate.xlsx"),
            "abc123",
            "abc123",
            &[],
            &[],
            &[],
            50,
        );
        assert!(rendered.contains("No cell-level differences detected."));
        assert!(rendered.contains("Diff summary: changed=0, added=0, removed=0"));
    }

    #[test]
    fn collects_xlsx_files_recursively_in_deterministic_order() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!("rootcellar-cli-test-{stamp}"));
        let nested = root.join("nested");
        fs::create_dir_all(&nested).expect("mkdir");
        let a = root.join("b.xlsx");
        let b = nested.join("a.xlsx");
        let c = nested.join("ignore.txt");
        File::create(&a).expect("a");
        File::create(&b).expect("b");
        File::create(&c).expect("c");

        let files = collect_xlsx_files_recursive(&root).expect("collect");
        assert_eq!(files.len(), 2);
        assert!(files[0].extension().and_then(|x| x.to_str()) == Some("xlsx"));
        assert!(files[1].extension().and_then(|x| x.to_str()) == Some("xlsx"));
        assert!(files.iter().any(|p| p.ends_with("b.xlsx")));
        assert!(files.iter().any(|p| p.ends_with("a.xlsx")));

        fs::remove_dir_all(&root).expect("cleanup");
    }

    #[test]
    fn corpus_report_default_path_is_in_directory() {
        let dir = Path::new("./corpus");
        let out = default_part_graph_corpus_report_path(dir);
        assert_eq!(
            out,
            PathBuf::from("./corpus").join("rootcellar-part-graph-corpus-report.json")
        );
    }

    #[test]
    fn batch_report_default_path_is_in_directory() {
        let dir = Path::new("./corpus");
        let out = default_batch_recalc_report_path(dir);
        assert_eq!(
            out,
            PathBuf::from("./corpus").join("rootcellar-batch-recalc-report.json")
        );
    }

    #[test]
    fn bench_report_default_path_is_stable() {
        assert_eq!(
            default_bench_recalc_report_path(),
            PathBuf::from("rootcellar-bench-recalc-report.json")
        );
    }

    #[test]
    fn synthetic_benchmark_workbook_shape_matches_requested_dimensions() {
        let wb = build_synthetic_recalc_benchmark_workbook(3, 4).expect("workbook");
        let sheet = wb.sheets.get("Sheet1").expect("sheet1");
        assert_eq!(sheet.cells.len(), 15);

        for row in 1..=3u32 {
            let root = sheet.cells.get(&CellRef { row, col: 1 }).expect("root");
            assert!(root.formula.is_none());
            for col in 2..=5u32 {
                let cell = sheet.cells.get(&CellRef { row, col }).expect("formula");
                assert!(cell.formula.is_some());
            }
        }
    }

    #[test]
    fn batch_detail_level_controls_payload_inclusion() {
        assert!(!CliBatchDetailLevel::Minimal.include_recalc_payload());
        assert!(CliBatchDetailLevel::Diagnostic.include_recalc_payload());
        assert!(CliBatchDetailLevel::Forensic.include_recalc_payload());
    }

    #[test]
    fn default_batch_threads_is_positive() {
        assert!(default_batch_threads() >= 1);
    }

    #[test]
    fn computes_files_per_second_from_wall_clock() {
        assert_eq!(compute_files_per_second(0, 100), 0.0);
        assert_eq!(compute_files_per_second(5, 0), 5.0);
        assert!((compute_files_per_second(10, 250) - 40.0).abs() < f64::EPSILON);
    }

    #[test]
    fn computes_duration_ratio() {
        assert_eq!(compute_duration_ratio(10, 0), 0.0);
        assert!((compute_duration_ratio(90, 30) - 3.0).abs() < f64::EPSILON);
    }

    #[test]
    fn normalizes_report_paths_with_forward_slashes() {
        let path = Path::new(r".\nested\file.xlsx");
        let normalized = normalize_path_for_report(path);
        assert_eq!(normalized, "./nested/file.xlsx");
    }
}
