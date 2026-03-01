use crate::model::{CellRef, CellValue, Workbook};
use crate::telemetry::{EventEnvelope, EventSink, TelemetryError, TraceContext};
use chrono::{Datelike, Duration, NaiveDate, NaiveTime, Timelike};
use serde::{Deserialize, Serialize};
use serde_json::json;
use std::collections::VecDeque;
use std::collections::{BTreeMap, BTreeSet};
use std::time::Instant;
use thiserror::Error;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RecalcReport {
    pub sheet: String,
    pub evaluated_cells: usize,
    pub cycle_count: usize,
    pub parse_error_count: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DependencyGraphReport {
    pub sheet: String,
    pub formula_cell_count: usize,
    pub function_call_count: usize,
    pub ast_node_count: usize,
    pub ast_unique_node_count: usize,
    pub dependency_edge_count: usize,
    pub formula_edge_count: usize,
    pub topo_order: Vec<String>,
    pub cyclic_cells: Vec<String>,
    pub parse_error_cells: Vec<String>,
    pub formula_ast_ids: BTreeMap<String, u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RecalcDagNodeTiming {
    pub cell: String,
    pub duration_us: u64,
    pub status: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RecalcDagNodeDegree {
    pub cell: String,
    pub fan_in: usize,
    pub fan_out: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RecalcDagTimingReport {
    pub sheet: String,
    pub mode: String,
    pub formula_cell_count: usize,
    pub evaluated_cells: usize,
    pub changed_root_count: Option<usize>,
    pub total_node_duration_us: u64,
    pub max_node_duration_us: u64,
    pub node_timings: Vec<RecalcDagNodeTiming>,
    pub node_degrees: Vec<RecalcDagNodeDegree>,
    pub max_fan_in: usize,
    pub max_fan_out: usize,
    pub critical_path: Vec<String>,
    pub critical_path_duration_us: u64,
    pub critical_path_truncated_by_cycles: bool,
    pub slow_nodes_threshold_us: u64,
    pub slow_nodes: Vec<RecalcDagNodeTiming>,
}

#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize)]
pub struct RecalcDagTimingOptions {
    pub slow_nodes_threshold_us: Option<u64>,
}

#[derive(Debug, Error)]
pub enum CalcError {
    #[error("sheet not found: {0}")]
    SheetNotFound(String),
    #[error("telemetry error: {0}")]
    Telemetry(#[from] TelemetryError),
}

#[derive(Debug, Clone, Error)]
enum EvalError {
    #[error("cycle detected at {0:?}")]
    Cycle(CellRef),
    #[error("parse error")]
    Parse,
    #[error("division by zero")]
    DivisionByZero,
}

#[derive(Debug, Clone)]
enum Expr {
    Number(f64),
    Cell(CellRef),
    Function {
        name: String,
        args: Vec<Expr>,
    },
    UnaryMinus(Box<Expr>),
    Binary {
        op: BinaryOp,
        left: Box<Expr>,
        right: Box<Expr>,
    },
}

#[derive(Debug, Clone, Copy)]
enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
}

impl BinaryOp {
    fn as_str(self) -> &'static str {
        match self {
            BinaryOp::Add => "add",
            BinaryOp::Sub => "sub",
            BinaryOp::Mul => "mul",
            BinaryOp::Div => "div",
        }
    }
}

#[derive(Debug, Clone)]
struct SheetDependencyAnalysis {
    parsed_formulas: BTreeMap<CellRef, Result<Expr, EvalError>>,
    dependency_refs: BTreeMap<CellRef, BTreeSet<CellRef>>,
    dependents_by_ref: BTreeMap<CellRef, BTreeSet<CellRef>>,
    formula_nodes: BTreeSet<CellRef>,
    function_call_count: usize,
    ast_node_count: usize,
    ast_unique_node_count: usize,
    formula_ast_ids: BTreeMap<CellRef, u32>,
    ast_intern_nodes: BTreeMap<u32, String>,
    dependency_edge_count: usize,
    formula_edge_count: usize,
    topo_order: Vec<CellRef>,
    cyclic_cells: Vec<CellRef>,
    parse_error_cells: Vec<CellRef>,
}

#[derive(Debug)]
struct InternalRecalcResult {
    report: RecalcReport,
    dag_timing: Option<RecalcDagTimingReport>,
}

#[derive(Debug)]
struct DagInsights {
    node_degrees: Vec<RecalcDagNodeDegree>,
    max_fan_in: usize,
    max_fan_out: usize,
    critical_path: Vec<String>,
    critical_path_duration_us: u64,
    critical_path_truncated_by_cycles: bool,
    slow_nodes_threshold_us: u64,
    slow_nodes: Vec<RecalcDagNodeTiming>,
}

#[derive(Debug, Default, Clone)]
struct AstInternPool {
    next_id: u32,
    id_by_key: BTreeMap<String, u32>,
    key_by_id: BTreeMap<u32, String>,
}

impl AstInternPool {
    fn new() -> Self {
        Self {
            next_id: 1,
            id_by_key: BTreeMap::new(),
            key_by_id: BTreeMap::new(),
        }
    }

    fn intern_expr(&mut self, expr: &Expr) -> u32 {
        let key = match expr {
            Expr::Number(n) => format!("num:{n:.15}"),
            Expr::Cell(cell) => format!("cell:r{}c{}", cell.row, cell.col),
            Expr::Function { name, args } => {
                let arg_ids = args
                    .iter()
                    .map(|arg| self.intern_expr(arg).to_string())
                    .collect::<Vec<_>>()
                    .join(",");
                format!("fn:{}:[{}]", name.to_ascii_uppercase(), arg_ids)
            }
            Expr::UnaryMinus(inner) => {
                let inner_id = self.intern_expr(inner);
                format!("neg:{inner_id}")
            }
            Expr::Binary { op, left, right } => {
                let left_id = self.intern_expr(left);
                let right_id = self.intern_expr(right);
                format!("bin:{}:{left_id}:{right_id}", op.as_str())
            }
        };
        self.intern_key(key)
    }

    fn intern_key(&mut self, key: String) -> u32 {
        if let Some(id) = self.id_by_key.get(&key) {
            return *id;
        }
        let id = self.next_id;
        self.next_id = self.next_id.saturating_add(1);
        self.id_by_key.insert(key.clone(), id);
        self.key_by_id.insert(id, key);
        id
    }

    fn unique_node_count(&self) -> usize {
        self.key_by_id.len()
    }
}

pub fn recalc_sheet(
    workbook: &mut Workbook,
    sheet_name: &str,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<RecalcReport, CalcError> {
    Ok(recalc_sheet_impl(
        workbook,
        sheet_name,
        None,
        sink,
        trace,
        "calc.recalc.start",
        "calc.recalc.end",
        false,
        RecalcDagTimingOptions::default(),
    )?
    .report)
}

pub fn recalc_sheet_with_dag_timing(
    workbook: &mut Workbook,
    sheet_name: &str,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<(RecalcReport, RecalcDagTimingReport), CalcError> {
    recalc_sheet_with_dag_timing_options(
        workbook,
        sheet_name,
        RecalcDagTimingOptions::default(),
        sink,
        trace,
    )
}

pub fn recalc_sheet_with_dag_timing_options(
    workbook: &mut Workbook,
    sheet_name: &str,
    options: RecalcDagTimingOptions,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<(RecalcReport, RecalcDagTimingReport), CalcError> {
    let result = recalc_sheet_impl(
        workbook,
        sheet_name,
        None,
        sink,
        trace,
        "calc.recalc.start",
        "calc.recalc.end",
        true,
        options,
    )?;
    let dag_timing = result
        .dag_timing
        .expect("dag timing report must be present when capture_dag_timing=true");
    Ok((result.report, dag_timing))
}

pub fn recalc_sheet_from_roots(
    workbook: &mut Workbook,
    sheet_name: &str,
    changed_roots: &[CellRef],
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<RecalcReport, CalcError> {
    let roots = changed_roots.iter().copied().collect::<BTreeSet<_>>();
    Ok(recalc_sheet_impl(
        workbook,
        sheet_name,
        Some(&roots),
        sink,
        trace,
        "calc.recalc.incremental.start",
        "calc.recalc.incremental.end",
        false,
        RecalcDagTimingOptions::default(),
    )?
    .report)
}

pub fn recalc_sheet_from_roots_with_dag_timing(
    workbook: &mut Workbook,
    sheet_name: &str,
    changed_roots: &[CellRef],
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<(RecalcReport, RecalcDagTimingReport), CalcError> {
    recalc_sheet_from_roots_with_dag_timing_options(
        workbook,
        sheet_name,
        changed_roots,
        RecalcDagTimingOptions::default(),
        sink,
        trace,
    )
}

pub fn recalc_sheet_from_roots_with_dag_timing_options(
    workbook: &mut Workbook,
    sheet_name: &str,
    changed_roots: &[CellRef],
    options: RecalcDagTimingOptions,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<(RecalcReport, RecalcDagTimingReport), CalcError> {
    let roots = changed_roots.iter().copied().collect::<BTreeSet<_>>();
    let result = recalc_sheet_impl(
        workbook,
        sheet_name,
        Some(&roots),
        sink,
        trace,
        "calc.recalc.incremental.start",
        "calc.recalc.incremental.end",
        true,
        options,
    )?;
    let dag_timing = result
        .dag_timing
        .expect("dag timing report must be present when capture_dag_timing=true");
    Ok((result.report, dag_timing))
}

fn recalc_sheet_impl(
    workbook: &mut Workbook,
    sheet_name: &str,
    changed_roots: Option<&BTreeSet<CellRef>>,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
    start_event: &str,
    end_event: &str,
    capture_dag_timing: bool,
    dag_timing_options: RecalcDagTimingOptions,
) -> Result<InternalRecalcResult, CalcError> {
    let start = Instant::now();
    let span = trace.child();
    let mode = if changed_roots.is_some() {
        "incremental"
    } else {
        "full"
    };
    let changed_roots_preview = changed_roots.map(|roots| {
        roots
            .iter()
            .take(20)
            .map(|cell| to_a1(*cell))
            .collect::<Vec<_>>()
    });

    sink.emit(
        EventEnvelope::info(start_event, &span)
            .with_workbook_id(workbook.workbook_id)
            .with_context(json!({
                "sheet": sheet_name,
                "mode": mode,
                "changed_root_count": changed_roots.map(|roots| roots.len()),
                "changed_roots_preview": changed_roots_preview,
            })),
    )?;

    let formula_cells = collect_formula_cells(workbook, sheet_name)?;
    let analysis = build_sheet_dependency_analysis(workbook, sheet_name, &formula_cells)?;
    emit_dependency_graph_event(sink, &span, workbook.workbook_id, sheet_name, &analysis)?;

    let target_formula_set = if let Some(roots) = changed_roots {
        collect_impacted_formula_cells(&analysis, roots)
    } else {
        formula_cells.iter().copied().collect::<BTreeSet<_>>()
    };
    let target_formula_cells = order_formula_cells(&analysis, &target_formula_set);

    let mut cache = BTreeMap::<CellRef, Result<f64, EvalError>>::new();
    let mut cycle_count = 0usize;
    let mut parse_error_count = 0usize;
    let mut duration_by_cell = BTreeMap::<CellRef, u64>::new();
    let mut node_timings = if capture_dag_timing {
        Some(Vec::<RecalcDagNodeTiming>::new())
    } else {
        None
    };

    for cell_ref in &target_formula_cells {
        let mut stack = BTreeSet::new();
        let eval_started = Instant::now();
        let result = eval_cell(
            workbook,
            sheet_name,
            *cell_ref,
            &analysis.parsed_formulas,
            &mut stack,
            &mut cache,
        );
        let duration_us = eval_started.elapsed().as_micros().min(u128::from(u64::MAX)) as u64;
        duration_by_cell.insert(*cell_ref, duration_us);
        if let Some(timings) = node_timings.as_mut() {
            timings.push(RecalcDagNodeTiming {
                cell: to_a1(*cell_ref),
                duration_us,
                status: eval_status(&result).to_string(),
            });
        }

        if let Some(sheet_mut) = workbook.sheets.get_mut(sheet_name) {
            if let Some(cell) = sheet_mut.cells.get_mut(cell_ref) {
                match result {
                    Ok(value) => {
                        cell.value = CellValue::Number(value);
                    }
                    Err(EvalError::Cycle(_)) => {
                        cycle_count += 1;
                        cell.value = CellValue::Error("#CYCLE!".to_string());
                        sink.emit(
                            EventEnvelope::info("calc.cycle.detected", &span)
                                .with_workbook_id(workbook.workbook_id)
                                .with_payload(json!({
                                    "sheet": sheet_name,
                                    "row": cell_ref.row,
                                    "col": cell_ref.col,
                                })),
                        )?;
                    }
                    Err(EvalError::Parse) => {
                        parse_error_count += 1;
                        cell.value = CellValue::Error("#PARSE!".to_string());
                    }
                    Err(EvalError::DivisionByZero) => {
                        cell.value = CellValue::Error("#DIV/0!".to_string());
                    }
                }
            }
        }
    }

    let report = RecalcReport {
        sheet: sheet_name.to_string(),
        evaluated_cells: target_formula_cells.len(),
        cycle_count,
        parse_error_count,
    };

    let dag_timing = if let Some(node_timings) = node_timings {
        let total_node_duration_us = node_timings.iter().map(|n| n.duration_us).sum::<u64>();
        let max_node_duration_us = node_timings
            .iter()
            .map(|n| n.duration_us)
            .max()
            .unwrap_or(0);
        let insights = build_dag_insights(
            &analysis,
            &target_formula_cells,
            &duration_by_cell,
            &node_timings,
            dag_timing_options.slow_nodes_threshold_us,
        );
        let mut slowest = node_timings.clone();
        slowest.sort_by(|a, b| {
            b.duration_us
                .cmp(&a.duration_us)
                .then_with(|| a.cell.cmp(&b.cell))
        });
        let slowest_preview = slowest
            .iter()
            .take(20)
            .map(|n| {
                json!({
                    "cell": n.cell,
                    "duration_us": n.duration_us,
                    "status": n.status,
                })
            })
            .collect::<Vec<_>>();
        let slow_nodes_preview = insights
            .slow_nodes
            .iter()
            .take(20)
            .map(|n| {
                json!({
                    "cell": n.cell,
                    "duration_us": n.duration_us,
                    "status": n.status,
                })
            })
            .collect::<Vec<_>>();
        let degree_preview = insights
            .node_degrees
            .iter()
            .take(20)
            .map(|node| {
                json!({
                    "cell": node.cell,
                    "fan_in": node.fan_in,
                    "fan_out": node.fan_out,
                })
            })
            .collect::<Vec<_>>();
        sink.emit(
            EventEnvelope::info("calc.recalc.dag_timing", &span)
                .with_workbook_id(workbook.workbook_id)
                .with_context(json!({
                    "sheet": sheet_name,
                    "mode": mode,
                    "changed_root_count": changed_roots.map(|roots| roots.len()),
                }))
                .with_metrics(json!({
                    "node_count": node_timings.len(),
                    "total_node_duration_us": total_node_duration_us,
                    "max_node_duration_us": max_node_duration_us,
                    "max_fan_in": insights.max_fan_in,
                    "max_fan_out": insights.max_fan_out,
                    "critical_path_duration_us": insights.critical_path_duration_us,
                    "critical_path_len": insights.critical_path.len(),
                    "slow_nodes_threshold_us": insights.slow_nodes_threshold_us,
                    "slow_node_count": insights.slow_nodes.len(),
                }))
                .with_payload(json!({
                    "slowest_nodes_preview": slowest_preview,
                    "slowest_preview_truncated": node_timings.len() > 20,
                    "slow_nodes_preview": slow_nodes_preview,
                    "slow_nodes_preview_truncated": insights.slow_nodes.len() > 20,
                    "node_degree_preview": degree_preview,
                    "node_degree_preview_truncated": insights.node_degrees.len() > 20,
                    "critical_path": insights.critical_path,
                    "critical_path_truncated_by_cycles": insights.critical_path_truncated_by_cycles,
                    "slow_nodes_threshold_override_us": dag_timing_options.slow_nodes_threshold_us,
                })),
        )?;
        Some(RecalcDagTimingReport {
            sheet: sheet_name.to_string(),
            mode: mode.to_string(),
            formula_cell_count: formula_cells.len(),
            evaluated_cells: report.evaluated_cells,
            changed_root_count: changed_roots.map(|roots| roots.len()),
            total_node_duration_us,
            max_node_duration_us,
            node_timings,
            node_degrees: insights.node_degrees,
            max_fan_in: insights.max_fan_in,
            max_fan_out: insights.max_fan_out,
            critical_path: insights.critical_path,
            critical_path_duration_us: insights.critical_path_duration_us,
            critical_path_truncated_by_cycles: insights.critical_path_truncated_by_cycles,
            slow_nodes_threshold_us: insights.slow_nodes_threshold_us,
            slow_nodes: insights.slow_nodes,
        })
    } else {
        None
    };

    sink.emit(
        EventEnvelope::info(end_event, &span)
            .with_workbook_id(workbook.workbook_id)
            .with_metrics(json!({
                "duration_ms": start.elapsed().as_secs_f64() * 1000.0,
                "evaluated_cells": report.evaluated_cells,
                "formula_cell_count": formula_cells.len(),
                "cycle_count": report.cycle_count,
                "parse_error_count": report.parse_error_count,
            }))
            .with_payload(json!({
                "sheet": report.sheet,
            })),
    )?;

    Ok(InternalRecalcResult { report, dag_timing })
}

pub fn analyze_sheet_dependencies(
    workbook: &Workbook,
    sheet_name: &str,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<DependencyGraphReport, CalcError> {
    let span = trace.child();
    let formula_cells = collect_formula_cells(workbook, sheet_name)?;
    let analysis = build_sheet_dependency_analysis(workbook, sheet_name, &formula_cells)?;
    emit_dependency_graph_event(sink, &span, workbook.workbook_id, sheet_name, &analysis)?;
    Ok(build_dependency_report(sheet_name, &analysis))
}

fn collect_formula_cells(workbook: &Workbook, sheet_name: &str) -> Result<Vec<CellRef>, CalcError> {
    let sheet = workbook
        .sheets
        .get(sheet_name)
        .ok_or_else(|| CalcError::SheetNotFound(sheet_name.to_string()))?;
    Ok(sheet
        .cells
        .iter()
        .filter_map(|(cell_ref, cell)| {
            if cell.formula.is_some() {
                Some(*cell_ref)
            } else {
                None
            }
        })
        .collect::<Vec<_>>())
}

fn build_sheet_dependency_analysis(
    workbook: &Workbook,
    sheet_name: &str,
    formula_cells: &[CellRef],
) -> Result<SheetDependencyAnalysis, CalcError> {
    let sheet = workbook
        .sheets
        .get(sheet_name)
        .ok_or_else(|| CalcError::SheetNotFound(sheet_name.to_string()))?;
    let formula_set = formula_cells.iter().copied().collect::<BTreeSet<_>>();

    let mut parsed_formulas = BTreeMap::<CellRef, Result<Expr, EvalError>>::new();
    let mut dependency_refs = BTreeMap::<CellRef, BTreeSet<CellRef>>::new();
    let mut dependents_by_ref = BTreeMap::<CellRef, BTreeSet<CellRef>>::new();
    let mut ast_pool = AstInternPool::new();
    let mut formula_ast_ids = BTreeMap::<CellRef, u32>::new();
    let mut parse_error_cells = Vec::<CellRef>::new();
    let mut function_call_count = 0usize;
    let mut ast_node_count = 0usize;
    let mut dependency_edge_count = 0usize;

    for cell_ref in formula_cells {
        let parsed = sheet
            .cells
            .get(cell_ref)
            .and_then(|cell| cell.formula.as_deref())
            .ok_or(EvalError::Parse)
            .and_then(parse_formula_expression);

        let refs = if let Ok(expr) = &parsed {
            let mut refs = BTreeSet::<CellRef>::new();
            collect_expr_references(expr, &mut refs);
            function_call_count += count_expr_functions(expr);
            ast_node_count += count_expr_nodes(expr);
            let root_id = ast_pool.intern_expr(expr);
            formula_ast_ids.insert(*cell_ref, root_id);
            refs
        } else {
            parse_error_cells.push(*cell_ref);
            BTreeSet::new()
        };

        dependency_edge_count += refs.len();
        for referenced in &refs {
            dependents_by_ref
                .entry(*referenced)
                .or_default()
                .insert(*cell_ref);
        }
        parsed_formulas.insert(*cell_ref, parsed);
        dependency_refs.insert(*cell_ref, refs);
    }

    let (topo_order, formula_edge_count) =
        build_formula_topological_order(&formula_set, &dependency_refs);
    let topo_set = topo_order.iter().copied().collect::<BTreeSet<_>>();
    let cyclic_cells = formula_set
        .difference(&topo_set)
        .copied()
        .collect::<Vec<_>>();

    Ok(SheetDependencyAnalysis {
        parsed_formulas,
        dependency_refs,
        dependents_by_ref,
        formula_nodes: formula_set,
        function_call_count,
        ast_node_count,
        ast_unique_node_count: ast_pool.unique_node_count(),
        formula_ast_ids,
        ast_intern_nodes: ast_pool.key_by_id,
        dependency_edge_count,
        formula_edge_count,
        topo_order,
        cyclic_cells,
        parse_error_cells,
    })
}

fn collect_impacted_formula_cells(
    analysis: &SheetDependencyAnalysis,
    changed_roots: &BTreeSet<CellRef>,
) -> BTreeSet<CellRef> {
    let mut impacted = BTreeSet::<CellRef>::new();
    let mut queue = VecDeque::<CellRef>::new();
    for root in changed_roots {
        if analysis.formula_nodes.contains(root) && impacted.insert(*root) {
            queue.push_back(*root);
        }
        if let Some(dependents) = analysis.dependents_by_ref.get(root) {
            for dependent in dependents {
                if impacted.insert(*dependent) {
                    queue.push_back(*dependent);
                }
            }
        }
    }

    while let Some(node) = queue.pop_front() {
        if let Some(dependents) = analysis.dependents_by_ref.get(&node) {
            for dependent in dependents {
                if impacted.insert(*dependent) {
                    queue.push_back(*dependent);
                }
            }
        }
    }

    impacted
}

fn order_formula_cells(
    analysis: &SheetDependencyAnalysis,
    target_cells: &BTreeSet<CellRef>,
) -> Vec<CellRef> {
    let mut ordered = analysis
        .topo_order
        .iter()
        .copied()
        .filter(|cell| target_cells.contains(cell))
        .collect::<Vec<_>>();
    let ordered_set = ordered.iter().copied().collect::<BTreeSet<_>>();
    let mut remaining = target_cells
        .difference(&ordered_set)
        .copied()
        .collect::<Vec<_>>();
    remaining.sort();
    ordered.extend(remaining);
    ordered
}

fn build_formula_topological_order(
    formula_set: &BTreeSet<CellRef>,
    dependency_refs: &BTreeMap<CellRef, BTreeSet<CellRef>>,
) -> (Vec<CellRef>, usize) {
    let mut indegree = formula_set
        .iter()
        .copied()
        .map(|cell| (cell, 0usize))
        .collect::<BTreeMap<_, _>>();
    let mut adjacency = BTreeMap::<CellRef, BTreeSet<CellRef>>::new();
    let mut formula_edge_count = 0usize;

    for (cell, refs) in dependency_refs {
        for referenced in refs {
            if formula_set.contains(referenced) {
                adjacency.entry(*referenced).or_default().insert(*cell);
                if let Some(entry) = indegree.get_mut(cell) {
                    *entry += 1;
                }
                formula_edge_count += 1;
            }
        }
    }

    let mut ready = indegree
        .iter()
        .filter_map(|(cell, degree)| if *degree == 0 { Some(*cell) } else { None })
        .collect::<BTreeSet<_>>();
    let mut order = Vec::<CellRef>::new();

    while let Some(next) = ready.iter().next().copied() {
        ready.remove(&next);
        order.push(next);

        if let Some(dependents) = adjacency.get(&next) {
            for dependent in dependents {
                if let Some(degree) = indegree.get_mut(dependent) {
                    *degree = degree.saturating_sub(1);
                    if *degree == 0 {
                        ready.insert(*dependent);
                    }
                }
            }
        }
    }

    (order, formula_edge_count)
}

fn emit_dependency_graph_event(
    sink: &mut dyn EventSink,
    span: &TraceContext,
    workbook_id: uuid::Uuid,
    sheet_name: &str,
    analysis: &SheetDependencyAnalysis,
) -> Result<(), CalcError> {
    let parse_errors = analysis
        .parse_error_cells
        .iter()
        .map(|cell| to_a1(*cell))
        .collect::<Vec<_>>();
    let cyclic_cells = analysis
        .cyclic_cells
        .iter()
        .map(|cell| to_a1(*cell))
        .collect::<Vec<_>>();
    let topo_preview = analysis
        .topo_order
        .iter()
        .take(20)
        .map(|cell| to_a1(*cell))
        .collect::<Vec<_>>();
    let ast_preview = analysis
        .ast_intern_nodes
        .iter()
        .take(20)
        .map(|(id, key)| {
            json!({
                "id": id,
                "key": key,
            })
        })
        .collect::<Vec<_>>();
    let formula_ast_ids = analysis
        .formula_ast_ids
        .iter()
        .map(|(cell, id)| {
            json!({
                "cell": to_a1(*cell),
                "ast_id": id,
            })
        })
        .collect::<Vec<_>>();

    sink.emit(
        EventEnvelope::info("calc.dependency_graph.built", span)
            .with_workbook_id(workbook_id)
            .with_context(json!({
                "sheet": sheet_name,
            }))
            .with_metrics(json!({
                "formula_cell_count": analysis.parsed_formulas.len(),
                "function_call_count": analysis.function_call_count,
                "ast_node_count": analysis.ast_node_count,
                "ast_unique_node_count": analysis.ast_unique_node_count,
                "dependency_edge_count": analysis.dependency_edge_count,
                "formula_edge_count": analysis.formula_edge_count,
                "topo_order_count": analysis.topo_order.len(),
                "cyclic_formula_count": analysis.cyclic_cells.len(),
                "parse_error_formula_count": analysis.parse_error_cells.len(),
            }))
            .with_payload(json!({
                "parse_error_cells": parse_errors,
                "cyclic_cells": cyclic_cells,
                "topo_order_preview": topo_preview,
                "topo_order_truncated": analysis.topo_order.len() > 20,
                "formula_ast_ids": formula_ast_ids,
                "ast_intern_preview": ast_preview,
                "ast_intern_preview_truncated": analysis.ast_intern_nodes.len() > 20,
                "dependency_refs": analysis
                    .dependency_refs
                    .iter()
                    .map(|(cell, refs)| {
                        json!({
                            "cell": to_a1(*cell),
                            "references": refs.iter().map(|dep| to_a1(*dep)).collect::<Vec<_>>(),
                        })
                    })
                    .collect::<Vec<_>>(),
            })),
    )?;

    Ok(())
}

fn build_dependency_report(
    sheet_name: &str,
    analysis: &SheetDependencyAnalysis,
) -> DependencyGraphReport {
    let formula_ast_ids = analysis
        .formula_ast_ids
        .iter()
        .map(|(cell, id)| (to_a1(*cell), *id))
        .collect::<BTreeMap<_, _>>();
    DependencyGraphReport {
        sheet: sheet_name.to_string(),
        formula_cell_count: analysis.parsed_formulas.len(),
        function_call_count: analysis.function_call_count,
        ast_node_count: analysis.ast_node_count,
        ast_unique_node_count: analysis.ast_unique_node_count,
        dependency_edge_count: analysis.dependency_edge_count,
        formula_edge_count: analysis.formula_edge_count,
        topo_order: analysis
            .topo_order
            .iter()
            .map(|cell| to_a1(*cell))
            .collect(),
        cyclic_cells: analysis
            .cyclic_cells
            .iter()
            .map(|cell| to_a1(*cell))
            .collect(),
        parse_error_cells: analysis
            .parse_error_cells
            .iter()
            .map(|cell| to_a1(*cell))
            .collect(),
        formula_ast_ids,
    }
}

fn build_dag_insights(
    analysis: &SheetDependencyAnalysis,
    target_formula_cells: &[CellRef],
    duration_by_cell: &BTreeMap<CellRef, u64>,
    node_timings: &[RecalcDagNodeTiming],
    slow_nodes_threshold_override_us: Option<u64>,
) -> DagInsights {
    let target_set = target_formula_cells
        .iter()
        .copied()
        .collect::<BTreeSet<_>>();
    let mut fan_in = target_formula_cells
        .iter()
        .copied()
        .map(|cell| (cell, 0usize))
        .collect::<BTreeMap<_, _>>();
    let mut fan_out = target_formula_cells
        .iter()
        .copied()
        .map(|cell| (cell, 0usize))
        .collect::<BTreeMap<_, _>>();
    let mut adjacency = BTreeMap::<CellRef, BTreeSet<CellRef>>::new();

    for referenced in target_formula_cells {
        if !analysis.formula_nodes.contains(referenced) {
            continue;
        }
        if let Some(dependents) = analysis.dependents_by_ref.get(referenced) {
            for dependent in dependents {
                if target_set.contains(dependent) {
                    *fan_in.entry(*dependent).or_insert(0) += 1;
                    *fan_out.entry(*referenced).or_insert(0) += 1;
                    adjacency.entry(*referenced).or_default().insert(*dependent);
                }
            }
        }
    }

    let node_degrees = target_formula_cells
        .iter()
        .map(|cell| RecalcDagNodeDegree {
            cell: to_a1(*cell),
            fan_in: *fan_in.get(cell).unwrap_or(&0),
            fan_out: *fan_out.get(cell).unwrap_or(&0),
        })
        .collect::<Vec<_>>();
    let max_fan_in = node_degrees.iter().map(|n| n.fan_in).max().unwrap_or(0);
    let max_fan_out = node_degrees.iter().map(|n| n.fan_out).max().unwrap_or(0);

    let topo_nodes = analysis
        .topo_order
        .iter()
        .copied()
        .filter(|cell| target_set.contains(cell))
        .collect::<Vec<_>>();
    let critical_path_truncated_by_cycles = topo_nodes.len() < target_set.len();
    let mut best_duration = BTreeMap::<CellRef, u64>::new();
    let mut predecessor = BTreeMap::<CellRef, CellRef>::new();
    for node in &topo_nodes {
        let self_duration = *duration_by_cell.get(node).unwrap_or(&0);
        best_duration.entry(*node).or_insert(self_duration);
        if let Some(dependents) = adjacency.get(node) {
            for dependent in dependents {
                let dependent_duration = *duration_by_cell.get(dependent).unwrap_or(&0);
                let candidate = best_duration.get(node).copied().unwrap_or(0) + dependent_duration;
                let current = best_duration
                    .get(dependent)
                    .copied()
                    .unwrap_or(dependent_duration);
                let should_update = if candidate > current {
                    true
                } else if candidate == current {
                    predecessor
                        .get(dependent)
                        .map(|existing| node < existing)
                        .unwrap_or(true)
                } else {
                    false
                };
                if should_update {
                    best_duration.insert(*dependent, candidate);
                    predecessor.insert(*dependent, *node);
                }
            }
        }
    }

    let critical_end = if !topo_nodes.is_empty() {
        topo_nodes
            .iter()
            .max_by(|a, b| {
                best_duration
                    .get(a)
                    .copied()
                    .unwrap_or(0)
                    .cmp(&best_duration.get(b).copied().unwrap_or(0))
                    .then_with(|| b.cmp(a))
            })
            .copied()
    } else {
        target_formula_cells
            .iter()
            .max_by(|a, b| {
                duration_by_cell
                    .get(a)
                    .copied()
                    .unwrap_or(0)
                    .cmp(&duration_by_cell.get(b).copied().unwrap_or(0))
                    .then_with(|| b.cmp(a))
            })
            .copied()
    };

    let mut critical_path_refs = Vec::<CellRef>::new();
    if let Some(mut cursor) = critical_end {
        critical_path_refs.push(cursor);
        while let Some(prev) = predecessor.get(&cursor).copied() {
            critical_path_refs.push(prev);
            cursor = prev;
        }
        critical_path_refs.reverse();
    }
    let critical_path_duration_us = critical_path_refs
        .iter()
        .map(|cell| duration_by_cell.get(cell).copied().unwrap_or(0))
        .sum::<u64>();
    let critical_path = critical_path_refs.iter().map(|cell| to_a1(*cell)).collect();

    let max_node_duration_us = node_timings
        .iter()
        .map(|n| n.duration_us)
        .max()
        .unwrap_or(0);
    let derived_slow_nodes_threshold_us = if max_node_duration_us == 0 {
        0
    } else {
        ((max_node_duration_us * 8) / 10).max(1)
    };
    let slow_nodes_threshold_us =
        slow_nodes_threshold_override_us.unwrap_or(derived_slow_nodes_threshold_us);
    let mut slow_nodes = node_timings
        .iter()
        .filter(|node| node.duration_us >= slow_nodes_threshold_us)
        .cloned()
        .collect::<Vec<_>>();
    slow_nodes.sort_by(|a, b| {
        b.duration_us
            .cmp(&a.duration_us)
            .then_with(|| a.cell.cmp(&b.cell))
    });

    DagInsights {
        node_degrees,
        max_fan_in,
        max_fan_out,
        critical_path,
        critical_path_duration_us,
        critical_path_truncated_by_cycles,
        slow_nodes_threshold_us,
        slow_nodes,
    }
}

fn eval_cell(
    workbook: &Workbook,
    sheet_name: &str,
    cell_ref: CellRef,
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut BTreeMap<CellRef, Result<f64, EvalError>>,
) -> Result<f64, EvalError> {
    if let Some(cached) = cache.get(&cell_ref) {
        return cached.clone();
    }

    if !stack.insert(cell_ref) {
        return Err(EvalError::Cycle(cell_ref));
    }

    let result = match workbook
        .sheets
        .get(sheet_name)
        .and_then(|sheet| sheet.cells.get(&cell_ref))
    {
        Some(cell) => {
            if let Some(formula) = &cell.formula {
                if let Some(parsed) = parsed_formulas.get(&cell_ref) {
                    match parsed {
                        Ok(expr) => {
                            eval_expr(workbook, sheet_name, expr, parsed_formulas, stack, cache)
                        }
                        Err(err) => Err(err.clone()),
                    }
                } else {
                    let parsed = parse_formula_expression(formula)?;
                    eval_expr(workbook, sheet_name, &parsed, parsed_formulas, stack, cache)
                }
            } else {
                value_as_number(&cell.value)
            }
        }
        None => Ok(0.0),
    };

    stack.remove(&cell_ref);
    cache.insert(cell_ref, result.clone());
    result
}

fn eval_expr(
    workbook: &Workbook,
    sheet_name: &str,
    expr: &Expr,
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut BTreeMap<CellRef, Result<f64, EvalError>>,
) -> Result<f64, EvalError> {
    match expr {
        Expr::Number(n) => Ok(*n),
        Expr::Cell(cell_ref) => eval_cell(
            workbook,
            sheet_name,
            *cell_ref,
            parsed_formulas,
            stack,
            cache,
        ),
        Expr::Function { name, args } => eval_function(
            workbook,
            sheet_name,
            name,
            args,
            parsed_formulas,
            stack,
            cache,
        ),
        Expr::UnaryMinus(inner) => Ok(-eval_expr(
            workbook,
            sheet_name,
            inner,
            parsed_formulas,
            stack,
            cache,
        )?),
        Expr::Binary { op, left, right } => {
            let lhs = eval_expr(workbook, sheet_name, left, parsed_formulas, stack, cache)?;
            let rhs = eval_expr(workbook, sheet_name, right, parsed_formulas, stack, cache)?;
            match op {
                BinaryOp::Add => Ok(lhs + rhs),
                BinaryOp::Sub => Ok(lhs - rhs),
                BinaryOp::Mul => Ok(lhs * rhs),
                BinaryOp::Div => {
                    if rhs == 0.0 {
                        Err(EvalError::DivisionByZero)
                    } else {
                        Ok(lhs / rhs)
                    }
                }
            }
        }
    }
}

fn eval_function(
    workbook: &Workbook,
    sheet_name: &str,
    name: &str,
    args: &[Expr],
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut BTreeMap<CellRef, Result<f64, EvalError>>,
) -> Result<f64, EvalError> {
    match name {
        "IF" => {
            if args.len() < 2 || args.len() > 3 {
                return Err(EvalError::Parse);
            }
            let condition = eval_expr(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            if condition != 0.0 {
                eval_expr(
                    workbook,
                    sheet_name,
                    &args[1],
                    parsed_formulas,
                    stack,
                    cache,
                )
            } else if args.len() == 3 {
                eval_expr(
                    workbook,
                    sheet_name,
                    &args[2],
                    parsed_formulas,
                    stack,
                    cache,
                )
            } else {
                Ok(0.0)
            }
        }
        "IFERROR" => {
            if args.len() != 2 {
                return Err(EvalError::Parse);
            }
            match eval_expr(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            ) {
                Ok(value) => Ok(value),
                Err(_) => eval_expr(
                    workbook,
                    sheet_name,
                    &args[1],
                    parsed_formulas,
                    stack,
                    cache,
                ),
            }
        }
        "CHOOSE" => {
            if args.len() < 2 {
                return Err(EvalError::Parse);
            }
            let index_value = eval_expr(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let option_index = trunc_f64_to_i64(index_value)?;
            if option_index < 1 || option_index > (args.len() - 1) as i64 {
                return Err(EvalError::Parse);
            }
            eval_expr(
                workbook,
                sheet_name,
                &args[option_index as usize],
                parsed_formulas,
                stack,
                cache,
            )
        }
        "INDEX" => {
            if args.len() < 2 {
                return Err(EvalError::Parse);
            }
            let index_value = eval_expr(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let selected_index = trunc_f64_to_i64(index_value)?;
            if selected_index < 1 || selected_index > (args.len() - 1) as i64 {
                return Err(EvalError::Parse);
            }
            eval_expr(
                workbook,
                sheet_name,
                &args[selected_index as usize],
                parsed_formulas,
                stack,
                cache,
            )
        }
        "LEN" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let text = eval_expr_as_text(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            Ok(text.chars().count() as f64)
        }
        "CODE" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let text = eval_expr_as_text(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let ch = text.chars().next().ok_or(EvalError::Parse)?;
            Ok(ch as u32 as f64)
        }
        "COUNT" => {
            let mut count = 0usize;
            for arg in args {
                let scalar = eval_expr_as_scalar_value(
                    workbook,
                    sheet_name,
                    arg,
                    parsed_formulas,
                    stack,
                    cache,
                );
                match scalar {
                    CellValue::Number(_) => count = count.saturating_add(1),
                    CellValue::Error(_) => return Err(EvalError::Parse),
                    CellValue::Bool(_) | CellValue::Text(_) | CellValue::Empty => {}
                }
            }
            Ok(count as f64)
        }
        "COUNTA" => {
            let mut count = 0usize;
            for arg in args {
                let scalar = eval_expr_as_scalar_value(
                    workbook,
                    sheet_name,
                    arg,
                    parsed_formulas,
                    stack,
                    cache,
                );
                match scalar {
                    CellValue::Empty => {}
                    CellValue::Error(_) => return Err(EvalError::Parse),
                    CellValue::Number(_) | CellValue::Bool(_) | CellValue::Text(_) => {
                        count = count.saturating_add(1)
                    }
                }
            }
            Ok(count as f64)
        }
        "COUNTBLANK" => {
            let mut count = 0usize;
            for arg in args {
                let scalar = eval_expr_as_scalar_value(
                    workbook,
                    sheet_name,
                    arg,
                    parsed_formulas,
                    stack,
                    cache,
                );
                match scalar {
                    CellValue::Empty => count = count.saturating_add(1),
                    CellValue::Error(_) => return Err(EvalError::Parse),
                    CellValue::Number(_) | CellValue::Bool(_) | CellValue::Text(_) => {}
                }
            }
            Ok(count as f64)
        }
        "N" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let scalar = eval_expr_as_scalar_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            );
            match scalar {
                CellValue::Number(n) => Ok(n),
                CellValue::Bool(true) => Ok(1.0),
                CellValue::Bool(false) => Ok(0.0),
                CellValue::Text(_) | CellValue::Empty => Ok(0.0),
                CellValue::Error(_) => Err(EvalError::Parse),
            }
        }
        "VALUE" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let scalar = eval_expr_as_scalar_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            );
            match scalar {
                CellValue::Number(n) => Ok(n),
                CellValue::Text(text) => parse_text_number(&text),
                CellValue::Bool(_) | CellValue::Empty | CellValue::Error(_) => {
                    Err(EvalError::Parse)
                }
            }
        }
        "DATEVALUE" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let scalar = eval_expr_as_scalar_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            );
            match scalar {
                CellValue::Number(n) => Ok(validated_excel_serial_day(n)? as f64),
                CellValue::Text(text) => parse_text_date_serial(&text),
                CellValue::Bool(_) | CellValue::Empty | CellValue::Error(_) => {
                    Err(EvalError::Parse)
                }
            }
        }
        "TIMEVALUE" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let scalar = eval_expr_as_scalar_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            );
            match scalar {
                CellValue::Number(n) => normalized_time_fraction(n),
                CellValue::Text(text) => parse_text_time_fraction(&text),
                CellValue::Bool(_) | CellValue::Empty | CellValue::Error(_) => {
                    Err(EvalError::Parse)
                }
            }
        }
        "TIME" => {
            if args.len() != 3 {
                return Err(EvalError::Parse);
            }
            let hour = trunc_f64_to_i64(eval_expr(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?)?;
            let minute = trunc_f64_to_i64(eval_expr(
                workbook,
                sheet_name,
                &args[1],
                parsed_formulas,
                stack,
                cache,
            )?)?;
            let second = trunc_f64_to_i64(eval_expr(
                workbook,
                sheet_name,
                &args[2],
                parsed_formulas,
                stack,
                cache,
            )?)?;
            if hour < 0 || minute < 0 || second < 0 {
                return Err(EvalError::Parse);
            }
            let hour_seconds = hour.checked_mul(3600).ok_or(EvalError::Parse)?;
            let minute_seconds = minute.checked_mul(60).ok_or(EvalError::Parse)?;
            let total_seconds = hour_seconds
                .checked_add(minute_seconds)
                .and_then(|value| value.checked_add(second))
                .ok_or(EvalError::Parse)?;
            Ok((total_seconds.rem_euclid(86_400) as f64) / 86_400.0)
        }
        "HOUR" | "MINUTE" | "SECOND" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let scalar = eval_expr_as_scalar_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            );
            let serial = match scalar {
                CellValue::Number(n) => n,
                CellValue::Text(text) => parse_text_time_fraction(&text)?,
                CellValue::Bool(_) | CellValue::Empty | CellValue::Error(_) => {
                    return Err(EvalError::Parse);
                }
            };
            let (hour, minute, second) = extract_hms_from_serial(serial)?;
            match name {
                "HOUR" => Ok(hour as f64),
                "MINUTE" => Ok(minute as f64),
                "SECOND" => Ok(second as f64),
                _ => Err(EvalError::Parse),
            }
        }
        "ISNUMBER" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let scalar = eval_expr_as_scalar_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            );
            Ok(if matches!(scalar, CellValue::Number(_)) {
                1.0
            } else {
                0.0
            })
        }
        "ISTEXT" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let scalar = eval_expr_as_scalar_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            );
            Ok(if matches!(scalar, CellValue::Text(_)) {
                1.0
            } else {
                0.0
            })
        }
        "ISBLANK" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let scalar = eval_expr_as_scalar_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            );
            Ok(if matches!(scalar, CellValue::Empty) {
                1.0
            } else {
                0.0
            })
        }
        "ISLOGICAL" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let scalar = eval_expr_as_scalar_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            );
            Ok(if matches!(scalar, CellValue::Bool(_)) {
                1.0
            } else {
                0.0
            })
        }
        "ISERROR" => {
            if args.len() != 1 {
                return Err(EvalError::Parse);
            }
            let scalar = eval_expr_as_scalar_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            );
            Ok(if matches!(scalar, CellValue::Error(_)) {
                1.0
            } else {
                0.0
            })
        }
        "EXACT" => {
            if args.len() != 2 {
                return Err(EvalError::Parse);
            }
            let left = eval_expr_as_text(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let right = eval_expr_as_text(
                workbook,
                sheet_name,
                &args[1],
                parsed_formulas,
                stack,
                cache,
            )?;
            Ok(if left == right { 1.0 } else { 0.0 })
        }
        "FIND" | "SEARCH" => {
            if args.len() < 2 || args.len() > 3 {
                return Err(EvalError::Parse);
            }
            let needle = eval_expr_as_text(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let haystack = eval_expr_as_text(
                workbook,
                sheet_name,
                &args[1],
                parsed_formulas,
                stack,
                cache,
            )?;
            let start_pos = if args.len() == 3 {
                let raw_start = eval_expr(
                    workbook,
                    sheet_name,
                    &args[2],
                    parsed_formulas,
                    stack,
                    cache,
                )?;
                parse_text_start_position(raw_start)?
            } else {
                1
            };
            let is_case_insensitive = name == "SEARCH";
            let found = find_text_position(&needle, &haystack, start_pos, is_case_insensitive)
                .ok_or(EvalError::Parse)?;
            Ok(found as f64)
        }
        _ => {
            let mut values = Vec::<f64>::with_capacity(args.len());
            for arg in args {
                values.push(eval_expr(
                    workbook,
                    sheet_name,
                    arg,
                    parsed_formulas,
                    stack,
                    cache,
                )?);
            }

            match name {
                "SUM" => Ok(values.into_iter().sum()),
                "PRODUCT" => {
                    if values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    Ok(values.into_iter().product())
                }
                "AVERAGE" | "AVG" => {
                    if values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    let sum = values.iter().sum::<f64>();
                    Ok(sum / values.len() as f64)
                }
                "MIN" => {
                    let min = values
                        .into_iter()
                        .min_by(f64::total_cmp)
                        .ok_or(EvalError::Parse)?;
                    Ok(min)
                }
                "MAX" => {
                    let max = values
                        .into_iter()
                        .max_by(f64::total_cmp)
                        .ok_or(EvalError::Parse)?;
                    Ok(max)
                }
                "MEDIAN" => {
                    if values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    let mut sorted = values;
                    sorted.sort_by(f64::total_cmp);
                    let len = sorted.len();
                    if len % 2 == 1 {
                        Ok(sorted[len / 2])
                    } else {
                        Ok((sorted[len / 2 - 1] + sorted[len / 2]) / 2.0)
                    }
                }
                "SMALL" => {
                    if values.len() < 2 {
                        return Err(EvalError::Parse);
                    }
                    let rank = parse_rank_index(values[values.len() - 1], values.len() - 1)?;
                    let mut candidates = values[..values.len() - 1].to_vec();
                    candidates.sort_by(f64::total_cmp);
                    Ok(candidates[rank])
                }
                "LARGE" => {
                    if values.len() < 2 {
                        return Err(EvalError::Parse);
                    }
                    let rank = parse_rank_index(values[values.len() - 1], values.len() - 1)?;
                    let mut candidates = values[..values.len() - 1].to_vec();
                    candidates.sort_by(f64::total_cmp);
                    Ok(candidates[candidates.len() - 1 - rank])
                }
                "ABS" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    Ok(values[0].abs())
                }
                "INT" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    Ok(values[0].floor())
                }
                "QUOTIENT" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    if values[1] == 0.0 {
                        return Err(EvalError::Parse);
                    }
                    let result = (values[0] / values[1]).trunc();
                    if !result.is_finite() {
                        return Err(EvalError::Parse);
                    }
                    Ok(result)
                }
                "MOD" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let divisor = values[1];
                    if divisor == 0.0 {
                        return Err(EvalError::Parse);
                    }
                    let quotient = (values[0] / divisor).floor();
                    Ok(values[0] - divisor * quotient)
                }
                "ROUND" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let digits = trunc_f64_to_i64(values[1])?;
                    round_with_digits(values[0], digits, RoundMode::Nearest)
                }
                "ROUNDUP" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let digits = trunc_f64_to_i64(values[1])?;
                    round_with_digits(values[0], digits, RoundMode::AwayFromZero)
                }
                "ROUNDDOWN" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let digits = trunc_f64_to_i64(values[1])?;
                    round_with_digits(values[0], digits, RoundMode::TowardZero)
                }
                "TRUNC" => {
                    if values.is_empty() || values.len() > 2 {
                        return Err(EvalError::Parse);
                    }
                    let digits = if values.len() == 2 {
                        trunc_f64_to_i64(values[1])?
                    } else {
                        0
                    };
                    trunc_with_digits(values[0], digits)
                }
                "MROUND" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    mround(values[0], values[1])
                }
                "POWER" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let result = values[0].powf(values[1]);
                    if !result.is_finite() {
                        return Err(EvalError::Parse);
                    }
                    Ok(result)
                }
                "SQRT" => {
                    if values.len() != 1 || values[0] < 0.0 {
                        return Err(EvalError::Parse);
                    }
                    let result = values[0].sqrt();
                    if !result.is_finite() {
                        return Err(EvalError::Parse);
                    }
                    Ok(result)
                }
                "SIGN" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    Ok(if values[0] > 0.0 {
                        1.0
                    } else if values[0] < 0.0 {
                        -1.0
                    } else {
                        0.0
                    })
                }
                "EVEN" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    round_to_even_away_from_zero(values[0])
                }
                "ODD" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    round_to_odd_away_from_zero(values[0])
                }
                "ISEVEN" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    Ok(if is_even_with_trunc(values[0])? {
                        1.0
                    } else {
                        0.0
                    })
                }
                "ISODD" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    Ok(if is_even_with_trunc(values[0])? {
                        0.0
                    } else {
                        1.0
                    })
                }
                "CEILING" => {
                    if values.is_empty() || values.len() > 2 {
                        return Err(EvalError::Parse);
                    }
                    let significance = if values.len() == 2 { values[1] } else { 1.0 };
                    ceiling_with_significance(values[0], significance)
                }
                "FLOOR" => {
                    if values.is_empty() || values.len() > 2 {
                        return Err(EvalError::Parse);
                    }
                    let significance = if values.len() == 2 { values[1] } else { 1.0 };
                    floor_with_significance(values[0], significance)
                }
                "AND" => {
                    if values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    Ok(if values.iter().all(|v| *v != 0.0) {
                        1.0
                    } else {
                        0.0
                    })
                }
                "OR" => {
                    if values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    Ok(if values.iter().any(|v| *v != 0.0) {
                        1.0
                    } else {
                        0.0
                    })
                }
                "NOT" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    Ok(if values[0] == 0.0 { 1.0 } else { 0.0 })
                }
                "MATCH" => {
                    if values.len() < 2 {
                        return Err(EvalError::Parse);
                    }
                    let needle = values[0];
                    let mut candidates = &values[1..];
                    let mut mode = 0i64;
                    if values.len() >= 4 {
                        let mode_candidate = values[values.len() - 1];
                        if let Ok(parsed_mode) = trunc_f64_to_i64(mode_candidate) {
                            if parsed_mode == -1 || parsed_mode == 0 || parsed_mode == 1 {
                                mode = parsed_mode;
                                candidates = &values[1..values.len() - 1];
                            }
                        }
                    }
                    let index = find_match_index_match(needle, candidates, mode)?;
                    Ok(index as f64)
                }
                "XMATCH" => {
                    if values.len() < 2 {
                        return Err(EvalError::Parse);
                    }
                    let needle = values[0];
                    let mut candidates = &values[1..];
                    let mut mode = 0i64;
                    if values.len() >= 4 {
                        let mode_candidate = values[values.len() - 1];
                        if let Ok(parsed_mode) = trunc_f64_to_i64(mode_candidate) {
                            if parsed_mode == -1 || parsed_mode == 0 || parsed_mode == 1 {
                                mode = parsed_mode;
                                candidates = &values[1..values.len() - 1];
                            }
                        }
                    }
                    let index = find_match_index_xmatch(needle, candidates, mode)?;
                    Ok(index as f64)
                }
                "DATE" => {
                    if values.len() != 3 {
                        return Err(EvalError::Parse);
                    }
                    excel_date_to_serial(values[0], values[1], values[2])
                }
                "YEAR" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    let (year, _, _) = excel_serial_to_ymd(values[0])?;
                    Ok(year as f64)
                }
                "MONTH" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    let (_, month, _) = excel_serial_to_ymd(values[0])?;
                    Ok(month as f64)
                }
                "DAY" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    let (_, _, day) = excel_serial_to_ymd(values[0])?;
                    Ok(day as f64)
                }
                "DAYS" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let end_serial = validated_excel_serial_day(values[0])?;
                    let start_serial = validated_excel_serial_day(values[1])?;
                    Ok((end_serial - start_serial) as f64)
                }
                "EDATE" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let start_serial = validated_excel_serial_day(values[0])?;
                    let start_date = excel_serial_to_naive_date(start_serial)?;
                    let month_offset = trunc_f64_to_i64(values[1])?;
                    let (target_year, target_month) =
                        shift_year_month(start_date.year(), start_date.month(), month_offset)?;
                    let target_day = start_date
                        .day()
                        .min(last_day_of_month(target_year, target_month)?);
                    let target_date =
                        NaiveDate::from_ymd_opt(target_year, target_month, target_day)
                            .ok_or(EvalError::Parse)?;
                    Ok(excel_naive_date_to_serial(target_date)? as f64)
                }
                "EOMONTH" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let start_serial = validated_excel_serial_day(values[0])?;
                    let start_date = excel_serial_to_naive_date(start_serial)?;
                    let month_offset = trunc_f64_to_i64(values[1])?;
                    let (target_year, target_month) =
                        shift_year_month(start_date.year(), start_date.month(), month_offset)?;
                    let target_day = last_day_of_month(target_year, target_month)?;
                    let target_date =
                        NaiveDate::from_ymd_opt(target_year, target_month, target_day)
                            .ok_or(EvalError::Parse)?;
                    Ok(excel_naive_date_to_serial(target_date)? as f64)
                }
                "WEEKDAY" => {
                    if values.is_empty() || values.len() > 2 {
                        return Err(EvalError::Parse);
                    }
                    let serial = validated_excel_serial_day(values[0])?;
                    let weekday_serial = if serial >= 60 { serial - 1 } else { serial };
                    let return_type = if values.len() == 2 {
                        trunc_f64_to_i64(values[1])?
                    } else {
                        1
                    };
                    match return_type {
                        1 => Ok((weekday_serial.rem_euclid(7) + 1) as f64),
                        2 => Ok(((weekday_serial + 6).rem_euclid(7) + 1) as f64),
                        3 => Ok(((weekday_serial + 6).rem_euclid(7)) as f64),
                        _ => Err(EvalError::Parse),
                    }
                }
                "WEEKNUM" => {
                    if values.is_empty() || values.len() > 2 {
                        return Err(EvalError::Parse);
                    }
                    let serial = validated_excel_serial_day(values[0])?;
                    let date = excel_serial_to_naive_date(serial)?;
                    let return_type = if values.len() == 2 {
                        trunc_f64_to_i64(values[1])?
                    } else {
                        1
                    };
                    let jan1 =
                        NaiveDate::from_ymd_opt(date.year(), 1, 1).ok_or(EvalError::Parse)?;
                    let ordinal = date.ordinal() as i64;
                    let week = match return_type {
                        1 => {
                            let jan1_weekday = jan1.weekday().num_days_from_sunday() as i64 + 1;
                            ((ordinal + jan1_weekday - 2) / 7) + 1
                        }
                        2 => {
                            let jan1_weekday = jan1.weekday().num_days_from_monday() as i64 + 1;
                            ((ordinal + jan1_weekday - 2) / 7) + 1
                        }
                        _ => return Err(EvalError::Parse),
                    };
                    Ok(week as f64)
                }
                "ISOWEEKNUM" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    let serial = validated_excel_serial_day(values[0])?;
                    let date = excel_serial_to_naive_date(serial)?;
                    Ok(date.iso_week().week() as f64)
                }
                _ => Err(EvalError::Parse),
            }
        }
    }
}

fn eval_status(result: &Result<f64, EvalError>) -> &'static str {
    match result {
        Ok(_) => "ok",
        Err(EvalError::Cycle(_)) => "cycle",
        Err(EvalError::Parse) => "parse_error",
        Err(EvalError::DivisionByZero) => "division_by_zero",
    }
}

fn parse_formula_expression(formula: &str) -> Result<Expr, EvalError> {
    let body = formula.strip_prefix('=').unwrap_or(formula).trim();
    if body.is_empty() {
        return Err(EvalError::Parse);
    }
    let mut parser = FormulaParser::new(body);
    let expr = parser.parse_expression()?;
    parser.skip_ws();
    if !parser.is_eof() {
        return Err(EvalError::Parse);
    }
    Ok(expr)
}

fn collect_expr_references(expr: &Expr, out: &mut BTreeSet<CellRef>) {
    match expr {
        Expr::Number(_) => {}
        Expr::Cell(cell) => {
            out.insert(*cell);
        }
        Expr::Function { args, .. } => {
            for arg in args {
                collect_expr_references(arg, out);
            }
        }
        Expr::UnaryMinus(inner) => collect_expr_references(inner, out),
        Expr::Binary { left, right, .. } => {
            collect_expr_references(left, out);
            collect_expr_references(right, out);
        }
    }
}

fn count_expr_functions(expr: &Expr) -> usize {
    match expr {
        Expr::Number(_) | Expr::Cell(_) => 0,
        Expr::Function { args, .. } => 1 + args.iter().map(count_expr_functions).sum::<usize>(),
        Expr::UnaryMinus(inner) => count_expr_functions(inner),
        Expr::Binary { left, right, .. } => {
            count_expr_functions(left) + count_expr_functions(right)
        }
    }
}

fn count_expr_nodes(expr: &Expr) -> usize {
    match expr {
        Expr::Number(_) | Expr::Cell(_) => 1,
        Expr::Function { args, .. } => 1 + args.iter().map(count_expr_nodes).sum::<usize>(),
        Expr::UnaryMinus(inner) => 1 + count_expr_nodes(inner),
        Expr::Binary { left, right, .. } => 1 + count_expr_nodes(left) + count_expr_nodes(right),
    }
}

struct FormulaParser<'a> {
    input: &'a str,
    index: usize,
}

impl<'a> FormulaParser<'a> {
    fn new(input: &'a str) -> Self {
        Self { input, index: 0 }
    }

    fn is_eof(&self) -> bool {
        self.index >= self.input.len()
    }

    fn skip_ws(&mut self) {
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_whitespace() {
                self.bump_char();
            } else {
                break;
            }
        }
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.index..].chars().next()
    }

    fn bump_char(&mut self) -> Option<char> {
        let ch = self.peek_char()?;
        self.index += ch.len_utf8();
        Some(ch)
    }

    fn parse_expression(&mut self) -> Result<Expr, EvalError> {
        let mut expr = self.parse_term()?;
        loop {
            self.skip_ws();
            let op = match self.peek_char() {
                Some('+') => BinaryOp::Add,
                Some('-') => BinaryOp::Sub,
                _ => break,
            };
            self.bump_char();
            let right = self.parse_term()?;
            expr = Expr::Binary {
                op,
                left: Box::new(expr),
                right: Box::new(right),
            };
        }
        Ok(expr)
    }

    fn parse_term(&mut self) -> Result<Expr, EvalError> {
        let mut expr = self.parse_factor()?;
        loop {
            self.skip_ws();
            let op = match self.peek_char() {
                Some('*') => BinaryOp::Mul,
                Some('/') => BinaryOp::Div,
                _ => break,
            };
            self.bump_char();
            let right = self.parse_factor()?;
            expr = Expr::Binary {
                op,
                left: Box::new(expr),
                right: Box::new(right),
            };
        }
        Ok(expr)
    }

    fn parse_factor(&mut self) -> Result<Expr, EvalError> {
        self.skip_ws();
        match self.peek_char() {
            Some('+') => {
                self.bump_char();
                self.parse_factor()
            }
            Some('-') => {
                self.bump_char();
                Ok(Expr::UnaryMinus(Box::new(self.parse_factor()?)))
            }
            Some('(') => {
                self.bump_char();
                let expr = self.parse_expression()?;
                self.skip_ws();
                if self.bump_char() != Some(')') {
                    return Err(EvalError::Parse);
                }
                Ok(expr)
            }
            Some(ch) if ch.is_ascii_digit() || ch == '.' => self.parse_number(),
            Some(ch) if ch.is_ascii_alphabetic() || ch == '_' => self.parse_identifier_expr(),
            _ => Err(EvalError::Parse),
        }
    }

    fn parse_number(&mut self) -> Result<Expr, EvalError> {
        let start = self.index;
        let mut saw_digit = false;
        let mut saw_dot = false;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_digit() {
                saw_digit = true;
                self.bump_char();
            } else if ch == '.' && !saw_dot {
                saw_dot = true;
                self.bump_char();
            } else {
                break;
            }
        }
        if !saw_digit {
            return Err(EvalError::Parse);
        }
        let token = &self.input[start..self.index];
        let number = token.parse::<f64>().map_err(|_| EvalError::Parse)?;
        Ok(Expr::Number(number))
    }

    fn parse_identifier_expr(&mut self) -> Result<Expr, EvalError> {
        let token = self.parse_identifier_token()?;
        self.skip_ws();
        if self.peek_char() == Some('(') {
            self.bump_char();
            self.skip_ws();
            let mut args = Vec::<Expr>::new();
            if self.peek_char() != Some(')') {
                loop {
                    let arg = self.parse_expression()?;
                    args.push(arg);
                    self.skip_ws();
                    match self.peek_char() {
                        Some(',') => {
                            self.bump_char();
                            self.skip_ws();
                        }
                        Some(')') => break,
                        _ => return Err(EvalError::Parse),
                    }
                }
            }
            if self.bump_char() != Some(')') {
                return Err(EvalError::Parse);
            }
            return Ok(Expr::Function {
                name: token.to_ascii_uppercase(),
                args,
            });
        }

        parse_a1_cell(&token)
            .map(Expr::Cell)
            .ok_or(EvalError::Parse)
    }

    fn parse_identifier_token(&mut self) -> Result<String, EvalError> {
        let start = self.index;
        match self.peek_char() {
            Some(ch) if ch.is_ascii_alphabetic() || ch == '_' => {
                self.bump_char();
            }
            _ => return Err(EvalError::Parse),
        }
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_alphanumeric() || ch == '_' {
                self.bump_char();
            } else {
                break;
            }
        }
        Ok(self.input[start..self.index].to_string())
    }
}

fn parse_a1_cell(input: &str) -> Option<CellRef> {
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
        let v = (ch as u32).checked_sub('A' as u32)? + 1;
        col = col.checked_mul(26)?.checked_add(v)?;
    }

    let row = row_part.parse::<u32>().ok()?;
    if row == 0 || col == 0 {
        return None;
    }

    Some(CellRef { row, col })
}

fn value_as_number(value: &CellValue) -> Result<f64, EvalError> {
    match value {
        CellValue::Number(n) => Ok(*n),
        CellValue::Bool(true) => Ok(1.0),
        CellValue::Bool(false) => Ok(0.0),
        CellValue::Text(_) => Ok(0.0),
        CellValue::Empty => Ok(0.0),
        CellValue::Error(_) => Err(EvalError::Parse),
    }
}

fn eval_expr_as_text(
    workbook: &Workbook,
    sheet_name: &str,
    expr: &Expr,
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut BTreeMap<CellRef, Result<f64, EvalError>>,
) -> Result<String, EvalError> {
    match expr {
        Expr::Cell(cell_ref) => {
            let maybe_cell = workbook
                .sheets
                .get(sheet_name)
                .and_then(|sheet| sheet.cells.get(cell_ref));
            match maybe_cell {
                Some(cell) if cell.formula.is_none() => cell_value_as_text(&cell.value),
                _ => {
                    let value = eval_cell(
                        workbook,
                        sheet_name,
                        *cell_ref,
                        parsed_formulas,
                        stack,
                        cache,
                    )?;
                    Ok(format_number_for_text(value))
                }
            }
        }
        _ => {
            let value = eval_expr(workbook, sheet_name, expr, parsed_formulas, stack, cache)?;
            Ok(format_number_for_text(value))
        }
    }
}

fn cell_value_as_text(value: &CellValue) -> Result<String, EvalError> {
    match value {
        CellValue::Number(n) => Ok(format_number_for_text(*n)),
        CellValue::Bool(true) => Ok("TRUE".to_string()),
        CellValue::Bool(false) => Ok("FALSE".to_string()),
        CellValue::Text(text) => Ok(text.clone()),
        CellValue::Empty => Ok(String::new()),
        CellValue::Error(_) => Err(EvalError::Parse),
    }
}

fn format_number_for_text(value: f64) -> String {
    if value == 0.0 {
        return "0".to_string();
    }
    let mut rendered = format!("{value}");
    if rendered.contains('.') && !rendered.contains('e') && !rendered.contains('E') {
        while rendered.ends_with('0') {
            rendered.pop();
        }
        if rendered.ends_with('.') {
            rendered.pop();
        }
    }
    if rendered == "-0" {
        "0".to_string()
    } else {
        rendered
    }
}

#[derive(Clone, Copy)]
enum RoundMode {
    Nearest,
    AwayFromZero,
    TowardZero,
}

fn round_with_digits(value: f64, digits: i64, mode: RoundMode) -> Result<f64, EvalError> {
    if !value.is_finite() {
        return Err(EvalError::Parse);
    }
    if digits < i32::MIN as i64 || digits > i32::MAX as i64 {
        return Err(EvalError::Parse);
    }

    let digits_i32 = digits as i32;
    let factor = 10f64.powi(digits_i32.abs());
    if !factor.is_finite() || factor == 0.0 {
        return Err(EvalError::Parse);
    }

    let scaled = if digits_i32 >= 0 {
        value * factor
    } else {
        value / factor
    };

    if !scaled.is_finite() {
        return Err(EvalError::Parse);
    }

    let rounded = match mode {
        RoundMode::Nearest => scaled.round(),
        RoundMode::AwayFromZero => scaled.signum() * scaled.abs().ceil(),
        RoundMode::TowardZero => scaled.signum() * scaled.abs().floor(),
    };

    let unscaled = if digits_i32 >= 0 {
        rounded / factor
    } else {
        rounded * factor
    };

    if !unscaled.is_finite() {
        return Err(EvalError::Parse);
    }
    Ok(unscaled)
}

fn trunc_with_digits(value: f64, digits: i64) -> Result<f64, EvalError> {
    if !value.is_finite() {
        return Err(EvalError::Parse);
    }
    if digits < i32::MIN as i64 || digits > i32::MAX as i64 {
        return Err(EvalError::Parse);
    }

    let digits_i32 = digits as i32;
    let factor = 10f64.powi(digits_i32.abs());
    if !factor.is_finite() || factor == 0.0 {
        return Err(EvalError::Parse);
    }

    let scaled = if digits_i32 >= 0 {
        value * factor
    } else {
        value / factor
    };
    if !scaled.is_finite() {
        return Err(EvalError::Parse);
    }

    let truncated = scaled.trunc();
    let unscaled = if digits_i32 >= 0 {
        truncated / factor
    } else {
        truncated * factor
    };
    if !unscaled.is_finite() {
        return Err(EvalError::Parse);
    }
    Ok(unscaled)
}

fn mround(value: f64, multiple: f64) -> Result<f64, EvalError> {
    if !value.is_finite() || !multiple.is_finite() {
        return Err(EvalError::Parse);
    }
    if multiple == 0.0 {
        return Ok(0.0);
    }
    if value != 0.0 && value.signum() != multiple.signum() {
        return Err(EvalError::Parse);
    }
    let result = (value / multiple).round() * multiple;
    if !result.is_finite() {
        return Err(EvalError::Parse);
    }
    if result == -0.0 {
        Ok(0.0)
    } else {
        Ok(result)
    }
}

fn round_to_even_away_from_zero(value: f64) -> Result<f64, EvalError> {
    if !value.is_finite() {
        return Err(EvalError::Parse);
    }
    if value == 0.0 {
        return Ok(0.0);
    }
    let mut magnitude = value.abs().ceil();
    let parity = magnitude.rem_euclid(2.0);
    if parity >= 1.0 {
        magnitude += 1.0;
    }
    let result = if value < 0.0 { -magnitude } else { magnitude };
    if !result.is_finite() {
        return Err(EvalError::Parse);
    }
    Ok(result)
}

fn round_to_odd_away_from_zero(value: f64) -> Result<f64, EvalError> {
    if !value.is_finite() {
        return Err(EvalError::Parse);
    }
    if value == 0.0 {
        return Ok(1.0);
    }
    let mut magnitude = value.abs().ceil();
    let parity = magnitude.rem_euclid(2.0);
    if parity < 1.0 {
        magnitude += 1.0;
    }
    let result = if value < 0.0 { -magnitude } else { magnitude };
    if !result.is_finite() {
        return Err(EvalError::Parse);
    }
    Ok(result)
}

fn is_even_with_trunc(value: f64) -> Result<bool, EvalError> {
    let integer = trunc_f64_to_i64(value)?;
    Ok(integer.rem_euclid(2) == 0)
}

fn parse_rank_index(rank_value: f64, candidate_len: usize) -> Result<usize, EvalError> {
    if candidate_len == 0 {
        return Err(EvalError::Parse);
    }
    let rank = trunc_f64_to_i64(rank_value)?;
    if rank < 1 || rank > candidate_len as i64 {
        return Err(EvalError::Parse);
    }
    Ok((rank - 1) as usize)
}

fn ceiling_with_significance(value: f64, significance: f64) -> Result<f64, EvalError> {
    if !value.is_finite() || !significance.is_finite() || significance == 0.0 {
        return Err(EvalError::Parse);
    }
    let result = (value / significance).ceil() * significance;
    if !result.is_finite() {
        return Err(EvalError::Parse);
    }
    if result == -0.0 {
        Ok(0.0)
    } else {
        Ok(result)
    }
}

fn floor_with_significance(value: f64, significance: f64) -> Result<f64, EvalError> {
    if !value.is_finite() || !significance.is_finite() || significance == 0.0 {
        return Err(EvalError::Parse);
    }
    let result = (value / significance).floor() * significance;
    if !result.is_finite() {
        return Err(EvalError::Parse);
    }
    if result == -0.0 {
        Ok(0.0)
    } else {
        Ok(result)
    }
}

fn trunc_f64_to_i64(value: f64) -> Result<i64, EvalError> {
    if !value.is_finite() {
        return Err(EvalError::Parse);
    }
    let truncated = value.trunc();
    if truncated < i64::MIN as f64 || truncated > i64::MAX as f64 {
        return Err(EvalError::Parse);
    }
    Ok(truncated as i64)
}

fn parse_text_start_position(value: f64) -> Result<usize, EvalError> {
    let start = trunc_f64_to_i64(value)?;
    if start < 1 || start > usize::MAX as i64 {
        return Err(EvalError::Parse);
    }
    Ok(start as usize)
}

fn parse_text_number(input: &str) -> Result<f64, EvalError> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return Err(EvalError::Parse);
    }
    trimmed.parse::<f64>().map_err(|_| EvalError::Parse)
}

fn parse_text_date_serial(input: &str) -> Result<f64, EvalError> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return Err(EvalError::Parse);
    }
    for format in ["%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y"] {
        if let Ok(date) = NaiveDate::parse_from_str(trimmed, format) {
            return Ok(excel_naive_date_to_serial(date)? as f64);
        }
    }
    Err(EvalError::Parse)
}

fn parse_text_time_fraction(input: &str) -> Result<f64, EvalError> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return Err(EvalError::Parse);
    }
    for format in [
        "%H:%M:%S",
        "%H:%M",
        "%I:%M:%S %p",
        "%I:%M %p",
        "%I:%M:%S%p",
        "%I:%M%p",
    ] {
        if let Ok(time) = NaiveTime::parse_from_str(trimmed, format) {
            return Ok(time.num_seconds_from_midnight() as f64 / 86_400.0);
        }
    }
    Err(EvalError::Parse)
}

fn normalized_time_fraction(value: f64) -> Result<f64, EvalError> {
    if !value.is_finite() || value < 0.0 {
        return Err(EvalError::Parse);
    }
    let mut fraction = value.fract();
    if fraction < 0.0 {
        fraction += 1.0;
    }
    Ok(fraction)
}

fn extract_hms_from_serial(serial: f64) -> Result<(u32, u32, u32), EvalError> {
    let fraction = normalized_time_fraction(serial)?;
    let total_seconds = (fraction * 86_400.0).floor() as u32;
    let hour = total_seconds / 3600;
    let minute = (total_seconds % 3600) / 60;
    let second = total_seconds % 60;
    Ok((hour, minute, second))
}

fn find_match_index_match(needle: f64, candidates: &[f64], mode: i64) -> Result<usize, EvalError> {
    if candidates.is_empty() {
        return Err(EvalError::Parse);
    }
    match mode {
        0 => candidates
            .iter()
            .position(|candidate| needle.total_cmp(candidate).is_eq())
            .map(|idx| idx + 1)
            .ok_or(EvalError::Parse),
        1 => {
            let mut best: Option<(usize, f64)> = None;
            for (idx, candidate) in candidates.iter().enumerate() {
                if *candidate <= needle {
                    match best {
                        None => best = Some((idx + 1, *candidate)),
                        Some((best_idx, best_value)) => {
                            if *candidate > best_value
                                || (candidate.total_cmp(&best_value).is_eq() && idx + 1 > best_idx)
                            {
                                best = Some((idx + 1, *candidate));
                            }
                        }
                    }
                }
            }
            best.map(|(idx, _)| idx).ok_or(EvalError::Parse)
        }
        -1 => {
            let mut best: Option<(usize, f64)> = None;
            for (idx, candidate) in candidates.iter().enumerate() {
                if *candidate >= needle {
                    match best {
                        None => best = Some((idx + 1, *candidate)),
                        Some((best_idx, best_value)) => {
                            if *candidate < best_value
                                || (candidate.total_cmp(&best_value).is_eq() && idx + 1 < best_idx)
                            {
                                best = Some((idx + 1, *candidate));
                            }
                        }
                    }
                }
            }
            best.map(|(idx, _)| idx).ok_or(EvalError::Parse)
        }
        _ => Err(EvalError::Parse),
    }
}

fn find_match_index_xmatch(needle: f64, candidates: &[f64], mode: i64) -> Result<usize, EvalError> {
    match mode {
        0 => find_match_index_match(needle, candidates, 0),
        -1 => find_match_index_match(needle, candidates, 1),
        1 => find_match_index_match(needle, candidates, -1),
        _ => Err(EvalError::Parse),
    }
}

fn find_text_position(
    needle: &str,
    haystack: &str,
    start_pos: usize,
    case_insensitive: bool,
) -> Option<usize> {
    if start_pos == 0 {
        return None;
    }

    let (needle_norm, haystack_norm) = if case_insensitive {
        (needle.to_lowercase(), haystack.to_lowercase())
    } else {
        (needle.to_string(), haystack.to_string())
    };
    let needle_chars = needle_norm.chars().collect::<Vec<_>>();
    let haystack_chars = haystack_norm.chars().collect::<Vec<_>>();

    if start_pos > haystack_chars.len() + 1 {
        return None;
    }
    if needle_chars.is_empty() {
        return Some(start_pos);
    }
    if start_pos > haystack_chars.len() {
        return None;
    }
    if needle_chars.len() > haystack_chars.len() {
        return None;
    }

    let start_index = start_pos - 1;
    let max_start = haystack_chars.len() - needle_chars.len();
    for index in start_index..=max_start {
        if haystack_chars[index..index + needle_chars.len()] == needle_chars[..] {
            return Some(index + 1);
        }
    }
    None
}

fn validated_excel_serial_day(value: f64) -> Result<i64, EvalError> {
    let serial = trunc_f64_to_i64(value)?;
    if serial <= 0 {
        return Err(EvalError::Parse);
    }
    let _ = excel_serial_to_ymd(serial as f64)?;
    Ok(serial)
}

fn eval_expr_as_scalar_value(
    workbook: &Workbook,
    sheet_name: &str,
    expr: &Expr,
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut BTreeMap<CellRef, Result<f64, EvalError>>,
) -> CellValue {
    match expr {
        Expr::Cell(cell_ref) => {
            let maybe_cell = workbook
                .sheets
                .get(sheet_name)
                .and_then(|sheet| sheet.cells.get(cell_ref));
            match maybe_cell {
                None => CellValue::Empty,
                Some(cell) if cell.formula.is_none() => cell.value.clone(),
                Some(_) => match eval_cell(
                    workbook,
                    sheet_name,
                    *cell_ref,
                    parsed_formulas,
                    stack,
                    cache,
                ) {
                    Ok(value) => CellValue::Number(value),
                    Err(_) => CellValue::Error("#ERROR".to_string()),
                },
            }
        }
        _ => match eval_expr(workbook, sheet_name, expr, parsed_formulas, stack, cache) {
            Ok(value) => CellValue::Number(value),
            Err(_) => CellValue::Error("#ERROR".to_string()),
        },
    }
}

fn excel_serial_to_naive_date(serial: i64) -> Result<NaiveDate, EvalError> {
    if serial <= 0 {
        return Err(EvalError::Parse);
    }
    let adjusted = if serial > 60 { serial - 1 } else { serial };
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 31).ok_or(EvalError::Parse)?;
    epoch
        .checked_add_signed(Duration::days(adjusted))
        .ok_or(EvalError::Parse)
}

fn excel_naive_date_to_serial(date: NaiveDate) -> Result<i64, EvalError> {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 31).ok_or(EvalError::Parse)?;
    let mut serial = date.signed_duration_since(epoch).num_days();
    if serial >= 60 {
        serial += 1;
    }
    if serial <= 0 {
        return Err(EvalError::Parse);
    }
    Ok(serial)
}

fn shift_year_month(year: i32, month: u32, month_offset: i64) -> Result<(i32, u32), EvalError> {
    if !(1..=12).contains(&month) {
        return Err(EvalError::Parse);
    }
    let month_index = month as i64 - 1 + month_offset;
    let shifted_year = year as i64 + month_index.div_euclid(12);
    if shifted_year < 0 || shifted_year > i32::MAX as i64 {
        return Err(EvalError::Parse);
    }
    let shifted_month = month_index.rem_euclid(12) as u32 + 1;
    Ok((shifted_year as i32, shifted_month))
}

fn last_day_of_month(year: i32, month: u32) -> Result<u32, EvalError> {
    let first_day = NaiveDate::from_ymd_opt(year, month, 1).ok_or(EvalError::Parse)?;
    let (next_year, next_month) = if month == 12 {
        (year.saturating_add(1), 1)
    } else {
        (year, month + 1)
    };
    let next_first = NaiveDate::from_ymd_opt(next_year, next_month, 1).ok_or(EvalError::Parse)?;
    let last_day = next_first
        .checked_sub_signed(Duration::days(1))
        .ok_or(EvalError::Parse)?;
    if last_day < first_day {
        return Err(EvalError::Parse);
    }
    Ok(last_day.day())
}

fn excel_serial_to_ymd(serial_value: f64) -> Result<(i32, u32, u32), EvalError> {
    let serial = trunc_f64_to_i64(serial_value)?;
    if serial <= 0 {
        return Err(EvalError::Parse);
    }
    if serial == 60 {
        return Ok((1900, 2, 29));
    }

    let adjusted = if serial > 60 { serial - 1 } else { serial };
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 31).ok_or(EvalError::Parse)?;
    let date = epoch
        .checked_add_signed(Duration::days(adjusted))
        .ok_or(EvalError::Parse)?;
    Ok((date.year(), date.month(), date.day()))
}

fn excel_date_to_serial(
    year_value: f64,
    month_value: f64,
    day_value: f64,
) -> Result<f64, EvalError> {
    let mut year = trunc_f64_to_i64(year_value)?;
    if (0..=1899).contains(&year) {
        year += 1900;
    }
    if !(0..=9999).contains(&year) {
        return Err(EvalError::Parse);
    }

    let month = trunc_f64_to_i64(month_value)?;
    let day = trunc_f64_to_i64(day_value)?;

    let month_index = month - 1;
    let adjusted_year = year + month_index.div_euclid(12);
    if !(0..=9999).contains(&adjusted_year) {
        return Err(EvalError::Parse);
    }
    let adjusted_month = (month_index.rem_euclid(12) + 1) as u32;

    let first_of_month =
        NaiveDate::from_ymd_opt(adjusted_year as i32, adjusted_month, 1).ok_or(EvalError::Parse)?;
    let date = first_of_month
        .checked_add_signed(Duration::days(day - 1))
        .ok_or(EvalError::Parse)?;
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 31).ok_or(EvalError::Parse)?;

    let mut serial = date.signed_duration_since(epoch).num_days();
    if serial >= 60 {
        serial += 1;
    }
    if serial <= 0 {
        return Err(EvalError::Parse);
    }
    Ok(serial as f64)
}

fn to_a1(cell_ref: CellRef) -> String {
    let mut col = cell_ref.col;
    let mut letters = Vec::<char>::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        letters.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    letters.reverse();
    format!(
        "{}{}",
        letters.into_iter().collect::<String>(),
        cell_ref.row
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{Mutation, Workbook};
    use crate::telemetry::NoopEventSink;

    #[test]
    fn recalculates_formula_cells() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(2.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(3.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=A1+B1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.evaluated_cells, 1);

        let value = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 3 }))
            .expect("cell")
            .value
            .clone();

        assert_eq!(value, CellValue::Number(5.0));
    }

    #[test]
    fn marks_cycles_as_error() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            formula: "=B1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            formula: "=A1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.cycle_count, 2);
    }

    #[test]
    fn supports_operator_precedence_and_parentheses() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(2.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(3.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Number(4.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            formula: "=A1+B1*C1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            formula: "=(A1+B1)*C1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        let d1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 4 }))
            .expect("d1")
            .value
            .clone();
        let e1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 5 }))
            .expect("e1")
            .value
            .clone();
        assert_eq!(d1, CellValue::Number(14.0));
        assert_eq!(e1, CellValue::Number(20.0));
    }

    #[test]
    fn evaluates_builtin_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(8.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(3.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=SUM(A1,B1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            formula: "=MIN(A1,B1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            formula: "=MAX(A1,B1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 6,
            formula: "=IF(A1-B1,10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 7,
            formula: "=AVERAGE(A1,B1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 8,
            formula: "=ABS(B1-A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 9,
            formula: "=AND(A1,B1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 10,
            formula: "=OR(0,0,B1-A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 11,
            formula: "=NOT(B1-A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 12,
            formula: "=NOT(0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            value: CellValue::Text("RootCellar".to_string()),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 13,
            formula: "=LEN(A2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 14,
            formula: "=CHOOSE(2,A1,B1,99)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 15,
            formula: "=MATCH(B1,A1,B1,7)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 16,
            formula: "=DATE(2026,3,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 17,
            formula: "=YEAR(P1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 18,
            formula: "=MONTH(P1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 19,
            formula: "=DAY(P1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 20,
            formula: "=YEAR(60)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 21,
            formula: "=MONTH(60)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 22,
            formula: "=DAY(60)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        let c1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 3 }))
            .expect("c1")
            .value
            .clone();
        let d1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 4 }))
            .expect("d1")
            .value
            .clone();
        let e1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 5 }))
            .expect("e1")
            .value
            .clone();
        let f1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 6 }))
            .expect("f1")
            .value
            .clone();
        let g1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 7 }))
            .expect("g1")
            .value
            .clone();
        let h1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 8 }))
            .expect("h1")
            .value
            .clone();
        let i1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 9 }))
            .expect("i1")
            .value
            .clone();
        let j1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 10 }))
            .expect("j1")
            .value
            .clone();
        let k1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 11 }))
            .expect("k1")
            .value
            .clone();
        let l1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 12 }))
            .expect("l1")
            .value
            .clone();
        let m1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 13 }))
            .expect("m1")
            .value
            .clone();
        let n1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 14 }))
            .expect("n1")
            .value
            .clone();
        let o1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 15 }))
            .expect("o1")
            .value
            .clone();
        let p1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 16 }))
            .expect("p1")
            .value
            .clone();
        let q1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 17 }))
            .expect("q1")
            .value
            .clone();
        let r1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 18 }))
            .expect("r1")
            .value
            .clone();
        let s1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 19 }))
            .expect("s1")
            .value
            .clone();
        let t1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 20 }))
            .expect("t1")
            .value
            .clone();
        let u1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 21 }))
            .expect("u1")
            .value
            .clone();
        let v1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 22 }))
            .expect("v1")
            .value
            .clone();
        assert_eq!(c1, CellValue::Number(13.0));
        assert_eq!(d1, CellValue::Number(2.0));
        assert_eq!(e1, CellValue::Number(8.0));
        assert_eq!(f1, CellValue::Number(10.0));
        assert_eq!(g1, CellValue::Number(4.0));
        assert_eq!(h1, CellValue::Number(5.0));
        assert_eq!(i1, CellValue::Number(1.0));
        assert_eq!(j1, CellValue::Number(1.0));
        assert_eq!(k1, CellValue::Number(0.0));
        assert_eq!(l1, CellValue::Number(1.0));
        assert_eq!(m1, CellValue::Number(10.0));
        assert_eq!(n1, CellValue::Number(3.0));
        assert_eq!(o1, CellValue::Number(2.0));
        assert_eq!(p1, CellValue::Number(46082.0));
        assert_eq!(q1, CellValue::Number(2026.0));
        assert_eq!(r1, CellValue::Number(3.0));
        assert_eq!(s1, CellValue::Number(1.0));
        assert_eq!(t1, CellValue::Number(1900.0));
        assert_eq!(u1, CellValue::Number(2.0));
        assert_eq!(v1, CellValue::Number(29.0));
    }

    #[test]
    fn unknown_function_yields_parse_error() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            formula: "=NOPE(1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 1);
        let a1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 1 }))
            .expect("a1")
            .value
            .clone();
        assert_eq!(a1, CellValue::Error("#PARSE!".to_string()));
    }

    #[test]
    fn evaluates_text_and_date_extension_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Text("Cellar".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Text("RootCellar".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Text("root".to_string()),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=EXACT(A1,A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=EXACT(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=FIND(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=SEARCH(C1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=DAYS(DATE(2026,3,1),DATE(2026,2,27))".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=WEEKDAY(DATE(2026,3,1))".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=WEEKDAY(DATE(2026,3,1),2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=WEEKDAY(DATE(2026,3,1),3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=EDATE(DATE(2026,3,1),1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=EOMONTH(DATE(2026,3,1),0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=WEEKNUM(DATE(2026,3,1))".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=WEEKNUM(DATE(2026,3,1),2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=ISOWEEKNUM(DATE(2026,3,1))".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=CODE(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        let a2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 1 }))
            .expect("a2")
            .value
            .clone();
        let b2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 2 }))
            .expect("b2")
            .value
            .clone();
        let c2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 3 }))
            .expect("c2")
            .value
            .clone();
        let d2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 4 }))
            .expect("d2")
            .value
            .clone();
        let e2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 5 }))
            .expect("e2")
            .value
            .clone();
        let f2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 6 }))
            .expect("f2")
            .value
            .clone();
        let g2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 7 }))
            .expect("g2")
            .value
            .clone();
        let h2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 8 }))
            .expect("h2")
            .value
            .clone();
        let i2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 9 }))
            .expect("i2")
            .value
            .clone();
        let j2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 10 }))
            .expect("j2")
            .value
            .clone();
        let k2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 11 }))
            .expect("k2")
            .value
            .clone();
        let l2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 12 }))
            .expect("l2")
            .value
            .clone();
        let m2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 13 }))
            .expect("m2")
            .value
            .clone();
        let n2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 14 }))
            .expect("n2")
            .value
            .clone();

        assert_eq!(a2, CellValue::Number(1.0));
        assert_eq!(b2, CellValue::Number(0.0));
        assert_eq!(c2, CellValue::Number(5.0));
        assert_eq!(d2, CellValue::Number(1.0));
        assert_eq!(e2, CellValue::Number(2.0));
        assert_eq!(f2, CellValue::Number(1.0));
        assert_eq!(g2, CellValue::Number(7.0));
        assert_eq!(h2, CellValue::Number(6.0));
        assert_eq!(i2, CellValue::Number(46113.0));
        assert_eq!(j2, CellValue::Number(46112.0));
        assert_eq!(k2, CellValue::Number(10.0));
        assert_eq!(l2, CellValue::Number(9.0));
        assert_eq!(m2, CellValue::Number(9.0));
        assert_eq!(n2, CellValue::Number(67.0));
    }

    #[test]
    fn find_and_weekday_invalid_inputs_yield_parse_error() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Text("missing".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Text("RootCellar".to_string()),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=FIND(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=WEEKDAY(1,4)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=WEEKNUM(1,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 3);

        let a2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 1 }))
            .expect("a2")
            .value
            .clone();
        let b2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 2 }))
            .expect("b2")
            .value
            .clone();
        let c2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 3 }))
            .expect("c2")
            .value
            .clone();
        assert_eq!(a2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(b2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(c2, CellValue::Error("#PARSE!".to_string()));
    }

    #[test]
    fn evaluates_value_and_type_probe_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(42.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Text("42.5".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Bool(true),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Text("abc".to_string()),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            formula: "=1/0".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=N(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=N(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=N(C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=VALUE(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=ISNUMBER(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=ISNUMBER(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=ISTEXT(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=ISBLANK(Z1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=ISLOGICAL(C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=ISERROR(E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=ISERROR(1/0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=VALUE(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=VALUE(D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 1);

        let a2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 1 }))
            .expect("a2")
            .value
            .clone();
        let b2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 2 }))
            .expect("b2")
            .value
            .clone();
        let c2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 3 }))
            .expect("c2")
            .value
            .clone();
        let d2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 4 }))
            .expect("d2")
            .value
            .clone();
        let e2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 5 }))
            .expect("e2")
            .value
            .clone();
        let f2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 6 }))
            .expect("f2")
            .value
            .clone();
        let g2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 7 }))
            .expect("g2")
            .value
            .clone();
        let h2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 8 }))
            .expect("h2")
            .value
            .clone();
        let i2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 9 }))
            .expect("i2")
            .value
            .clone();
        let j2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 10 }))
            .expect("j2")
            .value
            .clone();
        let k2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 11 }))
            .expect("k2")
            .value
            .clone();
        let l2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 12 }))
            .expect("l2")
            .value
            .clone();
        let m2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 13 }))
            .expect("m2")
            .value
            .clone();
        assert_eq!(a2, CellValue::Number(42.0));
        assert_eq!(b2, CellValue::Number(0.0));
        assert_eq!(c2, CellValue::Number(1.0));
        assert_eq!(d2, CellValue::Number(42.5));
        assert_eq!(e2, CellValue::Number(1.0));
        assert_eq!(f2, CellValue::Number(0.0));
        assert_eq!(g2, CellValue::Number(1.0));
        assert_eq!(h2, CellValue::Number(1.0));
        assert_eq!(i2, CellValue::Number(1.0));
        assert_eq!(j2, CellValue::Number(1.0));
        assert_eq!(k2, CellValue::Number(1.0));
        assert_eq!(l2, CellValue::Number(42.0));
        assert_eq!(m2, CellValue::Error("#PARSE!".to_string()));
    }

    #[test]
    fn evaluates_math_extension_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(12.345),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(-12.345),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Number(10.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Number(3.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            value: CellValue::Number(123.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=ROUND(A1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=ROUND(B1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=ROUND(E1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=ROUNDUP(B1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=ROUNDDOWN(B1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=INT(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=MOD(C1,D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=MOD(B1,D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=POWER(2,10)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=SQRT(81)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=PRODUCT(C1,D1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=ROUNDUP(E1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=ROUNDDOWN(E1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=MOD(5,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=SQRT(-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 16,
            formula: "=ROUND(1,2,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 17,
            formula: "=POWER(-1,0.5)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 4);

        let a2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 1 }))
            .expect("a2")
            .value
            .clone();
        let b2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 2 }))
            .expect("b2")
            .value
            .clone();
        let c2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 3 }))
            .expect("c2")
            .value
            .clone();
        let d2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 4 }))
            .expect("d2")
            .value
            .clone();
        let e2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 5 }))
            .expect("e2")
            .value
            .clone();
        let f2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 6 }))
            .expect("f2")
            .value
            .clone();
        let g2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 7 }))
            .expect("g2")
            .value
            .clone();
        let h2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 8 }))
            .expect("h2")
            .value
            .clone();
        let i2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 9 }))
            .expect("i2")
            .value
            .clone();
        let j2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 10 }))
            .expect("j2")
            .value
            .clone();
        let k2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 11 }))
            .expect("k2")
            .value
            .clone();
        let l2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 12 }))
            .expect("l2")
            .value
            .clone();
        let m2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 13 }))
            .expect("m2")
            .value
            .clone();
        let n2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 14 }))
            .expect("n2")
            .value
            .clone();
        let o2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 15 }))
            .expect("o2")
            .value
            .clone();
        let p2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 16 }))
            .expect("p2")
            .value
            .clone();
        let q2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 17 }))
            .expect("q2")
            .value
            .clone();

        assert_eq!(a2, CellValue::Number(12.35));
        assert_eq!(b2, CellValue::Number(-12.35));
        assert_eq!(c2, CellValue::Number(120.0));
        assert_eq!(d2, CellValue::Number(-12.4));
        assert_eq!(e2, CellValue::Number(-12.3));
        assert_eq!(f2, CellValue::Number(-13.0));
        assert_eq!(g2, CellValue::Number(1.0));
        assert_eq!(i2, CellValue::Number(1024.0));
        assert_eq!(j2, CellValue::Number(9.0));
        assert_eq!(k2, CellValue::Number(60.0));
        assert_eq!(l2, CellValue::Number(130.0));
        assert_eq!(m2, CellValue::Number(120.0));
        assert_eq!(n2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(o2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(p2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(q2, CellValue::Error("#PARSE!".to_string()));

        let expected_mod_negative = -12.345_f64 - 3.0_f64 * (-12.345_f64 / 3.0_f64).floor();
        match h2 {
            CellValue::Number(value) => {
                assert!((value - expected_mod_negative).abs() < 1e-12);
            }
            other => panic!("expected numeric H2 value, got {other:?}"),
        }
    }

    #[test]
    fn evaluates_stat_and_error_extension_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(10.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(20.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Number(30.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Number(-3.2),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            value: CellValue::Number(3.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=IFERROR(1/0,99)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=IFERROR(A1/2,99)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=MEDIAN(A1,B1,C1,40)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=SMALL(A1,B1,C1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=LARGE(A1,B1,C1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=SIGN(D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=SIGN(0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=CEILING(D1,E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=FLOOR(D1,E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=CEILING(7.1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=FLOOR(7.9,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=MEDIAN(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=SMALL(A1,B1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=LARGE(A1,B1,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=CEILING(5,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 16,
            formula: "=IFERROR(NOPE(1),A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 17,
            formula: "=IFERROR(A1,1/0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 18,
            formula: "=CEILING(2.2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 3);

        let a2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 1 }))
            .expect("a2")
            .value
            .clone();
        let b2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 2 }))
            .expect("b2")
            .value
            .clone();
        let c2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 3 }))
            .expect("c2")
            .value
            .clone();
        let d2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 4 }))
            .expect("d2")
            .value
            .clone();
        let e2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 5 }))
            .expect("e2")
            .value
            .clone();
        let f2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 6 }))
            .expect("f2")
            .value
            .clone();
        let g2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 7 }))
            .expect("g2")
            .value
            .clone();
        let h2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 8 }))
            .expect("h2")
            .value
            .clone();
        let i2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 9 }))
            .expect("i2")
            .value
            .clone();
        let j2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 10 }))
            .expect("j2")
            .value
            .clone();
        let k2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 11 }))
            .expect("k2")
            .value
            .clone();
        let l2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 12 }))
            .expect("l2")
            .value
            .clone();
        let m2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 13 }))
            .expect("m2")
            .value
            .clone();
        let n2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 14 }))
            .expect("n2")
            .value
            .clone();
        let o2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 15 }))
            .expect("o2")
            .value
            .clone();
        let p2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 16 }))
            .expect("p2")
            .value
            .clone();
        let q2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 17 }))
            .expect("q2")
            .value
            .clone();
        let r2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 18 }))
            .expect("r2")
            .value
            .clone();

        assert_eq!(a2, CellValue::Number(99.0));
        assert_eq!(b2, CellValue::Number(5.0));
        assert_eq!(c2, CellValue::Number(25.0));
        assert_eq!(d2, CellValue::Number(20.0));
        assert_eq!(e2, CellValue::Number(20.0));
        assert_eq!(f2, CellValue::Number(-1.0));
        assert_eq!(g2, CellValue::Number(0.0));
        assert_eq!(h2, CellValue::Number(-3.0));
        assert_eq!(i2, CellValue::Number(-6.0));
        assert_eq!(j2, CellValue::Number(8.0));
        assert_eq!(k2, CellValue::Number(6.0));
        assert_eq!(l2, CellValue::Number(10.0));
        assert_eq!(m2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(n2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(o2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(p2, CellValue::Number(10.0));
        assert_eq!(q2, CellValue::Number(10.0));
        assert_eq!(r2, CellValue::Number(3.0));
    }

    #[test]
    fn evaluates_advanced_rounding_and_parity_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(17.8),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(-17.8),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Number(5.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Number(-5.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            value: CellValue::Number(2.5),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=QUOTIENT(A1,E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=QUOTIENT(B1,E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=TRUNC(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=TRUNC(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=TRUNC(1234,-2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=TRUNC(987.65,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=MROUND(11,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=MROUND(-11,-2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=MROUND(5,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=EVEN(3.2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=EVEN(-3.2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=ODD(4)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=ODD(-4)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=ODD(0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=EVEN(0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 16,
            formula: "=ISEVEN(2.9)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 17,
            formula: "=ISODD(2.9)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 18,
            formula: "=ISEVEN(3.1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 19,
            formula: "=ISODD(3.1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 20,
            formula: "=MROUND(5,-2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 21,
            formula: "=QUOTIENT(1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 22,
            formula: "=TRUNC(1,2,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 3);

        let a2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 1 }))
            .expect("a2")
            .value
            .clone();
        let b2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 2 }))
            .expect("b2")
            .value
            .clone();
        let c2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 3 }))
            .expect("c2")
            .value
            .clone();
        let d2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 4 }))
            .expect("d2")
            .value
            .clone();
        let e2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 5 }))
            .expect("e2")
            .value
            .clone();
        let f2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 6 }))
            .expect("f2")
            .value
            .clone();
        let g2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 7 }))
            .expect("g2")
            .value
            .clone();
        let h2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 8 }))
            .expect("h2")
            .value
            .clone();
        let i2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 9 }))
            .expect("i2")
            .value
            .clone();
        let j2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 10 }))
            .expect("j2")
            .value
            .clone();
        let k2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 11 }))
            .expect("k2")
            .value
            .clone();
        let l2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 12 }))
            .expect("l2")
            .value
            .clone();
        let m2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 13 }))
            .expect("m2")
            .value
            .clone();
        let n2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 14 }))
            .expect("n2")
            .value
            .clone();
        let o2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 15 }))
            .expect("o2")
            .value
            .clone();
        let p2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 16 }))
            .expect("p2")
            .value
            .clone();
        let q2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 17 }))
            .expect("q2")
            .value
            .clone();
        let r2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 18 }))
            .expect("r2")
            .value
            .clone();
        let s2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 19 }))
            .expect("s2")
            .value
            .clone();
        let t2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 20 }))
            .expect("t2")
            .value
            .clone();
        let u2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 21 }))
            .expect("u2")
            .value
            .clone();
        let v2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 22 }))
            .expect("v2")
            .value
            .clone();

        assert_eq!(a2, CellValue::Number(7.0));
        assert_eq!(b2, CellValue::Number(-7.0));
        assert_eq!(c2, CellValue::Number(17.0));
        assert_eq!(d2, CellValue::Number(-17.0));
        assert_eq!(e2, CellValue::Number(1200.0));
        assert_eq!(f2, CellValue::Number(980.0));
        assert_eq!(g2, CellValue::Number(12.0));
        assert_eq!(h2, CellValue::Number(-12.0));
        assert_eq!(i2, CellValue::Number(0.0));
        assert_eq!(j2, CellValue::Number(4.0));
        assert_eq!(k2, CellValue::Number(-4.0));
        assert_eq!(l2, CellValue::Number(5.0));
        assert_eq!(m2, CellValue::Number(-5.0));
        assert_eq!(n2, CellValue::Number(1.0));
        assert_eq!(o2, CellValue::Number(0.0));
        assert_eq!(p2, CellValue::Number(1.0));
        assert_eq!(q2, CellValue::Number(0.0));
        assert_eq!(r2, CellValue::Number(0.0));
        assert_eq!(s2, CellValue::Number(1.0));
        assert_eq!(t2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(u2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(v2, CellValue::Error("#PARSE!".to_string()));
    }

    #[test]
    fn evaluates_count_and_time_extension_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(42.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Text("2026-03-01".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Text("18:45:30".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Text("not-a-date".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            value: CellValue::Bool(true),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=DATEVALUE(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=TIME(18,45,30)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=HOUR(B2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=MINUTE(B2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=SECOND(B2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=TIMEVALUE(C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=HOUR(C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=COUNT(A1,B1,Z1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=COUNTA(A1,B1,Z1,E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=COUNTBLANK(A1,B1,Z1,E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=DATEVALUE(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=TIMEVALUE(B2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=DATEVALUE(D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=TIMEVALUE(D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=COUNT(1/0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 3);

        let a2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 1 }))
            .expect("a2")
            .value
            .clone();
        let b2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 2 }))
            .expect("b2")
            .value
            .clone();
        let c2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 3 }))
            .expect("c2")
            .value
            .clone();
        let d2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 4 }))
            .expect("d2")
            .value
            .clone();
        let e2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 5 }))
            .expect("e2")
            .value
            .clone();
        let f2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 6 }))
            .expect("f2")
            .value
            .clone();
        let g2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 7 }))
            .expect("g2")
            .value
            .clone();
        let h2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 8 }))
            .expect("h2")
            .value
            .clone();
        let i2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 9 }))
            .expect("i2")
            .value
            .clone();
        let j2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 10 }))
            .expect("j2")
            .value
            .clone();
        let k2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 11 }))
            .expect("k2")
            .value
            .clone();
        let l2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 12 }))
            .expect("l2")
            .value
            .clone();
        let m2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 13 }))
            .expect("m2")
            .value
            .clone();
        let n2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 14 }))
            .expect("n2")
            .value
            .clone();
        let o2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 15 }))
            .expect("o2")
            .value
            .clone();

        assert_eq!(a2, CellValue::Number(46082.0));
        assert_eq!(c2, CellValue::Number(18.0));
        assert_eq!(d2, CellValue::Number(45.0));
        assert_eq!(e2, CellValue::Number(30.0));
        assert_eq!(g2, CellValue::Number(18.0));
        assert_eq!(h2, CellValue::Number(1.0));
        assert_eq!(i2, CellValue::Number(3.0));
        assert_eq!(j2, CellValue::Number(1.0));
        assert_eq!(k2, CellValue::Number(42.0));
        assert_eq!(m2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(n2, CellValue::Error("#PARSE!".to_string()));
        assert_eq!(o2, CellValue::Error("#PARSE!".to_string()));

        let expected_time = (18.0 * 3600.0 + 45.0 * 60.0 + 30.0) / 86_400.0;
        match b2 {
            CellValue::Number(value) => {
                assert!((value - expected_time).abs() < 1e-12);
            }
            other => panic!("expected numeric B2 value, got {other:?}"),
        }
        match f2 {
            CellValue::Number(value) => {
                assert!((value - expected_time).abs() < 1e-12);
            }
            other => panic!("expected numeric F2 value, got {other:?}"),
        }
        match l2 {
            CellValue::Number(value) => {
                assert!((value - expected_time).abs() < 1e-12);
            }
            other => panic!("expected numeric L2 value, got {other:?}"),
        }
    }

    #[test]
    fn evaluates_index_and_match_mode_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(10.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(20.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Number(30.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Number(40.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=INDEX(3,A1,B1,C1,D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=MATCH(25,A1,B1,C1,D1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=MATCH(25,A1,B1,C1,D1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=MATCH(30,A1,B1,C1,D1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=XMATCH(25,A1,B1,C1,D1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=XMATCH(25,A1,B1,C1,D1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=XMATCH(30,A1,B1,C1,D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        let a2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 1 }))
            .expect("a2")
            .value
            .clone();
        let b2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 2 }))
            .expect("b2")
            .value
            .clone();
        let c2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 3 }))
            .expect("c2")
            .value
            .clone();
        let d2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 4 }))
            .expect("d2")
            .value
            .clone();
        let e2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 5 }))
            .expect("e2")
            .value
            .clone();
        let f2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 6 }))
            .expect("f2")
            .value
            .clone();
        let g2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 7 }))
            .expect("g2")
            .value
            .clone();
        assert_eq!(a2, CellValue::Number(30.0));
        assert_eq!(b2, CellValue::Number(2.0));
        assert_eq!(c2, CellValue::Number(3.0));
        assert_eq!(d2, CellValue::Number(3.0));
        assert_eq!(e2, CellValue::Number(2.0));
        assert_eq!(f2, CellValue::Number(3.0));
        assert_eq!(g2, CellValue::Number(3.0));
    }

    #[test]
    fn analyzes_dependency_graph() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(2.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            formula: "=A1+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=B1*2".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let graph = analyze_sheet_dependencies(&wb, "Sheet1", &mut sink, &trace).expect("graph");
        assert_eq!(graph.formula_cell_count, 2);
        assert_eq!(graph.function_call_count, 0);
        assert!(graph.ast_node_count >= graph.ast_unique_node_count);
        assert_eq!(graph.formula_ast_ids.len(), 2);
        assert_eq!(graph.dependency_edge_count, 2);
        assert_eq!(graph.formula_edge_count, 1);
        assert_eq!(graph.topo_order, vec!["B1".to_string(), "C1".to_string()]);
        assert!(graph.cyclic_cells.is_empty());
        assert!(graph.parse_error_cells.is_empty());
    }

    #[test]
    fn reports_cyclic_cells_in_dependency_graph() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            formula: "=B1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            formula: "=A1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let graph = analyze_sheet_dependencies(&wb, "Sheet1", &mut sink, &trace).expect("graph");
        assert_eq!(graph.formula_cell_count, 2);
        assert_eq!(graph.function_call_count, 0);
        assert!(graph.ast_node_count >= graph.ast_unique_node_count);
        assert_eq!(graph.formula_ast_ids.len(), 2);
        assert_eq!(graph.formula_edge_count, 2);
        assert!(graph.topo_order.is_empty());
        assert_eq!(graph.cyclic_cells, vec!["A1".to_string(), "B1".to_string()]);
    }

    #[test]
    fn dependency_graph_counts_function_calls() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(2.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(4.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=SUM(A1,MAX(B1,3))".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let graph = analyze_sheet_dependencies(&wb, "Sheet1", &mut sink, &trace).expect("graph");
        assert_eq!(graph.formula_cell_count, 1);
        assert_eq!(graph.function_call_count, 2);
        assert_eq!(graph.formula_ast_ids.len(), 1);
        assert!(graph.ast_node_count >= graph.ast_unique_node_count);
        assert_eq!(graph.dependency_edge_count, 2);
    }

    #[test]
    fn ast_interning_deduplicates_repeated_formulas() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(1.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(2.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=SUM(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            formula: "=SUM(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let graph = analyze_sheet_dependencies(&wb, "Sheet1", &mut sink, &trace).expect("graph");
        let c1_id = graph.formula_ast_ids.get("C1").copied().expect("C1 ast id");
        let d1_id = graph.formula_ast_ids.get("D1").copied().expect("D1 ast id");
        assert_eq!(c1_id, d1_id);
        assert!(graph.ast_unique_node_count < graph.ast_node_count);
    }

    #[test]
    fn incremental_recalc_recomputes_only_impacted_formulas() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(10.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(5.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=A1+B1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            formula: "=C1*2".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            formula: "=B1*3".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let full = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("full");
        assert_eq!(full.evaluated_cells, 3);

        let mut txn2 = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn2.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(20.0),
        });
        txn2.commit(&mut wb, &mut sink, &trace).expect("commit");

        let changed_roots = [CellRef { row: 1, col: 1 }];
        let incremental =
            recalc_sheet_from_roots(&mut wb, "Sheet1", &changed_roots, &mut sink, &trace)
                .expect("incremental");
        assert_eq!(incremental.evaluated_cells, 2);

        let d1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 4 }))
            .expect("d1")
            .value
            .clone();
        let e1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 5 }))
            .expect("e1")
            .value
            .clone();
        assert_eq!(d1, CellValue::Number(50.0));
        assert_eq!(e1, CellValue::Number(15.0));
    }

    #[test]
    fn incremental_recalc_includes_formula_root_and_dependents() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(4.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            formula: "=A1+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=B1*2".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");
        recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("full");

        let mut txn2 = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn2.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            formula: "=A1+2".to_string(),
            cached_value: CellValue::Empty,
        });
        txn2.commit(&mut wb, &mut sink, &trace).expect("commit");

        let changed_roots = [CellRef { row: 1, col: 2 }];
        let incremental =
            recalc_sheet_from_roots(&mut wb, "Sheet1", &changed_roots, &mut sink, &trace)
                .expect("incremental");
        assert_eq!(incremental.evaluated_cells, 2);

        let b1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 2 }))
            .expect("b1")
            .value
            .clone();
        let c1 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 1, col: 3 }))
            .expect("c1")
            .value
            .clone();
        assert_eq!(b1, CellValue::Number(6.0));
        assert_eq!(c1, CellValue::Number(12.0));
    }

    #[test]
    fn recalc_dag_timing_report_contains_node_timings() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(2.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            formula: "=A1+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=B1*2".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let (recalc_report, dag) =
            recalc_sheet_with_dag_timing(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(recalc_report.evaluated_cells, 2);
        assert_eq!(dag.mode, "full");
        assert_eq!(dag.formula_cell_count, 2);
        assert_eq!(dag.evaluated_cells, 2);
        assert_eq!(dag.node_timings.len(), 2);
        assert_eq!(dag.node_degrees.len(), 2);
        assert_eq!(dag.max_fan_in, 1);
        assert_eq!(dag.max_fan_out, 1);
        assert!(dag.critical_path.contains(&"C1".to_string()));
        assert!(dag.critical_path_duration_us >= dag.max_node_duration_us);
        assert!(dag
            .slow_nodes
            .iter()
            .all(|node| node.duration_us >= dag.slow_nodes_threshold_us));
        assert!(dag.node_timings.iter().all(|node| node.status == "ok"));
        assert!(dag.total_node_duration_us >= dag.max_node_duration_us);
    }

    #[test]
    fn incremental_dag_timing_report_tracks_changed_roots() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(10.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(5.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=A1+B1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            formula: "=C1*2".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");
        recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("full");

        let mut txn2 = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn2.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(20.0),
        });
        txn2.commit(&mut wb, &mut sink, &trace).expect("commit");

        let changed_roots = [CellRef { row: 1, col: 1 }];
        let (_report, dag) = recalc_sheet_from_roots_with_dag_timing(
            &mut wb,
            "Sheet1",
            &changed_roots,
            &mut sink,
            &trace,
        )
        .expect("incremental");
        assert_eq!(dag.mode, "incremental");
        assert_eq!(dag.changed_root_count, Some(1));
        assert_eq!(dag.evaluated_cells, 2);
        assert_eq!(dag.node_timings.len(), 2);
        assert_eq!(dag.node_degrees.len(), 2);
        assert_eq!(dag.max_fan_in, 1);
        assert_eq!(dag.max_fan_out, 1);
        assert!(dag.critical_path_duration_us >= dag.max_node_duration_us);
        assert!(dag
            .slow_nodes
            .iter()
            .all(|node| node.duration_us >= dag.slow_nodes_threshold_us));
    }

    #[test]
    fn dag_timing_supports_slow_node_threshold_override() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(2.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            formula: "=A1+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=B1*2".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let options = RecalcDagTimingOptions {
            slow_nodes_threshold_us: Some(u64::MAX),
        };
        let (_recalc_report, dag) =
            recalc_sheet_with_dag_timing_options(&mut wb, "Sheet1", options, &mut sink, &trace)
                .expect("recalc");
        assert_eq!(dag.slow_nodes_threshold_us, u64::MAX);
        assert!(dag.slow_nodes.is_empty());
    }
}
