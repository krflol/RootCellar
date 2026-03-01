use crate::model::{CellRef, CellValue, Workbook};
use crate::telemetry::{EventEnvelope, EventSink, TelemetryError, TraceContext};
use chrono::{Datelike, Duration, NaiveDate, NaiveTime, Timelike};
use serde::{Deserialize, Serialize};
use serde_json::json;
use std::collections::{BTreeMap, BTreeSet, HashSet, VecDeque};
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

type EvalCache = BTreeMap<CellRef, Result<CellValue, EvalError>>;

#[derive(Debug, Clone)]
enum Expr {
    Number(f64),
    Text(String),
    Bool(bool),
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
    dependents_by_ref: BTreeMap<CellRef, Vec<CellRef>>,
    formula_nodes: BTreeSet<CellRef>,
    function_call_count: usize,
    ast_node_count: usize,
    ast_unique_node_count: usize,
    formula_ast_ids: BTreeMap<CellRef, u32>,
    ast_intern_nodes: BTreeMap<u32, String>,
    dependency_edge_count: usize,
    formula_edge_count: usize,
    topo_order: Vec<CellRef>,
    topo_position_by_cell: BTreeMap<CellRef, usize>,
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
            Expr::Text(text) => format!("text:{text:?}"),
            Expr::Bool(value) => format!("bool:{value}"),
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
    Ok(recalc_sheet_impl(
        workbook,
        sheet_name,
        Some(changed_roots),
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
    let result = recalc_sheet_impl(
        workbook,
        sheet_name,
        Some(changed_roots),
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
    changed_roots: Option<&[CellRef]>,
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

    let target_formula_cells = if let Some(roots) = changed_roots {
        let impacted_formula_cells = collect_impacted_formula_cells(&analysis, roots);
        order_formula_cells(&analysis, &impacted_formula_cells)
    } else {
        order_all_formula_cells(&analysis)
    };

    let mut cache = BTreeMap::<CellRef, Result<CellValue, EvalError>>::new();
    let mut cycle_count = 0usize;
    let mut parse_error_count = 0usize;
    let mut duration_by_cell = if capture_dag_timing {
        Some(BTreeMap::<CellRef, u64>::new())
    } else {
        None
    };
    let mut node_timings = if capture_dag_timing {
        Some(Vec::<RecalcDagNodeTiming>::new())
    } else {
        None
    };
    let mut stack = BTreeSet::new();

    for cell_ref in &target_formula_cells {
        stack.clear();
        let eval_started = Instant::now();
        let result = eval_cell_value(
            workbook,
            sheet_name,
            *cell_ref,
            &analysis.parsed_formulas,
            &mut stack,
            &mut cache,
        );
        let duration_us = eval_started.elapsed().as_micros().min(u128::from(u64::MAX)) as u64;
        if let Some(duration_map) = duration_by_cell.as_mut() {
            duration_map.insert(*cell_ref, duration_us);
        }
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
                        cell.value = value;
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
            duration_by_cell
                .as_ref()
                .expect("duration map must be present when capture_dag_timing=true"),
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
    let mut dependents_by_ref = BTreeMap::<CellRef, Vec<CellRef>>::new();
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
                .push(*cell_ref);
        }
        parsed_formulas.insert(*cell_ref, parsed);
        dependency_refs.insert(*cell_ref, refs);
    }

    let (topo_order, formula_edge_count, cyclic_cells) =
        build_formula_topological_order(&formula_set, &dependents_by_ref);
    let topo_position_by_cell = topo_order
        .iter()
        .enumerate()
        .map(|(idx, cell)| (*cell, idx))
        .collect::<BTreeMap<_, _>>();

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
        topo_position_by_cell,
        cyclic_cells,
        parse_error_cells,
    })
}

fn collect_impacted_formula_cells(
    analysis: &SheetDependencyAnalysis,
    changed_roots: &[CellRef],
) -> Vec<CellRef> {
    let mut impacted = Vec::<CellRef>::new();
    let mut seen = HashSet::<CellRef>::with_capacity(changed_roots.len());
    let mut queue = VecDeque::<CellRef>::with_capacity(changed_roots.len());
    for root in changed_roots {
        if analysis.formula_nodes.contains(root) && seen.insert(*root) {
            impacted.push(*root);
            queue.push_back(*root);
        }
        if let Some(dependents) = analysis.dependents_by_ref.get(root) {
            for dependent in dependents {
                if seen.insert(*dependent) {
                    impacted.push(*dependent);
                    queue.push_back(*dependent);
                }
            }
        }
    }

    while let Some(node) = queue.pop_front() {
        if let Some(dependents) = analysis.dependents_by_ref.get(&node) {
            for dependent in dependents {
                if seen.insert(*dependent) {
                    impacted.push(*dependent);
                    queue.push_back(*dependent);
                }
            }
        }
    }

    impacted
}

fn order_formula_cells(
    analysis: &SheetDependencyAnalysis,
    target_cells: &[CellRef],
) -> Vec<CellRef> {
    let mut in_topo = Vec::<(usize, CellRef)>::with_capacity(target_cells.len());
    let mut remaining = Vec::<CellRef>::new();
    for cell in target_cells {
        if let Some(position) = analysis.topo_position_by_cell.get(cell).copied() {
            in_topo.push((position, *cell));
        } else {
            remaining.push(*cell);
        }
    }

    in_topo.sort_by(|(pos_a, cell_a), (pos_b, cell_b)| {
        pos_a.cmp(pos_b).then_with(|| cell_a.cmp(cell_b))
    });
    remaining.sort();

    let mut ordered = in_topo
        .into_iter()
        .map(|(_, cell)| cell)
        .collect::<Vec<_>>();
    ordered.extend(remaining);
    ordered
}

fn order_all_formula_cells(analysis: &SheetDependencyAnalysis) -> Vec<CellRef> {
    let mut ordered = Vec::<CellRef>::with_capacity(analysis.formula_nodes.len());
    ordered.extend(analysis.topo_order.iter().copied());
    ordered.extend(analysis.cyclic_cells.iter().copied());
    ordered
}

fn build_formula_topological_order(
    formula_set: &BTreeSet<CellRef>,
    dependents_by_ref: &BTreeMap<CellRef, Vec<CellRef>>,
) -> (Vec<CellRef>, usize, Vec<CellRef>) {
    let mut indegree = formula_set
        .iter()
        .copied()
        .map(|cell| (cell, 0usize))
        .collect::<BTreeMap<_, _>>();
    let mut formula_edge_count = 0usize;

    for (referenced, dependents) in dependents_by_ref {
        if !formula_set.contains(referenced) {
            continue;
        }
        for dependent in dependents {
            if let Some(entry) = indegree.get_mut(dependent) {
                *entry += 1;
            }
            formula_edge_count += 1;
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

        if let Some(dependents) = dependents_by_ref.get(&next) {
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

    let cyclic_cells = indegree
        .iter()
        .filter_map(|(cell, degree)| if *degree > 0 { Some(*cell) } else { None })
        .collect::<Vec<_>>();

    (order, formula_edge_count, cyclic_cells)
}

fn emit_dependency_graph_event(
    sink: &mut dyn EventSink,
    span: &TraceContext,
    workbook_id: uuid::Uuid,
    sheet_name: &str,
    analysis: &SheetDependencyAnalysis,
) -> Result<(), CalcError> {
    let payload = if sink.supports_expensive_payloads() {
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

        json!({
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
        })
    } else {
        json!({
            "topo_order_preview_omitted_for_sink": true,
            "dependency_refs_omitted_for_sink": true,
            "ast_intern_preview_omitted_for_sink": true,
            "formula_ast_ids_omitted_for_sink": true,
        })
    };

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
            .with_payload(payload),
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

    let mut topo_ranked = target_set
        .iter()
        .filter_map(|cell| {
            analysis
                .topo_position_by_cell
                .get(cell)
                .copied()
                .map(|position| (position, *cell))
        })
        .collect::<Vec<_>>();
    topo_ranked.sort_by(|(pos_a, cell_a), (pos_b, cell_b)| {
        pos_a.cmp(pos_b).then_with(|| cell_a.cmp(cell_b))
    });
    let topo_nodes = topo_ranked
        .into_iter()
        .map(|(_, cell)| cell)
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

fn eval_cell_value(
    workbook: &Workbook,
    sheet_name: &str,
    cell_ref: CellRef,
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut EvalCache,
) -> Result<CellValue, EvalError> {
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
                        Ok(expr) => eval_expr_value(
                            workbook,
                            sheet_name,
                            expr,
                            parsed_formulas,
                            stack,
                            cache,
                        ),
                        Err(err) => Err(err.clone()),
                    }
                } else {
                    let parsed = parse_formula_expression(formula)?;
                    eval_expr_value(workbook, sheet_name, &parsed, parsed_formulas, stack, cache)
                }
            } else {
                Ok(cell.value.clone())
            }
        }
        None => Ok(CellValue::Empty),
    };

    stack.remove(&cell_ref);
    cache.insert(cell_ref, result.clone());
    result
}

fn eval_expr_value(
    workbook: &Workbook,
    sheet_name: &str,
    expr: &Expr,
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut EvalCache,
) -> Result<CellValue, EvalError> {
    match expr {
        Expr::Number(n) => Ok(CellValue::Number(*n)),
        Expr::Text(text) => Ok(CellValue::Text(text.clone())),
        Expr::Bool(value) => Ok(CellValue::Bool(*value)),
        Expr::Cell(cell_ref) => eval_cell_value(
            workbook,
            sheet_name,
            *cell_ref,
            parsed_formulas,
            stack,
            cache,
        ),
        Expr::Function { name, args } => eval_function_value(
            workbook,
            sheet_name,
            name,
            args,
            parsed_formulas,
            stack,
            cache,
        ),
        Expr::UnaryMinus(inner) => {
            let value =
                eval_expr_value(workbook, sheet_name, inner, parsed_formulas, stack, cache)?;
            Ok(CellValue::Number(-value_as_arithmetic_number(&value)?))
        }
        Expr::Binary { op, left, right } => {
            let lhs_value =
                eval_expr_value(workbook, sheet_name, left, parsed_formulas, stack, cache)?;
            let rhs_value =
                eval_expr_value(workbook, sheet_name, right, parsed_formulas, stack, cache)?;
            let lhs = value_as_arithmetic_number(&lhs_value)?;
            let rhs = value_as_arithmetic_number(&rhs_value)?;
            match op {
                BinaryOp::Add => Ok(CellValue::Number(lhs + rhs)),
                BinaryOp::Sub => Ok(CellValue::Number(lhs - rhs)),
                BinaryOp::Mul => Ok(CellValue::Number(lhs * rhs)),
                BinaryOp::Div => {
                    if rhs == 0.0 {
                        Err(EvalError::DivisionByZero)
                    } else {
                        Ok(CellValue::Number(lhs / rhs))
                    }
                }
            }
        }
    }
}

fn eval_expr(
    workbook: &Workbook,
    sheet_name: &str,
    expr: &Expr,
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut EvalCache,
) -> Result<f64, EvalError> {
    match expr {
        Expr::Number(n) => Ok(*n),
        Expr::Text(_) => Ok(0.0),
        Expr::Bool(value) => Ok(if *value { 1.0 } else { 0.0 }),
        Expr::Cell(cell_ref) => {
            let value = eval_cell_value(
                workbook,
                sheet_name,
                *cell_ref,
                parsed_formulas,
                stack,
                cache,
            )?;
            value_as_number(&value)
        }
        Expr::Function { name, args } => {
            let value = eval_function_value(
                workbook,
                sheet_name,
                name,
                args,
                parsed_formulas,
                stack,
                cache,
            )?;
            value_as_number(&value)
        }
        Expr::UnaryMinus(inner) => {
            let value =
                eval_expr_value(workbook, sheet_name, inner, parsed_formulas, stack, cache)?;
            Ok(-value_as_arithmetic_number(&value)?)
        }
        Expr::Binary { op, left, right } => {
            let lhs_value =
                eval_expr_value(workbook, sheet_name, left, parsed_formulas, stack, cache)?;
            let rhs_value =
                eval_expr_value(workbook, sheet_name, right, parsed_formulas, stack, cache)?;
            let lhs = value_as_arithmetic_number(&lhs_value)?;
            let rhs = value_as_arithmetic_number(&rhs_value)?;
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

fn switch_values_equal(left: &CellValue, right: &CellValue) -> Result<bool, EvalError> {
    match (left, right) {
        (CellValue::Error(_), _) | (_, CellValue::Error(_)) => Err(EvalError::Parse),
        (CellValue::Text(left_text), CellValue::Text(right_text)) => Ok(left_text == right_text),
        (CellValue::Text(_), _) | (_, CellValue::Text(_)) => Ok(false),
        _ => {
            let left_number = value_as_number(left)?;
            let right_number = value_as_number(right)?;
            Ok(left_number.total_cmp(&right_number).is_eq())
        }
    }
}

fn value_as_condition(value: &CellValue) -> Result<bool, EvalError> {
    match value {
        CellValue::Number(n) => Ok(*n != 0.0),
        CellValue::Bool(value) => Ok(*value),
        CellValue::Empty => Ok(false),
        CellValue::Text(text) => {
            let trimmed = text.trim();
            if trimmed.is_empty() {
                return Ok(false);
            }
            if trimmed.eq_ignore_ascii_case("TRUE") {
                return Ok(true);
            }
            if trimmed.eq_ignore_ascii_case("FALSE") {
                return Ok(false);
            }
            Ok(parse_text_number(trimmed)? != 0.0)
        }
        CellValue::Error(_) => Err(EvalError::Parse),
    }
}

fn eval_function_value(
    workbook: &Workbook,
    sheet_name: &str,
    name: &str,
    args: &[Expr],
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut EvalCache,
) -> Result<CellValue, EvalError> {
    match name {
        "IF" => {
            if args.len() < 2 || args.len() > 3 {
                return Err(EvalError::Parse);
            }
            let condition_value = eval_expr_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            if value_as_condition(&condition_value)? {
                eval_expr_value(
                    workbook,
                    sheet_name,
                    &args[1],
                    parsed_formulas,
                    stack,
                    cache,
                )
            } else if args.len() == 3 {
                eval_expr_value(
                    workbook,
                    sheet_name,
                    &args[2],
                    parsed_formulas,
                    stack,
                    cache,
                )
            } else {
                Ok(CellValue::Number(0.0))
            }
        }
        "IFERROR" => {
            if args.len() != 2 {
                return Err(EvalError::Parse);
            }
            match eval_expr_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            ) {
                Ok(CellValue::Error(_)) => eval_expr_value(
                    workbook,
                    sheet_name,
                    &args[1],
                    parsed_formulas,
                    stack,
                    cache,
                ),
                Ok(value) => Ok(value),
                Err(_) => eval_expr_value(
                    workbook,
                    sheet_name,
                    &args[1],
                    parsed_formulas,
                    stack,
                    cache,
                ),
            }
        }
        "IFS" => {
            if args.len() < 2 || args.len() % 2 != 0 {
                return Err(EvalError::Parse);
            }
            for pair in args.chunks(2) {
                let condition_value = eval_expr_value(
                    workbook,
                    sheet_name,
                    &pair[0],
                    parsed_formulas,
                    stack,
                    cache,
                )?;
                if value_as_condition(&condition_value)? {
                    return eval_expr_value(
                        workbook,
                        sheet_name,
                        &pair[1],
                        parsed_formulas,
                        stack,
                        cache,
                    );
                }
            }
            Err(EvalError::Parse)
        }
        "SWITCH" => {
            if args.len() < 3 {
                return Err(EvalError::Parse);
            }
            let expression_value = eval_expr_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let options = &args[1..];
            let has_default = options.len() % 2 == 1;
            let pair_end = if has_default {
                options.len() - 1
            } else {
                options.len()
            };
            for idx in (0..pair_end).step_by(2) {
                let case_value = eval_expr_value(
                    workbook,
                    sheet_name,
                    &options[idx],
                    parsed_formulas,
                    stack,
                    cache,
                )?;
                if switch_values_equal(&expression_value, &case_value)? {
                    return eval_expr_value(
                        workbook,
                        sheet_name,
                        &options[idx + 1],
                        parsed_formulas,
                        stack,
                        cache,
                    );
                }
            }
            if has_default {
                eval_expr_value(
                    workbook,
                    sheet_name,
                    &options[pair_end],
                    parsed_formulas,
                    stack,
                    cache,
                )
            } else {
                Err(EvalError::Parse)
            }
        }
        "CHOOSE" => {
            if args.len() < 2 {
                return Err(EvalError::Parse);
            }
            let index_value = eval_expr_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let option_index = trunc_f64_to_i64(value_as_arithmetic_number(&index_value)?)?;
            if option_index < 1 || option_index > (args.len() - 1) as i64 {
                return Err(EvalError::Parse);
            }
            eval_expr_value(
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
            let index_value = eval_expr_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let selected_index = trunc_f64_to_i64(value_as_arithmetic_number(&index_value)?)?;
            if selected_index < 1 || selected_index > (args.len() - 1) as i64 {
                return Err(EvalError::Parse);
            }
            eval_expr_value(
                workbook,
                sheet_name,
                &args[selected_index as usize],
                parsed_formulas,
                stack,
                cache,
            )
        }
        "LOWER" => {
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
            Ok(CellValue::Text(text.to_lowercase()))
        }
        "UPPER" => {
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
            Ok(CellValue::Text(text.to_uppercase()))
        }
        "TRIM" => {
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
            let trimmed = text.split_whitespace().collect::<Vec<_>>().join(" ");
            Ok(CellValue::Text(trimmed))
        }
        "LEFT" => {
            if args.is_empty() || args.len() > 2 {
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
            let count = if args.len() == 2 {
                parse_non_negative_usize(eval_expr(
                    workbook,
                    sheet_name,
                    &args[1],
                    parsed_formulas,
                    stack,
                    cache,
                )?)?
            } else {
                1
            };
            Ok(CellValue::Text(text.chars().take(count).collect()))
        }
        "RIGHT" => {
            if args.is_empty() || args.len() > 2 {
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
            let count = if args.len() == 2 {
                parse_non_negative_usize(eval_expr(
                    workbook,
                    sheet_name,
                    &args[1],
                    parsed_formulas,
                    stack,
                    cache,
                )?)?
            } else {
                1
            };
            let chars = text.chars().collect::<Vec<_>>();
            let start = chars.len().saturating_sub(count);
            Ok(CellValue::Text(chars[start..].iter().collect()))
        }
        "MID" => {
            if args.len() != 3 {
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
            let start_value = eval_expr(
                workbook,
                sheet_name,
                &args[1],
                parsed_formulas,
                stack,
                cache,
            )?;
            let count_value = eval_expr(
                workbook,
                sheet_name,
                &args[2],
                parsed_formulas,
                stack,
                cache,
            )?;
            let start = parse_text_start_position(start_value)?;
            let count = parse_non_negative_usize(count_value)?;
            let chars = text.chars().collect::<Vec<_>>();
            if start > chars.len() || count == 0 {
                return Ok(CellValue::Text(String::new()));
            }
            let start_index = start - 1;
            let end_index = start_index.saturating_add(count).min(chars.len());
            Ok(CellValue::Text(
                chars[start_index..end_index].iter().collect(),
            ))
        }
        "SUBSTITUTE" => {
            if args.len() < 3 || args.len() > 4 {
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
            let old_text = eval_expr_as_text(
                workbook,
                sheet_name,
                &args[1],
                parsed_formulas,
                stack,
                cache,
            )?;
            if old_text.is_empty() {
                return Err(EvalError::Parse);
            }
            let new_text = eval_expr_as_text(
                workbook,
                sheet_name,
                &args[2],
                parsed_formulas,
                stack,
                cache,
            )?;
            if args.len() == 4 {
                let occurrence = parse_positive_usize(eval_expr(
                    workbook,
                    sheet_name,
                    &args[3],
                    parsed_formulas,
                    stack,
                    cache,
                )?)?;
                Ok(CellValue::Text(substitute_nth_occurrence(
                    &text, &old_text, &new_text, occurrence,
                )))
            } else {
                Ok(CellValue::Text(text.replace(&old_text, &new_text)))
            }
        }
        "REPLACE" => {
            if args.len() != 4 {
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
            let start = parse_text_start_position(eval_expr(
                workbook,
                sheet_name,
                &args[1],
                parsed_formulas,
                stack,
                cache,
            )?)?;
            let count = parse_non_negative_usize(eval_expr(
                workbook,
                sheet_name,
                &args[2],
                parsed_formulas,
                stack,
                cache,
            )?)?;
            let replacement = eval_expr_as_text(
                workbook,
                sheet_name,
                &args[3],
                parsed_formulas,
                stack,
                cache,
            )?;
            let chars = text.chars().collect::<Vec<_>>();
            let start_index = start.saturating_sub(1).min(chars.len());
            let end_index = start_index.saturating_add(count).min(chars.len());
            let prefix = chars[..start_index].iter().collect::<String>();
            let suffix = chars[end_index..].iter().collect::<String>();
            Ok(CellValue::Text(format!("{prefix}{replacement}{suffix}")))
        }
        "CONCAT" => {
            if args.is_empty() {
                return Err(EvalError::Parse);
            }
            let mut joined = String::new();
            for arg in args {
                let text =
                    eval_expr_as_text(workbook, sheet_name, arg, parsed_formulas, stack, cache)?;
                joined.push_str(&text);
            }
            Ok(CellValue::Text(joined))
        }
        "TEXTJOIN" => {
            if args.len() < 3 {
                return Err(EvalError::Parse);
            }
            let delimiter = eval_expr_as_text(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let ignore_empty = eval_expr(
                workbook,
                sheet_name,
                &args[1],
                parsed_formulas,
                stack,
                cache,
            )? != 0.0;
            let mut parts = Vec::with_capacity(args.len().saturating_sub(2));
            for arg in &args[2..] {
                let value =
                    eval_expr_value(workbook, sheet_name, arg, parsed_formulas, stack, cache)?;
                let text = cell_value_as_text(&value)?;
                if ignore_empty && text.is_empty() {
                    continue;
                }
                parts.push(text);
            }
            Ok(CellValue::Text(parts.join(&delimiter)))
        }
        _ => Ok(CellValue::Number(eval_function(
            workbook,
            sheet_name,
            name,
            args,
            parsed_formulas,
            stack,
            cache,
        )?)),
    }
}

fn eval_function(
    workbook: &Workbook,
    sheet_name: &str,
    name: &str,
    args: &[Expr],
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut EvalCache,
) -> Result<f64, EvalError> {
    match name {
        "IF" => {
            if args.len() < 2 || args.len() > 3 {
                return Err(EvalError::Parse);
            }
            let condition_value = eval_expr_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            if value_as_condition(&condition_value)? {
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
        "IFS" => {
            if args.len() < 2 || args.len() % 2 != 0 {
                return Err(EvalError::Parse);
            }
            for pair in args.chunks(2) {
                let condition_value = eval_expr_value(
                    workbook,
                    sheet_name,
                    &pair[0],
                    parsed_formulas,
                    stack,
                    cache,
                )?;
                if value_as_condition(&condition_value)? {
                    return eval_expr(
                        workbook,
                        sheet_name,
                        &pair[1],
                        parsed_formulas,
                        stack,
                        cache,
                    );
                }
            }
            Err(EvalError::Parse)
        }
        "SWITCH" => {
            if args.len() < 3 {
                return Err(EvalError::Parse);
            }
            let expression_value = eval_expr_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let options = &args[1..];
            let has_default = options.len() % 2 == 1;
            let pair_end = if has_default {
                options.len() - 1
            } else {
                options.len()
            };
            let mut idx = 0usize;
            while idx < pair_end {
                let case_value = eval_expr_value(
                    workbook,
                    sheet_name,
                    &options[idx],
                    parsed_formulas,
                    stack,
                    cache,
                )?;
                if switch_values_equal(&expression_value, &case_value)? {
                    return eval_expr(
                        workbook,
                        sheet_name,
                        &options[idx + 1],
                        parsed_formulas,
                        stack,
                        cache,
                    );
                }
                idx += 2;
            }
            if has_default {
                eval_expr(
                    workbook,
                    sheet_name,
                    &options[pair_end],
                    parsed_formulas,
                    stack,
                    cache,
                )
            } else {
                Err(EvalError::Parse)
            }
        }
        "CHOOSE" => {
            if args.len() < 2 {
                return Err(EvalError::Parse);
            }
            let index_value = eval_expr_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let option_index = trunc_f64_to_i64(value_as_arithmetic_number(&index_value)?)?;
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
            let index_value = eval_expr_value(
                workbook,
                sheet_name,
                &args[0],
                parsed_formulas,
                stack,
                cache,
            )?;
            let selected_index = trunc_f64_to_i64(value_as_arithmetic_number(&index_value)?)?;
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
            match name {
                "AND" | "OR" | "XOR" => {
                    if args.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    let mut bool_values = Vec::<bool>::with_capacity(args.len());
                    for arg in args {
                        let value = eval_expr_value(
                            workbook,
                            sheet_name,
                            arg,
                            parsed_formulas,
                            stack,
                            cache,
                        )?;
                        bool_values.push(value_as_condition(&value)?);
                    }
                    return Ok(match name {
                        "AND" => {
                            if bool_values.iter().all(|value| *value) {
                                1.0
                            } else {
                                0.0
                            }
                        }
                        "OR" => {
                            if bool_values.iter().any(|value| *value) {
                                1.0
                            } else {
                                0.0
                            }
                        }
                        "XOR" => {
                            let true_count = bool_values.iter().filter(|value| **value).count();
                            if true_count % 2 == 1 {
                                1.0
                            } else {
                                0.0
                            }
                        }
                        _ => unreachable!(),
                    });
                }
                "NOT" => {
                    if args.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    let value = eval_expr_value(
                        workbook,
                        sheet_name,
                        &args[0],
                        parsed_formulas,
                        stack,
                        cache,
                    )?;
                    return Ok(if value_as_condition(&value)? {
                        0.0
                    } else {
                        1.0
                    });
                }
                _ => {}
            }

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
                "SUMSQ" => Ok(values.into_iter().map(|value| value * value).sum()),
                "PRODUCT" => {
                    if values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    Ok(values.into_iter().product())
                }
                "FACT" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    let n = parse_non_negative_i64(values[0])?;
                    let mut result = 1.0_f64;
                    for value in 2..=n {
                        result *= value as f64;
                        if !result.is_finite() {
                            return Err(EvalError::Parse);
                        }
                    }
                    Ok(result)
                }
                "FACTDOUBLE" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    let n = parse_non_negative_i64(values[0])?;
                    if n <= 1 {
                        return Ok(1.0);
                    }
                    let mut result = 1.0_f64;
                    let mut current = n;
                    while current > 1 {
                        result *= current as f64;
                        if !result.is_finite() {
                            return Err(EvalError::Parse);
                        }
                        current -= 2;
                    }
                    Ok(result)
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
                "COMBIN" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let n = parse_non_negative_i64(values[0])?;
                    let k = parse_non_negative_i64(values[1])?;
                    let result = combin_i128(n, k)?;
                    ensure_finite_result(result as f64)
                }
                "PERMUT" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let n = parse_non_negative_i64(values[0])?;
                    let k = parse_non_negative_i64(values[1])?;
                    if k > n {
                        return Err(EvalError::Parse);
                    }
                    let mut result = 1.0_f64;
                    for value in (n - k + 1)..=n {
                        result *= value as f64;
                        if !result.is_finite() {
                            return Err(EvalError::Parse);
                        }
                    }
                    Ok(result)
                }
                "GEOMEAN" => {
                    if values.is_empty() || values.iter().any(|value| *value <= 0.0) {
                        return Err(EvalError::Parse);
                    }
                    let log_sum = values.iter().map(|value| value.ln()).sum::<f64>();
                    ensure_finite_result((log_sum / values.len() as f64).exp())
                }
                "HARMEAN" => {
                    if values.is_empty() || values.iter().any(|value| *value <= 0.0) {
                        return Err(EvalError::Parse);
                    }
                    let reciprocal_sum = values.iter().map(|value| 1.0 / value).sum::<f64>();
                    if reciprocal_sum == 0.0 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values.len() as f64 / reciprocal_sum)
                }
                "VARP" => {
                    if values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    let mean = values.iter().sum::<f64>() / values.len() as f64;
                    let squared_dev_sum = values
                        .iter()
                        .map(|value| {
                            let delta = *value - mean;
                            delta * delta
                        })
                        .sum::<f64>();
                    ensure_finite_result(squared_dev_sum / values.len() as f64)
                }
                "VAR" | "VARS" => {
                    if values.len() < 2 {
                        return Err(EvalError::Parse);
                    }
                    let mean = values.iter().sum::<f64>() / values.len() as f64;
                    let squared_dev_sum = values
                        .iter()
                        .map(|value| {
                            let delta = *value - mean;
                            delta * delta
                        })
                        .sum::<f64>();
                    ensure_finite_result(squared_dev_sum / (values.len() as f64 - 1.0))
                }
                "STDEVP" => {
                    if values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    let mean = values.iter().sum::<f64>() / values.len() as f64;
                    let squared_dev_sum = values
                        .iter()
                        .map(|value| {
                            let delta = *value - mean;
                            delta * delta
                        })
                        .sum::<f64>();
                    ensure_finite_result((squared_dev_sum / values.len() as f64).sqrt())
                }
                "STDEV" | "STDEVS" => {
                    if values.len() < 2 {
                        return Err(EvalError::Parse);
                    }
                    let mean = values.iter().sum::<f64>() / values.len() as f64;
                    let squared_dev_sum = values
                        .iter()
                        .map(|value| {
                            let delta = *value - mean;
                            delta * delta
                        })
                        .sum::<f64>();
                    ensure_finite_result((squared_dev_sum / (values.len() as f64 - 1.0)).sqrt())
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
                "PI" => {
                    if !values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    Ok(std::f64::consts::PI)
                }
                "EXP" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].exp())
                }
                "LN" => {
                    if values.len() != 1 || values[0] <= 0.0 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].ln())
                }
                "LOG10" => {
                    if values.len() != 1 || values[0] <= 0.0 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].log10())
                }
                "LOG" => {
                    if values.is_empty() || values.len() > 2 {
                        return Err(EvalError::Parse);
                    }
                    let number = values[0];
                    if number <= 0.0 {
                        return Err(EvalError::Parse);
                    }
                    let base = if values.len() == 2 { values[1] } else { 10.0 };
                    if base <= 0.0 || base == 1.0 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(number.log(base))
                }
                "SIN" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].sin())
                }
                "COS" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].cos())
                }
                "TAN" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].tan())
                }
                "SINH" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].sinh())
                }
                "COSH" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].cosh())
                }
                "TANH" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].tanh())
                }
                "ASINH" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].asinh())
                }
                "ACOSH" => {
                    if values.len() != 1 || values[0] < 1.0 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].acosh())
                }
                "ATANH" => {
                    if values.len() != 1 || values[0] <= -1.0 || values[0] >= 1.0 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].atanh())
                }
                "ASIN" => {
                    if values.len() != 1 || values[0] < -1.0 || values[0] > 1.0 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].asin())
                }
                "ACOS" => {
                    if values.len() != 1 || values[0] < -1.0 || values[0] > 1.0 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].acos())
                }
                "ATAN" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].atan())
                }
                "ATAN2" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0].atan2(values[1]))
                }
                "RADIANS" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0] * std::f64::consts::PI / 180.0)
                }
                "DEGREES" => {
                    if values.len() != 1 {
                        return Err(EvalError::Parse);
                    }
                    ensure_finite_result(values[0] * 180.0 / std::f64::consts::PI)
                }
                "PMT" => {
                    if values.len() < 3 || values.len() > 5 {
                        return Err(EvalError::Parse);
                    }
                    let rate = values[0];
                    let nper = values[1];
                    let pv = values[2];
                    let fv = if values.len() >= 4 { values[3] } else { 0.0 };
                    let payment_type = parse_payment_type(values.get(4).copied())?;
                    if !rate.is_finite() || !nper.is_finite() || !pv.is_finite() || !fv.is_finite()
                    {
                        return Err(EvalError::Parse);
                    }
                    if nper == 0.0 {
                        return Err(EvalError::Parse);
                    }
                    if rate == 0.0 {
                        ensure_finite_result(-(pv + fv) / nper)
                    } else {
                        let one_plus_rate = 1.0 + rate;
                        let growth = one_plus_rate.powf(nper);
                        if !growth.is_finite() || growth == 1.0 {
                            return Err(EvalError::Parse);
                        }
                        let numerator = rate * (fv + pv * growth);
                        let denominator = (1.0 + rate * payment_type) * (growth - 1.0);
                        if denominator == 0.0 {
                            return Err(EvalError::Parse);
                        }
                        ensure_finite_result(-numerator / denominator)
                    }
                }
                "PV" => {
                    if values.len() < 3 || values.len() > 5 {
                        return Err(EvalError::Parse);
                    }
                    let rate = values[0];
                    let nper = values[1];
                    let pmt = values[2];
                    let fv = if values.len() >= 4 { values[3] } else { 0.0 };
                    let payment_type = parse_payment_type(values.get(4).copied())?;
                    if !rate.is_finite() || !nper.is_finite() || !pmt.is_finite() || !fv.is_finite()
                    {
                        return Err(EvalError::Parse);
                    }
                    if nper == 0.0 {
                        return Err(EvalError::Parse);
                    }
                    if rate == 0.0 {
                        ensure_finite_result(-(fv + pmt * nper))
                    } else {
                        let one_plus_rate = 1.0 + rate;
                        let growth = one_plus_rate.powf(nper);
                        if !growth.is_finite() || growth == 0.0 {
                            return Err(EvalError::Parse);
                        }
                        let annuity_term =
                            pmt * (1.0 + rate * payment_type) * (growth - 1.0) / rate;
                        ensure_finite_result(-(fv + annuity_term) / growth)
                    }
                }
                "FV" => {
                    if values.len() < 3 || values.len() > 5 {
                        return Err(EvalError::Parse);
                    }
                    let rate = values[0];
                    let nper = values[1];
                    let pmt = values[2];
                    let pv = if values.len() >= 4 { values[3] } else { 0.0 };
                    let payment_type = parse_payment_type(values.get(4).copied())?;
                    if !rate.is_finite() || !nper.is_finite() || !pmt.is_finite() || !pv.is_finite()
                    {
                        return Err(EvalError::Parse);
                    }
                    if nper == 0.0 {
                        return Err(EvalError::Parse);
                    }
                    if rate == 0.0 {
                        ensure_finite_result(-(pv + pmt * nper))
                    } else {
                        let one_plus_rate = 1.0 + rate;
                        let growth = one_plus_rate.powf(nper);
                        if !growth.is_finite() {
                            return Err(EvalError::Parse);
                        }
                        let annuity_term =
                            pmt * (1.0 + rate * payment_type) * (growth - 1.0) / rate;
                        ensure_finite_result(-(pv * growth + annuity_term))
                    }
                }
                "NPV" => {
                    if values.len() < 2 {
                        return Err(EvalError::Parse);
                    }
                    let rate = values[0];
                    if !rate.is_finite() || rate == -1.0 {
                        return Err(EvalError::Parse);
                    }
                    let mut total = 0.0;
                    for (index, value) in values.iter().skip(1).enumerate() {
                        let period = (index + 1) as f64;
                        let discount = (1.0 + rate).powf(period);
                        if !discount.is_finite() || discount == 0.0 {
                            return Err(EvalError::Parse);
                        }
                        total += *value / discount;
                    }
                    ensure_finite_result(total)
                }
                "BITAND" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let lhs = parse_bit_operand(values[0])?;
                    let rhs = parse_bit_operand(values[1])?;
                    Ok((lhs & rhs) as f64)
                }
                "BITOR" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let lhs = parse_bit_operand(values[0])?;
                    let rhs = parse_bit_operand(values[1])?;
                    Ok((lhs | rhs) as f64)
                }
                "BITXOR" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let lhs = parse_bit_operand(values[0])?;
                    let rhs = parse_bit_operand(values[1])?;
                    Ok((lhs ^ rhs) as f64)
                }
                "BITLSHIFT" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let operand = parse_bit_operand(values[0])?;
                    let shift = parse_shift_amount(values[1])?;
                    let result = apply_shift(operand, shift, true)?;
                    Ok(result as f64)
                }
                "BITRSHIFT" => {
                    if values.len() != 2 {
                        return Err(EvalError::Parse);
                    }
                    let operand = parse_bit_operand(values[0])?;
                    let shift = parse_shift_amount(values[1])?;
                    let result = apply_shift(operand, shift, false)?;
                    Ok(result as f64)
                }
                "GCD" => {
                    if values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    let mut result = 0i64;
                    for value in values {
                        let parsed = parse_non_negative_i64(value)?;
                        result = gcd_i64(result, parsed);
                    }
                    Ok(result as f64)
                }
                "LCM" => {
                    if values.is_empty() {
                        return Err(EvalError::Parse);
                    }
                    let mut result = 1i64;
                    let mut saw_zero = false;
                    for value in values {
                        let parsed = parse_non_negative_i64(value)?;
                        if parsed == 0 {
                            saw_zero = true;
                            break;
                        }
                        result = lcm_i64(result, parsed)?;
                    }
                    if saw_zero {
                        Ok(0.0)
                    } else {
                        Ok(result as f64)
                    }
                }
                "MATCH" => {
                    if args.len() < 2 {
                        return Err(EvalError::Parse);
                    }
                    let mut scalar_args = Vec::<CellValue>::with_capacity(args.len());
                    for arg in args {
                        scalar_args.push(eval_expr_value(
                            workbook,
                            sheet_name,
                            arg,
                            parsed_formulas,
                            stack,
                            cache,
                        )?);
                    }

                    let needle = value_as_arithmetic_number(&scalar_args[0])?;
                    let mut candidate_values = &scalar_args[1..];
                    let mut mode = 0i64;
                    if scalar_args.len() >= 4 {
                        let mode_candidate =
                            value_as_arithmetic_number(&scalar_args[scalar_args.len() - 1]);
                        if let Ok(parsed_mode) = mode_candidate.and_then(trunc_f64_to_i64) {
                            if parsed_mode == -1 || parsed_mode == 0 || parsed_mode == 1 {
                                mode = parsed_mode;
                                candidate_values = &scalar_args[1..scalar_args.len() - 1];
                            }
                        }
                    }
                    let candidates = candidate_values
                        .iter()
                        .map(value_as_arithmetic_number)
                        .collect::<Result<Vec<_>, _>>()?;
                    let index = find_match_index_match(needle, &candidates, mode)?;
                    Ok(index as f64)
                }
                "XMATCH" => {
                    if args.len() < 2 {
                        return Err(EvalError::Parse);
                    }
                    let mut scalar_args = Vec::<CellValue>::with_capacity(args.len());
                    for arg in args {
                        scalar_args.push(eval_expr_value(
                            workbook,
                            sheet_name,
                            arg,
                            parsed_formulas,
                            stack,
                            cache,
                        )?);
                    }

                    let needle = value_as_arithmetic_number(&scalar_args[0])?;
                    let mut candidate_values = &scalar_args[1..];
                    let mut mode = 0i64;
                    if scalar_args.len() >= 4 {
                        let mode_candidate =
                            value_as_arithmetic_number(&scalar_args[scalar_args.len() - 1]);
                        if let Ok(parsed_mode) = mode_candidate.and_then(trunc_f64_to_i64) {
                            if parsed_mode == -1 || parsed_mode == 0 || parsed_mode == 1 {
                                mode = parsed_mode;
                                candidate_values = &scalar_args[1..scalar_args.len() - 1];
                            }
                        }
                    }
                    let candidates = candidate_values
                        .iter()
                        .map(value_as_arithmetic_number)
                        .collect::<Result<Vec<_>, _>>()?;
                    let index = find_match_index_xmatch(needle, &candidates, mode)?;
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

fn eval_status(result: &Result<CellValue, EvalError>) -> &'static str {
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
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) => {}
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
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Cell(_) => 0,
        Expr::Function { args, .. } => 1 + args.iter().map(count_expr_functions).sum::<usize>(),
        Expr::UnaryMinus(inner) => count_expr_functions(inner),
        Expr::Binary { left, right, .. } => {
            count_expr_functions(left) + count_expr_functions(right)
        }
    }
}

fn count_expr_nodes(expr: &Expr) -> usize {
    match expr {
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Cell(_) => 1,
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
            Some('"') => self.parse_string_literal(),
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

        let upper_token = token.to_ascii_uppercase();
        if upper_token == "TRUE" {
            return Ok(Expr::Bool(true));
        }
        if upper_token == "FALSE" {
            return Ok(Expr::Bool(false));
        }

        parse_a1_cell(&token)
            .map(Expr::Cell)
            .ok_or(EvalError::Parse)
    }

    fn parse_string_literal(&mut self) -> Result<Expr, EvalError> {
        if self.bump_char() != Some('"') {
            return Err(EvalError::Parse);
        }

        let mut value = String::new();
        loop {
            match self.bump_char() {
                Some('"') => {
                    if self.peek_char() == Some('"') {
                        self.bump_char();
                        value.push('"');
                    } else {
                        break;
                    }
                }
                Some(ch) => value.push(ch),
                None => return Err(EvalError::Parse),
            }
        }
        Ok(Expr::Text(value))
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

fn value_as_arithmetic_number(value: &CellValue) -> Result<f64, EvalError> {
    match value {
        CellValue::Number(n) => Ok(*n),
        CellValue::Bool(true) => Ok(1.0),
        CellValue::Bool(false) => Ok(0.0),
        CellValue::Empty => Ok(0.0),
        CellValue::Text(text) => {
            let trimmed = text.trim();
            if trimmed.is_empty() {
                Ok(0.0)
            } else {
                parse_text_number(trimmed)
            }
        }
        CellValue::Error(_) => Err(EvalError::Parse),
    }
}

fn eval_expr_as_text(
    workbook: &Workbook,
    sheet_name: &str,
    expr: &Expr,
    parsed_formulas: &BTreeMap<CellRef, Result<Expr, EvalError>>,
    stack: &mut BTreeSet<CellRef>,
    cache: &mut EvalCache,
) -> Result<String, EvalError> {
    let value = eval_expr_value(workbook, sheet_name, expr, parsed_formulas, stack, cache)?;
    cell_value_as_text(&value)
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

fn ensure_finite_result(value: f64) -> Result<f64, EvalError> {
    if value.is_finite() {
        Ok(value)
    } else {
        Err(EvalError::Parse)
    }
}

fn parse_payment_type(value: Option<f64>) -> Result<f64, EvalError> {
    let raw = value.unwrap_or(0.0);
    let parsed = trunc_f64_to_i64(raw)?;
    match parsed {
        0 => Ok(0.0),
        1 => Ok(1.0),
        _ => Err(EvalError::Parse),
    }
}

fn parse_non_negative_i64(value: f64) -> Result<i64, EvalError> {
    let parsed = trunc_f64_to_i64(value)?;
    if parsed < 0 {
        return Err(EvalError::Parse);
    }
    Ok(parsed)
}

fn parse_non_negative_usize(value: f64) -> Result<usize, EvalError> {
    let parsed = parse_non_negative_i64(value)?;
    usize::try_from(parsed).map_err(|_| EvalError::Parse)
}

fn parse_positive_usize(value: f64) -> Result<usize, EvalError> {
    let parsed = trunc_f64_to_i64(value)?;
    if parsed < 1 {
        return Err(EvalError::Parse);
    }
    usize::try_from(parsed).map_err(|_| EvalError::Parse)
}

const MAX_BIT_OPERAND: u64 = (1u64 << 48) - 1;

fn parse_bit_operand(value: f64) -> Result<u64, EvalError> {
    let parsed = parse_non_negative_i64(value)? as u64;
    if parsed > MAX_BIT_OPERAND {
        return Err(EvalError::Parse);
    }
    Ok(parsed)
}

fn parse_shift_amount(value: f64) -> Result<i64, EvalError> {
    let parsed = trunc_f64_to_i64(value)?;
    if !(-63..=63).contains(&parsed) {
        return Err(EvalError::Parse);
    }
    Ok(parsed)
}

fn apply_shift(operand: u64, shift: i64, left_if_positive: bool) -> Result<u64, EvalError> {
    let mut effective = shift;
    if !left_if_positive {
        effective = -effective;
    }
    let shifted = if effective >= 0 {
        operand
            .checked_shl(effective as u32)
            .ok_or(EvalError::Parse)?
    } else {
        operand
            .checked_shr((-effective) as u32)
            .ok_or(EvalError::Parse)?
    };
    if shifted > MAX_BIT_OPERAND {
        return Err(EvalError::Parse);
    }
    Ok(shifted)
}

fn combin_i128(n: i64, k: i64) -> Result<i128, EvalError> {
    if n < 0 || k < 0 || k > n {
        return Err(EvalError::Parse);
    }
    let k = k.min(n - k);
    let mut result = 1i128;
    for i in 1..=k {
        let factor = (n - k + i) as i128;
        result = result.checked_mul(factor).ok_or(EvalError::Parse)?;
        result /= i as i128;
    }
    Ok(result)
}

fn gcd_i64(lhs: i64, rhs: i64) -> i64 {
    let mut a = lhs.abs();
    let mut b = rhs.abs();
    while b != 0 {
        let r = a % b;
        a = b;
        b = r;
    }
    a
}

fn lcm_i64(lhs: i64, rhs: i64) -> Result<i64, EvalError> {
    if lhs == 0 || rhs == 0 {
        return Ok(0);
    }
    let gcd = gcd_i64(lhs, rhs);
    if gcd == 0 {
        return Err(EvalError::Parse);
    }
    let scaled = lhs / gcd;
    let product = (scaled as i128)
        .checked_mul(rhs as i128)
        .ok_or(EvalError::Parse)?;
    if product < i64::MIN as i128 || product > i64::MAX as i128 {
        return Err(EvalError::Parse);
    }
    Ok(product as i64)
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
    if start < 1 {
        return Err(EvalError::Parse);
    }
    usize::try_from(start).map_err(|_| EvalError::Parse)
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

fn substitute_nth_occurrence(
    text: &str,
    old_text: &str,
    new_text: &str,
    occurrence: usize,
) -> String {
    let mut seen = 0usize;
    let mut replaced: Option<(usize, usize)> = None;
    for (index, matched) in text.match_indices(old_text) {
        seen = seen.saturating_add(1);
        if seen == occurrence {
            replaced = Some((index, matched.len()));
            break;
        }
    }

    if let Some((match_index, match_len)) = replaced {
        let mut result = String::new();
        result.push_str(&text[..match_index]);
        result.push_str(new_text);
        result.push_str(&text[match_index + match_len..]);
        result
    } else {
        text.to_string()
    }
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
    cache: &mut EvalCache,
) -> CellValue {
    match eval_expr_value(workbook, sheet_name, expr, parsed_formulas, stack, cache) {
        Ok(value) => value,
        Err(_) => CellValue::Error("#ERROR".to_string()),
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
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Text("  Root   Cellar  ".to_string()),
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
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=LOWER(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 16,
            formula: "=UPPER(C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 17,
            formula: "=TRIM(D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 18,
            formula: "=LEFT(B1,4)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 19,
            formula: "=RIGHT(B1,6)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 20,
            formula: "=MID(B1,5,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 21,
            formula: "=LEN(LOWER(B1))".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 22,
            formula: "=LEFT(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 23,
            formula: "=RIGHT(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 24,
            formula: "=MID(B1,20,3)".to_string(),
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
        let w2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 23 }))
            .expect("w2")
            .value
            .clone();
        let x2 = wb
            .sheets
            .get("Sheet1")
            .and_then(|s| s.cells.get(&CellRef { row: 2, col: 24 }))
            .expect("x2")
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
        assert_eq!(o2, CellValue::Text("rootcellar".to_string()));
        assert_eq!(p2, CellValue::Text("ROOT".to_string()));
        assert_eq!(q2, CellValue::Text("Root Cellar".to_string()));
        assert_eq!(r2, CellValue::Text("Root".to_string()));
        assert_eq!(s2, CellValue::Text("Cellar".to_string()));
        assert_eq!(t2, CellValue::Text("Cel".to_string()));
        assert_eq!(u2, CellValue::Number(10.0));
        assert_eq!(v2, CellValue::Text("R".to_string()));
        assert_eq!(w2, CellValue::Text("r".to_string()));
        assert_eq!(x2, CellValue::Text(String::new()));
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
    fn text_transform_invalid_inputs_yield_parse_error() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Text("RootCellar".to_string()),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=LEFT(A1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=RIGHT(A1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=MID(A1,0,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=MID(A1,1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=LOWER()".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=TRIM(A1,A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 6);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(2), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(3), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(4), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(5), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(6), CellValue::Error("#PARSE!".to_string()));
    }

    #[test]
    fn evaluates_text_composition_extension_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Text("Root Cellar Root".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Text("Root".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Text("Barrel".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Text("|".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 6,
            value: CellValue::Number(42.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 7,
            value: CellValue::Bool(true),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 8,
            value: CellValue::Text(String::new()),
        });

        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=SUBSTITUTE(A1,B1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=SUBSTITUTE(A1,B1,C1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=SUBSTITUTE(A1,B1,C1,5)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=REPLACE(A1,6,6,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=REPLACE(A1,20,3,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=REPLACE(B1,1,0,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=CONCAT(B1,D1,C1,D1,F1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=TEXTJOIN(D1,1,B1,H1,C1,E1,F1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=TEXTJOIN(D1,0,B1,H1,C1,E1,F1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=TEXTJOIN(D1,1,B1,G1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=LEN(CONCAT(B1,C1))".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=TEXTJOIN(D1,1,H1,E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(
            value_at(1),
            CellValue::Text("Barrel Cellar Barrel".to_string())
        );
        assert_eq!(
            value_at(2),
            CellValue::Text("Root Cellar Barrel".to_string())
        );
        assert_eq!(value_at(3), CellValue::Text("Root Cellar Root".to_string()));
        assert_eq!(value_at(4), CellValue::Text("Root Barrel Root".to_string()));
        assert_eq!(
            value_at(5),
            CellValue::Text("Root Cellar RootBarrel".to_string())
        );
        assert_eq!(value_at(6), CellValue::Text("BarrelRoot".to_string()));
        assert_eq!(value_at(7), CellValue::Text("Root|Barrel|42".to_string()));
        assert_eq!(value_at(8), CellValue::Text("Root|Barrel|42".to_string()));
        assert_eq!(value_at(9), CellValue::Text("Root||Barrel||42".to_string()));
        assert_eq!(
            value_at(10),
            CellValue::Text("Root|TRUE|Barrel".to_string())
        );
        assert_eq!(value_at(11), CellValue::Number(10.0));
        assert_eq!(value_at(12), CellValue::Text(String::new()));
    }

    #[test]
    fn text_composition_invalid_inputs_yield_parse_error() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Text("Root Cellar".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Text("Root".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Text("Barrel".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Text("|".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            value: CellValue::Text(String::new()),
        });

        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=SUBSTITUTE(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=SUBSTITUTE(A1,B1,C1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=SUBSTITUTE(A1,B1,C1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=SUBSTITUTE(A1,E1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=SUBSTITUTE(A1,B1,C1,1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=REPLACE(A1,0,1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=REPLACE(A1,1,-1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=REPLACE(A1,1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=CONCAT()".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=TEXTJOIN(D1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 10);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(2), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(3), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(4), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(5), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(6), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(7), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(8), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(9), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(10), CellValue::Error("#PARSE!".to_string()));
    }

    #[test]
    fn evaluates_formula_text_and_boolean_literals() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            value: CellValue::Error("#N/A".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            value: CellValue::Text("barrel".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            value: CellValue::Bool(false),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            formula: "=\"Root\"".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            formula: "=CONCAT(\"Root\",\"Cellar\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=CONCAT(\"He said \"\"Hi\"\"\",\"!\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            formula: "=IF(TRUE,10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            formula: "=IF(FALSE,10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 6,
            formula: "=N(TRUE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 7,
            formula: "=N(FALSE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 8,
            formula: "=TEXTJOIN(\"|\",TRUE,\"A\",\"\",\"B\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 9,
            formula: "=TEXTJOIN(\"|\",FALSE,\"A\",\"\",\"B\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 10,
            formula: "=SUBSTITUTE(\"Root Root\",\"Root\",\"Cellar\",2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 11,
            formula: "=TRUE+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 12,
            formula: "=LOWER(\"RoOt\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 13,
            formula: "=IF(TRUE,\"Yes\",\"No\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 14,
            formula: "=IF(FALSE,TRUE,FALSE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 15,
            formula: "=IFS(FALSE,\"A\",TRUE,\"B\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 16,
            formula: "=SWITCH(2,1,\"one\",2,\"two\",\"other\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 17,
            formula: "=SWITCH(3,1,\"one\",2,\"two\",\"other\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 18,
            formula: "=CHOOSE(2,10,\"pick\",TRUE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 19,
            formula: "=INDEX(3,10,\"pick\",TRUE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 20,
            formula: "=IFERROR(1/0,\"fallback\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 21,
            formula: "=IFERROR(\"ok\",\"fallback\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 22,
            formula: "=IFERROR(A2,\"fallback-from-cell-error\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 23,
            formula: "=IFERROR(IF(TRUE,A2,0),TRUE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 24,
            formula: "=IFERROR(IF(FALSE,A2,\"ok\"),\"fallback\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 25,
            formula: "=SWITCH(\"beta\",\"alpha\",\"A\",\"beta\",\"B\",\"Z\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 26,
            formula: "=SWITCH(TRUE,FALSE,\"no\",TRUE,\"yes\",\"default\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 27,
            formula: "=SWITCH(FALSE,\"text-case\",1,TRUE,2,\"none\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 28,
            formula: "=SWITCH(B2,\"root\",\"R\",\"barrel\",\"B\",\"X\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 29,
            formula: "=SWITCH(C2,TRUE,\"T\",FALSE,\"F\",\"D\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 0);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 1, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Text("Root".to_string()));
        assert_eq!(value_at(2), CellValue::Text("RootCellar".to_string()));
        assert_eq!(value_at(3), CellValue::Text("He said \"Hi\"!".to_string()));
        assert_eq!(value_at(4), CellValue::Number(10.0));
        assert_eq!(value_at(5), CellValue::Number(20.0));
        assert_eq!(value_at(6), CellValue::Number(1.0));
        assert_eq!(value_at(7), CellValue::Number(0.0));
        assert_eq!(value_at(8), CellValue::Text("A|B".to_string()));
        assert_eq!(value_at(9), CellValue::Text("A||B".to_string()));
        assert_eq!(value_at(10), CellValue::Text("Root Cellar".to_string()));
        assert_eq!(value_at(11), CellValue::Number(2.0));
        assert_eq!(value_at(12), CellValue::Text("root".to_string()));
        assert_eq!(value_at(13), CellValue::Text("Yes".to_string()));
        assert_eq!(value_at(14), CellValue::Bool(false));
        assert_eq!(value_at(15), CellValue::Text("B".to_string()));
        assert_eq!(value_at(16), CellValue::Text("two".to_string()));
        assert_eq!(value_at(17), CellValue::Text("other".to_string()));
        assert_eq!(value_at(18), CellValue::Text("pick".to_string()));
        assert_eq!(value_at(19), CellValue::Bool(true));
        assert_eq!(value_at(20), CellValue::Text("fallback".to_string()));
        assert_eq!(value_at(21), CellValue::Text("ok".to_string()));
        assert_eq!(
            value_at(22),
            CellValue::Text("fallback-from-cell-error".to_string())
        );
        assert_eq!(value_at(23), CellValue::Bool(true));
        assert_eq!(value_at(24), CellValue::Text("ok".to_string()));
        assert_eq!(value_at(25), CellValue::Text("B".to_string()));
        assert_eq!(value_at(26), CellValue::Text("yes".to_string()));
        assert_eq!(value_at(27), CellValue::Text("none".to_string()));
        assert_eq!(value_at(28), CellValue::Text("B".to_string()));
        assert_eq!(value_at(29), CellValue::Text("F".to_string()));
    }

    #[test]
    fn literal_formula_invalid_inputs_yield_parse_error() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            formula: "=CONCAT(\"unterminated)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            formula: "=\"still unterminated".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=IF(TRUE(),1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            formula: "=TRUEFALSE".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            formula: "=TEXTJOIN(\"|\",TRUE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 5);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 1, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(2), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(3), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(4), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(5), CellValue::Error("#PARSE!".to_string()));
    }

    #[test]
    fn coerces_text_conditions_in_if_and_ifs() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            value: CellValue::Text("TRUE".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            value: CellValue::Text("0".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            value: CellValue::Text("foo".to_string()),
        });

        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            formula: "=IF(\"TRUE\",10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            formula: "=IF(\"FALSE\",10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            formula: "=IF(\"2\",10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            formula: "=IF(\"0\",10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            formula: "=IF(\"\",10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 6,
            formula: "=IFS(\"FALSE\",\"A\",\"TRUE\",\"B\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 7,
            formula: "=IF(\"foo\",10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 8,
            formula: "=IFS(\"foo\",\"A\",TRUE,\"B\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 9,
            formula: "=IF(A2,1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 10,
            formula: "=IF(B2,1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 11,
            formula: "=IF(C2,1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 12,
            formula: "=IFS(B2,\"zero\",A2,\"truthy\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 3);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 1, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Number(10.0));
        assert_eq!(value_at(2), CellValue::Number(20.0));
        assert_eq!(value_at(3), CellValue::Number(10.0));
        assert_eq!(value_at(4), CellValue::Number(20.0));
        assert_eq!(value_at(5), CellValue::Number(20.0));
        assert_eq!(value_at(6), CellValue::Text("B".to_string()));
        assert_eq!(value_at(7), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(8), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(9), CellValue::Number(1.0));
        assert_eq!(value_at(10), CellValue::Number(0.0));
        assert_eq!(value_at(11), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(12), CellValue::Text("truthy".to_string()));
    }

    #[test]
    fn coerces_text_conditions_in_logical_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Text("TRUE".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Text("FALSE".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Text("2".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Text("0".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            value: CellValue::Text("foo".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 6,
            value: CellValue::Bool(true),
        });

        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=AND(A1,C1,F1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=OR(B1,D1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=XOR(A1,B1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=NOT(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=NOT(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=AND(A1,E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=OR(E1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=XOR(A1,E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=NOT(E1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=AND(\"\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=OR(\"\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=NOT(\"\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=AND(\" 2 \",\"TRUE\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=XOR(\"TRUE\",\"TRUE\",\"FALSE\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=OR(\"FALSE\",\" 0 \",FALSE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 16,
            formula: "=OR(\"FALSE\",\" 1 \",FALSE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 17,
            formula: "=AND(\"FALSE\",TRUE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 18,
            formula: "=XOR(\"FALSE\",\"1\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 4);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Number(1.0));
        assert_eq!(value_at(2), CellValue::Number(0.0));
        assert_eq!(value_at(3), CellValue::Number(0.0));
        assert_eq!(value_at(4), CellValue::Number(1.0));
        assert_eq!(value_at(5), CellValue::Number(0.0));
        assert_eq!(value_at(6), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(7), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(8), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(9), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(10), CellValue::Number(0.0));
        assert_eq!(value_at(11), CellValue::Number(0.0));
        assert_eq!(value_at(12), CellValue::Number(1.0));
        assert_eq!(value_at(13), CellValue::Number(1.0));
        assert_eq!(value_at(14), CellValue::Number(0.0));
        assert_eq!(value_at(15), CellValue::Number(0.0));
        assert_eq!(value_at(16), CellValue::Number(1.0));
        assert_eq!(value_at(17), CellValue::Number(0.0));
        assert_eq!(value_at(18), CellValue::Number(1.0));
    }

    #[test]
    fn coerces_numeric_text_in_arithmetic_and_rejects_invalid_text() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Text("2".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Text("foo".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Text(String::new()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Bool(true),
        });

        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=A1+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=B1+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=\"2\"+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=\"foo\"+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=-A1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=-B1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=C1+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=D1+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 3);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Number(3.0));
        assert_eq!(value_at(2), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(3), CellValue::Number(3.0));
        assert_eq!(value_at(4), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(5), CellValue::Number(-2.0));
        assert_eq!(value_at(6), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(7), CellValue::Number(1.0));
        assert_eq!(value_at(8), CellValue::Number(2.0));
    }

    #[test]
    fn coerces_text_indexes_in_choose_and_index() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Text("2".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Text("foo".to_string()),
        });

        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=CHOOSE(\"2\",10,\"pick\",TRUE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=INDEX(\"3\",10,\"pick\",TRUE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=CHOOSE(A1,10,\"pick\",TRUE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=INDEX(A1,10,\"pick\",TRUE)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=CHOOSE(\"2.9\",10,20,30)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=CHOOSE(\"foo\",10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=INDEX(\"foo\",10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=CHOOSE(B1,10,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 3);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Text("pick".to_string()));
        assert_eq!(value_at(2), CellValue::Bool(true));
        assert_eq!(value_at(3), CellValue::Text("pick".to_string()));
        assert_eq!(value_at(4), CellValue::Text("pick".to_string()));
        assert_eq!(value_at(5), CellValue::Number(20.0));
        assert_eq!(value_at(6), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(7), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(8), CellValue::Error("#PARSE!".to_string()));
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
    fn evaluates_xor_ifs_and_switch_extension_functions() {
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
            value: CellValue::Number(0.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Number(2.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=XOR(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=XOR(A1,B1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=XOR(B1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=XOR()".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=IFS(A1,10,B1,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=IFS(B1,10,C1,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=IFS(B1,10,0,20)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=IFS(A1,10,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=SWITCH(A1,1,100,2,200,300)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=SWITCH(C1,1,100,2,200,300)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=SWITCH(B1,1,100,2,200,300)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=SWITCH(B1,1,100,2,200)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=SWITCH(A1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=SWITCH(A1,1,100,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=IFS(A1,1,1/0,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 16,
            formula: "=SWITCH(A1,1,1,1/0,2,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 5);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Number(1.0));
        assert_eq!(value_at(2), CellValue::Number(0.0));
        assert_eq!(value_at(3), CellValue::Number(0.0));
        assert_eq!(value_at(4), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(5), CellValue::Number(10.0));
        assert_eq!(value_at(6), CellValue::Number(20.0));
        assert_eq!(value_at(7), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(8), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(9), CellValue::Number(100.0));
        assert_eq!(value_at(10), CellValue::Number(200.0));
        assert_eq!(value_at(11), CellValue::Number(300.0));
        assert_eq!(value_at(12), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(13), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(14), CellValue::Number(100.0));
        assert_eq!(value_at(15), CellValue::Number(1.0));
        assert_eq!(value_at(16), CellValue::Number(1.0));
    }

    #[test]
    fn evaluates_statistical_aggregate_functions() {
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
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Number(4.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Number(4.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            value: CellValue::Number(5.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 6,
            value: CellValue::Number(5.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 7,
            value: CellValue::Number(7.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 8,
            value: CellValue::Number(9.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=SUMSQ(A1,B1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=GEOMEAN(A1,8)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=HARMEAN(A1,8)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=VARP(A1,B1,C1,D1,E1,F1,G1,H1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=STDEVP(A1,B1,C1,D1,E1,F1,G1,H1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=VAR(A1,B1,C1,D1,E1,F1,G1,H1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=STDEV(A1,B1,C1,D1,E1,F1,G1,H1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=VARS(A1,B1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=STDEVS(A1,B1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=VARP(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=STDEVP(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=GEOMEAN(-1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=HARMEAN(0,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=STDEV(5)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=VAR(5)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 16,
            formula: "=GEOMEAN()".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 17,
            formula: "=SUMSQ()".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 5);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };
        let assert_close = |value: CellValue, expected: f64, label: &str| match value {
            CellValue::Number(actual) => {
                assert!(
                    (actual - expected).abs() < 1e-12,
                    "{label} expected {expected}, got {actual}"
                );
            }
            other => panic!("{label} expected number, got {other:?}"),
        };

        assert_eq!(value_at(1), CellValue::Number(36.0));
        assert_eq!(value_at(2), CellValue::Number(4.0));
        assert_eq!(value_at(3), CellValue::Number(3.2));
        assert_eq!(value_at(4), CellValue::Number(4.0));
        assert_eq!(value_at(5), CellValue::Number(2.0));
        assert_close(value_at(6), 32.0 / 7.0, "F2");
        assert_close(value_at(7), (32.0_f64 / 7.0_f64).sqrt(), "G2");
        assert_close(value_at(8), 1.3333333333333333, "H2");
        assert_close(value_at(9), 1.1547005383792515, "I2");
        assert_eq!(value_at(10), CellValue::Number(0.0));
        assert_eq!(value_at(11), CellValue::Number(0.0));
        assert_eq!(value_at(12), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(13), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(14), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(15), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(16), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(17), CellValue::Number(0.0));
    }

    #[test]
    fn evaluates_financial_extension_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(0.05),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(10.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Number(1000.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=PMT(A1,B1,C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=FV(A1,B1,A2,C1,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=PV(A1,B1,A2,0,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=NPV(0.1,100,100)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=PMT(0,10,1000)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=PV(0,10,-100)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=FV(0,10,-100,0,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=PMT(0.05,10,1000,0,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=NPV(0,10,20,30)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=PMT(0.05,10,1000,0,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=NPV(-1,100)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=PMT(0,0,1000)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=PV(0.05,10,-100,0,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=FV(0.05,10,-100,0,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=NPV(0.1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 6);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };
        let assert_close = |value: CellValue, expected: f64, label: &str| match value {
            CellValue::Number(actual) => {
                assert!(
                    (actual - expected).abs() < 1e-9,
                    "{label} expected {expected}, got {actual}"
                );
            }
            other => panic!("{label} expected number, got {other:?}"),
        };

        let rate = 0.05_f64;
        let nper = 10.0_f64;
        let pv = 1000.0_f64;
        let growth = (1.0_f64 + rate).powf(nper);
        let expected_pmt = -(rate * (0.0 + pv * growth)) / ((1.0 + rate * 0.0) * (growth - 1.0));
        let expected_pmt_type1 =
            -(rate * (0.0 + pv * growth)) / ((1.0 + rate * 1.0) * (growth - 1.0));

        assert_close(value_at(1), expected_pmt, "A2");
        assert_close(value_at(2), 0.0, "B2");
        assert_close(value_at(3), 1000.0, "C2");
        assert_close(value_at(4), 100.0 / 1.1 + 100.0 / (1.1 * 1.1), "D2");
        assert_eq!(value_at(5), CellValue::Number(-100.0));
        assert_eq!(value_at(6), CellValue::Number(1000.0));
        assert_eq!(value_at(7), CellValue::Number(1000.0));
        assert_close(value_at(8), expected_pmt_type1, "H2");
        assert_eq!(value_at(9), CellValue::Number(60.0));
        assert_eq!(value_at(10), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(11), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(12), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(13), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(14), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(15), CellValue::Error("#PARSE!".to_string()));
    }

    #[test]
    fn evaluates_combinatorics_and_number_theory_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(5.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(6.0),
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
            value: CellValue::Number(24.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 6,
            value: CellValue::Number(36.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 7,
            value: CellValue::Number(60.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 8,
            value: CellValue::Number(4.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 9,
            value: CellValue::Number(8.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=FACT(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=FACTDOUBLE(B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=COMBIN(C1,D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=PERMUT(C1,D1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=GCD(E1,F1,G1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=LCM(H1,F1,I1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=LCM(0,5)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=FACT(-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=COMBIN(5,7)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=PERMUT(5,7)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=GCD(-1,5)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=LCM(4,-2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=SUM(FACT(3),COMBIN(6,2),PERMUT(6,2))".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=GCD(0,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=FACTDOUBLE(7)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 16,
            formula: "=COMBIN(52,5)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 17,
            formula: "=PERMUT(10,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 18,
            formula: "=LCM(1,1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 19,
            formula: "=FACT()".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 20,
            formula: "=GCD()".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 7);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Number(120.0));
        assert_eq!(value_at(2), CellValue::Number(48.0));
        assert_eq!(value_at(3), CellValue::Number(120.0));
        assert_eq!(value_at(4), CellValue::Number(720.0));
        assert_eq!(value_at(5), CellValue::Number(12.0));
        assert_eq!(value_at(6), CellValue::Number(72.0));
        assert_eq!(value_at(7), CellValue::Number(0.0));
        assert_eq!(value_at(8), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(9), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(10), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(11), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(12), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(13), CellValue::Number(51.0));
        assert_eq!(value_at(14), CellValue::Number(0.0));
        assert_eq!(value_at(15), CellValue::Number(105.0));
        assert_eq!(value_at(16), CellValue::Number(2598960.0));
        assert_eq!(value_at(17), CellValue::Number(1.0));
        assert_eq!(value_at(18), CellValue::Number(1.0));
        assert_eq!(value_at(19), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(20), CellValue::Error("#PARSE!".to_string()));
    }

    #[test]
    fn evaluates_bitwise_extension_functions() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Number(6.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Number(3.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=BITAND(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=BITOR(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=BITXOR(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=BITLSHIFT(A1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=BITRSHIFT(A1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=BITLSHIFT(A1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=BITRSHIFT(A1,-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=BITLSHIFT(1,48)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=BITAND(-1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=BITOR(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=BITRSHIFT(8,1000)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=BITXOR(281474976710655,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=BITLSHIFT(281474976710655,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=BITRSHIFT(1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 5);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Number(2.0));
        assert_eq!(value_at(2), CellValue::Number(7.0));
        assert_eq!(value_at(3), CellValue::Number(5.0));
        assert_eq!(value_at(4), CellValue::Number(12.0));
        assert_eq!(value_at(5), CellValue::Number(3.0));
        assert_eq!(value_at(6), CellValue::Number(3.0));
        assert_eq!(value_at(7), CellValue::Number(12.0));
        assert_eq!(value_at(8), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(9), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(10), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(11), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(12), CellValue::Number(281474976710654.0));
        assert_eq!(value_at(13), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(14), CellValue::Number(0.0));
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
    fn evaluates_trig_and_log_extension_functions() {
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
            value: CellValue::Number(2.0),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Number(0.5),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            value: CellValue::Number(90.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=PI()".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=EXP(1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=LN(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=LOG(A1,B1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=LOG10(A1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=LOG(100)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=SIN(RADIANS(D1))".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=COS(0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 9,
            formula: "=TAN(0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 10,
            formula: "=ASIN(C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 11,
            formula: "=ACOS(C1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 12,
            formula: "=ATAN(1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 13,
            formula: "=ATAN2(1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 14,
            formula: "=DEGREES(PI())".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 15,
            formula: "=RADIANS(180)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 16,
            formula: "=EXP(1000)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 17,
            formula: "=LN(-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 18,
            formula: "=ACOS(2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 19,
            formula: "=PI(1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 20,
            formula: "=LOG(A1,1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 21,
            formula: "=LOG(A1,-2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 22,
            formula: "=LOG(-10)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 23,
            formula: "=ATAN2(1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 24,
            formula: "=SIN(PI()/2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 25,
            formula: "=SINH(1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 26,
            formula: "=COSH(1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 27,
            formula: "=TANH(1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 28,
            formula: "=ASINH(1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 29,
            formula: "=ACOSH(1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 30,
            formula: "=ATANH(0.5)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 31,
            formula: "=ACOSH(0.5)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 32,
            formula: "=ATANH(1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 33,
            formula: "=SINH(1,2)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 34,
            formula: "=ATANH(-1)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 12);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };
        let assert_close = |value: CellValue, expected: f64, label: &str| match value {
            CellValue::Number(actual) => {
                assert!(
                    (actual - expected).abs() < 1e-12,
                    "{label} expected {expected}, got {actual}"
                );
            }
            other => panic!("{label} expected number, got {other:?}"),
        };

        assert_close(value_at(1), std::f64::consts::PI, "A2");
        assert_close(value_at(2), std::f64::consts::E, "B2");
        assert_close(value_at(3), 8.0_f64.ln(), "C2");
        assert_eq!(value_at(4), CellValue::Number(3.0));
        assert_close(value_at(5), 8.0_f64.log10(), "E2");
        assert_eq!(value_at(6), CellValue::Number(2.0));
        assert_close(value_at(7), 1.0, "G2");
        assert_close(value_at(8), 1.0, "H2");
        assert_close(value_at(9), 0.0, "I2");
        assert_close(value_at(10), std::f64::consts::PI / 6.0, "J2");
        assert_close(value_at(11), std::f64::consts::PI / 3.0, "K2");
        assert_close(value_at(12), std::f64::consts::PI / 4.0, "L2");
        assert_close(value_at(13), std::f64::consts::PI / 4.0, "M2");
        assert_close(value_at(14), 180.0, "N2");
        assert_close(value_at(15), std::f64::consts::PI, "O2");
        assert_eq!(value_at(16), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(17), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(18), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(19), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(20), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(21), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(22), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(23), CellValue::Error("#PARSE!".to_string()));
        assert_close(value_at(24), 1.0, "X2");
        assert_close(value_at(25), 1.0_f64.sinh(), "Y2");
        assert_close(value_at(26), 1.0_f64.cosh(), "Z2");
        assert_close(value_at(27), 1.0_f64.tanh(), "AA2");
        assert_close(value_at(28), 1.0_f64.asinh(), "AB2");
        assert_close(value_at(29), 0.0, "AC2");
        assert_close(value_at(30), 0.5_f64.atanh(), "AD2");
        assert_eq!(value_at(31), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(32), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(33), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(34), CellValue::Error("#PARSE!".to_string()));
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
    fn coerces_text_values_in_match_and_xmatch() {
        let mut wb = Workbook::new();
        let mut sink = NoopEventSink;
        let trace = TraceContext::root();

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            value: CellValue::Text("2".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 2,
            value: CellValue::Text("3".to_string()),
        });
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 3,
            value: CellValue::Text("foo".to_string()),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            formula: "=MATCH(\"2\",1,2,3,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 2,
            formula: "=XMATCH(\"3\",1,2,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 3,
            formula: "=MATCH(A1,1,2,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 4,
            formula: "=XMATCH(B1,1,2,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 5,
            formula: "=MATCH(\"2\",1,2,3,\"0\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 6,
            formula: "=MATCH(\"foo\",1,2,3,0)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 7,
            formula: "=XMATCH(C1,1,2,3)".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 8,
            formula: "=MATCH(2,1,2,\"foo\")".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let report = recalc_sheet(&mut wb, "Sheet1", &mut sink, &trace).expect("recalc");
        assert_eq!(report.parse_error_count, 3);

        let value_at = |col: u32| -> CellValue {
            wb.sheets
                .get("Sheet1")
                .and_then(|s| s.cells.get(&CellRef { row: 2, col }))
                .expect("cell")
                .value
                .clone()
        };

        assert_eq!(value_at(1), CellValue::Number(2.0));
        assert_eq!(value_at(2), CellValue::Number(3.0));
        assert_eq!(value_at(3), CellValue::Number(2.0));
        assert_eq!(value_at(4), CellValue::Number(3.0));
        assert_eq!(value_at(5), CellValue::Number(2.0));
        assert_eq!(value_at(6), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(7), CellValue::Error("#PARSE!".to_string()));
        assert_eq!(value_at(8), CellValue::Error("#PARSE!".to_string()));
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
    fn orders_incremental_formula_subset_by_topo_position_with_cycle_tail() {
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
            formula: "=B1+1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 4,
            formula: "=E1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 5,
            formula: "=D1".to_string(),
            cached_value: CellValue::Empty,
        });
        txn.commit(&mut wb, &mut sink, &trace).expect("commit");

        let formula_cells = collect_formula_cells(&wb, "Sheet1").expect("formula cells");
        let analysis =
            build_sheet_dependency_analysis(&wb, "Sheet1", &formula_cells).expect("analysis");

        let subset = vec![
            CellRef { row: 1, col: 3 },
            CellRef { row: 1, col: 2 },
            CellRef { row: 1, col: 5 },
        ];
        let ordered_subset = order_formula_cells(&analysis, &subset);
        assert_eq!(
            ordered_subset,
            vec![
                CellRef { row: 1, col: 2 },
                CellRef { row: 1, col: 3 },
                CellRef { row: 1, col: 5 }
            ]
        );

        let cycle_only = vec![CellRef { row: 1, col: 5 }, CellRef { row: 1, col: 4 }];
        let ordered_cycle_only = order_formula_cells(&analysis, &cycle_only);
        assert_eq!(
            ordered_cycle_only,
            vec![CellRef { row: 1, col: 4 }, CellRef { row: 1, col: 5 }]
        );

        let full_set = analysis.formula_nodes.iter().copied().collect::<Vec<_>>();
        let ordered_full_set = order_formula_cells(&analysis, &full_set);
        let cached_full_order = order_all_formula_cells(&analysis);
        assert_eq!(ordered_full_set, cached_full_order);
        assert_eq!(
            cached_full_order,
            vec![
                CellRef { row: 1, col: 2 },
                CellRef { row: 1, col: 3 },
                CellRef { row: 1, col: 4 },
                CellRef { row: 1, col: 5 }
            ]
        );
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
