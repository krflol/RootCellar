#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use rootcellar_core::model::CellRef;
use rootcellar_core::{
    inspect_xlsx, load_workbook_model, preserve_xlsx_passthrough,
    preserve_xlsx_with_sheet_overrides, recalc_sheet, save_workbook_model, CellValue,
    CompatibilityIssue, Mutation, NoopEventSink, RecalcReport, SaveMode, TraceContext, Workbook,
    WorkbookPartGraphSummary, XlsxInspectionReport, XlsxSaveReport,
};
use serde::{Deserialize, Serialize};
use std::collections::BTreeSet;
use std::path::PathBuf;
use std::sync::Mutex;
use tauri::State;
use uuid::Uuid;

#[derive(Default)]
struct DesktopState {
    session: Mutex<Option<InteropSession>>,
}

#[derive(Debug, Clone)]
struct InteropSession {
    input_path: PathBuf,
    workbook: Workbook,
    inspection: XlsxInspectionReport,
    dirty_sheets: BTreeSet<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct UiTraceContext {
    trace_id: Option<String>,
    span_id: Option<String>,
    parent_span_id: Option<String>,
    session_id: Option<String>,
}

impl UiTraceContext {
    fn into_trace_context(self) -> TraceContext {
        let fallback = TraceContext::root();
        TraceContext {
            trace_id: self
                .trace_id
                .as_deref()
                .and_then(parse_uuid)
                .unwrap_or(fallback.trace_id),
            span_id: self
                .span_id
                .as_deref()
                .and_then(parse_uuid)
                .unwrap_or(fallback.span_id),
            parent_span_id: self.parent_span_id.as_deref().and_then(parse_uuid),
            session_id: self
                .session_id
                .as_deref()
                .and_then(parse_uuid)
                .or(fallback.session_id),
        }
    }
}

fn parse_uuid(value: &str) -> Option<Uuid> {
    Uuid::parse_str(value).ok()
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct AppStatusResponse {
    app: &'static str,
    ui_ready: bool,
    engine_ready: bool,
    interop_ready: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct TraceEcho {
    trace_id: String,
    span_id: String,
    parent_span_id: Option<String>,
    session_id: Option<String>,
}

impl From<&TraceContext> for TraceEcho {
    fn from(value: &TraceContext) -> Self {
        Self {
            trace_id: value.trace_id.to_string(),
            span_id: value.span_id.to_string(),
            parent_span_id: value.parent_span_id.map(|id| id.to_string()),
            session_id: value.session_id.map(|id| id.to_string()),
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct EngineRoundTripResponse {
    sheet: String,
    formula_cell: String,
    value: CellValue,
    evaluated_cells: usize,
    cycle_count: usize,
    parse_error_count: usize,
    workbook_id: String,
    trace: TraceEcho,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropSessionStatusResponse {
    loaded: bool,
    input_path: Option<String>,
    workbook_id: Option<String>,
    sheet_count: usize,
    cell_count: usize,
    issue_count: usize,
    unknown_part_count: usize,
    dirty_sheet_count: usize,
    dirty_sheets: Vec<String>,
    sheets: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropOpenResponse {
    input_path: String,
    workbook_id: String,
    sheet_count: usize,
    cell_count: usize,
    sheets: Vec<String>,
    feature_score: u8,
    issue_count: usize,
    unknown_part_count: usize,
    issues: Vec<CompatibilityIssue>,
    unknown_parts: Vec<String>,
    part_graph: WorkbookPartGraphSummary,
    trace: TraceEcho,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropRecalcResponse {
    workbook_id: String,
    reports: Vec<RecalcReport>,
    trace: TraceEcho,
}

#[derive(Debug, Clone, Copy, Deserialize)]
#[serde(rename_all = "snake_case")]
enum UiEditMode {
    Value,
    Formula,
}

impl UiEditMode {
    fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::Formula => "formula",
        }
    }
}

#[derive(Debug, Clone, Copy, Deserialize)]
#[serde(rename_all = "snake_case")]
enum UiSaveMode {
    Preserve,
    Normalize,
}

impl UiSaveMode {
    fn as_str(self) -> &'static str {
        match self {
            Self::Preserve => "preserve",
            Self::Normalize => "normalize",
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropSaveResponse {
    input_path: String,
    output_path: String,
    workbook_id: String,
    mode: String,
    sheet_count: usize,
    cell_count: usize,
    copied_bytes: u64,
    part_graph: WorkbookPartGraphSummary,
    part_graph_flags: rootcellar_core::SavePartGraphFlags,
    trace: TraceEcho,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropCellEditResponse {
    workbook_id: String,
    sheet: String,
    cell: String,
    anchor_cell: String,
    applied_cell_count: usize,
    mode: String,
    value: CellValue,
    formula: Option<String>,
    dirty_sheet_count: usize,
    dirty_sheets: Vec<String>,
    trace: TraceEcho,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropPreviewCell {
    cell: String,
    row: u32,
    col: u32,
    value: CellValue,
    formula: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropSheetPreviewResponse {
    workbook_id: String,
    sheet: String,
    total_cells: usize,
    shown_cells: usize,
    truncated: bool,
    cells: Vec<InteropPreviewCell>,
    trace: TraceEcho,
}

fn workbook_sheet_names(workbook: &Workbook) -> Vec<String> {
    workbook.sheets.keys().cloned().collect()
}

fn workbook_cell_count(workbook: &Workbook) -> usize {
    workbook
        .sheets
        .values()
        .map(|sheet| sheet.cells.len())
        .sum::<usize>()
}

fn normalize_input_path(path: &str) -> Result<PathBuf, String> {
    let trimmed = path.trim();
    if trimmed.is_empty() {
        return Err("workbook path is required".to_string());
    }

    let provided = PathBuf::from(trimmed);
    Ok(provided.canonicalize().unwrap_or(provided))
}

fn map_session_status(session: Option<&InteropSession>) -> InteropSessionStatusResponse {
    match session {
        Some(session) => InteropSessionStatusResponse {
            loaded: true,
            input_path: Some(session.input_path.display().to_string()),
            workbook_id: Some(session.workbook.workbook_id.to_string()),
            sheet_count: session.workbook.sheets.len(),
            cell_count: workbook_cell_count(&session.workbook),
            issue_count: session.inspection.summary.issue_count,
            unknown_part_count: session.inspection.summary.unknown_part_count,
            dirty_sheet_count: session.dirty_sheets.len(),
            dirty_sheets: session.dirty_sheets.iter().cloned().collect(),
            sheets: workbook_sheet_names(&session.workbook),
        },
        None => InteropSessionStatusResponse {
            loaded: false,
            input_path: None,
            workbook_id: None,
            sheet_count: 0,
            cell_count: 0,
            issue_count: 0,
            unknown_part_count: 0,
            dirty_sheet_count: 0,
            dirty_sheets: Vec::new(),
            sheets: Vec::new(),
        },
    }
}

fn map_open_response(
    input_path: &PathBuf,
    workbook: &Workbook,
    inspection: &XlsxInspectionReport,
    trace: &TraceContext,
) -> InteropOpenResponse {
    InteropOpenResponse {
        input_path: input_path.display().to_string(),
        workbook_id: workbook.workbook_id.to_string(),
        sheet_count: workbook.sheets.len(),
        cell_count: workbook_cell_count(workbook),
        sheets: workbook_sheet_names(workbook),
        feature_score: inspection.summary.workbook_feature_score,
        issue_count: inspection.summary.issue_count,
        unknown_part_count: inspection.summary.unknown_part_count,
        issues: inspection.issues.clone(),
        unknown_parts: inspection.unknown_parts.clone(),
        part_graph: inspection.part_graph.summary(),
        trace: TraceEcho::from(trace),
    }
}

fn map_save_response(
    session: &InteropSession,
    report: XlsxSaveReport,
    mode: UiSaveMode,
    trace: &TraceContext,
) -> InteropSaveResponse {
    InteropSaveResponse {
        input_path: session.input_path.display().to_string(),
        output_path: report.output_path.display().to_string(),
        workbook_id: session.workbook.workbook_id.to_string(),
        mode: mode.as_str().to_string(),
        sheet_count: report.sheet_count,
        cell_count: report.cell_count,
        copied_bytes: report.copied_bytes,
        part_graph: report.part_graph,
        part_graph_flags: report.part_graph_flags,
        trace: TraceEcho::from(trace),
    }
}

fn parse_a1_cell_ref(value: &str) -> Option<(u32, u32)> {
    let input = value.trim().to_ascii_uppercase();
    if input.is_empty() {
        return None;
    }

    let bytes = input.as_bytes();
    let mut index = 0usize;
    let mut col: u32 = 0;
    while index < bytes.len() {
        let ch = bytes[index];
        if !ch.is_ascii_alphabetic() {
            break;
        }
        col = col
            .saturating_mul(26)
            .saturating_add((ch - b'A' + 1) as u32);
        index += 1;
    }

    if col == 0 || index >= bytes.len() {
        return None;
    }

    let row_str = &input[index..];
    if !row_str.chars().all(|ch| ch.is_ascii_digit()) {
        return None;
    }
    let row = row_str.parse::<u32>().ok()?;
    if row == 0 {
        return None;
    }

    Some((row, col))
}

fn parse_a1_range_ref(value: &str) -> Option<((u32, u32), (u32, u32))> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return None;
    }

    if let Some((left, right)) = trimmed.split_once(':') {
        let (left_row, left_col) = parse_a1_cell_ref(left)?;
        let (right_row, right_col) = parse_a1_cell_ref(right)?;
        let start_row = left_row.min(right_row);
        let end_row = left_row.max(right_row);
        let start_col = left_col.min(right_col);
        let end_col = left_col.max(right_col);
        Some(((start_row, start_col), (end_row, end_col)))
    } else {
        let (row, col) = parse_a1_cell_ref(trimmed)?;
        Some(((row, col), (row, col)))
    }
}

fn parse_cell_value_input(value: &str) -> CellValue {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        CellValue::Empty
    } else if trimmed.eq_ignore_ascii_case("true") {
        CellValue::Bool(true)
    } else if trimmed.eq_ignore_ascii_case("false") {
        CellValue::Bool(false)
    } else if let Ok(number) = trimmed.parse::<f64>() {
        CellValue::Number(number)
    } else {
        CellValue::Text(value.to_string())
    }
}

fn normalize_formula_input(value: &str) -> Result<String, String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return Err("formula input is required".to_string());
    }

    if trimmed.starts_with('=') {
        Ok(trimmed.to_string())
    } else {
        Ok(format!("={trimmed}"))
    }
}

fn column_to_a1(mut col: u32) -> String {
    let mut chars = Vec::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        chars.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    chars.into_iter().rev().collect()
}

fn to_a1(row: u32, col: u32) -> String {
    format!("{}{}", column_to_a1(col), row)
}

fn apply_cell_edit_to_workbook(
    workbook: &mut Workbook,
    dirty_sheets: &mut BTreeSet<String>,
    sheet: &str,
    cell: &str,
    input: &str,
    mode: UiEditMode,
    command_trace: &TraceContext,
) -> Result<InteropCellEditResponse, String> {
    let mut sink = NoopEventSink;

    let resolved_sheet = {
        let trimmed = sheet.trim();
        if trimmed.is_empty() {
            workbook_sheet_names(workbook)
                .into_iter()
                .next()
                .ok_or_else(|| "loaded workbook has no sheets".to_string())?
        } else {
            trimmed.to_string()
        }
    };

    if !workbook.sheets.contains_key(&resolved_sheet) {
        return Err(format!("sheet not found: {resolved_sheet}"));
    }

    let normalized_target = cell.trim().to_ascii_uppercase();
    let ((start_row, start_col), (end_row, end_col)) = parse_a1_range_ref(&normalized_target)
        .ok_or_else(|| format!("invalid A1 cell/range reference: {}", cell.trim()))?;
    let cell_count = (end_row - start_row + 1) as usize * (end_col - start_col + 1) as usize;
    if cell_count > 20_000 {
        return Err("range too large; max 20,000 cells per edit".to_string());
    }

    let mut txn = workbook
        .begin_txn(&mut sink, command_trace)
        .map_err(|error| error.to_string())?;

    match mode {
        UiEditMode::Value => {
            let parsed_value = parse_cell_value_input(input);
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    txn.apply(Mutation::SetCellValue {
                        sheet: resolved_sheet.clone(),
                        row,
                        col,
                        value: parsed_value.clone(),
                    });
                }
            }
        }
        UiEditMode::Formula => {
            let formula = normalize_formula_input(input)?;
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    txn.apply(Mutation::SetCellFormula {
                        sheet: resolved_sheet.clone(),
                        row,
                        col,
                        formula: formula.clone(),
                        cached_value: CellValue::Empty,
                    });
                }
            }
        }
    }

    txn.commit(workbook, &mut sink, command_trace)
        .map_err(|error| error.to_string())?;
    dirty_sheets.insert(resolved_sheet.clone());

    let cell_ref = CellRef {
        row: start_row,
        col: start_col,
    };
    let cell_record = workbook
        .sheets
        .get(&resolved_sheet)
        .and_then(|sheet_data| sheet_data.cells.get(&cell_ref))
        .ok_or_else(|| "edited cell not found after commit".to_string())?;
    let anchor_cell = to_a1(start_row, start_col);

    Ok(InteropCellEditResponse {
        workbook_id: workbook.workbook_id.to_string(),
        sheet: resolved_sheet,
        cell: normalized_target,
        anchor_cell,
        applied_cell_count: cell_count,
        mode: mode.as_str().to_string(),
        value: cell_record.value.clone(),
        formula: cell_record.formula.clone(),
        dirty_sheet_count: dirty_sheets.len(),
        dirty_sheets: dirty_sheets.iter().cloned().collect(),
        trace: TraceEcho::from(command_trace),
    })
}

#[tauri::command]
fn app_status() -> AppStatusResponse {
    AppStatusResponse {
        app: "rootcellar-desktop-shell",
        ui_ready: true,
        engine_ready: true,
        interop_ready: true,
    }
}

#[tauri::command]
fn engine_round_trip(trace: Option<UiTraceContext>) -> Result<EngineRoundTripResponse, String> {
    let incoming_trace = trace
        .map(UiTraceContext::into_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let mut sink = NoopEventSink;

    let mut workbook = Workbook::new();
    let mut txn = workbook
        .begin_txn(&mut sink, &command_trace)
        .map_err(|error| error.to_string())?;

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

    txn.commit(&mut workbook, &mut sink, &command_trace)
        .map_err(|error| error.to_string())?;

    let report = recalc_sheet(&mut workbook, "Sheet1", &mut sink, &command_trace)
        .map_err(|error| error.to_string())?;

    let value = workbook
        .sheets
        .get("Sheet1")
        .and_then(|sheet| {
            sheet
                .cells
                .iter()
                .find(|(cell_ref, _)| cell_ref.row == 1 && cell_ref.col == 3)
                .map(|(_, record)| record.value.clone())
        })
        .ok_or_else(|| "formula cell C1 not found after recalc".to_string())?;

    Ok(EngineRoundTripResponse {
        sheet: "Sheet1".to_string(),
        formula_cell: "C1".to_string(),
        value,
        evaluated_cells: report.evaluated_cells,
        cycle_count: report.cycle_count,
        parse_error_count: report.parse_error_count,
        workbook_id: workbook.workbook_id.to_string(),
        trace: TraceEcho::from(&command_trace),
    })
}

#[tauri::command]
fn interop_open_workbook(
    state: State<DesktopState>,
    path: String,
    trace: Option<UiTraceContext>,
) -> Result<InteropOpenResponse, String> {
    let input_path = normalize_input_path(&path)?;
    let incoming_trace = trace
        .map(UiTraceContext::into_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let mut sink = NoopEventSink;

    let inspection =
        inspect_xlsx(&input_path, &mut sink, &command_trace).map_err(|error| error.to_string())?;
    let workbook = load_workbook_model(&input_path, &mut sink, &command_trace)
        .map_err(|error| error.to_string())?;
    let response = map_open_response(&input_path, &workbook, &inspection, &command_trace);

    let mut session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    *session_guard = Some(InteropSession {
        input_path,
        workbook,
        inspection,
        dirty_sheets: BTreeSet::new(),
    });

    Ok(response)
}

#[tauri::command]
fn interop_session_status(
    state: State<DesktopState>,
) -> Result<InteropSessionStatusResponse, String> {
    let session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    Ok(map_session_status(session_guard.as_ref()))
}

#[tauri::command]
fn interop_recalc_loaded(
    state: State<DesktopState>,
    sheet: Option<String>,
    trace: Option<UiTraceContext>,
) -> Result<InteropRecalcResponse, String> {
    let incoming_trace = trace
        .map(UiTraceContext::into_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let mut sink = NoopEventSink;

    let mut session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    let session = session_guard
        .as_mut()
        .ok_or_else(|| "no workbook loaded; open a workbook first".to_string())?;

    let target_sheets = match sheet
        .as_deref()
        .map(str::trim)
        .filter(|value| !value.is_empty())
    {
        Some(name) => vec![name.to_string()],
        None => workbook_sheet_names(&session.workbook),
    };

    if target_sheets.is_empty() {
        return Err("loaded workbook has no sheets".to_string());
    }

    let mut reports = Vec::with_capacity(target_sheets.len());
    for sheet_name in target_sheets {
        let report = recalc_sheet(
            &mut session.workbook,
            &sheet_name,
            &mut sink,
            &command_trace,
        )
        .map_err(|error| error.to_string())?;
        session.dirty_sheets.insert(sheet_name);
        reports.push(report);
    }

    Ok(InteropRecalcResponse {
        workbook_id: session.workbook.workbook_id.to_string(),
        reports,
        trace: TraceEcho::from(&command_trace),
    })
}

#[tauri::command]
fn interop_apply_cell_edit(
    state: State<DesktopState>,
    sheet: String,
    cell: String,
    input: String,
    mode: UiEditMode,
    trace: Option<UiTraceContext>,
) -> Result<InteropCellEditResponse, String> {
    let incoming_trace = trace
        .map(UiTraceContext::into_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();

    let mut session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    let session = session_guard
        .as_mut()
        .ok_or_else(|| "no workbook loaded; open a workbook first".to_string())?;

    apply_cell_edit_to_workbook(
        &mut session.workbook,
        &mut session.dirty_sheets,
        &sheet,
        &cell,
        &input,
        mode,
        &command_trace,
    )
}

#[tauri::command]
fn interop_sheet_preview(
    state: State<DesktopState>,
    sheet: Option<String>,
    limit: Option<usize>,
    trace: Option<UiTraceContext>,
) -> Result<InteropSheetPreviewResponse, String> {
    let incoming_trace = trace
        .map(UiTraceContext::into_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();

    let session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    let session = session_guard
        .as_ref()
        .ok_or_else(|| "no workbook loaded; open a workbook first".to_string())?;

    let resolved_sheet = match sheet
        .as_deref()
        .map(str::trim)
        .filter(|value| !value.is_empty())
    {
        Some(name) => name.to_string(),
        None => workbook_sheet_names(&session.workbook)
            .into_iter()
            .next()
            .ok_or_else(|| "loaded workbook has no sheets".to_string())?,
    };

    let sheet_data = session
        .workbook
        .sheets
        .get(&resolved_sheet)
        .ok_or_else(|| format!("sheet not found: {resolved_sheet}"))?;

    let preview_limit = limit.unwrap_or(120).clamp(1, 400);
    let cells = sheet_data
        .cells
        .iter()
        .take(preview_limit)
        .map(|(cell_ref, cell_record)| InteropPreviewCell {
            cell: to_a1(cell_ref.row, cell_ref.col),
            row: cell_ref.row,
            col: cell_ref.col,
            value: cell_record.value.clone(),
            formula: cell_record.formula.clone(),
        })
        .collect::<Vec<_>>();

    Ok(InteropSheetPreviewResponse {
        workbook_id: session.workbook.workbook_id.to_string(),
        sheet: resolved_sheet,
        total_cells: sheet_data.cells.len(),
        shown_cells: cells.len(),
        truncated: sheet_data.cells.len() > cells.len(),
        cells,
        trace: TraceEcho::from(&command_trace),
    })
}

#[tauri::command]
fn interop_save_workbook(
    state: State<DesktopState>,
    output_path: String,
    mode: UiSaveMode,
    promote_output_as_input: Option<bool>,
    trace: Option<UiTraceContext>,
) -> Result<InteropSaveResponse, String> {
    let output_path = output_path.trim();
    if output_path.is_empty() {
        return Err("output path is required".to_string());
    }

    let incoming_trace = trace
        .map(UiTraceContext::into_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let mut sink = NoopEventSink;

    let mut session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    let session = session_guard
        .as_mut()
        .ok_or_else(|| "no workbook loaded; open a workbook first".to_string())?;

    let report = match mode {
        UiSaveMode::Preserve if session.dirty_sheets.is_empty() => preserve_xlsx_passthrough(
            &session.input_path,
            output_path,
            &session.workbook,
            &mut sink,
            &command_trace,
        ),
        UiSaveMode::Preserve => {
            let changed_sheets = session.dirty_sheets.iter().cloned().collect::<Vec<_>>();
            preserve_xlsx_with_sheet_overrides(
                &session.input_path,
                output_path,
                &session.workbook,
                &changed_sheets,
                &mut sink,
                &command_trace,
            )
        }
        UiSaveMode::Normalize => save_workbook_model(
            &session.workbook,
            output_path,
            SaveMode::Normalize,
            &mut sink,
            &command_trace,
        ),
    }
    .map_err(|error| error.to_string())?;
    let saved_output_path = report.output_path.clone();
    let response = map_save_response(session, report, mode, &command_trace);

    if promote_output_as_input.unwrap_or(false) {
        session.input_path = saved_output_path
            .canonicalize()
            .unwrap_or(saved_output_path);
        session.dirty_sheets.clear();
    }

    Ok(response)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn seeded_workbook(trace: &TraceContext) -> Workbook {
        let mut sink = NoopEventSink;
        let mut workbook = Workbook::new();
        let mut txn = workbook
            .begin_txn(&mut sink, trace)
            .expect("seed transaction should start");

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

        txn.commit(&mut workbook, &mut sink, trace)
            .expect("seed transaction should commit");
        workbook
    }

    fn sheet_cell_value(workbook: &Workbook, row: u32, col: u32) -> CellValue {
        workbook
            .sheets
            .get("Sheet1")
            .and_then(|sheet| sheet.cells.get(&CellRef { row, col }))
            .map(|record| record.value.clone())
            .expect("cell should exist in Sheet1")
    }

    fn sheet_cell_formula(workbook: &Workbook, row: u32, col: u32) -> Option<String> {
        workbook
            .sheets
            .get("Sheet1")
            .and_then(|sheet| sheet.cells.get(&CellRef { row, col }))
            .map(|record| record.formula.clone())
            .expect("cell should exist in Sheet1")
    }

    #[test]
    fn apply_cell_edit_range_value_sets_anchor_and_all_cells() {
        let trace = TraceContext::root();
        let mut workbook = seeded_workbook(&trace);
        let mut dirty_sheets = BTreeSet::new();

        let response = apply_cell_edit_to_workbook(
            &mut workbook,
            &mut dirty_sheets,
            "Sheet1",
            "B2:A1",
            "9",
            UiEditMode::Value,
            &trace,
        )
        .expect("range value edit should succeed");

        assert_eq!(response.sheet, "Sheet1");
        assert_eq!(response.cell, "B2:A1");
        assert_eq!(response.anchor_cell, "A1");
        assert_eq!(response.applied_cell_count, 4);
        assert_eq!(response.value, CellValue::Number(9.0));
        assert_eq!(response.formula, None);
        assert_eq!(response.dirty_sheet_count, 1);
        assert_eq!(response.dirty_sheets, vec!["Sheet1".to_string()]);

        for row in 1..=2 {
            for col in 1..=2 {
                assert_eq!(
                    sheet_cell_value(&workbook, row, col),
                    CellValue::Number(9.0)
                );
                assert_eq!(sheet_cell_formula(&workbook, row, col), None);
            }
        }
    }

    #[test]
    fn apply_cell_edit_range_formula_normalizes_and_updates_all_cells() {
        let trace = TraceContext::root();
        let mut workbook = seeded_workbook(&trace);
        let mut dirty_sheets = BTreeSet::new();

        let response = apply_cell_edit_to_workbook(
            &mut workbook,
            &mut dirty_sheets,
            "Sheet1",
            "C1:C2",
            "A1+B1",
            UiEditMode::Formula,
            &trace,
        )
        .expect("range formula edit should succeed");

        assert_eq!(response.sheet, "Sheet1");
        assert_eq!(response.cell, "C1:C2");
        assert_eq!(response.anchor_cell, "C1");
        assert_eq!(response.applied_cell_count, 2);
        assert_eq!(response.value, CellValue::Empty);
        assert_eq!(response.formula.as_deref(), Some("=A1+B1"));

        assert_eq!(
            sheet_cell_formula(&workbook, 1, 3).as_deref(),
            Some("=A1+B1")
        );
        assert_eq!(
            sheet_cell_formula(&workbook, 2, 3).as_deref(),
            Some("=A1+B1")
        );
    }
}

fn main() {
    tauri::Builder::default()
        .manage(DesktopState::default())
        .invoke_handler(tauri::generate_handler![
            app_status,
            engine_round_trip,
            interop_open_workbook,
            interop_session_status,
            interop_recalc_loaded,
            interop_apply_cell_edit,
            interop_sheet_preview,
            interop_save_workbook
        ])
        .plugin(tauri_plugin_dialog::init())
        .run(tauri::generate_context!())
        .expect("failed to run rootcellar desktop shell");
}
