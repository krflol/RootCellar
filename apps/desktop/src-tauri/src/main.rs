#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod script;

use rootcellar_core::model::CellRef;
use rootcellar_core::{
    inspect_xlsx, load_workbook_model, preserve_xlsx_passthrough,
    preserve_xlsx_with_sheet_overrides, recalc_sheet, recalc_sheet_from_roots, save_workbook_model,
    CellValue, CompatibilityIssue, EventEnvelope, EventSink, JsonlEventSink, Mutation,
    NoopEventSink, RecalcReport, SaveMode, TraceContext, Workbook, WorkbookPartGraphSummary,
    XlsxInspectionReport, XlsxSaveReport,
};
use serde::{Deserialize, Serialize};
use serde_json::json;
use std::collections::{hash_map::DefaultHasher, BTreeMap, BTreeSet};
use std::fs::OpenOptions;
use std::hash::{Hash, Hasher};
use std::io::{BufWriter, Write};
use std::path::Path;
use std::path::PathBuf;
use std::sync::Mutex;
use std::time::Instant;
use tauri::State;
use uuid::Uuid;

const MAX_INTEROP_HISTORY_STEPS: usize = 50;

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
    undo_stack: Vec<InteropSessionSnapshot>,
    redo_stack: Vec<InteropSessionSnapshot>,
}

#[derive(Debug, Clone)]
struct InteropSessionSnapshot {
    workbook: Workbook,
    dirty_sheets: BTreeSet<String>,
}

impl InteropSession {
    fn push_undo_snapshot(&mut self) {
        self.undo_stack.push(InteropSessionSnapshot {
            workbook: self.workbook.clone(),
            dirty_sheets: self.dirty_sheets.clone(),
        });
        enforce_history_limit(&mut self.undo_stack, MAX_INTEROP_HISTORY_STEPS);
        self.redo_stack.clear();
    }

    fn undo(&mut self) -> Result<(), String> {
        let previous_state = self
            .undo_stack
            .pop()
            .ok_or_else(|| "no undo history available".to_string())?;

        let current_state = InteropSessionSnapshot {
            workbook: self.workbook.clone(),
            dirty_sheets: self.dirty_sheets.clone(),
        };
        self.redo_stack.push(current_state);
        enforce_history_limit(&mut self.redo_stack, MAX_INTEROP_HISTORY_STEPS);

        self.workbook = previous_state.workbook;
        self.dirty_sheets = previous_state.dirty_sheets;
        Ok(())
    }

    fn redo(&mut self) -> Result<(), String> {
        let next_state = self
            .redo_stack
            .pop()
            .ok_or_else(|| "no redo history available".to_string())?;

        let current_state = InteropSessionSnapshot {
            workbook: self.workbook.clone(),
            dirty_sheets: self.dirty_sheets.clone(),
        };
        self.undo_stack.push(current_state);
        enforce_history_limit(&mut self.undo_stack, MAX_INTEROP_HISTORY_STEPS);

        self.workbook = next_state.workbook;
        self.dirty_sheets = next_state.dirty_sheets;
        Ok(())
    }

    fn history_depths(&self) -> (usize, usize) {
        (self.undo_stack.len(), self.redo_stack.len())
    }
}

fn enforce_history_limit(stack: &mut Vec<InteropSessionSnapshot>, max_steps: usize) {
    if stack.len() > max_steps {
        let overflow = stack.len() - max_steps;
        stack.drain(0..overflow);
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct UiTraceContext {
    trace_id: Option<String>,
    span_id: Option<String>,
    parent_span_id: Option<String>,
    session_id: Option<String>,
    command_id: Option<String>,
    command_name: Option<String>,
    event_log_path: Option<String>,
    artifact_index_path: Option<String>,
}

impl UiTraceContext {
    fn as_trace_context(&self) -> TraceContext {
        self.clone().into_trace_context()
    }

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
            command_id: self.command_id.as_deref().and_then(parse_uuid),
            command_name: self.command_name,
        }
    }
}

fn parse_uuid(value: &str) -> Option<Uuid> {
    Uuid::parse_str(value).ok()
}

const DESKTOP_EVENT_JSONL_ENV_VAR: &str = "ROOTCELLAR_DESKTOP_EVENT_JSONL";
const DESKTOP_ARTIFACT_INDEX_JSONL_ENV_VAR: &str = "ROOTCELLAR_DESKTOP_ARTIFACT_INDEX";

fn make_event_sink(trace: Option<&UiTraceContext>) -> Result<Box<dyn EventSink>, String> {
    let path = resolve_event_log_path(trace).map(PathBuf::from);

    match path {
        Some(path) => Ok(Box::new(
            JsonlEventSink::new_append(&path).map_err(|error| error.to_string())?,
        )),
        None => Ok(Box::new(NoopEventSink)),
    }
}

fn resolve_event_log_path(trace: Option<&UiTraceContext>) -> Option<String> {
    trace
        .and_then(|value| value.event_log_path.as_deref())
        .filter(|value| !value.trim().is_empty())
        .map(std::string::ToString::to_string)
        .or_else(|| {
            std::env::var(DESKTOP_EVENT_JSONL_ENV_VAR)
                .ok()
                .filter(|value| !value.trim().is_empty())
        })
}

fn make_artifact_index_path(trace: Option<&UiTraceContext>) -> Option<PathBuf> {
    trace
        .and_then(|value| value.artifact_index_path.as_deref())
        .filter(|value| !value.trim().is_empty())
        .map(PathBuf::from)
        .or_else(|| {
            std::env::var(DESKTOP_ARTIFACT_INDEX_JSONL_ENV_VAR)
                .ok()
                .filter(|value| !value.trim().is_empty())
                .map(PathBuf::from)
        })
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct TraceArtifactIndexRecord {
    command_name: String,
    command_status: String,
    duration_ms: u128,
    trace_id: String,
    trace_root_id: String,
    span_id: String,
    parent_span_id: Option<String>,
    session_id: Option<String>,
    ui_command_id: Option<String>,
    ui_command_name: Option<String>,
    event_log_path: Option<String>,
    linked_artifact_ids: Vec<String>,
    artifact_refs: Vec<TraceArtifactRef>,
}

fn append_trace_artifact_index(
    artifact_index_path: &Path,
    trace: &TraceEcho,
    command_status: &str,
    duration_ms: u128,
    event_log_path: Option<&str>,
    artifact_refs: &[TraceArtifactRef],
) -> Result<(), String> {
    let record = TraceArtifactIndexRecord {
        command_name: trace
            .ui_command_name
            .clone()
            .unwrap_or_else(|| "unknown".to_string()),
        command_status: command_status.to_string(),
        duration_ms,
        trace_id: trace.trace_id.clone(),
        trace_root_id: trace.trace_root_id.clone(),
        span_id: trace.span_id.clone(),
        parent_span_id: trace.parent_span_id.clone(),
        session_id: trace.session_id.clone(),
        ui_command_id: trace.ui_command_id.clone(),
        ui_command_name: trace.ui_command_name.clone(),
        event_log_path: event_log_path.map(std::string::ToString::to_string),
        linked_artifact_ids: trace.linked_artifact_ids.clone(),
        artifact_refs: artifact_refs.to_vec(),
    };

    let file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(artifact_index_path)
        .map_err(|error| error.to_string())?;
    let mut writer = BufWriter::new(file);
    serde_json::to_writer(&mut writer, &record).map_err(|error| error.to_string())?;
    writer.write_all(b"\n").map_err(|error| error.to_string())?;
    writer.flush().map_err(|error| error.to_string())?;
    Ok(())
}

fn finalize_command_trace(
    command_trace: &TraceContext,
    command_status: &str,
    duration_ms: u128,
    artifact_refs: Vec<TraceArtifactRef>,
    artifact_index_path: Option<&Path>,
    event_log_path: Option<&str>,
) -> Result<TraceEcho, String> {
    let trace = finalize_trace(
        command_trace,
        command_status,
        duration_ms,
        artifact_refs.clone(),
        artifact_index_path,
        event_log_path,
    );

    if let Some(index_path) = artifact_index_path {
        append_trace_artifact_index(
            index_path,
            &trace,
            command_status,
            duration_ms,
            event_log_path,
            &artifact_refs,
        )?;
    }
    Ok(trace)
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct AppStatusResponse {
    app: &'static str,
    ui_ready: bool,
    engine_ready: bool,
    interop_ready: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct TraceArtifactRef {
    artifact_id: String,
    artifact_type: String,
    relation: String,
    path: Option<String>,
    description: Option<String>,
}

fn trace_artifact_id(
    trace: &TraceContext,
    command_name: &str,
    artifact_type: &str,
    relation: &str,
    path: Option<&str>,
) -> String {
    let mut path_hasher = DefaultHasher::new();
    let payload = match path {
        Some(path) => format!(
            "command={command_name}|type={artifact_type}|relation={relation}|path={path}|trace_id={}|span_id={}",
            trace.trace_id, trace.span_id
        ),
        None => format!(
            "command={command_name}|type={artifact_type}|relation={relation}|trace_id={}|span_id={}",
            trace.trace_id, trace.span_id
        ),
    };
    payload.hash(&mut path_hasher);
    format!("{:016x}", path_hasher.finish())
}

fn trace_artifact_ref(
    trace: &TraceContext,
    command_name: &str,
    artifact_type: &str,
    relation: &str,
    path: Option<&str>,
    description: Option<String>,
) -> TraceArtifactRef {
    TraceArtifactRef {
        artifact_id: trace_artifact_id(trace, command_name, artifact_type, relation, path),
        artifact_type: artifact_type.to_string(),
        relation: relation.to_string(),
        path: path.map(std::string::ToString::to_string),
        description,
    }
}

fn finalize_trace(
    trace: &TraceContext,
    command_status: &str,
    duration_ms: u128,
    artifact_refs: Vec<TraceArtifactRef>,
    artifact_index_path: Option<&Path>,
    event_log_path: Option<&str>,
) -> TraceEcho {
    let linked_artifact_ids = artifact_refs
        .iter()
        .map(|entry| entry.artifact_id.clone())
        .collect::<Vec<_>>();
    TraceEcho {
        trace_id: trace.trace_id.to_string(),
        trace_root_id: trace.trace_id.to_string(),
        span_id: trace.span_id.to_string(),
        parent_span_id: trace.parent_span_id.map(|id| id.to_string()),
        session_id: trace.session_id.map(|id| id.to_string()),
        ui_command_id: trace.command_id.map(|id| id.to_string()),
        ui_command_name: trace.command_name.clone(),
        command_status: command_status.to_string(),
        duration_ms,
        event_log_path: event_log_path.map(std::string::ToString::to_string),
        linked_artifact_ids,
        artifact_refs,
        artifact_index_path: artifact_index_path.map(|path| path.to_string_lossy().to_string()),
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct TraceEcho {
    trace_id: String,
    trace_root_id: String,
    span_id: String,
    parent_span_id: Option<String>,
    session_id: Option<String>,
    ui_command_id: Option<String>,
    ui_command_name: Option<String>,
    command_status: String,
    duration_ms: u128,
    event_log_path: Option<String>,
    linked_artifact_ids: Vec<String>,
    artifact_index_path: Option<String>,
    artifact_refs: Vec<TraceArtifactRef>,
}

impl From<&TraceContext> for TraceEcho {
    fn from(value: &TraceContext) -> Self {
        Self {
            trace_id: value.trace_id.to_string(),
            trace_root_id: value.trace_id.to_string(),
            span_id: value.span_id.to_string(),
            parent_span_id: value.parent_span_id.map(|id| id.to_string()),
            session_id: value.session_id.map(|id| id.to_string()),
            ui_command_id: value.command_id.map(|id| id.to_string()),
            ui_command_name: value.command_name.clone(),
            command_status: "not_measured".to_string(),
            duration_ms: 0,
            event_log_path: None,
            linked_artifact_ids: Vec::new(),
            artifact_index_path: None,
            artifact_refs: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropScriptPermissionEvent {
    event_name: String,
    permission: String,
    allowed: bool,
    reason: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropScriptRuntimeEvent {
    event_name: String,
    payload: serde_json::Value,
    severity: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropMacroTrustProvenance {
    mode: String,
    manifest_path: Option<String>,
    manifest_name: Option<String>,
    manifest_version: Option<String>,
    publisher: Option<String>,
    api_min_version: Option<u32>,
    permissions_required: Vec<String>,
    permissions_declared: Vec<String>,
    runtime_api_version: u32,
    signature_present: bool,
    signature_verified: Option<bool>,
    fingerprint: String,
    trusted: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct InteropMacroPermissionConfig {
    fs_read: bool,
    fs_write: bool,
    net_http: bool,
    clipboard: bool,
    process_exec: bool,
    udf: bool,
    events_emit: bool,
}

impl InteropMacroPermissionConfig {
    fn as_requested_permissions(&self) -> Vec<String> {
        let mut requested = Vec::new();
        if self.fs_read {
            requested.push(script::ScriptPermission::FsRead.as_str().to_string());
        }
        if self.fs_write {
            requested.push(script::ScriptPermission::FsWrite.as_str().to_string());
        }
        if self.net_http {
            requested.push(script::ScriptPermission::NetHttp.as_str().to_string());
        }
        if self.clipboard {
            requested.push(script::ScriptPermission::Clipboard.as_str().to_string());
        }
        if self.process_exec {
            requested.push(script::ScriptPermission::ProcessExec.as_str().to_string());
        }
        if self.udf {
            requested.push(script::ScriptPermission::Udf.as_str().to_string());
        }
        if self.events_emit {
            requested.push(script::ScriptPermission::EventsEmit.as_str().to_string());
        }
        requested
    }

    fn as_script_permissions(&self) -> Vec<script::ScriptPermission> {
        let mut permissions = Vec::new();
        if self.fs_read {
            permissions.push(script::ScriptPermission::FsRead);
        }
        if self.fs_write {
            permissions.push(script::ScriptPermission::FsWrite);
        }
        if self.net_http {
            permissions.push(script::ScriptPermission::NetHttp);
        }
        if self.clipboard {
            permissions.push(script::ScriptPermission::Clipboard);
        }
        if self.process_exec {
            permissions.push(script::ScriptPermission::ProcessExec);
        }
        if self.udf {
            permissions.push(script::ScriptPermission::Udf);
        }
        if self.events_emit {
            permissions.push(script::ScriptPermission::EventsEmit);
        }
        permissions
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropMacroMutationPreview {
    sheet: String,
    cell: String,
    kind: String,
    value: Option<CellValue>,
    formula: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropRunMacroResponse {
    workbook_id: String,
    script_path: String,
    macro_name: String,
    script_fingerprint: String,
    trust: Option<InteropMacroTrustProvenance>,
    runtime_events: Vec<InteropScriptRuntimeEvent>,
    requested_permissions: Vec<String>,
    permission_events: Vec<InteropScriptPermissionEvent>,
    permission_granted: usize,
    permission_denied: usize,
    mutation_count: usize,
    changed_sheets: Vec<String>,
    mutations: Vec<InteropMacroMutationPreview>,
    recalc_reports: Vec<RecalcReport>,
    stdout: Option<String>,
    stderr: Option<String>,
    trace: TraceEcho,
}

#[derive(Debug)]
enum MacroMutationValue {
    Value { value: CellValue },
    Formula { formula: String },
}

#[derive(Debug)]
struct MacroMutationAssignment {
    sheet: String,
    cell_ref: String,
    row: u32,
    col: u32,
    kind: MacroMutationValue,
}

fn macro_script_cell_value_to_core_value(value: &script::ScriptCellValue) -> CellValue {
    match value {
        script::ScriptCellValue::Number(value) => CellValue::Number(*value),
        script::ScriptCellValue::Text(value) => CellValue::Text(value.clone()),
        script::ScriptCellValue::Bool(value) => CellValue::Bool(*value),
        script::ScriptCellValue::Error(error) => CellValue::Error(error.clone()),
        script::ScriptCellValue::Empty => CellValue::Empty,
    }
}

fn normalize_macro_formula(formula: &str) -> String {
    if formula.trim_start().starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    }
}

fn desugar_script_mutation(
    mutation: &script::ScriptMutation,
) -> Result<Vec<MacroMutationAssignment>, String> {
    match mutation {
        script::ScriptMutation::SetCellValue { sheet, cell, value } => {
            let (row, col) = parse_a1_cell_ref(cell)
                .ok_or_else(|| format!("invalid cell reference '{cell}' in script mutation"))?;
            Ok(vec![MacroMutationAssignment {
                sheet: sheet.to_string(),
                cell_ref: to_a1(row, col),
                row,
                col,
                kind: MacroMutationValue::Value {
                    value: macro_script_cell_value_to_core_value(value),
                },
            }])
        }
        script::ScriptMutation::SetCellFormula {
            sheet,
            cell,
            formula,
        } => {
            let (row, col) = parse_a1_cell_ref(cell)
                .ok_or_else(|| format!("invalid cell reference '{cell}' in script mutation"))?;
            Ok(vec![MacroMutationAssignment {
                sheet: sheet.to_string(),
                cell_ref: to_a1(row, col),
                row,
                col,
                kind: MacroMutationValue::Formula {
                    formula: normalize_macro_formula(formula),
                },
            }])
        }
        script::ScriptMutation::SetCellRangeValue {
            sheet,
            start,
            end,
            value,
        } => {
            let range = format!("{start}:{end}");
            let Some((start_row, start_col, end_row, end_col)) = parse_range_bounds(&range) else {
                return Err(format!(
                    "invalid range reference '{start}:{end}' in script mutation"
                ));
            };
            let mut output = Vec::new();
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    output.push(MacroMutationAssignment {
                        sheet: sheet.to_string(),
                        cell_ref: to_a1(row, col),
                        row,
                        col,
                        kind: MacroMutationValue::Value {
                            value: macro_script_cell_value_to_core_value(value),
                        },
                    });
                }
            }
            Ok(output)
        }
        script::ScriptMutation::SetCellRangeFormula {
            sheet,
            start,
            end,
            formula,
        } => {
            let range = format!("{start}:{end}");
            let Some((start_row, start_col, end_row, end_col)) = parse_range_bounds(&range) else {
                return Err(format!(
                    "invalid range reference '{start}:{end}' in script mutation"
                ));
            };
            let normalized = normalize_macro_formula(formula);
            let mut output = Vec::new();
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    output.push(MacroMutationAssignment {
                        sheet: sheet.to_string(),
                        cell_ref: to_a1(row, col),
                        row,
                        col,
                        kind: MacroMutationValue::Formula {
                            formula: normalized.clone(),
                        },
                    });
                }
            }
            Ok(output)
        }
    }
}

fn parse_macro_args(raw: &str) -> Result<BTreeMap<String, String>, String> {
    let mut parsed = BTreeMap::<String, String>::new();
    for raw_entry in raw
        .split(|c: char| matches!(c, '\n' | ';' | ','))
        .map(str::trim)
        .filter(|line| !line.is_empty())
    {
        let Some((name, value)) = raw_entry.split_once('=') else {
            return Err(format!(
                "invalid macro arg (expected key=value): {raw_entry}"
            ));
        };
        let key = name.trim();
        if key.is_empty() {
            return Err(format!("macro arg key cannot be empty: {raw_entry}"));
        }
        let value = value.trim();
        if parsed.insert(key.to_string(), value.to_string()).is_some() {
            return Err(format!("duplicate macro arg key: {key}"));
        }
    }
    Ok(parsed)
}

fn parse_range_bounds(range: &str) -> Option<(u32, u32, u32, u32)> {
    let Some((start, end)) = range.split_once(':') else {
        return None;
    };
    let (start_row, start_col) = parse_a1_cell_ref(start)?;
    let (end_row, end_col) = parse_a1_cell_ref(end)?;
    Some((
        start_row.min(end_row),
        start_col.min(end_col),
        start_row.max(end_row),
        start_col.max(end_col),
    ))
}

fn parse_macro_permission_events(
    permission_events: Vec<script::ScriptPermissionEvent>,
) -> Vec<InteropScriptPermissionEvent> {
    permission_events
        .into_iter()
        .map(|event| InteropScriptPermissionEvent {
            event_name: event.event_name,
            permission: event.permission,
            allowed: event.allowed,
            reason: event.reason,
        })
        .collect()
}

fn parse_macro_runtime_events(
    runtime_events: Vec<script::ScriptRuntimeEvent>,
) -> Vec<InteropScriptRuntimeEvent> {
    runtime_events
        .into_iter()
        .map(|event| InteropScriptRuntimeEvent {
            event_name: event.event_name,
            payload: event.payload,
            severity: event.severity,
        })
        .collect()
}

fn parse_macro_trust(
    trust: Option<script::ScriptTrustProvenance>,
) -> Option<InteropMacroTrustProvenance> {
    trust.map(|trust| InteropMacroTrustProvenance {
        mode: trust.mode,
        manifest_path: trust.manifest_path,
        manifest_name: trust.manifest_name,
        manifest_version: trust.manifest_version,
        publisher: trust.publisher,
        api_min_version: trust.api_min_version,
        permissions_required: trust.permissions_required,
        permissions_declared: trust.permissions_declared,
        runtime_api_version: trust.runtime_api_version,
        signature_present: trust.signature_present,
        signature_verified: trust.signature_verified,
        fingerprint: trust.fingerprint,
        trusted: trust.trusted,
    })
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
    undo_count: usize,
    redo_count: usize,
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

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct InteropUndoRedoResponse {
    action: String,
    workbook_id: String,
    dirty_sheet_count: usize,
    dirty_sheets: Vec<String>,
    undo_count: usize,
    redo_count: usize,
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
        Some(session) => {
            let (undo_count, redo_count) = session.history_depths();
            InteropSessionStatusResponse {
                loaded: true,
                input_path: Some(session.input_path.display().to_string()),
                workbook_id: Some(session.workbook.workbook_id.to_string()),
                sheet_count: session.workbook.sheets.len(),
                cell_count: workbook_cell_count(&session.workbook),
                issue_count: session.inspection.summary.issue_count,
                unknown_part_count: session.inspection.summary.unknown_part_count,
                dirty_sheet_count: session.dirty_sheets.len(),
                dirty_sheets: session.dirty_sheets.iter().cloned().collect(),
                undo_count,
                redo_count,
                sheets: workbook_sheet_names(&session.workbook),
            }
        }
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
            undo_count: 0,
            redo_count: 0,
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
    sink: &mut dyn EventSink,
    command_trace: &TraceContext,
) -> Result<InteropCellEditResponse, String> {
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
        .begin_txn(sink, command_trace)
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

    txn.commit(workbook, sink, command_trace)
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
    let mut sink = make_event_sink(trace.as_ref())?;
    let artifact_index_path = make_artifact_index_path(trace.as_ref());
    let event_log_path = resolve_event_log_path(trace.as_ref());
    let incoming_trace = trace
        .as_ref()
        .map(UiTraceContext::as_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let started = Instant::now();

    let mut workbook = Workbook::new();
    let mut txn = workbook
        .begin_txn(&mut *sink, &command_trace)
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

    txn.commit(&mut workbook, &mut *sink, &command_trace)
        .map_err(|error| error.to_string())?;

    let report = recalc_sheet(&mut workbook, "Sheet1", &mut *sink, &command_trace)
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

    let artifact_refs = vec![trace_artifact_ref(
        &command_trace,
        "engine_round_trip",
        "engine_roundtrip",
        "runtime_report",
        None,
        Some(format!("workbook_id={}", workbook.workbook_id)),
    )];

    let command_duration = started.elapsed().as_millis();
    let response = EngineRoundTripResponse {
        sheet: "Sheet1".to_string(),
        formula_cell: "C1".to_string(),
        value,
        evaluated_cells: report.evaluated_cells,
        cycle_count: report.cycle_count,
        parse_error_count: report.parse_error_count,
        workbook_id: workbook.workbook_id.to_string(),
        trace: finalize_command_trace(
            &command_trace,
            "success",
            command_duration,
            artifact_refs,
            artifact_index_path.as_deref(),
            event_log_path.as_deref(),
        )
        .map_err(|error| error.to_string())?,
    };

    Ok(response)
}

#[tauri::command]
fn interop_open_workbook(
    state: State<DesktopState>,
    path: String,
    trace: Option<UiTraceContext>,
) -> Result<InteropOpenResponse, String> {
    let input_path = normalize_input_path(&path)?;
    let mut sink = make_event_sink(trace.as_ref())?;
    let artifact_index_path = make_artifact_index_path(trace.as_ref());
    let event_log_path = resolve_event_log_path(trace.as_ref());
    let incoming_trace = trace
        .as_ref()
        .map(UiTraceContext::as_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let started = Instant::now();

    let inspection =
        inspect_xlsx(&input_path, &mut *sink, &command_trace).map_err(|error| error.to_string())?;
    let workbook = load_workbook_model(&input_path, &mut *sink, &command_trace)
        .map_err(|error| error.to_string())?;
    let workbook_sheet_count = workbook.sheets.len();
    let input_path_display = input_path.display().to_string();
    let mut response = map_open_response(&input_path, &workbook, &inspection, &command_trace);

    let mut session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    *session_guard = Some(InteropSession {
        input_path,
        workbook,
        inspection,
        dirty_sheets: BTreeSet::new(),
        undo_stack: Vec::new(),
        redo_stack: Vec::new(),
    });

    let artifact_refs = vec![
        trace_artifact_ref(
            &command_trace,
            "interop_open_workbook",
            "workbook_input",
            "input_source",
            Some(&input_path_display),
            Some("opened workbook artifact".to_string()),
        ),
        trace_artifact_ref(
            &command_trace,
            "interop_open_workbook",
            "compatibility_report",
            "open_report",
            Some(&input_path_display),
            Some(format!("sheets={}", workbook_sheet_count)),
        ),
    ];
    let command_duration = started.elapsed().as_millis();
    response.trace = finalize_command_trace(
        &command_trace,
        "success",
        command_duration,
        artifact_refs,
        artifact_index_path.as_deref(),
        event_log_path.as_deref(),
    )
    .map_err(|error| error.to_string())?;
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
    let mut sink = make_event_sink(trace.as_ref())?;
    let artifact_index_path = make_artifact_index_path(trace.as_ref());
    let event_log_path = resolve_event_log_path(trace.as_ref());
    let incoming_trace = trace
        .as_ref()
        .map(UiTraceContext::as_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let started = Instant::now();

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

    let target_sheet_list = target_sheets.join(",");
    let mut reports = Vec::with_capacity(target_sheets.len());
    for sheet_name in &target_sheets {
        let report = recalc_sheet(
            &mut session.workbook,
            sheet_name,
            &mut *sink,
            &command_trace,
        )
        .map_err(|error| error.to_string())?;
        session.dirty_sheets.insert(sheet_name.to_string());
        reports.push(report);
    }

    let artifact_refs = vec![trace_artifact_ref(
        &command_trace,
        "interop_recalc_loaded",
        "recalc_report",
        "sheet_recalc",
        Some(&target_sheet_list),
        Some(format!("dirty_sheets={}", session.dirty_sheets.len())),
    )];
    let command_duration = started.elapsed().as_millis();
    Ok(InteropRecalcResponse {
        workbook_id: session.workbook.workbook_id.to_string(),
        reports,
        trace: finalize_command_trace(
            &command_trace,
            "success",
            command_duration,
            artifact_refs,
            artifact_index_path.as_deref(),
            event_log_path.as_deref(),
        )
        .map_err(|error| error.to_string())?,
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
    let mut sink = make_event_sink(trace.as_ref())?;
    let artifact_index_path = make_artifact_index_path(trace.as_ref());
    let event_log_path = resolve_event_log_path(trace.as_ref());
    let incoming_trace = trace
        .as_ref()
        .map(UiTraceContext::as_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let started = Instant::now();
    let resolved_sheet = sheet.trim().to_string();
    let normalized_cell = cell.trim().to_ascii_uppercase();

    let mut session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    let session = session_guard
        .as_mut()
        .ok_or_else(|| "no workbook loaded; open a workbook first".to_string())?;
    session.push_undo_snapshot();

    let mut response = apply_cell_edit_to_workbook(
        &mut session.workbook,
        &mut session.dirty_sheets,
        &resolved_sheet,
        &normalized_cell,
        &input,
        mode,
        &mut *sink,
        &command_trace,
    )?;

    let artifact_refs = vec![trace_artifact_ref(
        &command_trace,
        "interop_apply_cell_edit",
        "workbook_mutation",
        "sheet_edit",
        Some(&format!("{resolved_sheet}!{normalized_cell}")),
        Some(format!("applied_cells={}", response.applied_cell_count)),
    )];
    let command_duration = started.elapsed().as_millis();
    response.trace = finalize_command_trace(
        &command_trace,
        "success",
        command_duration,
        artifact_refs,
        artifact_index_path.as_deref(),
        event_log_path.as_deref(),
    )
    .map_err(|error| error.to_string())?;

    Ok(response)
}

#[tauri::command]
fn interop_run_macro(
    state: State<DesktopState>,
    script_path: String,
    macro_name: String,
    args: String,
    permissions: InteropMacroPermissionConfig,
    trace: Option<UiTraceContext>,
) -> Result<InteropRunMacroResponse, String> {
    let script_path = script_path.trim();
    if script_path.is_empty() {
        return Err("macro script path is required".to_string());
    }
    let macro_name = macro_name.trim();
    if macro_name.is_empty() {
        return Err("macro name is required".to_string());
    }

    let parsed_args = parse_macro_args(&args)?;
    let requested_permissions = permissions.as_requested_permissions();
    let requested_script_permissions = permissions.as_script_permissions();

    let mut sink = make_event_sink(trace.as_ref())?;
    let artifact_index_path = make_artifact_index_path(trace.as_ref());
    let event_log_path = resolve_event_log_path(trace.as_ref());
    let incoming_trace = trace
        .as_ref()
        .map(UiTraceContext::as_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let started = Instant::now();

    let mut session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    let session = session_guard
        .as_mut()
        .ok_or_else(|| "no workbook loaded; open a workbook first".to_string())?;

    let request = script::MacroRunRequest::new(
        command_trace.trace_id.to_string(),
        Path::new(script_path),
        macro_name.to_string(),
        &session.input_path,
        requested_script_permissions,
        parsed_args,
    );

    let invocation_started = Instant::now();
    sink.emit(
        EventEnvelope::info("script.session.start", &command_trace)
            .with_context(json!({
                "macro_script": script_path,
                "macro_name": macro_name,
                "input": session.input_path.display().to_string(),
            }))
            .with_payload(json!({
                "status": "started",
                "requested_permissions": requested_permissions,
            }))
            .with_metrics(json!({
                "requested_permission_count": request.permissions.len(),
                "macro_arg_count": request.args.len(),
            })),
    )
    .map_err(|error| error.to_string())?;

    let response = match script::run_macro(&request) {
        Ok(response) => response,
        Err(error) => {
            sink.emit(
                EventEnvelope::info("script.rpc.error", &command_trace)
                    .with_context(json!({
                        "operation": "script.macro.run",
                        "macro_script": script_path,
                        "macro_name": macro_name,
                        "input": session.input_path.display().to_string(),
                    }))
                    .with_payload(json!({
                        "status": "error",
                        "message": error.to_string(),
                    }))
                    .with_metrics(json!({
                        "duration_ms": invocation_started.elapsed().as_secs_f64() * 1000.0,
                    })),
            )
            .map_err(|event_error| format!("{error} ({event_error})"))?;

            return Err(error.to_string());
        }
    };

    let script_fingerprint = response
        .script_fingerprint
        .clone()
        .unwrap_or_else(|| "unavailable".to_string());
    let parsed_permissions = parse_macro_permission_events(response.permission_events.clone());
    let parsed_runtime_events = parse_macro_runtime_events(response.runtime_events.clone());
    let parsed_trust = parse_macro_trust(response.trust.clone());
    let permission_granted = response
        .permission_events
        .iter()
        .filter(|event| event.allowed)
        .count();
    let permission_denied = response
        .permission_events
        .len()
        .saturating_sub(permission_granted);

    for event in response.permission_events.iter() {
        let event_name = if event.allowed {
            "script.permission.granted"
        } else {
            "script.permission.denied"
        };
        sink.emit(
            EventEnvelope::info(event_name, &command_trace)
                .with_context(json!({
                    "permission": event.permission,
                    "allowed": event.allowed,
                    "macro_script": script_path,
                    "macro_name": macro_name,
                    "macro_script_fingerprint": script_fingerprint,
                    "reason": event.reason,
                }))
                .with_payload(json!({
                "status": if event.allowed { "granted" } else { "denied" },
                })),
        )
        .map_err(|error| error.to_string())?;
    }

    if let Some(trust) = parsed_trust.as_ref() {
        sink.emit(
            EventEnvelope::info("script.trust", &command_trace)
                .with_context(json!({
                    "macro_script": script_path,
                    "macro_name": macro_name,
                    "macro_script_fingerprint": script_fingerprint,
                }))
                .with_payload(json!({
                    "trust": trust,
                }))
                .with_metrics(json!({
                    "required_permissions": trust.permissions_required.len(),
                    "declared_permissions": trust.permissions_declared.len(),
                    "runtime_api_version": trust.runtime_api_version,
                    "signature_present": trust.signature_present,
                    "trusted": trust.trusted,
                })),
        )
        .map_err(|error| error.to_string())?;
    }

    if !parsed_runtime_events.is_empty() {
        sink.emit(
            EventEnvelope::info("script.runtime.events", &command_trace)
                .with_context(json!({
                    "macro_script": script_path,
                    "macro_name": macro_name,
                    "macro_script_fingerprint": script_fingerprint,
                }))
                .with_payload(json!({
                    "runtime_events": parsed_runtime_events,
                }))
                .with_metrics(json!({
                    "runtime_event_count": parsed_runtime_events.len(),
                })),
        )
        .map_err(|error| error.to_string())?;
    }

    if response.status.to_lowercase() != "ok" {
        return Err(format!(
            "macro execution failed: {}",
            response
                .message
                .clone()
                .unwrap_or_else(|| "non-ok script status".to_string())
        ));
    }

    let mut assignments = Vec::<MacroMutationAssignment>::new();
    for mutation in &response.mutations {
        assignments.extend(desugar_script_mutation(mutation)?);
    }

    let mutations = assignments
        .iter()
        .map(|assignment| InteropMacroMutationPreview {
            sheet: assignment.sheet.clone(),
            cell: assignment.cell_ref.clone(),
            kind: match assignment.kind {
                MacroMutationValue::Value { .. } => "value".to_string(),
                MacroMutationValue::Formula { .. } => "formula".to_string(),
            },
            value: match &assignment.kind {
                MacroMutationValue::Value { value } => Some(value.clone()),
                MacroMutationValue::Formula { .. } => None,
            },
            formula: match &assignment.kind {
                MacroMutationValue::Formula { formula } => Some(formula.clone()),
                MacroMutationValue::Value { .. } => None,
            },
        })
        .collect::<Vec<_>>();

    let mut recalc_reports = Vec::<RecalcReport>::new();
    let changed_sheets = if !assignments.is_empty() {
        session.push_undo_snapshot();

        let mut txn = session
            .workbook
            .begin_txn(&mut *sink, &command_trace)
            .map_err(|error| error.to_string())?;

        for assignment in assignments.iter() {
            match &assignment.kind {
                MacroMutationValue::Value { value } => txn.apply(Mutation::SetCellValue {
                    sheet: assignment.sheet.clone(),
                    row: assignment.row,
                    col: assignment.col,
                    value: value.clone(),
                }),
                MacroMutationValue::Formula { formula } => txn.apply(Mutation::SetCellFormula {
                    sheet: assignment.sheet.clone(),
                    row: assignment.row,
                    col: assignment.col,
                    formula: formula.clone(),
                    cached_value: CellValue::Empty,
                }),
            }
        }

        let commit = txn
            .commit(&mut session.workbook, &mut *sink, &command_trace)
            .map_err(|error| error.to_string())?;

        let mut sheet_list = commit
            .changed_cells
            .keys()
            .map(|name| name.to_string())
            .collect::<Vec<_>>();
        sheet_list.sort_unstable();

        for sheet_name in &sheet_list {
            session.dirty_sheets.insert(sheet_name.to_string());
            let changed_roots = commit
                .changed_cells
                .get(sheet_name)
                .map(|cells| cells.as_slice())
                .unwrap_or(&[]);
            let report = recalc_sheet_from_roots(
                &mut session.workbook,
                sheet_name,
                changed_roots,
                &mut *sink,
                &command_trace,
            )
            .map_err(|error| error.to_string())?;
            recalc_reports.push(report);
        }

        sheet_list
    } else {
        Vec::new()
    };

    let command_duration = started.elapsed().as_millis();
    let mut artifact_refs = vec![trace_artifact_ref(
        &command_trace,
        "interop_run_macro",
        "script_session",
        "macro_request",
        Some(script_path),
        Some(format!("macro_name={macro_name}")),
    )];
    if !changed_sheets.is_empty() {
        artifact_refs.push(trace_artifact_ref(
            &command_trace,
            "interop_run_macro",
            "workbook_mutation",
            "macro_cells",
            Some(&changed_sheets.join(",")),
            Some(format!("mutation_count={}", assignments.len())),
        ));
    }
    if !recalc_reports.is_empty() {
        artifact_refs.push(trace_artifact_ref(
            &command_trace,
            "interop_run_macro",
            "recalc_report",
            "macro_recalc",
            Some(&changed_sheets.join(",")),
            Some(format!("reports={}", recalc_reports.len())),
        ));
    }
    sink.emit(
        EventEnvelope::info("script.macro.run", &command_trace)
            .with_context(json!({
                "macro_script": script_path,
                "macro_name": macro_name,
                "input": session.input_path.display().to_string(),
                "status": "ok",
                "macro_script_fingerprint": script_fingerprint,
                "result": response.result,
                "stdout": response.stdout,
                "stderr": response.stderr,
            }))
            .with_metrics(json!({
                "macro_invocation_ms": invocation_started.elapsed().as_secs_f64() * 1000.0,
                "macro_mutation_count": assignments.len(),
                "changed_sheet_count": changed_sheets.len(),
                "permission_granted": permission_granted,
                "permission_denied": permission_denied,
                "permission_event_count": response.permission_events.len(),
                "runtime_event_count": parsed_runtime_events.len(),
                "trust_present": parsed_trust.is_some(),
            })),
    )
    .map_err(|error| error.to_string())?;

    sink.emit(
        EventEnvelope::info("script.session.end", &command_trace)
            .with_context(json!({
                "macro_script": script_path,
                "macro_name": macro_name,
                "status": "ok",
                "macro_script_fingerprint": script_fingerprint,
            }))
            .with_metrics(json!({
                "mutations": assignments.len(),
                "changed_sheet_count": changed_sheets.len(),
                "macro_invocation_ms": invocation_started.elapsed().as_secs_f64() * 1000.0,
            })),
    )
    .map_err(|error| error.to_string())?;

    Ok(InteropRunMacroResponse {
        workbook_id: session.workbook.workbook_id.to_string(),
        script_path: script_path.to_string(),
        macro_name: macro_name.to_string(),
        script_fingerprint,
        requested_permissions,
        permission_events: parsed_permissions,
        trust: parsed_trust,
        runtime_events: parsed_runtime_events,
        permission_granted,
        permission_denied,
        mutation_count: assignments.len(),
        changed_sheets,
        mutations,
        recalc_reports,
        stdout: response.stdout,
        stderr: response.stderr,
        trace: finalize_command_trace(
            &command_trace,
            "success",
            command_duration,
            artifact_refs,
            artifact_index_path.as_deref(),
            event_log_path.as_deref(),
        )
        .map_err(|error| error.to_string())?,
    })
}

fn make_undo_redo_response(
    session: &InteropSession,
    action: &str,
    command_trace: &TraceContext,
    command_duration: u128,
    artifact_index_path: Option<&Path>,
    event_log_path: Option<&str>,
) -> Result<InteropUndoRedoResponse, String> {
    let (undo_count, redo_count) = session.history_depths();
    let artifact_refs = vec![trace_artifact_ref(
        command_trace,
        &format!("interop_{action}"),
        "workbook_snapshot",
        "history_restore",
        Some(&session.input_path.display().to_string()),
        Some(format!("{action} history restore")),
    )];
    Ok(InteropUndoRedoResponse {
        action: action.to_string(),
        workbook_id: session.workbook.workbook_id.to_string(),
        dirty_sheet_count: session.dirty_sheets.len(),
        dirty_sheets: session.dirty_sheets.iter().cloned().collect(),
        undo_count,
        redo_count,
        trace: finalize_command_trace(
            command_trace,
            "success",
            command_duration,
            artifact_refs,
            artifact_index_path,
            event_log_path,
        )
        .map_err(|error| error.to_string())?,
    })
}

#[tauri::command]
fn interop_undo_edit(
    state: State<DesktopState>,
    trace: Option<UiTraceContext>,
) -> Result<InteropUndoRedoResponse, String> {
    let _sink = make_event_sink(trace.as_ref())?;
    let artifact_index_path = make_artifact_index_path(trace.as_ref());
    let event_log_path = resolve_event_log_path(trace.as_ref());
    let incoming_trace = trace
        .as_ref()
        .map(UiTraceContext::as_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let started = Instant::now();

    let _ = _sink;
    let mut session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    let session = session_guard
        .as_mut()
        .ok_or_else(|| "no workbook loaded; open a workbook first".to_string())?;

    session.undo()?;
    let response = make_undo_redo_response(
        session,
        "undo",
        &command_trace,
        started.elapsed().as_millis(),
        artifact_index_path.as_deref(),
        event_log_path.as_deref(),
    )?;
    Ok(response)
}

#[tauri::command]
fn interop_redo_edit(
    state: State<DesktopState>,
    trace: Option<UiTraceContext>,
) -> Result<InteropUndoRedoResponse, String> {
    let _sink = make_event_sink(trace.as_ref())?;
    let artifact_index_path = make_artifact_index_path(trace.as_ref());
    let event_log_path = resolve_event_log_path(trace.as_ref());
    let incoming_trace = trace
        .as_ref()
        .map(UiTraceContext::as_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let started = Instant::now();

    let _ = _sink;
    let mut session_guard = state
        .session
        .lock()
        .map_err(|_| "desktop state lock poisoned".to_string())?;
    let session = session_guard
        .as_mut()
        .ok_or_else(|| "no workbook loaded; open a workbook first".to_string())?;

    session.redo()?;
    let response = make_undo_redo_response(
        session,
        "redo",
        &command_trace,
        started.elapsed().as_millis(),
        artifact_index_path.as_deref(),
        event_log_path.as_deref(),
    )?;
    Ok(response)
}

#[tauri::command]
fn interop_sheet_preview(
    state: State<DesktopState>,
    sheet: Option<String>,
    limit: Option<usize>,
    trace: Option<UiTraceContext>,
) -> Result<InteropSheetPreviewResponse, String> {
    let artifact_index_path = make_artifact_index_path(trace.as_ref());
    let event_log_path = resolve_event_log_path(trace.as_ref());
    let incoming_trace = trace
        .as_ref()
        .map(UiTraceContext::as_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let started = Instant::now();

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

    let artifact_refs = vec![trace_artifact_ref(
        &command_trace,
        "interop_sheet_preview",
        "sheet_snapshot",
        "sheet_preview",
        Some(&resolved_sheet),
        Some(format!("limit={preview_limit},shown={}", cells.len())),
    )];
    let command_duration = started.elapsed().as_millis();
    Ok(InteropSheetPreviewResponse {
        workbook_id: session.workbook.workbook_id.to_string(),
        sheet: resolved_sheet,
        total_cells: sheet_data.cells.len(),
        shown_cells: cells.len(),
        truncated: sheet_data.cells.len() > cells.len(),
        cells,
        trace: finalize_command_trace(
            &command_trace,
            "success",
            command_duration,
            artifact_refs,
            artifact_index_path.as_deref(),
            event_log_path.as_deref(),
        )
        .map_err(|error| error.to_string())?,
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

    let mut sink = make_event_sink(trace.as_ref())?;
    let artifact_index_path = make_artifact_index_path(trace.as_ref());
    let event_log_path = resolve_event_log_path(trace.as_ref());
    let incoming_trace = trace
        .as_ref()
        .map(UiTraceContext::as_trace_context)
        .unwrap_or_else(TraceContext::root);
    let command_trace = incoming_trace.child();
    let started = Instant::now();

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
            &mut *sink,
            &command_trace,
        ),
        UiSaveMode::Preserve => {
            let changed_sheets = session.dirty_sheets.iter().cloned().collect::<Vec<_>>();
            preserve_xlsx_with_sheet_overrides(
                &session.input_path,
                output_path,
                &session.workbook,
                &changed_sheets,
                &mut *sink,
                &command_trace,
            )
        }
        UiSaveMode::Normalize => save_workbook_model(
            &session.workbook,
            output_path,
            SaveMode::Normalize,
            &mut *sink,
            &command_trace,
        ),
    }
    .map_err(|error| error.to_string())?;
    let saved_output_path = report.output_path.clone();
    let artifact_refs = vec![
        trace_artifact_ref(
            &command_trace,
            "interop_save_workbook",
            "workbook_output",
            "saved_workbook",
            Some(&saved_output_path.display().to_string()),
            Some(format!("mode={}", mode.as_str())),
        ),
        trace_artifact_ref(
            &command_trace,
            "interop_save_workbook",
            "workbook_snapshot",
            "session_input",
            Some(&session.input_path.display().to_string()),
            Some("source_workbook_path".to_string()),
        ),
    ];
    let command_duration = started.elapsed().as_millis();
    let mut response = map_save_response(session, report, mode, &command_trace);
    response.trace = finalize_command_trace(
        &command_trace,
        "success",
        command_duration,
        artifact_refs,
        artifact_index_path.as_deref(),
        event_log_path.as_deref(),
    )
    .map_err(|error| error.to_string())?;

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
    use std::sync::OnceLock;

    static TEST_ENV_MUTEX: OnceLock<Mutex<()>> = OnceLock::new();

    fn make_test_dir(test_name: &str) -> PathBuf {
        let suffix = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|clock| clock.as_nanos())
            .unwrap_or(0);
        let path =
            std::env::temp_dir().join(format!("rootcellar-desktop-macro-{test_name}-{suffix}"));
        std::fs::create_dir_all(&path).expect("create temporary test directory");
        path
    }

    fn python_available() -> bool {
        std::process::Command::new("python")
            .arg("-V")
            .output()
            .is_ok()
    }

    fn with_env_vars<T, F>(entries: &[(&str, Option<&str>)], action: F) -> T
    where
        F: FnOnce() -> T,
    {
        let _guard = TEST_ENV_MUTEX
            .get_or_init(|| Mutex::new(()))
            .lock()
            .expect("env lock must be acquireable");

        let mut originals: Vec<(String, Option<String>)> = Vec::new();
        for (key, value) in entries {
            originals.push((key.to_string(), std::env::var(key).ok()));
            match value {
                Some(value) => std::env::set_var(key, value),
                None => std::env::remove_var(key),
            }
        }

        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(action));

        for (key, original) in originals {
            match original {
                Some(value) => std::env::set_var(&key, value),
                None => std::env::remove_var(&key),
            }
        }

        match result {
            Ok(value) => value,
            Err(payload) => std::panic::resume_unwind(payload),
        }
    }

    fn temp_workbook_with_macro_inputs(test_name: &str) -> PathBuf {
        let test_dir = make_test_dir(test_name);
        let input_path = test_dir.join("input.xlsx");
        std::fs::copy(sample_xlsx_path(), &input_path).expect("copy fixture workbook");
        input_path
    }

    fn write_temp_script(path: &Path, body: &str) {
        std::fs::write(path, body).expect("write temporary script");
    }

    fn read_cell_value(session: &DesktopState, row: u32, col: u32) -> CellValue {
        let session_guard = session
            .session
            .lock()
            .expect("desktop state lock should not be poisoned");
        let opened = session_guard
            .as_ref()
            .expect("workbook should be opened for session");
        opened
            .workbook
            .sheets
            .get("Sheet1")
            .and_then(|sheet_data| sheet_data.cells.get(&CellRef { row, col }))
            .map(|record| record.value.clone())
            .expect("cell should exist in loaded workbook")
    }

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

    fn sample_xlsx_path() -> String {
        let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        manifest_dir
            .join("..")
            .join("..")
            .join("..")
            .join("sample-formula.xlsx")
            .to_string_lossy()
            .into_owned()
    }

    fn make_command_trace(
        seed_trace_id: &str,
        seed_session_id: &str,
        command_name: &str,
        event_log_path: Option<&str>,
        artifact_index_path: Option<&str>,
    ) -> UiTraceContext {
        UiTraceContext {
            trace_id: Some(seed_trace_id.to_string()),
            span_id: Some(Uuid::now_v7().to_string()),
            parent_span_id: None,
            session_id: Some(seed_session_id.to_string()),
            command_id: Some(Uuid::now_v7().to_string()),
            command_name: Some(command_name.to_string()),
            artifact_index_path: artifact_index_path.map(ToString::to_string),
            event_log_path: event_log_path.map(ToString::to_string),
        }
    }

    fn event_trace_present_in_log(path: &Path, trace_id: &str) -> bool {
        let expected = format!("\"trace_id\":\"{trace_id}\"");
        std::fs::read_to_string(path)
            .map(|contents| contents.contains(&expected))
            .unwrap_or(false)
    }

    fn event_line_count(path: &Path) -> usize {
        std::fs::read_to_string(path)
            .map(|contents| {
                contents
                    .lines()
                    .filter(|line| !line.trim().is_empty())
                    .count()
            })
            .unwrap_or(0)
    }

    fn artifact_index_present_in_log(path: &Path, trace_id: &str, command_name: &str) -> bool {
        std::fs::read_to_string(path)
            .map(|contents| {
                let trace_marker = format!("\"traceId\":\"{trace_id}\"");
                let command_marker = format!("\"commandName\":\"{command_name}\"");
                contents.contains(&trace_marker) && contents.contains(&command_marker)
            })
            .unwrap_or(false)
    }

    fn artifact_index_line_count(path: &Path) -> usize {
        std::fs::read_to_string(path)
            .map(|contents| {
                contents
                    .lines()
                    .filter(|line| !line.trim().is_empty())
                    .count()
            })
            .unwrap_or(0)
    }

    #[allow(dead_code)]
    #[derive(Debug, Clone, Deserialize)]
    #[serde(rename_all = "camelCase")]
    struct ArtifactIndexRecord {
        command_name: String,
        command_status: String,
        duration_ms: u128,
        trace_id: String,
        trace_root_id: String,
        span_id: String,
        parent_span_id: Option<String>,
        session_id: Option<String>,
        ui_command_id: Option<String>,
        ui_command_name: Option<String>,
        event_log_path: Option<String>,
        linked_artifact_ids: Vec<String>,
        artifact_refs: Vec<TraceArtifactRef>,
    }

    fn load_artifact_index_records(path: &Path) -> Vec<ArtifactIndexRecord> {
        std::fs::read_to_string(path)
            .unwrap_or_default()
            .lines()
            .filter_map(|line| {
                let trimmed = line.trim();
                if trimmed.is_empty() {
                    return None;
                }
                serde_json::from_str::<ArtifactIndexRecord>(trimmed).ok()
            })
            .collect()
    }

    fn assert_trace_has_manifest_coverage(trace: &TraceEcho, records: &[ArtifactIndexRecord]) {
        let command_name = trace
            .ui_command_name
            .clone()
            .unwrap_or_else(|| "unknown".to_string());
        let matching = records
            .iter()
            .filter(|record| {
                record.trace_id == trace.trace_id && record.command_name == command_name
            })
            .collect::<Vec<_>>();
        assert!(
            !matching.is_empty(),
            "artifact index should include {command_name} record for trace_id {}",
            trace.trace_id
        );

        for artifact_id in &trace.linked_artifact_ids {
            let resolved = matching.iter().any(|record| {
                record
                    .linked_artifact_ids
                    .iter()
                    .any(|id| id == artifact_id)
            });
            assert!(
                resolved,
                "artifact index for {command_name} should include linked_artifact_id {artifact_id}"
            );
        }

        for artifact_ref in &trace.artifact_refs {
            assert!(
                trace
                    .linked_artifact_ids
                    .iter()
                    .any(|artifact_id| artifact_id == &artifact_ref.artifact_id),
                "artifact_ref id {} should be present in linked_artifact_ids for {command_name}",
                artifact_ref.artifact_id
            );
            let present = matching.iter().any(|record| {
                record.artifact_refs.iter().any(|candidate| {
                    candidate.artifact_id == artifact_ref.artifact_id
                        && candidate.artifact_type == artifact_ref.artifact_type
                        && candidate.relation == artifact_ref.relation
                })
            });
            assert!(
                present,
                "artifact index for {command_name} should include artifact_ref {}",
                artifact_ref.artifact_id
            );
        }

        for record in &matching {
            assert_eq!(
                record.trace_root_id, trace.trace_root_id,
                "artifact index trace_root_id should match trace payload"
            );
            assert_eq!(
                record.trace_id, trace.trace_id,
                "artifact index trace_id should match command trace payload"
            );
            assert_eq!(
                record.command_status, trace.command_status,
                "artifact index command_status should match command trace payload"
            );
        }
    }

    fn as_state<'a>(session: &'a DesktopState) -> tauri::State<'a, DesktopState> {
        unsafe {
            // SAFETY: `State` is a transparent wrapper around a reference for read-only access.
            std::mem::transmute::<&'a DesktopState, tauri::State<'a, DesktopState>>(session)
        }
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

    fn cell_value_from_session(
        session: &DesktopState,
        sheet: &str,
        row: u32,
        col: u32,
    ) -> CellValue {
        let session_guard = session
            .session
            .lock()
            .expect("desktop state lock should not be poisoned");
        let opened = session_guard
            .as_ref()
            .expect("workbook should be opened for session");
        opened
            .workbook
            .sheets
            .get(sheet)
            .and_then(|sheet_data| sheet_data.cells.get(&CellRef { row, col }))
            .map(|record| record.value.clone())
            .expect("cell should exist in loaded workbook")
    }

    fn sheet1_cell_from_session(session: &DesktopState, row: u32, col: u32) -> CellValue {
        cell_value_from_session(session, "Sheet1", row, col)
    }

    fn sheet1_formula_from_session(session: &DesktopState, row: u32, col: u32) -> Option<String> {
        let session_guard = session
            .session
            .lock()
            .expect("desktop state lock should not be poisoned");
        let opened = session_guard
            .as_ref()
            .expect("workbook should be opened for session");
        opened
            .workbook
            .sheets
            .get("Sheet1")
            .and_then(|sheet| sheet.cells.get(&CellRef { row, col }))
            .and_then(|record| record.formula.clone())
    }

    #[test]
    fn desktop_trace_continuity_smoke_open_edit_save_recalc() {
        let sample_path = sample_xlsx_path();
        let session = DesktopState::default();
        let state = as_state(&session);

        let command_root = TraceContext::root();
        let trace_root = command_root.trace_id.to_string();
        let trace_session = command_root
            .session_id
            .expect("root session id should be set")
            .to_string();
        let event_log_path = std::env::temp_dir()
            .join(format!("rootcellar-desktop-events-{trace_root}.jsonl"))
            .to_string_lossy()
            .into_owned();
        let artifact_index_path = std::env::temp_dir()
            .join(format!("rootcellar-desktop-artifacts-{trace_root}.jsonl"))
            .to_string_lossy()
            .into_owned();

        let open_trace_context = make_command_trace(
            &trace_root,
            &trace_session,
            "interop_open_workbook",
            Some(&event_log_path),
            Some(&artifact_index_path),
        );
        let open_payload = interop_open_workbook(
            state.clone(),
            sample_path.clone(),
            Some(open_trace_context.clone()),
        )
        .expect("open should succeed for sample workbook");
        assert_eq!(open_payload.trace.command_status, "success");
        assert!(
            event_line_count(Path::new(&event_log_path)) > 0,
            "event log should be written for open"
        );
        assert!(
            event_trace_present_in_log(Path::new(&event_log_path), &trace_root),
            "events should include open trace_id"
        );
        assert!(
            artifact_index_present_in_log(
                Path::new(&artifact_index_path),
                &trace_root,
                "interop_open_workbook"
            ),
            "artifact index should include open command"
        );
        assert!(!open_payload.trace.linked_artifact_ids.is_empty());
        assert_trace_has_manifest_coverage(
            &open_payload.trace,
            &load_artifact_index_records(Path::new(&artifact_index_path)),
        );

        let first_sheet = open_payload
            .sheets
            .first()
            .cloned()
            .expect("sample workbook should expose at least one sheet");
        assert_eq!(open_payload.trace.trace_id, trace_root);
        assert_eq!(
            open_payload.trace.trace_root_id,
            open_payload.trace.trace_id
        );
        assert_eq!(
            open_payload.trace.parent_span_id,
            Some(open_trace_context.span_id.expect("open trace span"))
        );
        assert!(open_payload.trace.ui_command_id.is_some());
        assert_eq!(
            open_payload.trace.ui_command_name.as_deref(),
            Some("interop_open_workbook")
        );
        let events_after_open = event_line_count(Path::new(&event_log_path));

        let preview_trace_context = make_command_trace(
            &trace_root,
            &trace_session,
            "interop_sheet_preview",
            Some(&event_log_path),
            Some(&artifact_index_path),
        );
        let preview_payload = interop_sheet_preview(
            state.clone(),
            Some(first_sheet.clone()),
            Some(20),
            Some(preview_trace_context.clone()),
        )
        .expect("preview should succeed after open");
        assert_eq!(preview_payload.trace.trace_id, trace_root);
        assert_eq!(
            preview_payload.trace.trace_root_id,
            preview_payload.trace.trace_id
        );
        assert_eq!(preview_payload.trace.command_status, "success");
        assert!(!preview_payload.trace.linked_artifact_ids.is_empty());
        assert_trace_has_manifest_coverage(
            &preview_payload.trace,
            &load_artifact_index_records(Path::new(&artifact_index_path)),
        );
        assert_eq!(
            preview_payload.trace.parent_span_id,
            Some(preview_trace_context.span_id.expect("preview trace span"))
        );
        assert_eq!(
            preview_payload.trace.ui_command_name.as_deref(),
            Some("interop_sheet_preview")
        );
        assert!(
            artifact_index_present_in_log(
                Path::new(&artifact_index_path),
                &trace_root,
                "interop_sheet_preview",
            ),
            "artifact index should include preview command"
        );
        assert!(
            artifact_index_line_count(Path::new(&artifact_index_path)) >= 2,
            "artifact index should include open + preview records"
        );

        let edit_trace_context = make_command_trace(
            &trace_root,
            &trace_session,
            "interop_apply_cell_edit",
            Some(&event_log_path),
            Some(&artifact_index_path),
        );
        let edit_payload = interop_apply_cell_edit(
            state.clone(),
            first_sheet.clone(),
            "A1".to_string(),
            "11".to_string(),
            UiEditMode::Value,
            Some(edit_trace_context.clone()),
        )
        .expect("edit should succeed on sample workbook");
        assert_eq!(edit_payload.trace.trace_id, trace_root);
        assert_eq!(
            edit_payload.trace.trace_root_id,
            edit_payload.trace.trace_id
        );
        assert_eq!(edit_payload.trace.command_status, "success");
        assert!(!edit_payload.trace.linked_artifact_ids.is_empty());
        assert_trace_has_manifest_coverage(
            &edit_payload.trace,
            &load_artifact_index_records(Path::new(&artifact_index_path)),
        );
        assert_eq!(
            edit_payload.trace.parent_span_id,
            Some(edit_trace_context.span_id.expect("edit trace span"))
        );
        assert_eq!(
            edit_payload.trace.ui_command_name.as_deref(),
            Some("interop_apply_cell_edit")
        );
        let events_after_edit = event_line_count(Path::new(&event_log_path));
        assert!(
            events_after_edit > events_after_open,
            "event log should grow after edit"
        );
        assert!(
            event_trace_present_in_log(Path::new(&event_log_path), &trace_root),
            "events should include edit trace_id"
        );
        assert!(
            artifact_index_present_in_log(
                Path::new(&artifact_index_path),
                &trace_root,
                "interop_apply_cell_edit",
            ),
            "artifact index should include edit command"
        );
        assert!(
            artifact_index_line_count(Path::new(&artifact_index_path)) >= 2,
            "artifact index should include at least open + edit records"
        );

        let output_path = std::env::temp_dir()
            .join(format!("rootcellar-desktop-smoke-{trace_root}.xlsx"))
            .to_string_lossy()
            .into_owned();
        let save_trace_context = make_command_trace(
            &trace_root,
            &trace_session,
            "interop_save_workbook",
            Some(&event_log_path),
            Some(&artifact_index_path),
        );
        let save_payload = interop_save_workbook(
            state.clone(),
            output_path.clone(),
            UiSaveMode::Preserve,
            Some(false),
            Some(save_trace_context.clone()),
        )
        .expect("save should succeed after mutation");
        assert_eq!(save_payload.trace.command_status, "success");
        assert!(!save_payload.trace.linked_artifact_ids.is_empty());
        assert_trace_has_manifest_coverage(
            &save_payload.trace,
            &load_artifact_index_records(Path::new(&artifact_index_path)),
        );
        assert_eq!(save_payload.trace.trace_id, trace_root);
        assert_eq!(
            save_payload.trace.trace_root_id,
            save_payload.trace.trace_id
        );
        assert_eq!(
            save_payload.trace.parent_span_id,
            Some(save_trace_context.span_id.expect("save trace span"))
        );
        assert_eq!(
            save_payload.trace.ui_command_name.as_deref(),
            Some("interop_save_workbook")
        );
        let events_after_save = event_line_count(Path::new(&event_log_path));
        assert!(
            events_after_save > events_after_edit,
            "event log should grow after save"
        );
        assert!(
            event_trace_present_in_log(Path::new(&event_log_path), &trace_root),
            "events should include save trace_id"
        );
        assert!(
            artifact_index_present_in_log(
                Path::new(&artifact_index_path),
                &trace_root,
                "interop_save_workbook",
            ),
            "artifact index should include save command"
        );
        assert!(
            artifact_index_line_count(Path::new(&artifact_index_path)) >= 4,
            "artifact index should include open + preview + edit + save records"
        );

        let recalc_trace_context = make_command_trace(
            &trace_root,
            &trace_session,
            "interop_recalc_loaded",
            Some(&event_log_path),
            Some(&artifact_index_path),
        );
        let recalc_payload =
            interop_recalc_loaded(state, Some(first_sheet), Some(recalc_trace_context.clone()))
                .expect("recalc should succeed for loaded workbook");
        assert_eq!(recalc_payload.trace.command_status, "success");
        assert!(!recalc_payload.trace.linked_artifact_ids.is_empty());
        assert_trace_has_manifest_coverage(
            &recalc_payload.trace,
            &load_artifact_index_records(Path::new(&artifact_index_path)),
        );
        assert_eq!(recalc_payload.trace.trace_id, trace_root);
        let events_after_recalc = event_line_count(Path::new(&event_log_path));
        assert_eq!(
            recalc_payload.trace.trace_root_id,
            recalc_payload.trace.trace_id
        );
        assert_eq!(
            recalc_payload.trace.parent_span_id,
            Some(recalc_trace_context.span_id.expect("recalc trace span"))
        );
        assert!(!recalc_payload.reports.is_empty());
        assert_eq!(
            recalc_payload.trace.ui_command_name.as_deref(),
            Some("interop_recalc_loaded")
        );
        assert!(
            event_line_count(Path::new(&event_log_path)) > events_after_save,
            "event log should grow after recalc"
        );
        assert!(
            event_trace_present_in_log(Path::new(&event_log_path), &trace_root),
            "events should include recalc trace_id"
        );
        assert!(
            artifact_index_present_in_log(
                Path::new(&artifact_index_path),
                &trace_root,
                "interop_recalc_loaded",
            ),
            "artifact index should include recalc command"
        );
        assert!(
            artifact_index_line_count(Path::new(&artifact_index_path)) >= 5,
            "artifact index should include open + preview + edit + save + recalc records"
        );

        let round_trip_trace_context = make_command_trace(
            &trace_root,
            &trace_session,
            "engine_round_trip",
            Some(&event_log_path),
            Some(&artifact_index_path),
        );
        let round_trip_payload = engine_round_trip(Some(round_trip_trace_context.clone()))
            .expect("round-trip should succeed");
        assert_eq!(round_trip_payload.trace.command_status, "success");
        assert!(!round_trip_payload.trace.linked_artifact_ids.is_empty());
        assert_trace_has_manifest_coverage(
            &round_trip_payload.trace,
            &load_artifact_index_records(Path::new(&artifact_index_path)),
        );
        assert_eq!(round_trip_payload.trace.trace_id, trace_root);
        assert_eq!(
            round_trip_payload.trace.trace_root_id,
            round_trip_payload.trace.trace_id
        );
        assert_eq!(
            round_trip_payload.trace.parent_span_id,
            Some(
                round_trip_trace_context
                    .span_id
                    .expect("round-trip trace span")
            )
        );
        assert_eq!(
            round_trip_payload.trace.ui_command_name.as_deref(),
            Some("engine_round_trip")
        );
        assert!(
            event_line_count(Path::new(&event_log_path)) > events_after_recalc,
            "event log should grow after round-trip"
        );
        assert!(
            event_trace_present_in_log(Path::new(&event_log_path), &trace_root),
            "events should include round-trip trace_id"
        );
        assert!(
            artifact_index_present_in_log(
                Path::new(&artifact_index_path),
                &trace_root,
                "engine_round_trip",
            ),
            "artifact index should include round-trip command"
        );
        assert!(
            artifact_index_line_count(Path::new(&artifact_index_path)) >= 6,
            "artifact index should include open + preview + edit + save + recalc + round-trip records"
        );

        let _ = std::fs::remove_file(output_path);
        let _ = std::fs::remove_file(&event_log_path);
        let _ = std::fs::remove_file(&artifact_index_path);
    }

    #[test]
    fn interop_undo_redo_restores_cell_state_and_counts() {
        let sample_path = sample_xlsx_path();
        let session = DesktopState::default();
        let state = as_state(&session);

        interop_open_workbook(state.clone(), sample_path.clone(), None)
            .expect("open should succeed");

        let initial = sheet1_cell_from_session(&session, 1, 1);

        let status_after_open = interop_session_status(state.clone())
            .expect("session status should be available after open");
        assert_eq!(status_after_open.undo_count, 0);
        assert_eq!(status_after_open.redo_count, 0);

        interop_apply_cell_edit(
            state.clone(),
            "Sheet1".to_string(),
            "A1".to_string(),
            "10".to_string(),
            UiEditMode::Value,
            None,
        )
        .expect("first edit should succeed");
        assert_eq!(
            sheet1_cell_from_session(&session, 1, 1),
            CellValue::Number(10.0)
        );

        let status_after_first_edit = interop_session_status(state.clone())
            .expect("session status should be available after first edit");
        assert_eq!(status_after_first_edit.undo_count, 1);
        assert_eq!(status_after_first_edit.redo_count, 0);

        interop_apply_cell_edit(
            state.clone(),
            "Sheet1".to_string(),
            "A1".to_string(),
            "20".to_string(),
            UiEditMode::Value,
            None,
        )
        .expect("second edit should succeed");
        assert_eq!(
            sheet1_cell_from_session(&session, 1, 1),
            CellValue::Number(20.0)
        );

        let status_after_second_edit = interop_session_status(state.clone())
            .expect("session status should be available after second edit");
        assert_eq!(status_after_second_edit.undo_count, 2);
        assert_eq!(status_after_second_edit.redo_count, 0);

        let undo_payload =
            interop_undo_edit(state.clone(), None).expect("undo should restore previous edit");
        assert_eq!(undo_payload.undo_count, 1);
        assert_eq!(undo_payload.redo_count, 1);
        assert_eq!(
            sheet1_cell_from_session(&session, 1, 1),
            CellValue::Number(10.0)
        );

        let undo_again_payload = interop_undo_edit(state.clone(), None)
            .expect("second undo should restore initial value");
        assert_eq!(undo_again_payload.undo_count, 0);
        assert_eq!(undo_again_payload.redo_count, 2);
        assert_eq!(sheet1_cell_from_session(&session, 1, 1), initial);

        let redo_payload =
            interop_redo_edit(state.clone(), None).expect("redo should reapply edit");
        assert_eq!(redo_payload.undo_count, 1);
        assert_eq!(redo_payload.redo_count, 1);
        assert_eq!(
            sheet1_cell_from_session(&session, 1, 1),
            CellValue::Number(10.0)
        );

        let redo_final_payload =
            interop_redo_edit(state.clone(), None).expect("redo should restore latest edit");
        assert_eq!(redo_final_payload.undo_count, 2);
        assert_eq!(redo_final_payload.redo_count, 0);
        assert_eq!(
            sheet1_cell_from_session(&session, 1, 1),
            CellValue::Number(20.0)
        );
    }

    #[test]
    fn interop_history_clears_redo_after_new_edit() {
        let sample_path = sample_xlsx_path();
        let session = DesktopState::default();
        let state = as_state(&session);

        interop_open_workbook(state.clone(), sample_path, None).expect("open should succeed");

        interop_apply_cell_edit(
            state.clone(),
            "Sheet1".to_string(),
            "A1".to_string(),
            "10".to_string(),
            UiEditMode::Value,
            None,
        )
        .expect("first edit should succeed");
        interop_apply_cell_edit(
            state.clone(),
            "Sheet1".to_string(),
            "A1".to_string(),
            "20".to_string(),
            UiEditMode::Value,
            None,
        )
        .expect("second edit should succeed");

        interop_undo_edit(state.clone(), None).expect("undo should succeed");
        let status_after_undo =
            interop_session_status(state.clone()).expect("session status should be available");
        assert_eq!(status_after_undo.undo_count, 1);
        assert_eq!(status_after_undo.redo_count, 1);

        interop_apply_cell_edit(
            state.clone(),
            "Sheet1".to_string(),
            "A1".to_string(),
            "30".to_string(),
            UiEditMode::Value,
            None,
        )
        .expect("new edit should succeed");

        let status_after_branch =
            interop_session_status(state.clone()).expect("session status should be available");
        assert_eq!(status_after_branch.undo_count, 2);
        assert_eq!(status_after_branch.redo_count, 0);
    }

    #[test]
    fn interop_run_macro_applies_mutations_and_recalculates_cells() {
        if !python_available() {
            eprintln!("python missing; skipping desktop macro integration test");
            return;
        }

        let session = DesktopState::default();
        let state = as_state(&session);
        let input_path = temp_workbook_with_macro_inputs("success");
        interop_open_workbook(
            state.clone(),
            input_path.to_string_lossy().to_string(),
            None,
        )
        .expect("workbook should open");
        let macro_path = input_path
            .parent()
            .expect("macro fixture directory")
            .join("macro_user.py");
        let worker_path = input_path
            .parent()
            .expect("macro fixture directory")
            .join("macro_worker.py");

        let success_worker = r#"
import json
import sys

json.load(sys.stdin)
print(
    json.dumps(
        {
            "status": "ok",
            "message": "ok",
            "stdout": "macro completed",
            "stderr": "",
            "permission_events": [
                {
                    "event_name": "script.permission.granted",
                    "permission": "fs.write",
                    "allowed": True,
                    "reason": "allowed by test override"
                },
                {
                    "event_name": "script.permission.granted",
                    "permission": "fs.read",
                    "allowed": True,
                    "reason": "allowed by test override"
                }
            ],
            "mutations": [
                {"op": "set_cell_value", "sheet": "Sheet1", "cell": "A1", "value": {"kind": "number", "value": 11}},
                {"op": "set_cell_formula", "sheet": "Sheet1", "cell": "B1", "formula": "=A1+1"},
                {"op": "set_cell_range_value", "sheet": "Sheet1", "start": "C1", "end": "D2", "value": {"kind": "number", "value": 22}}
            ],
            "result": {}
        }
    )
)
"#;

        write_temp_script(
            &macro_path,
            "def run_macro(ctx, args):\n    return {'ok': True}\n",
        );
        write_temp_script(&worker_path, success_worker);

        let response = with_env_vars(
            &[
                ("ROOTCELLAR_PYTHON", Some("python")),
                (
                    "ROOTCELLAR_SCRIPT_WORKER",
                    Some(worker_path.to_string_lossy().as_ref()),
                ),
            ],
            || {
                interop_run_macro(
                    state.clone(),
                    macro_path.to_string_lossy().to_string(),
                    "run_macro".to_string(),
                    String::new(),
                    InteropMacroPermissionConfig {
                        fs_read: true,
                        fs_write: true,
                        net_http: false,
                        clipboard: false,
                        process_exec: false,
                        udf: false,
                        events_emit: false,
                    },
                    None,
                )
                .expect("macro execution should succeed")
            },
        );

        assert_eq!(response.permission_granted, 2);
        assert_eq!(response.permission_denied, 0);
        assert_eq!(response.mutation_count, 6);
        assert_eq!(response.changed_sheets, vec!["Sheet1".to_string()]);
        assert!(!response.recalc_reports.is_empty());
        assert_eq!(response.stdout.as_deref(), Some("macro completed"));
        assert_eq!(response.script_fingerprint.len(), 16);
        assert_eq!(
            sheet1_cell_from_session(&session, 1, 1),
            CellValue::Number(11.0)
        );
        assert_eq!(
            sheet1_formula_from_session(&session, 1, 2),
            Some("=A1+1".to_string())
        );
        assert_eq!(read_cell_value(&session, 1, 3), CellValue::Number(22.0));
        assert_eq!(read_cell_value(&session, 2, 3), CellValue::Number(22.0));
        assert_eq!(read_cell_value(&session, 1, 4), CellValue::Number(22.0));
    }

    #[test]
    fn interop_run_macro_denies_mutations_without_permission() {
        if !python_available() {
            eprintln!("python missing; skipping desktop macro denial integration test");
            return;
        }

        let session = DesktopState::default();
        let state = as_state(&session);
        let input_path = temp_workbook_with_macro_inputs("denied");
        interop_open_workbook(
            state.clone(),
            input_path.to_string_lossy().to_string(),
            None,
        )
        .expect("workbook should open");
        let macro_path = input_path
            .parent()
            .expect("macro fixture directory")
            .join("macro_user.py");
        let worker_path = input_path
            .parent()
            .expect("macro fixture directory")
            .join("macro_worker.py");
        let denied_worker = r#"
import json
import sys

json.load(sys.stdin)
print(
    json.dumps(
        {
            "status": "ok",
            "message": "ok",
            "stdout": "permission denied",
            "stderr": "",
            "permission_events": [
                {
                    "event_name": "script.permission.denied",
                    "permission": "fs.write",
                    "allowed": False,
                    "reason": "permission denied for test"
                }
            ],
            "mutations": [],
            "result": {}
        }
    )
)
"#;

        let before = {
            let initial = read_cell_value(&session, 1, 1);
            let initial_formula = sheet1_formula_from_session(&session, 1, 2);
            (initial, initial_formula)
        };
        write_temp_script(
            &macro_path,
            "def run_macro(ctx, args):\n    raise RuntimeError('should be blocked by test worker')\n",
        );
        write_temp_script(&worker_path, denied_worker);

        let response = with_env_vars(
            &[
                ("ROOTCELLAR_PYTHON", Some("python")),
                (
                    "ROOTCELLAR_SCRIPT_WORKER",
                    Some(worker_path.to_string_lossy().as_ref()),
                ),
            ],
            || {
                interop_run_macro(
                    state.clone(),
                    macro_path.to_string_lossy().to_string(),
                    "run_macro".to_string(),
                    String::new(),
                    InteropMacroPermissionConfig {
                        fs_read: true,
                        fs_write: false,
                        net_http: false,
                        clipboard: false,
                        process_exec: false,
                        udf: false,
                        events_emit: false,
                    },
                    None,
                )
                .expect("macro execution should succeed even without mutation")
            },
        );

        assert_eq!(response.permission_granted, 0);
        assert_eq!(response.permission_denied, 1);
        assert_eq!(response.mutation_count, 0);
        assert!(response.changed_sheets.is_empty());
        assert!(response.recalc_reports.is_empty());
        assert_eq!(response.script_fingerprint.len(), 16);
        assert_eq!(sheet1_cell_from_session(&session, 1, 1), before.0);
        assert_eq!(sheet1_formula_from_session(&session, 1, 2), before.1);
        assert_eq!(response.stdout.as_deref(), Some("permission denied"));
    }

    #[test]
    fn interop_run_macro_fails_with_invalid_worker_response() {
        if !python_available() {
            eprintln!("python missing; skipping desktop macro invalid worker integration test");
            return;
        }

        let session = DesktopState::default();
        let state = as_state(&session);
        let input_path = temp_workbook_with_macro_inputs("invalid-worker");
        interop_open_workbook(
            state.clone(),
            input_path.to_string_lossy().to_string(),
            None,
        )
        .expect("workbook should open");
        let macro_path = input_path
            .parent()
            .expect("macro fixture directory")
            .join("macro_user.py");
        let worker_path = input_path
            .parent()
            .expect("macro fixture directory")
            .join("macro_worker.py");

        write_temp_script(
            &macro_path,
            "def run_macro(ctx, args):\n    return {'ok': True}\n",
        );
        write_temp_script(&worker_path, "print('not-json')");

        let response = with_env_vars(
            &[
                ("ROOTCELLAR_PYTHON", Some("python")),
                (
                    "ROOTCELLAR_SCRIPT_WORKER",
                    Some(worker_path.to_string_lossy().as_ref()),
                ),
            ],
            || {
                interop_run_macro(
                    state.clone(),
                    macro_path.to_string_lossy().to_string(),
                    "run_macro".to_string(),
                    String::new(),
                    InteropMacroPermissionConfig {
                        fs_read: true,
                        fs_write: true,
                        net_http: false,
                        clipboard: false,
                        process_exec: false,
                        udf: false,
                        events_emit: false,
                    },
                    None,
                )
            },
        );

        let error =
            response.expect_err("macro execution should fail with malformed worker response");
        assert!(error.contains("no JSON response payload"));
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
            &mut NoopEventSink,
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
            &mut NoopEventSink,
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
            interop_run_macro,
            interop_undo_edit,
            interop_redo_edit,
            interop_sheet_preview,
            interop_save_workbook
        ])
        .plugin(tauri_plugin_dialog::init())
        .run(tauri::generate_context!())
        .expect("failed to run rootcellar desktop shell");
}
