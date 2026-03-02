use crate::model::{CellRecord, CellRef, CellValue, Sheet, Workbook};
use crate::telemetry::{EventEnvelope, EventSink, TelemetryError, TraceContext};
use chrono::{DateTime, Utc};
use regex::Regex;
use serde::{Deserialize, Serialize};
use serde_json::json;
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::fs::File;
use std::io::{BufReader, Read, Write};
use std::path::{Path, PathBuf};
use std::time::Instant;
use thiserror::Error;
use zip::result::ZipError;
use zip::write::SimpleFileOptions;
use zip::ZipArchive;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum CompatibilityStatus {
    Supported,
    PartiallySupported,
    PreservedOnly,
    NotSupported,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CompatibilityIssue {
    pub code: String,
    pub title: String,
    pub status: CompatibilityStatus,
    pub details: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RequiredParts {
    pub content_types: bool,
    pub workbook_xml: bool,
    pub worksheet_count: usize,
    pub styles_xml: bool,
    pub shared_strings_xml: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ReportSummary {
    pub issue_count: usize,
    pub unknown_part_count: usize,
    pub workbook_feature_score: u8,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XlsxInspectionReport {
    pub workbook_path: PathBuf,
    pub generated_at: DateTime<Utc>,
    pub required_parts: RequiredParts,
    pub summary: ReportSummary,
    pub issues: Vec<CompatibilityIssue>,
    pub unknown_parts: Vec<String>,
    pub part_graph: WorkbookPartGraph,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookPartGraph {
    pub node_count: usize,
    pub edge_count: usize,
    pub dangling_edge_count: usize,
    pub external_edge_count: usize,
    pub unknown_part_count: usize,
    pub nodes: Vec<WorkbookPartNode>,
    pub edges: Vec<WorkbookPartEdge>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookPartNode {
    pub path: String,
    pub known: bool,
    pub synthetic: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookPartEdge {
    pub source: String,
    pub relationship_id: String,
    pub relationship_type: String,
    pub target: String,
    pub target_mode: String,
    pub dangling_target: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookPartGraphSummary {
    pub node_count: usize,
    pub edge_count: usize,
    pub dangling_edge_count: usize,
    pub external_edge_count: usize,
    pub unknown_part_count: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SavePartGraphFlags {
    pub strategy: String,
    pub source_graph_reused: bool,
    pub relationships_preserved: bool,
    pub unknown_parts_preserved: bool,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum SaveMode {
    Preserve,
    Normalize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XlsxSaveReport {
    pub output_path: PathBuf,
    pub generated_at: DateTime<Utc>,
    pub mode: SaveMode,
    pub sheet_count: usize,
    pub cell_count: usize,
    pub copied_bytes: u64,
    pub part_graph: WorkbookPartGraphSummary,
    pub part_graph_flags: SavePartGraphFlags,
}

#[derive(Debug, Error)]
pub enum InteropError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("zip error: {0}")]
    Zip(#[from] ZipError),
    #[error("telemetry error: {0}")]
    Telemetry(#[from] TelemetryError),
    #[error("invalid file extension; expected .xlsx")]
    InvalidExtension,
    #[error("missing required part: {0}")]
    MissingPart(String),
    #[error("parse error: {0}")]
    Parse(String),
}

#[derive(Debug, Clone)]
struct SheetMeta {
    name: String,
    rel_id: Option<String>,
}

#[derive(Debug, Clone)]
struct RelationshipEntry {
    id: String,
    rel_type: String,
    target: String,
    target_mode: Option<String>,
}

impl WorkbookPartGraph {
    pub fn summary(&self) -> WorkbookPartGraphSummary {
        WorkbookPartGraphSummary {
            node_count: self.node_count,
            edge_count: self.edge_count,
            dangling_edge_count: self.dangling_edge_count,
            external_edge_count: self.external_edge_count,
            unknown_part_count: self.unknown_part_count,
        }
    }
}

pub fn inspect_xlsx(
    path: impl AsRef<Path>,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<XlsxInspectionReport, InteropError> {
    let path = path.as_ref();
    ensure_xlsx_extension(path)?;

    let start = Instant::now();
    sink.emit(
        EventEnvelope::info("interop.xlsx.load.start", trace).with_context(json!({
            "path": path.display().to_string(),
        })),
    )?;

    let parts = read_archive_parts(path)?;
    let part_names = parts.keys().cloned().collect::<Vec<_>>();
    let worksheet_re = Regex::new(r"^xl/worksheets/[^/]+\.xml$").expect("valid regex");

    let required = RequiredParts {
        content_types: part_names.iter().any(|p| p == "[Content_Types].xml"),
        workbook_xml: part_names.iter().any(|p| p == "xl/workbook.xml"),
        worksheet_count: part_names
            .iter()
            .filter(|p| worksheet_re.is_match(p))
            .count(),
        styles_xml: part_names.iter().any(|p| p == "xl/styles.xml"),
        shared_strings_xml: part_names.iter().any(|p| p == "xl/sharedStrings.xml"),
    };

    let mut issues = Vec::new();

    if !required.content_types {
        issues.push(CompatibilityIssue {
            code: "XLSX_REQUIRED_CONTENT_TYPES_MISSING".to_string(),
            title: "Missing [Content_Types].xml".to_string(),
            status: CompatibilityStatus::NotSupported,
            details: "Workbook package is missing a required OpenXML part.".to_string(),
        });
    }

    if !required.workbook_xml {
        issues.push(CompatibilityIssue {
            code: "XLSX_REQUIRED_WORKBOOK_XML_MISSING".to_string(),
            title: "Missing xl/workbook.xml".to_string(),
            status: CompatibilityStatus::NotSupported,
            details: "Workbook structure cannot be loaded without workbook.xml.".to_string(),
        });
    }

    if required.worksheet_count == 0 {
        issues.push(CompatibilityIssue {
            code: "XLSX_WORKSHEETS_MISSING".to_string(),
            title: "No worksheet parts found".to_string(),
            status: CompatibilityStatus::NotSupported,
            details: "At least one worksheet XML part is required.".to_string(),
        });
    }

    if !required.styles_xml {
        issues.push(CompatibilityIssue {
            code: "XLSX_STYLES_MISSING".to_string(),
            title: "Missing xl/styles.xml".to_string(),
            status: CompatibilityStatus::PartiallySupported,
            details: "Workbook can load values but style fidelity may degrade on round-trip save."
                .to_string(),
        });
    }

    if !required.shared_strings_xml {
        issues.push(CompatibilityIssue {
            code: "XLSX_SHARED_STRINGS_MISSING".to_string(),
            title: "Missing xl/sharedStrings.xml".to_string(),
            status: CompatibilityStatus::PreservedOnly,
            details:
                "Workbook may still load inline strings; shared string optimization metadata is absent."
                    .to_string(),
        });
    }

    for issue in &issues {
        sink.emit(
            EventEnvelope::info("compat.issue.detected", trace).with_payload(json!({
                "code": issue.code,
                "status": issue.status,
                "title": issue.title,
            })),
        )?;
    }

    let unknown_parts = detect_unknown_parts(&part_names);
    let part_graph = build_workbook_part_graph(&parts);

    let score = calculate_feature_score(&required, issues.len(), unknown_parts.len());
    let report = XlsxInspectionReport {
        workbook_path: path.to_path_buf(),
        generated_at: Utc::now(),
        required_parts: required,
        summary: ReportSummary {
            issue_count: issues.len(),
            unknown_part_count: unknown_parts.len(),
            workbook_feature_score: score,
        },
        issues,
        unknown_parts,
        part_graph: part_graph.clone(),
    };

    let node_preview = part_graph
        .nodes
        .iter()
        .take(20)
        .map(|node| {
            json!({
                "path": node.path,
                "known": node.known,
                "synthetic": node.synthetic,
            })
        })
        .collect::<Vec<_>>();
    let edge_preview = part_graph
        .edges
        .iter()
        .take(20)
        .map(|edge| {
            json!({
                "source": edge.source,
                "relationship_id": edge.relationship_id,
                "relationship_type": edge.relationship_type,
                "target": edge.target,
                "target_mode": edge.target_mode,
                "dangling_target": edge.dangling_target,
            })
        })
        .collect::<Vec<_>>();
    sink.emit(
        EventEnvelope::info("interop.xlsx.part_graph.built", trace)
            .with_metrics(json!({
                "node_count": part_graph.node_count,
                "edge_count": part_graph.edge_count,
                "dangling_edge_count": part_graph.dangling_edge_count,
                "external_edge_count": part_graph.external_edge_count,
                "unknown_part_count": part_graph.unknown_part_count,
            }))
            .with_payload(json!({
                "node_preview": node_preview,
                "node_preview_truncated": part_graph.nodes.len() > 20,
                "edge_preview": edge_preview,
                "edge_preview_truncated": part_graph.edges.len() > 20,
            })),
    )?;

    sink.emit(
        EventEnvelope::info("interop.xlsx.load.end", trace)
            .with_metrics(json!({
                "duration_ms": start.elapsed().as_secs_f64() * 1000.0,
                "issue_count": report.summary.issue_count,
                "unknown_part_count": report.summary.unknown_part_count,
                "worksheet_count": report.required_parts.worksheet_count,
                "part_graph_node_count": report.part_graph.node_count,
                "part_graph_edge_count": report.part_graph.edge_count,
                "part_graph_dangling_edge_count": report.part_graph.dangling_edge_count,
            }))
            .with_payload(json!({
                "workbook_feature_score": report.summary.workbook_feature_score,
            })),
    )?;

    Ok(report)
}

pub fn load_workbook_model(
    path: impl AsRef<Path>,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<Workbook, InteropError> {
    let path = path.as_ref();
    ensure_xlsx_extension(path)?;

    let start = Instant::now();
    sink.emit(
        EventEnvelope::info("interop.xlsx.model_load.start", trace).with_context(json!({
            "path": path.display().to_string(),
        })),
    )?;

    let parts = read_archive_parts(path)?;
    let workbook_xml = part_as_string(&parts, "xl/workbook.xml")?
        .ok_or_else(|| InteropError::MissingPart("xl/workbook.xml".to_string()))?;

    let rel_map = parts
        .get("xl/_rels/workbook.xml.rels")
        .map(|raw| parse_workbook_relationships(&String::from_utf8_lossy(raw)))
        .unwrap_or_default();

    let shared_strings = parts
        .get("xl/sharedStrings.xml")
        .map(|raw| parse_shared_strings(&String::from_utf8_lossy(raw)))
        .unwrap_or_default();

    let sheet_metas = parse_workbook_sheet_metas(&workbook_xml)?;
    let mut fallback_sheet_paths = parts
        .keys()
        .filter(|part| part.starts_with("xl/worksheets/") && part.ends_with(".xml"))
        .cloned()
        .collect::<Vec<_>>();
    fallback_sheet_paths.sort();

    let mut workbook = Workbook::new();

    for (index, meta) in sheet_metas.iter().enumerate() {
        let resolved_path = meta
            .rel_id
            .as_ref()
            .and_then(|id| rel_map.get(id))
            .cloned()
            .or_else(|| fallback_sheet_paths.get(index).cloned());

        let cells = if let Some(sheet_path) = resolved_path {
            if let Some(sheet_xml) = part_as_string(&parts, &sheet_path)? {
                parse_worksheet_cells(&sheet_xml, &shared_strings)?
            } else {
                sink.emit(
                    EventEnvelope::info("interop.xlsx.model_load.missing_sheet_part", trace)
                        .with_payload(json!({
                            "sheet": meta.name,
                            "sheet_path": sheet_path,
                        })),
                )?;
                BTreeMap::new()
            }
        } else {
            BTreeMap::new()
        };

        workbook.sheets.insert(
            meta.name.clone(),
            Sheet {
                name: meta.name.clone(),
                cells,
            },
        );
    }

    if workbook.sheets.is_empty() {
        for (index, path) in fallback_sheet_paths.iter().enumerate() {
            if let Some(sheet_xml) = part_as_string(&parts, path)? {
                let cells = parse_worksheet_cells(&sheet_xml, &shared_strings)?;
                let name = format!("Sheet{}", index + 1);
                workbook.sheets.insert(name.clone(), Sheet { name, cells });
            }
        }
    }

    if workbook.sheets.is_empty() {
        return Err(InteropError::Parse(
            "could not resolve any worksheet XML parts".to_string(),
        ));
    }

    let total_cells = workbook
        .sheets
        .values()
        .map(|sheet| sheet.cells.len())
        .sum::<usize>();

    sink.emit(
        EventEnvelope::info("interop.xlsx.model_load.end", trace)
            .with_workbook_id(workbook.workbook_id)
            .with_metrics(json!({
                "duration_ms": start.elapsed().as_secs_f64() * 1000.0,
                "sheet_count": workbook.sheets.len(),
                "cell_count": total_cells,
                "shared_string_count": shared_strings.len(),
            })),
    )?;

    Ok(workbook)
}

pub fn save_workbook_model(
    workbook: &Workbook,
    output_path: impl AsRef<Path>,
    mode: SaveMode,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<XlsxSaveReport, InteropError> {
    let output_path = output_path.as_ref();
    ensure_xlsx_extension(output_path)?;
    let start = Instant::now();

    sink.emit(
        EventEnvelope::info("interop.xlsx.save.start", trace)
            .with_workbook_id(workbook.workbook_id)
            .with_context(json!({
                "path": output_path.display().to_string(),
                "mode": mode,
            })),
    )?;

    let sheets = if workbook.sheets.is_empty() {
        vec![Sheet {
            name: "Sheet1".to_string(),
            cells: BTreeMap::new(),
        }]
    } else {
        workbook.sheets.values().cloned().collect::<Vec<_>>()
    };

    let content_types_xml = build_content_types_xml(sheets.len());
    let root_rels_xml = build_root_relationships_xml();
    let workbook_xml = build_workbook_xml(&sheets);
    let workbook_rels_xml = build_workbook_relationships_xml(sheets.len());

    let file = File::create(output_path)?;
    let mut zip_writer = zip::ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    write_zip_part(
        &mut zip_writer,
        "[Content_Types].xml",
        content_types_xml.as_bytes(),
        options,
    )?;
    write_zip_part(
        &mut zip_writer,
        "_rels/.rels",
        root_rels_xml.as_bytes(),
        options,
    )?;
    write_zip_part(
        &mut zip_writer,
        "xl/workbook.xml",
        workbook_xml.as_bytes(),
        options,
    )?;
    write_zip_part(
        &mut zip_writer,
        "xl/_rels/workbook.xml.rels",
        workbook_rels_xml.as_bytes(),
        options,
    )?;

    for (index, sheet) in sheets.iter().enumerate() {
        let sheet_xml = build_sheet_xml(sheet);
        write_zip_part(
            &mut zip_writer,
            &format!("xl/worksheets/sheet{}.xml", index + 1),
            sheet_xml.as_bytes(),
            options,
        )?;
    }

    zip_writer.finish()?;

    let output_parts = read_archive_parts(output_path)?;
    let part_graph = build_workbook_part_graph(&output_parts);
    let part_graph_summary = part_graph.summary();
    let cell_count = sheets.iter().map(|s| s.cells.len()).sum::<usize>();
    let report = XlsxSaveReport {
        output_path: output_path.to_path_buf(),
        generated_at: Utc::now(),
        mode,
        sheet_count: sheets.len(),
        cell_count,
        copied_bytes: 0,
        part_graph: part_graph_summary,
        part_graph_flags: SavePartGraphFlags {
            strategy: "normalized_rebuild".to_string(),
            source_graph_reused: false,
            relationships_preserved: false,
            unknown_parts_preserved: false,
        },
    };

    sink.emit(
        EventEnvelope::info("interop.xlsx.save.end", trace)
            .with_workbook_id(workbook.workbook_id)
            .with_metrics(json!({
                "duration_ms": start.elapsed().as_secs_f64() * 1000.0,
                "sheet_count": report.sheet_count,
                "cell_count": report.cell_count,
                "part_graph_node_count": report.part_graph.node_count,
                "part_graph_edge_count": report.part_graph.edge_count,
                "part_graph_dangling_edge_count": report.part_graph.dangling_edge_count,
                "part_graph_unknown_part_count": report.part_graph.unknown_part_count,
            }))
            .with_payload(json!({
                "path": report.output_path.display().to_string(),
                "mode": report.mode,
                "part_graph_flags": report.part_graph_flags,
            })),
    )?;

    Ok(report)
}

pub fn preserve_xlsx_passthrough(
    input_path: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
    workbook: &Workbook,
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<XlsxSaveReport, InteropError> {
    let input_path = input_path.as_ref();
    let output_path = output_path.as_ref();
    ensure_xlsx_extension(input_path)?;
    ensure_xlsx_extension(output_path)?;

    let start = Instant::now();
    sink.emit(
        EventEnvelope::info("interop.xlsx.save.start", trace)
            .with_workbook_id(workbook.workbook_id)
            .with_context(json!({
                "path": output_path.display().to_string(),
                "mode": SaveMode::Preserve,
                "strategy": "passthrough_copy",
            })),
    )?;

    let copied_bytes = fs::copy(input_path, output_path)?;
    let output_parts = read_archive_parts(output_path)?;
    let part_graph = build_workbook_part_graph(&output_parts);
    let part_graph_summary = part_graph.summary();
    let cell_count = workbook
        .sheets
        .values()
        .map(|sheet| sheet.cells.len())
        .sum::<usize>();

    let report = XlsxSaveReport {
        output_path: output_path.to_path_buf(),
        generated_at: Utc::now(),
        mode: SaveMode::Preserve,
        sheet_count: workbook.sheets.len(),
        cell_count,
        copied_bytes,
        part_graph: part_graph_summary,
        part_graph_flags: SavePartGraphFlags {
            strategy: "passthrough_copy".to_string(),
            source_graph_reused: true,
            relationships_preserved: true,
            unknown_parts_preserved: true,
        },
    };

    sink.emit(
        EventEnvelope::info("interop.xlsx.save.end", trace)
            .with_workbook_id(workbook.workbook_id)
            .with_metrics(json!({
                "duration_ms": start.elapsed().as_secs_f64() * 1000.0,
                "sheet_count": report.sheet_count,
                "cell_count": report.cell_count,
                "copied_bytes": report.copied_bytes,
                "part_graph_node_count": report.part_graph.node_count,
                "part_graph_edge_count": report.part_graph.edge_count,
                "part_graph_dangling_edge_count": report.part_graph.dangling_edge_count,
                "part_graph_unknown_part_count": report.part_graph.unknown_part_count,
            }))
            .with_payload(json!({
                "path": report.output_path.display().to_string(),
                "mode": report.mode,
                "strategy": "passthrough_copy",
                "part_graph_flags": report.part_graph_flags,
            })),
    )?;

    Ok(report)
}

pub fn preserve_xlsx_with_sheet_overrides(
    input_path: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
    workbook: &Workbook,
    changed_sheets: &[String],
    sink: &mut dyn EventSink,
    trace: &TraceContext,
) -> Result<XlsxSaveReport, InteropError> {
    let input_path = input_path.as_ref();
    let output_path = output_path.as_ref();
    ensure_xlsx_extension(input_path)?;
    ensure_xlsx_extension(output_path)?;

    if changed_sheets.is_empty() {
        return preserve_xlsx_passthrough(input_path, output_path, workbook, sink, trace);
    }

    let start = Instant::now();
    let changed_set = changed_sheets.iter().cloned().collect::<BTreeSet<_>>();
    sink.emit(
        EventEnvelope::info("interop.xlsx.save.start", trace)
            .with_workbook_id(workbook.workbook_id)
            .with_context(json!({
                "path": output_path.display().to_string(),
                "mode": SaveMode::Preserve,
                "strategy": "sheet_overrides",
                "changed_sheet_count": changed_set.len(),
            })),
    )?;

    let mut parts = read_archive_parts(input_path)?;
    let sheet_part_paths = resolve_sheet_part_paths(&parts)?;
    let mut rewritten_sheet_count = 0usize;

    for sheet_name in &changed_set {
        let sheet = workbook.sheets.get(sheet_name).ok_or_else(|| {
            InteropError::Parse(format!(
                "changed sheet '{}' not found in workbook model",
                sheet_name
            ))
        })?;
        let sheet_part = sheet_part_paths.get(sheet_name).ok_or_else(|| {
            InteropError::Parse(format!(
                "could not resolve worksheet part for '{}'",
                sheet_name
            ))
        })?;

        let xml = build_sheet_xml(sheet);
        parts.insert(sheet_part.clone(), xml.into_bytes());
        rewritten_sheet_count += 1;
    }

    let copied_bytes = write_archive_parts(output_path, &parts)?;
    let part_graph = build_workbook_part_graph(&parts);
    let part_graph_summary = part_graph.summary();
    let cell_count = workbook
        .sheets
        .values()
        .map(|sheet| sheet.cells.len())
        .sum::<usize>();

    let report = XlsxSaveReport {
        output_path: output_path.to_path_buf(),
        generated_at: Utc::now(),
        mode: SaveMode::Preserve,
        sheet_count: workbook.sheets.len(),
        cell_count,
        copied_bytes,
        part_graph: part_graph_summary,
        part_graph_flags: SavePartGraphFlags {
            strategy: "sheet_overrides".to_string(),
            source_graph_reused: true,
            relationships_preserved: true,
            unknown_parts_preserved: true,
        },
    };

    sink.emit(
        EventEnvelope::info("interop.xlsx.save.end", trace)
            .with_workbook_id(workbook.workbook_id)
            .with_metrics(json!({
                "duration_ms": start.elapsed().as_secs_f64() * 1000.0,
                "sheet_count": report.sheet_count,
                "cell_count": report.cell_count,
                "copied_bytes": report.copied_bytes,
                "rewritten_sheet_count": rewritten_sheet_count,
                "part_graph_node_count": report.part_graph.node_count,
                "part_graph_edge_count": report.part_graph.edge_count,
                "part_graph_dangling_edge_count": report.part_graph.dangling_edge_count,
                "part_graph_unknown_part_count": report.part_graph.unknown_part_count,
            }))
            .with_payload(json!({
                "path": report.output_path.display().to_string(),
                "mode": report.mode,
                "strategy": "sheet_overrides",
                "part_graph_flags": report.part_graph_flags,
            })),
    )?;

    Ok(report)
}

fn ensure_xlsx_extension(path: &Path) -> Result<(), InteropError> {
    if !path
        .extension()
        .and_then(|x| x.to_str())
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx"))
    {
        return Err(InteropError::InvalidExtension);
    }
    Ok(())
}

fn write_archive_parts(
    output_path: &Path,
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<u64, InteropError> {
    let file = File::create(output_path)?;
    let mut zip_writer = zip::ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    for (part_name, bytes) in parts {
        write_zip_part(&mut zip_writer, part_name, bytes, options)?;
    }
    zip_writer.finish()?;
    Ok(fs::metadata(output_path)?.len())
}

fn resolve_sheet_part_paths(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<BTreeMap<String, String>, InteropError> {
    let workbook_xml = part_as_string(parts, "xl/workbook.xml")?
        .ok_or_else(|| InteropError::MissingPart("xl/workbook.xml".to_string()))?;
    let sheet_metas = parse_workbook_sheet_metas(&workbook_xml)?;
    let rel_map = parts
        .get("xl/_rels/workbook.xml.rels")
        .map(|raw| parse_workbook_relationships(&String::from_utf8_lossy(raw)))
        .unwrap_or_default();
    let mut fallback_sheet_paths = parts
        .keys()
        .filter(|part| part.starts_with("xl/worksheets/") && part.ends_with(".xml"))
        .cloned()
        .collect::<Vec<_>>();
    fallback_sheet_paths.sort();

    let mut result = BTreeMap::new();
    for (index, meta) in sheet_metas.iter().enumerate() {
        let resolved = meta
            .rel_id
            .as_ref()
            .and_then(|id| rel_map.get(id))
            .cloned()
            .or_else(|| fallback_sheet_paths.get(index).cloned());
        if let Some(path) = resolved {
            result.insert(meta.name.clone(), path);
        }
    }

    Ok(result)
}

fn write_zip_part(
    writer: &mut zip::ZipWriter<File>,
    part_path: &str,
    content: &[u8],
    options: SimpleFileOptions,
) -> Result<(), InteropError> {
    writer.start_file(part_path, options)?;
    writer.write_all(content)?;
    Ok(())
}

fn build_content_types_xml(sheet_count: usize) -> String {
    let mut overrides = String::new();
    overrides.push_str("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
    for i in 1..=sheet_count {
        overrides.push_str(&format!(
            "<Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        ));
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
<Default Extension=\"xml\" ContentType=\"application/xml\"/>\
{overrides}\
</Types>"
    )
}

fn build_root_relationships_xml() -> String {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\
</Relationships>"
        .to_string()
}

fn build_workbook_xml(sheets: &[Sheet]) -> String {
    let mut sheet_nodes = String::new();
    for (index, sheet) in sheets.iter().enumerate() {
        let sheet_name = escape_xml_attr(&sheet.name);
        let sheet_id = index + 1;
        sheet_nodes.push_str(&format!(
            "<sheet name=\"{sheet_name}\" sheetId=\"{sheet_id}\" r:id=\"rId{sheet_id}\"/>"
        ));
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
<sheets>{sheet_nodes}</sheets>\
</workbook>"
    )
}

fn build_workbook_relationships_xml(sheet_count: usize) -> String {
    let mut rel_nodes = String::new();
    for i in 1..=sheet_count {
        rel_nodes.push_str(&format!(
            "<Relationship Id=\"rId{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i}.xml\"/>"
        ));
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
{rel_nodes}\
</Relationships>"
    )
}

fn build_sheet_xml(sheet: &Sheet) -> String {
    let mut xml = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>",
    );

    let mut current_row: Option<u32> = None;
    for (cell_ref, cell) in &sheet.cells {
        let Some(cell_xml) = render_cell_xml(cell_ref, cell) else {
            continue;
        };

        if current_row != Some(cell_ref.row) {
            if current_row.is_some() {
                xml.push_str("</row>");
            }
            xml.push_str(&format!("<row r=\"{}\">", cell_ref.row));
            current_row = Some(cell_ref.row);
        }

        xml.push_str(&cell_xml);
    }

    if current_row.is_some() {
        xml.push_str("</row>");
    }

    xml.push_str("</sheetData></worksheet>");
    xml
}

fn render_cell_xml(cell_ref: &CellRef, cell: &CellRecord) -> Option<String> {
    if cell.formula.is_none() && matches!(cell.value, CellValue::Empty) {
        return None;
    }

    let cell_addr = to_a1(cell_ref.row, cell_ref.col);

    if let Some(formula) = &cell.formula {
        let formula_body = formula.strip_prefix('=').unwrap_or(formula);
        match &cell.value {
            CellValue::Number(n) => Some(format!(
                "<c r=\"{cell_addr}\"><f>{}</f><v>{}</v></c>",
                escape_xml_text(formula_body),
                format_number(*n)
            )),
            CellValue::Bool(v) => Some(format!(
                "<c r=\"{cell_addr}\" t=\"b\"><f>{}</f><v>{}</v></c>",
                escape_xml_text(formula_body),
                if *v { "1" } else { "0" }
            )),
            CellValue::Text(v) => Some(format!(
                "<c r=\"{cell_addr}\" t=\"str\"><f>{}</f><v>{}</v></c>",
                escape_xml_text(formula_body),
                escape_xml_text(v)
            )),
            CellValue::Error(v) => Some(format!(
                "<c r=\"{cell_addr}\" t=\"e\"><f>{}</f><v>{}</v></c>",
                escape_xml_text(formula_body),
                escape_xml_text(v)
            )),
            CellValue::Empty => Some(format!(
                "<c r=\"{cell_addr}\"><f>{}</f></c>",
                escape_xml_text(formula_body)
            )),
        }
    } else {
        match &cell.value {
            CellValue::Number(n) => Some(format!(
                "<c r=\"{cell_addr}\"><v>{}</v></c>",
                format_number(*n)
            )),
            CellValue::Bool(v) => Some(format!(
                "<c r=\"{cell_addr}\" t=\"b\"><v>{}</v></c>",
                if *v { "1" } else { "0" }
            )),
            CellValue::Text(v) => Some(format!(
                "<c r=\"{cell_addr}\" t=\"inlineStr\"><is><t>{}</t></is></c>",
                escape_xml_text(v)
            )),
            CellValue::Error(v) => Some(format!(
                "<c r=\"{cell_addr}\" t=\"e\"><v>{}</v></c>",
                escape_xml_text(v)
            )),
            CellValue::Empty => None,
        }
    }
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

fn format_number(value: f64) -> String {
    if value.fract() == 0.0 {
        format!("{:.0}", value)
    } else {
        format!("{value}")
    }
}

fn escape_xml_attr(input: &str) -> String {
    escape_xml_text(input).replace('\"', "&quot;")
}

fn escape_xml_text(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('\"', "&quot;")
        .replace('\'', "&apos;")
}

fn read_archive_parts(path: &Path) -> Result<BTreeMap<String, Vec<u8>>, InteropError> {
    let file = File::open(path)?;
    let reader = BufReader::new(file);
    let mut archive = ZipArchive::new(reader)?;
    let mut parts = BTreeMap::new();

    for i in 0..archive.len() {
        let mut zip_file = archive.by_index(i)?;
        if zip_file.is_dir() {
            continue;
        }

        let mut content = Vec::new();
        zip_file.read_to_end(&mut content)?;
        parts.insert(zip_file.name().to_string(), content);
    }

    Ok(parts)
}

fn part_as_string(
    parts: &BTreeMap<String, Vec<u8>>,
    part_name: &str,
) -> Result<Option<String>, InteropError> {
    if let Some(content) = parts.get(part_name) {
        return Ok(Some(String::from_utf8_lossy(content).to_string()));
    }
    Ok(None)
}

fn parse_workbook_sheet_metas(workbook_xml: &str) -> Result<Vec<SheetMeta>, InteropError> {
    let sheet_re = Regex::new(r"<sheet\b[^>]*>").expect("valid regex");
    let mut metas = Vec::new();

    for capture in sheet_re.captures_iter(workbook_xml) {
        let tag = capture.get(0).map(|x| x.as_str()).unwrap_or_default();
        let attrs = tag.trim_start_matches("<sheet").trim_end_matches('>');
        let name = get_attr(attrs, "name").unwrap_or_else(|| format!("Sheet{}", metas.len() + 1));
        let rel_id = get_attr(attrs, "r:id");
        metas.push(SheetMeta { name, rel_id });
    }

    if metas.is_empty() {
        return Err(InteropError::Parse(
            "workbook.xml did not contain any <sheet> definitions".to_string(),
        ));
    }

    Ok(metas)
}

fn parse_workbook_relationships(rels_xml: &str) -> BTreeMap<String, String> {
    let mut map = BTreeMap::new();

    for entry in parse_relationship_entries(rels_xml) {
        if entry
            .target_mode
            .as_deref()
            .map(|mode| mode.eq_ignore_ascii_case("External"))
            .unwrap_or(false)
        {
            continue;
        }
        map.insert(entry.id, normalize_zip_path("xl", &entry.target));
    }

    map
}

fn parse_relationship_entries(rels_xml: &str) -> Vec<RelationshipEntry> {
    let rel_re = Regex::new(r"<Relationship\b[^>]*/?>").expect("valid regex");
    let mut entries = Vec::<RelationshipEntry>::new();
    for capture in rel_re.captures_iter(rels_xml) {
        let tag = capture.get(0).map(|x| x.as_str()).unwrap_or_default();
        let attrs = tag
            .trim_start_matches("<Relationship")
            .trim_end_matches('>')
            .trim_end_matches('/');
        let Some(id) = get_attr(attrs, "Id") else {
            continue;
        };
        let Some(target) = get_attr(attrs, "Target") else {
            continue;
        };
        entries.push(RelationshipEntry {
            id,
            rel_type: get_attr(attrs, "Type").unwrap_or_default(),
            target,
            target_mode: get_attr(attrs, "TargetMode"),
        });
    }
    entries
}

fn relationship_source_from_rels_part(rels_part: &str) -> Option<String> {
    if rels_part == "_rels/.rels" {
        return Some(String::new());
    }
    let (prefix, rest) = rels_part.split_once("/_rels/")?;
    let source_file = rest.strip_suffix(".rels")?;
    if prefix.is_empty() {
        return Some(source_file.to_string());
    }
    Some(format!("{prefix}/{source_file}"))
}

fn parent_part_dir(part: &str) -> String {
    part.rsplit_once('/')
        .map(|(parent, _)| parent.to_string())
        .unwrap_or_default()
}

fn is_relationship_part(part: &str) -> bool {
    part == "_rels/.rels" || (part.ends_with(".rels") && part.contains("/_rels/"))
}

fn build_workbook_part_graph(parts: &BTreeMap<String, Vec<u8>>) -> WorkbookPartGraph {
    let part_names = parts.keys().cloned().collect::<Vec<_>>();
    let unknown_parts = detect_unknown_parts(&part_names);
    let unknown_set = unknown_parts.iter().cloned().collect::<BTreeSet<_>>();
    let mut node_map = BTreeMap::<String, WorkbookPartNode>::new();

    for part in &part_names {
        node_map.insert(
            part.clone(),
            WorkbookPartNode {
                path: part.clone(),
                known: !unknown_set.contains(part),
                synthetic: false,
            },
        );
    }

    let mut edges = Vec::<WorkbookPartEdge>::new();
    for rels_part in part_names.iter().filter(|part| is_relationship_part(part)) {
        let Some(source_part) = relationship_source_from_rels_part(rels_part) else {
            continue;
        };
        let source_node = if source_part.is_empty() {
            "__package__".to_string()
        } else {
            source_part.clone()
        };
        if !node_map.contains_key(&source_node) {
            node_map.insert(
                source_node.clone(),
                WorkbookPartNode {
                    path: source_node.clone(),
                    known: source_part.is_empty() || is_known_part(&source_node),
                    synthetic: true,
                },
            );
        }

        let rels_xml = parts
            .get(rels_part)
            .map(|raw| String::from_utf8_lossy(raw).to_string())
            .unwrap_or_default();
        let entries = parse_relationship_entries(&rels_xml);
        let base_dir = parent_part_dir(&source_part);
        for entry in entries {
            let external = entry
                .target_mode
                .as_deref()
                .map(|mode| mode.eq_ignore_ascii_case("External"))
                .unwrap_or(false);
            let target = if external {
                entry.target.clone()
            } else {
                normalize_zip_path(&base_dir, &entry.target)
            };
            let dangling_target = !external && !parts.contains_key(&target);
            if dangling_target && !node_map.contains_key(&target) {
                node_map.insert(
                    target.clone(),
                    WorkbookPartNode {
                        path: target.clone(),
                        known: is_known_part(&target),
                        synthetic: true,
                    },
                );
            }
            edges.push(WorkbookPartEdge {
                source: source_node.clone(),
                relationship_id: entry.id,
                relationship_type: entry.rel_type,
                target,
                target_mode: if external {
                    "external".to_string()
                } else {
                    "internal".to_string()
                },
                dangling_target,
            });
        }
    }

    edges.sort_by(|a, b| {
        a.source
            .cmp(&b.source)
            .then_with(|| a.relationship_id.cmp(&b.relationship_id))
            .then_with(|| a.target.cmp(&b.target))
            .then_with(|| a.relationship_type.cmp(&b.relationship_type))
            .then_with(|| a.target_mode.cmp(&b.target_mode))
    });

    let nodes = node_map.into_values().collect::<Vec<_>>();
    let dangling_edge_count = edges.iter().filter(|edge| edge.dangling_target).count();
    let external_edge_count = edges
        .iter()
        .filter(|edge| edge.target_mode == "external")
        .count();
    let unknown_part_count = nodes
        .iter()
        .filter(|node| !node.synthetic && !node.known)
        .count();

    WorkbookPartGraph {
        node_count: nodes.len(),
        edge_count: edges.len(),
        dangling_edge_count,
        external_edge_count,
        unknown_part_count,
        nodes,
        edges,
    }
}

fn normalize_zip_path(base: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }

    let mut segments = base
        .split('/')
        .filter(|x| !x.is_empty())
        .map(|x| x.to_string())
        .collect::<Vec<_>>();

    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            other => segments.push(other.to_string()),
        }
    }

    segments.join("/")
}

fn parse_shared_strings(shared_strings_xml: &str) -> Vec<String> {
    let si_re = Regex::new(r"(?s)<si\b[^>]*>(.*?)</si>").expect("valid regex");
    let t_re = Regex::new(r"(?s)<t(?:\s[^>]*)?>(.*?)</t>").expect("valid regex");

    si_re
        .captures_iter(shared_strings_xml)
        .map(|cap| {
            let body = cap.get(1).map(|x| x.as_str()).unwrap_or_default();
            let mut text = String::new();
            for t_cap in t_re.captures_iter(body) {
                if let Some(segment) = t_cap.get(1) {
                    text.push_str(&decode_xml_entities(segment.as_str()));
                }
            }
            text
        })
        .collect::<Vec<_>>()
}

fn parse_worksheet_cells(
    sheet_xml: &str,
    shared_strings: &[String],
) -> Result<BTreeMap<CellRef, CellRecord>, InteropError> {
    let mut cells = BTreeMap::new();

    for (attrs, body) in extract_cell_nodes(sheet_xml)? {
        let Some(reference) = get_attr(&attrs, "r") else {
            continue;
        };

        let Some(cell_ref) = parse_a1_cell(&reference) else {
            continue;
        };

        let cell_type = get_attr(&attrs, "t");
        let formula = body
            .as_deref()
            .and_then(|xml| extract_first_tag_text(xml, "f"))
            .map(|formula| {
                if formula.starts_with('=') {
                    formula
                } else {
                    format!("={formula}")
                }
            });

        let value = if formula.is_some() {
            parse_formula_cached_value(cell_type.as_deref(), body.as_deref(), shared_strings)
        } else {
            parse_literal_cell_value(cell_type.as_deref(), body.as_deref(), shared_strings)
        };

        cells.insert(cell_ref, CellRecord { value, formula });
    }

    Ok(cells)
}

fn parse_formula_cached_value(
    cell_type: Option<&str>,
    body: Option<&str>,
    shared_strings: &[String],
) -> CellValue {
    let Some(body_xml) = body else {
        return CellValue::Empty;
    };
    let value_text = extract_first_tag_text(body_xml, "v");
    parse_value_from_components(cell_type, value_text.as_deref(), body_xml, shared_strings)
}

fn parse_literal_cell_value(
    cell_type: Option<&str>,
    body: Option<&str>,
    shared_strings: &[String],
) -> CellValue {
    let Some(body_xml) = body else {
        return CellValue::Empty;
    };

    let value_text = extract_first_tag_text(body_xml, "v");
    parse_value_from_components(cell_type, value_text.as_deref(), body_xml, shared_strings)
}

fn parse_value_from_components(
    cell_type: Option<&str>,
    value_text: Option<&str>,
    body_xml: &str,
    shared_strings: &[String],
) -> CellValue {
    match cell_type {
        Some("s") => {
            if let Some(idx_text) = value_text {
                if let Ok(idx) = idx_text.trim().parse::<usize>() {
                    if let Some(value) = shared_strings.get(idx) {
                        return CellValue::Text(value.clone());
                    }
                }
                return CellValue::Text(decode_xml_entities(idx_text));
            }
            CellValue::Empty
        }
        Some("b") => match value_text.unwrap_or("0").trim() {
            "1" => CellValue::Bool(true),
            _ => CellValue::Bool(false),
        },
        Some("e") => CellValue::Error(decode_xml_entities(value_text.unwrap_or("#ERROR"))),
        Some("inlineStr") => {
            if let Some(text) = extract_inline_string(body_xml) {
                CellValue::Text(text)
            } else {
                CellValue::Empty
            }
        }
        Some("str") => CellValue::Text(decode_xml_entities(value_text.unwrap_or_default())),
        _ => {
            if let Some(v) = value_text {
                let trimmed = v.trim();
                if trimmed.is_empty() {
                    CellValue::Empty
                } else if let Ok(n) = trimmed.parse::<f64>() {
                    CellValue::Number(n)
                } else {
                    CellValue::Text(decode_xml_entities(trimmed))
                }
            } else if let Some(text) = extract_inline_string(body_xml) {
                CellValue::Text(text)
            } else {
                CellValue::Empty
            }
        }
    }
}

fn extract_inline_string(body_xml: &str) -> Option<String> {
    let t_re = Regex::new(r"(?s)<t(?:\s[^>]*)?>(.*?)</t>").expect("valid regex");
    let mut out = String::new();
    for capture in t_re.captures_iter(body_xml) {
        if let Some(segment) = capture.get(1) {
            out.push_str(&decode_xml_entities(segment.as_str()));
        }
    }
    if out.is_empty() {
        None
    } else {
        Some(out)
    }
}

fn extract_first_tag_text(xml: &str, tag_name: &str) -> Option<String> {
    let open_token = format!("<{tag_name}");
    let open_start = xml.find(&open_token)?;
    let open_end = xml[open_start..].find('>')? + open_start;
    let close_token = format!("</{tag_name}>");
    let close_start = xml[open_end + 1..].find(&close_token)? + open_end + 1;
    let value = &xml[open_end + 1..close_start];
    Some(decode_xml_entities(value.trim()))
}

fn extract_cell_nodes(sheet_xml: &str) -> Result<Vec<(String, Option<String>)>, InteropError> {
    let mut out = Vec::new();
    let mut cursor = 0usize;

    while let Some(relative_start) = sheet_xml[cursor..].find("<c") {
        let start = cursor + relative_start;
        let Some(after_tag) = sheet_xml.as_bytes().get(start + 2).copied() else {
            break;
        };

        if !matches!(after_tag, b' ' | b'\t' | b'\n' | b'\r' | b'>' | b'/') {
            cursor = start + 2;
            continue;
        }

        let Some(open_end_rel) = sheet_xml[start..].find('>') else {
            return Err(InteropError::Parse(
                "unterminated <c> tag in worksheet xml".to_string(),
            ));
        };
        let open_end = start + open_end_rel;
        let open_tag = &sheet_xml[start + 2..open_end];

        if open_tag.trim_end().ends_with('/') {
            out.push((open_tag.trim_end_matches('/').trim().to_string(), None));
            cursor = open_end + 1;
            continue;
        }

        let Some(close_rel) = sheet_xml[open_end + 1..].find("</c>") else {
            return Err(InteropError::Parse(
                "missing </c> terminator in worksheet xml".to_string(),
            ));
        };
        let close_start = open_end + 1 + close_rel;
        let body = sheet_xml[open_end + 1..close_start].to_string();

        out.push((open_tag.trim().to_string(), Some(body)));
        cursor = close_start + 4;
    }

    Ok(out)
}

fn parse_a1_cell(reference: &str) -> Option<CellRef> {
    let mut col_part = String::new();
    let mut row_part = String::new();

    for ch in reference.chars() {
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

    Some(CellRef { row, col })
}

fn get_attr(attrs: &str, name: &str) -> Option<String> {
    for quote in ['\"', '\''] {
        let needle = format!("{name}={quote}");
        if let Some(start) = attrs.find(&needle) {
            let value_start = start + needle.len();
            if let Some(end_offset) = attrs[value_start..].find(quote) {
                let raw = &attrs[value_start..value_start + end_offset];
                return Some(decode_xml_entities(raw));
            }
        }
    }
    None
}

fn decode_xml_entities(input: &str) -> String {
    input
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

fn detect_unknown_parts(parts: &[String]) -> Vec<String> {
    let mut unknown = parts
        .iter()
        .filter(|part| !is_known_part(part))
        .cloned()
        .collect::<Vec<_>>();

    unknown.sort();
    unknown
}

fn is_known_part(part: &str) -> bool {
    let known_exact = [
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/styles.xml",
        "xl/sharedStrings.xml",
        "xl/calcChain.xml",
    ];
    let known_prefixes = [
        "docProps/",
        "_rels/",
        "xl/_rels/",
        "xl/worksheets/",
        "xl/worksheets/_rels/",
        "xl/theme/",
        "xl/charts/",
        "xl/drawings/",
        "xl/drawings/_rels/",
        "xl/tables/",
        "xl/tables/_rels/",
        "xl/pivotTables/",
        "xl/pivotCache/",
        "xl/media/",
        "customXml/",
        "xl/printerSettings/",
        "xl/externalLinks/",
        "xl/externalLinks/_rels/",
    ];

    known_exact.contains(&part)
        || known_prefixes.iter().any(|prefix| part.starts_with(prefix))
        || is_relationship_part(part)
}

fn calculate_feature_score(
    required: &RequiredParts,
    issue_count: usize,
    unknown_count: usize,
) -> u8 {
    let mut score: i32 = 100;
    if !required.content_types {
        score -= 30;
    }
    if !required.workbook_xml {
        score -= 30;
    }
    if required.worksheet_count == 0 {
        score -= 20;
    }
    if !required.styles_xml {
        score -= 8;
    }
    if !required.shared_strings_xml {
        score -= 4;
    }
    score -= (issue_count as i32) * 4;
    score -= std::cmp::min(10, unknown_count as i32);

    score.clamp(0, 100) as u8
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::telemetry::NoopEventSink;
    use std::io::Write;
    use tempfile::tempdir;
    use zip::write::SimpleFileOptions;

    #[test]
    fn inspects_minimal_xlsx_and_reports_expected_issues() {
        let dir = tempdir().expect("tempdir");
        let path = dir.path().join("test.xlsx");
        {
            let file = File::create(&path).expect("create file");
            let mut writer = zip::ZipWriter::new(file);
            let options = SimpleFileOptions::default();
            writer
                .start_file("[Content_Types].xml", options)
                .expect("start content types");
            writer.write_all(b"<Types/>").expect("write");
            writer
                .start_file("xl/workbook.xml", options)
                .expect("start workbook");
            writer.write_all(b"<workbook/>").expect("write");
            writer
                .start_file("xl/worksheets/sheet1.xml", options)
                .expect("start sheet");
            writer.write_all(b"<worksheet/>").expect("write");
            writer.finish().expect("finish zip");
        }

        let mut sink = NoopEventSink;
        let trace = TraceContext::root();
        let report = inspect_xlsx(&path, &mut sink, &trace).expect("inspect");

        assert!(report.required_parts.content_types);
        assert!(report.required_parts.workbook_xml);
        assert_eq!(report.required_parts.worksheet_count, 1);
        assert_eq!(report.summary.issue_count, 2);
        assert!(report
            .issues
            .iter()
            .any(|x| x.code == "XLSX_STYLES_MISSING"));
        assert_eq!(report.part_graph.edge_count, 0);
        assert_eq!(report.part_graph.dangling_edge_count, 0);
        assert!(report.part_graph.node_count >= 3);
        assert_eq!(report.part_graph.unknown_part_count, 0);
    }

    #[test]
    fn loads_workbook_model_values_and_formulas() {
        let dir = tempdir().expect("tempdir");
        let path = dir.path().join("model.xlsx");
        {
            let file = File::create(&path).expect("create file");
            let mut writer = zip::ZipWriter::new(file);
            let options = SimpleFileOptions::default();
            writer
                .start_file("[Content_Types].xml", options)
                .expect("content types");
            writer.write_all(b"<Types/>").expect("write");
            writer
                .start_file("xl/workbook.xml", options)
                .expect("workbook");
            writer
                .write_all(
                    br#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
                )
                .expect("write");
            writer
                .start_file("xl/_rels/workbook.xml.rels", options)
                .expect("rels");
            writer
                .write_all(
                    br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
                )
                .expect("write");
            writer
                .start_file("xl/sharedStrings.xml", options)
                .expect("shared strings");
            writer
                .write_all(br#"<sst><si><t>Root</t></si></sst>"#)
                .expect("write");
            writer
                .start_file("xl/worksheets/sheet1.xml", options)
                .expect("sheet");
            writer
                .write_all(
                    br#"<worksheet><sheetData><row r="1"><c r="A1"><v>40</v></c><c r="B1" t="s"><v>0</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><f>A1+A2</f><v>42</v></c></row></sheetData></worksheet>"#,
                )
                .expect("write");
            writer.finish().expect("finish");
        }

        let mut sink = NoopEventSink;
        let trace = TraceContext::root();
        let workbook = load_workbook_model(&path, &mut sink, &trace).expect("load workbook");

        let sheet = workbook.sheets.get("Sheet1").expect("sheet1");
        let text_value = sheet
            .cells
            .get(&CellRef { row: 1, col: 2 })
            .expect("b1")
            .value
            .clone();
        assert_eq!(text_value, CellValue::Text("Root".to_string()));

        let formula_cell = sheet.cells.get(&CellRef { row: 3, col: 1 }).expect("a3");
        assert_eq!(formula_cell.formula.as_deref(), Some("=A1+A2"));
        assert_eq!(formula_cell.value, CellValue::Number(42.0));
    }

    #[test]
    fn saves_and_reloads_workbook_model() {
        let dir = tempdir().expect("tempdir");
        let path = dir.path().join("saved.xlsx");

        let mut workbook = Workbook::new();
        let mut sheet_cells = BTreeMap::new();
        sheet_cells.insert(
            CellRef { row: 1, col: 1 },
            CellRecord {
                value: CellValue::Number(10.0),
                formula: None,
            },
        );
        sheet_cells.insert(
            CellRef { row: 1, col: 2 },
            CellRecord {
                value: CellValue::Number(5.0),
                formula: None,
            },
        );
        sheet_cells.insert(
            CellRef { row: 2, col: 1 },
            CellRecord {
                value: CellValue::Empty,
                formula: Some("=A1+B1".to_string()),
            },
        );
        sheet_cells.insert(
            CellRef { row: 3, col: 1 },
            CellRecord {
                value: CellValue::Text("hello".to_string()),
                formula: None,
            },
        );
        workbook.sheets.insert(
            "SheetA".to_string(),
            Sheet {
                name: "SheetA".to_string(),
                cells: sheet_cells,
            },
        );

        let mut sink = NoopEventSink;
        let trace = TraceContext::root();
        let save_report =
            save_workbook_model(&workbook, &path, SaveMode::Normalize, &mut sink, &trace)
                .expect("save");
        assert_eq!(save_report.sheet_count, 1);
        assert_eq!(save_report.cell_count, 4);
        assert_eq!(save_report.copied_bytes, 0);
        assert_eq!(save_report.part_graph_flags.strategy, "normalized_rebuild");
        assert!(!save_report.part_graph_flags.source_graph_reused);
        assert!(!save_report.part_graph_flags.relationships_preserved);
        assert!(!save_report.part_graph_flags.unknown_parts_preserved);
        assert!(save_report.part_graph.edge_count >= 2);
        assert_eq!(save_report.part_graph.dangling_edge_count, 0);

        let reloaded = load_workbook_model(&path, &mut sink, &trace).expect("reload");
        let sheet = reloaded.sheets.get("SheetA").expect("sheet");
        let formula_cell = sheet.cells.get(&CellRef { row: 2, col: 1 }).expect("a2");
        assert_eq!(formula_cell.formula.as_deref(), Some("=A1+B1"));
        let text_cell = sheet.cells.get(&CellRef { row: 3, col: 1 }).expect("a3");
        assert_eq!(text_cell.value, CellValue::Text("hello".to_string()));
    }

    #[test]
    fn preserve_passthrough_keeps_unknown_parts() {
        let dir = tempdir().expect("tempdir");
        let source = dir.path().join("source.xlsx");
        let output = dir.path().join("preserved.xlsx");
        let custom_payload = b"{\"custom\":\"payload\"}";

        {
            let file = File::create(&source).expect("create source");
            let mut writer = zip::ZipWriter::new(file);
            let options = SimpleFileOptions::default();
            writer
                .start_file("[Content_Types].xml", options)
                .expect("ct");
            writer.write_all(b"<Types/>").expect("write");
            writer.start_file("xl/workbook.xml", options).expect("wb");
            writer
                .write_all(
                    br#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
                )
                .expect("write");
            writer
                .start_file("xl/_rels/workbook.xml.rels", options)
                .expect("rels");
            writer
                .write_all(
                    br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
                )
                .expect("write");
            writer
                .start_file("xl/worksheets/sheet1.xml", options)
                .expect("sheet");
            writer
                .write_all(
                    br#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#,
                )
                .expect("write");
            writer
                .start_file("customXml/item1.bin", options)
                .expect("custom");
            writer.write_all(custom_payload).expect("write");
            writer.finish().expect("finish");
        }

        let mut sink = NoopEventSink;
        let trace = TraceContext::root();
        let workbook = load_workbook_model(&source, &mut sink, &trace).expect("load");
        let report = preserve_xlsx_passthrough(&source, &output, &workbook, &mut sink, &trace)
            .expect("copy");
        assert!(report.copied_bytes > 0);
        assert_eq!(report.part_graph_flags.strategy, "passthrough_copy");
        assert!(report.part_graph_flags.source_graph_reused);
        assert!(report.part_graph_flags.relationships_preserved);
        assert!(report.part_graph_flags.unknown_parts_preserved);
        assert!(report.part_graph.edge_count >= 1);

        let file = File::open(&output).expect("open output");
        let reader = BufReader::new(file);
        let mut archive = ZipArchive::new(reader).expect("archive");
        let mut custom = archive
            .by_name("customXml/item1.bin")
            .expect("custom part exists");
        let mut data = Vec::new();
        custom.read_to_end(&mut data).expect("read part");
        assert_eq!(data, custom_payload);
    }

    #[test]
    fn preserve_sheet_overrides_updates_target_sheet_only() {
        let dir = tempdir().expect("tempdir");
        let source = dir.path().join("source_override.xlsx");
        let output = dir.path().join("override_output.xlsx");
        let custom_payload = b"PRESERVE_ME";

        {
            let file = File::create(&source).expect("create source");
            let mut writer = zip::ZipWriter::new(file);
            let options = SimpleFileOptions::default();
            writer
                .start_file("[Content_Types].xml", options)
                .expect("ct");
            writer.write_all(b"<Types/>").expect("write");
            writer.start_file("xl/workbook.xml", options).expect("wb");
            writer
                .write_all(
                    br#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
                )
                .expect("write");
            writer
                .start_file("xl/_rels/workbook.xml.rels", options)
                .expect("rels");
            writer
                .write_all(
                    br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
                )
                .expect("write");
            writer
                .start_file("xl/worksheets/sheet1.xml", options)
                .expect("sheet");
            writer
                .write_all(
                    br#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#,
                )
                .expect("write");
            writer
                .start_file("customXml/item42.bin", options)
                .expect("custom");
            writer.write_all(custom_payload).expect("write");
            writer.finish().expect("finish");
        }

        let mut sink = NoopEventSink;
        let trace = TraceContext::root();
        let mut workbook = load_workbook_model(&source, &mut sink, &trace).expect("load");
        if let Some(sheet) = workbook.sheets.get_mut("Sheet1") {
            if let Some(cell) = sheet.cells.get_mut(&CellRef { row: 1, col: 1 }) {
                cell.value = CellValue::Number(99.0);
            }
        }

        let changed = vec!["Sheet1".to_string()];
        let report = preserve_xlsx_with_sheet_overrides(
            &source, &output, &workbook, &changed, &mut sink, &trace,
        )
        .expect("preserve override");
        assert!(report.copied_bytes > 0);
        assert_eq!(report.part_graph_flags.strategy, "sheet_overrides");
        assert!(report.part_graph_flags.source_graph_reused);
        assert!(report.part_graph_flags.relationships_preserved);
        assert!(report.part_graph_flags.unknown_parts_preserved);
        assert!(report.part_graph.edge_count >= 1);

        let file = File::open(&output).expect("open output");
        let reader = BufReader::new(file);
        let mut archive = ZipArchive::new(reader).expect("archive");

        {
            let mut sheet = archive
                .by_name("xl/worksheets/sheet1.xml")
                .expect("sheet xml");
            let mut sheet_data = String::new();
            sheet.read_to_string(&mut sheet_data).expect("read sheet");
            assert!(sheet_data.contains("<v>99</v>"));
        }

        let mut custom = archive
            .by_name("customXml/item42.bin")
            .expect("custom part exists");
        let mut data = Vec::new();
        custom.read_to_end(&mut data).expect("read custom");
        assert_eq!(data, custom_payload);
    }

    #[test]
    fn inspection_part_graph_flags_dangling_relationship_targets() {
        let dir = tempdir().expect("tempdir");
        let path = dir.path().join("dangling.xlsx");
        {
            let file = File::create(&path).expect("create file");
            let mut writer = zip::ZipWriter::new(file);
            let options = SimpleFileOptions::default();
            writer
                .start_file("[Content_Types].xml", options)
                .expect("content types");
            writer.write_all(b"<Types/>").expect("write");
            writer
                .start_file("_rels/.rels", options)
                .expect("root rels");
            writer
                .write_all(
                    br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
                )
                .expect("write");
            writer
                .start_file("xl/workbook.xml", options)
                .expect("workbook");
            writer
                .write_all(
                    br#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
                )
                .expect("write");
            writer
                .start_file("xl/_rels/workbook.xml.rels", options)
                .expect("workbook rels");
            writer
                .write_all(
                    br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/missing-sheet.xml"/></Relationships>"#,
                )
                .expect("write");
            writer
                .start_file("xl/worksheets/sheet1.xml", options)
                .expect("sheet");
            writer
                .write_all(
                    br#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#,
                )
                .expect("write");
            writer.finish().expect("finish");
        }

        let mut sink = NoopEventSink;
        let trace = TraceContext::root();
        let report = inspect_xlsx(&path, &mut sink, &trace).expect("inspect");
        assert!(report.part_graph.edge_count >= 2);
        assert_eq!(report.part_graph.dangling_edge_count, 1);
        assert!(report
            .part_graph
            .edges
            .iter()
            .any(|edge| edge.dangling_target));
    }

    #[test]
    fn accepts_uppercase_xlsx_extension() {
        let path = Path::new("BOOK.XLSX");
        let result = ensure_xlsx_extension(path);
        assert!(result.is_ok());
    }
}
