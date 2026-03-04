use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::fmt::{self, Display};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::str::FromStr;
use std::{env, fs, io};

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum ScriptPermission {
    #[serde(rename = "fs.read")]
    FsRead,
    #[serde(rename = "fs.write")]
    FsWrite,
    #[serde(rename = "net.http")]
    NetHttp,
    #[serde(rename = "clipboard")]
    Clipboard,
    #[serde(rename = "process.exec")]
    ProcessExec,
}

impl ScriptPermission {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::FsRead => "fs.read",
            Self::FsWrite => "fs.write",
            Self::NetHttp => "net.http",
            Self::Clipboard => "clipboard",
            Self::ProcessExec => "process.exec",
        }
    }
}

impl Display for ScriptPermission {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

impl FromStr for ScriptPermission {
    type Err = String;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "fs.read" => Ok(Self::FsRead),
            "fs.write" => Ok(Self::FsWrite),
            "net.http" => Ok(Self::NetHttp),
            "clipboard" => Ok(Self::Clipboard),
            "process.exec" => Ok(Self::ProcessExec),
            _ => Err(format!("unsupported permission: {value}")),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MacroRunRequest {
    pub command: String,
    pub trace_id: String,
    pub script_path: String,
    pub macro_name: String,
    pub workbook_path: String,
    pub permissions: Vec<ScriptPermission>,
    pub args: BTreeMap<String, String>,
}

impl MacroRunRequest {
    pub fn new(
        trace_id: String,
        script_path: &Path,
        macro_name: String,
        workbook_path: &Path,
        permissions: Vec<ScriptPermission>,
        args: BTreeMap<String, String>,
    ) -> Self {
        Self {
            command: "macro.run".to_string(),
            trace_id,
            script_path: script_path.to_string_lossy().to_string(),
            macro_name,
            workbook_path: workbook_path.to_string_lossy().to_string(),
            permissions,
            args,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", content = "value", rename_all = "lowercase")]
pub enum ScriptCellValue {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
    Empty,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "op", rename_all = "snake_case")]
pub enum ScriptMutation {
    SetCellValue {
        sheet: String,
        cell: String,
        value: ScriptCellValue,
    },
    SetCellFormula {
        sheet: String,
        cell: String,
        formula: String,
    },
    SetCellRangeValue {
        sheet: String,
        start: String,
        end: String,
        value: ScriptCellValue,
    },
    SetCellRangeFormula {
        sheet: String,
        start: String,
        end: String,
        formula: String,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScriptPermissionEvent {
    pub event_name: String,
    pub permission: String,
    pub allowed: bool,
    pub reason: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScriptRunResponse {
    pub status: String,
    pub message: Option<String>,
    pub stdout: Option<String>,
    pub stderr: Option<String>,
    pub permission_events: Vec<ScriptPermissionEvent>,
    pub mutations: Vec<ScriptMutation>,
    pub result: Option<serde_json::Value>,
}

#[derive(Debug, thiserror::Error)]
pub enum ScriptError {
    #[error("python invocation failed: {message}")]
    ExecutionFailed {
        message: String,
        status: Option<String>,
        permission_events: Vec<ScriptPermissionEvent>,
        stdout: Option<String>,
        stderr: Option<String>,
    },
    #[error("worker process missing: {0}")]
    WorkerMissing(String),
    #[error("worker transport failed: {0}")]
    Transport(String),
    #[error("invalid worker response: {0}")]
    InvalidResponse(String),
    #[error("io error: {0}")]
    Io(#[from] io::Error),
    #[error("serde error: {0}")]
    Serde(#[from] serde_json::Error),
}

fn resolve_python_binary() -> String {
    env::var("ROOTCELLAR_PYTHON")
        .ok()
        .filter(|value| !value.trim().is_empty())
        .unwrap_or_else(|| "python".to_string())
}

fn resolve_worker_path() -> Result<PathBuf, ScriptError> {
    if let Ok(explicit) = env::var("ROOTCELLAR_SCRIPT_WORKER") {
        let explicit_path = PathBuf::from(explicit);
        if explicit_path.exists() {
            return Ok(explicit_path);
        }
        return Err(ScriptError::WorkerMissing(format!(
            "explicit worker path does not exist: {}",
            explicit_path.display()
        )));
    }

    let candidate = PathBuf::from("python").join("worker_stub.py");
    if candidate.exists() {
        return Ok(candidate);
    }

    // Fallback for tests or non-root working directories.
    let cwd = env::current_dir()?;
    let candidate = cwd.join("python").join("worker_stub.py");
    if candidate.exists() {
        return Ok(candidate);
    }

    // Allow local developer runs from nested directories with explicit workspace root.
    if let Ok(workspace_root) = env::var("ROOTCELLAR_WORKSPACE") {
        let candidate = PathBuf::from(workspace_root)
            .join("python")
            .join("worker_stub.py");
        if candidate.exists() {
            return Ok(candidate);
        }
    }

    Err(ScriptError::WorkerMissing(
        "could not locate python/worker_stub.py; set ROOTCELLAR_SCRIPT_WORKER or run from repo root"
            .to_string(),
    ))
}

fn parse_worker_response(payload: &str) -> Result<ScriptRunResponse, ScriptError> {
    let mut text = payload
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty());
    let mut last_json: Option<&str> = None;
    for line in text.by_ref() {
        if line.starts_with('{') {
            last_json = Some(line);
        }
    }

    let line = last_json.ok_or_else(|| {
        ScriptError::InvalidResponse("worker produced no JSON response payload".to_string())
    })?;
    serde_json::from_str(line).map_err(|error| {
        ScriptError::InvalidResponse(format!("cannot parse worker response: {error}"))
    })
}

pub fn run_macro(request: &MacroRunRequest) -> Result<ScriptRunResponse, ScriptError> {
    let request_json = serde_json::to_string(request)?;
    let python = resolve_python_binary();
    let worker = resolve_worker_path()?;
    let _ = fs::metadata(&worker).map_err(ScriptError::Io)?;

    let mut child = Command::new(python)
        .arg(worker)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|error| ScriptError::ExecutionFailed {
            message: error.to_string(),
            status: None,
            permission_events: Vec::new(),
            stdout: None,
            stderr: None,
        })?;

    if let Some(mut stdin) = child.stdin.take() {
        stdin.write_all(request_json.as_bytes())?;
        stdin.write_all(b"\n")?;
    } else {
        return Err(ScriptError::Transport(
            "failed to open worker stdin pipe".to_string(),
        ));
    }

    let output = child
        .wait_with_output()
        .map_err(|error| ScriptError::Transport(error.to_string()))?;

    let stdout = String::from_utf8(output.stdout).map_err(|error| {
        ScriptError::Transport(format!("invalid UTF-8 output from worker: {error}"))
    })?;
    let stderr = String::from_utf8(output.stderr).map_err(|error| {
        ScriptError::Transport(format!("invalid UTF-8 error output from worker: {error}"))
    })?;

    let mut response = parse_worker_response(&stdout)?;
    if response.stdout.is_none() {
        response.stdout = Some(stdout.trim().to_string());
    }
    if response.stderr.is_none() && !stderr.trim().is_empty() {
        response.stderr = Some(stderr.clone());
    }

    if !output.status.success() {
        return Err(ScriptError::ExecutionFailed {
            status: Some(output.status.to_string()),
            message: format!(
                "worker process exited with status {} and reported status '{}'",
                output.status, response.status
            ),
            permission_events: response.permission_events,
            stdout: response.stdout,
            stderr: response.stderr,
        });
    }

    if response.status.to_lowercase() != "ok" {
        return Err(ScriptError::ExecutionFailed {
            status: Some(response.status.clone()),
            message: response
                .message
                .unwrap_or_else(|| "script execution returned non-ok status".to_string()),
            permission_events: response.permission_events,
            stdout: response.stdout,
            stderr: response.stderr,
        });
    }

    Ok(response)
}
