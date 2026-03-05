use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::fmt::{self, Display};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::str::FromStr;
use std::{env, fs, io};
use sha2::{Digest, Sha256};

const SCRIPT_API_VERSION: u32 = 1;
const SCRIPT_TRUST_MODE_ENV: &str = "ROOTCELLAR_MACRO_TRUST_MODE";
const SCRIPT_SIGNATURE_SECRET_ENV: &str = "ROOTCELLAR_MACRO_SIGNATURE_SECRET";
const SCRIPT_PUBLISHER_ALLOWLIST_ENV: &str = "ROOTCELLAR_MACRO_PUBLISHER_ALLOWLIST";
const SCRIPT_TRUST_MANIFEST_ENV: &str = "ROOTCELLAR_MACRO_MANIFEST";

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
    #[serde(rename = "udf")]
    Udf,
    #[serde(rename = "events.emit")]
    EventsEmit,
}

impl ScriptPermission {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::FsRead => "fs.read",
            Self::FsWrite => "fs.write",
            Self::NetHttp => "net.http",
            Self::Clipboard => "clipboard",
            Self::ProcessExec => "process.exec",
            Self::Udf => "udf",
            Self::EventsEmit => "events.emit",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum ScriptTrustMode {
    Legacy,
    Manifest,
    Signed,
}

impl ScriptTrustMode {
    fn as_str(self) -> &'static str {
        match self {
            Self::Legacy => "legacy",
            Self::Manifest => "manifest",
            Self::Signed => "signed",
        }
    }
}

impl std::str::FromStr for ScriptTrustMode {
    type Err = String;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value.to_lowercase().as_str() {
            "" | "legacy" | "off" | "fingerprint" | "allowlist" => Ok(Self::Legacy),
            "manifest" | "manifest-only" => Ok(Self::Manifest),
            "signed" | "signed-only" => Ok(Self::Signed),
            _ => Err(format!("unsupported macro trust mode: {value}")),
        }
    }
}

const SCRIPT_TRUST_FINGERPRINT_ALLOWLIST_ENV: &str = "ROOTCELLAR_MACRO_FINGERPRINT_ALLOWLIST";

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScriptTrustProvenance {
    pub mode: String,
    pub manifest_path: Option<String>,
    pub manifest_name: Option<String>,
    pub manifest_version: Option<String>,
    pub publisher: Option<String>,
    pub api_min_version: Option<u32>,
    pub permissions_required: Vec<String>,
    pub permissions_declared: Vec<String>,
    pub runtime_api_version: u32,
    pub signature_present: bool,
    pub signature_verified: Option<bool>,
    pub fingerprint: String,
    pub trusted: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScriptManifest {
    pub name: String,
    pub version: String,
    pub publisher: String,
    pub permissions: Vec<String>,
    pub api_min_version: u32,
    pub signature: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScriptRuntimeEvent {
    pub event_name: String,
    pub payload: serde_json::Value,
    pub severity: Option<String>,
}

fn script_fingerprint(path: &Path) -> Result<String, ScriptError> {
    let bytes = fs::read(path).map_err(|error| {
        ScriptError::Io(io::Error::new(
            io::ErrorKind::Other,
            format!("failed to read macro script for fingerprinting: {error}"),
        ))
    })?;

    let mut hash: u64 = 0xcbf29ce484222325;
    for byte in bytes {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    Ok(format!("{hash:016x}"))
}

fn is_truthy(value: &str) -> bool {
    matches!(
        value.trim().to_lowercase().as_str(),
        "1" | "true" | "yes" | "on"
    )
}

fn parse_trust_allowlist(raw: &str) -> (bool, Vec<String>) {
    raw.split(|c| c == ',' || c == ';')
        .map(|entry| entry.trim())
        .filter(|entry| !entry.is_empty())
        .fold(
            (false, Vec::new()),
            |(has_truthy, mut fingerprints), entry| {
                if is_truthy(entry) {
                    (true, fingerprints)
                } else {
                    fingerprints.push(entry.to_string());
                    (has_truthy, fingerprints)
                }
            },
        )
}

fn canonicalized_permissions(permissions: &[ScriptPermission]) -> Vec<String> {
    let mut values = permissions.iter().map(|permission| permission.as_str().to_string()).collect::<Vec<_>>();
    values.sort_unstable();
    values
}

fn resolve_trust_mode() -> Result<ScriptTrustMode, ScriptError> {
    let raw_mode = env::var(SCRIPT_TRUST_MODE_ENV).unwrap_or_else(|_| "legacy".to_string());
    ScriptTrustMode::from_str(&raw_mode).map_err(|error| ScriptError::ExecutionFailed {
        message: error,
        status: Some("invalid_macro_trust_mode".to_string()),
        permission_events: Vec::new(),
        stdout: None,
        stderr: None,
    })
}

fn resolve_manifest_path(script_path: &Path) -> Option<PathBuf> {
    if let Ok(manifest_env) = env::var(SCRIPT_TRUST_MANIFEST_ENV) {
        let trimmed = manifest_env.trim();
        if !trimmed.is_empty() {
            return Some(PathBuf::from(trimmed));
        }
    }

    let sidecar = script_path.with_extension("macro.json");
    if sidecar.exists() {
        return Some(sidecar);
    }

    None
}

fn load_script_manifest(manifest_path: &Path) -> Result<ScriptManifest, ScriptError> {
    if !manifest_path.exists() {
        return Err(ScriptError::ExecutionFailed {
            message: format!(
                "macro trust manifest not found: {}",
                manifest_path.display()
            ),
            status: Some("macro_manifest_missing".to_string()),
            permission_events: Vec::new(),
            stdout: None,
            stderr: None,
        });
    }

    let manifest_content = fs::read_to_string(manifest_path).map_err(|error| {
        ScriptError::ExecutionFailed {
            message: format!(
                "failed to read macro manifest {}: {error}",
                manifest_path.display()
            ),
            status: Some("macro_manifest_read_failed".to_string()),
            permission_events: Vec::new(),
            stdout: None,
            stderr: None,
        }
    })?;
    let manifest: ScriptManifest = serde_json::from_str(&manifest_content).map_err(|error| {
        ScriptError::InvalidResponse(format!("invalid macro manifest JSON: {error}"))
    })?;
    Ok(manifest)
}

fn compute_manifest_signature(manifest: &ScriptManifest, fingerprint: &str, secret: &str) -> String {
    let mut hasher = Sha256::new();
    let mut normalized_permissions = manifest.permissions.clone();
    normalized_permissions.sort_unstable();

    hasher.update(secret.as_bytes());
    hasher.update(fingerprint.as_bytes());
    hasher.update(manifest.name.as_bytes());
    hasher.update(manifest.version.as_bytes());
    hasher.update(manifest.publisher.as_bytes());
    hasher.update(manifest.api_min_version.to_string().as_bytes());
    hasher.update(manifest.permissions.len().to_string().as_bytes());
    for permission in normalized_permissions {
        hasher.update(permission.as_bytes());
    }
    hasher
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>()
}

fn validate_manifest_permissions(
    manifest: &ScriptManifest,
    requested_permissions: &[ScriptPermission],
) -> Result<(), ScriptError> {
    let mut granted = manifest
        .permissions
        .iter()
        .map(|permission| permission.trim().to_lowercase())
        .collect::<Vec<_>>();
    granted.sort_unstable();
    for requested in requested_permissions {
        if !granted.contains(&requested.as_str().to_string()) {
            return Err(ScriptError::ExecutionFailed {
                message: format!(
                    "requested permission is not declared in manifest: {}",
                    requested.as_str()
                ),
                status: Some("manifest_permission_violation".to_string()),
                permission_events: Vec::new(),
                stdout: None,
                stderr: None,
            });
        }
    }
    Ok(())
}

fn verify_publisher_allowlist(publisher: &str) -> bool {
    let allowlist = env::var(SCRIPT_PUBLISHER_ALLOWLIST_ENV).unwrap_or_default();
    if allowlist.trim().is_empty() {
        return true;
    }
    let (_, publishers) = parse_trust_allowlist(&allowlist);
    publishers
        .iter()
        .any(|entry| entry.eq_ignore_ascii_case(&publisher.trim().to_lowercase()))
}

fn enforce_script_trust(
    script_path: &str,
    permissions: &[ScriptPermission],
    fingerprint: &str,
) -> Result<ScriptTrustProvenance, ScriptError> {
    let mode = resolve_trust_mode()?;
    let requested_permissions = canonicalized_permissions(permissions);

    match mode {
        ScriptTrustMode::Legacy => {
            let allowlist = env::var(SCRIPT_TRUST_FINGERPRINT_ALLOWLIST_ENV).unwrap_or_default();
            if allowlist.trim().is_empty() {
                return Ok(ScriptTrustProvenance {
                    mode: mode.as_str().to_string(),
                    manifest_path: None,
                    manifest_name: None,
                    manifest_version: None,
                    publisher: None,
                    api_min_version: None,
                    permissions_required: requested_permissions,
                    permissions_declared: Vec::new(),
                    runtime_api_version: SCRIPT_API_VERSION,
                    signature_present: false,
                    signature_verified: None,
                    fingerprint: fingerprint.to_string(),
                    trusted: true,
                });
            }

            let (allowlist_allows_all, fingerprint_allowlist) = parse_trust_allowlist(&allowlist);
            if allowlist_allows_all || fingerprint_allowlist.iter().any(|entry| entry == fingerprint) {
                return Ok(ScriptTrustProvenance {
                    mode: "legacy".to_string(),
                    manifest_path: None,
                    manifest_name: None,
                    manifest_version: None,
                    publisher: None,
                    api_min_version: None,
                    permissions_required: requested_permissions,
                    permissions_declared: Vec::new(),
                    runtime_api_version: SCRIPT_API_VERSION,
                    signature_present: false,
                    signature_verified: None,
                    fingerprint: fingerprint.to_string(),
                    trusted: true,
                });
            }

            Err(ScriptError::ExecutionFailed {
                message: format!("macro script is not trusted: {script_path} (fingerprint {fingerprint})"),
                status: Some("macro_not_trusted".to_string()),
                permission_events: Vec::new(),
                stdout: None,
                stderr: None,
            })
        }
        ScriptTrustMode::Manifest | ScriptTrustMode::Signed => {
            let manifest_path = resolve_manifest_path(Path::new(script_path)).ok_or_else(|| {
                ScriptError::ExecutionFailed {
                    message: format!("missing macro trust manifest for: {script_path}"),
                    status: Some("macro_manifest_missing".to_string()),
                    permission_events: Vec::new(),
                    stdout: None,
                    stderr: None,
                }
            })?;
            let manifest = load_script_manifest(manifest_path.as_path())?;
            if manifest.api_min_version > SCRIPT_API_VERSION {
                return Err(ScriptError::ExecutionFailed {
                    message: format!(
                        "macro API requires version {} but host is {}",
                        manifest.api_min_version, SCRIPT_API_VERSION
                    ),
                    status: Some("incompatible_macro_api_version".to_string()),
                    permission_events: Vec::new(),
                    stdout: None,
                    stderr: None,
                });
            }

            if !verify_publisher_allowlist(&manifest.publisher) {
                return Err(ScriptError::ExecutionFailed {
                    message: format!("macro publisher is not allow-listed: {}", manifest.publisher),
                    status: Some("macro_publisher_not_allowlisted".to_string()),
                    permission_events: Vec::new(),
                    stdout: None,
                    stderr: None,
                });
            }

            validate_manifest_permissions(&manifest, permissions)?;

            let signature = manifest.signature.clone();
            let signature_present = signature.as_deref().is_some();
            let signature_verified = match (mode, signature.clone()) {
                (ScriptTrustMode::Manifest, Some(signature)) => {
                    let secret = env::var(SCRIPT_SIGNATURE_SECRET_ENV).unwrap_or_default();
                    if secret.trim().is_empty() {
                        None
                    } else {
                        Some(signature == compute_manifest_signature(&manifest, fingerprint, &secret))
                    }
                }
                (ScriptTrustMode::Signed, Some(signature)) => {
                    let secret = env::var(SCRIPT_SIGNATURE_SECRET_ENV).unwrap_or_default();
                    if secret.trim().is_empty() {
                        return Err(ScriptError::ExecutionFailed {
                            message: "missing ROOTCELLAR_MACRO_SIGNATURE_SECRET for signed mode".to_string(),
                            status: Some("macro_signature_secret_missing".to_string()),
                            permission_events: Vec::new(),
                            stdout: None,
                            stderr: None,
                        });
                    }
                    Some(signature == compute_manifest_signature(&manifest, fingerprint, &secret))
                }
                (ScriptTrustMode::Signed, None) => {
                    return Err(ScriptError::ExecutionFailed {
                        message: "macro manifest signature required in signed mode".to_string(),
                        status: Some("macro_signature_missing".to_string()),
                        permission_events: Vec::new(),
                        stdout: None,
                        stderr: None,
                    });
                }
                _ => None,
            };

            if mode == ScriptTrustMode::Signed {
                if !signature_verified.unwrap_or(false) {
                    return Err(ScriptError::ExecutionFailed {
                        message: "macro manifest signature verification failed".to_string(),
                        status: Some("macro_signature_invalid".to_string()),
                        permission_events: Vec::new(),
                        stdout: None,
                        stderr: None,
                    });
                }
            }

            Ok(ScriptTrustProvenance {
                mode: mode.as_str().to_string(),
                manifest_path: Some(manifest_path.to_string_lossy().to_string()),
                manifest_name: Some(manifest.name),
                manifest_version: Some(manifest.version),
                publisher: Some(manifest.publisher),
                api_min_version: Some(manifest.api_min_version),
                permissions_required: requested_permissions,
                permissions_declared: canonicalized_permissions(
                    &manifest
                        .permissions
                        .iter()
                        .filter_map(|value| ScriptPermission::from_str(value).ok())
                        .collect::<Vec<_>>(),
                ),
                runtime_api_version: SCRIPT_API_VERSION,
                signature_present,
                signature_verified,
                fingerprint: fingerprint.to_string(),
                trusted: true,
            })
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
            "udf" => Ok(Self::Udf),
            "events.emit" => Ok(Self::EventsEmit),
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
    pub script_fingerprint: Option<String>,
    #[serde(default)]
    pub trust: Option<ScriptTrustProvenance>,
    #[serde(default)]
    pub runtime_events: Vec<ScriptRuntimeEvent>,
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
    let script_path = Path::new(&request.script_path);
    let fingerprint = script_fingerprint(script_path)?;
    let trust_provenance = enforce_script_trust(&request.script_path, &request.permissions, &fingerprint)?;

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
    response.trust = response.trust.or(Some(trust_provenance));
    if response.runtime_events.is_empty() {
        response.runtime_events = Vec::new();
    }
    if response.script_fingerprint.is_none() {
        response.script_fingerprint = Some(fingerprint);
    } else if response
        .script_fingerprint
        .as_deref()
        .is_some_and(|value| value != fingerprint)
    {
        return Err(ScriptError::ExecutionFailed {
            message: format!(
                "script-provided fingerprint mismatch for {}: declared={} actual={}",
                request.script_path,
                response.script_fingerprint.as_deref().unwrap_or_default(),
                fingerprint
            ),
            status: Some("fingerprint_mismatch".to_string()),
            permission_events: response.permission_events,
            stdout: response.stdout,
            stderr: response.stderr,
        });
    } else {
        response.script_fingerprint = Some(fingerprint);
    }

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

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs::write;
    use std::sync::{Mutex, OnceLock};
    use std::time::{SystemTime, UNIX_EPOCH};

    static TEST_ENV_MUTEX: OnceLock<Mutex<()>> = OnceLock::new();

    fn with_trust_env<'a, T, F>(vars: &[(&'a str, Option<&'a str>)], action: F) -> T
    where
        F: FnOnce() -> T,
    {
        let _guard = TEST_ENV_MUTEX
            .get_or_init(|| Mutex::new(()))
            .lock()
            .expect("test env lock");

        let mut original: Vec<(String, Option<String>)> = vars
            .iter()
            .map(|(name, value)| {
                let previous = env::var(name).ok();
                match value {
                    Some(value) => env::set_var(name, value),
                    None => env::remove_var(name),
                }
                (name.to_string(), previous)
            })
            .collect();

        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(action));

        original.reverse();
        for (key, value) in original {
            match value {
                Some(value) => env::set_var(&key, value),
                None => env::remove_var(&key),
            }
        }

        match result {
            Ok(value) => value,
            Err(payload) => std::panic::resume_unwind(payload),
        }
    }

    fn with_allowlist_only<T, F>(value: Option<&str>, action: F) -> T
    where
        F: FnOnce() -> T,
    {
        with_trust_env(&[(SCRIPT_TRUST_FINGERPRINT_ALLOWLIST_ENV, value)], action)
    }

    fn write_manifest(path: &Path, manifest: &ScriptManifest, signature: Option<&str>) {
        let mut manifest = manifest.clone();
        manifest.signature = signature.map(ToString::to_string);
        let serialized = serde_json::to_string_pretty(&manifest).expect("serialize macro manifest");
        write(path, serialized).expect("write temporary macro manifest");
    }

    fn write_macro_with_manifest(
        script_path: &Path,
        manifest_path: &Path,
        permissions: &[&str],
        publisher: &str,
        signature: Option<&str>,
        secret: Option<&str>,
    ) -> String {
        if let Some(secret) = secret {
            let _ = secret;
        }

        write_script(
            script_path,
            "def run_macro(ctx, args):\n    return {'ok': True}\n",
        );
        let fingerprint = script_fingerprint(script_path).expect("compute fingerprint for manifest");
        let manifest = ScriptManifest {
            name: "macro-fixture".to_string(),
            version: "1.0.0".to_string(),
            publisher: publisher.to_string(),
            permissions: permissions.iter().map(|item| item.to_string()).collect(),
            api_min_version: SCRIPT_API_VERSION,
            signature: None,
        };
        let manifest_signature = match signature {
            Some(value) => Some(value.to_string()),
            None => secret.map(|secret| compute_manifest_signature(&manifest, &fingerprint, secret)),
        };
        write_manifest(manifest_path, &manifest, manifest_signature.as_deref());
        fingerprint
    }

    fn with_script_trust_mode<T, F>(trust_mode: &str, action: F) -> T
    where
        F: FnOnce() -> T,
    {
        with_trust_env(
            &[
                (SCRIPT_TRUST_MODE_ENV, Some(trust_mode)),
                (SCRIPT_TRUST_FINGERPRINT_ALLOWLIST_ENV, None),
            ],
            action,
        )
    }

    fn temp_script_path(test_name: &str) -> PathBuf {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|clock| clock.as_nanos())
            .unwrap_or(0);
        std::env::temp_dir().join(format!(
            "rootcellar-macro-fingerprint-{test_name}-{suffix}.py"
        ))
    }

    fn write_script(path: &Path, body: &str) {
        write(path, body).expect("write fixture script");
    }

    #[test]
    fn script_fingerprint_is_deterministic() {
        let path = temp_script_path("deterministic");
        write_script(&path, "print('macro')\n");
        let first = script_fingerprint(&path).expect("compute fingerprint");
        let second = script_fingerprint(&path).expect("recompute fingerprint");
        assert_eq!(first, second);
        assert_eq!(first.len(), 16);
        std::fs::remove_file(path).expect("cleanup script");
    }

    #[test]
    fn enforce_script_trust_no_allowlist_skips_validation() {
        let path = temp_script_path("no_allowlist");
        write_script(&path, "print('macro')\n");
        let fingerprint = script_fingerprint(&path).expect("compute fingerprint");

        with_trust_env(&[], || {
            assert!(
                enforce_script_trust(
                    path.to_str().expect("path string"),
                    &[],
                    &fingerprint,
                )
                    .is_ok(),
            );
        });
        assert!(
            with_trust_env(&[(SCRIPT_TRUST_FINGERPRINT_ALLOWLIST_ENV, Some("")),], || {
                enforce_script_trust(
                    path.to_str().expect("path string"),
                    &[],
                    &fingerprint,
                )
                .is_ok()
            })
        );
        std::fs::remove_file(path).expect("cleanup script");
    }

    #[test]
    fn enforce_script_trust_accepts_matching_or_truthy_allowlist() {
        let path = temp_script_path("allowlist_match");
        write_script(&path, "print('macro')\n");
        let fingerprint = script_fingerprint(&path).expect("compute fingerprint");

        with_allowlist_only(Some("ff00bad,  1 , other"), || {
            assert!(
                enforce_script_trust(
                    path.to_str().expect("path string"),
                    &[],
                    &fingerprint,
                )
                .is_ok()
            );
        });
        with_allowlist_only(Some("on"), || {
            assert!(
                enforce_script_trust(
                    path.to_str().expect("path string"),
                    &[],
                    &fingerprint,
                )
                .is_ok()
            );
        });
        std::fs::remove_file(path).expect("cleanup script");
    }

    #[test]
    fn enforce_script_trust_rejects_untrusted_fingerprint() {
        let path = temp_script_path("allowlist_reject");
        write_script(&path, "print('macro')\n");
        let _ = script_fingerprint(&path).expect("compute fingerprint");

        let error = with_allowlist_only(Some("ff00bad,deadbeef"), || {
            enforce_script_trust(path.to_str().expect("path string"), &[], "not-present")
        })
        .expect_err("untrusted fingerprint should be rejected");
        assert!(error.to_string().contains("not trusted"));
        std::fs::remove_file(path).expect("cleanup script");
    }

    #[test]
    fn enforce_script_trust_accepts_manifest_mode_with_permissions_and_publisher() {
        let script_path = temp_script_path("manifest_ok");
        let manifest_path = script_path.with_extension("macro.json");
        let fingerprint = write_macro_with_manifest(
            &script_path,
            &manifest_path,
            &["fs.read", "fs.write"],
            "trusted-author",
            None,
            None,
        );
        let permissions = [ScriptPermission::FsRead, ScriptPermission::FsWrite];

        with_script_trust_mode("manifest", || {
            let result = enforce_script_trust(
                script_path.to_str().expect("path string"),
                &permissions,
                &fingerprint,
            );
            let provenance = result.expect("manifest trust should pass");
            assert_eq!(provenance.mode, "manifest");
            assert!(provenance.trusted);
            assert!(!provenance.signature_present);
            assert!(provenance.signature_verified.is_none());
            assert_eq!(provenance.publisher.as_deref(), Some("trusted-author"));
            assert_eq!(provenance.permissions_required, vec!["fs.read", "fs.write"]);
            assert_eq!(provenance.permissions_declared, vec!["fs.read", "fs.write"]);
        });

        std::fs::remove_file(script_path).expect("cleanup script");
        std::fs::remove_file(manifest_path).expect("cleanup manifest");
    }

    #[test]
    fn enforce_script_trust_manifest_mode_rejects_unknown_publisher() {
        let script_path = temp_script_path("manifest_unknown_publisher");
        let manifest_path = script_path.with_extension("macro.json");
        let _fingerprint = write_macro_with_manifest(
            &script_path,
            &manifest_path,
            &["fs.read"],
            "untrusted-author",
            None,
            None,
        );

        with_script_trust_mode("manifest", || {
            with_trust_env(&[(SCRIPT_PUBLISHER_ALLOWLIST_ENV, Some("trusted-author"))], || {
                let error = enforce_script_trust(
                    script_path.to_str().expect("path string"),
                    &[ScriptPermission::FsRead],
                    "fingerprint",
                )
                .expect_err("manifest with unknown publisher should fail");
                assert!(error.to_string().contains("publisher is not allow-listed"));
            });
        });

        std::fs::remove_file(script_path).expect("cleanup script");
        std::fs::remove_file(manifest_path).expect("cleanup manifest");
    }

    #[test]
    fn enforce_script_trust_signed_mode_rejects_invalid_signature() {
        let script_path = temp_script_path("signed_invalid");
        let manifest_path = script_path.with_extension("macro.json");
        let fingerprint = write_macro_with_manifest(
            &script_path,
            &manifest_path,
            &["fs.read"],
            "trusted-author",
            Some("deadbeef"),
            None,
        );

        with_script_trust_mode("signed", || {
            with_trust_env(&[(SCRIPT_SIGNATURE_SECRET_ENV, Some("topsecret"))], || {
                let error = enforce_script_trust(
                    script_path.to_str().expect("path string"),
                    &[ScriptPermission::FsRead],
                    &fingerprint,
                )
                .expect_err("signed mode should reject invalid signature");
                assert!(error
                    .to_string()
                    .contains("macro manifest signature verification failed"));
            });
        });

        std::fs::remove_file(script_path).expect("cleanup script");
        std::fs::remove_file(manifest_path).expect("cleanup manifest");
    }

    #[test]
    fn enforce_script_trust_signed_mode_rejects_missing_secret() {
        let script_path = temp_script_path("signed_missing_secret");
        let manifest_path = script_path.with_extension("macro.json");
        let fingerprint = write_macro_with_manifest(
            &script_path,
            &manifest_path,
            &["fs.read"],
            "trusted-author",
            None,
            Some("topsecret"),
        );

        with_script_trust_mode("signed", || {
            let error = enforce_script_trust(
                script_path.to_str().expect("path string"),
                &[ScriptPermission::FsRead],
                &fingerprint,
            )
            .expect_err("signed mode should require ROOTCELLAR_MACRO_SIGNATURE_SECRET");
            assert!(error.to_string().contains("ROOTCELLAR_MACRO_SIGNATURE_SECRET"));
        });

        std::fs::remove_file(script_path).expect("cleanup script");
        std::fs::remove_file(manifest_path).expect("cleanup manifest");
    }

    #[test]
    fn enforce_script_trust_signed_mode_accepts_valid_signature() {
        let script_path = temp_script_path("signed_valid");
        let manifest_path = script_path.with_extension("macro.json");
        let secret = "topsecret";
        let fingerprint = write_macro_with_manifest(
            &script_path,
            &manifest_path,
            &["fs.read"],
            "trusted-author",
            None,
            Some(secret),
        );

        with_script_trust_mode("signed", || {
            with_trust_env(&[(SCRIPT_SIGNATURE_SECRET_ENV, Some(secret))], || {
                let provenance = enforce_script_trust(
                    script_path.to_str().expect("path string"),
                    &[ScriptPermission::FsRead],
                    &fingerprint,
                )
                .expect("valid signed manifest should pass");
                assert_eq!(provenance.mode, "signed");
                assert!(provenance.trusted);
                assert!(provenance.signature_verified.unwrap_or(false));
            });
        });

        std::fs::remove_file(script_path).expect("cleanup script");
        std::fs::remove_file(manifest_path).expect("cleanup manifest");
    }
}
