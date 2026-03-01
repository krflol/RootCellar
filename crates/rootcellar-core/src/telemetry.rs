use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::Path;
use thiserror::Error;
use uuid::Uuid;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum Severity {
    Debug,
    Info,
    Warn,
    Error,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TraceContext {
    pub trace_id: Uuid,
    pub span_id: Uuid,
    pub parent_span_id: Option<Uuid>,
    pub session_id: Option<Uuid>,
}

impl TraceContext {
    pub fn root() -> Self {
        Self {
            trace_id: Uuid::now_v7(),
            span_id: Uuid::now_v7(),
            parent_span_id: None,
            session_id: Some(Uuid::now_v7()),
        }
    }

    pub fn child(&self) -> Self {
        Self {
            trace_id: self.trace_id,
            span_id: Uuid::now_v7(),
            parent_span_id: Some(self.span_id),
            session_id: self.session_id,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EventEnvelope {
    pub event_name: String,
    pub event_version: String,
    pub timestamp: DateTime<Utc>,
    pub severity: Severity,
    pub trace_id: Uuid,
    pub span_id: Uuid,
    pub parent_span_id: Option<Uuid>,
    pub session_id: Option<Uuid>,
    pub workbook_id: Option<Uuid>,
    pub txn_id: Option<Uuid>,
    pub context: Value,
    pub metrics: Value,
    pub payload: Value,
}

impl EventEnvelope {
    pub fn info(event_name: impl Into<String>, trace: &TraceContext) -> Self {
        Self {
            event_name: event_name.into(),
            event_version: "1.0.0".to_string(),
            timestamp: Utc::now(),
            severity: Severity::Info,
            trace_id: trace.trace_id,
            span_id: trace.span_id,
            parent_span_id: trace.parent_span_id,
            session_id: trace.session_id,
            workbook_id: None,
            txn_id: None,
            context: json!({}),
            metrics: json!({}),
            payload: json!({}),
        }
    }

    pub fn with_workbook_id(mut self, workbook_id: Uuid) -> Self {
        self.workbook_id = Some(workbook_id);
        self
    }

    pub fn with_txn_id(mut self, txn_id: Uuid) -> Self {
        self.txn_id = Some(txn_id);
        self
    }

    pub fn with_context(mut self, context: Value) -> Self {
        self.context = context;
        self
    }

    pub fn with_metrics(mut self, metrics: Value) -> Self {
        self.metrics = metrics;
        self
    }

    pub fn with_payload(mut self, payload: Value) -> Self {
        self.payload = payload;
        self
    }
}

#[derive(Debug, Error)]
pub enum TelemetryError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("serialization error: {0}")]
    Serialization(#[from] serde_json::Error),
}

pub trait EventSink {
    fn emit(&mut self, event: EventEnvelope) -> Result<(), TelemetryError>;

    fn supports_expensive_payloads(&self) -> bool {
        true
    }
}

#[derive(Default)]
pub struct NoopEventSink;

impl EventSink for NoopEventSink {
    fn emit(&mut self, _event: EventEnvelope) -> Result<(), TelemetryError> {
        Ok(())
    }

    fn supports_expensive_payloads(&self) -> bool {
        false
    }
}

pub struct JsonlEventSink {
    writer: BufWriter<File>,
}

impl JsonlEventSink {
    pub fn new(path: impl AsRef<Path>) -> Result<Self, TelemetryError> {
        let file = File::create(path)?;
        Ok(Self {
            writer: BufWriter::new(file),
        })
    }
}

impl EventSink for JsonlEventSink {
    fn emit(&mut self, event: EventEnvelope) -> Result<(), TelemetryError> {
        serde_json::to_writer(&mut self.writer, &event)?;
        self.writer.write_all(b"\n")?;
        self.writer.flush()?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn writes_jsonl_event() {
        let dir = tempdir().expect("tempdir");
        let log_path = dir.path().join("events.jsonl");
        let mut sink = JsonlEventSink::new(&log_path).expect("sink");
        let trace = TraceContext::root();

        sink.emit(EventEnvelope::info("engine.txn.begin", &trace))
            .expect("emit");

        let data = std::fs::read_to_string(log_path).expect("read");
        assert!(data.contains("engine.txn.begin"));
    }

    #[test]
    fn noop_sink_disables_expensive_payloads() {
        let sink = NoopEventSink;
        assert!(!sink.supports_expensive_payloads());
    }
}
