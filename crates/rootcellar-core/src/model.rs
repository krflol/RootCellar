use crate::telemetry::{EventEnvelope, EventSink, TelemetryError, TraceContext};
use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use serde_json::json;
use std::collections::{BTreeMap, BTreeSet};
use thiserror::Error;
use uuid::Uuid;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum CellValue {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
    Empty,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct CellRef {
    pub row: u32,
    pub col: u32,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Mutation {
    SetCellValue {
        sheet: String,
        row: u32,
        col: u32,
        value: CellValue,
    },
    SetCellFormula {
        sheet: String,
        row: u32,
        col: u32,
        formula: String,
        cached_value: CellValue,
    },
}

#[derive(Debug, Error)]
pub enum ModelError {
    #[error("invalid cell address row={row}, col={col}; both must be >= 1")]
    InvalidCellAddress { row: u32, col: u32 },
    #[error("telemetry error: {0}")]
    Telemetry(#[from] TelemetryError),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellRecord {
    pub value: CellValue,
    pub formula: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    pub name: String,
    pub cells: BTreeMap<CellRef, CellRecord>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Workbook {
    pub workbook_id: Uuid,
    pub created_at: DateTime<Utc>,
    pub sheets: BTreeMap<String, Sheet>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Transaction {
    pub txn_id: Uuid,
    pub workbook_id: Uuid,
    pub created_at: DateTime<Utc>,
    pub mutations: Vec<Mutation>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommitResult {
    pub txn_id: Uuid,
    pub workbook_id: Uuid,
    pub committed_at: DateTime<Utc>,
    pub mutation_count: usize,
    pub changed_cells: BTreeMap<String, Vec<CellRef>>,
}

impl Workbook {
    pub fn new() -> Self {
        Self {
            workbook_id: Uuid::now_v7(),
            created_at: Utc::now(),
            sheets: BTreeMap::new(),
        }
    }

    pub fn begin_txn(
        &self,
        sink: &mut dyn EventSink,
        trace: &TraceContext,
    ) -> Result<Transaction, ModelError> {
        let txn = Transaction {
            txn_id: Uuid::now_v7(),
            workbook_id: self.workbook_id,
            created_at: Utc::now(),
            mutations: Vec::new(),
        };

        sink.emit(
            EventEnvelope::info("engine.txn.begin", trace)
                .with_workbook_id(self.workbook_id)
                .with_txn_id(txn.txn_id),
        )?;

        Ok(txn)
    }

    pub fn snapshot_json(&self) -> serde_json::Value {
        json!({
            "workbook_id": self.workbook_id,
            "sheet_count": self.sheets.len(),
            "sheets": self.sheets.iter().map(|(name, sheet)| {
                json!({
                    "name": name,
                    "cell_count": sheet.cells.len(),
                })
            }).collect::<Vec<_>>()
        })
    }

    fn ensure_sheet_mut(&mut self, name: &str) -> &mut Sheet {
        self.sheets
            .entry(name.to_string())
            .or_insert_with(|| Sheet {
                name: name.to_string(),
                cells: BTreeMap::new(),
            })
    }

    fn set_cell_value(
        &mut self,
        sheet: &str,
        row: u32,
        col: u32,
        value: CellValue,
    ) -> Result<(), ModelError> {
        validate_cell(row, col)?;
        let sheet = self.ensure_sheet_mut(sheet);
        sheet.cells.insert(
            CellRef { row, col },
            CellRecord {
                value,
                formula: None,
            },
        );
        Ok(())
    }

    fn set_cell_formula(
        &mut self,
        sheet: &str,
        row: u32,
        col: u32,
        formula: String,
        cached_value: CellValue,
    ) -> Result<(), ModelError> {
        validate_cell(row, col)?;
        let sheet = self.ensure_sheet_mut(sheet);
        sheet.cells.insert(
            CellRef { row, col },
            CellRecord {
                value: cached_value,
                formula: Some(formula),
            },
        );
        Ok(())
    }
}

impl Transaction {
    pub fn apply(&mut self, mutation: Mutation) {
        self.mutations.push(mutation);
    }

    pub fn commit(
        self,
        workbook: &mut Workbook,
        sink: &mut dyn EventSink,
        trace: &TraceContext,
    ) -> Result<CommitResult, ModelError> {
        let mut changed = BTreeMap::<String, BTreeSet<CellRef>>::new();
        for mutation in &self.mutations {
            match mutation {
                Mutation::SetCellValue {
                    sheet,
                    row,
                    col,
                    value,
                } => {
                    workbook.set_cell_value(sheet, *row, *col, value.clone())?;
                    changed
                        .entry(sheet.to_string())
                        .or_default()
                        .insert(CellRef {
                            row: *row,
                            col: *col,
                        });
                }
                Mutation::SetCellFormula {
                    sheet,
                    row,
                    col,
                    formula,
                    cached_value,
                } => {
                    workbook.set_cell_formula(
                        sheet,
                        *row,
                        *col,
                        formula.clone(),
                        cached_value.clone(),
                    )?;
                    changed
                        .entry(sheet.to_string())
                        .or_default()
                        .insert(CellRef {
                            row: *row,
                            col: *col,
                        });
                }
            }
        }

        let changed_cells = changed
            .into_iter()
            .map(|(k, v)| (k, v.into_iter().collect::<Vec<_>>()))
            .collect::<BTreeMap<_, _>>();

        let result = CommitResult {
            txn_id: self.txn_id,
            workbook_id: workbook.workbook_id,
            committed_at: Utc::now(),
            mutation_count: self.mutations.len(),
            changed_cells,
        };

        sink.emit(
            EventEnvelope::info("engine.txn.commit", trace)
                .with_workbook_id(workbook.workbook_id)
                .with_txn_id(self.txn_id)
                .with_metrics(json!({
                    "mutation_count": result.mutation_count,
                    "changed_sheet_count": result.changed_cells.len(),
                }))
                .with_payload(json!({
                    "changed_cells": result.changed_cells,
                })),
        )?;

        Ok(result)
    }
}

fn validate_cell(row: u32, col: u32) -> Result<(), ModelError> {
    if row == 0 || col == 0 {
        return Err(ModelError::InvalidCellAddress { row, col });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::telemetry::NoopEventSink;

    #[test]
    fn commit_tracks_changed_cells_in_stable_order() {
        let mut wb = Workbook::new();
        let trace = TraceContext::root();
        let mut sink = NoopEventSink;

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 2,
            col: 1,
            value: CellValue::Number(7.0),
        });
        txn.apply(Mutation::SetCellFormula {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
            formula: "=1+1".to_string(),
            cached_value: CellValue::Number(2.0),
        });

        let result = txn.commit(&mut wb, &mut sink, &trace).expect("commit");
        let changed = result.changed_cells.get("Sheet1").expect("sheet");
        assert_eq!(changed[0], CellRef { row: 1, col: 1 });
        assert_eq!(changed[1], CellRef { row: 2, col: 1 });
    }

    #[test]
    fn rejects_zero_row_or_col() {
        let mut wb = Workbook::new();
        let trace = TraceContext::root();
        let mut sink = NoopEventSink;

        let mut txn = wb.begin_txn(&mut sink, &trace).expect("begin");
        txn.apply(Mutation::SetCellValue {
            sheet: "Sheet1".to_string(),
            row: 0,
            col: 1,
            value: CellValue::Empty,
        });

        let err = txn
            .commit(&mut wb, &mut sink, &trace)
            .expect_err("invalid cell");
        assert!(matches!(err, ModelError::InvalidCellAddress { .. }));
    }
}
