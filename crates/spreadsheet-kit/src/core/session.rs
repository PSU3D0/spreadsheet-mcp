use crate::model::{GridCell, GridColumnHint, GridPayload, GridRow, RangeValuesEntry, StylePatch};
use crate::styles::descriptor_from_style;
use crate::workbook::cell_to_value;
use anyhow::{Context, Result, anyhow};
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::Path;
use umya_spreadsheet::{Spreadsheet, Worksheet};

/// Surface-agnostic in-memory workbook session.
///
/// This API avoids workbook IDs, fork handles, and MCP-specific wiring so it can
/// be reused by CLI, SDK, or WASM bindings.
pub struct WorkbookSession {
    spreadsheet: Spreadsheet,
}

impl WorkbookSession {
    /// Open a workbook session from raw XLSX bytes.
    pub fn from_bytes(bytes: impl AsRef<[u8]>) -> Result<Self> {
        let workbook_bytes = bytes.as_ref();
        let cursor = std::io::Cursor::new(workbook_bytes);
        let spreadsheet = umya_spreadsheet::reader::xlsx::read_reader(cursor, true)
            .context("failed to parse workbook bytes")?;
        Ok(Self { spreadsheet })
    }

    /// Open a workbook session from a filesystem path.
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let bytes = fs::read(path)
            .with_context(|| format!("failed to read workbook '{}'", path.display()))?;
        Self::from_bytes(bytes)
    }

    /// Return sheet names in workbook order.
    pub fn list_sheets(&self) -> Vec<String> {
        self.spreadsheet
            .get_sheet_collection()
            .iter()
            .map(|sheet| sheet.get_name().to_string())
            .collect()
    }

    /// Read one or more A1 ranges from a sheet.
    pub fn range_values(
        &self,
        sheet_name: &str,
        ranges: impl Into<SessionRangeSelection>,
    ) -> Result<Vec<RangeValuesEntry>> {
        let sheet = self.sheet_by_name(sheet_name)?;
        let ranges = ranges.into().into_vec();
        if ranges.is_empty() {
            return Err(anyhow!("at least one range is required"));
        }

        let mut out = Vec::with_capacity(ranges.len());
        for range in ranges {
            let bounds = parse_range_bounds(&range)?;
            let mut rows = Vec::new();

            for row in bounds.min_row..=bounds.max_row {
                let mut row_values = Vec::new();
                for col in bounds.min_col..=bounds.max_col {
                    let value = sheet.get_cell((col, row)).and_then(cell_to_value);
                    row_values.push(value);
                }
                rows.push(row_values);
            }

            out.push(RangeValuesEntry {
                range,
                rows: Some(rows),
                formulas: None,
                values: None,
                csv: None,
                next_start_row: None,
            });
        }

        Ok(out)
    }

    /// Export a range as grid payload (value/formula/style patch surface).
    pub fn grid_export(&self, sheet_name: &str, range: &str) -> Result<GridPayload> {
        let sheet = self.sheet_by_name(sheet_name)?;
        let bounds = parse_range_bounds(range)?;

        let mut columns = Vec::new();
        for col_idx in bounds.min_col..=bounds.max_col {
            if let Some(dim) = sheet.get_column_dimension_by_number(&col_idx) {
                let width = *dim.get_width();
                if width > 0.0 {
                    columns.push(GridColumnHint {
                        offset: col_idx - bounds.min_col,
                        width_chars: width,
                    });
                }
            }
        }

        let mut merges = Vec::new();
        for merge_cell in sheet.get_merge_cells() {
            let merge_range = merge_cell.get_range();
            if let Ok(merge_bounds) = parse_range_bounds(&merge_range)
                && merge_bounds.min_col <= bounds.max_col
                && merge_bounds.max_col >= bounds.min_col
                && merge_bounds.min_row <= bounds.max_row
                && merge_bounds.max_row >= bounds.min_row
            {
                merges.push(merge_range.to_string());
            }
        }

        let mut rows = Vec::new();
        for row in bounds.min_row..=bounds.max_row {
            let mut cells = Vec::new();
            for col in bounds.min_col..=bounds.max_col {
                let Some(cell) = sheet.get_cell((&col, &row)) else {
                    continue;
                };

                let (value, formula) = if cell.is_formula() {
                    (None, Some(format!("={}", cell.get_formula())))
                } else {
                    (cell_to_json_value(cell_to_value(cell)), None)
                };

                let descriptor = descriptor_from_style(cell.get_style());
                let number_format = descriptor.number_format.clone();
                let style_patch = style_descriptor_to_patch(descriptor);

                if value.is_some()
                    || formula.is_some()
                    || number_format.is_some()
                    || style_patch.is_some()
                {
                    cells.push(GridCell {
                        offset: [row - bounds.min_row, col - bounds.min_col],
                        v: value,
                        f: formula,
                        fmt: number_format,
                        style: style_patch,
                    });
                }
            }
            if !cells.is_empty() {
                rows.push(GridRow { cells });
            }
        }

        Ok(GridPayload {
            sheet: sheet_name.to_string(),
            anchor: crate::utils::cell_address(bounds.min_col, bounds.min_row),
            columns,
            merges,
            rows,
        })
    }

    /// Apply a transform batch in-session.
    ///
    /// For this extraction pass we support `write_matrix` operations.
    pub fn apply_ops(&mut self, ops: &[SessionTransformOp]) -> Result<SessionApplySummary> {
        // Validate all operations first so failures are atomic for this batch.
        self.validate_ops(ops)?;

        let mut summary = SessionApplySummary {
            ops_applied: ops.len(),
            ..Default::default()
        };

        for op in ops {
            match op {
                SessionTransformOp::WriteMatrix {
                    sheet_name,
                    anchor,
                    rows,
                    overwrite_formulas,
                } => {
                    let sheet = self.sheet_by_name_mut(sheet_name)?;
                    let (anchor_col, anchor_row) = parse_cell_ref(anchor)?;

                    for (row_offset, row_values) in rows.iter().enumerate() {
                        let row_idx = anchor_row + row_offset as u32;
                        for (col_offset, cell_value) in row_values.iter().enumerate() {
                            let Some(cell_value) = cell_value else {
                                continue;
                            };
                            let col_idx = anchor_col + col_offset as u32;
                            let cell = sheet.get_cell_mut((col_idx, row_idx));
                            summary.cells_touched += 1;

                            if cell.is_formula() {
                                if !*overwrite_formulas {
                                    summary.cells_skipped_keep_formulas += 1;
                                    continue;
                                }
                                cell.set_formula(String::new());
                                summary.cells_formula_cleared += 1;
                            }

                            match cell_value {
                                SessionMatrixCell::Value(raw) => {
                                    cell.set_value(json_value_to_cell_string(raw));
                                    summary.cells_value_set += 1;
                                }
                                SessionMatrixCell::Formula(formula) => {
                                    let formula = formula.strip_prefix('=').unwrap_or(formula);
                                    cell.set_formula(formula.to_string());
                                    cell.set_formula_result_default("");
                                    summary.cells_formula_set += 1;
                                }
                            }
                        }
                    }
                }
            }
        }

        Ok(summary)
    }

    /// Convenience wrapper for a single `write_matrix` operation.
    pub fn apply_write_matrix(
        &mut self,
        sheet_name: impl Into<String>,
        anchor: impl Into<String>,
        rows: Vec<Vec<Option<SessionMatrixCell>>>,
        overwrite_formulas: bool,
    ) -> Result<SessionApplySummary> {
        let op = SessionTransformOp::WriteMatrix {
            sheet_name: sheet_name.into(),
            anchor: anchor.into(),
            rows,
            overwrite_formulas,
        };
        self.apply_ops(&[op])
    }

    /// Serialize the current in-memory workbook state back to XLSX bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        umya_spreadsheet::writer::xlsx::write_writer(&self.spreadsheet, &mut bytes)
            .context("failed to serialize workbook to bytes")?;
        Ok(bytes)
    }

    /// Serialize and consume the current in-memory workbook state.
    pub fn into_bytes(self) -> Result<Vec<u8>> {
        self.to_bytes()
    }

    fn sheet_by_name(&self, sheet_name: &str) -> Result<&Worksheet> {
        self.spreadsheet
            .get_sheet_by_name(sheet_name)
            .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))
    }

    fn sheet_by_name_mut(&mut self, sheet_name: &str) -> Result<&mut Worksheet> {
        self.spreadsheet
            .get_sheet_by_name_mut(sheet_name)
            .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))
    }

    fn validate_ops(&self, ops: &[SessionTransformOp]) -> Result<()> {
        for op in ops {
            match op {
                SessionTransformOp::WriteMatrix {
                    sheet_name, anchor, ..
                } => {
                    self.sheet_by_name(sheet_name)?;
                    let _ = parse_cell_ref(anchor)?;
                }
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub enum SessionRangeSelection {
    Single(String),
    Multi(Vec<String>),
}

impl SessionRangeSelection {
    fn into_vec(self) -> Vec<String> {
        match self {
            SessionRangeSelection::Single(range) => vec![range],
            SessionRangeSelection::Multi(ranges) => ranges,
        }
    }
}

impl From<String> for SessionRangeSelection {
    fn from(value: String) -> Self {
        SessionRangeSelection::Single(value)
    }
}

impl From<&str> for SessionRangeSelection {
    fn from(value: &str) -> Self {
        SessionRangeSelection::Single(value.to_string())
    }
}

impl From<Vec<String>> for SessionRangeSelection {
    fn from(value: Vec<String>) -> Self {
        SessionRangeSelection::Multi(value)
    }
}

impl From<Vec<&str>> for SessionRangeSelection {
    fn from(value: Vec<&str>) -> Self {
        SessionRangeSelection::Multi(value.into_iter().map(str::to_string).collect())
    }
}

impl From<&[String]> for SessionRangeSelection {
    fn from(value: &[String]) -> Self {
        SessionRangeSelection::Multi(value.to_vec())
    }
}

impl From<&[&str]> for SessionRangeSelection {
    fn from(value: &[&str]) -> Self {
        SessionRangeSelection::Multi(value.iter().map(|entry| entry.to_string()).collect())
    }
}

impl<const N: usize> From<[&str; N]> for SessionRangeSelection {
    fn from(value: [&str; N]) -> Self {
        SessionRangeSelection::Multi(value.into_iter().map(str::to_string).collect())
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub enum SessionMatrixCell {
    #[serde(rename = "v")]
    Value(serde_json::Value),
    #[serde(rename = "f")]
    Formula(String),
}

fn default_overwrite_formulas() -> bool {
    false
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum SessionTransformOp {
    WriteMatrix {
        sheet_name: String,
        anchor: String,
        rows: Vec<Vec<Option<SessionMatrixCell>>>,
        #[serde(default = "default_overwrite_formulas")]
        overwrite_formulas: bool,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, Default, PartialEq, Eq)]
pub struct SessionApplySummary {
    pub ops_applied: usize,
    pub cells_touched: u64,
    pub cells_value_set: u64,
    pub cells_formula_set: u64,
    pub cells_formula_cleared: u64,
    pub cells_skipped_keep_formulas: u64,
}

#[derive(Debug, Clone, Copy)]
struct RangeBounds {
    min_col: u32,
    min_row: u32,
    max_col: u32,
    max_row: u32,
}

fn parse_cell_ref(cell: &str) -> Result<(u32, u32)> {
    use umya_spreadsheet::helper::coordinate::index_from_coordinate;

    let (col, row, _, _) = index_from_coordinate(cell);
    match (col, row) {
        (Some(c), Some(r)) => Ok((c, r)),
        _ => Err(anyhow!("invalid cell reference: {}", cell)),
    }
}

fn parse_range_bounds(range: &str) -> Result<RangeBounds> {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.is_empty() || parts.len() > 2 {
        return Err(anyhow!(
            "invalid range '{}'; expected 'A1' or 'A1:Z99'",
            range
        ));
    }

    let start = parse_cell_ref(parts[0])?;
    let end = if parts.len() == 2 {
        parse_cell_ref(parts[1])?
    } else {
        start
    };

    Ok(RangeBounds {
        min_col: start.0.min(end.0),
        min_row: start.1.min(end.1),
        max_col: start.0.max(end.0),
        max_row: start.1.max(end.1),
    })
}

fn json_value_to_cell_string(value: &serde_json::Value) -> String {
    match value {
        serde_json::Value::Null => String::new(),
        serde_json::Value::Bool(raw) => raw.to_string(),
        serde_json::Value::Number(raw) => raw.to_string(),
        serde_json::Value::String(raw) => raw.clone(),
        serde_json::Value::Array(_) | serde_json::Value::Object(_) => value.to_string(),
    }
}

fn cell_to_json_value(value: Option<crate::model::CellValue>) -> Option<serde_json::Value> {
    match value {
        Some(crate::model::CellValue::Text(text)) => Some(serde_json::Value::String(text)),
        Some(crate::model::CellValue::Number(number)) => Some(serde_json::json!(number)),
        Some(crate::model::CellValue::Bool(value)) => Some(serde_json::Value::Bool(value)),
        Some(crate::model::CellValue::Error(text)) => Some(serde_json::Value::String(text)),
        Some(crate::model::CellValue::Date(text)) => Some(serde_json::Value::String(text)),
        None => None,
    }
}

fn style_descriptor_to_patch(desc: crate::model::StyleDescriptor) -> Option<StylePatch> {
    if desc.font.is_none()
        && desc.fill.is_none()
        && desc.borders.is_none()
        && desc.alignment.is_none()
    {
        return None;
    }

    Some(StylePatch {
        font: desc.font.map(|font| {
            Some(crate::model::FontPatch {
                name: font.name.map(Some),
                size: font.size.map(Some),
                bold: font.bold.map(Some),
                italic: font.italic.map(Some),
                underline: font.underline.map(Some),
                strikethrough: font.strikethrough.map(Some),
                color: font.color.map(Some),
            })
        }),
        fill: desc.fill.map(|fill| {
            Some(match fill {
                crate::model::FillDescriptor::Pattern(pattern) => {
                    crate::model::FillPatch::Pattern(crate::model::PatternFillPatch {
                        pattern_type: pattern.pattern_type.map(Some),
                        foreground_color: pattern.foreground_color.map(Some),
                        background_color: pattern.background_color.map(Some),
                    })
                }
                crate::model::FillDescriptor::Gradient(gradient) => {
                    crate::model::FillPatch::Gradient(crate::model::GradientFillPatch {
                        degree: gradient.degree.map(Some),
                        stops: Some(
                            gradient
                                .stops
                                .into_iter()
                                .map(|stop| crate::model::GradientStopPatch {
                                    position: stop.position,
                                    color: stop.color,
                                })
                                .collect(),
                        ),
                    })
                }
            })
        }),
        borders: desc.borders.map(|borders| {
            Some(crate::model::BordersPatch {
                left: borders.left.map(|side| {
                    Some(crate::model::BorderSidePatch {
                        style: side.style.map(Some),
                        color: side.color.map(Some),
                    })
                }),
                right: borders.right.map(|side| {
                    Some(crate::model::BorderSidePatch {
                        style: side.style.map(Some),
                        color: side.color.map(Some),
                    })
                }),
                top: borders.top.map(|side| {
                    Some(crate::model::BorderSidePatch {
                        style: side.style.map(Some),
                        color: side.color.map(Some),
                    })
                }),
                bottom: borders.bottom.map(|side| {
                    Some(crate::model::BorderSidePatch {
                        style: side.style.map(Some),
                        color: side.color.map(Some),
                    })
                }),
                diagonal: borders.diagonal.map(|side| {
                    Some(crate::model::BorderSidePatch {
                        style: side.style.map(Some),
                        color: side.color.map(Some),
                    })
                }),
                vertical: borders.vertical.map(|side| {
                    Some(crate::model::BorderSidePatch {
                        style: side.style.map(Some),
                        color: side.color.map(Some),
                    })
                }),
                horizontal: borders.horizontal.map(|side| {
                    Some(crate::model::BorderSidePatch {
                        style: side.style.map(Some),
                        color: side.color.map(Some),
                    })
                }),
                diagonal_up: borders.diagonal_up.map(Some),
                diagonal_down: borders.diagonal_down.map(Some),
            })
        }),
        alignment: desc.alignment.map(|alignment| {
            Some(crate::model::AlignmentPatch {
                horizontal: alignment.horizontal.map(Some),
                vertical: alignment.vertical.map(Some),
                wrap_text: alignment.wrap_text.map(Some),
                text_rotation: alignment.text_rotation.map(Some),
            })
        }),
        number_format: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::CellValue;
    use anyhow::Result;
    use tempfile::tempdir;

    fn workbook_bytes(setup: impl FnOnce(&mut Spreadsheet)) -> Vec<u8> {
        let mut book = umya_spreadsheet::new_file();
        setup(&mut book);

        let mut bytes = Vec::new();
        umya_spreadsheet::writer::xlsx::write_writer(&book, &mut bytes).expect("write workbook");
        bytes
    }

    fn grid_cell(payload: &GridPayload, row_offset: u32, col_offset: u32) -> Option<&GridCell> {
        payload
            .rows
            .iter()
            .flat_map(|row| row.cells.iter())
            .find(|cell| cell.offset == [row_offset, col_offset])
    }

    #[test]
    fn bytes_roundtrip_and_multi_range_reads() -> Result<()> {
        let bytes = workbook_bytes(|book| {
            book.get_sheet_by_name_mut("Sheet1")
                .expect("sheet")
                .get_cell_mut("A1")
                .set_value("hello");
            let _ = book.new_sheet("Data");
            book.get_sheet_by_name_mut("Data")
                .expect("data sheet")
                .get_cell_mut("B2")
                .set_value_number(42.0);
        });

        let session = WorkbookSession::from_bytes(bytes)?;
        assert_eq!(session.list_sheets(), vec!["Sheet1", "Data"]);

        let entries = session.range_values("Data", vec!["A1:B2", "B2:B2"])?;
        assert_eq!(entries.len(), 2);

        let rows = entries[0].rows.as_ref().expect("rows");
        assert!(
            matches!(rows[1][1], Some(CellValue::Number(v)) if (v - 42.0).abs() < f64::EPSILON)
        );

        let out_bytes = session.into_bytes()?;
        let reopened = WorkbookSession::from_bytes(out_bytes)?;
        let reopened_entries = reopened.range_values("Sheet1", "A1")?;
        let reopened_rows = reopened_entries[0].rows.as_ref().expect("rows");
        assert!(matches!(
            reopened_rows[0][0],
            Some(CellValue::Text(ref value)) if value == "hello"
        ));

        Ok(())
    }

    #[test]
    fn apply_write_matrix_updates_in_session_and_roundtrips() -> Result<()> {
        let bytes = workbook_bytes(|book| {
            let sheet = book.get_sheet_by_name_mut("Sheet1").expect("sheet");
            sheet.get_cell_mut("A1").set_value("before");
            sheet.get_cell_mut("B1").set_formula("1+1");
        });

        let mut session = WorkbookSession::from_bytes(bytes)?;

        let summary = session.apply_write_matrix(
            "Sheet1",
            "A1",
            vec![vec![
                Some(SessionMatrixCell::Value(serde_json::json!("after"))),
                Some(SessionMatrixCell::Value(serde_json::json!(99))),
            ]],
            false,
        )?;

        assert_eq!(summary.ops_applied, 1);
        assert_eq!(summary.cells_skipped_keep_formulas, 1);

        let before_overwrite = session.grid_export("Sheet1", "A1:B1")?;
        let b1_before = grid_cell(&before_overwrite, 0, 1).expect("B1 cell");
        assert_eq!(b1_before.f.as_deref(), Some("=1+1"));

        session.apply_write_matrix(
            "Sheet1",
            "B1",
            vec![vec![Some(SessionMatrixCell::Formula(
                "=SUM(1,2)".to_string(),
            ))]],
            true,
        )?;

        let after = session.grid_export("Sheet1", "A1:B1")?;
        let a1_after = grid_cell(&after, 0, 0).expect("A1 cell");
        let b1_after = grid_cell(&after, 0, 1).expect("B1 cell");
        assert_eq!(a1_after.v, Some(serde_json::json!("after")));
        assert_eq!(b1_after.f.as_deref(), Some("=SUM(1,2)"));

        let roundtrip = WorkbookSession::from_bytes(session.into_bytes()?)?;
        let persisted = roundtrip.grid_export("Sheet1", "B1")?;
        let persisted_b1 = grid_cell(&persisted, 0, 0).expect("persisted B1");
        assert_eq!(persisted_b1.f.as_deref(), Some("=SUM(1,2)"));

        Ok(())
    }

    #[test]
    fn write_matrix_default_overwrite_formulas_is_false() -> Result<()> {
        let raw = serde_json::json!({
            "kind": "write_matrix",
            "sheet_name": "Sheet1",
            "anchor": "A1",
            "rows": [[{"v": "x"}]]
        });

        let op: SessionTransformOp = serde_json::from_value(raw)?;
        match op {
            SessionTransformOp::WriteMatrix {
                overwrite_formulas, ..
            } => {
                assert!(!overwrite_formulas);
            }
        }

        Ok(())
    }

    #[test]
    fn apply_ops_is_atomic_on_validation_failure() -> Result<()> {
        let bytes = workbook_bytes(|book| {
            let sheet = book.get_sheet_by_name_mut("Sheet1").expect("sheet");
            sheet.get_cell_mut("A1").set_value("before");
        });

        let mut session = WorkbookSession::from_bytes(bytes)?;

        let ops = vec![
            SessionTransformOp::WriteMatrix {
                sheet_name: "Sheet1".to_string(),
                anchor: "A1".to_string(),
                rows: vec![vec![Some(SessionMatrixCell::Value(serde_json::json!(
                    "after"
                )))]],
                overwrite_formulas: false,
            },
            SessionTransformOp::WriteMatrix {
                sheet_name: "MissingSheet".to_string(),
                anchor: "A1".to_string(),
                rows: vec![vec![Some(SessionMatrixCell::Value(serde_json::json!(
                    "bad"
                )))]],
                overwrite_formulas: false,
            },
        ];

        let err = session.apply_ops(&ops).unwrap_err();
        assert!(err.to_string().contains("MissingSheet"));

        let after = session.range_values("Sheet1", "A1")?;
        let rows = after[0].rows.as_ref().expect("rows");
        assert!(matches!(
            rows[0][0],
            Some(CellValue::Text(ref value)) if value == "before"
        ));

        Ok(())
    }

    #[test]
    fn from_path_loads_workbook() -> Result<()> {
        let dir = tempdir()?;
        let path = dir.path().join("session-path.xlsx");

        let bytes = workbook_bytes(|book| {
            book.get_sheet_by_name_mut("Sheet1")
                .expect("sheet")
                .get_cell_mut("C3")
                .set_value("path-load");
        });
        fs::write(&path, bytes)?;

        let session = WorkbookSession::from_path(&path)?;
        let entries = session.range_values("Sheet1", "C3")?;
        let rows = entries[0].rows.as_ref().expect("rows");
        assert!(matches!(
            rows[0][0],
            Some(CellValue::Text(ref value)) if value == "path-load"
        ));

        Ok(())
    }
}
