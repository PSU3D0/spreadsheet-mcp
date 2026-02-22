use crate::config::{OutputProfile, RecalcBackendKind, ServerConfig, TransportKind};
use crate::model::{
    CellSnapshot, CellValue, GridCell, GridColumnHint, GridPayload, GridRow, NamedRangesResponse,
    RangeValuesEntry, RowSnapshot, SheetOverviewResponse, SheetPageCompact, SheetPageFormat,
    SheetPageResponse, SheetPageValues, StylePatch, WorkbookDescription, WorkbookId,
};
use crate::styles::descriptor_from_style;
use crate::workbook::{WorkbookContext, cell_to_value};
use anyhow::{Context, Result, anyhow};
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Arc;
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

    /// Return workbook-level descriptor for this in-memory session.
    pub fn describe_workbook(&self) -> Result<WorkbookDescription> {
        let workbook = self.as_workbook_context()?;
        Ok(workbook.describe())
    }

    /// Return workbook defined names and table descriptors.
    pub fn named_ranges(&self) -> Result<NamedRangesResponse> {
        let workbook = self.as_workbook_context()?;
        let items = workbook.named_items()?;
        Ok(NamedRangesResponse {
            workbook_id: workbook.id.clone(),
            items,
        })
    }

    /// Return overview/classification information for a sheet.
    pub fn sheet_overview(
        &self,
        params: SessionSheetOverviewParams,
    ) -> Result<SheetOverviewResponse> {
        let workbook = self.as_workbook_context()?;
        let mut overview = workbook.sheet_overview(&params.sheet_name)?;

        let max_regions = params.max_regions.unwrap_or(25).max(1);
        let max_headers = params.max_headers.unwrap_or(50).max(1);
        let include_headers = params.include_headers.unwrap_or(true);

        let region_limit = if params.max_regions == Some(0) {
            usize::MAX
        } else {
            max_regions as usize
        };
        let header_limit = if params.max_headers == Some(0) {
            usize::MAX
        } else {
            max_headers as usize
        };

        let total_regions = overview.detected_regions.len() as u32;
        let mut headers_truncated = false;

        for region in &mut overview.detected_regions {
            let header_count = region.header_count.max(region.headers.len() as u32);
            region.header_count = header_count;
            if !include_headers {
                region.headers.clear();
            } else if region.headers.len() > header_limit {
                region.headers.truncate(header_limit);
            }
            region.headers_truncated = region.headers.len() as u32 != header_count;
            headers_truncated |= region.headers_truncated;
        }

        let regions_truncated = if overview.detected_regions.len() > region_limit {
            overview.detected_regions.truncate(region_limit);
            true
        } else {
            false
        };

        overview.detected_region_count = total_regions;
        overview.detected_regions_truncated = regions_truncated;

        if regions_truncated {
            overview.notes.push(format!(
                "Detected regions truncated to {} ({} total).",
                region_limit, total_regions
            ));
        }
        if headers_truncated {
            overview.notes.push(format!(
                "Region headers truncated to {} columns.",
                header_limit
            ));
        }

        Ok(overview)
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

    /// Read a page-oriented snapshot from a sheet.
    pub fn sheet_page(&self, params: SessionSheetPageParams) -> Result<SheetPageResponse> {
        if params.page_size == 0 {
            return Err(anyhow!("page_size must be greater than zero"));
        }

        let sheet = self.sheet_by_name(&params.sheet_name)?;
        let start_row = params.start_row.max(1);
        let page_size = params.page_size.min(500);
        let max_row = sheet.get_highest_row();

        let page = build_sheet_page(
            sheet,
            start_row,
            page_size,
            params.columns.as_ref(),
            params.columns_by_header.as_ref(),
            params.include_formulas,
            params.include_styles,
            params.include_header,
        )?;

        let last_row_index = page
            .rows
            .last()
            .map(|row| row.row_index)
            .unwrap_or(start_row.saturating_sub(1));
        let next_start_row = if last_row_index < max_row {
            Some(last_row_index + 1)
        } else {
            None
        };

        Ok(build_sheet_page_response(
            WorkbookId("session".to_string()),
            params.sheet_name,
            params.format,
            params.include_header,
            page.header,
            page.rows,
            next_start_row,
        ))
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

    fn as_workbook_context(&self) -> Result<WorkbookContext> {
        let bytes = self.to_bytes()?;
        let workbook_id = WorkbookId("session".to_string());
        let short_id = crate::utils::make_short_workbook_id("session", workbook_id.as_str());
        let config = Arc::new(ServerConfig {
            workspace_root: PathBuf::from("."),
            screenshot_dir: PathBuf::from("screenshots"),
            path_mappings: Vec::new(),
            cache_capacity: 2,
            supported_extensions: vec![
                "xlsx".to_string(),
                "xlsm".to_string(),
                "xls".to_string(),
                "xlsb".to_string(),
            ],
            single_workbook: None,
            enabled_tools: None,
            transport: TransportKind::Stdio,
            http_bind_address: "127.0.0.1:8079"
                .parse()
                .expect("hardcoded bind address is valid"),
            recalc_enabled: false,
            recalc_backend: RecalcBackendKind::Auto,
            vba_enabled: false,
            max_concurrent_recalcs: 1,
            tool_timeout_ms: Some(30_000),
            max_response_bytes: Some(1_000_000),
            output_profile: OutputProfile::Verbose,
            max_payload_bytes: Some(65_536),
            max_cells: Some(10_000),
            max_items: Some(500),
            allow_overwrite: true,
        });

        WorkbookContext::load_from_bytes(
            &config,
            "session.xlsx",
            &bytes,
            workbook_id,
            short_id,
            None,
        )
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

fn default_start_row() -> u32 {
    1
}

fn default_page_size() -> u32 {
    50
}

fn default_include_formulas() -> bool {
    true
}

fn default_include_styles() -> bool {
    false
}

fn default_include_header() -> bool {
    true
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "snake_case")]
pub struct SessionSheetOverviewParams {
    pub sheet_name: String,
    #[serde(default)]
    pub max_regions: Option<u32>,
    #[serde(default)]
    pub max_headers: Option<u32>,
    #[serde(default)]
    pub include_headers: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub struct SessionSheetPageParams {
    pub sheet_name: String,
    #[serde(default = "default_start_row")]
    pub start_row: u32,
    #[serde(default = "default_page_size")]
    pub page_size: u32,
    #[serde(default)]
    pub columns: Option<Vec<String>>,
    #[serde(default)]
    pub columns_by_header: Option<Vec<String>>,
    #[serde(default = "default_include_formulas")]
    pub include_formulas: bool,
    #[serde(default = "default_include_styles")]
    pub include_styles: bool,
    #[serde(default = "default_include_header")]
    pub include_header: bool,
    #[serde(default)]
    pub format: SheetPageFormat,
}

impl SessionSheetPageParams {
    pub fn with_sheet_name(sheet_name: impl Into<String>) -> Self {
        Self {
            sheet_name: sheet_name.into(),
            ..Self::default()
        }
    }
}

impl Default for SessionSheetPageParams {
    fn default() -> Self {
        Self {
            sheet_name: String::new(),
            start_row: default_start_row(),
            page_size: default_page_size(),
            columns: None,
            columns_by_header: None,
            include_formulas: default_include_formulas(),
            include_styles: default_include_styles(),
            include_header: default_include_header(),
            format: SheetPageFormat::default(),
        }
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

struct PageBuildResult {
    rows: Vec<RowSnapshot>,
    header: Option<RowSnapshot>,
}

#[allow(clippy::too_many_arguments)]
fn build_sheet_page(
    sheet: &umya_spreadsheet::Worksheet,
    start_row: u32,
    page_size: u32,
    columns: Option<&Vec<String>>,
    columns_by_header: Option<&Vec<String>>,
    include_formulas: bool,
    include_styles: bool,
    include_header: bool,
) -> Result<PageBuildResult> {
    let max_col = sheet.get_highest_column();
    let end_row = start_row
        .saturating_add(page_size.saturating_sub(1))
        .min(sheet.get_highest_row());
    let column_indices = resolve_columns_with_headers(sheet, columns, columns_by_header, max_col)?;

    let header = if include_header {
        Some(build_row_snapshot(
            sheet,
            1,
            &column_indices,
            include_formulas,
            include_styles,
        ))
    } else {
        None
    };

    let mut rows = Vec::new();
    for row_idx in start_row..=end_row {
        rows.push(build_row_snapshot(
            sheet,
            row_idx,
            &column_indices,
            include_formulas,
            include_styles,
        ));
    }

    Ok(PageBuildResult { rows, header })
}

fn build_row_snapshot(
    sheet: &umya_spreadsheet::Worksheet,
    row_index: u32,
    columns: &[u32],
    include_formulas: bool,
    include_styles: bool,
) -> RowSnapshot {
    let mut cells = Vec::new();
    for &col in columns {
        if let Some(cell) = sheet.get_cell((col, row_index)) {
            cells.push(build_cell_snapshot(cell, include_formulas, include_styles));
        } else {
            let address = crate::utils::cell_address(col, row_index);
            cells.push(CellSnapshot {
                address,
                value: None,
                formula: None,
                cached_value: None,
                number_format: None,
                style_tags: Vec::new(),
                notes: Vec::new(),
            });
        }
    }

    RowSnapshot { row_index, cells }
}

fn build_cell_snapshot(
    cell: &umya_spreadsheet::Cell,
    include_formulas: bool,
    include_styles: bool,
) -> CellSnapshot {
    let address = cell.get_coordinate().get_coordinate();
    let value = crate::workbook::cell_to_value(cell);
    let formula = if include_formulas && cell.is_formula() {
        Some(cell.get_formula().to_string())
    } else {
        None
    };
    let cached_value = if cell.is_formula() {
        value.clone()
    } else {
        None
    };
    let number_format = if include_styles {
        cell.get_style()
            .get_number_format()
            .map(|fmt| fmt.get_format_code().to_string())
    } else {
        None
    };
    let style_tags = if include_styles {
        crate::analysis::style::tag_cell(cell)
            .map(|(_, tagging)| tagging.tags)
            .unwrap_or_default()
    } else {
        Vec::new()
    };

    CellSnapshot {
        address,
        value,
        formula,
        cached_value,
        number_format,
        style_tags,
        notes: Vec::new(),
    }
}

fn parse_column_index(spec: &str) -> Result<u32> {
    use umya_spreadsheet::helper::coordinate::column_index_from_string;

    let trimmed = spec.trim();
    if trimmed.is_empty() {
        return Err(anyhow!("invalid column spec: empty"));
    }
    if !trimmed.chars().all(|c| c.is_ascii_alphabetic()) {
        return Err(anyhow!(
            "invalid column spec '{}'; expected letters like 'A' or 'A:C'",
            spec
        ));
    }
    if trimmed.len() > 3 {
        return Err(anyhow!(
            "invalid column spec '{}'; expected at most 3 column letters",
            spec
        ));
    }

    Ok(column_index_from_string(trimmed.to_ascii_uppercase()))
}

fn resolve_columns(columns: Option<&Vec<String>>, max_column: u32) -> Result<Vec<u32>> {
    use std::collections::BTreeSet;

    let mut indices = BTreeSet::new();
    if let Some(specs) = columns {
        for spec in specs {
            if let Some((start, end)) = spec.split_once(':') {
                let start_idx = parse_column_index(start)?;
                let end_idx = parse_column_index(end)?;
                let (min_idx, max_idx) = if start_idx <= end_idx {
                    (start_idx, end_idx)
                } else {
                    (end_idx, start_idx)
                };
                for idx in min_idx..=max_idx {
                    indices.insert(idx);
                }
            } else {
                indices.insert(parse_column_index(spec)?);
            }
        }
    } else {
        for idx in 1..=max_column.max(1) {
            indices.insert(idx);
        }
    }

    Ok(indices.into_iter().collect())
}

fn resolve_columns_with_headers(
    sheet: &umya_spreadsheet::Worksheet,
    columns: Option<&Vec<String>>,
    columns_by_header: Option<&Vec<String>>,
    max_column: u32,
) -> Result<Vec<u32>> {
    use std::collections::BTreeSet;

    if columns_by_header.is_none() {
        return resolve_columns(columns, max_column);
    }

    let mut selected: BTreeSet<u32> = if columns.is_some() {
        resolve_columns(columns, max_column)?.into_iter().collect()
    } else {
        BTreeSet::new()
    };
    let mut matched_header = false;
    let header_targets: Vec<String> = columns_by_header
        .expect("checked")
        .iter()
        .map(|h| h.trim().to_ascii_lowercase())
        .collect();

    for col_idx in 1..=max_column.max(1) {
        let header_cell = sheet.get_cell((col_idx, 1u32));
        let header_value = header_cell
            .and_then(cell_to_value)
            .map(cell_value_to_string_lower);
        if let Some(hval) = header_value
            && header_targets.iter().any(|target| target == &hval)
        {
            selected.insert(col_idx);
            matched_header = true;
        }
    }

    if !matched_header && columns.is_none() {
        resolve_columns(None, max_column)
    } else {
        Ok(selected.into_iter().collect())
    }
}

fn cell_value_to_string_lower(value: CellValue) -> String {
    match value {
        CellValue::Text(s) => s.to_ascii_lowercase(),
        CellValue::Number(n) => n.to_string().to_ascii_lowercase(),
        CellValue::Bool(b) => b.to_string(),
        CellValue::Error(e) => e.to_ascii_lowercase(),
        CellValue::Date(d) => d.to_ascii_lowercase(),
    }
}

fn build_compact_payload(
    header: &Option<RowSnapshot>,
    rows: &[RowSnapshot],
    include_header: bool,
) -> SheetPageCompact {
    let headers = derive_headers(header, rows);
    let header_row = if include_header {
        header
            .as_ref()
            .map(|h| h.cells.iter().map(|c| c.value.clone()).collect())
            .unwrap_or_default()
    } else {
        Vec::new()
    };
    let data_rows = rows
        .iter()
        .map(|row| {
            let mut vals: Vec<Option<CellValue>> = Vec::new();
            vals.push(Some(CellValue::Number(row.row_index as f64)));
            vals.extend(row.cells.iter().map(|c| c.value.clone()));
            vals
        })
        .collect();

    SheetPageCompact {
        headers,
        header_row,
        rows: data_rows,
    }
}

fn build_values_only_payload(
    header: &Option<RowSnapshot>,
    rows: &[RowSnapshot],
    include_header: bool,
) -> SheetPageValues {
    let mut data = Vec::new();
    if include_header && let Some(h) = header {
        data.push(h.cells.iter().map(|c| c.value.clone()).collect());
    }
    for row in rows {
        data.push(row.cells.iter().map(|c| c.value.clone()).collect());
    }

    SheetPageValues { rows: data }
}

fn build_sheet_page_response(
    workbook_id: WorkbookId,
    sheet_name: String,
    format: SheetPageFormat,
    include_header: bool,
    header: Option<RowSnapshot>,
    rows: Vec<RowSnapshot>,
    next_start_row: Option<u32>,
) -> SheetPageResponse {
    let compact_payload = if matches!(format, SheetPageFormat::Compact) {
        Some(build_compact_payload(&header, &rows, include_header))
    } else {
        None
    };

    let values_only_payload = if matches!(format, SheetPageFormat::ValuesOnly) {
        Some(build_values_only_payload(&header, &rows, include_header))
    } else {
        None
    };

    let rows_payload = if matches!(format, SheetPageFormat::Full) {
        rows
    } else {
        Vec::new()
    };

    let header_row = if include_header && matches!(format, SheetPageFormat::Full) {
        header
    } else {
        None
    };

    SheetPageResponse {
        workbook_id,
        sheet_name,
        rows: rows_payload,
        next_start_row,
        header_row,
        compact: compact_payload,
        values_only: values_only_payload,
        format,
    }
}

fn derive_headers(header: &Option<RowSnapshot>, rows: &[RowSnapshot]) -> Vec<String> {
    if let Some(h) = header {
        let mut headers: Vec<String> = h
            .cells
            .iter()
            .map(|c| match &c.value {
                Some(CellValue::Text(t)) => t.clone(),
                Some(CellValue::Number(n)) => n.to_string(),
                Some(CellValue::Bool(b)) => b.to_string(),
                Some(CellValue::Date(d)) => d.clone(),
                Some(CellValue::Error(e)) => e.clone(),
                None => c.address.clone(),
            })
            .collect();
        headers.insert(0, "Row".to_string());
        headers
    } else if let Some(first) = rows.first() {
        let mut headers = Vec::new();
        headers.push("Row".to_string());
        for cell in &first.cells {
            headers.push(cell.address.clone());
        }
        headers
    } else {
        vec![]
    }
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
    fn sheet_page_full_supports_paging_and_formulas() -> Result<()> {
        let bytes = workbook_bytes(|book| {
            let sheet = book.get_sheet_by_name_mut("Sheet1").expect("sheet");
            sheet.get_cell_mut("A1").set_value("Name");
            sheet.get_cell_mut("B1").set_value("Calc");
            sheet.get_cell_mut("A2").set_value("alpha");
            sheet.get_cell_mut("B2").set_formula("1+1");
            sheet.get_cell_mut("A3").set_value("beta");
            sheet.get_cell_mut("B3").set_value_number(7.0);
        });

        let session = WorkbookSession::from_bytes(bytes)?;
        let page = session.sheet_page(SessionSheetPageParams {
            sheet_name: "Sheet1".to_string(),
            start_row: 2,
            page_size: 1,
            include_formulas: true,
            format: SheetPageFormat::Full,
            ..SessionSheetPageParams::default()
        })?;

        assert_eq!(page.sheet_name, "Sheet1");
        assert!(matches!(page.workbook_id, WorkbookId(ref id) if id == "session"));
        assert_eq!(page.next_start_row, Some(3));
        let header = page.header_row.as_ref().expect("header row");
        assert_eq!(header.row_index, 1);
        assert_eq!(page.rows.len(), 1);
        assert_eq!(page.rows[0].row_index, 2);
        assert_eq!(page.rows[0].cells[1].formula.as_deref(), Some("1+1"));

        Ok(())
    }

    #[test]
    fn sheet_page_compact_respects_columns_by_header() -> Result<()> {
        let bytes = workbook_bytes(|book| {
            let sheet = book.get_sheet_by_name_mut("Sheet1").expect("sheet");
            sheet.get_cell_mut("A1").set_value("Name");
            sheet.get_cell_mut("B1").set_value("Score");
            sheet.get_cell_mut("C1").set_value("Ignore");
            sheet.get_cell_mut("A2").set_value("alpha");
            sheet.get_cell_mut("B2").set_value_number(99.0);
            sheet.get_cell_mut("C2").set_value("x");
        });

        let session = WorkbookSession::from_bytes(bytes)?;
        let page = session.sheet_page(SessionSheetPageParams {
            sheet_name: "Sheet1".to_string(),
            start_row: 2,
            page_size: 1,
            columns_by_header: Some(vec!["score".to_string()]),
            format: SheetPageFormat::Compact,
            ..SessionSheetPageParams::default()
        })?;

        let compact = page.compact.as_ref().expect("compact payload");
        assert_eq!(compact.headers, vec!["Row", "Score"]);
        assert_eq!(compact.rows.len(), 1);
        assert!(
            matches!(compact.rows[0][0], Some(CellValue::Number(n)) if (n - 2.0).abs() < f64::EPSILON)
        );
        assert!(
            matches!(compact.rows[0][1], Some(CellValue::Number(n)) if (n - 99.0).abs() < f64::EPSILON)
        );

        Ok(())
    }

    #[test]
    fn sheet_page_rejects_invalid_column_specs() -> Result<()> {
        let bytes = workbook_bytes(|book| {
            book.get_sheet_by_name_mut("Sheet1")
                .expect("sheet")
                .get_cell_mut("A1")
                .set_value("x");
        });

        let session = WorkbookSession::from_bytes(bytes)?;
        let err = session
            .sheet_page(SessionSheetPageParams {
                sheet_name: "Sheet1".to_string(),
                columns: Some(vec!["1".to_string()]),
                ..SessionSheetPageParams::default()
            })
            .expect_err("invalid columns should error");
        assert!(err.to_string().contains("invalid column spec"));

        Ok(())
    }

    #[test]
    fn sheet_page_handles_large_start_rows_without_overflow() -> Result<()> {
        let bytes = workbook_bytes(|book| {
            book.get_sheet_by_name_mut("Sheet1")
                .expect("sheet")
                .get_cell_mut("A1")
                .set_value("x");
        });

        let session = WorkbookSession::from_bytes(bytes)?;
        let page = session.sheet_page(SessionSheetPageParams {
            sheet_name: "Sheet1".to_string(),
            start_row: u32::MAX,
            page_size: 500,
            format: SheetPageFormat::ValuesOnly,
            ..SessionSheetPageParams::default()
        })?;

        assert_eq!(page.next_start_row, None);

        Ok(())
    }

    #[test]
    fn describe_and_named_ranges_return_session_metadata() -> Result<()> {
        let bytes = workbook_bytes(|book| {
            let sheet = book.get_sheet_by_name_mut("Sheet1").expect("sheet");
            sheet.get_cell_mut("A1").set_value("Revenue");
            sheet.get_cell_mut("A2").set_value_number(100.0);
            sheet
                .add_defined_name("TotalRevenue", "Sheet1!$A$2")
                .expect("defined name");
        });

        let session = WorkbookSession::from_bytes(bytes)?;
        let desc = session.describe_workbook()?;
        assert!(matches!(desc.workbook_id, WorkbookId(ref id) if id == "session"));
        assert!(desc.sheet_count >= 1);

        let named = session.named_ranges()?;
        assert!(matches!(named.workbook_id, WorkbookId(ref id) if id == "session"));
        assert!(named.items.iter().all(|item| !item.name.trim().is_empty()));

        Ok(())
    }

    #[test]
    fn sheet_overview_applies_region_and_header_limits() -> Result<()> {
        let bytes = workbook_bytes(|book| {
            let sheet = book.get_sheet_by_name_mut("Sheet1").expect("sheet");
            sheet.get_cell_mut("A1").set_value("Name");
            sheet.get_cell_mut("B1").set_value("Score");
            sheet.get_cell_mut("A2").set_value("alpha");
            sheet.get_cell_mut("B2").set_value_number(10.0);
            sheet.get_cell_mut("A10").set_value("Name");
            sheet.get_cell_mut("B10").set_value("Score");
            sheet.get_cell_mut("A11").set_value("beta");
            sheet.get_cell_mut("B11").set_value_number(20.0);
        });

        let session = WorkbookSession::from_bytes(bytes)?;
        let overview = session.sheet_overview(SessionSheetOverviewParams {
            sheet_name: "Sheet1".to_string(),
            max_regions: Some(1),
            max_headers: Some(1),
            include_headers: Some(true),
        })?;

        assert_eq!(overview.sheet_name, "Sheet1");
        assert!(overview.detected_region_count >= overview.detected_regions.len() as u32);
        assert!(overview.detected_regions.len() <= 1);

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
