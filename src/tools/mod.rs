pub mod filters;
#[cfg(feature = "recalc")]
pub mod fork;

use crate::analysis::{formula::FormulaGraph, stats};
use crate::model::*;
use crate::state::AppState;
use crate::workbook::{WorkbookContext, cell_to_value};
use anyhow::{Result, anyhow};
use regex::Regex;
use schemars::JsonSchema;
use serde::Deserialize;
use std::cmp::Ordering;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

const DEFAULT_TRACE_PAGE_SIZE: usize = 20;
const TRACE_PAGE_MIN: usize = 5;
const TRACE_PAGE_MAX: usize = 200;
const TRACE_RANGE_THRESHOLD: usize = 4;
const TRACE_RANGE_HIGHLIGHT_LIMIT: usize = 3;
const TRACE_GROUP_HIGHLIGHT_LIMIT: usize = 3;
const TRACE_CELL_HIGHLIGHT_LIMIT: usize = 5;
const TRACE_RANGE_VALUE_SAMPLES: usize = 3;
const TRACE_RANGE_FORMULA_SAMPLES: usize = 2;
const TRACE_GROUP_SAMPLE_LIMIT: usize = 5;
const TRACE_DEPENDENTS_PER_CELL_LIMIT: usize = 500;

pub async fn list_workbooks(
    state: Arc<AppState>,
    params: ListWorkbooksParams,
) -> Result<WorkbookListResponse> {
    let filter = params.into_filter()?;
    state.list_workbooks(filter)
}

pub async fn describe_workbook(
    state: Arc<AppState>,
    params: DescribeWorkbookParams,
) -> Result<WorkbookDescription> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let desc = workbook.describe();
    Ok(desc)
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ListWorkbooksParams {
    pub slug_prefix: Option<String>,
    pub folder: Option<String>,
    pub path_glob: Option<String>,
}

impl ListWorkbooksParams {
    fn into_filter(self) -> Result<filters::WorkbookFilter> {
        filters::WorkbookFilter::new(self.slug_prefix, self.folder, self.path_glob)
    }
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DescribeWorkbookParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ListSheetsParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
}

pub async fn list_sheets(
    state: Arc<AppState>,
    params: ListSheetsParams,
) -> Result<SheetListResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let summaries = workbook.list_summaries()?;
    let response = SheetListResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        sheets: summaries,
    };
    Ok(response)
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct SheetOverviewParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub sheet_name: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct WorkbookSummaryParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
}

pub async fn workbook_summary(
    state: Arc<AppState>,
    params: WorkbookSummaryParams,
) -> Result<WorkbookSummaryResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let sheet_names = workbook.sheet_names();

    let mut total_cells: u64 = 0;
    let mut total_formulas: u64 = 0;
    let mut breakdown = WorkbookBreakdown::default();
    let mut region_counts = RegionCountSummary::default();
    let mut entry_points: Vec<EntryPoint> = Vec::new();
    let mut key_named_ranges: Vec<NamedRangeDescriptor> = Vec::new();

    for sheet_name in &sheet_names {
        let entry = workbook.get_sheet_metrics(sheet_name)?;
        total_cells += (entry.metrics.row_count as u64) * (entry.metrics.column_count as u64);
        total_formulas += entry.metrics.formula_cells as u64;
        match entry.metrics.classification {
            SheetClassification::Calculator => breakdown.calculator_sheets += 1,
            SheetClassification::Metadata => breakdown.metadata_sheets += 1,
            SheetClassification::Empty => {}
            _ => breakdown.data_sheets += 1,
        }

        for region in &entry.detected_regions {
            match region
                .region_kind
                .clone()
                .unwrap_or(region.classification.clone())
            {
                RegionKind::Calculator => region_counts.calculator += 1,
                RegionKind::Metadata => region_counts.metadata += 1,
                RegionKind::Parameters => region_counts.parameters += 1,
                RegionKind::Outputs => region_counts.outputs += 1,
                RegionKind::Data | RegionKind::Table => region_counts.data += 1,
                _ => region_counts.other += 1,
            }
            if region.confidence >= 0.3 {
                let kind = region
                    .region_kind
                    .as_ref()
                    .unwrap_or(&region.classification);
                let priority = match kind {
                    RegionKind::Parameters => 0,
                    RegionKind::Data | RegionKind::Table => 1,
                    RegionKind::Outputs => 2,
                    RegionKind::Calculator => 3,
                    RegionKind::Metadata => 4,
                    _ => 5,
                };
                entry_points.push(EntryPoint {
                    sheet_name: sheet_name.clone(),
                    region_id: Some(region.id),
                    bounds: Some(region.bounds.clone()),
                    rationale: format!(
                        "{:?} region ({} rows, {:.0}% conf, p{})",
                        kind,
                        region.row_count,
                        region.confidence * 100.0,
                        priority
                    ),
                });
            }
        }

        if entry.detected_regions.is_empty() && entry.metrics.non_empty_cells > 0 {
            entry_points.push(EntryPoint {
                sheet_name: sheet_name.clone(),
                region_id: None,
                bounds: None,
                rationale: "Whole sheet is non-empty; start at top-left".to_string(),
            });
        }
    }

    entry_points.sort_by(|a, b| {
        let pa = priority_from_rationale(&a.rationale);
        let pb = priority_from_rationale(&b.rationale);
        pa.cmp(&pb)
            .then_with(|| {
                a.bounds
                    .as_ref()
                    .map(|_| 1)
                    .cmp(&b.bounds.as_ref().map(|_| 1))
            })
            .then_with(|| a.sheet_name.cmp(&b.sheet_name))
    });
    entry_points.truncate(5);

    let mut seen_ranges = std::collections::HashSet::new();
    for item in workbook.named_items()? {
        if item.kind != NamedItemKind::NamedRange && item.kind != NamedItemKind::Table {
            continue;
        }
        if !seen_ranges.insert(item.refers_to.clone()) {
            continue;
        }
        key_named_ranges.push(item);
        if key_named_ranges.len() >= 10 {
            break;
        }
    }

    Ok(WorkbookSummaryResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        slug: workbook.slug.clone(),
        sheet_count: sheet_names.len(),
        total_cells,
        total_formulas,
        breakdown,
        region_counts,
        key_named_ranges,
        suggested_entry_points: entry_points,
    })
}

fn priority_from_rationale(rationale: &str) -> u32 {
    if rationale.contains("p0") {
        0
    } else if rationale.contains("p1") {
        1
    } else if rationale.contains("p2") {
        2
    } else if rationale.contains("p3") {
        3
    } else if rationale.contains("p4") {
        4
    } else {
        5
    }
}

pub async fn sheet_overview(
    state: Arc<AppState>,
    params: SheetOverviewParams,
) -> Result<SheetOverviewResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let overview = workbook.sheet_overview(&params.sheet_name)?;
    Ok(overview)
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

fn default_include_header() -> bool {
    true
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct SheetPageParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
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
    #[serde(default)]
    pub include_styles: bool,
    #[serde(default = "default_include_header")]
    pub include_header: bool,
    #[serde(default)]
    pub format: Option<SheetPageFormat>,
}

impl Default for SheetPageParams {
    fn default() -> Self {
        SheetPageParams {
            workbook_or_fork_id: WorkbookId(String::new()),
            sheet_name: String::new(),
            start_row: default_start_row(),
            page_size: default_page_size(),
            columns: None,
            columns_by_header: None,
            include_formulas: default_include_formulas(),
            include_styles: false,
            include_header: default_include_header(),
            format: None,
        }
    }
}

fn default_find_limit() -> u32 {
    50
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct FindValueParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub query: String,
    #[serde(default)]
    pub label: Option<String>,
    #[serde(default)]
    pub mode: Option<FindMode>,
    #[serde(default)]
    pub match_mode: Option<String>,
    #[serde(default)]
    pub case_sensitive: bool,
    #[serde(default)]
    pub sheet_name: Option<String>,
    #[serde(default)]
    pub region_id: Option<u32>,
    #[serde(default)]
    pub table_name: Option<String>,
    #[serde(default)]
    pub value_types: Option<Vec<String>>,
    #[serde(default)]
    pub search_headers_only: bool,
    #[serde(default)]
    pub direction: Option<LabelDirection>,
    #[serde(default = "default_find_limit")]
    pub limit: u32,
}

impl Default for FindValueParams {
    fn default() -> Self {
        Self {
            workbook_or_fork_id: WorkbookId(String::new()),
            query: String::new(),
            label: None,
            mode: None,
            match_mode: None,
            case_sensitive: false,
            sheet_name: None,
            region_id: None,
            table_name: None,
            value_types: None,
            search_headers_only: false,
            direction: None,
            limit: default_find_limit(),
        }
    }
}

#[derive(Debug, Deserialize, JsonSchema, Default)]
pub struct ReadTableParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    #[serde(default)]
    pub sheet_name: Option<String>,
    #[serde(default)]
    pub table_name: Option<String>,
    #[serde(default)]
    pub region_id: Option<u32>,
    #[serde(default)]
    pub range: Option<String>,
    #[serde(default)]
    pub header_row: Option<u32>,
    #[serde(default)]
    pub header_rows: Option<u32>,
    #[serde(default)]
    pub columns: Option<Vec<String>>,
    #[serde(default)]
    pub filters: Option<Vec<TableFilter>>,
    #[serde(default)]
    pub sample_mode: Option<String>,
    #[serde(default)]
    pub limit: Option<u32>,
    #[serde(default)]
    pub offset: Option<u32>,
}

#[derive(Debug, Deserialize, JsonSchema, Clone)]
pub struct TableFilter {
    pub column: String,
    pub op: String,
    pub value: serde_json::Value,
}

#[derive(Debug, Deserialize, JsonSchema, Default)]
pub struct TableProfileParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    #[serde(default)]
    pub sheet_name: Option<String>,
    #[serde(default)]
    pub region_id: Option<u32>,
    #[serde(default)]
    pub table_name: Option<String>,
    #[serde(default)]
    pub sample_mode: Option<String>,
    #[serde(default)]
    pub sample_size: Option<u32>,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct RangeValuesParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub sheet_name: String,
    pub ranges: Vec<String>,
    #[serde(default)]
    pub include_headers: Option<bool>,
}

pub async fn sheet_page(
    state: Arc<AppState>,
    params: SheetPageParams,
) -> Result<SheetPageResponse> {
    if params.page_size == 0 {
        return Err(anyhow!("page_size must be greater than zero"));
    }

    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let metrics = workbook.get_sheet_metrics(&params.sheet_name)?;
    let format = params.format.unwrap_or_default();

    let start_row = params.start_row.max(1);
    let page_size = params.page_size.min(500);
    let include_formulas = params.include_formulas;
    let include_styles = params.include_styles;
    let columns = params.columns.clone();
    let columns_by_header = params.columns_by_header.clone();
    let include_header = params.include_header;

    let page = workbook.with_sheet(&params.sheet_name, |sheet| {
        build_page(
            sheet,
            start_row,
            page_size,
            columns.clone(),
            columns_by_header.clone(),
            include_formulas,
            include_styles,
            include_header,
        )
    })?;

    let has_more = page.end_row < metrics.metrics.row_count;
    let next_start_row = if has_more {
        Some(page.end_row + 1)
    } else {
        None
    };

    let compact_payload = if matches!(format, SheetPageFormat::Compact) {
        Some(build_compact_payload(
            &page.header,
            &page.rows,
            include_header,
        ))
    } else {
        None
    };

    let values_only_payload = if matches!(format, SheetPageFormat::ValuesOnly) {
        Some(build_values_only_payload(
            &page.header,
            &page.rows,
            include_header,
        ))
    } else {
        None
    };

    let response = SheetPageResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        sheet_name: params.sheet_name,
        rows: if matches!(format, SheetPageFormat::Full) {
            page.rows
        } else {
            Vec::new()
        },
        has_more,
        next_start_row,
        header_row: if include_header { page.header } else { None },
        compact: compact_payload,
        values_only: values_only_payload,
        format,
    };
    Ok(response)
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct SheetFormulaMapParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub sheet_name: String,
    pub range: Option<String>,
    #[serde(default)]
    pub expand: bool,
    #[serde(default)]
    pub limit: Option<u32>,
    #[serde(default)]
    pub sort_by: Option<FormulaSortBy>,
}

#[derive(Debug, Clone, Copy, Default, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum FormulaSortBy {
    #[default]
    Address,
    Complexity,
    Count,
}

pub async fn sheet_formula_map(
    state: Arc<AppState>,
    params: SheetFormulaMapParams,
) -> Result<SheetFormulaMapResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let graph = workbook.formula_graph(&params.sheet_name)?;
    let mut groups = Vec::new();
    let mut truncated = false;

    for mut group in graph.groups() {
        if let Some(range) = &params.range {
            group.addresses.retain(|addr| address_in_range(addr, range));
            if group.addresses.is_empty() {
                continue;
            }
        }
        if !params.expand && group.addresses.len() > 15 {
            group.addresses.truncate(15);
            truncated = true;
        }
        groups.push(group);
    }

    let sort_by = params.sort_by.unwrap_or_default();
    match sort_by {
        FormulaSortBy::Address => {
            groups.sort_by(|a, b| a.addresses.first().cmp(&b.addresses.first()));
        }
        FormulaSortBy::Complexity => {
            groups.sort_by(|a, b| b.formula.len().cmp(&a.formula.len()));
        }
        FormulaSortBy::Count => {
            groups.sort_by(|a, b| b.addresses.len().cmp(&a.addresses.len()));
        }
    }

    if let Some(limit) = params.limit
        && groups.len() > limit as usize
    {
        groups.truncate(limit as usize);
        truncated = true;
    }

    let response = SheetFormulaMapResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        sheet_name: params.sheet_name.clone(),
        groups,
        truncated,
    };
    Ok(response)
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct FormulaTraceParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub sheet_name: String,
    pub cell_address: String,
    pub direction: TraceDirection,
    pub depth: Option<u32>,
    pub limit: Option<u32>,
    #[serde(default)]
    pub page_size: Option<usize>,
    #[serde(default)]
    pub cursor: Option<TraceCursor>,
}

pub async fn formula_trace(
    state: Arc<AppState>,
    params: FormulaTraceParams,
) -> Result<FormulaTraceResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let graph = workbook.formula_graph(&params.sheet_name)?;
    let formula_lookup = build_formula_lookup(&graph);
    let depth = params.depth.unwrap_or(3).clamp(1, 5);
    let page_size = params
        .page_size
        .or_else(|| params.limit.map(|v| v as usize))
        .unwrap_or(DEFAULT_TRACE_PAGE_SIZE)
        .clamp(TRACE_PAGE_MIN, TRACE_PAGE_MAX);

    let origin = params.cell_address.to_uppercase();
    let config = TraceConfig {
        direction: &params.direction,
        origin: &origin,
        sheet_name: &params.sheet_name,
        depth_limit: depth,
        page_size,
    };
    let (layers, next_cursor, notes) = build_trace_layers(
        &workbook,
        &graph,
        &formula_lookup,
        &config,
        params.cursor.clone(),
    )?;

    let response = FormulaTraceResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        sheet_name: params.sheet_name.clone(),
        origin,
        direction: params.direction.clone(),
        layers,
        next_cursor,
        notes,
    };
    Ok(response)
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct NamedRangesParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub sheet_name: Option<String>,
    pub name_prefix: Option<String>,
}

pub async fn named_ranges(
    state: Arc<AppState>,
    params: NamedRangesParams,
) -> Result<NamedRangesResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let mut items = workbook.named_items()?;

    if let Some(sheet_filter) = &params.sheet_name {
        items.retain(|item| {
            item.sheet_name
                .as_ref()
                .map(|name| name.eq_ignore_ascii_case(sheet_filter))
                .unwrap_or(false)
        });
    }
    if let Some(prefix) = &params.name_prefix {
        let prefix_lower = prefix.to_ascii_lowercase();
        items.retain(|item| item.name.to_ascii_lowercase().starts_with(&prefix_lower));
    }

    let response = NamedRangesResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        items,
    };
    Ok(response)
}

struct PageBuildResult {
    rows: Vec<RowSnapshot>,
    header: Option<RowSnapshot>,
    end_row: u32,
}

#[allow(clippy::too_many_arguments)]
fn build_page(
    sheet: &umya_spreadsheet::Worksheet,
    start_row: u32,
    page_size: u32,
    columns: Option<Vec<String>>,
    columns_by_header: Option<Vec<String>>,
    include_formulas: bool,
    include_styles: bool,
    include_header: bool,
) -> PageBuildResult {
    let max_col = sheet.get_highest_column();
    let end_row = (start_row + page_size - 1).min(sheet.get_highest_row().max(start_row));
    let column_indices =
        resolve_columns_with_headers(sheet, columns.as_ref(), columns_by_header.as_ref(), max_col);

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

    PageBuildResult {
        rows,
        header,
        end_row,
    }
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

fn resolve_columns(columns: Option<&Vec<String>>, max_column: u32) -> Vec<u32> {
    use std::collections::BTreeSet;
    use umya_spreadsheet::helper::coordinate::column_index_from_string;

    let mut indices = BTreeSet::new();
    if let Some(specs) = columns {
        for spec in specs {
            if let Some((start, end)) = spec.split_once(':') {
                let start_idx = column_index_from_string(start);
                let end_idx = column_index_from_string(end);
                let (min_idx, max_idx) = if start_idx <= end_idx {
                    (start_idx, end_idx)
                } else {
                    (end_idx, start_idx)
                };
                for idx in min_idx..=max_idx {
                    indices.insert(idx);
                }
            } else {
                indices.insert(column_index_from_string(spec));
            }
        }
    } else {
        for idx in 1..=max_column.max(1) {
            indices.insert(idx);
        }
    }

    indices.into_iter().collect()
}

fn resolve_columns_with_headers(
    sheet: &umya_spreadsheet::Worksheet,
    columns: Option<&Vec<String>>,
    columns_by_header: Option<&Vec<String>>,
    max_column: u32,
) -> Vec<u32> {
    if columns_by_header.is_none() {
        return resolve_columns(columns, max_column);
    }

    let mut selected = Vec::new();
    let header_targets: Vec<String> = columns_by_header
        .unwrap()
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
            selected.push(col_idx);
        }
    }

    if selected.is_empty() {
        resolve_columns(columns, max_column)
    } else {
        selected
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
fn default_stats_sample() -> usize {
    500
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct SheetStatisticsParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub sheet_name: String,
    #[serde(default)]
    pub sample_rows: Option<usize>,
}

pub async fn sheet_statistics(
    state: Arc<AppState>,
    params: SheetStatisticsParams,
) -> Result<SheetStatisticsResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let sheet_metrics = workbook.get_sheet_metrics(&params.sheet_name)?;
    let sample_rows = params.sample_rows.unwrap_or_else(default_stats_sample);
    let stats = workbook.with_sheet(&params.sheet_name, |sheet| {
        stats::compute_sheet_statistics(sheet, sample_rows)
    })?;
    let response = SheetStatisticsResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        sheet_name: params.sheet_name,
        row_count: sheet_metrics.metrics.row_count,
        column_count: sheet_metrics.metrics.column_count,
        density: stats.density,
        numeric_columns: stats.numeric_columns,
        text_columns: stats.text_columns,
        null_counts: stats.null_counts,
        duplicate_warnings: stats.duplicate_warnings,
    };
    Ok(response)
}

fn address_in_range(address: &str, range: &str) -> bool {
    parse_range(range).is_none_or(|((start_col, start_row), (end_col, end_row))| {
        if let Some((col, row)) = parse_address(address) {
            col >= start_col && col <= end_col && row >= start_row && row <= end_row
        } else {
            false
        }
    })
}

fn parse_range(range: &str) -> Option<((u32, u32), (u32, u32))> {
    let mut parts = range.split(':');
    let start = parts.next()?;
    let end = parts.next().unwrap_or(start);
    let start_idx = parse_address(start)?;
    let end_idx = parse_address(end)?;
    Some((
        (start_idx.0.min(end_idx.0), start_idx.1.min(end_idx.1)),
        (start_idx.0.max(end_idx.0), start_idx.1.max(end_idx.1)),
    ))
}

fn parse_address(address: &str) -> Option<(u32, u32)> {
    use umya_spreadsheet::helper::coordinate::index_from_coordinate;
    let (col, row, _, _) = index_from_coordinate(address);
    match (col, row) {
        (Some(c), Some(r)) => Some((c, r)),
        _ => None,
    }
}

#[derive(Clone)]
struct TableTarget {
    sheet_name: String,
    table_name: Option<String>,
    range: ((u32, u32), (u32, u32)),
    header_hint: Option<u32>,
}

fn resolve_table_target(
    workbook: &WorkbookContext,
    params: &ReadTableParams,
) -> Result<TableTarget> {
    if let Some(region_id) = params.region_id
        && let Some(sheet) = &params.sheet_name
        && let Ok(region) = workbook.detected_region(sheet, region_id)
    {
        return Ok(TableTarget {
            sheet_name: sheet.clone(),
            table_name: None,
            range: parse_range(&region.bounds).unwrap_or(((1, 1), (1, 1))),
            header_hint: region.header_row,
        });
    }

    if let Some(table_name) = &params.table_name {
        let items = workbook.named_items()?;
        for item in items {
            if item.name.eq_ignore_ascii_case(table_name)
                || item
                    .name
                    .to_ascii_lowercase()
                    .contains(&table_name.to_ascii_lowercase())
            {
                let mut sheet_name = item
                    .sheet_name
                    .clone()
                    .or_else(|| params.sheet_name.clone())
                    .unwrap_or_else(|| workbook.sheet_names().first().cloned().unwrap_or_default());
                let refers_to = item.refers_to.trim_start_matches('=');
                let mut range_part = refers_to;
                if let Some((sheet_part, rest)) = refers_to.split_once('!') {
                    sheet_name = sheet_part.trim_matches('\'').to_string();
                    range_part = rest;
                }
                if let Some(range) = parse_range(range_part) {
                    return Ok(TableTarget {
                        sheet_name,
                        table_name: Some(item.name.clone()),
                        range,
                        header_hint: if item.kind == NamedItemKind::Table {
                            Some(range.0.1)
                        } else {
                            None
                        },
                    });
                }
            }
        }
    }

    let sheet_name = params
        .sheet_name
        .clone()
        .unwrap_or_else(|| workbook.sheet_names().first().cloned().unwrap_or_default());

    if let Some(rng) = &params.range
        && let Some(range) = parse_range(rng)
    {
        return Ok(TableTarget {
            sheet_name,
            table_name: None,
            range,
            header_hint: None,
        });
    }

    let metrics = workbook.get_sheet_metrics(&sheet_name)?;
    let end_col = metrics.metrics.column_count.max(1);
    let end_row = metrics.metrics.row_count.max(1);
    Ok(TableTarget {
        sheet_name,
        table_name: None,
        range: ((1, 1), (end_col, end_row)),
        header_hint: None,
    })
}

#[allow(clippy::too_many_arguments)]
fn extract_table_rows(
    sheet: &umya_spreadsheet::Worksheet,
    target: &TableTarget,
    header_row: Option<u32>,
    header_rows: Option<u32>,
    columns: Option<Vec<String>>,
    filters: Option<Vec<TableFilter>>,
    limit: usize,
    offset: usize,
    sample_mode: &str,
) -> Result<(Vec<String>, Vec<TableRow>, u32)> {
    let ((start_col, start_row), (end_col, end_row)) = target.range;
    let mut header_start = header_row.or(target.header_hint).unwrap_or(start_row);
    if header_start < start_row {
        header_start = start_row;
    }
    if header_start > end_row {
        header_start = start_row;
    }
    let header_rows_count = header_rows.unwrap_or(1).max(1);
    let data_start_row = (header_start + header_rows_count).max(start_row + header_rows_count);
    let column_indices: Vec<u32> = if let Some(cols) = columns.as_ref() {
        resolve_columns(Some(cols), end_col).into_iter().collect()
    } else {
        (start_col..=end_col).collect()
    };

    let headers = build_headers(sheet, &column_indices, header_start, header_rows_count);
    let mut all_rows: Vec<TableRow> = Vec::new();
    let mut total_rows: u32 = 0;

    for row_idx in data_start_row..=end_row {
        let mut row = BTreeMap::new();
        for (i, col_idx) in column_indices.iter().enumerate() {
            let header = headers
                .get(i)
                .cloned()
                .unwrap_or_else(|| format!("Col{col_idx}"));
            let value = sheet.get_cell((*col_idx, row_idx)).and_then(cell_to_value);
            row.insert(header, value);
        }
        if !row_passes_filters(&row, filters.as_ref()) {
            continue;
        }
        total_rows += 1;
        if matches!(sample_mode, "first" | "all") && total_rows as usize > offset + limit {
            continue;
        }
        all_rows.push(row);
    }

    let rows = sample_rows(all_rows, limit, offset, sample_mode);

    Ok((headers, rows, total_rows))
}

fn build_headers(
    sheet: &umya_spreadsheet::Worksheet,
    columns: &[u32],
    header_start: u32,
    header_rows: u32,
) -> Vec<String> {
    let mut headers = Vec::new();
    for col_idx in columns {
        let mut parts = Vec::new();
        for h in header_start..(header_start + header_rows) {
            let (origin_col, origin_row) = sheet.map_merged_cell((*col_idx, h));
            if let Some(value) = sheet
                .get_cell((origin_col, origin_row))
                .and_then(cell_to_value)
            {
                match value {
                    CellValue::Text(ref s) if s.trim().is_empty() => {}
                    CellValue::Text(s) => parts.push(s),
                    CellValue::Number(n) => parts.push(n.to_string()),
                    CellValue::Bool(b) => parts.push(b.to_string()),
                    CellValue::Error(e) => parts.push(e),
                    CellValue::Date(d) => parts.push(d),
                }
            }
        }
        if parts.is_empty() {
            headers.push(crate::utils::column_number_to_name(*col_idx));
        } else {
            headers.push(parts.join(" / "));
        }
    }

    if headers.iter().all(|h| h.trim().is_empty()) {
        return columns
            .iter()
            .map(|c| crate::utils::column_number_to_name(*c))
            .collect();
    }

    dedupe_headers(headers)
}

fn dedupe_headers(mut headers: Vec<String>) -> Vec<String> {
    let mut seen: HashMap<String, u32> = HashMap::new();
    for h in headers.iter_mut() {
        let key = h.clone();
        if key.trim().is_empty() {
            continue;
        }
        let count = seen.entry(key.clone()).or_insert(0);
        if *count > 0 {
            h.push_str(&format!("_{}", *count + 1));
        }
        *count += 1;
    }
    headers
}

fn row_passes_filters(row: &TableRow, filters: Option<&Vec<TableFilter>>) -> bool {
    if let Some(filters) = filters {
        for filter in filters {
            if let Some(value) = row.get(&filter.column) {
                match filter.op.as_str() {
                    "eq" => {
                        if !value_eq(value, &filter.value) {
                            return false;
                        }
                    }
                    "neq" => {
                        if value_eq(value, &filter.value) {
                            return false;
                        }
                    }
                    "contains" => {
                        if !value_contains(value, &filter.value) {
                            return false;
                        }
                    }
                    "gt" => {
                        if !value_gt(value, &filter.value) {
                            return false;
                        }
                    }
                    "lt" => {
                        if !value_lt(value, &filter.value) {
                            return false;
                        }
                    }
                    "in" => {
                        let list = filter
                            .value
                            .as_array()
                            .cloned()
                            .unwrap_or_else(|| vec![filter.value.clone()]);
                        if !list.iter().any(|cmp| value_eq(value, cmp)) {
                            return false;
                        }
                    }
                    _ => {}
                }
            }
        }
    }
    true
}

fn value_eq(cell: &Option<CellValue>, cmp: &serde_json::Value) -> bool {
    match (cell, cmp) {
        (Some(CellValue::Text(s)), serde_json::Value::String(t)) => s == t,
        (Some(CellValue::Number(n)), serde_json::Value::Number(v)) => {
            v.as_f64().is_some_and(|f| (*n - f).abs() < f64::EPSILON)
        }
        (Some(CellValue::Number(n)), serde_json::Value::String(t)) => t
            .parse::<f64>()
            .map(|f| (*n - f).abs() < f64::EPSILON)
            .unwrap_or(false),
        (Some(CellValue::Bool(b)), serde_json::Value::Bool(v)) => b == v,
        (Some(CellValue::Bool(b)), serde_json::Value::String(t)) => {
            t.eq_ignore_ascii_case("true") == *b
        }
        (Some(CellValue::Date(d)), serde_json::Value::String(t)) => d == t,
        (None, serde_json::Value::Null) => true,
        _ => false,
    }
}

fn value_contains(cell: &Option<CellValue>, cmp: &serde_json::Value) -> bool {
    if let (Some(CellValue::Text(s)), serde_json::Value::String(t)) = (cell, cmp) {
        return s.to_ascii_lowercase().contains(&t.to_ascii_lowercase());
    }
    false
}

fn value_gt(cell: &Option<CellValue>, cmp: &serde_json::Value) -> bool {
    match (cell, cmp) {
        (Some(CellValue::Number(n)), serde_json::Value::Number(v)) => {
            v.as_f64().is_some_and(|f| *n > f)
        }
        _ => false,
    }
}

fn value_lt(cell: &Option<CellValue>, cmp: &serde_json::Value) -> bool {
    match (cell, cmp) {
        (Some(CellValue::Number(n)), serde_json::Value::Number(v)) => {
            v.as_f64().is_some_and(|f| *n < f)
        }
        _ => false,
    }
}

fn sample_rows(rows: Vec<TableRow>, limit: usize, offset: usize, mode: &str) -> Vec<TableRow> {
    if rows.is_empty() {
        return rows;
    }

    match mode {
        "distributed" => {
            if limit == 0 {
                return Vec::new();
            }
            let mut indices = Vec::new();
            let span = rows.len().saturating_sub(1);
            let step = std::cmp::max(1, span / std::cmp::max(1, limit.saturating_sub(1)));
            let mut idx = offset;
            while idx < rows.len() && indices.len() < limit {
                indices.push(idx);
                idx = idx.saturating_add(step);
                if idx == indices.last().copied().unwrap_or(0) {
                    idx += 1;
                }
            }
            if indices.len() < limit {
                let last_idx = rows.len().saturating_sub(1);
                if !indices.contains(&last_idx) {
                    indices.push(last_idx);
                }
            }
            indices
                .into_iter()
                .filter_map(|i| rows.get(i).cloned())
                .collect()
        }
        "last" => {
            let start = rows.len().saturating_sub(limit + offset);
            rows.into_iter().skip(start + offset).take(limit).collect()
        }
        _ => rows.into_iter().skip(offset).take(limit).collect(),
    }
}

fn summarize_columns(headers: &[String], rows: &[TableRow]) -> Vec<ColumnTypeSummary> {
    let mut summaries = Vec::new();
    for header in headers {
        let mut nulls = 0u32;
        let mut distinct_set: HashSet<String> = HashSet::new();
        let mut values: Vec<f64> = Vec::new();
        let mut top_counts: HashMap<String, u32> = HashMap::new();

        for row in rows {
            match row.get(header) {
                Some(Some(CellValue::Number(n))) => {
                    values.push(*n);
                    let key = n.to_string();
                    *top_counts.entry(key).or_default() += 1;
                }
                Some(Some(CellValue::Text(s))) => {
                    distinct_set.insert(s.clone());
                    *top_counts.entry(s.clone()).or_default() += 1;
                }
                Some(Some(CellValue::Bool(b))) => {
                    let key = b.to_string();
                    distinct_set.insert(key.clone());
                    *top_counts.entry(key).or_default() += 1;
                }
                Some(Some(CellValue::Date(d))) => {
                    distinct_set.insert(d.clone());
                    *top_counts.entry(d.clone()).or_default() += 1;
                }
                Some(Some(CellValue::Error(e))) => {
                    distinct_set.insert(e.clone());
                    *top_counts.entry(e.clone()).or_default() += 1;
                }
                _ => {
                    nulls += 1;
                }
            }
        }

        let inferred_type = if !values.is_empty() {
            "number"
        } else if !distinct_set.is_empty() {
            "text"
        } else {
            "unknown"
        }
        .to_string();

        let min = values.iter().cloned().reduce(f64::min);
        let max = values.iter().cloned().reduce(f64::max);
        let mean = if values.is_empty() {
            None
        } else {
            Some(values.iter().sum::<f64>() / values.len() as f64)
        };

        let mut top_values: Vec<(String, u32)> = top_counts.into_iter().collect();
        top_values.sort_by(|a, b| b.1.cmp(&a.1));
        let top_values = top_values.into_iter().take(3).map(|(v, _)| v).collect();

        summaries.push(ColumnTypeSummary {
            name: header.clone(),
            inferred_type,
            nulls,
            distinct: distinct_set.len() as u32,
            top_values,
            min,
            max,
            mean,
        });
    }
    summaries
}

#[allow(clippy::too_many_arguments)]
fn collect_value_matches(
    sheet: &umya_spreadsheet::Worksheet,
    sheet_name: &str,
    mode: &FindMode,
    match_mode: &str,
    direction: &LabelDirection,
    params: &FindValueParams,
    region: Option<&DetectedRegion>,
    default_bounds: ((u32, u32), (u32, u32)),
) -> Result<Vec<FindValueMatch>> {
    let mut results = Vec::new();
    let regex = if match_mode == "regex" {
        Regex::new(&params.query).ok()
    } else {
        None
    };
    let bounds = region
        .as_ref()
        .and_then(|r| parse_range(&r.bounds))
        .unwrap_or(default_bounds);

    let header_row = region.and_then(|r| r.header_row).unwrap_or(1);

    for cell in sheet.get_cell_collection() {
        let coord = cell.get_coordinate();
        let col = *coord.get_col_num();
        let row = *coord.get_row_num();
        if col < bounds.0.0 || col > bounds.1.0 || row < bounds.0.1 || row > bounds.1.1 {
            continue;
        }
        if params.search_headers_only && row != header_row {
            continue;
        }

        let value = cell_to_value(cell);
        if let Some(ref allowed) = params.value_types
            && !value_type_matches(&value, allowed)
        {
            continue;
        }
        if matches!(mode, FindMode::Value) {
            if !value_matches(
                &value,
                &params.query,
                match_mode,
                params.case_sensitive,
                &regex,
            ) {
                continue;
            }
        } else if let Some(label) = &params.label {
            if !label_matches(cell, label, match_mode, params.case_sensitive, &regex) {
                continue;
            }
        } else {
            continue;
        }

        let neighbors = collect_neighbors(sheet, row, col);
        let (label_hit, match_value) = if matches!(mode, FindMode::Label) {
            let target_value = match direction {
                LabelDirection::Right => sheet.get_cell((col + 1, row)),
                LabelDirection::Below => sheet.get_cell((col, row + 1)),
                LabelDirection::Any => sheet
                    .get_cell((col + 1, row))
                    .or_else(|| sheet.get_cell((col, row + 1))),
            }
            .and_then(cell_to_value);
            if target_value.is_none() {
                continue;
            }
            (
                Some(LabelHit {
                    label_address: coord.get_coordinate(),
                    label: label_from_cell(cell),
                }),
                target_value,
            )
        } else {
            (None, value.clone())
        };

        let row_context = build_row_context(sheet, row, col);

        results.push(FindValueMatch {
            address: coord.get_coordinate(),
            sheet_name: sheet_name.to_string(),
            value: match_value,
            row_context,
            neighbors,
            label_hit,
        });
    }

    Ok(results)
}

fn label_from_cell(cell: &umya_spreadsheet::Cell) -> String {
    cell_to_value(cell)
        .map(|v| match v {
            CellValue::Text(s) => s,
            CellValue::Number(n) => n.to_string(),
            CellValue::Bool(b) => b.to_string(),
            CellValue::Date(d) => d,
            CellValue::Error(e) => e,
        })
        .unwrap_or_else(|| cell.get_value().to_string())
}

fn value_matches(
    value: &Option<CellValue>,
    query: &str,
    mode: &str,
    case_sensitive: bool,
    regex: &Option<Regex>,
) -> bool {
    if value.is_none() {
        return false;
    }
    let haystack = cell_value_to_string_lower(value.clone().unwrap());
    let needle = if case_sensitive {
        query.to_string()
    } else {
        query.to_ascii_lowercase()
    };

    match mode {
        "exact" => haystack == needle,
        "regex" => regex
            .as_ref()
            .map(|re| re.is_match(&haystack))
            .unwrap_or(false),
        _ => haystack.contains(&needle),
    }
}

fn label_matches(
    cell: &umya_spreadsheet::Cell,
    label: &str,
    mode: &str,
    case_sensitive: bool,
    regex: &Option<Regex>,
) -> bool {
    let value = cell_to_value(cell);
    if value.is_none() {
        return false;
    }
    let haystack = cell_value_to_string_lower(value.unwrap());
    let needle = if case_sensitive {
        label.to_string()
    } else {
        label.to_ascii_lowercase()
    };
    match mode {
        "exact" => haystack == needle,
        "regex" => regex
            .as_ref()
            .map(|re| re.is_match(&haystack))
            .unwrap_or(false),
        _ => haystack.contains(&needle),
    }
}

fn value_type_matches(value: &Option<CellValue>, allowed: &[String]) -> bool {
    if value.is_none() {
        return allowed.iter().any(|v| v == "null");
    }
    match value.as_ref().unwrap() {
        CellValue::Text(_) => allowed.iter().any(|v| v.eq_ignore_ascii_case("text")),
        CellValue::Number(_) => allowed.iter().any(|v| v.eq_ignore_ascii_case("number")),
        CellValue::Bool(_) => allowed.iter().any(|v| v.eq_ignore_ascii_case("bool")),
        CellValue::Date(_) => allowed.iter().any(|v| v.eq_ignore_ascii_case("date")),
        CellValue::Error(_) => true,
    }
}

fn collect_neighbors(
    sheet: &umya_spreadsheet::Worksheet,
    row: u32,
    col: u32,
) -> Option<NeighborValues> {
    Some(NeighborValues {
        left: if col > 1 {
            sheet.get_cell((col - 1, row)).and_then(cell_to_value)
        } else {
            None
        },
        right: sheet.get_cell((col + 1, row)).and_then(cell_to_value),
        up: if row > 1 {
            sheet.get_cell((col, row - 1)).and_then(cell_to_value)
        } else {
            None
        },
        down: sheet.get_cell((col, row + 1)).and_then(cell_to_value),
    })
}

fn build_row_context(
    sheet: &umya_spreadsheet::Worksheet,
    row: u32,
    col: u32,
) -> Option<RowContext> {
    let header_value = sheet
        .get_cell((col, 1u32))
        .and_then(cell_to_value)
        .map(|v| match v {
            CellValue::Text(s) => s,
            CellValue::Number(n) => n.to_string(),
            CellValue::Bool(b) => b.to_string(),
            CellValue::Date(d) => d,
            CellValue::Error(e) => e,
        })
        .unwrap_or_else(|| format!("Col{}", col));
    let value = sheet.get_cell((col, row)).and_then(cell_to_value);
    Some(RowContext {
        headers: vec![header_value],
        values: vec![value],
    })
}

fn default_find_formula_limit() -> u32 {
    50
}

#[derive(Debug, Deserialize, JsonSchema, Default)]
pub struct FindFormulaParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub query: String,
    pub sheet_name: Option<String>,
    #[serde(default)]
    pub case_sensitive: bool,
    #[serde(default)]
    pub include_context: bool,
    #[serde(default = "default_find_formula_limit")]
    pub limit: u32,
    #[serde(default)]
    pub offset: u32,
}

pub async fn find_formula(
    state: Arc<AppState>,
    params: FindFormulaParams,
) -> Result<FindFormulaResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let query = if params.case_sensitive {
        params.query.clone()
    } else {
        params.query.to_ascii_lowercase()
    };

    let sheet_names: Vec<String> = if let Some(sheet) = &params.sheet_name {
        vec![sheet.clone()]
    } else {
        workbook.sheet_names()
    };

    let limit = params.limit.clamp(1, 500);
    let offset = params.offset;

    let mut matches = Vec::new();
    let mut seen: u32 = 0;
    let mut truncated = false;

    for sheet_name in sheet_names {
        let (sheet_matches, sheet_seen, sheet_truncated) =
            workbook.with_sheet(&sheet_name, |sheet| {
                collect_formula_matches(
                    sheet,
                    &sheet_name,
                    &query,
                    params.case_sensitive,
                    params.include_context,
                    offset,
                    limit,
                    seen,
                )
            })?;

        seen = sheet_seen;
        truncated |= sheet_truncated;
        matches.extend(sheet_matches);

        if truncated {
            break;
        }
    }

    let response = FindFormulaResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        matches,
        truncated,
        next_offset: truncated.then(|| offset.saturating_add(limit)),
    };
    Ok(response)
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ScanVolatilesParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub sheet_name: Option<String>,
}

pub async fn scan_volatiles(
    state: Arc<AppState>,
    params: ScanVolatilesParams,
) -> Result<VolatileScanResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let target_sheets: Vec<String> = if let Some(sheet) = &params.sheet_name {
        vec![sheet.clone()]
    } else {
        workbook.sheet_names()
    };

    let mut items = Vec::new();
    let mut truncated = false;

    for sheet_name in target_sheets {
        let graph = workbook.formula_graph(&sheet_name)?;
        for group in graph.groups() {
            if !group.is_volatile {
                continue;
            }
            for address in group.addresses.iter().take(50) {
                items.push(VolatileScanEntry {
                    address: address.clone(),
                    sheet_name: sheet_name.clone(),
                    function: "volatile".to_string(),
                    note: Some(group.formula.clone()),
                });
            }
            if group.addresses.len() > 50 {
                truncated = true;
            }
        }
    }

    let response = VolatileScanResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        items,
        truncated,
    };
    Ok(response)
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct WorkbookStyleSummaryParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub max_styles: Option<u32>,
    pub max_conditional_formats: Option<u32>,
    pub max_cells_scan: Option<u32>,
}

#[derive(Debug)]
struct WorkbookStyleAccum {
    descriptor: StyleDescriptor,
    occurrences: u32,
    tags: HashSet<String>,
    example_cells: Vec<String>,
}

impl WorkbookStyleAccum {
    fn new(descriptor: StyleDescriptor) -> Self {
        Self {
            descriptor,
            occurrences: 0,
            tags: HashSet::new(),
            example_cells: Vec::new(),
        }
    }
}

pub async fn workbook_style_summary(
    state: Arc<AppState>,
    params: WorkbookStyleSummaryParams,
) -> Result<WorkbookStyleSummaryResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let sheet_names = workbook.sheet_names();

    const STYLE_EXAMPLE_LIMIT: usize = 5;
    const STYLE_LIMIT_DEFAULT: usize = 200;
    const CF_LIMIT_DEFAULT: usize = 200;
    const CELL_SCAN_LIMIT_DEFAULT: usize = 500_000;

    let style_limit = params
        .max_styles
        .map(|v| v as usize)
        .unwrap_or(STYLE_LIMIT_DEFAULT);
    let cf_limit = params
        .max_conditional_formats
        .map(|v| v as usize)
        .unwrap_or(CF_LIMIT_DEFAULT);
    let cell_scan_limit = params
        .max_cells_scan
        .map(|v| v as usize)
        .unwrap_or(CELL_SCAN_LIMIT_DEFAULT);

    let mut acc: HashMap<String, WorkbookStyleAccum> = HashMap::new();
    let mut scanned_cells: usize = 0;
    let mut scan_truncated = false;

    for sheet_name in &sheet_names {
        if scan_truncated {
            break;
        }
        workbook.with_sheet(sheet_name, |sheet| {
            for cell in sheet.get_cell_collection() {
                if scanned_cells >= cell_scan_limit {
                    scan_truncated = true;
                    break;
                }
                scanned_cells += 1;

                let address = cell.get_coordinate().get_coordinate().to_string();
                let descriptor = crate::styles::descriptor_from_style(cell.get_style());
                let style_id = crate::styles::stable_style_id(&descriptor);

                let entry = acc
                    .entry(style_id.clone())
                    .or_insert_with(|| WorkbookStyleAccum::new(descriptor.clone()));
                entry.occurrences += 1;
                if entry.example_cells.len() < STYLE_EXAMPLE_LIMIT {
                    entry.example_cells.push(format!("{sheet_name}!{address}"));
                }

                if let Some((_, tagging)) = crate::analysis::style::tag_cell(cell) {
                    for tag in tagging.tags {
                        entry.tags.insert(tag);
                    }
                }
            }
        })?;
    }

    let total_styles = acc.len() as u32;
    let mut styles: Vec<WorkbookStyleUsage> = acc
        .into_iter()
        .map(|(style_id, entry)| {
            let mut tags: Vec<String> = entry.tags.into_iter().collect();
            tags.sort();
            WorkbookStyleUsage {
                style_id,
                occurrences: entry.occurrences,
                tags,
                example_cells: entry.example_cells,
                descriptor: Some(entry.descriptor),
            }
        })
        .collect();

    styles.sort_by(|a, b| {
        b.occurrences
            .cmp(&a.occurrences)
            .then_with(|| a.style_id.cmp(&b.style_id))
    });

    let inferred_default_style_id = styles.first().map(|s| s.style_id.clone());
    let mut inferred_default_font = styles
        .first()
        .and_then(|s| s.descriptor.as_ref().and_then(|d| d.font.clone()));

    let styles_truncated = if styles.len() > style_limit {
        styles.truncate(style_limit);
        true
    } else {
        false
    };

    let theme = workbook.with_spreadsheet(|book| {
        let theme = book.get_theme();
        let elements = theme.get_theme_elements();
        let scheme = elements.get_color_scheme();
        let mut colors = BTreeMap::new();

        let mut insert_color = |name: &str, value: String| {
            if !value.trim().is_empty() {
                colors.insert(name.to_string(), value);
            }
        };

        insert_color("dk1", scheme.get_dk1().get_val());
        insert_color("lt1", scheme.get_lt1().get_val());
        insert_color("dk2", scheme.get_dk2().get_val());
        insert_color("lt2", scheme.get_lt2().get_val());
        insert_color("accent1", scheme.get_accent1().get_val());
        insert_color("accent2", scheme.get_accent2().get_val());
        insert_color("accent3", scheme.get_accent3().get_val());
        insert_color("accent4", scheme.get_accent4().get_val());
        insert_color("accent5", scheme.get_accent5().get_val());
        insert_color("accent6", scheme.get_accent6().get_val());
        insert_color("hlink", scheme.get_hlink().get_val());
        insert_color("fol_hlink", scheme.get_fol_hlink().get_val());

        let font_scheme = elements.get_font_scheme();
        let major = font_scheme.get_major_font();
        let minor = font_scheme.get_minor_font();
        let font_scheme_summary = ThemeFontSchemeSummary {
            major_latin: Some(major.get_latin_font().get_typeface().to_string())
                .filter(|s| !s.trim().is_empty()),
            major_east_asian: Some(major.get_east_asian_font().get_typeface().to_string())
                .filter(|s| !s.trim().is_empty()),
            major_complex_script: Some(major.get_complex_script_font().get_typeface().to_string())
                .filter(|s| !s.trim().is_empty()),
            minor_latin: Some(minor.get_latin_font().get_typeface().to_string())
                .filter(|s| !s.trim().is_empty()),
            minor_east_asian: Some(minor.get_east_asian_font().get_typeface().to_string())
                .filter(|s| !s.trim().is_empty()),
            minor_complex_script: Some(minor.get_complex_script_font().get_typeface().to_string())
                .filter(|s| !s.trim().is_empty()),
        };

        ThemeSummary {
            name: Some(theme.get_name().to_string()).filter(|s| !s.trim().is_empty()),
            colors,
            font_scheme: font_scheme_summary,
        }
    })?;

    if inferred_default_font.is_none()
        && let Some(name) = theme
            .font_scheme
            .minor_latin
            .clone()
            .or_else(|| theme.font_scheme.major_latin.clone())
    {
        inferred_default_font = Some(FontDescriptor {
            name: Some(name),
            size: None,
            bold: None,
            italic: None,
            underline: None,
            strikethrough: None,
            color: None,
        });
    }

    let mut conditional_formats: Vec<ConditionalFormatSummary> = Vec::new();
    let mut conditional_formats_truncated = false;
    {
        use umya_spreadsheet::structs::EnumTrait;
        for sheet_name in &sheet_names {
            if conditional_formats_truncated {
                break;
            }
            workbook.with_sheet(sheet_name, |sheet| {
                for cf in sheet.get_conditional_formatting_collection() {
                    if conditional_formats.len() >= cf_limit {
                        conditional_formats_truncated = true;
                        break;
                    }
                    let range = cf.get_sequence_of_references().get_sqref().to_string();
                    let mut types: HashSet<String> = HashSet::new();
                    for rule in cf.get_conditional_collection() {
                        types.insert(rule.get_type().get_value_string().to_string());
                    }
                    let mut rule_types: Vec<String> = types.into_iter().collect();
                    rule_types.sort();
                    conditional_formats.push(ConditionalFormatSummary {
                        sheet_name: sheet_name.clone(),
                        range,
                        rule_types,
                        rule_count: cf.get_conditional_collection().len() as u32,
                    });
                }
            })?;
        }
    }

    let mut notes: Vec<String> = Vec::new();
    if scan_truncated {
        notes.push(format!(
            "Stopped scanning after {cell_scan_limit} cells; style counts may be incomplete."
        ));
    }
    notes.push(
        "Named styles are not directly exposed by umya-spreadsheet; styles here are inferred from cell formatting."
            .to_string(),
    );

    Ok(WorkbookStyleSummaryResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        theme: Some(theme),
        inferred_default_style_id,
        inferred_default_font,
        styles,
        total_styles,
        styles_truncated,
        conditional_formats,
        conditional_formats_truncated,
        scan_truncated,
        notes,
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct SheetStylesParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub sheet_name: String,
    #[serde(default)]
    pub scope: Option<SheetStylesScope>,
    #[serde(default)]
    pub granularity: Option<String>,
    #[serde(default)]
    pub max_items: Option<usize>,
}

#[derive(Debug, Deserialize, JsonSchema)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum SheetStylesScope {
    Range { range: String },
    Region { region_id: u32 },
}

#[derive(Debug)]
struct StyleAccum {
    descriptor: StyleDescriptor,
    occurrences: u32,
    tags: HashSet<String>,
    example_cells: Vec<String>,
    positions: Vec<(u32, u32)>,
}

impl StyleAccum {
    fn new(descriptor: StyleDescriptor) -> Self {
        Self {
            descriptor,
            occurrences: 0,
            tags: HashSet::new(),
            example_cells: Vec::new(),
            positions: Vec::new(),
        }
    }
}

pub async fn sheet_styles(
    state: Arc<AppState>,
    params: SheetStylesParams,
) -> Result<SheetStylesResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    const STYLE_EXAMPLE_LIMIT: usize = 5;
    const STYLE_RANGE_LIMIT: usize = 50;
    const STYLE_LIMIT: usize = 200;
    const MAX_MAX_ITEMS: usize = 5000;

    let metrics = workbook.get_sheet_metrics(&params.sheet_name)?;
    let full_bounds = (
        (1, 1),
        (
            metrics.metrics.column_count.max(1),
            metrics.metrics.row_count.max(1),
        ),
    );

    let bounds = match &params.scope {
        Some(SheetStylesScope::Range { range }) => {
            parse_range(range).ok_or_else(|| anyhow!("invalid range: {}", range))?
        }
        Some(SheetStylesScope::Region { region_id }) => {
            let region = workbook.detected_region(&params.sheet_name, *region_id)?;
            parse_range(&region.bounds)
                .ok_or_else(|| anyhow!("invalid region bounds: {}", region.bounds))?
        }
        None => full_bounds,
    };

    let granularity = params
        .granularity
        .as_deref()
        .unwrap_or("runs")
        .to_ascii_lowercase();
    if granularity != "runs" && granularity != "cells" {
        return Err(anyhow!(
            "invalid granularity: {} (expected runs|cells)",
            granularity
        ));
    }

    let max_items = params
        .max_items
        .unwrap_or(STYLE_RANGE_LIMIT)
        .clamp(1, MAX_MAX_ITEMS);

    let (styles, total_styles, styles_truncated) =
        workbook.with_sheet(&params.sheet_name, |sheet| {
            let mut acc: HashMap<String, StyleAccum> = HashMap::new();

            for cell in sheet.get_cell_collection() {
                let address = cell.get_coordinate().get_coordinate().to_string();
                let Some((col, row)) = parse_address(&address) else {
                    continue;
                };
                if col < bounds.0.0 || col > bounds.1.0 || row < bounds.0.1 || row > bounds.1.1 {
                    continue;
                }

                let descriptor = crate::styles::descriptor_from_style(cell.get_style());
                let style_id = crate::styles::stable_style_id(&descriptor);

                let entry = acc
                    .entry(style_id.clone())
                    .or_insert_with(|| StyleAccum::new(descriptor.clone()));
                entry.occurrences += 1;
                if entry.example_cells.len() < STYLE_EXAMPLE_LIMIT {
                    entry.example_cells.push(address.clone());
                }

                if let Some((_, tagging)) = crate::analysis::style::tag_cell(cell) {
                    for tag in tagging.tags {
                        entry.tags.insert(tag);
                    }
                }

                entry.positions.push((row, col));
            }

            let mut summaries: Vec<StyleSummary> = acc
                .into_iter()
                .map(|(style_id, mut entry)| {
                    entry.positions.sort_unstable();
                    entry.positions.dedup();

                    let (cell_ranges, ranges_truncated) = if granularity == "cells" {
                        let mut out = Vec::new();
                        for (row, col) in entry.positions.iter().take(max_items) {
                            out.push(crate::utils::cell_address(*col, *row));
                        }
                        (out, entry.positions.len() > max_items)
                    } else {
                        crate::styles::compress_positions_to_ranges(&entry.positions, max_items)
                    };

                    StyleSummary {
                        style_id,
                        occurrences: entry.occurrences,
                        tags: entry.tags.into_iter().collect(),
                        example_cells: entry.example_cells,
                        descriptor: Some(entry.descriptor),
                        cell_ranges,
                        ranges_truncated,
                    }
                })
                .collect();

            summaries.sort_by(|a, b| {
                b.occurrences
                    .cmp(&a.occurrences)
                    .then_with(|| a.style_id.cmp(&b.style_id))
            });

            let total = summaries.len() as u32;
            let truncated = if summaries.len() > STYLE_LIMIT {
                summaries.truncate(STYLE_LIMIT);
                true
            } else {
                false
            };

            Ok::<_, anyhow::Error>((summaries, total, truncated))
        })??;

    Ok(SheetStylesResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        sheet_name: params.sheet_name.clone(),
        styles,
        conditional_rules: Vec::new(),
        total_styles,
        styles_truncated,
    })
}

pub async fn range_values(
    state: Arc<AppState>,
    params: RangeValuesParams,
) -> Result<RangeValuesResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let include_headers = params.include_headers.unwrap_or(false);
    let values = workbook.with_sheet(&params.sheet_name, |sheet| {
        params
            .ranges
            .iter()
            .filter_map(|range| {
                parse_range(range).map(|((start_col, start_row), (end_col, end_row))| {
                    let mut rows = Vec::new();
                    for r in start_row..=end_row {
                        let mut row_vals = Vec::new();
                        for c in start_col..=end_col {
                            if include_headers && r == start_row && start_row == 1 {
                                row_vals.push(sheet.get_cell((c, 1u32)).and_then(cell_to_value));
                            } else {
                                row_vals.push(sheet.get_cell((c, r)).and_then(cell_to_value));
                            }
                        }
                        rows.push(row_vals);
                    }
                    RangeValuesEntry {
                        range: range.clone(),
                        rows,
                    }
                })
            })
            .collect()
    })?;

    Ok(RangeValuesResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        sheet_name: params.sheet_name,
        values,
    })
}

pub async fn find_value(
    state: Arc<AppState>,
    params: FindValueParams,
) -> Result<FindValueResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let mut matches = Vec::new();
    let mut truncated = false;
    let mode = params.mode.clone().unwrap_or_else(|| {
        if params.label.is_some() {
            FindMode::Label
        } else {
            FindMode::Value
        }
    });
    let match_mode = params
        .match_mode
        .as_deref()
        .unwrap_or("contains")
        .to_ascii_lowercase();
    let direction = params.direction.clone().unwrap_or(LabelDirection::Any);

    let target_sheets: Vec<String> = if let Some(sheet) = &params.sheet_name {
        vec![sheet.clone()]
    } else {
        workbook.sheet_names()
    };

    for sheet_name in target_sheets {
        let metrics_entry = workbook.get_sheet_metrics(&sheet_name)?;
        let default_bounds = (
            (1, 1),
            (
                metrics_entry.metrics.column_count.max(1),
                metrics_entry.metrics.row_count.max(1),
            ),
        );
        let region_bounds = params
            .region_id
            .and_then(|id| workbook.detected_region(&sheet_name, id).ok());
        let sheet_matches = workbook.with_sheet(&sheet_name, |sheet| {
            collect_value_matches(
                sheet,
                &sheet_name,
                &mode,
                &match_mode,
                &direction,
                &params,
                region_bounds.as_ref(),
                default_bounds,
            )
        })??;
        matches.extend(sheet_matches);
        if matches.len() as u32 >= params.limit {
            truncated = true;
            matches.truncate(params.limit as usize);
            break;
        }
    }

    Ok(FindValueResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        matches,
        truncated,
    })
}

pub async fn read_table(
    state: Arc<AppState>,
    params: ReadTableParams,
) -> Result<ReadTableResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let resolved = resolve_table_target(&workbook, &params)?;
    let limit = params.limit.unwrap_or(100) as usize;
    let offset = params.offset.unwrap_or(0) as usize;
    let sample_mode = params
        .sample_mode
        .clone()
        .unwrap_or_else(|| "first".to_string());

    let (headers, rows, total_rows) = workbook.with_sheet(&resolved.sheet_name, |sheet| {
        extract_table_rows(
            sheet,
            &resolved,
            params.header_row,
            params.header_rows,
            params.columns.clone(),
            params.filters.clone(),
            limit,
            offset,
            &sample_mode,
        )
    })??;

    let has_more = offset + rows.len() < total_rows as usize;

    Ok(ReadTableResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        sheet_name: resolved.sheet_name,
        table_name: resolved.table_name,
        headers,
        rows,
        total_rows,
        has_more,
    })
}

pub async fn table_profile(
    state: Arc<AppState>,
    params: TableProfileParams,
) -> Result<TableProfileResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let resolved = resolve_table_target(
        &workbook,
        &ReadTableParams {
            workbook_or_fork_id: params.workbook_or_fork_id.clone(),
            sheet_name: params.sheet_name.clone(),
            table_name: params.table_name.clone(),
            region_id: params.region_id,
            range: None,
            header_row: None,
            header_rows: None,
            columns: None,
            filters: None,
            sample_mode: params.sample_mode.clone(),
            limit: params.sample_size,
            offset: Some(0),
        },
    )?;

    let sample_size = params.sample_size.unwrap_or(10) as usize;
    let sample_mode = params
        .sample_mode
        .clone()
        .unwrap_or_else(|| "distributed".to_string());

    let (headers, rows, total_rows) = workbook.with_sheet(&resolved.sheet_name, |sheet| {
        extract_table_rows(
            sheet,
            &resolved,
            None,
            None,
            None,
            None,
            sample_size,
            0,
            &sample_mode,
        )
    })??;

    let column_types = summarize_columns(&headers, &rows);

    Ok(TableProfileResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        sheet_name: resolved.sheet_name,
        table_name: resolved.table_name,
        headers,
        column_types,
        row_count: total_rows,
        samples: rows,
        notes: Vec::new(),
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ManifestStubParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
    pub sheet_filter: Option<String>,
}

pub async fn get_manifest_stub(
    state: Arc<AppState>,
    params: ManifestStubParams,
) -> Result<ManifestStubResponse> {
    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
    let mut summaries = workbook.list_summaries()?;

    if let Some(filter) = &params.sheet_filter {
        summaries.retain(|summary| summary.name.eq_ignore_ascii_case(filter));
    }

    let sheets = summaries
        .into_iter()
        .map(|summary| ManifestSheetStub {
            sheet_name: summary.name.clone(),
            classification: summary.classification.clone(),
            candidate_expectations: vec![format!(
                "Review {} sheet for expectation candidates",
                format!("{:?}", summary.classification).to_ascii_lowercase()
            )],
            notes: summary.style_tags,
        })
        .collect();

    let response = ManifestStubResponse {
        workbook_id: workbook.id.clone(),
        workbook_short_id: workbook.short_id.clone(),
        slug: workbook.slug.clone(),
        sheets,
    };
    Ok(response)
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct CloseWorkbookParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
}

pub async fn close_workbook(
    state: Arc<AppState>,
    params: CloseWorkbookParams,
) -> Result<CloseWorkbookResponse> {
    state.close_workbook(&params.workbook_or_fork_id)?;
    Ok(CloseWorkbookResponse {
        workbook_id: params.workbook_or_fork_id.clone(),
        message: format!("workbook {} evicted", params.workbook_or_fork_id.as_str()),
    })
}
#[allow(clippy::too_many_arguments)]
fn collect_formula_matches(
    sheet: &umya_spreadsheet::Worksheet,
    sheet_name: &str,
    query: &str,
    case_sensitive: bool,
    include_context: bool,
    offset: u32,
    limit: u32,
    seen_so_far: u32,
) -> (Vec<FindFormulaMatch>, u32, bool) {
    use crate::workbook::cell_to_value;

    let mut results = Vec::new();
    let mut seen = seen_so_far;

    for cell in sheet.get_cell_collection() {
        if !cell.is_formula() {
            continue;
        }
        let formula = cell.get_formula();
        let haystack = if case_sensitive {
            formula.to_string()
        } else {
            formula.to_ascii_lowercase()
        };
        if !haystack.contains(query) {
            continue;
        }

        if seen < offset {
            seen += 1;
            continue;
        }

        if results.len() as u32 >= limit {
            return (results, seen, true);
        }

        let coord = cell.get_coordinate();
        let column = *coord.get_col_num();
        let row = *coord.get_row_num();

        let context = if include_context {
            let columns = vec![column];
            let context_row = build_row_snapshot(sheet, row, &columns, true, false);
            let header_row = build_row_snapshot(sheet, 1, &columns, false, false);
            vec![header_row, context_row]
        } else {
            Vec::new()
        };

        results.push(FindFormulaMatch {
            address: coord.get_coordinate(),
            sheet_name: sheet_name.to_string(),
            formula: formula.to_string(),
            cached_value: cell_to_value(cell),
            context,
        });

        seen += 1;
    }

    (results, seen, false)
}

#[derive(Clone)]
struct TraceFormulaInfo {
    fingerprint: String,
    formula: String,
}

#[derive(Clone)]
struct TraceEdgeRaw {
    from: String,
    to: String,
    neighbor: String,
}

#[derive(Clone)]
struct LayerLinks {
    depth: u32,
    edges: Vec<TraceEdgeRaw>,
    truncated_cells: Vec<String>,
}

#[derive(Clone)]
struct NeighborDetail {
    address: String,
    column: Option<u32>,
    row: Option<u32>,
    kind: TraceCellKind,
    value: Option<CellValue>,
    formula: Option<String>,
    fingerprint: Option<String>,
    external: bool,
}

fn build_formula_lookup(graph: &FormulaGraph) -> HashMap<String, TraceFormulaInfo> {
    let mut map = HashMap::new();
    for group in graph.groups() {
        for address in group.addresses.clone() {
            map.insert(
                address.to_ascii_uppercase(),
                TraceFormulaInfo {
                    fingerprint: group.fingerprint.clone(),
                    formula: group.formula.clone(),
                },
            );
        }
    }
    map
}

struct TraceConfig<'a> {
    direction: &'a TraceDirection,
    origin: &'a str,
    sheet_name: &'a str,
    depth_limit: u32,
    page_size: usize,
}

fn build_trace_layers(
    workbook: &WorkbookContext,
    graph: &FormulaGraph,
    formula_lookup: &HashMap<String, TraceFormulaInfo>,
    config: &TraceConfig<'_>,
    cursor: Option<TraceCursor>,
) -> Result<(Vec<TraceLayer>, Option<TraceCursor>, Vec<String>)> {
    let layer_links =
        collect_layer_links(graph, config.direction, config.origin, config.depth_limit);
    let mut layers = Vec::new();
    let mut next_cursor = None;
    let mut notes = Vec::new();
    let focus_depth = cursor.as_ref().map(|c| c.depth);

    for layer in layer_links {
        let produce_edges = focus_depth.is_none_or(|depth| depth == layer.depth);
        let offset = cursor
            .as_ref()
            .filter(|c| c.depth == layer.depth)
            .map(|c| c.offset)
            .unwrap_or(0);

        let mut node_set: HashSet<String> = HashSet::new();
        for edge in &layer.edges {
            node_set.insert(edge.neighbor.clone());
        }
        let mut nodes: Vec<String> = node_set.into_iter().collect();
        nodes.sort_by(|a, b| compare_addresses(a, b));

        let details = workbook.with_sheet(config.sheet_name, |sheet| {
            collect_neighbor_details(sheet, config.sheet_name, &nodes, formula_lookup)
        })?;
        let total_nodes = details.len();
        let start = offset.min(total_nodes);
        let end = if produce_edges {
            (start + config.page_size).min(total_nodes)
        } else {
            start
        };
        let selected_slice = if produce_edges {
            &details[start..end]
        } else {
            &details[0..0]
        };
        let selected_addresses: HashSet<String> = selected_slice
            .iter()
            .map(|detail| detail.address.clone())
            .collect();

        let summary = build_layer_summary(&details);
        let range_highlights = build_range_highlights(&details);
        let group_highlights = build_formula_group_highlights(&details);
        let notable_cells = build_notable_cells(&details, &range_highlights, &group_highlights);

        let highlights = TraceLayerHighlights {
            top_ranges: range_highlights.clone(),
            top_formula_groups: group_highlights.clone(),
            notable_cells,
        };

        let edges = if produce_edges {
            build_edges_for_layer(&layer.edges, &selected_addresses, formula_lookup)
        } else {
            Vec::new()
        };

        let has_more = produce_edges && end < total_nodes;
        if has_more && next_cursor.is_none() {
            next_cursor = Some(TraceCursor {
                depth: layer.depth,
                offset: end,
            });
        }
        if has_more {
            notes.push(format!(
                "Layer {} truncated at {} of {} nodes; supply cursor.depth={} and cursor.offset={} to continue",
                layer.depth, end, total_nodes, layer.depth, end
            ));
        }

        if !layer.truncated_cells.is_empty() {
            let cell_list = if layer.truncated_cells.len() <= 3 {
                layer.truncated_cells.join(", ")
            } else {
                format!(
                    "{}, ... ({} more)",
                    layer.truncated_cells[..3].join(", "),
                    layer.truncated_cells.len() - 3
                )
            };
            notes.push(format!(
                "Layer {}: dependents truncated at {} per cell for: {}",
                layer.depth, TRACE_DEPENDENTS_PER_CELL_LIMIT, cell_list
            ));
        }

        layers.push(TraceLayer {
            depth: layer.depth,
            summary,
            highlights,
            edges,
            has_more,
        });
    }

    Ok((layers, next_cursor, notes))
}

fn collect_layer_links(
    graph: &FormulaGraph,
    direction: &TraceDirection,
    origin: &str,
    depth_limit: u32,
) -> Vec<LayerLinks> {
    let mut visited: HashSet<String> = HashSet::new();
    visited.insert(origin.to_string());
    let mut frontier = vec![origin.to_string()];
    let mut layers = Vec::new();

    for depth in 1..=depth_limit {
        let mut next_frontier_set: HashSet<String> = HashSet::new();
        let mut edges = Vec::new();
        let mut truncated_cells = Vec::new();

        for cell in &frontier {
            let (neighbors, was_truncated) = match direction {
                TraceDirection::Precedents => (graph.precedents(cell), false),
                TraceDirection::Dependents => {
                    graph.dependents_limited(cell, Some(TRACE_DEPENDENTS_PER_CELL_LIMIT))
                }
            };

            if was_truncated {
                truncated_cells.push(cell.clone());
            }

            for neighbor in neighbors {
                let neighbor_upper = neighbor.to_ascii_uppercase();
                let edge = match direction {
                    TraceDirection::Precedents => TraceEdgeRaw {
                        from: cell.clone(),
                        to: neighbor_upper.clone(),
                        neighbor: neighbor_upper.clone(),
                    },
                    TraceDirection::Dependents => TraceEdgeRaw {
                        from: neighbor_upper.clone(),
                        to: cell.clone(),
                        neighbor: neighbor_upper.clone(),
                    },
                };
                edges.push(edge);
                if visited.insert(neighbor_upper.clone()) {
                    next_frontier_set.insert(neighbor_upper);
                }
            }
        }

        if edges.is_empty() {
            break;
        }

        layers.push(LayerLinks {
            depth,
            edges,
            truncated_cells,
        });
        if next_frontier_set.is_empty() {
            break;
        }
        let mut next_frontier: Vec<String> = next_frontier_set.into_iter().collect();
        next_frontier.sort();
        frontier = next_frontier;
    }

    layers
}

fn collect_neighbor_details(
    sheet: &umya_spreadsheet::Worksheet,
    current_sheet: &str,
    addresses: &[String],
    formula_lookup: &HashMap<String, TraceFormulaInfo>,
) -> Vec<NeighborDetail> {
    let mut details = Vec::new();
    for address in addresses {
        let (sheet_part, cell_part) = split_sheet_and_cell(address);
        let normalized_sheet = sheet_part
            .as_ref()
            .map(|s| clean_sheet_name(s).to_ascii_lowercase());
        let is_external = normalized_sheet
            .as_ref()
            .map(|s| !s.eq_ignore_ascii_case(current_sheet))
            .unwrap_or(false);

        let Some(cell_ref) = cell_part else {
            details.push(NeighborDetail {
                address: address.clone(),
                column: None,
                row: None,
                kind: TraceCellKind::External,
                value: None,
                formula: None,
                fingerprint: None,
                external: true,
            });
            continue;
        };

        let cell_ref_upper = cell_ref.to_ascii_uppercase();

        if is_external {
            let formula_info = lookup_formula_info(formula_lookup, &cell_ref_upper, address);
            details.push(NeighborDetail {
                address: address.clone(),
                column: None,
                row: None,
                kind: TraceCellKind::External,
                value: None,
                formula: formula_info.map(|info| info.formula.clone()),
                fingerprint: formula_info.map(|info| info.fingerprint.clone()),
                external: true,
            });
            continue;
        }

        let Some((col, row)) = parse_address(&cell_ref_upper) else {
            details.push(NeighborDetail {
                address: address.clone(),
                column: None,
                row: None,
                kind: TraceCellKind::External,
                value: None,
                formula: None,
                fingerprint: None,
                external: true,
            });
            continue;
        };

        let cell_opt = sheet.get_cell((&col, &row));
        let formula_info = lookup_formula_info(formula_lookup, &cell_ref_upper, address);
        if let Some(cell) = cell_opt {
            let value = cell_to_value(cell);
            let kind = if cell.is_formula() {
                TraceCellKind::Formula
            } else if value.is_some() {
                TraceCellKind::Literal
            } else {
                TraceCellKind::Blank
            };
            details.push(NeighborDetail {
                address: address.clone(),
                column: Some(col),
                row: Some(row),
                kind,
                value,
                formula: formula_info.map(|info| info.formula.clone()),
                fingerprint: formula_info.map(|info| info.fingerprint.clone()),
                external: false,
            });
        } else {
            details.push(NeighborDetail {
                address: address.clone(),
                column: Some(col),
                row: Some(row),
                kind: TraceCellKind::Blank,
                value: None,
                formula: formula_info.map(|info| info.formula.clone()),
                fingerprint: formula_info.map(|info| info.fingerprint.clone()),
                external: false,
            });
        }
    }
    details
}

fn build_layer_summary(details: &[NeighborDetail]) -> TraceLayerSummary {
    let mut summary = TraceLayerSummary {
        total_nodes: details.len(),
        formula_nodes: 0,
        value_nodes: 0,
        blank_nodes: 0,
        external_nodes: 0,
        unique_formula_groups: 0,
    };

    let mut fingerprints: HashSet<String> = HashSet::new();

    for detail in details {
        match detail.kind {
            TraceCellKind::Formula => {
                summary.formula_nodes += 1;
                if let Some(fp) = &detail.fingerprint {
                    fingerprints.insert(fp.clone());
                }
            }
            TraceCellKind::Literal => summary.value_nodes += 1,
            TraceCellKind::Blank => summary.blank_nodes += 1,
            TraceCellKind::External => summary.external_nodes += 1,
        }
    }

    summary.unique_formula_groups = fingerprints.len();
    summary
}

fn build_formula_group_highlights(details: &[NeighborDetail]) -> Vec<TraceFormulaGroupHighlight> {
    let mut aggregates: HashMap<String, (String, usize, Vec<String>)> = HashMap::new();
    for detail in details {
        if let (Some(fingerprint), Some(formula)) = (&detail.fingerprint, &detail.formula) {
            let entry = aggregates
                .entry(fingerprint.clone())
                .or_insert_with(|| (formula.clone(), 0, Vec::new()));
            entry.1 += 1;
            if entry.2.len() < TRACE_GROUP_SAMPLE_LIMIT {
                entry.2.push(detail.address.clone());
            }
        }
    }

    let mut highlights: Vec<TraceFormulaGroupHighlight> = aggregates
        .into_iter()
        .map(
            |(fingerprint, (formula, count, sample_addresses))| TraceFormulaGroupHighlight {
                fingerprint,
                formula,
                count,
                sample_addresses,
            },
        )
        .collect();

    highlights.sort_by(|a, b| b.count.cmp(&a.count));
    highlights.truncate(TRACE_GROUP_HIGHLIGHT_LIMIT);
    highlights
}

fn build_range_highlights(details: &[NeighborDetail]) -> Vec<TraceRangeHighlight> {
    let mut by_column: HashMap<u32, Vec<&NeighborDetail>> = HashMap::new();
    for detail in details {
        if let (Some(col), Some(_row)) = (detail.column, detail.row)
            && !detail.external
        {
            by_column.entry(col).or_default().push(detail);
        }
    }

    for column_entries in by_column.values_mut() {
        column_entries.sort_by(|a, b| a.row.cmp(&b.row));
    }

    let mut ranges = Vec::new();
    for entries in by_column.values() {
        let mut current: Vec<&NeighborDetail> = Vec::new();
        for detail in entries {
            if current.is_empty() {
                current.push(detail);
                continue;
            }
            let prev_row = current.last().and_then(|d| d.row).unwrap_or(0);
            if detail.row.unwrap_or(0) == prev_row + 1 {
                current.push(detail);
            } else {
                if current.len() >= TRACE_RANGE_THRESHOLD {
                    ranges.push(make_range_highlight(&current));
                }
                current.clear();
                current.push(detail);
            }
        }
        if current.len() >= TRACE_RANGE_THRESHOLD {
            ranges.push(make_range_highlight(&current));
        }
    }

    ranges.sort_by(|a, b| b.count.cmp(&a.count));
    ranges.truncate(TRACE_RANGE_HIGHLIGHT_LIMIT);
    ranges
}

fn make_range_highlight(details: &[&NeighborDetail]) -> TraceRangeHighlight {
    let mut literals = 0usize;
    let mut formulas = 0usize;
    let mut blanks = 0usize;
    let mut sample_values = Vec::new();
    let mut sample_formulas = Vec::new();
    let mut sample_addresses = Vec::new();

    for detail in details {
        match detail.kind {
            TraceCellKind::Formula => {
                formulas += 1;
                if let Some(formula) = &detail.formula
                    && sample_formulas.len() < TRACE_RANGE_FORMULA_SAMPLES
                    && !sample_formulas.contains(formula)
                {
                    sample_formulas.push(formula.clone());
                }
            }
            TraceCellKind::Literal => {
                literals += 1;
                if let Some(value) = &detail.value
                    && sample_values.len() < TRACE_RANGE_VALUE_SAMPLES
                {
                    sample_values.push(value.clone());
                }
            }
            TraceCellKind::Blank => blanks += 1,
            TraceCellKind::External => {}
        }
        if sample_addresses.len() < TRACE_RANGE_VALUE_SAMPLES {
            sample_addresses.push(detail.address.clone());
        }
    }

    TraceRangeHighlight {
        start: details
            .first()
            .map(|d| d.address.clone())
            .unwrap_or_default(),
        end: details
            .last()
            .map(|d| d.address.clone())
            .unwrap_or_default(),
        count: details.len(),
        literals,
        formulas,
        blanks,
        sample_values,
        sample_formulas,
        sample_addresses,
    }
}

fn build_notable_cells(
    details: &[NeighborDetail],
    ranges: &[TraceRangeHighlight],
    groups: &[TraceFormulaGroupHighlight],
) -> Vec<TraceCellHighlight> {
    let mut exclude: HashSet<String> = HashSet::new();
    for range in ranges {
        exclude.insert(range.start.clone());
        exclude.insert(range.end.clone());
        for addr in &range.sample_addresses {
            exclude.insert(addr.clone());
        }
    }
    for group in groups {
        for addr in &group.sample_addresses {
            exclude.insert(addr.clone());
        }
    }

    let mut highlights = Vec::new();
    let mut kind_counts: HashMap<TraceCellKind, usize> = HashMap::new();

    for detail in details {
        if highlights.len() >= TRACE_CELL_HIGHLIGHT_LIMIT {
            break;
        }
        if exclude.contains(&detail.address) {
            continue;
        }
        let counter = kind_counts.entry(detail.kind.clone()).or_insert(0);
        if *counter >= 2 && detail.kind != TraceCellKind::External {
            continue;
        }
        highlights.push(TraceCellHighlight {
            address: detail.address.clone(),
            kind: detail.kind.clone(),
            value: detail.value.clone(),
            formula: detail.formula.clone(),
        });
        *counter += 1;
    }

    highlights
}

fn build_edges_for_layer(
    raw_edges: &[TraceEdgeRaw],
    selected: &HashSet<String>,
    formula_lookup: &HashMap<String, TraceFormulaInfo>,
) -> Vec<FormulaTraceEdge> {
    let mut edges = Vec::new();
    for edge in raw_edges {
        if selected.contains(&edge.neighbor) {
            let formula = lookup_formula_info(formula_lookup, &edge.neighbor, &edge.neighbor)
                .map(|info| info.formula.clone());
            edges.push(FormulaTraceEdge {
                from: edge.from.clone(),
                to: edge.to.clone(),
                formula,
                note: None,
            });
        }
    }
    edges.sort_by(|a, b| compare_addresses(&a.to, &b.to));
    edges
}

fn lookup_formula_info<'a>(
    lookup: &'a HashMap<String, TraceFormulaInfo>,
    cell_ref: &str,
    original: &str,
) -> Option<&'a TraceFormulaInfo> {
    if let Some(info) = lookup.get(cell_ref) {
        return Some(info);
    }
    if let (Some(_sheet), Some(cell)) = split_sheet_and_cell(original) {
        let upper = cell.to_ascii_uppercase();
        return lookup.get(&upper);
    }
    None
}

fn compare_addresses(left: &str, right: &str) -> Ordering {
    let (sheet_left, cell_left) = split_sheet_and_cell(left);
    let (sheet_right, cell_right) = split_sheet_and_cell(right);

    let sheet_left_key = sheet_left
        .as_ref()
        .map(|s| clean_sheet_name(s).to_ascii_uppercase())
        .unwrap_or_default();
    let sheet_right_key = sheet_right
        .as_ref()
        .map(|s| clean_sheet_name(s).to_ascii_uppercase())
        .unwrap_or_default();

    match sheet_left_key.cmp(&sheet_right_key) {
        Ordering::Equal => {
            let left_core = cell_left.unwrap_or_else(|| left.to_string());
            let right_core = cell_right.unwrap_or_else(|| right.to_string());
            let left_coords = parse_address(&left_core.to_ascii_uppercase());
            let right_coords = parse_address(&right_core.to_ascii_uppercase());
            match (left_coords, right_coords) {
                (Some((lc, lr)), Some((rc, rr))) => lc
                    .cmp(&rc)
                    .then_with(|| lr.cmp(&rr))
                    .then_with(|| left_core.cmp(&right_core)),
                _ => left_core.cmp(&right_core),
            }
        }
        other => other,
    }
}

fn split_sheet_and_cell(address: &str) -> (Option<String>, Option<String>) {
    if let Some(idx) = address.rfind('!') {
        let sheet = address[..idx].to_string();
        let cell = address[idx + 1..].to_string();
        (Some(sheet), Some(cell))
    } else {
        (None, Some(address.to_string()))
    }
}

fn clean_sheet_name(sheet: &str) -> String {
    let trimmed = sheet.trim_matches(|c| c == '\'' || c == '"');
    let after_bracket = trimmed.rsplit(']').next().unwrap_or(trimmed);
    after_bracket
        .trim_matches(|c| c == '\'' || c == '"')
        .to_string()
}
