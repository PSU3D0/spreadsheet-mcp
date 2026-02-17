use anyhow::{Result, anyhow, bail};
use serde_json::Value;
use std::path::PathBuf;

use crate::cli::{
    FindValueMode, FormulaSort, LabelDirectionArg, SheetPageFormatArg, TableReadFormat,
    TableSampleModeArg, TraceDirectionArg,
};
use crate::model::{
    FindMode, FormulaParsePolicy, LabelDirection, SheetPageFormat, TableOutputFormat, TraceCursor,
    TraceDirection,
};
use crate::runtime::stateless::StatelessRuntime;
use crate::tools;
use crate::tools::{
    DescribeWorkbookParams, FindFormulaParams, FindValueParams, FormulaSortBy, FormulaTraceParams,
    ListSheetsParams, NamedRangesParams, RangeValuesParams, ReadTableParams, SampleMode,
    ScanVolatilesParams, SheetFormulaMapParams, SheetOverviewParams, SheetPageParams,
    SheetStatisticsParams, TableFilter, TableProfileParams,
};

const TRACE_DEPTH_MIN: u32 = 1;
const TRACE_DEPTH_MAX: u32 = 5;
const TRACE_PAGE_SIZE_MIN: usize = 5;
const TRACE_PAGE_SIZE_MAX: usize = 200;

const SHEET_PAGE_DEFAULT_START_ROW: u32 = 1;
const SHEET_PAGE_DEFAULT_PAGE_SIZE: u32 = 50;
const SHEET_PAGE_DEFAULT_INCLUDE_FORMULAS: bool = true;
const SHEET_PAGE_DEFAULT_INCLUDE_STYLES: bool = false;
const SHEET_PAGE_DEFAULT_INCLUDE_HEADER: bool = true;

pub async fn list_sheets(file: PathBuf) -> Result<Value> {
    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let response = tools::list_sheets(
        state,
        ListSheetsParams {
            workbook_or_fork_id: workbook_id,
            limit: None,
            offset: None,
            include_bounds: None,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn sheet_overview(file: PathBuf, sheet: String) -> Result<Value> {
    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet = resolve_sheet_name(&state, &workbook_id, &sheet).await?;
    let response = tools::sheet_overview(
        state,
        SheetOverviewParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: sheet,
            max_regions: None,
            max_headers: None,
            include_headers: None,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn range_values(file: PathBuf, sheet: String, ranges: Vec<String>) -> Result<Value> {
    if ranges.is_empty() {
        bail!("at least one range must be provided");
    }
    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet = resolve_sheet_name(&state, &workbook_id, &sheet).await?;
    let response = tools::range_values(
        state,
        RangeValuesParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: sheet,
            ranges,
            include_headers: None,
            format: Some(TableOutputFormat::Json),
            page_size: None,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn sheet_page(
    file: PathBuf,
    sheet: String,
    start_row: Option<u32>,
    page_size: Option<u32>,
    columns: Option<Vec<String>>,
    columns_by_header: Option<Vec<String>>,
    include_formulas: Option<bool>,
    include_styles: Option<bool>,
    include_header: Option<bool>,
    format: SheetPageFormatArg,
) -> Result<Value> {
    validate_sheet_page_arguments(page_size, columns.as_ref())?;

    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet = resolve_sheet_name(&state, &workbook_id, &sheet).await?;
    let response = tools::sheet_page(
        state,
        SheetPageParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: sheet,
            start_row: start_row.unwrap_or(SHEET_PAGE_DEFAULT_START_ROW),
            page_size: page_size.unwrap_or(SHEET_PAGE_DEFAULT_PAGE_SIZE),
            columns,
            columns_by_header,
            include_formulas: include_formulas.unwrap_or(SHEET_PAGE_DEFAULT_INCLUDE_FORMULAS),
            include_styles: include_styles.unwrap_or(SHEET_PAGE_DEFAULT_INCLUDE_STYLES),
            include_header: include_header.unwrap_or(SHEET_PAGE_DEFAULT_INCLUDE_HEADER),
            format: Some(map_sheet_page_format(format)),
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn describe(file: PathBuf) -> Result<Value> {
    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let response = tools::describe_workbook(
        state,
        DescribeWorkbookParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn read_table(
    file: PathBuf,
    sheet: Option<String>,
    range: Option<String>,
    table_name: Option<String>,
    region_id: Option<u32>,
    limit: Option<u32>,
    offset: Option<u32>,
    sample_mode: Option<TableSampleModeArg>,
    filters_json: Option<String>,
    filters_file: Option<PathBuf>,
    format: Option<TableReadFormat>,
) -> Result<Value> {
    validate_read_table_arguments(limit, offset, sample_mode)?;
    let filters = parse_table_filters(filters_json, filters_file)?;

    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet_name = match sheet {
        Some(name) => Some(resolve_sheet_name(&state, &workbook_id, &name).await?),
        None => None,
    };
    let response = tools::read_table(
        state,
        ReadTableParams {
            workbook_or_fork_id: workbook_id,
            sheet_name,
            table_name,
            region_id,
            range,
            header_row: None,
            header_rows: None,
            columns: None,
            filters,
            sample_mode: sample_mode.map(map_table_sample_mode),
            limit,
            offset,
            format: format.map(map_table_read_format),
            include_headers: None,
            include_types: None,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn find_value(
    file: PathBuf,
    query: String,
    sheet: Option<String>,
    mode: Option<FindValueMode>,
    label_direction: Option<LabelDirectionArg>,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet_name = match sheet {
        Some(name) => Some(resolve_sheet_name(&state, &workbook_id, &name).await?),
        None => None,
    };

    let mapped_mode = mode.map(map_find_value_mode);
    let label = if matches!(mapped_mode, Some(FindMode::Label)) {
        Some(query.clone())
    } else {
        None
    };

    let response = tools::find_value(
        state,
        FindValueParams {
            workbook_or_fork_id: workbook_id,
            query,
            label,
            mode: mapped_mode,
            direction: label_direction.map(map_label_direction),
            sheet_name,
            ..FindValueParams::default()
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn named_ranges(
    file: PathBuf,
    sheet: Option<String>,
    name_prefix: Option<String>,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet_name = match sheet {
        Some(name) => Some(resolve_sheet_name(&state, &workbook_id, &name).await?),
        None => None,
    };

    let response = tools::named_ranges(
        state,
        NamedRangesParams {
            workbook_or_fork_id: workbook_id,
            sheet_name,
            name_prefix,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn find_formula(
    file: PathBuf,
    query: String,
    sheet: Option<String>,
    limit: Option<u32>,
    offset: Option<u32>,
) -> Result<Value> {
    validate_positive_limit(limit, "--limit")?;

    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet_name = match sheet {
        Some(name) => Some(resolve_sheet_name(&state, &workbook_id, &name).await?),
        None => None,
    };

    let response = tools::find_formula(
        state,
        FindFormulaParams {
            workbook_or_fork_id: workbook_id,
            query,
            sheet_name,
            case_sensitive: false,
            include_context: false,
            limit: limit.unwrap_or(50),
            offset: offset.unwrap_or(0),
            context_rows: None,
            context_cols: None,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn scan_volatiles(
    file: PathBuf,
    sheet: Option<String>,
    limit: Option<u32>,
    offset: Option<u32>,
    formula_parse_policy: Option<FormulaParsePolicy>,
) -> Result<Value> {
    validate_positive_limit(limit, "--limit")?;

    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet_name = match sheet {
        Some(name) => Some(resolve_sheet_name(&state, &workbook_id, &name).await?),
        None => None,
    };

    let response = tools::scan_volatiles(
        state,
        ScanVolatilesParams {
            workbook_or_fork_id: workbook_id,
            sheet_name,
            summary_only: None,
            include_addresses: None,
            addresses_limit: None,
            limit,
            offset,
            formula_parse_policy,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn sheet_statistics(file: PathBuf, sheet: String) -> Result<Value> {
    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet_name = resolve_sheet_name(&state, &workbook_id, &sheet).await?;

    let response = tools::sheet_statistics(
        state,
        SheetStatisticsParams {
            workbook_or_fork_id: workbook_id,
            sheet_name,
            sample_rows: None,
            summary_only: None,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn formula_map(
    file: PathBuf,
    sheet: String,
    limit: Option<u32>,
    sort_by: Option<FormulaSort>,
    formula_parse_policy: Option<FormulaParsePolicy>,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet = resolve_sheet_name(&state, &workbook_id, &sheet).await?;
    let response = tools::sheet_formula_map(
        state,
        SheetFormulaMapParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: sheet,
            range: None,
            expand: false,
            limit,
            sort_by: sort_by.map(map_formula_sort),
            summary_only: None,
            include_addresses: None,
            addresses_limit: None,
            formula_parse_policy,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn formula_trace(
    file: PathBuf,
    sheet: String,
    cell: String,
    direction: TraceDirectionArg,
    depth: Option<u32>,
    page_size: Option<usize>,
    cursor_depth: Option<u32>,
    cursor_offset: Option<usize>,
    formula_parse_policy: Option<FormulaParsePolicy>,
) -> Result<Value> {
    validate_formula_trace_arguments(depth, page_size)?;
    let cursor = build_trace_cursor(cursor_depth, cursor_offset)?;

    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet = resolve_sheet_name(&state, &workbook_id, &sheet).await?;
    let response = tools::formula_trace(
        state,
        FormulaTraceParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: sheet,
            cell_address: cell,
            direction: map_trace_direction(direction),
            depth,
            limit: None,
            page_size,
            cursor,
            formula_parse_policy,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

pub async fn table_profile(file: PathBuf, sheet: Option<String>) -> Result<Value> {
    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet_name = match sheet {
        Some(name) => Some(resolve_sheet_name(&state, &workbook_id, &name).await?),
        None => None,
    };
    let response = tools::table_profile(
        state,
        TableProfileParams {
            workbook_or_fork_id: workbook_id,
            sheet_name,
            region_id: None,
            table_name: None,
            sample_mode: None,
            sample_size: None,
            summary_only: None,
        },
    )
    .await?;
    Ok(serde_json::to_value(response)?)
}

fn map_table_read_format(format: TableReadFormat) -> TableOutputFormat {
    match format {
        TableReadFormat::Json => TableOutputFormat::Json,
        TableReadFormat::Values => TableOutputFormat::Values,
        TableReadFormat::Csv => TableOutputFormat::Csv,
    }
}

fn map_sheet_page_format(format: SheetPageFormatArg) -> SheetPageFormat {
    match format {
        SheetPageFormatArg::Full => SheetPageFormat::Full,
        SheetPageFormatArg::Compact => SheetPageFormat::Compact,
        SheetPageFormatArg::ValuesOnly => SheetPageFormat::ValuesOnly,
    }
}

fn map_table_sample_mode(mode: TableSampleModeArg) -> SampleMode {
    match mode {
        TableSampleModeArg::First => SampleMode::First,
        TableSampleModeArg::Last => SampleMode::Last,
        TableSampleModeArg::Distributed => SampleMode::Distributed,
    }
}

fn map_find_value_mode(mode: FindValueMode) -> FindMode {
    match mode {
        FindValueMode::Value => FindMode::Value,
        FindValueMode::Label => FindMode::Label,
    }
}

fn map_label_direction(direction: LabelDirectionArg) -> LabelDirection {
    match direction {
        LabelDirectionArg::Right => LabelDirection::Right,
        LabelDirectionArg::Below => LabelDirection::Below,
        LabelDirectionArg::Any => LabelDirection::Any,
    }
}

fn map_formula_sort(sort: FormulaSort) -> FormulaSortBy {
    match sort {
        FormulaSort::Complexity => FormulaSortBy::Complexity,
        FormulaSort::Count => FormulaSortBy::Count,
    }
}

fn map_trace_direction(direction: TraceDirectionArg) -> TraceDirection {
    match direction {
        TraceDirectionArg::Precedents => TraceDirection::Precedents,
        TraceDirectionArg::Dependents => TraceDirection::Dependents,
    }
}

fn validate_sheet_page_arguments(
    page_size: Option<u32>,
    columns: Option<&Vec<String>>,
) -> Result<()> {
    if matches!(page_size, Some(0)) {
        return Err(invalid_argument("--page-size must be at least 1"));
    }

    validate_sheet_page_columns(columns)?;
    Ok(())
}

fn validate_sheet_page_columns(columns: Option<&Vec<String>>) -> Result<()> {
    let Some(columns) = columns else {
        return Ok(());
    };

    for raw_spec in columns {
        let spec = raw_spec.trim();
        if spec.is_empty() {
            return Err(invalid_argument("invalid column spec: ''"));
        }

        let (start, end) = spec.split_once(':').unwrap_or((spec, spec));
        if !is_valid_column_token(start) || !is_valid_column_token(end) {
            return Err(invalid_argument(format!(
                "invalid column spec: '{raw_spec}'"
            )));
        }

        let start_idx = umya_spreadsheet::helper::coordinate::column_index_from_string(start);
        let end_idx = umya_spreadsheet::helper::coordinate::column_index_from_string(end);
        if start_idx == 0 || end_idx == 0 {
            return Err(invalid_argument(format!(
                "invalid column spec: '{raw_spec}'"
            )));
        }
    }

    Ok(())
}

fn is_valid_column_token(token: &str) -> bool {
    let token = token.trim();
    !token.is_empty() && token.chars().all(|ch| ch.is_ascii_alphabetic())
}

fn validate_positive_limit(limit: Option<u32>, flag_name: &'static str) -> Result<()> {
    if matches!(limit, Some(0)) {
        return Err(invalid_argument(format!("{flag_name} must be at least 1")));
    }
    Ok(())
}

fn validate_read_table_arguments(
    limit: Option<u32>,
    offset: Option<u32>,
    sample_mode: Option<TableSampleModeArg>,
) -> Result<()> {
    validate_positive_limit(limit, "--limit")?;

    if offset.unwrap_or(0) > 0 {
        if let Some(TableSampleModeArg::Last | TableSampleModeArg::Distributed) = sample_mode {
            return Err(invalid_argument(
                "--offset greater than 0 requires --sample-mode first",
            ));
        }
    }

    Ok(())
}

fn parse_table_filters(
    filters_json: Option<String>,
    filters_file: Option<PathBuf>,
) -> Result<Option<Vec<TableFilter>>> {
    match (filters_json, filters_file) {
        (Some(_), Some(_)) => Err(invalid_argument(
            "--filters-json and --filters-file are mutually exclusive",
        )),
        (Some(raw), None) => parse_table_filters_payload(&raw, "--filters-json").map(Some),
        (None, Some(path)) => {
            let raw = std::fs::read_to_string(&path).map_err(|err| {
                invalid_argument(format!(
                    "failed to read --filters-file '{}': {}",
                    path.display(),
                    err
                ))
            })?;
            parse_table_filters_payload(&raw, "--filters-file").map(Some)
        }
        (None, None) => Ok(None),
    }
}

fn parse_table_filters_payload(raw: &str, source: &str) -> Result<Vec<TableFilter>> {
    serde_json::from_str(raw).map_err(|err| {
        invalid_argument(format!(
            "{source} must be a valid JSON array of filters: {err}"
        ))
    })
}

fn validate_formula_trace_arguments(depth: Option<u32>, page_size: Option<usize>) -> Result<()> {
    if let Some(depth) = depth
        && !(TRACE_DEPTH_MIN..=TRACE_DEPTH_MAX).contains(&depth)
    {
        return Err(invalid_argument(format!(
            "--depth must be between {TRACE_DEPTH_MIN} and {TRACE_DEPTH_MAX}"
        )));
    }

    if let Some(page_size) = page_size
        && !(TRACE_PAGE_SIZE_MIN..=TRACE_PAGE_SIZE_MAX).contains(&page_size)
    {
        return Err(invalid_argument(format!(
            "--page-size must be between {TRACE_PAGE_SIZE_MIN} and {TRACE_PAGE_SIZE_MAX}"
        )));
    }

    Ok(())
}

fn build_trace_cursor(
    cursor_depth: Option<u32>,
    cursor_offset: Option<usize>,
) -> Result<Option<TraceCursor>> {
    match (cursor_depth, cursor_offset) {
        (None, None) => Ok(None),
        (Some(depth), Some(offset)) => {
            if depth < 1 {
                return Err(invalid_argument("--cursor-depth must be at least 1"));
            }
            Ok(Some(TraceCursor { depth, offset }))
        }
        (Some(_), None) | (None, Some(_)) => Err(invalid_argument(
            "--cursor-depth and --cursor-offset must be provided together",
        )),
    }
}

fn invalid_argument(message: impl Into<String>) -> anyhow::Error {
    anyhow!("invalid argument: {}", message.into())
}

async fn resolve_sheet_name(
    state: &std::sync::Arc<crate::state::AppState>,
    workbook_id: &crate::model::WorkbookId,
    requested: &str,
) -> Result<String> {
    let response = tools::list_sheets(
        state.clone(),
        ListSheetsParams {
            workbook_or_fork_id: workbook_id.clone(),
            limit: None,
            offset: None,
            include_bounds: None,
        },
    )
    .await?;

    let Some(exact) = response.sheets.iter().find(|entry| entry.name == requested) else {
        if let Some(case_insensitive) = response
            .sheets
            .iter()
            .find(|entry| entry.name.eq_ignore_ascii_case(requested))
        {
            return Ok(case_insensitive.name.clone());
        }

        let best = response
            .sheets
            .iter()
            .min_by_key(|entry| levenshtein(requested, &entry.name))
            .map(|entry| entry.name.clone());
        if let Some(suggestion) = best {
            bail!(
                "sheet '{}' not found; did you mean '{}' ?",
                requested,
                suggestion
            );
        }
        bail!("sheet '{}' not found", requested);
    };

    Ok(exact.name.clone())
}

fn levenshtein(left: &str, right: &str) -> usize {
    if left == right {
        return 0;
    }
    if left.is_empty() {
        return right.chars().count();
    }
    if right.is_empty() {
        return left.chars().count();
    }

    let left_chars = left.chars().collect::<Vec<_>>();
    let right_chars = right.chars().collect::<Vec<_>>();

    let mut prev = (0..=right_chars.len()).collect::<Vec<_>>();
    let mut curr = vec![0usize; right_chars.len() + 1];

    for (i, lc) in left_chars.iter().enumerate() {
        curr[0] = i + 1;
        for (j, rc) in right_chars.iter().enumerate() {
            let cost = usize::from(lc != rc);
            curr[j + 1] = (prev[j + 1] + 1).min(curr[j] + 1).min(prev[j] + cost);
        }
        std::mem::swap(&mut prev, &mut curr);
    }

    prev[right_chars.len()]
}
