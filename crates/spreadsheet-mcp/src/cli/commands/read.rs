use anyhow::{Result, bail};
use serde_json::Value;
use std::path::PathBuf;

use crate::cli::{FindValueMode, FormulaSort, TableReadFormat, TraceDirectionArg};
use crate::model::{FindMode, TableOutputFormat, TraceDirection};
use crate::runtime::stateless::StatelessRuntime;
use crate::tools;
use crate::tools::{
    DescribeWorkbookParams, FindValueParams, FormulaSortBy, FormulaTraceParams, ListSheetsParams,
    RangeValuesParams, ReadTableParams, SheetFormulaMapParams, SheetOverviewParams,
    TableProfileParams,
};

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
    format: Option<TableReadFormat>,
) -> Result<Value> {
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
            table_name: None,
            region_id: None,
            range,
            header_row: None,
            header_rows: None,
            columns: None,
            filters: None,
            sample_mode: None,
            limit: None,
            offset: None,
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
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let (state, workbook_id) = runtime.open_state_for_file(&file).await?;
    let sheet_name = match sheet {
        Some(name) => Some(resolve_sheet_name(&state, &workbook_id, &name).await?),
        None => None,
    };
    let response = tools::find_value(
        state,
        FindValueParams {
            workbook_or_fork_id: workbook_id,
            query,
            mode: mode.map(map_find_value_mode),
            sheet_name,
            ..FindValueParams::default()
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
) -> Result<Value> {
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
            depth: None,
            limit: None,
            page_size: None,
            cursor: None,
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

fn map_find_value_mode(mode: FindValueMode) -> FindMode {
    match mode {
        FindValueMode::Value => FindMode::Value,
        FindValueMode::Label => FindMode::Label,
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
