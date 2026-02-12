use anyhow::{Result, bail};
use serde_json::Value;
use std::path::PathBuf;

use crate::model::TableOutputFormat;
use crate::runtime::stateless::StatelessRuntime;
use crate::tools;
use crate::tools::{
    DescribeWorkbookParams, ListSheetsParams, RangeValuesParams, SheetOverviewParams,
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
