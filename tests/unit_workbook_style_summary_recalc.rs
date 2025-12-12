#![cfg(feature = "recalc")]

use anyhow::Result;
use spreadsheet_mcp::model::{FontPatch, StylePatch};
use spreadsheet_mcp::state::AppState;
use spreadsheet_mcp::tools::fork::{
    CreateForkParams, StyleBatchParams, StyleOp, StyleTarget, create_fork, style_batch,
};
use spreadsheet_mcp::tools::{
    ListWorkbooksParams, WorkbookStyleSummaryParams, list_workbooks, workbook_style_summary,
};

mod support;

fn recalc_state(workspace: &support::TestWorkspace) -> std::sync::Arc<AppState> {
    let config = workspace.config_with(|cfg| {
        cfg.recalc_enabled = true;
    });
    support::app_state_with_config(config)
}

#[tokio::test(flavor = "current_thread")]
async fn workbook_style_summary_reflects_styles_in_forks() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("fork_styles.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value("x");
        sheet.get_cell_mut("A2").set_value("y");
    });

    let state = recalc_state(&workspace);
    let list = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = list.workbooks[0].workbook_id.clone();

    let base_summary = workbook_style_summary(
        state.clone(),
        WorkbookStyleSummaryParams {
            workbook_id: workbook_id.clone(),
            max_styles: None,
            max_conditional_formats: None,
            max_cells_scan: None,
        },
    )
    .await?;
    let base_ids: std::collections::HashSet<String> = base_summary
        .styles
        .iter()
        .map(|s| s.style_id.clone())
        .collect();

    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    let patch = StylePatch {
        font: Some(Some(FontPatch {
            italic: Some(Some(true)),
            ..Default::default()
        })),
        ..Default::default()
    };

    style_batch(
        state.clone(),
        StyleBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StyleOp {
                sheet_name: "Sheet1".to_string(),
                target: StyleTarget::Cells {
                    cells: vec!["A2".to_string()],
                },
                patch,
                op_mode: Some("merge".to_string()),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_summary = workbook_style_summary(
        state,
        WorkbookStyleSummaryParams {
            workbook_id: spreadsheet_mcp::model::WorkbookId(fork.fork_id),
            max_styles: None,
            max_conditional_formats: None,
            max_cells_scan: None,
        },
    )
    .await?;

    let fork_ids: std::collections::HashSet<String> = fork_summary
        .styles
        .iter()
        .map(|s| s.style_id.clone())
        .collect();

    assert!(
        fork_ids.difference(&base_ids).next().is_some(),
        "expected fork to introduce a new style"
    );

    Ok(())
}
