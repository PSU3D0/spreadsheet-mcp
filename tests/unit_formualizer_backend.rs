#![cfg(all(feature = "recalc", feature = "recalc-formualizer"))]

use anyhow::Result;
use spreadsheet_mcp::model::WorkbookId;
use spreadsheet_mcp::tools::fork::{
    CreateForkParams, RecalculateParams, create_fork, edit_batch, recalculate,
};
use spreadsheet_mcp::tools::write_normalize::{CellEditInput, EditBatchParamsInput};
use spreadsheet_mcp::tools::{ListWorkbooksParams, list_workbooks};
use spreadsheet_mcp::{RecalcBackendKind, state::AppState};
use std::sync::Arc;

mod support;

async fn first_workbook_id(state: Arc<AppState>) -> Result<WorkbookId> {
    let list = list_workbooks(
        state,
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
            limit: None,
            offset: None,
            include_paths: None,
        },
    )
    .await?;
    Ok(list.workbooks[0].workbook_id.clone())
}

#[tokio::test(flavor = "current_thread")]
async fn recalculate_uses_formualizer_backend_and_updates_formula_cache() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("formualizer_recalc.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value_number(10);
        let out = sheet.get_cell_mut("A2");
        out.set_formula("A1*2");
        out.get_cell_value_mut().set_formula_result_default("0");
    });

    let config = Arc::new(workspace.config_with(|cfg| {
        cfg.recalc_enabled = true;
        cfg.recalc_backend = RecalcBackendKind::Formualizer;
    }));
    let state = Arc::new(AppState::new(config));

    let workbook_id = first_workbook_id(state.clone()).await?;
    let fork = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;

    edit_batch(
        state.clone(),
        EditBatchParamsInput {
            fork_id: fork.fork_id.clone(),
            sheet_name: "Sheet1".to_string(),
            edits: vec![CellEditInput::Shorthand("A1=11".to_string())],
        },
    )
    .await?;

    let recalc = recalculate(
        state.clone(),
        RecalculateParams {
            fork_id: fork.fork_id.clone(),
            timeout_ms: 30_000,
            backend: Some(RecalcBackendKind::Formualizer),
        },
    )
    .await?;

    assert_eq!(recalc.backend, "formualizer");
    assert!(recalc.cells_evaluated.unwrap_or_default() > 0);

    let fork_ctx = state
        .fork_registry()
        .expect("fork registry")
        .get_fork(&fork.fork_id)?;
    let saved = umya_spreadsheet::reader::xlsx::read(&fork_ctx.work_path)?;
    let sheet = saved.get_sheet_by_name("Sheet1").expect("Sheet1 exists");
    assert_eq!(sheet.get_cell("A2").expect("A2 exists").get_value(), "22");

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn recalculate_populates_eval_errors() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("formualizer_eval_errors.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_formula("UNKNOWNFN(1)");
        sheet.get_cell_mut("A2").set_formula("A3+1");
        sheet.get_cell_mut("A3").set_formula("A2+1");
    });

    let config = Arc::new(workspace.config_with(|cfg| {
        cfg.recalc_enabled = true;
        cfg.recalc_backend = RecalcBackendKind::Formualizer;
    }));
    let state = Arc::new(AppState::new(config));
    let workbook_id = first_workbook_id(state.clone()).await?;
    let fork = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;

    let recalc = recalculate(
        state,
        RecalculateParams {
            fork_id: fork.fork_id,
            timeout_ms: 30_000,
            backend: Some(RecalcBackendKind::Formualizer),
        },
    )
    .await?;

    let errors = recalc.eval_errors.unwrap_or_default();
    assert!(!errors.is_empty());
    assert!(errors.iter().any(|e| {
        let lower = e.to_ascii_lowercase();
        lower.contains("circular") || lower.contains("name") || lower.contains("unknown")
    }));
    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn recalculate_timeout_can_cancel_long_eval() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("formualizer_timeout.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value_number(1);
        for row in 2..=30_000u32 {
            sheet
                .get_cell_mut((1, row))
                .set_formula(format!("A{}+1", row - 1));
        }
    });

    let config = Arc::new(workspace.config_with(|cfg| {
        cfg.recalc_enabled = true;
        cfg.recalc_backend = RecalcBackendKind::Formualizer;
    }));
    let state = Arc::new(AppState::new(config));
    let workbook_id = first_workbook_id(state.clone()).await?;
    let fork = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;

    let result = recalculate(
        state,
        RecalculateParams {
            fork_id: fork.fork_id,
            timeout_ms: 1,
            backend: Some(RecalcBackendKind::Formualizer),
        },
    )
    .await;

    assert!(result.is_err(), "expected timeout cancellation error");
    Ok(())
}
