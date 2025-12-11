//! Integration tests for fork workflow: create_fork → edit_batch → get_edits → get_changeset → save_fork
//!
//! These tests exercise the full fork workflow without requiring LibreOffice for recalculation.

#![cfg(feature = "recalc")]

use std::sync::Arc;

use anyhow::Result;
use spreadsheet_mcp::ServerConfig;
use spreadsheet_mcp::diff::Change; // Add Change import
use spreadsheet_mcp::diff::merge::ModificationType;
use spreadsheet_mcp::model::WorkbookId;
use spreadsheet_mcp::state::AppState;
use spreadsheet_mcp::tools::fork::{
    CellEdit, CreateForkParams, DiscardForkParams, EditBatchParams, GetChangesetParams,
    GetEditsParams, ListForksParams, SaveForkParams, create_fork, discard_fork, edit_batch,
    get_changeset, get_edits, list_forks, save_fork,
};
use spreadsheet_mcp::tools::{ListWorkbooksParams, list_workbooks};

#[path = "./support/mod.rs"]
mod support;

fn recalc_enabled_config(workspace: &support::TestWorkspace) -> ServerConfig {
    workspace.config_with(|cfg| {
        cfg.recalc_enabled = true;
    })
}

fn app_state_with_recalc(workspace: &support::TestWorkspace) -> Arc<AppState> {
    let config = Arc::new(recalc_enabled_config(workspace));
    Arc::new(AppState::new(config))
}

async fn discover_workbook(state: Arc<AppState>) -> Result<WorkbookId> {
    let list = list_workbooks(
        state,
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    assert_eq!(list.workbooks.len(), 1, "expected exactly 1 workbook");
    Ok(list.workbooks[0].workbook_id.clone())
}

#[tokio::test]
async fn test_create_fork_basic() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("source.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value_number(100);
        sheet.get_cell_mut("B1").set_formula("A1*2");
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let response = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    assert!(!response.fork_id.is_empty());
    assert!(response.base_workbook.contains("source.xlsx"));
    assert_eq!(response.ttl_seconds, 3600);

    Ok(())
}

#[tokio::test]
async fn test_create_fork_rejects_non_xlsx() -> Result<()> {
    let workspace = support::TestWorkspace::new();

    support::touch_file(&workspace.path("data.xls"));

    let state = app_state_with_recalc(&workspace);

    let list = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: Some("*.xls".to_string()),
        },
    )
    .await?;

    if list.workbooks.is_empty() {
        return Ok(());
    }

    let result = create_fork(
        state,
        CreateForkParams {
            workbook_id: list.workbooks[0].workbook_id.clone(),
        },
    )
    .await;

    assert!(result.is_err());

    Ok(())
}

#[tokio::test]
async fn test_edit_batch_applies_values() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("editable.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Data");
        sheet.get_cell_mut("A1").set_value_number(10);
        sheet.get_cell_mut("A2").set_value_number(20);
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let fork_response = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    let edit_response = edit_batch(
        state.clone(),
        EditBatchParams {
            fork_id: fork_response.fork_id.clone(),
            sheet_name: "Data".to_string(),
            edits: vec![
                CellEdit {
                    address: "A1".to_string(),
                    value: "100".to_string(),
                    is_formula: false,
                },
                CellEdit {
                    address: "A3".to_string(),
                    value: "SUM(A1:A2)".to_string(),
                    is_formula: true,
                },
            ],
        },
    )
    .await?;

    assert_eq!(edit_response.edits_applied, 2);
    assert_eq!(edit_response.total_edits, 2);

    Ok(())
}

#[tokio::test]
async fn test_get_edits_returns_history() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("history.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Sheet1");
        sheet.get_cell_mut("A1").set_value_number(1);
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    // Apply multiple batches
    edit_batch(
        state.clone(),
        EditBatchParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: "Sheet1".to_string(),
            edits: vec![CellEdit {
                address: "A1".to_string(),
                value: "10".to_string(),
                is_formula: false,
            }],
        },
    )
    .await?;

    edit_batch(
        state.clone(),
        EditBatchParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: "Sheet1".to_string(),
            edits: vec![CellEdit {
                address: "B1".to_string(),
                value: "A1*2".to_string(),
                is_formula: true,
            }],
        },
    )
    .await?;

    let edits = get_edits(
        state.clone(),
        GetEditsParams {
            fork_id: fork.fork_id.clone(),
        },
    )
    .await?;

    assert_eq!(edits.edits.len(), 2);

    let a1_edit = &edits.edits[0];
    assert_eq!(a1_edit.address, "A1");
    assert_eq!(a1_edit.value, "10");
    assert!(!a1_edit.is_formula);

    let b1_edit = &edits.edits[1];
    assert_eq!(b1_edit.address, "B1");
    assert_eq!(b1_edit.value, "A1*2");
    assert!(b1_edit.is_formula);

    Ok(())
}

#[tokio::test]
async fn test_get_changeset_detects_modifications() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("changeset.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Sheet1");
        sheet.get_cell_mut("A1").set_value_number(100);
        sheet.get_cell_mut("A2").set_value("original");
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    edit_batch(
        state.clone(),
        EditBatchParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: "Sheet1".to_string(),
            edits: vec![
                CellEdit {
                    address: "A1".to_string(),
                    value: "200".to_string(),
                    is_formula: false,
                },
                CellEdit {
                    address: "A2".to_string(),
                    value: "modified".to_string(),
                    is_formula: false,
                },
            ],
        },
    )
    .await?;

    let changeset = get_changeset(
        state.clone(),
        GetChangesetParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: None,
        },
    )
    .await?;

    assert_eq!(changeset.changes.len(), 2);

    // Find A1 change
    let a1_change = changeset
        .changes
        .iter()
        .find(|c| {
            matches!(c, Change::Cell(cell) if matches!(&cell.diff, spreadsheet_mcp::diff::merge::CellDiff::Modified { address, .. } if address == "A1"))
        })
        .expect("A1 change not found");

    if let Change::Cell(c) = a1_change {
        match &c.diff {
            spreadsheet_mcp::diff::merge::CellDiff::Modified {
                subtype,
                old_value,
                new_value,
                ..
            } => {
                assert!(matches!(subtype, ModificationType::ValueEdit));
                assert_eq!(old_value.as_deref(), Some("100"));
                assert_eq!(new_value.as_deref(), Some("200"));
            }
            _ => panic!("Expected Modified diff"),
        }
    } else {
        panic!("Expected cell change");
    }

    Ok(())
}

#[tokio::test]
async fn test_get_changeset_with_sheet_filter() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("multi_sheet.xlsx", |book| {
        let sheet1 = book.get_sheet_mut(&0).unwrap();
        sheet1.set_name("Sheet1");
        sheet1.get_cell_mut("A1").set_value_number(1);

        let sheet2 = book.new_sheet("Sheet2").unwrap();
        sheet2.get_cell_mut("A1").set_value_number(2);
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    // Edit both sheets
    edit_batch(
        state.clone(),
        EditBatchParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: "Sheet1".to_string(),
            edits: vec![CellEdit {
                address: "A1".to_string(),
                value: "10".to_string(),
                is_formula: false,
            }],
        },
    )
    .await?;

    edit_batch(
        state.clone(),
        EditBatchParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: "Sheet2".to_string(),
            edits: vec![CellEdit {
                address: "A1".to_string(),
                value: "20".to_string(),
                is_formula: false,
            }],
        },
    )
    .await?;

    // Get changeset filtered to Sheet1 only
    let changeset = get_changeset(
        state.clone(),
        GetChangesetParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: Some("Sheet1".to_string()),
        },
    )
    .await?;

    assert_eq!(changeset.changes.len(), 1);

    if let Change::Cell(c) = &changeset.changes[0] {
        assert_eq!(c.sheet, "Sheet1");
    } else {
        panic!("Expected cell change");
    }

    Ok(())
}

#[tokio::test]
async fn test_list_forks() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("listable.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value_number(1);
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let list = list_forks(state.clone(), ListForksParams {}).await?;
    assert!(list.forks.is_empty());

    let fork1 = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_id: workbook_id.clone(),
        },
    )
    .await?;

    let fork2 = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    let list = list_forks(state.clone(), ListForksParams {}).await?;
    assert_eq!(list.forks.len(), 2);

    let fork_ids: Vec<_> = list.forks.iter().map(|f| f.fork_id.as_str()).collect();
    assert!(fork_ids.contains(&fork1.fork_id.as_str()));
    assert!(fork_ids.contains(&fork2.fork_id.as_str()));

    Ok(())
}

#[tokio::test]
async fn test_discard_fork() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("discardable.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value_number(1);
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    // Verify fork exists
    let list = list_forks(state.clone(), ListForksParams {}).await?;
    assert_eq!(list.forks.len(), 1);

    // Discard it
    let discard_response = discard_fork(
        state.clone(),
        DiscardForkParams {
            fork_id: fork.fork_id.clone(),
        },
    )
    .await?;

    assert!(discard_response.discarded);
    assert_eq!(discard_response.fork_id, fork.fork_id);

    // Verify fork is gone
    let list = list_forks(state.clone(), ListForksParams {}).await?;
    assert!(list.forks.is_empty());

    Ok(())
}

#[tokio::test]
async fn test_save_fork_overwrites_original() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("saveable.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Data");
        sheet.get_cell_mut("A1").set_value_number(1);
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    edit_batch(
        state.clone(),
        EditBatchParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: "Data".to_string(),
            edits: vec![CellEdit {
                address: "A1".to_string(),
                value: "999".to_string(),
                is_formula: false,
            }],
        },
    )
    .await?;

    let save_response = save_fork(
        state.clone(),
        SaveForkParams {
            fork_id: fork.fork_id.clone(),
            target_path: None, // Overwrite original
        },
    )
    .await?;

    assert!(save_response.saved_to.contains("saveable.xlsx"));

    // Verify the original file was updated
    let book = umya_spreadsheet::reader::xlsx::read(&path)?;
    let sheet = book.get_sheet_by_name("Data").unwrap();
    let value = sheet.get_cell("A1").unwrap().get_value();
    assert_eq!(value, "999");

    // Fork should be removed after save
    let list = list_forks(state.clone(), ListForksParams {}).await?;
    assert!(list.forks.is_empty());

    Ok(())
}

#[tokio::test]
async fn test_save_fork_to_new_path() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _original = workspace.create_workbook("original.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Data");
        sheet.get_cell_mut("A1").set_value_number(1);
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    edit_batch(
        state.clone(),
        EditBatchParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: "Data".to_string(),
            edits: vec![CellEdit {
                address: "A1".to_string(),
                value: "modified".to_string(),
                is_formula: false,
            }],
        },
    )
    .await?;

    let save_response = save_fork(
        state.clone(),
        SaveForkParams {
            fork_id: fork.fork_id.clone(),
            target_path: Some("copy.xlsx".to_string()),
        },
    )
    .await?;

    assert!(save_response.saved_to.contains("copy.xlsx"));

    // Verify original is unchanged
    let original_book = umya_spreadsheet::reader::xlsx::read(&workspace.path("original.xlsx"))?;
    let original_value = original_book
        .get_sheet_by_name("Data")
        .unwrap()
        .get_cell("A1")
        .unwrap()
        .get_value();
    assert_eq!(original_value, "1");

    // Verify copy has changes
    let copy_book = umya_spreadsheet::reader::xlsx::read(&workspace.path("copy.xlsx"))?;
    let copy_value = copy_book
        .get_sheet_by_name("Data")
        .unwrap()
        .get_cell("A1")
        .unwrap()
        .get_value();
    assert_eq!(copy_value, "modified");

    Ok(())
}

#[tokio::test]
async fn test_full_workflow_without_recalc() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("workflow.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Budget");
        sheet.get_cell_mut("A1").set_value("Item");
        sheet.get_cell_mut("B1").set_value("Amount");
        sheet.get_cell_mut("A2").set_value("Rent");
        sheet.get_cell_mut("B2").set_value_number(1000);
        sheet.get_cell_mut("A3").set_value("Food");
        sheet.get_cell_mut("B3").set_value_number(500);
        sheet.get_cell_mut("A4").set_value("Total");
        let cell = sheet.get_cell_mut("B4");
        cell.set_formula("SUM(B2:B3)");
        cell.set_formula_result_default("1500");
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;
    assert!(!fork.fork_id.is_empty());

    // Step 2: Apply edits (update rent amount)
    let edit_result = edit_batch(
        state.clone(),
        EditBatchParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: "Budget".to_string(),
            edits: vec![CellEdit {
                address: "B2".to_string(),
                value: "1200".to_string(),
                is_formula: false,
            }],
        },
    )
    .await?;
    assert_eq!(edit_result.edits_applied, 1);

    // Step 3: Review edits
    let edits = get_edits(
        state.clone(),
        GetEditsParams {
            fork_id: fork.fork_id.clone(),
        },
    )
    .await?;
    assert_eq!(edits.edits.len(), 1);
    assert_eq!(edits.edits[0].address, "B2");
    assert_eq!(edits.edits[0].value, "1200");

    // Step 4: Get changeset (without recalc - just value change)
    let changeset = get_changeset(
        state.clone(),
        GetChangesetParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: None,
        },
    )
    .await?;
    assert_eq!(changeset.changes.len(), 1);

    // Step 5: Save fork
    let save_result = save_fork(
        state.clone(),
        SaveForkParams {
            fork_id: fork.fork_id.clone(),
            target_path: Some("workflow_updated.xlsx".to_string()),
        },
    )
    .await?;
    assert!(save_result.saved_to.contains("workflow_updated.xlsx"));

    // Verify the saved file
    let saved_book =
        umya_spreadsheet::reader::xlsx::read(&workspace.path("workflow_updated.xlsx"))?;
    let saved_value = saved_book
        .get_sheet_by_name("Budget")
        .unwrap()
        .get_cell("B2")
        .unwrap()
        .get_value();
    assert_eq!(saved_value, "1200");

    Ok(())
}

#[tokio::test]
async fn test_fork_not_found_error() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("dummy.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.get_cell_mut("A1").set_value_number(1);
    });

    let state = app_state_with_recalc(&workspace);

    let result = get_edits(
        state.clone(),
        GetEditsParams {
            fork_id: "nonexistent-fork-id".to_string(),
        },
    )
    .await;

    assert!(result.is_err());
    let err = result.unwrap_err().to_string();
    assert!(
        err.contains("not found"),
        "error should mention not found: {}",
        err
    );

    Ok(())
}

#[tokio::test]
async fn test_edit_nonexistent_sheet_error() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("single_sheet.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("RealSheet");
        sheet.get_cell_mut("A1").set_value_number(1);
    });

    let state = app_state_with_recalc(&workspace);
    let workbook_id = discover_workbook(state.clone()).await?;

    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    let result = edit_batch(
        state.clone(),
        EditBatchParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: "FakeSheet".to_string(),
            edits: vec![CellEdit {
                address: "A1".to_string(),
                value: "test".to_string(),
                is_formula: false,
            }],
        },
    )
    .await;

    assert!(result.is_err());
    let err = result.unwrap_err().to_string();
    assert!(
        err.contains("not found") || err.contains("FakeSheet"),
        "error should mention sheet not found: {}",
        err
    );

    Ok(())
}
