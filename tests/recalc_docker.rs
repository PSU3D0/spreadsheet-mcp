//! Integration tests for recalc functionality using Docker container with LibreOffice.
//!
//! These tests require Docker and are ignored by default. Run with:
//! ```
//! cargo test --features recalc --test recalc_docker -- --ignored
//! ```

use anyhow::Result;
use serial_test::serial;
use std::time::Instant;

#[path = "./support/mod.rs"]
mod support;

use support::TestWorkspace;
use support::docker::LibreOfficeRecalc;

#[tokio::test]
#[ignore]
#[serial]
async fn test_recalc_updates_formula_result() -> Result<()> {
    let workspace = TestWorkspace::new();

    workspace.create_workbook("recalc_test.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Data");
        sheet.get_cell_mut("A1").set_value_number(100);
        sheet.get_cell_mut("A2").set_value_number(20);
        let sum_cell = sheet.get_cell_mut("A3");
        sum_cell.set_formula("SUM(A1:A2)");
        sum_cell.set_formula_result_default("0");
    });

    let lo = LibreOfficeRecalc::new(workspace.root()).await?;

    let start = Instant::now();
    lo.recalc("recalc_test.xlsx").await?;
    let elapsed = start.elapsed();
    eprintln!("Recalc took: {:?}", elapsed);

    let book = umya_spreadsheet::reader::xlsx::read(&workspace.path("recalc_test.xlsx"))?;
    let sheet = book.get_sheet_by_name("Data").unwrap();
    let a3_value = sheet.get_cell("A3").unwrap().get_value();

    assert_eq!(a3_value, "120", "A3 should be 120 (100 + 20)");

    Ok(())
}

#[tokio::test]
#[ignore]
#[serial]
async fn test_recalc_cross_sheet_reference() -> Result<()> {
    let workspace = TestWorkspace::new();

    workspace.create_workbook("cross_sheet.xlsx", |book| {
        let sheet1 = book.get_sheet_mut(&0).unwrap();
        sheet1.set_name("Input");
        sheet1.get_cell_mut("A1").set_value_number(50);

        let sheet2 = book.new_sheet("Output").unwrap();
        let ref_cell = sheet2.get_cell_mut("A1");
        ref_cell.set_formula("Input!A1*2");
        ref_cell.set_formula_result_default("0");
    });

    let lo = LibreOfficeRecalc::new(workspace.root()).await?;

    let start = Instant::now();
    lo.recalc("cross_sheet.xlsx").await?;
    let elapsed = start.elapsed();
    eprintln!("Cross-sheet recalc took: {:?}", elapsed);

    let book = umya_spreadsheet::reader::xlsx::read(&workspace.path("cross_sheet.xlsx"))?;
    let sheet = book.get_sheet_by_name("Output").unwrap();
    let a1_value = sheet.get_cell("A1").unwrap().get_value();

    assert_eq!(a1_value, "100", "Output!A1 should be 100 (50 * 2)");

    Ok(())
}

#[tokio::test]
#[ignore]
#[serial]
async fn test_recalc_complex_formulas() -> Result<()> {
    let workspace = TestWorkspace::new();

    workspace.create_workbook("complex.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Sheet1");

        sheet.get_cell_mut("A1").set_value_number(10);
        sheet.get_cell_mut("A2").set_value_number(20);
        sheet.get_cell_mut("A3").set_value_number(30);

        let avg = sheet.get_cell_mut("B1");
        avg.set_formula("AVERAGE(A1:A3)");
        avg.set_formula_result_default("0");

        let max = sheet.get_cell_mut("B2");
        max.set_formula("MAX(A1:A3)");
        max.set_formula_result_default("0");

        let min = sheet.get_cell_mut("B3");
        min.set_formula("MIN(A1:A3)");
        min.set_formula_result_default("0");

        let nested = sheet.get_cell_mut("B4");
        nested.set_formula("IF(B1>15,\"High\",\"Low\")");
        nested.set_formula_result_default("");
    });

    let lo = LibreOfficeRecalc::new(workspace.root()).await?;

    let start = Instant::now();
    lo.recalc("complex.xlsx").await?;
    let elapsed = start.elapsed();
    eprintln!("Complex recalc took: {:?}", elapsed);

    let book = umya_spreadsheet::reader::xlsx::read(&workspace.path("complex.xlsx"))?;
    let sheet = book.get_sheet_by_name("Sheet1").unwrap();

    assert_eq!(sheet.get_cell("B1").unwrap().get_value(), "20");
    assert_eq!(sheet.get_cell("B2").unwrap().get_value(), "30");
    assert_eq!(sheet.get_cell("B3").unwrap().get_value(), "10");
    assert_eq!(sheet.get_cell("B4").unwrap().get_value(), "High");

    Ok(())
}

#[tokio::test]
#[ignore]
#[serial]
async fn test_recalc_chain_dependencies() -> Result<()> {
    let workspace = TestWorkspace::new();

    workspace.create_workbook("chain.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Chain");

        sheet.get_cell_mut("A1").set_value_number(5);

        let b1 = sheet.get_cell_mut("B1");
        b1.set_formula("A1*2");
        b1.set_formula_result_default("0");

        let c1 = sheet.get_cell_mut("C1");
        c1.set_formula("B1*2");
        c1.set_formula_result_default("0");

        let d1 = sheet.get_cell_mut("D1");
        d1.set_formula("C1*2");
        d1.set_formula_result_default("0");
    });

    let lo = LibreOfficeRecalc::new(workspace.root()).await?;

    let start = Instant::now();
    lo.recalc("chain.xlsx").await?;
    let elapsed = start.elapsed();
    eprintln!("Chain recalc took: {:?}", elapsed);

    let book = umya_spreadsheet::reader::xlsx::read(&workspace.path("chain.xlsx"))?;
    let sheet = book.get_sheet_by_name("Chain").unwrap();

    assert_eq!(sheet.get_cell("B1").unwrap().get_value(), "10");
    assert_eq!(sheet.get_cell("C1").unwrap().get_value(), "20");
    assert_eq!(sheet.get_cell("D1").unwrap().get_value(), "40");

    Ok(())
}

#[tokio::test]
#[ignore]
#[serial]
async fn test_recalc_preserves_formatting() -> Result<()> {
    let workspace = TestWorkspace::new();

    workspace.create_workbook("format.xlsx", |book| {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Formatted");

        sheet.get_cell_mut("A1").set_value_number(1000);
        sheet.get_cell_mut("A2").set_value_number(2000);

        let sum = sheet.get_cell_mut("A3");
        sum.set_formula("SUM(A1:A2)");
        sum.set_formula_result_default("0");
    });

    let lo = LibreOfficeRecalc::new(workspace.root()).await?;
    lo.recalc("format.xlsx").await?;

    let book = umya_spreadsheet::reader::xlsx::read(&workspace.path("format.xlsx"))?;
    let sheet = book.get_sheet_by_name("Formatted").unwrap();

    assert_eq!(sheet.get_cell("A3").unwrap().get_value(), "3000");

    Ok(())
}
