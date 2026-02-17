use super::RecalcResult;
use crate::recalc::RecalcBackend;
use crate::utils::column_number_to_name;
use anyhow::{Context, Result, anyhow};
use async_trait::async_trait;
use formualizer::workbook::{
    LiteralValue, LoadStrategy, SpreadsheetReader, UmyaAdapter, Workbook, WorkbookMode,
};
use std::io::Cursor;
use std::path::Path;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::thread;
use std::time::{Duration, Instant};

pub struct FormualizerBackend;

#[async_trait]
impl RecalcBackend for FormualizerBackend {
    async fn recalculate(
        &self,
        fork_work_path: &Path,
        timeout_ms: Option<u64>,
    ) -> Result<RecalcResult> {
        let path = fork_work_path.to_path_buf();
        tokio::task::spawn_blocking(move || recalc_sync(&path, timeout_ms)).await?
    }

    fn is_available(&self) -> bool {
        true
    }

    fn name(&self) -> &'static str {
        "formualizer"
    }
}

fn recalc_sync(path: &Path, timeout_ms: Option<u64>) -> Result<RecalcResult> {
    let start = Instant::now();

    // Read the workbook bytes once so we can hydrate umya from memory.
    // If UmyaAdapter::open_bytes is supported by the active formualizer backend,
    // this path also avoids a second filesystem read for formualizer ingestion.
    // Otherwise we gracefully fall back to open_path (legacy behavior).
    let workbook_bytes =
        std::fs::read(path).with_context(|| format!("failed to read workbook bytes {:?}", path))?;

    let mut reader = Cursor::new(workbook_bytes.clone());
    let mut spreadsheet = umya_spreadsheet::reader::xlsx::read_reader(&mut reader, true)
        .with_context(|| format!("failed to parse workbook {:?}", path))?;

    let adapter = UmyaAdapter::open_bytes(workbook_bytes)
        .or_else(|_| UmyaAdapter::open_path(path))
        .map_err(|e| anyhow!("failed to open workbook adapter {:?}: {e}", path))?;
    let mut workbook =
        Workbook::from_reader_with_mode(adapter, LoadStrategy::EagerAll, WorkbookMode::Ephemeral)
            .map_err(|e| anyhow!("failed to construct formualizer workbook: {e}"))?;

    let (cells_evaluated, cycle_errors) = evaluate_with_optional_timeout(&mut workbook, timeout_ms)
        .map_err(|e| anyhow!("formualizer evaluate_all failed: {e}"))?;

    let mut eval_errors = Vec::new();
    if cycle_errors > 0 {
        eval_errors.push(format!(
            "Detected {} circular reference cycle(s)",
            cycle_errors
        ));
    }

    let sheet_names: Vec<String> = workbook.sheet_names();
    for sheet_name in sheet_names {
        let formula_cells = collect_formula_cells(&spreadsheet, &sheet_name);
        if formula_cells.is_empty() {
            continue;
        }

        let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) else {
            continue;
        };

        for (row, col) in formula_cells {
            let value = workbook
                .get_value(&sheet_name, row, col)
                .unwrap_or(LiteralValue::Empty);
            let cell = sheet.get_cell_mut((col, row));

            if let Some(cache) = literal_to_formula_cache(&value) {
                cell.get_cell_value_mut().set_formula_result_default(cache);
            } else {
                cell.get_cell_value_mut().set_formula_result_default("");
            }

            if let LiteralValue::Error(err) = value
                && eval_errors.len() < 200
            {
                let addr = format!("{}{}", column_number_to_name(col), row);
                eval_errors.push(format!("{}!{}: {}", sheet_name, addr, err));
            }
        }
    }

    umya_spreadsheet::writer::xlsx::write(&spreadsheet, path)
        .with_context(|| format!("failed to write workbook {:?}", path))?;

    Ok(RecalcResult {
        duration_ms: start.elapsed().as_millis() as u64,
        was_warm: true,
        backend_name: "formualizer",
        cells_evaluated: Some(cells_evaluated),
        eval_errors: if eval_errors.is_empty() {
            None
        } else {
            Some(eval_errors)
        },
    })
}

fn evaluate_with_optional_timeout(
    workbook: &mut Workbook,
    timeout_ms: Option<u64>,
) -> Result<(u64, u64), formualizer::workbook::IoError> {
    let Some(timeout_ms) = timeout_ms else {
        let eval = workbook.evaluate_all()?;
        return Ok((eval.computed_vertices as u64, eval.cycle_errors as u64));
    };

    let cancel_flag = Arc::new(AtomicBool::new(false));
    let done_flag = Arc::new(AtomicBool::new(false));
    let cancel_for_thread = cancel_flag.clone();
    let done_for_thread = done_flag.clone();

    let handle = thread::spawn(move || {
        let deadline = Instant::now() + Duration::from_millis(timeout_ms);
        // Relaxed is sufficient: flag is monotonic false->true, no data synchronized.
        while !done_for_thread.load(Ordering::Relaxed) {
            if Instant::now() >= deadline {
                cancel_for_thread.store(true, Ordering::Relaxed);
                break;
            }
            thread::sleep(Duration::from_millis(5));
        }
    });

    let result = workbook.evaluate_all_cancellable(cancel_flag);
    done_flag.store(true, Ordering::Relaxed);
    let _ = handle.join();
    let eval = result?;
    Ok((eval.computed_vertices as u64, eval.cycle_errors as u64))
}

fn collect_formula_cells(
    spreadsheet: &umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
) -> Vec<(u32, u32)> {
    let Some(sheet) = spreadsheet.get_sheet_by_name(sheet_name) else {
        return Vec::new();
    };

    sheet
        .get_cell_collection()
        .into_iter()
        .filter_map(|cell| {
            let cv = cell.get_cell_value();
            if !cv.is_formula() {
                return None;
            }
            let coord = cell.get_coordinate();
            Some((*coord.get_row_num(), *coord.get_col_num()))
        })
        .collect()
}

fn literal_to_formula_cache(value: &LiteralValue) -> Option<String> {
    match value {
        LiteralValue::Int(i) => Some(i.to_string()),
        LiteralValue::Number(n) => Some(n.to_string()),
        LiteralValue::Text(s) => Some(s.clone()),
        LiteralValue::Boolean(b) => Some(if *b { "TRUE" } else { "FALSE" }.to_string()),
        LiteralValue::Error(e) => Some(e.to_string()),
        LiteralValue::Date(_)
        | LiteralValue::DateTime(_)
        | LiteralValue::Time(_)
        | LiteralValue::Duration(_) => value.as_serial_number().map(|n| n.to_string()),
        // For array results, cache the top-left value for this cell.
        // Multi-cell spills are represented as independent formula cells and are
        // populated per-cell during the formula-cell writeback loop.
        LiteralValue::Array(values) => values
            .first()
            .and_then(|row| row.first())
            .and_then(literal_to_formula_cache),
        LiteralValue::Empty | LiteralValue::Pending => None,
    }
}
