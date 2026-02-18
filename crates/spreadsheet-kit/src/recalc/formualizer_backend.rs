use super::RecalcResult;
use crate::recalc::RecalcBackend;
use crate::utils::column_number_to_name;
use anyhow::{Context, Result, anyhow};
use async_trait::async_trait;
use formualizer::eval::engine::ingest::EngineLoadStream;
use formualizer::eval::engine::{Engine, EvalConfig};
use formualizer::workbook::workbook::WBResolver;
use formualizer::workbook::{LiteralValue, SpreadsheetReader, SpreadsheetWriter, UmyaAdapter};
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

type FormualizerEngine = Engine<WBResolver>;

fn recalc_sync(path: &Path, timeout_ms: Option<u64>) -> Result<RecalcResult> {
    let start = Instant::now();

    // Read workbook bytes once so we can open UmyaAdapter from memory when supported.
    // This avoids a second filesystem read and lets us reuse the same adapter for
    // ingest + cached-value writeback.
    let workbook_bytes =
        std::fs::read(path).with_context(|| format!("failed to read workbook bytes {:?}", path))?;

    let mut adapter = UmyaAdapter::open_bytes(workbook_bytes)
        .or_else(|_| UmyaAdapter::open_path(path))
        .map_err(|e| anyhow!("failed to open workbook adapter {:?}: {e}", path))?;

    let formula_cells = adapter.formula_cells();

    let mut engine = FormualizerEngine::new(WBResolver, EvalConfig::default());
    adapter
        .stream_into_engine(&mut engine)
        .map_err(|e| anyhow!("failed to ingest workbook into formualizer engine: {e}"))?;

    let (cells_evaluated, cycle_errors) =
        evaluate_with_optional_timeout(&mut engine, timeout_ms)
            .map_err(|e| anyhow!("formualizer evaluate_all failed: {e}"))?;

    let mut eval_errors = Vec::new();
    if cycle_errors > 0 {
        eval_errors.push(format!(
            "Detected {} circular reference cycle(s)",
            cycle_errors
        ));
    }

    let date_system = engine.config.date_system;
    for (sheet_name, row, col) in formula_cells {
        let value = engine
            .get_cell_value(&sheet_name, row, col)
            .unwrap_or(LiteralValue::Empty);

        adapter
            .set_formula_cached_value(&sheet_name, row, col, &value, date_system)
            .map_err(|e| {
                anyhow!(
                    "failed to write formula cache for {}!{}{}: {e}",
                    sheet_name,
                    column_number_to_name(col),
                    row
                )
            })?;

        if let LiteralValue::Error(err) = value
            && eval_errors.len() < 200
        {
            let addr = format!("{}{}", column_number_to_name(col), row);
            eval_errors.push(format!("{}!{}: {}", sheet_name, addr, err));
        }
    }

    adapter
        .save_as_path(path)
        .map_err(|e| anyhow!("failed to save recalculated workbook {:?}: {e}", path))?;

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
    engine: &mut FormualizerEngine,
    timeout_ms: Option<u64>,
) -> Result<(u64, u64)> {
    let Some(timeout_ms) = timeout_ms else {
        let eval = engine.evaluate_all()?;
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

    let result = engine.evaluate_all_cancellable(cancel_flag);
    done_flag.store(true, Ordering::Relaxed);
    let _ = handle.join();

    let eval = result?;
    Ok((eval.computed_vertices as u64, eval.cycle_errors as u64))
}
