use crate::cli::OutputFormat;
use crate::response_prune::prune_non_structural_empties;
use anyhow::{Result, bail};
use serde_json::Value;

pub fn emit_value(value: &Value, format: OutputFormat, compact: bool, quiet: bool) -> Result<()> {
    if matches!(format, OutputFormat::Csv) {
        bail!("csv output is not implemented yet for agent-spreadsheet")
    }

    let mut value = value.clone();
    prune_non_structural_empties(&mut value);

    let stdout = std::io::stdout();
    let mut handle = stdout.lock();
    if compact || quiet {
        serde_json::to_writer(&mut handle, &value)?;
    } else {
        serde_json::to_writer_pretty(&mut handle, &value)?;
    }
    use std::io::Write;
    handle.write_all(b"\n")?;
    Ok(())
}
