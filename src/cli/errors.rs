use crate::cli::OutputFormat;
use anyhow::{Result, bail};

pub fn ensure_output_supported(format: OutputFormat) -> Result<()> {
    match format {
        OutputFormat::Json => Ok(()),
        OutputFormat::Csv => {
            bail!("csv output is not implemented yet for spreadsheet-cli; use --format json")
        }
    }
}
