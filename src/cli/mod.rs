pub mod commands;
pub mod errors;
pub mod output;

use anyhow::Result;
use clap::{Parser, Subcommand, ValueEnum};
use serde_json::Value;
use std::path::PathBuf;

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum OutputFormat {
    Json,
    Csv,
}

#[derive(Debug, Parser)]
#[command(
    name = "spreadsheet-cli",
    version,
    about = "Spreadsheet command line interface"
)]
pub struct Cli {
    #[arg(long, value_enum, default_value_t = OutputFormat::Json, global = true)]
    pub format: OutputFormat,

    #[arg(long, global = true)]
    pub compact: bool,

    #[arg(long, global = true)]
    pub quiet: bool,

    #[command(subcommand)]
    pub command: Commands,
}

#[derive(Debug, Subcommand)]
pub enum Commands {
    ListSheets {
        file: PathBuf,
    },
    SheetOverview {
        file: PathBuf,
        sheet: String,
    },
    RangeValues {
        file: PathBuf,
        sheet: String,
        ranges: Vec<String>,
    },
    Describe {
        file: PathBuf,
    },
    Copy {
        source: PathBuf,
        dest: PathBuf,
    },
    Edit {
        file: PathBuf,
        sheet: String,
        edits: Vec<String>,
    },
    Recalculate {
        file: PathBuf,
    },
    Diff {
        original: PathBuf,
        modified: PathBuf,
    },
}

pub async fn run_command(command: Commands) -> Result<Value> {
    match command {
        Commands::ListSheets { file } => commands::read::list_sheets(file).await,
        Commands::SheetOverview { file, sheet } => {
            commands::read::sheet_overview(file, sheet).await
        }
        Commands::RangeValues {
            file,
            sheet,
            ranges,
        } => commands::read::range_values(file, sheet, ranges).await,
        Commands::Describe { file } => commands::read::describe(file).await,
        Commands::Copy { source, dest } => commands::write::copy(source, dest).await,
        Commands::Edit { file, sheet, edits } => commands::write::edit(file, sheet, edits).await,
        Commands::Recalculate { file } => commands::recalc::recalculate(file).await,
        Commands::Diff { original, modified } => commands::diff::diff(original, modified).await,
    }
}
