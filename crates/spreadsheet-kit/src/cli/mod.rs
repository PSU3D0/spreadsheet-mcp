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

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum TableReadFormat {
    Json,
    Values,
    Csv,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum FindValueMode {
    Value,
    Label,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum FormulaSort {
    Complexity,
    Count,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum TraceDirectionArg {
    Precedents,
    Dependents,
}

#[derive(Debug, Parser)]
#[command(
    name = "agent-spreadsheet",
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
    ReadTable {
        file: PathBuf,
        #[arg(long)]
        sheet: Option<String>,
        #[arg(long)]
        range: Option<String>,
        #[arg(long = "table-format", value_enum)]
        table_format: Option<TableReadFormat>,
    },
    FindValue {
        file: PathBuf,
        query: String,
        #[arg(long)]
        sheet: Option<String>,
        #[arg(long, value_enum)]
        mode: Option<FindValueMode>,
    },
    FormulaMap {
        file: PathBuf,
        sheet: String,
        #[arg(long)]
        limit: Option<u32>,
        #[arg(long, value_enum)]
        sort_by: Option<FormulaSort>,
    },
    FormulaTrace {
        file: PathBuf,
        sheet: String,
        cell: String,
        direction: TraceDirectionArg,
    },
    Describe {
        file: PathBuf,
    },
    TableProfile {
        file: PathBuf,
        #[arg(long)]
        sheet: Option<String>,
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
        Commands::ReadTable {
            file,
            sheet,
            range,
            table_format,
        } => commands::read::read_table(file, sheet, range, table_format).await,
        Commands::FindValue {
            file,
            query,
            sheet,
            mode,
        } => commands::read::find_value(file, query, sheet, mode).await,
        Commands::FormulaMap {
            file,
            sheet,
            limit,
            sort_by,
        } => commands::read::formula_map(file, sheet, limit, sort_by).await,
        Commands::FormulaTrace {
            file,
            sheet,
            cell,
            direction,
        } => commands::read::formula_trace(file, sheet, cell, direction).await,
        Commands::Describe { file } => commands::read::describe(file).await,
        Commands::TableProfile { file, sheet } => commands::read::table_profile(file, sheet).await,
        Commands::Copy { source, dest } => commands::write::copy(source, dest).await,
        Commands::Edit { file, sheet, edits } => commands::write::edit(file, sheet, edits).await,
        Commands::Recalculate { file } => commands::recalc::recalculate(file).await,
        Commands::Diff { original, modified } => commands::diff::diff(original, modified).await,
    }
}

pub async fn run() -> Result<()> {
    let cli = Cli::parse();
    run_with_options(cli.command, cli.format, cli.compact, cli.quiet).await
}

pub async fn run_with_options(
    command: Commands,
    format: OutputFormat,
    compact: bool,
    quiet: bool,
) -> Result<()> {
    if let Err(error) = errors::ensure_output_supported(format) {
        emit_error_and_exit(error);
    }

    match run_command(command).await {
        Ok(payload) => {
            if let Err(error) = output::emit_value(&payload, format, compact, quiet) {
                emit_error_and_exit(error);
            }
            Ok(())
        }
        Err(error) => emit_error_and_exit(error),
    }
}

fn emit_error_and_exit(error: anyhow::Error) -> ! {
    let envelope = errors::envelope_for(&error);
    let stderr = std::io::stderr();
    let mut handle = stderr.lock();
    if serde_json::to_writer(&mut handle, &envelope).is_err() {
        eprintln!("{{\"code\":\"COMMAND_FAILED\",\"message\":\"{}\"}}", error);
    } else {
        use std::io::Write;
        let _ = handle.write_all(b"\n");
    }
    std::process::exit(1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_global_flags_and_read_table() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "--format",
            "json",
            "--compact",
            "--quiet",
            "read-table",
            "workbook.xlsx",
            "--sheet",
            "Sheet1",
            "--range",
            "A1:B10",
            "--table-format",
            "values",
        ])
        .expect("parse command");

        assert!(cli.compact);
        assert!(cli.quiet);
        match cli.command {
            Commands::ReadTable {
                file,
                sheet,
                range,
                table_format,
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(sheet.as_deref(), Some("Sheet1"));
                assert_eq!(range.as_deref(), Some("A1:B10"));
                assert!(matches!(table_format, Some(TableReadFormat::Values)));
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_formula_trace_direction() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "formula-trace",
            "workbook.xlsx",
            "Sheet1",
            "C3",
            "dependents",
        ])
        .expect("parse command");

        match cli.command {
            Commands::FormulaTrace {
                direction,
                cell,
                sheet,
                ..
            } => {
                assert_eq!(cell, "C3");
                assert_eq!(sheet, "Sheet1");
                assert!(matches!(direction, TraceDirectionArg::Dependents));
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }
}
