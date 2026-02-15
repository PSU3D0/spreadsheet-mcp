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
pub enum OutputShape {
    Canonical,
    Compact,
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
    about = "Stateless spreadsheet CLI for reads, edits, and diffs",
    long_about = "Stateless spreadsheet CLI for AI and automation workflows.\n\nCommon workflows:\n  • Inspect a workbook: list-sheets → sheet-overview → table-profile\n  • Find labels or values: find-value --mode label|value\n  • Copy → edit → recalculate → diff for safe what-if changes\n\nTip: global --format csv is currently unsupported and returns an error. Use --format json, or command-level CSV options such as read-table --table-format csv."
)]
pub struct Cli {
    #[arg(
        long,
        value_enum,
        default_value_t = OutputFormat::Json,
        global = true,
        help = "Output format (csv is currently unsupported globally; use json or command-specific CSV options like read-table --table-format csv)"
    )]
    pub format: OutputFormat,

    #[arg(
        long,
        value_enum,
        default_value_t = OutputShape::Canonical,
        global = true,
        help = "Output shape (canonical default keeps full schema; compact only flattens single-range range-values responses and preserves continuation fields like next_start_row)"
    )]
    pub shape: OutputShape,

    #[arg(
        long,
        global = true,
        help = "Emit compact JSON without pretty-printing"
    )]
    pub compact: bool,

    #[arg(long, global = true, help = "Suppress non-fatal warnings")]
    pub quiet: bool,

    #[command(subcommand)]
    pub command: Commands,
}

#[derive(Debug, Subcommand)]
pub enum Commands {
    #[command(about = "List workbook sheets with basic summary metadata")]
    ListSheets {
        #[arg(value_name = "FILE", help = "Path to the workbook (.xlsx/.xlsm)")]
        file: PathBuf,
    },
    #[command(about = "Inspect one sheet and detect structured regions")]
    SheetOverview {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(
            value_name = "SHEET",
            help = "Exact sheet name (quote names with spaces)"
        )]
        sheet: String,
    },
    #[command(
        about = "Read raw values for one or more A1 ranges",
        after_long_help = "Examples:\n  agent-spreadsheet range-values data.xlsx Sheet1 A1:C20\n  agent-spreadsheet range-values data.xlsx \"Q1 Actuals\" A1:B5 D10:E20\n\nShape behavior:\n  --shape canonical (default/omitted): keep values as an array of per-range entries.\n  --shape compact with one range: flatten that entry to top-level fields (range, payload, optional next_start_row).\n  --shape compact with multiple ranges: keep values as an array with per-entry range."
    )]
    RangeValues {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Sheet name containing the ranges")]
        sheet: String,
        #[arg(
            value_name = "RANGE",
            help = "One or more A1 ranges (for example A1:C10)"
        )]
        ranges: Vec<String>,
    },
    #[command(about = "Read a table-like region as json, values, or csv")]
    ReadTable {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(long, value_name = "SHEET", help = "Restrict read to a specific sheet")]
        sheet: Option<String>,
        #[arg(long, value_name = "RANGE", help = "Optional A1 range override")]
        range: Option<String>,
        #[arg(
            long = "table-format",
            value_enum,
            value_name = "FORMAT",
            help = "Output format for this command"
        )]
        table_format: Option<TableReadFormat>,
    },
    #[command(
        about = "Find cells matching a text query by value or label",
        after_long_help = "Examples:\n  agent-spreadsheet find-value data.xlsx Revenue\n  agent-spreadsheet find-value data.xlsx \"Net Income\" --sheet \"Q1 Actuals\" --mode label"
    )]
    FindValue {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "QUERY", help = "Text to search for")]
        query: String,
        #[arg(long, value_name = "SHEET", help = "Limit search to one sheet")]
        sheet: Option<String>,
        #[arg(
            long,
            value_enum,
            value_name = "MODE",
            help = "Search mode: value or label"
        )]
        mode: Option<FindValueMode>,
    },
    #[command(
        about = "Summarize formulas on a sheet by complexity or frequency",
        after_long_help = "Examples:\n  agent-spreadsheet formula-map data.xlsx Sheet1\n  agent-spreadsheet formula-map data.xlsx \"Q1 Actuals\" --sort-by count --limit 25"
    )]
    FormulaMap {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Sheet to analyze")]
        sheet: String,
        #[arg(long, value_name = "LIMIT", help = "Maximum groups to return")]
        limit: Option<u32>,
        #[arg(
            long,
            value_enum,
            value_name = "ORDER",
            help = "Sort groups by complexity or count"
        )]
        sort_by: Option<FormulaSort>,
    },
    #[command(about = "Trace formula precedents or dependents from one origin cell")]
    FormulaTrace {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Sheet containing the origin cell")]
        sheet: String,
        #[arg(value_name = "CELL", help = "Origin cell in A1 notation")]
        cell: String,
        #[arg(
            value_name = "DIRECTION",
            help = "Trace direction: precedents or dependents"
        )]
        direction: TraceDirectionArg,
    },
    #[command(about = "Describe workbook-level metadata and sheet counts")]
    Describe {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
    },
    #[command(
        about = "Profile table headers, types, and column distributions",
        after_long_help = "Examples:\n  agent-spreadsheet table-profile data.xlsx\n  agent-spreadsheet table-profile data.xlsx --sheet \"Q1 Actuals\""
    )]
    TableProfile {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(long, value_name = "SHEET", help = "Optional sheet to profile")]
        sheet: Option<String>,
    },
    #[command(about = "Copy a workbook to a new path for safe edits")]
    Copy {
        #[arg(value_name = "SOURCE", help = "Original workbook path")]
        source: PathBuf,
        #[arg(value_name = "DEST", help = "Destination workbook path")]
        dest: PathBuf,
    },
    #[command(about = "Apply one or more shorthand cell edits to a sheet")]
    Edit {
        #[arg(value_name = "FILE", help = "Workbook path to modify")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Target sheet name")]
        sheet: String,
        #[arg(
            value_name = "EDIT",
            help = "Edit operations like A1=42 or B2==SUM(A1:A10)"
        )]
        edits: Vec<String>,
    },
    #[command(about = "Recalculate workbook formulas")]
    Recalculate {
        #[arg(value_name = "FILE", help = "Workbook path to recalculate")]
        file: PathBuf,
    },
    #[command(
        about = "Diff two workbook versions and report changed cells",
        after_long_help = "Examples:\n  agent-spreadsheet diff baseline.xlsx candidate.xlsx\n  agent-spreadsheet diff data.xlsx /tmp/data-edited.xlsx"
    )]
    Diff {
        #[arg(value_name = "ORIGINAL", help = "Baseline workbook path")]
        original: PathBuf,
        #[arg(value_name = "MODIFIED", help = "Modified workbook path")]
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
    run_with_options(cli.command, cli.format, cli.shape, cli.compact, cli.quiet).await
}

pub async fn run_with_options(
    command: Commands,
    format: OutputFormat,
    shape: OutputShape,
    compact: bool,
    quiet: bool,
) -> Result<()> {
    if let Err(error) = errors::ensure_output_supported(format) {
        emit_error_and_exit(error);
    }

    match run_command(command).await {
        Ok(payload) => {
            if let Err(error) = output::emit_value(&payload, format, shape, compact, quiet) {
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
            "--shape",
            "compact",
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

        assert!(matches!(cli.shape, OutputShape::Compact));
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

        assert!(matches!(cli.shape, OutputShape::Canonical));

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
