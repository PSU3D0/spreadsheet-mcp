pub mod commands;
pub mod errors;
pub mod output;

use crate::model::FormulaParsePolicy;
use anyhow::Result;
use clap::{Parser, Subcommand, ValueEnum};
use serde_json::Value;
use std::ffi::OsString;
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
pub enum RangeValuesFormatArg {
    Json,
    Values,
    Csv,
    Dense,
    Rows,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum SheetPageFormatArg {
    #[value(name = "full")]
    Full,
    #[value(name = "compact")]
    Compact,
    #[value(name = "values_only")]
    ValuesOnly,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum TableSampleModeArg {
    First,
    Last,
    Distributed,
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
pub enum LabelDirectionArg {
    Right,
    Below,
    Any,
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

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum LayoutModeArg {
    Values,
    Formulas,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum LayoutRenderArg {
    Json,
    Ascii,
    Both,
}

#[derive(Debug, Subcommand)]
pub enum SheetportManifestCommands {
    #[command(
        about = "Discover candidate SheetPort ports from workbook structure",
        after_long_help = "Examples:\n  agent-spreadsheet sheetport manifest candidates deal_model.xlsx\n  agent-spreadsheet sheetport manifest candidates deal_model.xlsx --sheet-filter Assumptions"
    )]
    Candidates {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(long, value_name = "SHEET", help = "Optional sheet filter")]
        sheet_filter: Option<String>,
    },
    #[command(about = "Print the canonical SheetPort JSON schema")]
    Schema,
    #[command(
        about = "Validate a SheetPort manifest",
        after_long_help = "Example:\n  agent-spreadsheet sheetport manifest validate manifest.yaml"
    )]
    Validate {
        #[arg(value_name = "MANIFEST", help = "Path to the YAML manifest")]
        manifest: PathBuf,
    },
    #[command(
        about = "Normalize a SheetPort manifest for deterministic diffs",
        after_long_help = "Examples:\n  agent-spreadsheet sheetport manifest normalize manifest.yaml\n  agent-spreadsheet sheetport manifest normalize manifest.yaml --output manifest.normalized.yaml"
    )]
    Normalize {
        #[arg(value_name = "MANIFEST", help = "Path to the YAML manifest")]
        manifest: PathBuf,
        #[arg(long, value_name = "PATH", help = "Write normalized YAML to this file")]
        output: Option<PathBuf>,
    },
}

#[derive(Debug, Subcommand)]
pub enum SheetportCommands {
    #[command(about = "Manifest lifecycle helpers", subcommand)]
    Manifest(SheetportManifestCommands),
    #[command(
        about = "Bind-check a workbook against a SheetPort manifest",
        after_long_help = "Example:\n  agent-spreadsheet sheetport bind-check deal_model.xlsx manifest.yaml"
    )]
    BindCheck {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "MANIFEST", help = "Path to the YAML manifest")]
        manifest: PathBuf,
    },
    #[command(
        about = "Execute a SheetPort manifest with JSON inputs",
        after_long_help = "Examples:\n  agent-spreadsheet sheetport run data.xlsx manifest.yaml --inputs '{\"loan\": 10000}'\n  agent-spreadsheet sheetport run data.xlsx manifest.yaml"
    )]
    Run {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "MANIFEST", help = "Path to the YAML manifest")]
        manifest: PathBuf,
        #[arg(long, help = "JSON string or @file containing input arguments")]
        inputs: Option<String>,
        #[arg(long, help = "Seed for deterministic RNG evaluation")]
        rng_seed: Option<u64>,
        #[arg(long, help = "Freeze volatile functions (e.g. NOW(), RAND())")]
        freeze_volatile: bool,
    },
}

#[derive(Debug, Subcommand)]
pub enum SessionCommands {
    #[command(about = "Start a new session tracking a base workbook file")]
    Start {
        #[arg(long, value_name = "FILE", help = "Path to the base workbook")]
        base: PathBuf,
        #[arg(long, value_name = "LABEL", help = "Human-readable session label")]
        label: Option<String>,
        #[arg(
            long,
            value_name = "PATH",
            help = "Workspace root directory (default: cwd)"
        )]
        workspace: Option<PathBuf>,
    },
    #[command(about = "View the event timeline for a session")]
    Log {
        #[arg(long, value_name = "ID", help = "Session identifier")]
        session: String,
        #[arg(long, value_name = "OP_ID", help = "Show events since this op_id")]
        since: Option<String>,
        #[arg(
            long,
            value_name = "KIND",
            help = "Filter by operation kind prefix (e.g. structure)"
        )]
        kind: Option<String>,
        #[arg(long, value_name = "PATH", help = "Workspace root directory")]
        workspace: Option<PathBuf>,
    },
    #[command(about = "List branches in a session")]
    Branches {
        #[arg(long, value_name = "ID", help = "Session identifier")]
        session: String,
        #[arg(long, value_name = "PATH", help = "Workspace root directory")]
        workspace: Option<PathBuf>,
    },
    #[command(about = "Switch to a different branch")]
    Switch {
        #[arg(long, value_name = "ID", help = "Session identifier")]
        session: String,
        #[arg(long, value_name = "NAME", help = "Branch name to switch to")]
        branch: String,
        #[arg(long, value_name = "PATH", help = "Workspace root directory")]
        workspace: Option<PathBuf>,
    },
    #[command(about = "Set HEAD to a specific event (time-travel)")]
    Checkout {
        #[arg(long, value_name = "ID", help = "Session identifier")]
        session: String,
        #[arg(value_name = "OP_ID", help = "Event identifier to checkout")]
        op_id: String,
        #[arg(long, value_name = "PATH", help = "Workspace root directory")]
        workspace: Option<PathBuf>,
    },
    #[command(about = "Move HEAD back one event (branch-local undo)")]
    Undo {
        #[arg(long, value_name = "ID", help = "Session identifier")]
        session: String,
        #[arg(long, value_name = "PATH", help = "Workspace root directory")]
        workspace: Option<PathBuf>,
    },
    #[command(about = "Move HEAD forward one event (branch-local redo)")]
    Redo {
        #[arg(long, value_name = "ID", help = "Session identifier")]
        session: String,
        #[arg(long, value_name = "PATH", help = "Workspace root directory")]
        workspace: Option<PathBuf>,
    },
    #[command(about = "Create a new branch forking from a given event")]
    Fork {
        #[arg(long, value_name = "ID", help = "Session identifier")]
        session: String,
        #[arg(
            long,
            value_name = "OP_ID",
            help = "Fork from this event (default: current HEAD)"
        )]
        from: Option<String>,
        #[arg(long, value_name = "LABEL", help = "Human-readable branch label")]
        label: Option<String>,
        #[arg(value_name = "NAME", help = "New branch name")]
        branch_name: String,
        #[arg(long, value_name = "PATH", help = "Workspace root directory")]
        workspace: Option<PathBuf>,
    },
    #[command(
        about = "Stage an operation (compute dry-run impact without advancing HEAD)",
        after_long_help = "Canonical session payload contract:\n  • Every payload must include a top-level kind field.\n  • transform.write_matrix is a flat object with sheet_name/anchor/rows.\n  • Batch families use an ops array envelope.\n\nExamples:\n  asp session op --session sess_abc123 --ops @write_matrix.json\n\n  write_matrix.json\n  {\n    \"kind\": \"transform.write_matrix\",\n    \"sheet_name\": \"Sheet1\",\n    \"anchor\": \"B7\",\n    \"rows\": [[\"Revenue\", 100]]\n  }\n\n  asp session op --session sess_abc123 --ops @structure_ops.json\n\n  structure_ops.json\n  {\n    \"kind\": \"structure.insert_rows\",\n    \"ops\": [{ \"sheet_name\": \"Sheet1\", \"at\": 12, \"count\": 2 }]\n  }"
    )]
    Op {
        #[arg(long, value_name = "ID", help = "Session identifier")]
        session: String,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)"
        )]
        ops: String,
        #[arg(long, value_name = "PATH", help = "Workspace root directory")]
        workspace: Option<PathBuf>,
    },
    #[command(about = "Apply a staged operation (compare-and-swap against current HEAD)")]
    Apply {
        #[arg(long, value_name = "ID", help = "Session identifier")]
        session: String,
        #[arg(value_name = "STAGED_ID", help = "Staged operation identifier")]
        staged_id: String,
        #[arg(long, value_name = "PATH", help = "Workspace root directory")]
        workspace: Option<PathBuf>,
    },
    #[command(about = "Compile the current HEAD into a standalone Excel file")]
    Materialize {
        #[arg(long, value_name = "ID", help = "Session identifier")]
        session: String,
        #[arg(long, value_name = "PATH", help = "Output file path")]
        output: PathBuf,
        #[arg(long, help = "Allow overwriting existing output file")]
        force: bool,
        #[arg(long, value_name = "PATH", help = "Workspace root directory")]
        workspace: Option<PathBuf>,
    },
}

#[derive(Debug, Subcommand)]
pub enum DiscoverabilityCommands {
    #[command(about = "Schema/example target for transform-batch payloads")]
    TransformBatch,
    #[command(about = "Schema/example target for style-batch payloads")]
    StyleBatch,
    #[command(about = "Schema/example target for apply-formula-pattern payloads")]
    ApplyFormulaPattern,
    #[command(about = "Schema/example target for structure-batch payloads")]
    StructureBatch,
    #[command(about = "Schema/example target for column-size-batch payloads")]
    ColumnSizeBatch,
    #[command(about = "Schema/example target for sheet-layout-batch payloads")]
    SheetLayoutBatch,
    #[command(about = "Schema/example target for rules-batch payloads")]
    RulesBatch,
    #[command(about = "Schema/example target for event-sourced session op payloads")]
    SessionOp {
        #[arg(
            value_name = "KIND",
            help = "Exact session op kind, e.g. transform.write_matrix"
        )]
        kind: String,
    },
}

#[derive(Debug, Parser)]
#[command(
    name = "asp",
    version,
    about = "Stateless spreadsheet CLI for reads, edits, and diffs",
    long_about = "Stateless spreadsheet CLI for AI and automation workflows.\n\nPrimary command: asp\nCompatibility alias: agent-spreadsheet\n\nVerify install:\n  asp --version\n  asp --help\n\nCommon workflows:\n  • Inspect a workbook: list-sheets → sheet-overview → table-profile\n  • Deterministic pagination loops: sheet-page (--format + next_start_row) and read-table (--limit/--offset + next_offset)\n  • Find labels or values: find-value --mode label|value\n  • Discover payload contracts: schema <target> / example <target>\n  • Stateless batch writes: transform/style/formula/structure/column/layout/rules via --ops @ops.json + one mode (--dry-run|--in-place|--output)\n  • Copy → edit → recalculate → diff for safe what-if changes\n  • SheetPort manifest loop: sheetport manifest candidates → draft/edit YAML → sheetport manifest validate → sheetport bind-check → sheetport run\n\nTip: global --output-format csv is currently unsupported and returns an error. Use --output-format json, or command-level CSV options such as read-table --table-format csv."
)]
pub struct Cli {
    #[arg(
        long = "output-format",
        value_enum,
        default_value_t = OutputFormat::Json,
        global = true,
        help = "Output format (csv is currently unsupported globally; use json or command-specific CSV options like read-table --table-format csv)"
    )]
    pub output_format: OutputFormat,

    #[arg(
        long,
        value_enum,
        default_value_t = OutputShape::Canonical,
        global = true,
        help = "Output shape (canonical keeps full schema; compact applies command-specific projections while preserving stable payload contracts for range-values/read-table/sheet-page; formula-trace compact omits per-layer highlights while preserving continuation fields)"
    )]
    pub shape: OutputShape,

    #[arg(
        long,
        global = true,
        help = "Emit compact JSON without pretty-printing (default behavior)"
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
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
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
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(
        about = "Read raw values for one or more A1 ranges",
        after_long_help = "Examples:\n  agent-spreadsheet range-values data.xlsx Sheet1 A1:C20\n  agent-spreadsheet range-values data.xlsx \"Q1 Actuals\" A1:B5 D10:E20\n  agent-spreadsheet range-values data.xlsx Sheet1 A1:C20 --include-formulas\n\nDense default:\n  range-values defaults to dense JSON encoding optimized for agent consumption:\n  dictionary + row_runs + optional sparse formulas.\n\nFormula semantics:\n  By default, range-values returns resolved values only.\n  Use --include-formulas to include formulas in the response (sparse list in dense mode, matrix in json mode).\n\nShape behavior:\n  range-values keeps a stable top-level shape in both canonical and compact modes (no single-range flattening).\n\nRelated:\n  Use inspect-cells when you need formula + value + style metadata in one response."
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
        #[arg(
            long,
            value_enum,
            value_name = "FORMAT",
            help = "Output payload format (dense default, or json/values/csv explicitly)"
        )]
        format: Option<RangeValuesFormatArg>,
        #[arg(
            long = "include-formulas",
            value_name = "BOOL",
            num_args = 0..=1,
            default_missing_value = "true",
            help = "Include formulas (sparse list in dense mode, matrix in json mode)"
        )]
        include_formulas: Option<bool>,
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(
        about = "Export a range to a specific format (e.g., csv, grid)",
        after_long_help = "Examples:\n  agent-spreadsheet range-export data.xlsx Sheet1 A1:C20 --format csv --output data.csv\n  agent-spreadsheet range-export data.xlsx Sheet1 A1:C20 --format csv --output -"
    )]
    RangeExport {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Sheet name containing the range")]
        sheet: String,
        #[arg(value_name = "RANGE", help = "A1 range (for example A1:C10)")]
        range: String,
        #[arg(long, help = "Output format (e.g. csv, grid)", default_value = "json")]
        format: String,
        #[arg(long, help = "Output path or '-' for stdout")]
        output: Option<String>,
        #[arg(
            long = "include-formulas",
            value_name = "BOOL",
            num_args = 0..=1,
            default_missing_value = "true",
            help = "Include parsed formulas in formula cells alongside evaluated values (JSON only)"
        )]
        include_formulas: Option<bool>,
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(
        about = "Import range data from grid JSON or CSV",
        after_long_help = "Examples:\n  agent-spreadsheet range-import data.xlsx Sheet1 --anchor B7 --from-grid region.json\n  agent-spreadsheet range-import data.xlsx Sheet1 --anchor B7 --from-csv data.csv --in-place"
    )]
    RangeImport {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Sheet name to import into")]
        sheet: String,
        #[arg(long, help = "Anchor cell for import (e.g. B7)")]
        anchor: String,
        #[arg(long, help = "Path to the grid JSON file to import")]
        from_grid: Option<String>,
        #[arg(long, help = "Path to the CSV file to import")]
        from_csv: Option<String>,
        #[arg(long, help = "Skip first CSV row when importing --from-csv")]
        header: bool,
        #[arg(long, help = "Clear the target area before import")]
        clear_target: bool,
        #[arg(long, help = "Validate ops without mutating files")]
        dry_run: bool,
        #[arg(long, help = "Apply imports by atomically replacing the source file")]
        in_place: bool,
        #[arg(long, help = "Apply imports to this output path")]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
    },
    #[command(
        about = "Inspect detail snapshots for targeted A1 cells/ranges (detail view, default max 25 cells)",
        after_long_help = "Examples:
  agent-spreadsheet inspect-cells data.xlsx Sheet1 A1:C3
  agent-spreadsheet inspect-cells data.xlsx \"Q1 Actuals\" D4 D7:F8
  agent-spreadsheet inspect-cells data.xlsx Sheet1 B2,C4 --include-empty
  agent-spreadsheet inspect-cells data.xlsx Sheet1 A1:J10 --budget 100

inspect-cells is a detail view for formula/value/cached/style triage and enforces a small per-request cell budget.
Use --budget to raise the limit for rect-style reads (up to 200).
For broader discovery, use sheet-page, range-values, or layout-page."
    )]
    InspectCells {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Sheet name containing the targets")]
        sheet: String,
        #[arg(
            value_name = "TARGET",
            value_delimiter = ',',
            num_args = 1..,
            help = "One or more A1 cells/ranges (e.g. B2, A1:C3, D7:F8)"
        )]
        targets: Vec<String>,
        #[arg(long, help = "Include empty cells in the response")]
        include_empty: bool,
        #[arg(
            long,
            value_name = "N",
            help = "Override the per-request cell budget (default 25, max 200)"
        )]
        budget: Option<u32>,
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(
        about = "Read one sheet page with deterministic continuation",
        after_long_help = "Examples:\n  agent-spreadsheet sheet-page data.xlsx Sheet1 --format compact --page-size 200\n  agent-spreadsheet sheet-page data.xlsx Sheet1 --format compact --page-size 200 --start-row 201\n  agent-spreadsheet sheet-page data.xlsx Sheet1 --format full --columns A,C:E --include-styles\n\nMachine contract:\n  - Inspect the top-level format field first.\n  - format=full: consume top-level rows/header_row/next_start_row.\n  - format=compact: consume compact.headers/compact.header_row/compact.rows plus next_start_row.\n  - format=values_only: consume values_only.rows plus next_start_row.\n  - Global --shape compact preserves the active sheet-page branch (no flattening).\n\nPagination loop:\n  1) Run without --start-row.\n  2) If next_start_row is present, pass it to --start-row for the next request.\n  3) Stop when next_start_row is omitted.\n\nMachine continuation example:\n  Request page 1, read next_start_row, then request page 2 with --start-row <next_start_row>."
    )]
    SheetPage {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Sheet to page through")]
        sheet: String,
        #[arg(long, value_name = "ROW", help = "1-based starting row")]
        start_row: Option<u32>,
        #[arg(
            long = "page-size",
            value_name = "N",
            help = "Rows per page (must be at least 1)"
        )]
        page_size: Option<u32>,
        #[arg(
            long,
            value_name = "COLUMNS",
            value_delimiter = ',',
            help = "Column selectors by letter/range, e.g. A,C,E:G"
        )]
        columns: Option<Vec<String>>,
        #[arg(
            long = "columns-by-header",
            value_name = "HEADERS",
            value_delimiter = ',',
            help = "Column selectors by header text (case-insensitive)"
        )]
        columns_by_header: Option<Vec<String>>,
        #[arg(
            long = "include-formulas",
            value_name = "BOOL",
            num_args = 0..=1,
            default_missing_value = "true",
            help = "Include formulas (default true)"
        )]
        include_formulas: Option<bool>,
        #[arg(
            long = "include-styles",
            value_name = "BOOL",
            num_args = 0..=1,
            default_missing_value = "true",
            help = "Include style metadata (default false)"
        )]
        include_styles: Option<bool>,
        #[arg(
            long = "include-header",
            value_name = "BOOL",
            num_args = 0..=1,
            default_missing_value = "true",
            help = "Include header row (default true)"
        )]
        include_header: Option<bool>,
        #[arg(
            long,
            value_enum,
            value_name = "FORMAT",
            required = true,
            help = "Page output format: full, compact, or values_only"
        )]
        format: SheetPageFormatArg,
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(
        about = "Read a table-like region as json, values, or csv",
        after_long_help = "Examples:\n  agent-spreadsheet read-table data.xlsx --sheet Sheet1 --table-format values\n  agent-spreadsheet read-table data.xlsx --sheet Sheet1 --table-format csv --limit 50 --offset 0\n  agent-spreadsheet read-table data.xlsx --table-name SalesTable --sample-mode distributed --limit 20\n\nPagination loop:\n  Repeat with --offset set to next_offset until next_offset is omitted."
    )]
    ReadTable {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(long, value_name = "SHEET", help = "Restrict read to a specific sheet")]
        sheet: Option<String>,
        #[arg(long, value_name = "RANGE", help = "Optional A1 range override")]
        range: Option<String>,
        #[arg(long, value_name = "NAME", help = "Read from a named Excel table")]
        table_name: Option<String>,
        #[arg(long, value_name = "ID", help = "Read from a detected region id")]
        region_id: Option<u32>,
        #[arg(
            long,
            value_name = "LIMIT",
            help = "Maximum rows to return (must be at least 1)"
        )]
        limit: Option<u32>,
        #[arg(long, value_name = "OFFSET", help = "Row offset for pagination")]
        offset: Option<u32>,
        #[arg(
            long = "sample-mode",
            value_enum,
            value_name = "MODE",
            help = "Sampling mode: first, last, or distributed"
        )]
        sample_mode: Option<TableSampleModeArg>,
        #[arg(
            long = "filters-json",
            value_name = "JSON",
            help = "Inline JSON array of filters (mutually exclusive with --filters-file)"
        )]
        filters_json: Option<String>,
        #[arg(
            long = "filters-file",
            value_name = "PATH",
            help = "Path to JSON array of filters (mutually exclusive with --filters-json)"
        )]
        filters_file: Option<PathBuf>,
        #[arg(
            long = "table-format",
            value_enum,
            value_name = "FORMAT",
            help = "Output format for this command"
        )]
        table_format: Option<TableReadFormat>,
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(
        about = "Find cells matching a text query by value or label",
        after_long_help = "Examples:\n  agent-spreadsheet find-value data.xlsx Revenue --mode value\n  agent-spreadsheet find-value data.xlsx \"Net Income\" --sheet \"Q1 Actuals\" --mode label --label-direction below\n\nLabel mode behavior:\n  - QUERY is matched against label cells.\n  - Result value is taken from an adjacent cell, not from the label itself.\n  - --label-direction any (default) checks right first, then below."
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
        #[arg(
            long = "label-direction",
            value_enum,
            value_name = "DIR",
            help = "For --mode label, read the value from right, below, or any (default: any)"
        )]
        label_direction: Option<LabelDirectionArg>,
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(
        about = "List workbook named ranges and table/formula named items",
        after_long_help = "Examples:\n  agent-spreadsheet named-ranges data.xlsx\n  agent-spreadsheet named-ranges data.xlsx --sheet \"Q1 Actuals\" --name-prefix Sales"
    )]
    NamedRanges {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(long, value_name = "SHEET", help = "Optional sheet name filter")]
        sheet: Option<String>,
        #[arg(
            long = "name-prefix",
            value_name = "PREFIX",
            help = "Optional case-insensitive prefix filter for item names"
        )]
        name_prefix: Option<String>,
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(
        about = "Define a new named range in a workbook",
        after_long_help = "Examples:\n  agent-spreadsheet define-name data.xlsx MyRange 'Sheet1!$A$1:$B$10'\n  agent-spreadsheet define-name data.xlsx SheetLocal 'Sheet1!$A$1' --scope sheet --scope-sheet-name Sheet1 --in-place"
    )]
    DefineName {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "NAME", help = "Name to define")]
        name: String,
        #[arg(value_name = "REFERS_TO", help = "Range or formula the name refers to")]
        refers_to: String,
        #[arg(
            long,
            value_name = "SCOPE",
            help = "Scope: workbook (default) or sheet"
        )]
        scope: Option<String>,
        #[arg(
            long = "scope-sheet-name",
            value_name = "SHEET",
            help = "Sheet name when scope is 'sheet'"
        )]
        scope_sheet_name: Option<String>,
        #[arg(long, help = "Validate without mutating files")]
        dry_run: bool,
        #[arg(long, help = "Apply by atomically replacing the source file")]
        in_place: bool,
        #[arg(long, value_name = "PATH", help = "Apply to this output path")]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
    },
    #[command(
        about = "Update an existing named range",
        after_long_help = "Examples:\n  agent-spreadsheet update-name data.xlsx MyRange 'Sheet1!$A$1:$C$20' --in-place\n  agent-spreadsheet update-name data.xlsx SheetLocal --scope sheet --scope-sheet-name Sheet1 --in-place\n\nNote: REFERS_TO is optional. Omit it to update scope metadata only."
    )]
    UpdateName {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "NAME", help = "Name to update")]
        name: String,
        #[arg(
            value_name = "REFERS_TO",
            help = "Optional new range or formula the name refers to"
        )]
        refers_to: Option<String>,
        #[arg(long, value_name = "SCOPE", help = "Scope filter: workbook or sheet")]
        scope: Option<String>,
        #[arg(
            long = "scope-sheet-name",
            value_name = "SHEET",
            help = "Sheet name to disambiguate"
        )]
        scope_sheet_name: Option<String>,
        #[arg(long, help = "Validate without mutating files")]
        dry_run: bool,
        #[arg(long, help = "Apply by atomically replacing the source file")]
        in_place: bool,
        #[arg(long, value_name = "PATH", help = "Apply to this output path")]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
    },
    #[command(
        about = "Delete a named range from a workbook",
        after_long_help = "Examples:\n  agent-spreadsheet delete-name data.xlsx MyRange --in-place\n  agent-spreadsheet delete-name data.xlsx SheetLocal --scope sheet --scope-sheet-name Sheet1 --in-place"
    )]
    DeleteName {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "NAME", help = "Name to delete")]
        name: String,
        #[arg(long, value_name = "SCOPE", help = "Scope filter: workbook or sheet")]
        scope: Option<String>,
        #[arg(
            long = "scope-sheet-name",
            value_name = "SHEET",
            help = "Sheet name to disambiguate"
        )]
        scope_sheet_name: Option<String>,
        #[arg(long, help = "Validate without mutating files")]
        dry_run: bool,
        #[arg(long, help = "Apply by atomically replacing the source file")]
        in_place: bool,
        #[arg(long, value_name = "PATH", help = "Apply to this output path")]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
    },
    #[command(
        about = "Find formulas containing a text query with pagination",
        after_long_help = "Examples:\n  agent-spreadsheet find-formula data.xlsx SUM(\n  agent-spreadsheet find-formula data.xlsx VLOOKUP --sheet \"Q1 Actuals\" --limit 25 --offset 50\n\nRelated:\n  Use inspect-cells for per-cell formula/value/cached/style snapshots in a target range."
    )]
    FindFormula {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "QUERY", help = "Text to search for within formulas")]
        query: String,
        #[arg(long, value_name = "SHEET", help = "Optional sheet name filter")]
        sheet: Option<String>,
        #[arg(
            long,
            value_name = "N",
            help = "Maximum matches to return (must be at least 1)"
        )]
        limit: Option<u32>,
        #[arg(long, value_name = "N", help = "Match offset for continuation")]
        offset: Option<u32>,
    },
    #[command(
        about = "Scan workbook formulas for volatile functions",
        after_long_help = "Examples:\n  agent-spreadsheet scan-volatiles data.xlsx\n  agent-spreadsheet scan-volatiles data.xlsx --sheet \"Q1 Actuals\" --limit 10 --offset 10"
    )]
    ScanVolatiles {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(long, value_name = "SHEET", help = "Optional sheet name filter")]
        sheet: Option<String>,
        #[arg(
            long,
            value_name = "N",
            help = "Maximum entries to return (must be at least 1)"
        )]
        limit: Option<u32>,
        #[arg(long, value_name = "N", help = "Entry offset for continuation")]
        offset: Option<u32>,
        #[arg(
            long = "formula-parse-policy",
            value_enum,
            value_name = "POLICY",
            help = "Formula parse policy: fail, warn (default), or off"
        )]
        formula_parse_policy: Option<FormulaParsePolicy>,
    },
    #[command(
        about = "Compute per-sheet statistics for density and column types",
        after_long_help = "Examples:\n  agent-spreadsheet sheet-statistics data.xlsx Sheet1\n  agent-spreadsheet sheet-statistics data.xlsx \"Q1 Actuals\""
    )]
    SheetStatistics {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Sheet to summarize")]
        sheet: String,
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
        #[arg(
            long = "formula-parse-policy",
            value_enum,
            value_name = "POLICY",
            help = "Formula parse policy: fail, warn (default), or off"
        )]
        formula_parse_policy: Option<FormulaParsePolicy>,
    },
    #[command(
        about = "Trace formula precedents or dependents from one origin cell",
        after_long_help = "Examples:\n  agent-spreadsheet formula-trace data.xlsx Sheet1 C2 precedents --depth 2\n  agent-spreadsheet formula-trace data.xlsx Sheet1 C2 dependents --page-size 25\n  agent-spreadsheet formula-trace data.xlsx Sheet1 C2 precedents --cursor-depth 1 --cursor-offset 25\n\nContinuation:\n  Reuse next_cursor.depth/next_cursor.offset as --cursor-depth/--cursor-offset to continue paged traces.\n\nRelated:\n  Use inspect-cells for a local per-cell triage view that includes formula/value/cached/style metadata."
    )]
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
        #[arg(
            long,
            value_name = "DEPTH",
            help = "Trace depth (must be between 1 and 5)"
        )]
        depth: Option<u32>,
        #[arg(
            long = "page-size",
            value_name = "N",
            help = "Page size for trace edges (must be between 5 and 200)"
        )]
        page_size: Option<usize>,
        #[arg(
            long = "cursor-depth",
            value_name = "DEPTH",
            help = "Continuation cursor depth (must be paired with --cursor-offset)"
        )]
        cursor_depth: Option<u32>,
        #[arg(
            long = "cursor-offset",
            value_name = "OFFSET",
            help = "Continuation cursor offset (must be paired with --cursor-depth)"
        )]
        cursor_offset: Option<usize>,
        #[arg(
            long = "formula-parse-policy",
            value_enum,
            value_name = "POLICY",
            help = "Formula parse policy: fail, warn (default), or off"
        )]
        formula_parse_policy: Option<FormulaParsePolicy>,
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(about = "Describe workbook-level metadata and sheet counts")]
    Describe {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
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
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(
        about = "Render a range with layout: column widths, borders, bold/italic, alignment",
        after_long_help = "Examples:\n  agent-spreadsheet layout-page data.xlsx Sheet1 --range A1:F30\n  agent-spreadsheet layout-page data.xlsx Sheet1 --range A1:H40 --render both\n  agent-spreadsheet layout-page data.xlsx Sheet1 --range B2:G20 --mode formulas\n  agent-spreadsheet layout-page data.xlsx Sheet1 --range B2:G20 --render ascii\n\nThe JSON output (default) includes per-column widths, merged cell spans, and per-cell style metadata.\nThe ASCII render gives a proportional grid with box-drawing borders and bold/italic markers.\n\nCLI notes:\n  --render ascii prints the grid directly (plain text) instead of JSON.\n  Empty edge columns are trimmed by default; use --skip-empty-columns-trim to keep them.\n\nLimits: 80 rows × 25 columns. Ranges exceeding these are silently capped."
    )]
    LayoutPage {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Sheet name")]
        sheet: String,
        #[arg(
            long,
            value_name = "RANGE",
            help = "A1 range to render (default: A1:T50)"
        )]
        range: Option<String>,
        #[arg(
            long,
            value_enum,
            value_name = "MODE",
            help = "Cell content: values (default) or formulas"
        )]
        mode: Option<LayoutModeArg>,
        #[arg(
            long = "max-col-width",
            value_name = "N",
            help = "Maximum column width in character units before truncating (default: 20)"
        )]
        max_col_width: Option<u32>,
        #[arg(
            long = "fit-columns",
            help = "Set each column width to the longest rendered cell so truncation is avoided (default off)"
        )]
        fit_columns: bool,
        #[arg(
            long = "skip-empty-columns-trim",
            help = "Disable default trimming of empty edge columns"
        )]
        skip_empty_columns_trim: bool,
        #[arg(
            long,
            value_enum,
            value_name = "RENDER",
            help = "Output format: json (default), ascii, or both"
        )]
        render: Option<LayoutRenderArg>,
        #[arg(
            long,
            value_name = "ID",
            help = "Read from a session's materialized state instead of the file"
        )]
        session: Option<String>,
        #[arg(
            long = "session-workspace",
            value_name = "PATH",
            help = "Workspace root for session resolution"
        )]
        session_workspace: Option<PathBuf>,
    },
    #[command(
        about = "Create a new workbook at a destination path",
        after_long_help = "Examples:
  agent-spreadsheet create-workbook new.xlsx
  agent-spreadsheet create-workbook model.xlsx --sheets Inputs,Calc,Output
  agent-spreadsheet create-workbook model.xlsx --overwrite"
    )]
    CreateWorkbook {
        #[arg(value_name = "PATH", help = "Destination workbook path")]
        path: PathBuf,
        #[arg(
            long,
            value_name = "SHEETS",
            value_delimiter = ',',
            help = "Comma-separated sheet names (default: Sheet1)"
        )]
        sheets: Option<Vec<String>>,
        #[arg(long, help = "Overwrite destination file when it exists")]
        overwrite: bool,
    },
    #[command(about = "Copy a workbook to a new path for safe edits")]
    Copy {
        #[arg(value_name = "SOURCE", help = "Original workbook path")]
        source: PathBuf,
        #[arg(value_name = "DEST", help = "Destination workbook path")]
        dest: PathBuf,
    },
    #[command(
        about = "Apply one or more shorthand cell edits to a sheet",
        after_long_help = r#"Examples:
  agent-spreadsheet edit workbook.xlsx Sheet1 A1=42 B2==SUM(A1:A10)
  agent-spreadsheet edit workbook.xlsx Sheet1 --dry-run A1=42 B2==SUM(A1:A10)
  agent-spreadsheet edit workbook.xlsx Sheet1 --output edited.xlsx --force A1=42 B2==SUM(A1:A10)

Mode selection:
  Default behavior (no mode flags): in-place edit of the source workbook.
  Optional explicit modes: --dry-run, --in-place, or --output <PATH>.

Formula shorthand:
  Use double equals for formulas, e.g. C2==SUM(A1:A10).
  Single equals writes a literal value/text, e.g. C2=SUM(A1:A10).

Cache note:
  Formula edits (values starting with =) clear cached results.
  Run recalculate to refresh computed values.

Diagnostics note:
  Formula writes include write_path_provenance (written_via + formula_targets)."#
    )]
    Edit {
        #[arg(value_name = "FILE", help = "Workbook path to modify")]
        file: PathBuf,
        #[arg(value_name = "SHEET", help = "Target sheet name")]
        sheet: String,
        #[arg(long, help = "Validate edits without mutating any workbook")]
        dry_run: bool,
        #[arg(long, help = "Apply edits by atomically replacing the source file")]
        in_place: bool,
        #[arg(long, value_name = "PATH", help = "Apply edits to this output path")]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
        #[arg(
            value_name = "EDIT",
            help = "Edit operations like A1=42 or B2==SUM(A1:A10)"
        )]
        edits: Vec<String>,
        #[arg(
            long = "formula-parse-policy",
            value_enum,
            value_name = "POLICY",
            help = "Formula parse policy: fail (default for edit), warn, or off"
        )]
        formula_parse_policy: Option<FormulaParsePolicy>,
    },
    #[command(
        about = "Apply stateless transform operations from an @ops payload",
        after_long_help = r#"Examples:
  agent-spreadsheet transform-batch workbook.xlsx --ops @ops.json --dry-run
  agent-spreadsheet transform-batch workbook.xlsx --ops @ops.json --in-place
  agent-spreadsheet transform-batch workbook.xlsx --ops @ops.json --output transformed.xlsx --force

Mode selection:
  Choose exactly one of --dry-run, --in-place, or --output <PATH>.

Payload examples (`--ops @transform_ops.json`):
  Minimal:
    {"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"range","range":"B2:B4"},"value":"0"}]}
  Advanced:
    {"ops":[{"kind":"replace_in_range","sheet_name":"Sheet1","target":{"kind":"region","region_id":1},"find":"N/A","replace":"","match_mode":"contains","case_sensitive":false,"include_formulas":true}]}

Required envelope:
  Top-level object with an `ops` array.
  Each op requires a `kind` discriminator and command-specific required fields.

Cache note:
  Formula writes (FillRange with is_formula, ReplaceInRange with include_formulas) clear cached results.
  Run recalculate to refresh computed values.

Diagnostics note:
  Formula writes include write_path_provenance (written_via + formula_targets)."#
    )]
    TransformBatch {
        #[arg(
            value_name = "FILE",
            help = "Workbook path to transform",
            required_unless_present = "print_schema"
        )]
        file: Option<PathBuf>,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)",
            required_unless_present = "print_schema"
        )]
        ops: Option<String>,
        #[arg(long, help = "Validate ops and report summary without mutating files")]
        dry_run: bool,
        #[arg(
            long,
            help = "Apply transforms by atomically replacing the source file"
        )]
        in_place: bool,
        #[arg(
            long,
            value_name = "PATH",
            help = "Apply transforms to this output path"
        )]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
        #[arg(
            long = "print-schema",
            hide = true,
            help = "Print the full JSON schema for the --ops payload and exit"
        )]
        print_schema: bool,
        #[arg(
            long = "formula-parse-policy",
            value_enum,
            value_name = "POLICY",
            help = "Formula parse policy: fail, warn (default for transform-batch), or off"
        )]
        formula_parse_policy: Option<FormulaParsePolicy>,
    },
    #[command(
        about = "Apply stateless style operations from an @ops payload",
        after_long_help = r#"Examples:
  agent-spreadsheet style-batch workbook.xlsx --ops @style_ops.json --dry-run
  agent-spreadsheet style-batch workbook.xlsx --ops @style_ops.json --output styled.xlsx --force

Payload examples (`--ops @style_ops.json`):
  Minimal:
    {"ops":[{"sheet_name":"Sheet1","target":{"kind":"range","range":"B2:B2"},"patch":{"font":{"bold":true}}}]}
  Advanced:
    {"ops":[{"sheet_name":"Sheet1","target":{"kind":"cells","cells":["B2","B3"]},"patch":{"number_format":"$#,##0.00","alignment":{"horizontal":"right"}},"op_mode":"merge"}]}

Required envelope:
  Top-level object with an `ops` array.
  Style ops require `sheet_name`, `target`, and `patch` (no top-level op `kind`)."#
    )]
    StyleBatch {
        #[arg(
            value_name = "FILE",
            help = "Workbook path to style",
            required_unless_present = "print_schema"
        )]
        file: Option<PathBuf>,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)",
            required_unless_present = "print_schema"
        )]
        ops: Option<String>,
        #[arg(long, help = "Validate ops and report summary without mutating files")]
        dry_run: bool,
        #[arg(long, help = "Apply style ops by atomically replacing the source file")]
        in_place: bool,
        #[arg(
            long,
            value_name = "PATH",
            help = "Apply style ops to this output path"
        )]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
        #[arg(
            long = "print-schema",
            hide = true,
            help = "Print the full JSON schema for the --ops payload and exit"
        )]
        print_schema: bool,
    },
    #[command(
        about = "Apply stateless formula pattern operations from an @ops payload",
        after_long_help = r#"Examples:
  agent-spreadsheet apply-formula-pattern workbook.xlsx --ops @formula_ops.json --in-place
  agent-spreadsheet apply-formula-pattern workbook.xlsx --ops @formula_ops.json --dry-run

Payload examples (`--ops @formula_ops.json`):
  Minimal:
    {"ops":[{"sheet_name":"Sheet1","target_range":"C2:C4","anchor_cell":"C2","base_formula":"B2*2"}]}
  Advanced:
    {"ops":[{"sheet_name":"Sheet1","target_range":"C2:E4","anchor_cell":"C2","base_formula":"B2*2","fill_direction":"both","relative_mode":"excel"}]}

Required envelope:
  Top-level object with an `ops` array.
  Each op requires `sheet_name`, `target_range`, `anchor_cell`, and `base_formula`.
  `relative_mode` valid values: excel|abs_cols|abs_rows.

Cache note:
  Updated formula cells clear cached results. Run recalculate to refresh computed values.

Diagnostics note:
  Formula writes include write_path_provenance (written_via + formula_targets)."#
    )]
    ApplyFormulaPattern {
        #[arg(
            value_name = "FILE",
            help = "Workbook path to update",
            required_unless_present = "print_schema"
        )]
        file: Option<PathBuf>,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)",
            required_unless_present = "print_schema"
        )]
        ops: Option<String>,
        #[arg(long, help = "Validate ops and report summary without mutating files")]
        dry_run: bool,
        #[arg(
            long,
            help = "Apply formula pattern ops by atomically replacing the source file"
        )]
        in_place: bool,
        #[arg(
            long,
            value_name = "PATH",
            help = "Apply formula pattern ops to this output path"
        )]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
        #[arg(
            long = "print-schema",
            hide = true,
            help = "Print the full JSON schema for the --ops payload and exit"
        )]
        print_schema: bool,
    },
    #[command(
        about = "Apply stateless structure operations from an @ops payload",
        after_long_help = r#"Examples:
  agent-spreadsheet structure-batch workbook.xlsx --ops @structure_ops.json --dry-run
  agent-spreadsheet structure-batch workbook.xlsx --ops @structure_ops.json --output structured.xlsx

Payload examples (`--ops @structure_ops.json`):
  Minimal:
    {"ops":[{"kind":"rename_sheet","old_name":"Summary","new_name":"Dashboard"}]}
  Advanced:
    {"ops":[{"kind":"copy_range","sheet_name":"Sheet1","dest_sheet_name":"Summary","src_range":"A1:C4","dest_anchor":"A1","include_styles":true,"include_formulas":true}]}

Required envelope:
  Top-level object with an `ops` array.
  Each op requires a `kind` discriminator and kind-specific required fields.

Cache note:
  Structural operations that rewrite formula references (row/column insert/delete, sheet rename,
  copy/move) clear cached formula results. Run recalculate to refresh computed values."#
    )]
    StructureBatch {
        #[arg(
            value_name = "FILE",
            help = "Workbook path to update",
            required_unless_present = "print_schema"
        )]
        file: Option<PathBuf>,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)",
            required_unless_present = "print_schema"
        )]
        ops: Option<String>,
        #[arg(long, help = "Validate ops and report summary without mutating files")]
        dry_run: bool,
        #[arg(
            long,
            help = "Apply structure ops by atomically replacing the source file"
        )]
        in_place: bool,
        #[arg(
            long,
            value_name = "PATH",
            help = "Apply structure ops to this output path"
        )]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
        #[arg(
            long = "print-schema",
            hide = true,
            help = "Print the full JSON schema for the --ops payload and exit"
        )]
        print_schema: bool,
        #[arg(
            long = "formula-parse-policy",
            value_enum,
            value_name = "POLICY",
            help = "Formula parse policy: fail, warn (default for structure-batch), or off"
        )]
        formula_parse_policy: Option<FormulaParsePolicy>,
        #[arg(
            long = "impact-report",
            help = "Include a structural impact report (shifted spans, absolute-ref warnings). Requires --dry-run."
        )]
        impact_report: bool,
        #[arg(
            long = "show-formula-delta",
            help = "Include before/after formula delta preview samples. Requires --dry-run."
        )]
        show_formula_delta: bool,
    },
    #[command(
        about = "Analyze structural operation impact without mutation (preflight ref-risk check)",
        after_long_help = r#"Examples:
  agent-spreadsheet check-ref-impact workbook.xlsx --ops @structure_ops.json
  agent-spreadsheet check-ref-impact workbook.xlsx --ops @structure_ops.json --show-formula-delta

Payload format is the same as structure-batch --ops.
This command is read-only: it never modifies the workbook.

Output includes:
  - shifted_spans: which rows/cols shift and by how much
  - absolute_ref_warnings: $-anchored references that cross insertion/deletion boundaries
  - tokens_affected / tokens_unaffected counts
  - optional formula_delta_preview (before/after formula samples)"#
    )]
    CheckRefImpact {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path) \u{2014} same format as structure-batch"
        )]
        ops: String,
        #[arg(
            long = "show-formula-delta",
            help = "Include before/after formula delta preview samples"
        )]
        show_formula_delta: bool,
    },
    #[command(
        about = "Apply stateless column sizing operations from an @ops payload",
        after_long_help = r#"Examples:
  agent-spreadsheet column-size-batch workbook.xlsx --ops @column_size_ops.json --in-place
  agent-spreadsheet column-size-batch workbook.xlsx --ops @column_size_ops.json --output columns.xlsx

Payload examples (`--ops @column_size_ops.json`):
  Minimal:
    {"sheet_name":"Sheet1","ops":[{"range":"A:A","size":{"kind":"width","width_chars":12.0}}]}
  Advanced:
    {"sheet_name":"Sheet1","ops":[{"target":{"kind":"columns","range":"A:C"},"size":{"kind":"auto","min_width_chars":8.0,"max_width_chars":24.0}}]}

Required envelope:
  Preferred: top-level object with `sheet_name` and `ops`.
  Also accepted: top-level `ops` where each op includes `sheet_name`.
  Each op requires `size.kind`; canonical form also includes `target.kind:"columns"`."#
    )]
    ColumnSizeBatch {
        #[arg(
            value_name = "FILE",
            help = "Workbook path to update",
            required_unless_present = "print_schema"
        )]
        file: Option<PathBuf>,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)",
            required_unless_present = "print_schema"
        )]
        ops: Option<String>,
        #[arg(long, help = "Validate ops and report summary without mutating files")]
        dry_run: bool,
        #[arg(
            long,
            help = "Apply column sizing ops by atomically replacing the source file"
        )]
        in_place: bool,
        #[arg(
            long,
            value_name = "PATH",
            help = "Apply column sizing ops to this output path"
        )]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
        #[arg(
            long = "print-schema",
            hide = true,
            help = "Print the full JSON schema for the --ops payload and exit"
        )]
        print_schema: bool,
    },
    #[command(
        about = "Apply stateless sheet layout operations from an @ops payload",
        after_long_help = r#"Examples:
  agent-spreadsheet sheet-layout-batch workbook.xlsx --ops @layout_ops.json --dry-run
  agent-spreadsheet sheet-layout-batch workbook.xlsx --ops @layout_ops.json --in-place

Payload examples (`--ops @layout_ops.json`):
  Minimal:
    {"ops":[{"kind":"freeze_panes","sheet_name":"Sheet1","freeze_rows":1,"freeze_cols":1}]}
  Advanced:
    {"ops":[{"kind":"set_page_setup","sheet_name":"Sheet1","orientation":"landscape","fit_to_width":1,"fit_to_height":1}]}

Required envelope:
  Top-level object with an `ops` array.
  Each op requires a `kind` discriminator plus kind-specific required fields."#
    )]
    SheetLayoutBatch {
        #[arg(
            value_name = "FILE",
            help = "Workbook path to update",
            required_unless_present = "print_schema"
        )]
        file: Option<PathBuf>,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)",
            required_unless_present = "print_schema"
        )]
        ops: Option<String>,
        #[arg(long, help = "Validate ops and report summary without mutating files")]
        dry_run: bool,
        #[arg(
            long,
            help = "Apply sheet layout ops by atomically replacing the source file"
        )]
        in_place: bool,
        #[arg(
            long,
            value_name = "PATH",
            help = "Apply sheet layout ops to this output path"
        )]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
        #[arg(
            long = "print-schema",
            hide = true,
            help = "Print the full JSON schema for the --ops payload and exit"
        )]
        print_schema: bool,
    },
    #[command(
        about = "Apply stateless data validation and conditional format operations from an @ops payload",
        after_long_help = r##"Examples:
  agent-spreadsheet rules-batch workbook.xlsx --ops @rules_ops.json --dry-run
  agent-spreadsheet rules-batch workbook.xlsx --ops @rules_ops.json --output ruled.xlsx --force

Payload examples (`--ops @rules_ops.json`):
  Minimal:
    {"ops":[{"kind":"set_data_validation","sheet_name":"Sheet1","target_range":"B2:B4","validation":{"kind":"list","formula1":"\"A,B,C\""}}]}
  Advanced:
    {"ops":[{"kind":"set_conditional_format","sheet_name":"Sheet1","target_range":"C2:C10","rule":{"kind":"expression","formula":"C2>100"},"style":{"fill_color":"#FFF2CC","bold":true}}]}

Required envelope:
  Top-level object with an `ops` array.
  Each op requires a `kind` discriminator and kind-specific required fields.

Note:
  Data-validation and conditional-format formulas are rule-level (not cell-level) and do not affect
  cell formula caches. No recalculate is needed after rules-batch operations."##
    )]
    RulesBatch {
        #[arg(
            value_name = "FILE",
            help = "Workbook path to update",
            required_unless_present = "print_schema"
        )]
        file: Option<PathBuf>,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)",
            required_unless_present = "print_schema"
        )]
        ops: Option<String>,
        #[arg(long, help = "Validate ops and report summary without mutating files")]
        dry_run: bool,
        #[arg(long, help = "Apply rules ops by atomically replacing the source file")]
        in_place: bool,
        #[arg(
            long,
            value_name = "PATH",
            help = "Apply rules ops to this output path"
        )]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
        #[arg(
            long = "print-schema",
            hide = true,
            help = "Print the full JSON schema for the --ops payload and exit"
        )]
        print_schema: bool,
        #[arg(
            long = "formula-parse-policy",
            value_enum,
            value_name = "POLICY",
            help = "Formula parse policy: fail, warn (default for rules-batch), or off"
        )]
        formula_parse_policy: Option<FormulaParsePolicy>,
    },
    #[command(
        about = "SheetPort manifest lifecycle and execution commands",
        after_long_help = "Examples:\n  agent-spreadsheet sheetport manifest candidates model.xlsx\n  agent-spreadsheet sheetport manifest validate manifest.yaml\n  agent-spreadsheet sheetport bind-check model.xlsx manifest.yaml\n  agent-spreadsheet sheetport run model.xlsx manifest.yaml --inputs @inputs.json"
    )]
    Sheetport {
        #[command(subcommand)]
        command: SheetportCommands,
    },
    #[command(
        about = "Find and replace text in formula bodies (not values)",
        after_long_help = r#"Examples:
  agent-spreadsheet replace-in-formulas data.xlsx Sheet1 --find '$64' --replace '$65' --dry-run
  agent-spreadsheet replace-in-formulas data.xlsx Sheet1 --find 'SUM' --replace 'SUMIFS' --in-place
  agent-spreadsheet replace-in-formulas data.xlsx Sheet1 --find 'Sheet1!' --replace 'Sheet2!' --range A1:Z100 --output fixed.xlsx
  agent-spreadsheet replace-in-formulas data.xlsx Sheet1 --find '(?i)old_name' --replace 'new_name' --regex --in-place

Mode selection:
  Choose exactly one of --dry-run, --in-place, or --output <PATH>.

Behavior:
  Only formula-bearing cells are considered. Literal values are never touched.
  When --range is omitted, the used range of the sheet is scanned.
  Output includes a count of changed formulas and sample diffs (address, before, after).

Regex mode:
  Use --regex for regular expression patterns. Capture groups are supported in --replace (e.g. $1).

Formula parse policy:
  After replacement, each new formula is validated. Policy controls behavior on malformed results:
    warn (default) => report diagnostics and skip invalid replacements
    fail => reject and error
    off => skip validation"#
    )]
    ReplaceInFormulas {
        #[arg(value_name = "FILE", help = "Workbook path to update")]
        file: PathBuf,
        #[arg(
            value_name = "SHEET",
            help = "Sheet name containing formulas to update"
        )]
        sheet: String,
        #[arg(long, help = "Text or pattern to find in formula bodies")]
        find: String,
        #[arg(long, help = "Replacement text")]
        replace: String,
        #[arg(
            long,
            value_name = "RANGE",
            help = "Optional A1 range to scope replacement (default: used range)"
        )]
        range: Option<String>,
        #[arg(long, help = "Interpret --find as a regular expression")]
        regex: bool,
        #[arg(long, help = "Case-sensitive matching (default: true)")]
        case_sensitive: Option<bool>,
        #[arg(long, help = "Validate ops and report summary without mutating files")]
        dry_run: bool,
        #[arg(
            long,
            help = "Apply replacement by atomically replacing the source file"
        )]
        in_place: bool,
        #[arg(
            long,
            value_name = "PATH",
            help = "Apply replacement to this output path"
        )]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
        #[arg(
            long = "formula-parse-policy",
            value_enum,
            value_name = "POLICY",
            help = "Formula parse policy: warn (default), fail, or off"
        )]
        formula_parse_policy: Option<FormulaParsePolicy>,
    },
    #[command(
        about = "Recalculate workbook formulas",
        after_long_help = "Examples:\n  agent-spreadsheet recalculate data.xlsx\n  agent-spreadsheet recalculate data.xlsx --output /tmp/recalced.xlsx\n  agent-spreadsheet recalculate data.xlsx --output /tmp/recalced.xlsx --force\n\nDefault (no flags): recalculate the file in-place.\n--output <PATH>: copy source to output, recalculate the copy, leave source unchanged.\n--force: allow overwriting an existing --output file."
    )]
    Recalculate {
        #[arg(value_name = "FILE", help = "Workbook path to recalculate")]
        file: PathBuf,
        #[arg(
            long,
            value_name = "PATH",
            help = "Recalculate into this output path (source stays unchanged)"
        )]
        output: Option<PathBuf>,
        #[arg(long, help = "Allow overwriting --output when it already exists")]
        force: bool,
        #[arg(
            long = "ignore-sheets",
            value_name = "SHEETS",
            value_delimiter = ',',
            help = "Comma-separated sheet names to exclude from changed-cells summary"
        )]
        ignore_sheets: Option<Vec<String>>,
        #[arg(
            long = "changed-cells",
            help = "Include a summary of cells whose values changed after recalculation"
        )]
        changed_cells: bool,
    },
    #[command(
        about = "Compare two workbook states and verify target deltas plus error provenance",
        after_long_help = "Examples:\n  asp verify baseline.xlsx candidate.xlsx --targets Summary!B2\n  asp verify baseline.xlsx candidate.xlsx --targets Sheet1!C2,Summary!B2 --named-ranges\n\nBehavior:\n  - target_deltas compares the exact Sheet!A1 cells you request\n  - new_errors reports error cells present only in the current workbook\n  - preexisting_errors reports error cells that existed in both baseline and current\n  - --named-ranges adds added/removed/changed named range deltas"
    )]
    Verify {
        #[arg(value_name = "BASELINE", help = "Baseline workbook path")]
        baseline: PathBuf,
        #[arg(value_name = "CURRENT", help = "Current workbook path")]
        current: PathBuf,
        #[arg(
            long = "targets",
            value_name = "SHEET!CELL",
            value_delimiter = ',',
            help = "One or more Sheet!A1 targets to compare (comma-separated)"
        )]
        targets: Option<Vec<String>>,
        #[arg(
            long = "named-ranges",
            help = "Include added/removed/changed named range deltas"
        )]
        named_ranges: bool,
    },
    #[command(
        about = "Diff two workbook versions with summary-first, paged details",
        after_long_help = "Examples:\n  agent-spreadsheet diff baseline.xlsx candidate.xlsx\n  agent-spreadsheet diff baseline.xlsx candidate.xlsx --details --limit 200 --offset 0\n  agent-spreadsheet diff baseline.xlsx candidate.xlsx --sheet \"GL Data\" --range A1:P200"
    )]
    Diff {
        #[arg(value_name = "ORIGINAL", help = "Baseline workbook path")]
        original: PathBuf,
        #[arg(value_name = "MODIFIED", help = "Modified workbook path")]
        modified: PathBuf,
        #[arg(long, help = "Limit diff to one sheet name")]
        sheet: Option<String>,
        #[arg(
            long,
            value_name = "SHEETS",
            value_delimiter = ',',
            help = "Limit diff to multiple sheet names (comma-separated)"
        )]
        sheets: Option<Vec<String>>,
        #[arg(
            long,
            value_name = "A1_RANGE",
            help = "Optional A1 range filter (e.g. A1:C100)"
        )]
        range: Option<String>,
        #[arg(
            long,
            help = "Include paged change items; default output is summary-only"
        )]
        details: bool,
        #[arg(
            long,
            default_value_t = 200,
            help = "Page size for --details (1..2000)"
        )]
        limit: u32,
        #[arg(long, default_value_t = 0, help = "Offset for --details pagination")]
        offset: u32,
    },
    #[command(
        about = "Print canonical JSON schema for a command or payload target",
        after_long_help = "Examples:\n  asp schema transform-batch\n  asp schema structure-batch\n  asp schema session-op transform.write_matrix"
    )]
    Schema {
        #[command(subcommand)]
        command: DiscoverabilityCommands,
    },
    #[command(
        about = "Print a copy-pastable canonical example for a command or payload target",
        after_long_help = "Examples:\n  asp example transform-batch\n  asp example rules-batch\n  asp example session-op structure.clone_row"
    )]
    Example {
        #[command(subcommand)]
        command: DiscoverabilityCommands,
    },
    #[command(
        about = "Event-sourced session management (start, navigate, stage, apply, materialize)",
        subcommand,
        after_long_help = "Session commands provide event-sourced workbook editing with undo/redo, branching, staged apply, and payload discovery.\n\nWorkflow:\n  1. asp session start --base model.xlsx\n  2. asp example session-op transform.write_matrix\n  3. asp session op --session <id> --ops @edits.json\n  4. asp session apply --session <id> <staged_id>\n  5. asp session materialize --session <id> --output result.xlsx\n\nDiscoverability:\n  • asp schema session-op transform.write_matrix\n  • asp example session-op transform.write_matrix"
    )]
    Session(Box<SessionCommands>),
    #[command(
        about = "[Deprecated] Execute a SheetPort manifest with JSON inputs",
        after_long_help = "Use `agent-spreadsheet sheetport run ...` for new workflows.\n\nExamples:\n  agent-spreadsheet run-manifest data.xlsx manifest.yaml --inputs '{\"loan\": 10000}'\n  agent-spreadsheet sheetport run data.xlsx manifest.yaml --inputs @inputs.json"
    )]
    RunManifest {
        #[arg(value_name = "FILE", help = "Path to the workbook")]
        file: PathBuf,
        #[arg(value_name = "MANIFEST", help = "Path to the YAML manifest")]
        manifest: PathBuf,
        #[arg(long, help = "JSON string or @file containing input arguments")]
        inputs: Option<String>,
        #[arg(long, help = "Seed for deterministic RNG evaluation")]
        rng_seed: Option<u64>,
        #[arg(long, help = "Freeze volatile functions (e.g. NOW(), RAND())")]
        freeze_volatile: bool,
    },
}

pub async fn run_command(command: Commands) -> Result<Value> {
    match command {
        Commands::ListSheets {
            file,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::list_sheets(resolved).await
        }
        Commands::SheetOverview {
            file,
            sheet,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::sheet_overview(resolved, sheet).await
        }
        Commands::RangeValues {
            file,
            sheet,
            ranges,
            format,
            include_formulas,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::range_values(resolved, sheet, ranges, format, include_formulas).await
        }
        Commands::RangeExport {
            file,
            sheet,
            range,
            format,
            output,
            include_formulas,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::range_export(resolved, sheet, range, format, output, include_formulas)
                .await
        }
        Commands::RangeImport {
            file,
            sheet,
            anchor,
            from_grid,
            from_csv,
            header,
            clear_target,
            dry_run,
            in_place,
            output,
            force,
        } => {
            commands::write::range_import(
                file,
                sheet,
                anchor,
                from_grid,
                from_csv,
                header,
                clear_target,
                dry_run,
                in_place,
                output,
                force,
            )
            .await
        }
        Commands::InspectCells {
            file,
            sheet,
            targets,
            include_empty,
            budget,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::inspect_cells(resolved, sheet, targets, include_empty, budget).await
        }
        Commands::SheetPage {
            file,
            sheet,
            start_row,
            page_size,
            columns,
            columns_by_header,
            include_formulas,
            include_styles,
            include_header,
            format,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::sheet_page(
                resolved,
                sheet,
                start_row,
                page_size,
                columns,
                columns_by_header,
                include_formulas,
                include_styles,
                include_header,
                format,
            )
            .await
        }
        Commands::ReadTable {
            file,
            sheet,
            range,
            table_name,
            region_id,
            limit,
            offset,
            sample_mode,
            filters_json,
            filters_file,
            table_format,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::read_table(
                resolved,
                sheet,
                range,
                table_name,
                region_id,
                limit,
                offset,
                sample_mode,
                filters_json,
                filters_file,
                table_format,
            )
            .await
        }
        Commands::FindValue {
            file,
            query,
            sheet,
            mode,
            label_direction,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::find_value(resolved, query, sheet, mode, label_direction).await
        }
        Commands::NamedRanges {
            file,
            sheet,
            name_prefix,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::named_ranges(resolved, sheet, name_prefix).await
        }
        Commands::DefineName {
            file,
            name,
            refers_to,
            scope,
            scope_sheet_name,
            dry_run,
            in_place,
            output,
            force,
        } => {
            commands::write::define_name(
                file,
                name,
                refers_to,
                scope,
                scope_sheet_name,
                dry_run,
                in_place,
                output,
                force,
            )
            .await
        }
        Commands::UpdateName {
            file,
            name,
            refers_to,
            scope,
            scope_sheet_name,
            dry_run,
            in_place,
            output,
            force,
        } => {
            commands::write::update_name(
                file,
                name,
                refers_to,
                scope,
                scope_sheet_name,
                dry_run,
                in_place,
                output,
                force,
            )
            .await
        }
        Commands::DeleteName {
            file,
            name,
            scope,
            scope_sheet_name,
            dry_run,
            in_place,
            output,
            force,
        } => {
            commands::write::delete_name(
                file,
                name,
                scope,
                scope_sheet_name,
                dry_run,
                in_place,
                output,
                force,
            )
            .await
        }
        Commands::FindFormula {
            file,
            query,
            sheet,
            limit,
            offset,
        } => commands::read::find_formula(file, query, sheet, limit, offset).await,
        Commands::ScanVolatiles {
            file,
            sheet,
            limit,
            offset,
            formula_parse_policy,
        } => commands::read::scan_volatiles(file, sheet, limit, offset, formula_parse_policy).await,
        Commands::SheetStatistics { file, sheet } => {
            commands::read::sheet_statistics(file, sheet).await
        }
        Commands::FormulaMap {
            file,
            sheet,
            limit,
            sort_by,
            formula_parse_policy,
        } => commands::read::formula_map(file, sheet, limit, sort_by, formula_parse_policy).await,
        Commands::FormulaTrace {
            file,
            sheet,
            cell,
            direction,
            depth,
            page_size,
            cursor_depth,
            cursor_offset,
            formula_parse_policy,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::formula_trace(
                resolved,
                sheet,
                cell,
                direction,
                depth,
                page_size,
                cursor_depth,
                cursor_offset,
                formula_parse_policy,
            )
            .await
        }
        Commands::Describe {
            file,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::describe(resolved).await
        }
        Commands::TableProfile {
            file,
            sheet,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::table_profile(resolved, sheet).await
        }
        Commands::LayoutPage {
            file,
            sheet,
            range,
            mode,
            max_col_width,
            fit_columns,
            skip_empty_columns_trim,
            render,
            session,
            session_workspace,
        } => {
            let (resolved, _guard) =
                commands::read::resolve_file_or_session(file, session, session_workspace)?;
            commands::read::layout_page(
                resolved,
                sheet,
                range,
                mode,
                max_col_width,
                fit_columns,
                skip_empty_columns_trim,
                render,
            )
            .await
        }
        Commands::CreateWorkbook {
            path,
            sheets,
            overwrite,
        } => commands::write::create_workbook(path, sheets, overwrite).await,
        Commands::Copy { source, dest } => commands::write::copy(source, dest).await,
        Commands::Edit {
            file,
            sheet,
            dry_run,
            in_place,
            output,
            force,
            edits,
            formula_parse_policy,
        } => {
            commands::write::edit(
                file,
                sheet,
                edits,
                dry_run,
                in_place,
                output,
                force,
                formula_parse_policy,
            )
            .await
        }
        Commands::TransformBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
            print_schema,
            formula_parse_policy,
        } => {
            if print_schema {
                commands::write::batch_payload_schema(
                    commands::write::BatchSchemaCommand::Transform,
                )
            } else {
                let file = file.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: transform-batch requires <FILE>")
                })?;
                let ops = ops.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: transform-batch requires --ops @<path>")
                })?;
                commands::write::transform_batch(
                    file,
                    ops,
                    dry_run,
                    in_place,
                    output,
                    force,
                    formula_parse_policy,
                )
                .await
            }
        }
        Commands::StyleBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
            print_schema,
        } => {
            if print_schema {
                commands::write::batch_payload_schema(commands::write::BatchSchemaCommand::Style)
            } else {
                let file = file.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: style-batch requires <FILE>")
                })?;
                let ops = ops.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: style-batch requires --ops @<path>")
                })?;
                commands::write::style_batch(file, ops, dry_run, in_place, output, force).await
            }
        }
        Commands::ApplyFormulaPattern {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
            print_schema,
        } => {
            if print_schema {
                commands::write::batch_payload_schema(
                    commands::write::BatchSchemaCommand::ApplyFormulaPattern,
                )
            } else {
                let file = file.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: apply-formula-pattern requires <FILE>")
                })?;
                let ops = ops.ok_or_else(|| {
                    anyhow::anyhow!(
                        "invalid argument: apply-formula-pattern requires --ops @<path>"
                    )
                })?;
                commands::write::apply_formula_pattern(file, ops, dry_run, in_place, output, force)
                    .await
            }
        }
        Commands::StructureBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
            print_schema,
            formula_parse_policy,
            impact_report,
            show_formula_delta,
        } => {
            if print_schema {
                commands::write::batch_payload_schema(
                    commands::write::BatchSchemaCommand::Structure,
                )
            } else {
                let file = file.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: structure-batch requires <FILE>")
                })?;
                let ops = ops.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: structure-batch requires --ops @<path>")
                })?;
                commands::write::structure_batch(
                    file,
                    ops,
                    dry_run,
                    in_place,
                    output,
                    force,
                    formula_parse_policy,
                    impact_report,
                    show_formula_delta,
                )
                .await
            }
        }
        Commands::CheckRefImpact {
            file,
            ops,
            show_formula_delta,
        } => commands::write::check_ref_impact(file, ops, show_formula_delta).await,
        Commands::ColumnSizeBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
            print_schema,
        } => {
            if print_schema {
                commands::write::batch_payload_schema(
                    commands::write::BatchSchemaCommand::ColumnSize,
                )
            } else {
                let file = file.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: column-size-batch requires <FILE>")
                })?;
                let ops = ops.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: column-size-batch requires --ops @<path>")
                })?;
                commands::write::column_size_batch(file, ops, dry_run, in_place, output, force)
                    .await
            }
        }
        Commands::SheetLayoutBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
            print_schema,
        } => {
            if print_schema {
                commands::write::batch_payload_schema(
                    commands::write::BatchSchemaCommand::SheetLayout,
                )
            } else {
                let file = file.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: sheet-layout-batch requires <FILE>")
                })?;
                let ops = ops.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: sheet-layout-batch requires --ops @<path>")
                })?;
                commands::write::sheet_layout_batch(file, ops, dry_run, in_place, output, force)
                    .await
            }
        }
        Commands::RulesBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
            print_schema,
            formula_parse_policy,
        } => {
            if print_schema {
                commands::write::batch_payload_schema(commands::write::BatchSchemaCommand::Rules)
            } else {
                let file = file.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: rules-batch requires <FILE>")
                })?;
                let ops = ops.ok_or_else(|| {
                    anyhow::anyhow!("invalid argument: rules-batch requires --ops @<path>")
                })?;
                commands::write::rules_batch(
                    file,
                    ops,
                    dry_run,
                    in_place,
                    output,
                    force,
                    formula_parse_policy,
                )
                .await
            }
        }
        Commands::Sheetport { command } => match command {
            SheetportCommands::Manifest(manifest_command) => match manifest_command {
                SheetportManifestCommands::Candidates { file, sheet_filter } => {
                    commands::read::sheetport_manifest_candidates(file, sheet_filter).await
                }
                SheetportManifestCommands::Schema => commands::read::sheetport_manifest_schema(),
                SheetportManifestCommands::Validate { manifest } => {
                    commands::read::sheetport_manifest_validate(manifest)
                }
                SheetportManifestCommands::Normalize { manifest, output } => {
                    commands::read::sheetport_manifest_normalize(manifest, output)
                }
            },
            SheetportCommands::BindCheck { file, manifest } => {
                commands::read::sheetport_bind_check(file, manifest).await
            }
            SheetportCommands::Run {
                file,
                manifest,
                inputs,
                rng_seed,
                freeze_volatile,
            } => {
                commands::read::sheetport_run(file, manifest, inputs, rng_seed, freeze_volatile)
                    .await
            }
        },
        Commands::ReplaceInFormulas {
            file,
            sheet,
            find,
            replace,
            range,
            regex,
            case_sensitive,
            dry_run,
            in_place,
            output,
            force,
            formula_parse_policy,
        } => {
            commands::write::replace_in_formulas(
                file,
                sheet,
                find,
                replace,
                range,
                regex,
                case_sensitive.unwrap_or(true),
                dry_run,
                in_place,
                output,
                force,
                formula_parse_policy,
            )
            .await
        }
        Commands::Recalculate {
            file,
            output,
            force,
            ignore_sheets,
            changed_cells,
        } => commands::recalc::recalculate(file, output, force, ignore_sheets, changed_cells).await,
        Commands::Verify {
            baseline,
            current,
            targets,
            named_ranges,
        } => commands::verify::verify(baseline, current, targets, named_ranges).await,
        Commands::Diff {
            original,
            modified,
            sheet,
            sheets,
            range,
            details,
            limit,
            offset,
        } => {
            commands::diff::diff(
                original, modified, sheet, sheets, range, details, limit, offset,
            )
            .await
        }
        Commands::Schema { command } => run_schema_command(command),
        Commands::Example { command } => run_example_command(command),
        Commands::Session(command) => match *command {
            SessionCommands::Start {
                base,
                label,
                workspace,
            } => commands::session::session_start(base, label, workspace).await,
            SessionCommands::Log {
                session,
                since,
                kind,
                workspace,
            } => commands::session::session_log(session, workspace, since, kind).await,
            SessionCommands::Branches { session, workspace } => {
                commands::session::session_branches(session, workspace).await
            }
            SessionCommands::Switch {
                session,
                branch,
                workspace,
            } => commands::session::session_switch(session, branch, workspace).await,
            SessionCommands::Checkout {
                session,
                op_id,
                workspace,
            } => commands::session::session_checkout(session, op_id, workspace).await,
            SessionCommands::Undo { session, workspace } => {
                commands::session::session_undo(session, workspace).await
            }
            SessionCommands::Redo { session, workspace } => {
                commands::session::session_redo(session, workspace).await
            }
            SessionCommands::Fork {
                session,
                from,
                label,
                branch_name,
                workspace,
            } => {
                commands::session::session_fork(session, from, label, branch_name, workspace).await
            }
            SessionCommands::Op {
                session,
                ops,
                workspace,
            } => commands::session::session_op_stage(session, ops, workspace).await,
            SessionCommands::Apply {
                session,
                staged_id,
                workspace,
            } => commands::session::session_apply(session, staged_id, workspace).await,
            SessionCommands::Materialize {
                session,
                output,
                force,
                workspace,
            } => commands::session::session_materialize(session, output, workspace, force).await,
        },
        Commands::RunManifest {
            file,
            manifest,
            inputs,
            rng_seed,
            freeze_volatile,
        } => commands::read::sheetport_run(file, manifest, inputs, rng_seed, freeze_volatile).await,
    }
}

fn run_schema_command(command: DiscoverabilityCommands) -> Result<Value> {
    match command {
        DiscoverabilityCommands::TransformBatch => {
            commands::write::batch_payload_schema(commands::write::BatchSchemaCommand::Transform)
        }
        DiscoverabilityCommands::StyleBatch => {
            commands::write::batch_payload_schema(commands::write::BatchSchemaCommand::Style)
        }
        DiscoverabilityCommands::ApplyFormulaPattern => commands::write::batch_payload_schema(
            commands::write::BatchSchemaCommand::ApplyFormulaPattern,
        ),
        DiscoverabilityCommands::StructureBatch => {
            commands::write::batch_payload_schema(commands::write::BatchSchemaCommand::Structure)
        }
        DiscoverabilityCommands::ColumnSizeBatch => {
            commands::write::batch_payload_schema(commands::write::BatchSchemaCommand::ColumnSize)
        }
        DiscoverabilityCommands::SheetLayoutBatch => {
            commands::write::batch_payload_schema(commands::write::BatchSchemaCommand::SheetLayout)
        }
        DiscoverabilityCommands::RulesBatch => {
            commands::write::batch_payload_schema(commands::write::BatchSchemaCommand::Rules)
        }
        DiscoverabilityCommands::SessionOp { kind } => {
            commands::session::session_payload_schema(kind)
        }
    }
}

fn run_example_command(command: DiscoverabilityCommands) -> Result<Value> {
    match command {
        DiscoverabilityCommands::TransformBatch => {
            commands::write::batch_payload_example(commands::write::BatchSchemaCommand::Transform)
        }
        DiscoverabilityCommands::StyleBatch => {
            commands::write::batch_payload_example(commands::write::BatchSchemaCommand::Style)
        }
        DiscoverabilityCommands::ApplyFormulaPattern => commands::write::batch_payload_example(
            commands::write::BatchSchemaCommand::ApplyFormulaPattern,
        ),
        DiscoverabilityCommands::StructureBatch => {
            commands::write::batch_payload_example(commands::write::BatchSchemaCommand::Structure)
        }
        DiscoverabilityCommands::ColumnSizeBatch => {
            commands::write::batch_payload_example(commands::write::BatchSchemaCommand::ColumnSize)
        }
        DiscoverabilityCommands::SheetLayoutBatch => {
            commands::write::batch_payload_example(commands::write::BatchSchemaCommand::SheetLayout)
        }
        DiscoverabilityCommands::RulesBatch => {
            commands::write::batch_payload_example(commands::write::BatchSchemaCommand::Rules)
        }
        DiscoverabilityCommands::SessionOp { kind } => {
            commands::session::session_payload_example(kind)
        }
    }
}

fn first_subcommand_index(argv: &[OsString]) -> Option<usize> {
    let mut expect_global_value = false;

    for (index, arg) in argv.iter().enumerate().skip(1) {
        let token = arg.to_string_lossy();

        if expect_global_value {
            expect_global_value = false;
            continue;
        }

        match token.as_ref() {
            "--output-format" | "--shape" | "--format" => {
                expect_global_value = true;
                continue;
            }
            "--compact" | "--quiet" => continue,
            _ => {}
        }

        if token.starts_with("--output-format=")
            || token.starts_with("--shape=")
            || token.starts_with("--format=")
        {
            continue;
        }

        if token.starts_with('-') {
            continue;
        }

        return Some(index);
    }

    None
}

fn is_legacy_output_format(value: &str) -> bool {
    matches!(value, "json" | "csv")
}

fn normalize_legacy_global_format_argv(argv: Vec<OsString>) -> Vec<OsString> {
    if argv.len() <= 1 {
        return argv;
    }

    let first_subcommand_index = first_subcommand_index(&argv);
    let first_subcommand_name = first_subcommand_index
        .map(|index| argv[index].to_string_lossy().into_owned())
        .unwrap_or_default();
    let preserve_sheet_page_format = first_subcommand_name == "sheet-page"
        || first_subcommand_name == "range-export"
        || first_subcommand_name == "range-values";

    let mut normalized = Vec::with_capacity(argv.len());
    normalized.push(argv[0].clone());

    let mut index = 1usize;
    while index < argv.len() {
        let token = argv[index].to_string_lossy();
        let can_rewrite_here = !preserve_sheet_page_format
            || first_subcommand_index
                .map(|subcommand_index| index < subcommand_index)
                .unwrap_or(true);

        if can_rewrite_here && token == "--format" && index + 1 < argv.len() {
            let value = argv[index + 1].to_string_lossy();
            if is_legacy_output_format(value.as_ref()) {
                normalized.push(OsString::from("--output-format"));
                normalized.push(argv[index + 1].clone());
                index += 2;
                continue;
            }
        }

        if can_rewrite_here
            && let Some(value) = token.strip_prefix("--format=")
            && is_legacy_output_format(value)
        {
            normalized.push(OsString::from(format!("--output-format={value}")));
            index += 1;
            continue;
        }

        normalized.push(argv[index].clone());
        index += 1;
    }

    normalized
}

pub async fn run() -> Result<()> {
    let argv = normalize_legacy_global_format_argv(std::env::args_os().collect());
    let cli = Cli::parse_from(argv);
    run_with_options(
        cli.command,
        cli.output_format,
        cli.shape,
        cli.compact,
        cli.quiet,
    )
    .await
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

    let projection_target = compact_projection_target_for_command(&command);
    let emit_layout_ascii_direct = matches!(
        &command,
        Commands::LayoutPage {
            render: Some(LayoutRenderArg::Ascii),
            ..
        }
    );

    match run_command(command).await {
        Ok(payload) => {
            if emit_layout_ascii_direct {
                if let Some(ascii) = payload.get("ascii_render").and_then(|v| v.as_str()) {
                    print!("{ascii}");
                    if !ascii.ends_with('\n') {
                        println!();
                    }
                    return Ok(());
                }
                emit_error_and_exit(anyhow::anyhow!(
                    "layout-page --render ascii expected ascii_render in response"
                ));
            }

            if let Err(error) =
                output::emit_value(&payload, format, shape, projection_target, compact, quiet)
            {
                emit_error_and_exit(error);
            }
            Ok(())
        }
        Err(error) => emit_error_and_exit(error),
    }
}

fn compact_projection_target_for_command(command: &Commands) -> output::CompactProjectionTarget {
    match command {
        Commands::RangeValues { .. } => output::CompactProjectionTarget::RangeValues,
        Commands::ReadTable { .. } => output::CompactProjectionTarget::ReadTable,
        Commands::SheetPage { .. } => output::CompactProjectionTarget::SheetPage,
        Commands::FormulaTrace { .. } => output::CompactProjectionTarget::FormulaTrace,
        _ => output::CompactProjectionTarget::None,
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
            "--output-format",
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
            "--table-name",
            "SalesTable",
            "--region-id",
            "7",
            "--limit",
            "10",
            "--offset",
            "2",
            "--sample-mode",
            "first",
            "--filters-json",
            r#"[{"column":"Name","op":"eq","value":"Alice"}]"#,
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
                table_name,
                region_id,
                limit,
                offset,
                sample_mode,
                filters_json,
                filters_file,
                table_format,
                ..
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(sheet.as_deref(), Some("Sheet1"));
                assert_eq!(range.as_deref(), Some("A1:B10"));
                assert_eq!(table_name.as_deref(), Some("SalesTable"));
                assert_eq!(region_id, Some(7));
                assert_eq!(limit, Some(10));
                assert_eq!(offset, Some(2));
                assert!(matches!(sample_mode, Some(TableSampleModeArg::First)));
                assert_eq!(
                    filters_json.as_deref(),
                    Some(r#"[{"column":"Name","op":"eq","value":"Alice"}]"#)
                );
                assert!(filters_file.is_none());
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
            "--depth",
            "2",
            "--page-size",
            "15",
            "--cursor-depth",
            "2",
            "--cursor-offset",
            "5",
        ])
        .expect("parse command");

        assert!(matches!(cli.shape, OutputShape::Canonical));

        match cli.command {
            Commands::FormulaTrace {
                direction,
                cell,
                sheet,
                depth,
                page_size,
                cursor_depth,
                cursor_offset,
                ..
            } => {
                assert_eq!(cell, "C3");
                assert_eq!(sheet, "Sheet1");
                assert_eq!(depth, Some(2));
                assert_eq!(page_size, Some(15));
                assert_eq!(cursor_depth, Some(2));
                assert_eq!(cursor_offset, Some(5));
                assert!(matches!(direction, TraceDirectionArg::Dependents));
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_range_values_include_formulas_flag() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "range-values",
            "workbook.xlsx",
            "Sheet1",
            "A1:C10",
            "--include-formulas",
        ])
        .expect("parse command");

        match cli.command {
            Commands::RangeValues {
                file,
                sheet,
                ranges,
                format,
                include_formulas,
                ..
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(sheet, "Sheet1");
                assert_eq!(ranges, vec!["A1:C10".to_string()]);
                assert!(format.is_none());
                assert_eq!(include_formulas, Some(true));
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_range_values_format_argument() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "range-values",
            "workbook.xlsx",
            "Sheet1",
            "A1:C10",
            "--format",
            "json",
        ])
        .expect("parse command");

        match cli.command {
            Commands::RangeValues { format, .. } => {
                assert!(matches!(format, Some(RangeValuesFormatArg::Json)));
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_diff_arguments_with_paging_and_filters() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "diff",
            "baseline.xlsx",
            "candidate.xlsx",
            "--sheet",
            "Sheet1",
            "--range",
            "A1:C20",
            "--details",
            "--limit",
            "150",
            "--offset",
            "300",
        ])
        .expect("parse diff command");

        match cli.command {
            Commands::Diff {
                original,
                modified,
                sheet,
                sheets,
                range,
                details,
                limit,
                offset,
            } => {
                assert_eq!(original, PathBuf::from("baseline.xlsx"));
                assert_eq!(modified, PathBuf::from("candidate.xlsx"));
                assert_eq!(sheet.as_deref(), Some("Sheet1"));
                assert!(sheets.is_none());
                assert_eq!(range.as_deref(), Some("A1:C20"));
                assert!(details);
                assert_eq!(limit, 150);
                assert_eq!(offset, 300);
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_diff_defaults_to_summary_only() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "diff",
            "baseline.xlsx",
            "candidate.xlsx",
        ])
        .expect("parse diff command defaults");

        match cli.command {
            Commands::Diff {
                details,
                limit,
                offset,
                ..
            } => {
                assert!(!details);
                assert_eq!(limit, 200);
                assert_eq!(offset, 0);
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_range_import_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "range-import",
            "workbook.xlsx",
            "Sheet1",
            "--anchor",
            "B7",
            "--from-grid",
            "region.json",
            "--in-place",
        ])
        .expect("parse range-import");

        match cli.command {
            Commands::RangeImport {
                file,
                sheet,
                anchor,
                from_grid,
                from_csv,
                header,
                clear_target,
                dry_run,
                in_place,
                output,
                force,
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(sheet, "Sheet1");
                assert_eq!(anchor, "B7");
                assert_eq!(from_grid.as_deref(), Some("region.json"));
                assert!(from_csv.is_none());
                assert!(!header);
                assert!(!clear_target);
                assert!(!dry_run);
                assert!(in_place);
                assert!(output.is_none());
                assert!(!force);
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_range_import_from_csv_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "range-import",
            "workbook.xlsx",
            "Sheet1",
            "--anchor",
            "B7",
            "--from-csv",
            "data.csv",
            "--header",
            "--in-place",
        ])
        .expect("parse range-import csv");

        match cli.command {
            Commands::RangeImport {
                from_grid,
                from_csv,
                header,
                ..
            } => {
                assert!(from_grid.is_none());
                assert_eq!(from_csv.as_deref(), Some("data.csv"));
                assert!(header);
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_inspect_cells_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "inspect-cells",
            "workbook.xlsx",
            "Sheet1",
            "A1:C10",
            "D4",
            "--include-empty",
        ])
        .expect("parse command");

        match cli.command {
            Commands::InspectCells {
                file,
                sheet,
                targets,
                include_empty,
                budget,
                ..
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(sheet, "Sheet1");
                assert_eq!(targets, vec!["A1:C10", "D4"]);
                assert!(include_empty);
                assert_eq!(budget, None);
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_sheet_page_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "sheet-page",
            "workbook.xlsx",
            "Sheet1",
            "--start-row",
            "2",
            "--page-size",
            "5",
            "--columns",
            "A,C:E",
            "--columns-by-header",
            "Name,Total",
            "--include-formulas",
            "--include-styles",
            "--include-header",
            "--format",
            "compact",
        ])
        .expect("parse command");

        match cli.command {
            Commands::SheetPage {
                file,
                sheet,
                start_row,
                page_size,
                columns,
                columns_by_header,
                include_formulas,
                include_styles,
                include_header,
                format,
                ..
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(sheet, "Sheet1");
                assert_eq!(start_row, Some(2));
                assert_eq!(page_size, Some(5));
                assert_eq!(columns, Some(vec!["A".to_string(), "C:E".to_string()]));
                assert_eq!(
                    columns_by_header,
                    Some(vec!["Name".to_string(), "Total".to_string()])
                );
                assert_eq!(include_formulas, Some(true));
                assert_eq!(include_styles, Some(true));
                assert_eq!(include_header, Some(true));
                assert!(matches!(format, SheetPageFormatArg::Compact));
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_create_workbook_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "create-workbook",
            "workbook.xlsx",
            "--sheets",
            "Inputs,Calc,Output",
            "--overwrite",
        ])
        .expect("parse create-workbook");

        match cli.command {
            Commands::CreateWorkbook {
                path,
                sheets,
                overwrite,
            } => {
                assert_eq!(path, PathBuf::from("workbook.xlsx"));
                assert_eq!(
                    sheets,
                    Some(vec![
                        "Inputs".to_string(),
                        "Calc".to_string(),
                        "Output".to_string(),
                    ])
                );
                assert!(overwrite);
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_transform_batch_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "transform-batch",
            "workbook.xlsx",
            "--ops",
            "@ops.json",
            "--output",
            "out.xlsx",
            "--force",
        ])
        .expect("parse transform-batch");

        match cli.command {
            Commands::TransformBatch {
                file,
                ops,
                dry_run,
                in_place,
                output,
                force,
                print_schema,
                formula_parse_policy,
            } => {
                assert_eq!(file, Some(PathBuf::from("workbook.xlsx")));
                assert_eq!(ops, Some("@ops.json".to_string()));
                assert!(!dry_run);
                assert!(!in_place);
                assert_eq!(output, Some(PathBuf::from("out.xlsx")));
                assert!(force);
                assert!(!print_schema);
                assert_eq!(formula_parse_policy, None);
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_style_batch_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "style-batch",
            "workbook.xlsx",
            "--ops",
            "@style.json",
            "--dry-run",
        ])
        .expect("parse style-batch");

        match cli.command {
            Commands::StyleBatch {
                file,
                ops,
                dry_run,
                in_place,
                output,
                force,
                print_schema,
            } => {
                assert_eq!(file, Some(PathBuf::from("workbook.xlsx")));
                assert_eq!(ops, Some("@style.json".to_string()));
                assert!(dry_run);
                assert!(!in_place);
                assert!(output.is_none());
                assert!(!force);
                assert!(!print_schema);
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_apply_formula_pattern_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "apply-formula-pattern",
            "workbook.xlsx",
            "--ops",
            "@formula.json",
            "--in-place",
        ])
        .expect("parse apply-formula-pattern");

        match cli.command {
            Commands::ApplyFormulaPattern {
                file,
                ops,
                dry_run,
                in_place,
                output,
                force,
                print_schema,
            } => {
                assert_eq!(file, Some(PathBuf::from("workbook.xlsx")));
                assert_eq!(ops, Some("@formula.json".to_string()));
                assert!(!dry_run);
                assert!(in_place);
                assert!(output.is_none());
                assert!(!force);
                assert!(!print_schema);
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_phase_b_batch_write_arguments() {
        let structure = Cli::try_parse_from([
            "agent-spreadsheet",
            "structure-batch",
            "workbook.xlsx",
            "--ops",
            "@structure.json",
            "--output",
            "out.xlsx",
        ])
        .expect("parse structure-batch");
        match structure.command {
            Commands::StructureBatch {
                file,
                ops,
                output,
                print_schema,
                ..
            } => {
                assert_eq!(file, Some(PathBuf::from("workbook.xlsx")));
                assert_eq!(ops, Some("@structure.json".to_string()));
                assert_eq!(output, Some(PathBuf::from("out.xlsx")));
                assert!(!print_schema);
            }
            other => panic!("unexpected command: {other:?}"),
        }

        let column = Cli::try_parse_from([
            "agent-spreadsheet",
            "column-size-batch",
            "workbook.xlsx",
            "--ops",
            "@columns.json",
            "--in-place",
        ])
        .expect("parse column-size-batch");
        match column.command {
            Commands::ColumnSizeBatch {
                ops,
                in_place,
                print_schema,
                ..
            } => {
                assert_eq!(ops, Some("@columns.json".to_string()));
                assert!(in_place);
                assert!(!print_schema);
            }
            other => panic!("unexpected command: {other:?}"),
        }

        let layout = Cli::try_parse_from([
            "agent-spreadsheet",
            "sheet-layout-batch",
            "workbook.xlsx",
            "--ops",
            "@layout.json",
            "--dry-run",
        ])
        .expect("parse sheet-layout-batch");
        match layout.command {
            Commands::SheetLayoutBatch {
                ops,
                dry_run,
                print_schema,
                ..
            } => {
                assert_eq!(ops, Some("@layout.json".to_string()));
                assert!(dry_run);
                assert!(!print_schema);
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_rules_batch_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "rules-batch",
            "workbook.xlsx",
            "--ops",
            "@rules.json",
            "--output",
            "rules.xlsx",
            "--force",
        ])
        .expect("parse rules-batch");

        match cli.command {
            Commands::RulesBatch {
                file,
                ops,
                dry_run,
                in_place,
                output,
                force,
                print_schema,
                formula_parse_policy,
            } => {
                assert_eq!(file, Some(PathBuf::from("workbook.xlsx")));
                assert_eq!(ops, Some("@rules.json".to_string()));
                assert!(!dry_run);
                assert!(!in_place);
                assert_eq!(output, Some(PathBuf::from("rules.xlsx")));
                assert!(force);
                assert!(!print_schema);
                assert!(formula_parse_policy.is_none());
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_global_schema_and_example_commands() {
        let transform = Cli::try_parse_from(["asp", "schema", "transform-batch"])
            .expect("parse schema transform-batch");
        match transform.command {
            Commands::Schema {
                command: DiscoverabilityCommands::TransformBatch,
            } => {}
            other => panic!("unexpected command: {other:?}"),
        }

        let style = Cli::try_parse_from(["asp", "example", "style-batch"])
            .expect("parse example style-batch");
        match style.command {
            Commands::Example {
                command: DiscoverabilityCommands::StyleBatch,
            } => {}
            other => panic!("unexpected command: {other:?}"),
        }

        let session_schema =
            Cli::try_parse_from(["asp", "schema", "session-op", "transform.write_matrix"])
                .expect("parse schema session-op");
        match session_schema.command {
            Commands::Schema {
                command: DiscoverabilityCommands::SessionOp { kind },
            } => {
                assert_eq!(kind, "transform.write_matrix");
            }
            other => panic!("unexpected command: {other:?}"),
        }

        let session_example =
            Cli::try_parse_from(["asp", "example", "session-op", "structure.insert_rows"])
                .expect("parse example session-op");
        match session_example.command {
            Commands::Example {
                command: DiscoverabilityCommands::SessionOp { kind },
            } => {
                assert_eq!(kind, "structure.insert_rows");
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_named_ranges_and_scan_volatiles_arguments() {
        let named = Cli::try_parse_from([
            "agent-spreadsheet",
            "named-ranges",
            "workbook.xlsx",
            "--sheet",
            "Sheet1",
            "--name-prefix",
            "Sales",
        ])
        .expect("parse named-ranges");

        match named.command {
            Commands::NamedRanges {
                file,
                sheet,
                name_prefix,
                ..
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(sheet.as_deref(), Some("Sheet1"));
                assert_eq!(name_prefix.as_deref(), Some("Sales"));
            }
            other => panic!("unexpected command: {other:?}"),
        }

        let volatiles = Cli::try_parse_from([
            "agent-spreadsheet",
            "scan-volatiles",
            "workbook.xlsx",
            "--sheet",
            "Sheet1",
            "--limit",
            "10",
            "--offset",
            "5",
        ])
        .expect("parse scan-volatiles");

        match volatiles.command {
            Commands::ScanVolatiles {
                file,
                sheet,
                limit,
                offset,
                formula_parse_policy,
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(sheet.as_deref(), Some("Sheet1"));
                assert_eq!(limit, Some(10));
                assert_eq!(offset, Some(5));
                assert!(formula_parse_policy.is_none());
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_find_value_label_direction_arguments() {
        let find = Cli::try_parse_from([
            "agent-spreadsheet",
            "find-value",
            "workbook.xlsx",
            "Amount",
            "--sheet",
            "Sheet1",
            "--mode",
            "label",
            "--label-direction",
            "below",
        ])
        .expect("parse find-value");

        match find.command {
            Commands::FindValue {
                file,
                query,
                sheet,
                mode,
                label_direction,
                ..
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(query, "Amount");
                assert_eq!(sheet.as_deref(), Some("Sheet1"));
                assert!(matches!(mode, Some(FindValueMode::Label)));
                assert!(matches!(label_direction, Some(LabelDirectionArg::Below)));
            }
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_find_formula_and_sheet_statistics_arguments() {
        let find = Cli::try_parse_from([
            "agent-spreadsheet",
            "find-formula",
            "workbook.xlsx",
            "SUM(",
            "--sheet",
            "Sheet1",
            "--limit",
            "25",
            "--offset",
            "50",
        ])
        .expect("parse find-formula");

        match find.command {
            Commands::FindFormula {
                file,
                query,
                sheet,
                limit,
                offset,
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(query, "SUM(");
                assert_eq!(sheet.as_deref(), Some("Sheet1"));
                assert_eq!(limit, Some(25));
                assert_eq!(offset, Some(50));
            }
            other => panic!("unexpected command: {other:?}"),
        }

        let stats = Cli::try_parse_from([
            "agent-spreadsheet",
            "sheet-statistics",
            "workbook.xlsx",
            "Summary",
        ])
        .expect("parse sheet-statistics");

        match stats.command {
            Commands::SheetStatistics { file, sheet } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(sheet, "Summary");
            }
            other => panic!("unexpected command: {other:?}"),
        }

        assert!(
            Cli::try_parse_from(["agent-spreadsheet", "find-formula", "workbook.xlsx"]).is_err(),
            "missing QUERY should fail clap parsing"
        );
    }

    #[test]
    fn parses_sheet_page_all_required_formats() {
        for (raw, expected) in [
            ("full", SheetPageFormatArg::Full),
            ("compact", SheetPageFormatArg::Compact),
            ("values_only", SheetPageFormatArg::ValuesOnly),
        ] {
            let cli = Cli::try_parse_from([
                "agent-spreadsheet",
                "sheet-page",
                "workbook.xlsx",
                "Sheet1",
                "--format",
                raw,
            ])
            .expect("parse format value");

            match cli.command {
                Commands::SheetPage { format, .. } => {
                    assert!(
                        matches!(
                            (format, expected),
                            (SheetPageFormatArg::Full, SheetPageFormatArg::Full)
                                | (SheetPageFormatArg::Compact, SheetPageFormatArg::Compact)
                                | (
                                    SheetPageFormatArg::ValuesOnly,
                                    SheetPageFormatArg::ValuesOnly
                                )
                        ),
                        "format mismatch for {raw}"
                    );
                }
                other => panic!("unexpected command: {other:?}"),
            }
        }
    }

    #[test]
    fn normalizes_legacy_global_format_for_non_sheet_page_commands() {
        let normalized = normalize_legacy_global_format_argv(
            [
                "agent-spreadsheet",
                "list-sheets",
                "workbook.xlsx",
                "--format",
                "json",
            ]
            .into_iter()
            .map(OsString::from)
            .collect(),
        );

        let tokens = normalized
            .iter()
            .map(|arg| arg.to_string_lossy().into_owned())
            .collect::<Vec<_>>();

        assert_eq!(
            tokens,
            vec![
                "agent-spreadsheet",
                "list-sheets",
                "workbook.xlsx",
                "--output-format",
                "json"
            ]
        );
    }

    #[test]
    fn preserves_sheet_page_local_format_flag() {
        let normalized = normalize_legacy_global_format_argv(
            [
                "agent-spreadsheet",
                "sheet-page",
                "workbook.xlsx",
                "Sheet1",
                "--format",
                "compact",
            ]
            .into_iter()
            .map(OsString::from)
            .collect(),
        );

        let tokens = normalized
            .iter()
            .map(|arg| arg.to_string_lossy().into_owned())
            .collect::<Vec<_>>();

        assert_eq!(
            tokens,
            vec![
                "agent-spreadsheet",
                "sheet-page",
                "workbook.xlsx",
                "Sheet1",
                "--format",
                "compact"
            ]
        );
    }

    #[test]
    fn preserves_range_values_local_format_flag() {
        let normalized = normalize_legacy_global_format_argv(
            [
                "agent-spreadsheet",
                "range-values",
                "workbook.xlsx",
                "Sheet1",
                "A1:B2",
                "--format",
                "json",
            ]
            .into_iter()
            .map(OsString::from)
            .collect(),
        );

        let tokens = normalized
            .iter()
            .map(|arg| arg.to_string_lossy().into_owned())
            .collect::<Vec<_>>();

        assert_eq!(
            tokens,
            vec![
                "agent-spreadsheet",
                "range-values",
                "workbook.xlsx",
                "Sheet1",
                "A1:B2",
                "--format",
                "json"
            ]
        );
    }

    #[test]
    fn parses_sheetport_manifest_validate_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "sheetport",
            "manifest",
            "validate",
            "manifest.yaml",
        ])
        .expect("parse sheetport manifest validate");

        match cli.command {
            Commands::Sheetport { command } => match command {
                SheetportCommands::Manifest(SheetportManifestCommands::Validate { manifest }) => {
                    assert_eq!(manifest, PathBuf::from("manifest.yaml"));
                }
                other => panic!("unexpected sheetport command: {other:?}"),
            },
            other => panic!("unexpected command: {other:?}"),
        }
    }

    #[test]
    fn parses_sheetport_run_arguments() {
        let cli = Cli::try_parse_from([
            "agent-spreadsheet",
            "sheetport",
            "run",
            "workbook.xlsx",
            "manifest.yaml",
            "--inputs",
            "@inputs.json",
            "--rng-seed",
            "42",
            "--freeze-volatile",
        ])
        .expect("parse sheetport run");

        match cli.command {
            Commands::Sheetport { command } => match command {
                SheetportCommands::Run {
                    file,
                    manifest,
                    inputs,
                    rng_seed,
                    freeze_volatile,
                } => {
                    assert_eq!(file, PathBuf::from("workbook.xlsx"));
                    assert_eq!(manifest, PathBuf::from("manifest.yaml"));
                    assert_eq!(inputs.as_deref(), Some("@inputs.json"));
                    assert_eq!(rng_seed, Some(42));
                    assert!(freeze_volatile);
                }
                other => panic!("unexpected sheetport command: {other:?}"),
            },
            other => panic!("unexpected command: {other:?}"),
        }
    }
}
