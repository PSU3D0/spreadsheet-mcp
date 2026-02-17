pub mod commands;
pub mod errors;
pub mod output;

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

#[derive(Debug, Parser)]
#[command(
    name = "agent-spreadsheet",
    version,
    about = "Stateless spreadsheet CLI for reads, edits, and diffs",
    long_about = "Stateless spreadsheet CLI for AI and automation workflows.\n\nCommon workflows:\n  • Inspect a workbook: list-sheets → sheet-overview → table-profile\n  • Deterministic pagination loops: sheet-page (--format + next_start_row) and read-table (--limit/--offset + next_offset)\n  • Find labels or values: find-value --mode label|value\n  • Stateless batch writes: transform/style/formula/structure/column/layout/rules via --ops @ops.json + one mode (--dry-run|--in-place|--output)\n  • Copy → edit → recalculate → diff for safe what-if changes\n\nTip: global --output-format csv is currently unsupported and returns an error. Use --output-format json, or command-level CSV options such as read-table --table-format csv."
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
        help = "Output shape (canonical keeps full schema; compact applies command-specific projections: range-values single-range flattening, read-table/sheet-page branch preservation, and formula-trace layer highlight omission while preserving continuation fields)"
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
    #[command(
        about = "Read one sheet page with deterministic continuation",
        after_long_help = "Examples:\n  agent-spreadsheet sheet-page data.xlsx Sheet1 --format compact --page-size 200\n  agent-spreadsheet sheet-page data.xlsx Sheet1 --format compact --page-size 200 --start-row 201\n  agent-spreadsheet sheet-page data.xlsx Sheet1 --format full --columns A,C:E --include-styles\n\nPagination loop:\n  1) Run without --start-row.\n  2) If next_start_row is present, pass it to --start-row for the next request.\n  3) Stop when next_start_row is omitted."
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
    },
    #[command(
        about = "Find formulas containing a text query with pagination",
        after_long_help = "Examples:\n  agent-spreadsheet find-formula data.xlsx SUM(\n  agent-spreadsheet find-formula data.xlsx VLOOKUP --sheet \"Q1 Actuals\" --limit 25 --offset 50"
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
    },
    #[command(
        about = "Trace formula precedents or dependents from one origin cell",
        after_long_help = "Examples:\n  agent-spreadsheet formula-trace data.xlsx Sheet1 C2 precedents --depth 2\n  agent-spreadsheet formula-trace data.xlsx Sheet1 C2 dependents --page-size 25\n  agent-spreadsheet formula-trace data.xlsx Sheet1 C2 precedents --cursor-depth 1 --cursor-offset 25\n\nContinuation:\n  Reuse next_cursor.depth/next_cursor.offset as --cursor-depth/--cursor-offset to continue paged traces."
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
    #[command(
        about = "Apply stateless transform operations from an @ops payload",
        after_long_help = "Examples:\n  agent-spreadsheet transform-batch workbook.xlsx --ops @ops.json --dry-run\n  agent-spreadsheet transform-batch workbook.xlsx --ops @ops.json --in-place\n  agent-spreadsheet transform-batch workbook.xlsx --ops @ops.json --output transformed.xlsx --force\n\nMode selection:\n  Choose exactly one of --dry-run, --in-place, or --output <PATH>."
    )]
    TransformBatch {
        #[arg(value_name = "FILE", help = "Workbook path to transform")]
        file: PathBuf,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)"
        )]
        ops: String,
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
    },
    #[command(
        about = "Apply stateless style operations from an @ops payload",
        after_long_help = "Examples:\n  agent-spreadsheet style-batch workbook.xlsx --ops @style_ops.json --dry-run\n  agent-spreadsheet style-batch workbook.xlsx --ops @style_ops.json --output styled.xlsx --force"
    )]
    StyleBatch {
        #[arg(value_name = "FILE", help = "Workbook path to style")]
        file: PathBuf,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)"
        )]
        ops: String,
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
    },
    #[command(
        about = "Apply stateless formula pattern operations from an @ops payload",
        after_long_help = "Examples:\n  agent-spreadsheet apply-formula-pattern workbook.xlsx --ops @formula_ops.json --in-place\n  agent-spreadsheet apply-formula-pattern workbook.xlsx --ops @formula_ops.json --dry-run\n\nCache note:\n  Updated formula cells clear cached results. Run recalculate to refresh computed values."
    )]
    ApplyFormulaPattern {
        #[arg(value_name = "FILE", help = "Workbook path to update")]
        file: PathBuf,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)"
        )]
        ops: String,
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
    },
    #[command(
        about = "Apply stateless structure operations from an @ops payload",
        after_long_help = "Examples:\n  agent-spreadsheet structure-batch workbook.xlsx --ops @structure_ops.json --dry-run\n  agent-spreadsheet structure-batch workbook.xlsx --ops @structure_ops.json --output structured.xlsx"
    )]
    StructureBatch {
        #[arg(value_name = "FILE", help = "Workbook path to update")]
        file: PathBuf,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)"
        )]
        ops: String,
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
    },
    #[command(
        about = "Apply stateless column sizing operations from an @ops payload",
        after_long_help = "Examples:\n  agent-spreadsheet column-size-batch workbook.xlsx --ops @column_size_ops.json --in-place\n  agent-spreadsheet column-size-batch workbook.xlsx --ops @column_size_ops.json --output columns.xlsx"
    )]
    ColumnSizeBatch {
        #[arg(value_name = "FILE", help = "Workbook path to update")]
        file: PathBuf,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)"
        )]
        ops: String,
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
    },
    #[command(
        about = "Apply stateless sheet layout operations from an @ops payload",
        after_long_help = "Examples:\n  agent-spreadsheet sheet-layout-batch workbook.xlsx --ops @layout_ops.json --dry-run\n  agent-spreadsheet sheet-layout-batch workbook.xlsx --ops @layout_ops.json --in-place"
    )]
    SheetLayoutBatch {
        #[arg(value_name = "FILE", help = "Workbook path to update")]
        file: PathBuf,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)"
        )]
        ops: String,
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
    },
    #[command(
        about = "Apply stateless data validation and conditional format operations from an @ops payload",
        after_long_help = "Examples:\n  agent-spreadsheet rules-batch workbook.xlsx --ops @rules_ops.json --dry-run\n  agent-spreadsheet rules-batch workbook.xlsx --ops @rules_ops.json --output ruled.xlsx --force"
    )]
    RulesBatch {
        #[arg(value_name = "FILE", help = "Workbook path to update")]
        file: PathBuf,
        #[arg(
            long,
            value_name = "OPS_REF",
            help = "Ops payload file reference (@path)"
        )]
        ops: String,
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
        } => {
            commands::read::sheet_page(
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
        } => {
            commands::read::read_table(
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
            )
            .await
        }
        Commands::FindValue {
            file,
            query,
            sheet,
            mode,
            label_direction,
        } => commands::read::find_value(file, query, sheet, mode, label_direction).await,
        Commands::NamedRanges {
            file,
            sheet,
            name_prefix,
        } => commands::read::named_ranges(file, sheet, name_prefix).await,
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
        } => commands::read::scan_volatiles(file, sheet, limit, offset).await,
        Commands::SheetStatistics { file, sheet } => {
            commands::read::sheet_statistics(file, sheet).await
        }
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
            depth,
            page_size,
            cursor_depth,
            cursor_offset,
        } => {
            commands::read::formula_trace(
                file,
                sheet,
                cell,
                direction,
                depth,
                page_size,
                cursor_depth,
                cursor_offset,
            )
            .await
        }
        Commands::Describe { file } => commands::read::describe(file).await,
        Commands::TableProfile { file, sheet } => commands::read::table_profile(file, sheet).await,
        Commands::Copy { source, dest } => commands::write::copy(source, dest).await,
        Commands::Edit { file, sheet, edits } => commands::write::edit(file, sheet, edits).await,
        Commands::TransformBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
        } => commands::write::transform_batch(file, ops, dry_run, in_place, output, force).await,
        Commands::StyleBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
        } => commands::write::style_batch(file, ops, dry_run, in_place, output, force).await,
        Commands::ApplyFormulaPattern {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
        } => {
            commands::write::apply_formula_pattern(file, ops, dry_run, in_place, output, force)
                .await
        }
        Commands::StructureBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
        } => commands::write::structure_batch(file, ops, dry_run, in_place, output, force).await,
        Commands::ColumnSizeBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
        } => commands::write::column_size_batch(file, ops, dry_run, in_place, output, force).await,
        Commands::SheetLayoutBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
        } => commands::write::sheet_layout_batch(file, ops, dry_run, in_place, output, force).await,
        Commands::RulesBatch {
            file,
            ops,
            dry_run,
            in_place,
            output,
            force,
        } => commands::write::rules_batch(file, ops, dry_run, in_place, output, force).await,
        Commands::Recalculate { file } => commands::recalc::recalculate(file).await,
        Commands::Diff { original, modified } => commands::diff::diff(original, modified).await,
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
    let preserve_sheet_page_format = first_subcommand_name == "sheet-page";

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

        if can_rewrite_here {
            if let Some(value) = token.strip_prefix("--format=") {
                if is_legacy_output_format(value) {
                    normalized.push(OsString::from(format!("--output-format={value}")));
                    index += 1;
                    continue;
                }
            }
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

    match run_command(command).await {
        Ok(payload) => {
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
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(ops, "@ops.json");
                assert!(!dry_run);
                assert!(!in_place);
                assert_eq!(output, Some(PathBuf::from("out.xlsx")));
                assert!(force);
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
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(ops, "@style.json");
                assert!(dry_run);
                assert!(!in_place);
                assert!(output.is_none());
                assert!(!force);
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
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(ops, "@formula.json");
                assert!(!dry_run);
                assert!(in_place);
                assert!(output.is_none());
                assert!(!force);
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
                file, ops, output, ..
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(ops, "@structure.json");
                assert_eq!(output, Some(PathBuf::from("out.xlsx")));
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
            Commands::ColumnSizeBatch { ops, in_place, .. } => {
                assert_eq!(ops, "@columns.json");
                assert!(in_place);
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
            Commands::SheetLayoutBatch { ops, dry_run, .. } => {
                assert_eq!(ops, "@layout.json");
                assert!(dry_run);
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
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(ops, "@rules.json");
                assert!(!dry_run);
                assert!(!in_place);
                assert_eq!(output, Some(PathBuf::from("rules.xlsx")));
                assert!(force);
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
            } => {
                assert_eq!(file, PathBuf::from("workbook.xlsx"));
                assert_eq!(sheet.as_deref(), Some("Sheet1"));
                assert_eq!(limit, Some(10));
                assert_eq!(offset, Some(5));
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
}
