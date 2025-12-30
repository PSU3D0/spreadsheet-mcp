use crate::caps::BackendCaps;
use schemars::JsonSchema;
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize, JsonSchema, Default)]
#[serde(transparent)]
pub struct WorkbookId(pub String);

impl WorkbookId {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct WorkbookDescriptor {
    pub workbook_id: WorkbookId,
    pub short_id: String,
    pub slug: String,
    pub folder: Option<String>,
    pub path: String,
    pub bytes: u64,
    pub last_modified: Option<String>,
    pub caps: BackendCaps,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct WorkbookListResponse {
    pub workbooks: Vec<WorkbookDescriptor>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct WorkbookDescription {
    pub workbook_id: WorkbookId,
    pub short_id: String,
    pub slug: String,
    pub path: String,
    pub bytes: u64,
    pub sheet_count: usize,
    pub defined_names: usize,
    pub tables: usize,
    pub macros_present: bool,
    pub last_modified: Option<String>,
    pub caps: BackendCaps,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct WorkbookSummaryResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub slug: String,
    pub sheet_count: usize,
    pub total_cells: u64,
    pub total_formulas: u64,
    pub breakdown: WorkbookBreakdown,
    pub region_counts: RegionCountSummary,
    pub region_counts_truncated: bool,
    pub key_named_ranges: Vec<NamedRangeDescriptor>,
    pub suggested_entry_points: Vec<EntryPoint>,
    pub entry_points_truncated: bool,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct WorkbookBreakdown {
    pub data_sheets: u32,
    pub calculator_sheets: u32,
    pub parameter_sheets: u32,
    pub metadata_sheets: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct RegionCountSummary {
    pub data: u32,
    pub parameters: u32,
    pub outputs: u32,
    pub calculator: u32,
    pub metadata: u32,
    pub other: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct EntryPoint {
    pub sheet_name: String,
    pub region_id: Option<u32>,
    pub bounds: Option<String>,
    pub rationale: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct SheetSummary {
    pub name: String,
    pub visible: bool,
    pub row_count: u32,
    pub column_count: u32,
    pub non_empty_cells: u32,
    pub formula_cells: u32,
    pub cached_values: u32,
    pub classification: SheetClassification,
    pub style_tags: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum SheetClassification {
    Data,
    Calculator,
    Mixed,
    Metadata,
    Empty,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct SheetListResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub sheets: Vec<SheetSummary>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct SheetOverviewResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub sheet_name: String,
    pub narrative: String,
    pub regions: Vec<SheetRegion>,
    pub detected_regions: Vec<DetectedRegion>,
    pub detected_region_count: u32,
    pub detected_regions_truncated: bool,
    pub key_ranges: Vec<String>,
    pub formula_ratio: f32,
    pub notable_features: Vec<String>,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct SheetRegion {
    pub kind: RegionKind,
    pub address: String,
    pub description: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub enum RegionKind {
    #[serde(rename = "likely_table")]
    Table,
    #[serde(rename = "likely_data")]
    Data,
    #[serde(rename = "likely_parameters")]
    Parameters,
    #[serde(rename = "likely_outputs")]
    Outputs,
    #[serde(rename = "likely_calculator")]
    Calculator,
    #[serde(rename = "likely_metadata")]
    Metadata,
    #[serde(rename = "likely_styles")]
    Styles,
    #[serde(rename = "likely_comments")]
    Comments,
    #[serde(rename = "unknown")]
    Other,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct DetectedRegion {
    pub id: u32,
    pub bounds: String,
    pub header_row: Option<u32>,
    pub headers: Vec<String>,
    pub header_count: u32,
    pub headers_truncated: bool,
    pub row_count: u32,
    pub classification: RegionKind,
    pub region_kind: Option<RegionKind>,
    pub confidence: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct SheetPageResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub sheet_name: String,
    pub rows: Vec<RowSnapshot>,
    pub has_more: bool,
    pub next_start_row: Option<u32>,
    pub header_row: Option<RowSnapshot>,
    pub compact: Option<SheetPageCompact>,
    pub values_only: Option<SheetPageValues>,
    pub format: SheetPageFormat,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct RowSnapshot {
    pub row_index: u32,
    pub cells: Vec<CellSnapshot>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct CellSnapshot {
    pub address: String,
    pub value: Option<CellValue>,
    pub formula: Option<String>,
    pub cached_value: Option<CellValue>,
    pub number_format: Option<String>,
    pub style_tags: Vec<String>,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "kind", content = "value")]
pub enum CellValue {
    Text(String),
    Number(f64),
    Bool(bool),
    Error(String),
    Date(String),
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
#[derive(Default)]
pub enum SheetPageFormat {
    #[default]
    Full,
    Compact,
    ValuesOnly,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct SheetPageCompact {
    pub headers: Vec<String>,
    pub header_row: Vec<Option<CellValue>>,
    pub rows: Vec<Vec<Option<CellValue>>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct SheetPageValues {
    pub rows: Vec<Vec<Option<CellValue>>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct SheetStatisticsResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub sheet_name: String,
    pub row_count: u32,
    pub column_count: u32,
    pub density: f32,
    pub numeric_columns: Vec<ColumnSummary>,
    pub text_columns: Vec<ColumnSummary>,
    pub null_counts: BTreeMap<String, u32>,
    pub duplicate_warnings: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct ColumnSummary {
    pub header: Option<String>,
    pub column: String,
    pub samples: Vec<CellValue>,
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub mean: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct SheetFormulaMapResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub sheet_name: String,
    pub groups: Vec<FormulaGroup>,
    pub truncated: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct FormulaGroup {
    pub fingerprint: String,
    pub addresses: Vec<String>,
    pub formula: String,
    pub is_array: bool,
    pub is_shared: bool,
    pub is_volatile: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct FormulaTraceResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub sheet_name: String,
    pub origin: String,
    pub direction: TraceDirection,
    pub layers: Vec<TraceLayer>,
    pub next_cursor: Option<TraceCursor>,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct FormulaTraceEdge {
    pub from: String,
    pub to: String,
    pub formula: Option<String>,
    pub note: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct TraceLayer {
    pub depth: u32,
    pub summary: TraceLayerSummary,
    pub highlights: TraceLayerHighlights,
    pub edges: Vec<FormulaTraceEdge>,
    pub has_more: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct TraceLayerSummary {
    pub total_nodes: usize,
    pub formula_nodes: usize,
    pub value_nodes: usize,
    pub blank_nodes: usize,
    pub external_nodes: usize,
    pub unique_formula_groups: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct TraceLayerHighlights {
    pub top_ranges: Vec<TraceRangeHighlight>,
    pub top_formula_groups: Vec<TraceFormulaGroupHighlight>,
    pub notable_cells: Vec<TraceCellHighlight>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct TraceRangeHighlight {
    pub start: String,
    pub end: String,
    pub count: usize,
    pub literals: usize,
    pub formulas: usize,
    pub blanks: usize,
    pub sample_values: Vec<CellValue>,
    pub sample_formulas: Vec<String>,
    pub sample_addresses: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct TraceFormulaGroupHighlight {
    pub fingerprint: String,
    pub formula: String,
    pub count: usize,
    pub sample_addresses: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct TraceCellHighlight {
    pub address: String,
    pub kind: TraceCellKind,
    pub value: Option<CellValue>,
    pub formula: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, PartialEq, Eq, Hash)]
#[serde(rename_all = "snake_case")]
pub enum TraceCellKind {
    Formula,
    Literal,
    Blank,
    External,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct TraceCursor {
    pub depth: u32,
    pub offset: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum TraceDirection {
    Precedents,
    Dependents,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, PartialEq, Eq)]
pub struct NamedRangeDescriptor {
    pub name: String,
    pub scope: Option<String>,
    pub refers_to: String,
    pub kind: NamedItemKind,
    pub sheet_name: Option<String>,
    pub comment: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum NamedItemKind {
    NamedRange,
    Table,
    Formula,
    Unknown,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct NamedRangesResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub items: Vec<NamedRangeDescriptor>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct FindFormulaMatch {
    pub address: String,
    pub sheet_name: String,
    pub formula: String,
    pub cached_value: Option<CellValue>,
    pub context: Vec<RowSnapshot>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct FindFormulaResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub matches: Vec<FindFormulaMatch>,
    pub truncated: bool,
    pub next_offset: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct VolatileScanEntry {
    pub address: String,
    pub sheet_name: String,
    pub function: String,
    pub note: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct VolatileScanResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub items: Vec<VolatileScanEntry>,
    pub truncated: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct StyleDescriptor {
    pub font: Option<FontDescriptor>,
    pub fill: Option<FillDescriptor>,
    pub borders: Option<BordersDescriptor>,
    pub alignment: Option<AlignmentDescriptor>,
    pub number_format: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct FontDescriptor {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<String>,
    pub strikethrough: Option<bool>,
    pub color: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum FillDescriptor {
    Pattern(PatternFillDescriptor),
    Gradient(GradientFillDescriptor),
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct PatternFillDescriptor {
    pub pattern_type: Option<String>,
    pub foreground_color: Option<String>,
    pub background_color: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct GradientFillDescriptor {
    pub degree: Option<f64>,
    pub stops: Vec<GradientStopDescriptor>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct GradientStopDescriptor {
    pub position: f64,
    pub color: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct BordersDescriptor {
    pub left: Option<BorderSideDescriptor>,
    pub right: Option<BorderSideDescriptor>,
    pub top: Option<BorderSideDescriptor>,
    pub bottom: Option<BorderSideDescriptor>,
    pub diagonal: Option<BorderSideDescriptor>,
    pub vertical: Option<BorderSideDescriptor>,
    pub horizontal: Option<BorderSideDescriptor>,
    pub diagonal_up: Option<bool>,
    pub diagonal_down: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct BorderSideDescriptor {
    pub style: Option<String>,
    pub color: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct AlignmentDescriptor {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: Option<bool>,
    pub text_rotation: Option<u32>,
}

// Patch variants for write tools (Phase 2+). Double-option fields distinguish:
// - missing field => no change (merge mode)
// - null => clear to default
// - value => set/merge that value
#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct StylePatch {
    #[serde(default)]
    pub font: Option<Option<FontPatch>>,
    #[serde(default)]
    pub fill: Option<Option<FillPatch>>,
    #[serde(default)]
    pub borders: Option<Option<BordersPatch>>,
    #[serde(default)]
    pub alignment: Option<Option<AlignmentPatch>>,
    #[serde(default)]
    pub number_format: Option<Option<String>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct FontPatch {
    #[serde(default)]
    pub name: Option<Option<String>>,
    #[serde(default)]
    pub size: Option<Option<f64>>,
    #[serde(default)]
    pub bold: Option<Option<bool>>,
    #[serde(default)]
    pub italic: Option<Option<bool>>,
    #[serde(default)]
    pub underline: Option<Option<String>>,
    #[serde(default)]
    pub strikethrough: Option<Option<bool>>,
    #[serde(default)]
    pub color: Option<Option<String>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum FillPatch {
    Pattern(PatternFillPatch),
    Gradient(GradientFillPatch),
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct PatternFillPatch {
    #[serde(default)]
    pub pattern_type: Option<Option<String>>,
    #[serde(default)]
    pub foreground_color: Option<Option<String>>,
    #[serde(default)]
    pub background_color: Option<Option<String>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct GradientFillPatch {
    #[serde(default)]
    pub degree: Option<Option<f64>>,
    #[serde(default)]
    pub stops: Option<Vec<GradientStopPatch>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct GradientStopPatch {
    pub position: f64,
    pub color: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct BordersPatch {
    #[serde(default)]
    pub left: Option<Option<BorderSidePatch>>,
    #[serde(default)]
    pub right: Option<Option<BorderSidePatch>>,
    #[serde(default)]
    pub top: Option<Option<BorderSidePatch>>,
    #[serde(default)]
    pub bottom: Option<Option<BorderSidePatch>>,
    #[serde(default)]
    pub diagonal: Option<Option<BorderSidePatch>>,
    #[serde(default)]
    pub vertical: Option<Option<BorderSidePatch>>,
    #[serde(default)]
    pub horizontal: Option<Option<BorderSidePatch>>,
    #[serde(default)]
    pub diagonal_up: Option<Option<bool>>,
    #[serde(default)]
    pub diagonal_down: Option<Option<bool>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct BorderSidePatch {
    #[serde(default)]
    pub style: Option<Option<String>>,
    #[serde(default)]
    pub color: Option<Option<String>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct AlignmentPatch {
    #[serde(default)]
    pub horizontal: Option<Option<String>>,
    #[serde(default)]
    pub vertical: Option<Option<String>>,
    #[serde(default)]
    pub wrap_text: Option<Option<bool>>,
    #[serde(default)]
    pub text_rotation: Option<Option<u32>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct SheetStylesResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub sheet_name: String,
    pub styles: Vec<StyleSummary>,
    pub conditional_rules: Vec<String>,
    pub total_styles: u32,
    pub styles_truncated: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct StyleSummary {
    pub style_id: String,
    pub occurrences: u32,
    pub tags: Vec<String>,
    pub example_cells: Vec<String>,
    pub descriptor: Option<StyleDescriptor>,
    pub cell_ranges: Vec<String>,
    pub ranges_truncated: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct WorkbookStyleSummaryResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub theme: Option<ThemeSummary>,
    pub inferred_default_style_id: Option<String>,
    pub inferred_default_font: Option<FontDescriptor>,
    pub styles: Vec<WorkbookStyleUsage>,
    pub total_styles: u32,
    pub styles_truncated: bool,
    pub conditional_formats: Vec<ConditionalFormatSummary>,
    pub conditional_formats_truncated: bool,
    pub scan_truncated: bool,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct WorkbookStyleUsage {
    pub style_id: String,
    pub occurrences: u32,
    pub tags: Vec<String>,
    pub example_cells: Vec<String>,
    pub descriptor: Option<StyleDescriptor>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct ThemeSummary {
    pub name: Option<String>,
    pub colors: BTreeMap<String, String>,
    pub font_scheme: ThemeFontSchemeSummary,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
pub struct ThemeFontSchemeSummary {
    pub major_latin: Option<String>,
    pub major_east_asian: Option<String>,
    pub major_complex_script: Option<String>,
    pub minor_latin: Option<String>,
    pub minor_east_asian: Option<String>,
    pub minor_complex_script: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct ConditionalFormatSummary {
    pub sheet_name: String,
    pub range: String,
    pub rule_types: Vec<String>,
    pub rule_count: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct ManifestStubResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub slug: String,
    pub sheets: Vec<ManifestSheetStub>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct ManifestSheetStub {
    pub sheet_name: String,
    pub classification: SheetClassification,
    pub candidate_expectations: Vec<String>,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema, Default)]
#[serde(rename_all = "snake_case")]
pub enum FindMode {
    #[default]
    Value,
    Label,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum LabelDirection {
    Right,
    Below,
    Any,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct FindValueMatch {
    pub address: String,
    pub sheet_name: String,
    pub value: Option<CellValue>,
    pub row_context: Option<RowContext>,
    pub neighbors: Option<NeighborValues>,
    pub label_hit: Option<LabelHit>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct RowContext {
    pub headers: Vec<String>,
    pub values: Vec<Option<CellValue>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct NeighborValues {
    pub left: Option<CellValue>,
    pub right: Option<CellValue>,
    pub up: Option<CellValue>,
    pub down: Option<CellValue>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct LabelHit {
    pub label_address: String,
    pub label: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct FindValueResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub matches: Vec<FindValueMatch>,
    pub truncated: bool,
}

pub type TableRow = BTreeMap<String, Option<CellValue>>;

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct ReadTableResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub sheet_name: String,
    pub table_name: Option<String>,
    pub headers: Vec<String>,
    pub rows: Vec<TableRow>,
    pub total_rows: u32,
    pub has_more: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct ColumnTypeSummary {
    pub name: String,
    pub inferred_type: String,
    pub nulls: u32,
    pub distinct: u32,
    pub top_values: Vec<String>,
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub mean: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct TableProfileResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub sheet_name: String,
    pub table_name: Option<String>,
    pub headers: Vec<String>,
    pub column_types: Vec<ColumnTypeSummary>,
    pub row_count: u32,
    pub samples: Vec<TableRow>,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct RangeValuesResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub sheet_name: String,
    pub values: Vec<RangeValuesEntry>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct RangeValuesEntry {
    pub range: String,
    pub rows: Vec<Vec<Option<CellValue>>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct CloseWorkbookResponse {
    pub workbook_id: WorkbookId,
    pub message: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct VbaProjectSummaryResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub has_vba: bool,
    pub code_page: Option<u16>,
    pub sys_kind: Option<String>,
    pub modules: Vec<VbaModuleDescriptor>,
    pub modules_truncated: bool,
    pub references: Vec<VbaReferenceDescriptor>,
    pub references_truncated: bool,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct VbaModuleDescriptor {
    pub name: String,
    pub stream_name: String,
    pub doc_string: String,
    pub text_offset: u64,
    pub help_context: u32,
    pub module_type: String,
    pub read_only: bool,
    pub private: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct VbaReferenceDescriptor {
    pub kind: String,
    pub debug: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct VbaModuleSourceResponse {
    pub workbook_id: WorkbookId,
    pub workbook_short_id: String,
    pub module_name: String,
    pub offset_lines: u32,
    pub limit_lines: u32,
    pub total_lines: u32,
    pub truncated: bool,
    pub source: String,
}
