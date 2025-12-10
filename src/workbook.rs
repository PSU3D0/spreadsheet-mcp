use crate::analysis::{
    classification,
    formula::{FormulaAtlas, FormulaGraph},
    style,
};
use crate::caps::BackendCaps;
use crate::config::ServerConfig;
use crate::model::{
    NamedItemKind, NamedRangeDescriptor, SheetClassification, SheetOverviewResponse, SheetSummary,
    WorkbookDescription, WorkbookId, WorkbookListResponse,
};
use crate::tools::filters::WorkbookFilter;
use crate::utils::{
    hash_path_metadata, make_short_workbook_id, path_to_forward_slashes, system_time_to_rfc3339,
};
use anyhow::{Context, Result, anyhow};
use chrono::{DateTime, Utc};
use parking_lot::RwLock;
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use umya_spreadsheet::reader::xlsx;
use umya_spreadsheet::{DefinedName, Spreadsheet, Worksheet};

pub struct WorkbookContext {
    pub id: WorkbookId,
    pub short_id: String,
    pub slug: String,
    pub path: PathBuf,
    pub caps: BackendCaps,
    pub bytes: u64,
    pub last_modified: Option<DateTime<Utc>>,
    spreadsheet: Arc<RwLock<Spreadsheet>>,
    sheet_cache: RwLock<HashMap<String, Arc<SheetCacheEntry>>>,
    formula_atlas: Arc<FormulaAtlas>,
}

pub struct SheetCacheEntry {
    pub metrics: SheetMetrics,
    pub style_tags: Vec<String>,
    pub named_ranges: Vec<NamedRangeDescriptor>,
    pub detected_regions: Vec<crate::model::DetectedRegion>,
}

#[derive(Debug, Clone)]
pub struct SheetMetrics {
    pub row_count: u32,
    pub column_count: u32,
    pub non_empty_cells: u32,
    pub formula_cells: u32,
    pub cached_values: u32,
    pub comments: u32,
    pub style_map: HashMap<String, StyleUsage>,
    pub classification: SheetClassification,
}

#[derive(Debug, Clone)]
pub struct StyleUsage {
    pub occurrences: u32,
    pub tags: Vec<String>,
    pub example_cells: Vec<String>,
}

impl WorkbookContext {
    pub fn load(_config: &Arc<ServerConfig>, path: &Path) -> Result<Self> {
        let metadata = fs::metadata(path)
            .with_context(|| format!("unable to read metadata for {:?}", path))?;
        let slug = path
            .file_stem()
            .map(|s| s.to_string_lossy().to_string())
            .unwrap_or_else(|| "workbook".to_string());
        let bytes = metadata.len();
        let last_modified = metadata.modified().ok().and_then(system_time_to_rfc3339);
        let id = WorkbookId(hash_path_metadata(path, &metadata));
        let spreadsheet =
            xlsx::read(path).with_context(|| format!("failed to parse workbook {:?}", path))?;
        let short_id = make_short_workbook_id(&slug, id.as_str());

        Ok(Self {
            id,
            short_id,
            slug,
            path: path.to_path_buf(),
            caps: BackendCaps::xlsx(),
            bytes,
            last_modified,
            spreadsheet: Arc::new(RwLock::new(spreadsheet)),
            sheet_cache: RwLock::new(HashMap::new()),
            formula_atlas: Arc::new(FormulaAtlas::default()),
        })
    }

    pub fn sheet_names(&self) -> Vec<String> {
        let book = self.spreadsheet.read();
        book.get_sheet_collection()
            .iter()
            .map(|sheet| sheet.get_name().to_string())
            .collect()
    }

    pub fn describe(&self) -> WorkbookDescription {
        let book = self.spreadsheet.read();
        let defined_names_count = book.get_defined_names().len();
        let table_count: usize = book
            .get_sheet_collection()
            .iter()
            .map(|sheet| sheet.get_tables().len())
            .sum();
        let macros_present = false;

        WorkbookDescription {
            workbook_id: self.id.clone(),
            short_id: self.short_id.clone(),
            slug: self.slug.clone(),
            path: path_to_forward_slashes(&self.path),
            bytes: self.bytes,
            sheet_count: book.get_sheet_collection().len(),
            defined_names: defined_names_count,
            tables: table_count,
            macros_present,
            last_modified: self
                .last_modified
                .map(|dt| dt.to_rfc3339_opts(chrono::SecondsFormat::Secs, true)),
            caps: self.caps.clone(),
        }
    }

    pub fn get_sheet_metrics(&self, sheet_name: &str) -> Result<Arc<SheetCacheEntry>> {
        if let Some(entry) = self.sheet_cache.read().get(sheet_name) {
            return Ok(entry.clone());
        }

        let mut writer = self.sheet_cache.write();
        if let Some(entry) = writer.get(sheet_name) {
            return Ok(entry.clone());
        }

        let book = self.spreadsheet.read();
        let sheet = book
            .get_sheet_by_name(sheet_name)
            .ok_or_else(|| anyhow!("sheet {} not found", sheet_name))?;
        let (metrics, style_tags) = compute_sheet_metrics(sheet);
        let named_ranges = gather_named_ranges(sheet, book.get_defined_names());
        let detected_regions = detect_regions(sheet, &metrics);

        let entry = Arc::new(SheetCacheEntry {
            metrics,
            style_tags,
            named_ranges,
            detected_regions,
        });

        writer.insert(sheet_name.to_string(), entry.clone());
        Ok(entry)
    }

    pub fn list_summaries(&self) -> Result<Vec<SheetSummary>> {
        let book = self.spreadsheet.read();
        let mut summaries = Vec::new();
        for sheet in book.get_sheet_collection() {
            let name = sheet.get_name().to_string();
            let entry = self.get_sheet_metrics(&name)?;
            summaries.push(SheetSummary {
                name: name.clone(),
                visible: sheet.get_sheet_state() != "hidden",
                row_count: entry.metrics.row_count,
                column_count: entry.metrics.column_count,
                non_empty_cells: entry.metrics.non_empty_cells,
                formula_cells: entry.metrics.formula_cells,
                cached_values: entry.metrics.cached_values,
                classification: entry.metrics.classification.clone(),
                style_tags: entry.style_tags.clone(),
            });
        }
        Ok(summaries)
    }

    pub fn with_sheet<T, F>(&self, sheet_name: &str, func: F) -> Result<T>
    where
        F: FnOnce(&Worksheet) -> T,
    {
        let book = self.spreadsheet.read();
        let sheet = book
            .get_sheet_by_name(sheet_name)
            .ok_or_else(|| anyhow!("sheet {} not found", sheet_name))?;
        Ok(func(sheet))
    }

    pub fn formula_graph(&self, sheet_name: &str) -> Result<FormulaGraph> {
        self.with_sheet(sheet_name, |sheet| {
            FormulaGraph::build(sheet, &self.formula_atlas)
        })?
    }

    pub fn named_items(&self) -> Result<Vec<NamedRangeDescriptor>> {
        let book = self.spreadsheet.read();
        let sheet_names: Vec<String> = book
            .get_sheet_collection()
            .iter()
            .map(|sheet| sheet.get_name().to_string())
            .collect();
        let mut items = Vec::new();

        for defined in book.get_defined_names() {
            let refers_to = defined.get_address();
            let scope = if defined.has_local_sheet_id() {
                let idx = *defined.get_local_sheet_id() as usize;
                sheet_names.get(idx).cloned()
            } else {
                None
            };
            let kind = if refers_to.starts_with('=') {
                NamedItemKind::Formula
            } else {
                NamedItemKind::NamedRange
            };

            items.push(NamedRangeDescriptor {
                name: defined.get_name().to_string(),
                scope: scope.clone(),
                refers_to: refers_to.clone(),
                kind,
                sheet_name: scope,
                comment: None,
            });
        }

        for sheet in book.get_sheet_collection() {
            for table in sheet.get_tables() {
                let start = table.get_area().0.get_coordinate();
                let end = table.get_area().1.get_coordinate();
                items.push(NamedRangeDescriptor {
                    name: table.get_name().to_string(),
                    scope: Some(sheet.get_name().to_string()),
                    refers_to: format!("{}:{}", start, end),
                    kind: NamedItemKind::Table,
                    sheet_name: Some(sheet.get_name().to_string()),
                    comment: None,
                });
            }
        }

        Ok(items)
    }

    pub fn sheet_overview(&self, sheet_name: &str) -> Result<SheetOverviewResponse> {
        let entry = self.get_sheet_metrics(sheet_name)?;
        let narrative = classification::narrative(&entry.metrics);
        let regions = classification::regions(&entry.metrics);
        let key_ranges = classification::key_ranges(&entry.metrics);
        let detected_regions = entry.detected_regions.clone();

        Ok(SheetOverviewResponse {
            workbook_id: self.id.clone(),
            workbook_short_id: self.short_id.clone(),
            sheet_name: sheet_name.to_string(),
            narrative,
            regions,
            detected_regions,
            key_ranges,
            formula_ratio: if entry.metrics.non_empty_cells == 0 {
                0.0
            } else {
                entry.metrics.formula_cells as f32 / entry.metrics.non_empty_cells as f32
            },
            notable_features: entry.style_tags.clone(),
        })
    }

    pub fn detected_region(
        &self,
        sheet_name: &str,
        id: u32,
    ) -> Result<crate::model::DetectedRegion> {
        let entry = self.get_sheet_metrics(sheet_name)?;
        entry
            .detected_regions
            .iter()
            .find(|r| r.id == id)
            .cloned()
            .ok_or_else(|| anyhow!("region {} not found on sheet {}", id, sheet_name))
    }
}

pub fn cell_to_value(cell: &umya_spreadsheet::Cell) -> Option<crate::model::CellValue> {
    let raw = cell.get_value();
    if raw.is_empty() {
        return None;
    }
    if let Ok(number) = raw.parse::<f64>() {
        return Some(crate::model::CellValue::Number(number));
    }

    let lower = raw.to_ascii_lowercase();
    if lower == "true" {
        return Some(crate::model::CellValue::Bool(true));
    }
    if lower == "false" {
        return Some(crate::model::CellValue::Bool(false));
    }

    Some(crate::model::CellValue::Text(raw.to_string()))
}

pub fn compute_sheet_metrics(sheet: &Worksheet) -> (SheetMetrics, Vec<String>) {
    use std::collections::HashMap as StdHashMap;
    let mut non_empty = 0u32;
    let mut formulas = 0u32;
    let mut cached = 0u32;
    let comments = sheet.get_comments().len() as u32;
    let mut style_usage: StdHashMap<String, StyleUsage> = StdHashMap::new();

    for cell in sheet.get_cell_collection() {
        let value = cell.get_value();
        if !value.is_empty() {
            non_empty += 1;
        }
        if cell.is_formula() {
            formulas += 1;
            if !cell.get_value().is_empty() {
                cached += 1;
            }
        }

        if let Some((style_key, usage)) = style::tag_cell(cell) {
            let entry = style_usage.entry(style_key).or_insert_with(|| StyleUsage {
                occurrences: 0,
                tags: usage.tags.clone(),
                example_cells: Vec::new(),
            });
            entry.occurrences += 1;
            if entry.example_cells.len() < 5 {
                entry.example_cells.push(usage.example_cell.clone());
            }
        }
    }

    let (max_col, max_row) = sheet.get_highest_column_and_row();

    let classification = classification::classify(
        non_empty,
        formulas,
        max_row,
        max_col,
        comments,
        &style_usage,
    );

    let style_tags: Vec<String> = style_usage
        .values()
        .flat_map(|usage| usage.tags.clone())
        .collect();

    let metrics = SheetMetrics {
        row_count: max_row,
        column_count: max_col,
        non_empty_cells: non_empty,
        formula_cells: formulas,
        cached_values: cached,
        comments,
        style_map: style_usage,
        classification,
    };
    (metrics, style_tags)
}

#[derive(Debug, Clone, Copy)]
struct Rect {
    start_row: u32,
    end_row: u32,
    start_col: u32,
    end_col: u32,
}

#[derive(Debug, Clone)]
struct CellInfo {
    value: Option<crate::model::CellValue>,
    is_formula: bool,
}

#[derive(Debug)]
struct Occupancy {
    cells: HashMap<(u32, u32), CellInfo>,
}

impl Occupancy {
    fn row_col_counts(&self, rect: &Rect) -> (Vec<u32>, Vec<u32>) {
        let mut row_counts = vec![0u32; (rect.end_row - rect.start_row + 1) as usize];
        let mut col_counts = vec![0u32; (rect.end_col - rect.start_col + 1) as usize];
        for ((row, col), _) in self.cells.iter() {
            if *row >= rect.start_row
                && *row <= rect.end_row
                && *col >= rect.start_col
                && *col <= rect.end_col
            {
                row_counts[(row - rect.start_row) as usize] += 1;
                col_counts[(col - rect.start_col) as usize] += 1;
            }
        }
        (row_counts, col_counts)
    }

    fn stats_in_rect(&self, rect: &Rect) -> RegionStats {
        let mut stats = RegionStats::default();
        for ((row, col), info) in self.cells.iter() {
            if *row < rect.start_row
                || *row > rect.end_row
                || *col < rect.start_col
                || *col > rect.end_col
            {
                continue;
            }
            stats.non_empty += 1;
            if info.is_formula {
                stats.formulas += 1;
            }
            if let Some(val) = &info.value {
                match val {
                    crate::model::CellValue::Text(_) => stats.text += 1,
                    crate::model::CellValue::Number(_) => stats.numbers += 1,
                    crate::model::CellValue::Bool(_) => stats.bools += 1,
                    crate::model::CellValue::Date(_) => stats.dates += 1,
                    crate::model::CellValue::Error(_) => stats.errors += 1,
                }
            }
        }
        stats
    }

    fn value_at(&self, row: u32, col: u32) -> Option<&crate::model::CellValue> {
        self.cells.get(&(row, col)).and_then(|c| c.value.as_ref())
    }
}

#[derive(Debug, Default, Clone)]
struct RegionStats {
    non_empty: u32,
    formulas: u32,
    text: u32,
    numbers: u32,
    bools: u32,
    dates: u32,
    errors: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Gutter {
    Row { start: u32, end: u32 },
    Col { start: u32, end: u32 },
}

fn detect_regions(sheet: &Worksheet, metrics: &SheetMetrics) -> Vec<crate::model::DetectedRegion> {
    if metrics.row_count == 0 || metrics.column_count == 0 {
        return Vec::new();
    }
    let occupancy = build_occupancy(sheet);
    let root = Rect {
        start_row: 1,
        end_row: metrics.row_count.max(1),
        start_col: 1,
        end_col: metrics.column_count.max(1),
    };

    let mut leaves = Vec::new();
    split_rect(&occupancy, &root, &mut leaves);

    let mut regions = Vec::new();
    for (idx, rect) in leaves.into_iter().enumerate() {
        if let Some(trimmed) = trim_rect(&occupancy, rect) {
            let region = build_region(&occupancy, &trimmed, metrics, idx as u32);
            regions.push(region);
        }
    }
    regions
}

fn build_occupancy(sheet: &Worksheet) -> Occupancy {
    let mut cells = HashMap::new();
    for cell in sheet.get_cell_collection() {
        let coord = cell.get_coordinate();
        let row = *coord.get_row_num();
        let col = *coord.get_col_num();
        let value = cell_to_value(cell);
        let is_formula = cell.is_formula();
        cells.insert((row, col), CellInfo { value, is_formula });
    }
    Occupancy { cells }
}

fn split_rect(occupancy: &Occupancy, rect: &Rect, leaves: &mut Vec<Rect>) {
    if rect.start_row >= rect.end_row && rect.start_col >= rect.end_col {
        leaves.push(*rect);
        return;
    }
    if let Some(gutter) = find_best_gutter(occupancy, rect) {
        match gutter {
            Gutter::Row { start, end } => {
                if start > rect.start_row {
                    let upper = Rect {
                        start_row: rect.start_row,
                        end_row: start - 1,
                        start_col: rect.start_col,
                        end_col: rect.end_col,
                    };
                    split_rect(occupancy, &upper, leaves);
                }
                if end < rect.end_row {
                    let lower = Rect {
                        start_row: end + 1,
                        end_row: rect.end_row,
                        start_col: rect.start_col,
                        end_col: rect.end_col,
                    };
                    split_rect(occupancy, &lower, leaves);
                }
            }
            Gutter::Col { start, end } => {
                if start > rect.start_col {
                    let left = Rect {
                        start_row: rect.start_row,
                        end_row: rect.end_row,
                        start_col: rect.start_col,
                        end_col: start - 1,
                    };
                    split_rect(occupancy, &left, leaves);
                }
                if end < rect.end_col {
                    let right = Rect {
                        start_row: rect.start_row,
                        end_row: rect.end_row,
                        start_col: end + 1,
                        end_col: rect.end_col,
                    };
                    split_rect(occupancy, &right, leaves);
                }
            }
        }
        return;
    }
    leaves.push(*rect);
}

fn find_best_gutter(occupancy: &Occupancy, rect: &Rect) -> Option<Gutter> {
    let (row_counts, col_counts) = occupancy.row_col_counts(rect);
    let width = rect.end_col - rect.start_col + 1;
    let height = rect.end_row - rect.start_row + 1;

    let row_blank_runs = find_blank_runs(&row_counts, width);
    let col_blank_runs = find_blank_runs(&col_counts, height);

    let mut best: Option<(Gutter, u32)> = None;

    if let Some((start, end, len)) = row_blank_runs {
        let gutter = Gutter::Row {
            start: rect.start_row + start,
            end: rect.start_row + end,
        };
        best = Some((gutter, len));
    }
    if let Some((start, end, len)) = col_blank_runs {
        let gutter = Gutter::Col {
            start: rect.start_col + start,
            end: rect.start_col + end,
        };
        if best.map(|(_, l)| len > l).unwrap_or(true) {
            best = Some((gutter, len));
        }
    }

    best.map(|(g, _)| g)
}

fn find_blank_runs(counts: &[u32], span: u32) -> Option<(u32, u32, u32)> {
    if counts.is_empty() {
        return None;
    }
    let mut best_start = 0;
    let mut best_end = 0;
    let mut best_len = 0;
    let mut current_start = None;
    for (idx, count) in counts.iter().enumerate() {
        let is_blank = *count == 0 || (*count as f32 / span as f32) < 0.05;
        if is_blank {
            if current_start.is_none() {
                current_start = Some(idx as u32);
            }
        } else if let Some(start) = current_start.take() {
            let end = idx as u32 - 1;
            let len = end - start + 1;
            if len > best_len && start > 0 && end + 1 < counts.len() as u32 {
                best_len = len;
                best_start = start;
                best_end = end;
            }
        }
    }
    if let Some(start) = current_start {
        let end = counts.len() as u32 - 1;
        let len = end - start + 1;
        if len > best_len && start > 0 && end + 1 < counts.len() as u32 {
            best_len = len;
            best_start = start;
            best_end = end;
        }
    }
    if best_len >= 2 {
        Some((best_start, best_end, best_len))
    } else {
        None
    }
}

fn trim_rect(occupancy: &Occupancy, rect: Rect) -> Option<Rect> {
    let mut r = rect;
    loop {
        let (row_counts, col_counts) = occupancy.row_col_counts(&r);
        let width = r.end_col - r.start_col + 1;
        let height = r.end_row - r.start_row + 1;
        let top_blank = row_counts
            .first()
            .map(|c| *c == 0 || (*c as f32 / width as f32) < 0.1)
            .unwrap_or(false);
        let bottom_blank = row_counts
            .last()
            .map(|c| *c == 0 || (*c as f32 / width as f32) < 0.1)
            .unwrap_or(false);
        let left_blank = col_counts
            .first()
            .map(|c| *c == 0 || (*c as f32 / height as f32) < 0.1)
            .unwrap_or(false);
        let right_blank = col_counts
            .last()
            .map(|c| *c == 0 || (*c as f32 / height as f32) < 0.1)
            .unwrap_or(false);

        let mut changed = false;
        if top_blank && r.start_row < r.end_row {
            r.start_row += 1;
            changed = true;
        }
        if bottom_blank && r.end_row > r.start_row {
            r.end_row -= 1;
            changed = true;
        }
        if left_blank && r.start_col < r.end_col {
            r.start_col += 1;
            changed = true;
        }
        if right_blank && r.end_col > r.start_col {
            r.end_col -= 1;
            changed = true;
        }

        if !changed {
            break;
        }
        if r.start_row > r.end_row || r.start_col > r.end_col {
            return None;
        }
    }
    Some(r)
}

fn build_region(
    occupancy: &Occupancy,
    rect: &Rect,
    metrics: &SheetMetrics,
    id: u32,
) -> crate::model::DetectedRegion {
    let header_info = detect_headers(occupancy, rect);
    let stats = occupancy.stats_in_rect(rect);
    let (kind, confidence) = classify_region(rect, &stats, &header_info, metrics);
    crate::model::DetectedRegion {
        id,
        bounds: format!(
            "{}{}:{}{}",
            crate::utils::column_number_to_name(rect.start_col),
            rect.start_row,
            crate::utils::column_number_to_name(rect.end_col),
            rect.end_row
        ),
        header_row: header_info.header_row,
        headers: header_info.headers,
        row_count: rect.end_row - rect.start_row + 1,
        classification: kind.clone(),
        region_kind: Some(kind),
        confidence,
    }
}

#[derive(Debug, Default)]
struct HeaderInfo {
    header_row: Option<u32>,
    headers: Vec<String>,
}

fn detect_headers(occupancy: &Occupancy, rect: &Rect) -> HeaderInfo {
    let mut candidates = Vec::new();
    let max_row = rect.start_row.saturating_add(2).min(rect.end_row);
    for row in rect.start_row..=max_row {
        let mut text = 0;
        let mut numbers = 0;
        let mut non_empty = 0;
        let mut unique = HashSet::new();
        for col in rect.start_col..=rect.end_col {
            if let Some(val) = occupancy.value_at(row, col) {
                non_empty += 1;
                match val {
                    crate::model::CellValue::Text(s) => {
                        text += 1;
                        unique.insert(s.clone());
                    }
                    crate::model::CellValue::Number(_) => numbers += 1,
                    crate::model::CellValue::Bool(_) => text += 1,
                    crate::model::CellValue::Date(_) => text += 1,
                    crate::model::CellValue::Error(_) => {}
                }
            }
        }
        if non_empty == 0 {
            continue;
        }
        let score = text as f32 + unique.len() as f32 * 0.2 - numbers as f32 * 0.3;
        candidates.push((row, score, text, non_empty));
    }

    let header_candidates: Vec<&(u32, f32, u32, u32)> = candidates
        .iter()
        .filter(|(_, _, text, non_empty)| *text >= 1 && *text * 2 >= *non_empty)
        .collect();

    let best = header_candidates.iter().copied().max_by(|a, b| {
        a.1.partial_cmp(&b.1)
            .unwrap_or(Ordering::Equal)
            .then_with(|| b.0.cmp(&a.0))
    });
    let earliest = header_candidates
        .iter()
        .copied()
        .min_by(|a, b| a.0.cmp(&b.0));

    let maybe_header = match (best, earliest) {
        (Some(best_row), Some(early_row)) => {
            if (best_row.1 - early_row.1).abs() <= 0.3 {
                Some(early_row.0)
            } else {
                Some(best_row.0)
            }
        }
        (Some(best_row), None) => Some(best_row.0),
        _ => None,
    };

    let mut header_rows = Vec::new();
    if let Some(hr) = maybe_header {
        header_rows.push(hr);
        if hr < rect.end_row
            && let Some((_, score_next, text_next, non_empty_next)) =
                candidates.iter().find(|(r, _, _, _)| *r == hr + 1)
            && *text_next >= 1
            && *text_next * 2 >= *non_empty_next
            && *score_next
                >= 0.6
                    * candidates
                        .iter()
                        .find(|(r, _, _, _)| *r == hr)
                        .map(|c| c.1)
                        .unwrap_or(0.0)
        {
            header_rows.push(hr + 1);
        }
    }

    let mut headers = Vec::new();
    for col in rect.start_col..=rect.end_col {
        let mut parts = Vec::new();
        for hr in &header_rows {
            if let Some(val) = occupancy.value_at(*hr, col) {
                match val {
                    crate::model::CellValue::Text(s) if !s.trim().is_empty() => {
                        parts.push(s.trim().to_string())
                    }
                    crate::model::CellValue::Number(n) => parts.push(n.to_string()),
                    crate::model::CellValue::Bool(b) => parts.push(b.to_string()),
                    crate::model::CellValue::Date(d) => parts.push(d.clone()),
                    crate::model::CellValue::Error(e) => parts.push(e.clone()),
                    _ => {}
                }
            }
        }
        if parts.is_empty() {
            headers.push(crate::utils::column_number_to_name(col));
        } else {
            headers.push(parts.join(" / "));
        }
    }

    HeaderInfo {
        header_row: header_rows.first().copied(),
        headers,
    }
}

fn classify_region(
    rect: &Rect,
    stats: &RegionStats,
    header_info: &HeaderInfo,
    metrics: &SheetMetrics,
) -> (crate::model::RegionKind, f32) {
    let width = rect.end_col - rect.start_col + 1;
    let height = rect.end_row - rect.start_row + 1;
    let area = width.max(1) * height.max(1);
    let density = if area == 0 {
        0.0
    } else {
        stats.non_empty as f32 / area as f32
    };
    let formula_ratio = if stats.non_empty == 0 {
        0.0
    } else {
        stats.formulas as f32 / stats.non_empty as f32
    };
    let text_ratio = if stats.non_empty == 0 {
        0.0
    } else {
        stats.text as f32 / stats.non_empty as f32
    };

    let mut kind = crate::model::RegionKind::Data;
    if formula_ratio > 0.25 && is_outputs_band(rect, metrics, height, width) {
        kind = crate::model::RegionKind::Outputs;
    } else if formula_ratio > 0.55 {
        kind = crate::model::RegionKind::Calculator;
    } else if height <= 3
        && width <= 4
        && text_ratio > 0.5
        && rect.end_row >= metrics.row_count.saturating_sub(3)
    {
        kind = crate::model::RegionKind::Metadata;
    } else if formula_ratio < 0.25
        && stats.numbers > 0
        && stats.text > 0
        && text_ratio >= 0.3
        && (width <= 2 || (width <= 3 && header_info.header_row.is_none()))
    {
        kind = crate::model::RegionKind::Parameters;
    } else if height <= 4 && width <= 6 && formula_ratio < 0.2 && text_ratio > 0.4 && density < 0.5
    {
        kind = crate::model::RegionKind::Metadata;
    }

    let mut confidence: f32 = 0.4;
    if header_info.header_row.is_some() {
        confidence += 0.2;
    }
    confidence += (density * 0.2).min(0.2);
    confidence += (formula_ratio * 0.2).min(0.2);
    if matches!(
        kind,
        crate::model::RegionKind::Parameters | crate::model::RegionKind::Metadata
    ) && width <= 4
    {
        confidence += 0.1;
    }
    if confidence > 1.0 {
        confidence = 1.0;
    }

    (kind, confidence)
}

fn is_outputs_band(rect: &Rect, metrics: &SheetMetrics, height: u32, width: u32) -> bool {
    let near_bottom = rect.end_row >= metrics.row_count.saturating_sub(6);
    let near_right = rect.end_col >= metrics.column_count.saturating_sub(3);
    let is_shallow = height <= 6;
    let is_narrow_at_edge = width <= 6 && near_right;
    let not_at_top_left = rect.start_row > 1 || rect.start_col > 1;
    let sheet_has_depth = metrics.row_count > 10 || metrics.column_count > 6;
    let is_band = (is_shallow && near_bottom) || is_narrow_at_edge;
    is_band && not_at_top_left && sheet_has_depth
}

fn gather_named_ranges(
    sheet: &Worksheet,
    defined_names: &[DefinedName],
) -> Vec<NamedRangeDescriptor> {
    let name_str = sheet.get_name();
    defined_names
        .iter()
        .filter(|name| name.get_address().contains(name_str))
        .map(|name| NamedRangeDescriptor {
            name: name.get_name().to_string(),
            scope: if name.has_local_sheet_id() {
                Some(name_str.to_string())
            } else {
                None
            },
            refers_to: name.get_address(),
            kind: NamedItemKind::NamedRange,
            sheet_name: Some(name_str.to_string()),
            comment: None,
        })
        .collect()
}

pub fn build_workbook_list(
    config: &Arc<ServerConfig>,
    filter: &WorkbookFilter,
) -> Result<WorkbookListResponse> {
    let mut descriptors = Vec::new();

    if let Some(single) = config.single_workbook() {
        let metadata = fs::metadata(single)
            .with_context(|| format!("unable to read metadata for {:?}", single))?;
        let id = WorkbookId(hash_path_metadata(single, &metadata));
        let slug = single
            .file_stem()
            .map(|s| s.to_string_lossy().to_string())
            .unwrap_or_else(|| "workbook".to_string());
        let folder = derive_folder(config, single);
        let short_id = make_short_workbook_id(&slug, id.as_str());
        let caps = BackendCaps::xlsx();

        if filter.matches(&slug, folder.as_deref(), single) {
            let relative = single
                .strip_prefix(&config.workspace_root)
                .unwrap_or(single);
            let descriptor = crate::model::WorkbookDescriptor {
                workbook_id: id,
                short_id,
                slug,
                folder,
                path: path_to_forward_slashes(relative),
                bytes: metadata.len(),
                last_modified: metadata
                    .modified()
                    .ok()
                    .and_then(system_time_to_rfc3339)
                    .map(|dt| dt.to_rfc3339_opts(chrono::SecondsFormat::Secs, true)),
                caps,
            };
            descriptors.push(descriptor);
        }

        return Ok(WorkbookListResponse {
            workbooks: descriptors,
        });
    }

    use walkdir::WalkDir;

    for entry in WalkDir::new(&config.workspace_root) {
        let entry = entry?;
        if !entry.file_type().is_file() {
            continue;
        }
        let path = entry.path();
        if !has_supported_extension(&config.supported_extensions, path) {
            continue;
        }
        let metadata = entry.metadata()?;
        let id = WorkbookId(hash_path_metadata(path, &metadata));
        let slug = path
            .file_stem()
            .map(|s| s.to_string_lossy().to_string())
            .unwrap_or_else(|| "workbook".to_string());
        let folder = derive_folder(config, path);
        let short_id = make_short_workbook_id(&slug, id.as_str());
        let caps = BackendCaps::xlsx();

        if !filter.matches(&slug, folder.as_deref(), path) {
            continue;
        }

        let relative = path.strip_prefix(&config.workspace_root).unwrap_or(path);
        let descriptor = crate::model::WorkbookDescriptor {
            workbook_id: id,
            short_id,
            slug,
            folder,
            path: path_to_forward_slashes(relative),
            bytes: metadata.len(),
            last_modified: metadata
                .modified()
                .ok()
                .and_then(system_time_to_rfc3339)
                .map(|dt| dt.to_rfc3339_opts(chrono::SecondsFormat::Secs, true)),
            caps,
        };
        descriptors.push(descriptor);
    }

    descriptors.sort_by(|a, b| a.slug.cmp(&b.slug));

    Ok(WorkbookListResponse {
        workbooks: descriptors,
    })
}

fn derive_folder(config: &Arc<ServerConfig>, path: &Path) -> Option<String> {
    path.strip_prefix(&config.workspace_root)
        .ok()
        .and_then(|relative| relative.parent())
        .and_then(|parent| parent.file_name())
        .map(|os| os.to_string_lossy().to_string())
}

fn has_supported_extension(allowed: &[String], path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| {
            let lower = ext.to_ascii_lowercase();
            allowed.iter().any(|candidate| candidate == &lower)
        })
        .unwrap_or(false)
}
