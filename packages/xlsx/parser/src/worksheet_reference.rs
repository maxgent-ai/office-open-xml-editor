use crate::worksheet_projector::{WorksheetProjectorItem, WorksheetRowProjector};
use crate::{attr_bool, resolve_sheet_path, Row};
use crate::{CellRange, CellValue, SharedString, SheetMeta, XlsxZip};
use ooxml_common::depth::parse_guarded;
use ooxml_common::ns::is_x_ns;
use ooxml_common::resource::{observe_hard_limit, HardResourceLimitKind};
use std::collections::{BTreeMap, HashMap};

pub(crate) const MAX_REFERENCE_CELLS: usize = 1_000_000;
pub(crate) const MAX_TOTAL_REFERENCE_CELLS: usize = 1_000_000;
const MAX_TOTAL_REFERENCE_STRING_BYTES: usize = 64 * 1024 * 1024;
const MAX_TOTAL_INDEXED_CELLS: usize = MAX_TOTAL_REFERENCE_CELLS;
const MAX_TOTAL_INDEXED_STRING_BYTES: usize = MAX_TOTAL_REFERENCE_STRING_BYTES;
const MAX_COL: u32 = 16_384;
const MAX_ROW: u32 = 1_048_576;

/// Reference admission is library resource policy, not an unsupported formula.
/// Latch before returning a provisional absence so every package/cursor boundary
/// rejects it, including callers that intentionally default unresolved formulas.
fn reference_limit(kind: HardResourceLimitKind, limit: usize, observed: usize) -> Option<()> {
    if observed <= limit {
        return Some(());
    }
    let _ = observe_hard_limit(kind, None, limit as u64, observed as u64);
    None
}

#[derive(Clone, Debug, PartialEq)]
pub(crate) enum ReferencedCellValue {
    Empty,
    Text(String),
    Number(f64),
}

impl ReferencedCellValue {
    fn string_bytes(&self) -> usize {
        match self {
            Self::Text(text) => text.len(),
            _ => 0,
        }
    }
}

fn referenced_cell_value(
    value: &CellValue,
    shared_strings: &[SharedString],
) -> ReferencedCellValue {
    match value {
        CellValue::Text { text, .. } => ReferencedCellValue::Text(text.clone()),
        CellValue::Shared { si } => shared_strings
            .get(*si)
            .map(|string| ReferencedCellValue::Text(string.text.clone()))
            .unwrap_or(ReferencedCellValue::Empty),
        CellValue::Number { number } if number.is_finite() => ReferencedCellValue::Number(*number),
        _ => ReferencedCellValue::Empty,
    }
}

pub(crate) struct WorksheetCellLookup {
    cells: HashMap<(u32, u32), ReferencedCell>,
    hidden_row_ranges: Vec<(u32, u32)>,
    hidden_column_ranges: Vec<(u32, u32)>,
    string_bytes: usize,
}

#[derive(Debug, PartialEq)]
struct ReferencedCell {
    value: ReferencedCellValue,
}

impl WorksheetCellLookup {
    fn resource_units(&self) -> usize {
        self.cells
            .len()
            .saturating_add(self.hidden_row_ranges.len())
            .saturating_add(self.hidden_column_ranges.len())
    }
}

pub(crate) type ReferenceIndexLimit = (HardResourceLimitKind, usize, usize);

fn report_index_limit((kind, limit, observed): ReferenceIndexLimit) {
    let _ = reference_limit(kind, limit, observed);
}

pub(crate) struct WorksheetCellLookupBuilder {
    worksheet: WorksheetCellLookup,
    max_cells: usize,
    max_string_bytes: usize,
    limit_failure: Option<ReferenceIndexLimit>,
}

impl WorksheetCellLookupBuilder {
    pub(crate) fn bounded() -> Self {
        Self::new(MAX_TOTAL_INDEXED_CELLS, MAX_TOTAL_INDEXED_STRING_BYTES)
    }

    pub(crate) fn new(max_cells: usize, max_string_bytes: usize) -> Self {
        Self {
            worksheet: WorksheetCellLookup {
                cells: HashMap::new(),
                hidden_row_ranges: Vec::new(),
                hidden_column_ranges: Vec::new(),
                string_bytes: 0,
            },
            max_cells,
            max_string_bytes,
            limit_failure: None,
        }
    }

    pub(crate) fn limit_failure(&self) -> Option<ReferenceIndexLimit> {
        self.limit_failure
    }

    // A bounded cache admission may recover by evicting an unused prebuilt
    // index. Preserve the refusal until the resolver knows whether content
    // would be lost; only an unrecoverable refusal poisons the operation.
    fn check_limit(
        &mut self,
        kind: HardResourceLimitKind,
        limit: usize,
        observed: usize,
    ) -> Option<()> {
        if observed > limit {
            self.limit_failure = Some((kind, limit, observed));
            return None;
        }
        Some(())
    }

    fn push_cell(&mut self, row: u32, col: u32, value: ReferencedCellValue) -> Option<()> {
        if !(1..=MAX_COL).contains(&col) || !(1..=MAX_ROW).contains(&row) {
            return Some(());
        }
        let key = (row, col);
        if value == ReferencedCellValue::Empty {
            if let Some(previous) = self.worksheet.cells.remove(&key) {
                self.worksheet.string_bytes = self
                    .worksheet
                    .string_bytes
                    .checked_sub(previous.value.string_bytes())?;
            }
            return Some(());
        }
        let old_string_bytes = self
            .worksheet
            .cells
            .get(&key)
            .map(|cell| cell.value.string_bytes())
            .unwrap_or(0);
        let next_cell_count =
            self.worksheet.cells.len() + usize::from(!self.worksheet.cells.contains_key(&key));
        let next_string_bytes = self
            .worksheet
            .string_bytes
            .checked_sub(old_string_bytes)?
            .checked_add(value.string_bytes())?;
        self.check_limit(
            HardResourceLimitKind::WorksheetReferenceIndexCells,
            self.max_cells,
            next_cell_count
                .saturating_add(self.worksheet.hidden_row_ranges.len())
                .saturating_add(self.worksheet.hidden_column_ranges.len()),
        )?;
        self.check_limit(
            HardResourceLimitKind::WorksheetReferenceIndexStringBytes,
            self.max_string_bytes,
            next_string_bytes,
        )?;
        self.worksheet.cells.insert(key, ReferencedCell { value });
        self.worksheet.string_bytes = next_string_bytes;
        Some(())
    }

    pub(crate) fn push_row(&mut self, row: &Row, shared_strings: &[SharedString]) -> Option<()> {
        if row.hidden {
            self.push_hidden_row(row.index)?;
        }
        for cell in &row.cells {
            if !(1..=MAX_COL).contains(&cell.col) || !(1..=MAX_ROW).contains(&cell.row) {
                continue;
            }
            let value = referenced_cell_value(&cell.value, shared_strings);
            self.push_cell(cell.row, cell.col, value)?;
        }
        Some(())
    }

    fn push_hidden_row(&mut self, row: u32) -> Option<()> {
        if !(1..=MAX_ROW).contains(&row) {
            return Some(());
        }
        if let Some(last) = self.worksheet.hidden_row_ranges.last_mut() {
            if row >= last.0 && row <= last.1.saturating_add(1) {
                last.1 = last.1.max(row);
                return Some(());
            }
        }
        self.check_limit(
            HardResourceLimitKind::WorksheetReferenceIndexCells,
            self.max_cells,
            self.worksheet.resource_units().saturating_add(1),
        )?;
        self.worksheet.hidden_row_ranges.push((row, row));
        Some(())
    }

    pub(crate) fn mark_hidden_columns(
        &mut self,
        hidden_columns: &BTreeMap<u32, bool>,
    ) -> Option<()> {
        let mut ranges = Vec::<(u32, u32)>::new();
        for column in hidden_columns
            .iter()
            .filter_map(|(column, hidden)| (*hidden).then_some(*column))
        {
            if let Some(last) = ranges.last_mut() {
                if column <= last.1.saturating_add(1) {
                    last.1 = last.1.max(column);
                    continue;
                }
            }
            ranges.push((column, column));
        }
        self.check_limit(
            HardResourceLimitKind::WorksheetReferenceIndexCells,
            self.max_cells,
            self.worksheet.resource_units().saturating_add(ranges.len()),
        )?;
        self.worksheet.hidden_column_ranges = ranges;
        Some(())
    }

    pub(crate) fn finish(mut self) -> WorksheetCellLookup {
        self.worksheet.hidden_row_ranges.sort_unstable();
        let mut merged = Vec::<(u32, u32)>::new();
        for (start, end) in self.worksheet.hidden_row_ranges.drain(..) {
            if let Some(last) = merged.last_mut() {
                if start <= last.1.saturating_add(1) {
                    last.1 = last.1.max(end);
                    continue;
                }
            }
            merged.push((start, end));
        }
        self.worksheet.hidden_row_ranges = merged;
        self.worksheet
    }
}

pub(crate) fn extend_lookup_transactionally(
    lookup: &mut Option<WorksheetCellLookupBuilder>,
    rows: &[Row],
    shared_strings: &[SharedString],
) -> Option<ReferenceIndexLimit> {
    let builder = lookup.as_mut()?;
    if rows
        .iter()
        .any(|row| builder.push_row(row, shared_strings).is_none())
    {
        let failure = builder.limit_failure();
        *lookup = None;
        return failure;
    }
    None
}

/// Per-worksheet-parse state shared by charts and sparklines. Source sheets
/// are parsed once into sparse non-empty-cell maps. Independent cumulative
/// cell and UTF-8 byte budgets bound both those indexes and the dense reference
/// vectors retained by the resulting model.
pub(crate) struct WorksheetReferenceSession {
    sheets: HashMap<String, Option<WorksheetCellLookup>>,
    current_sheet: CurrentSheetLookupEntry,
    remaining_cells: usize,
    remaining_string_bytes: usize,
    remaining_indexed_cells: usize,
    remaining_indexed_string_bytes: usize,
    remaining_physical_indexed_cells: usize,
    remaining_physical_indexed_string_bytes: usize,
    /// Chart source-linked number formats read the style of a reference's
    /// first cell; the value index above deliberately keeps no styles.
    pub(crate) chart_source_styles: crate::chart::ChartSourceStyleCache,
}

enum CurrentSheetLookupEntry {
    Unseeded,
    PrebuiltUncharged {
        name: String,
        lookup: WorksheetCellLookup,
    },
    Unavailable {
        name: String,
        limit_failure: Option<ReferenceIndexLimit>,
    },
}

impl Default for WorksheetReferenceSession {
    fn default() -> Self {
        Self {
            sheets: HashMap::new(),
            current_sheet: CurrentSheetLookupEntry::Unseeded,
            remaining_cells: MAX_TOTAL_REFERENCE_CELLS,
            remaining_string_bytes: MAX_TOTAL_REFERENCE_STRING_BYTES,
            remaining_indexed_cells: MAX_TOTAL_INDEXED_CELLS,
            remaining_indexed_string_bytes: MAX_TOTAL_INDEXED_STRING_BYTES,
            remaining_physical_indexed_cells: MAX_TOTAL_INDEXED_CELLS,
            remaining_physical_indexed_string_bytes: MAX_TOTAL_INDEXED_STRING_BYTES,
            chart_source_styles: Default::default(),
        }
    }
}

impl WorksheetReferenceSession {
    pub(crate) fn seed_current_sheet(
        &mut self,
        sheet_name: &str,
        worksheet: Option<WorksheetCellLookup>,
    ) {
        self.current_sheet = match worksheet {
            Some(lookup)
                if lookup.resource_units() <= self.remaining_physical_indexed_cells
                    && lookup.string_bytes <= self.remaining_physical_indexed_string_bytes =>
            {
                self.remaining_physical_indexed_cells -= lookup.resource_units();
                self.remaining_physical_indexed_string_bytes -= lookup.string_bytes;
                CurrentSheetLookupEntry::PrebuiltUncharged {
                    name: sheet_name.to_string(),
                    lookup,
                }
            }
            Some(lookup) => CurrentSheetLookupEntry::Unavailable {
                name: sheet_name.to_string(),
                limit_failure: Some(
                    if lookup.resource_units() > self.remaining_physical_indexed_cells {
                        (
                            HardResourceLimitKind::WorksheetReferenceIndexCells,
                            self.remaining_physical_indexed_cells,
                            lookup.resource_units(),
                        )
                    } else {
                        (
                            HardResourceLimitKind::WorksheetReferenceIndexStringBytes,
                            self.remaining_physical_indexed_string_bytes,
                            lookup.string_bytes,
                        )
                    },
                ),
            },
            None => CurrentSheetLookupEntry::Unavailable {
                name: sheet_name.to_string(),
                limit_failure: None,
            },
        };
    }

    pub(crate) fn seed_current_sheet_failure(
        &mut self,
        sheet_name: &str,
        failure: ReferenceIndexLimit,
    ) {
        self.current_sheet = CurrentSheetLookupEntry::Unavailable {
            name: sheet_name.to_string(),
            limit_failure: Some(failure),
        };
    }

    fn reservable_cell_count(&self, range: &CellRange) -> Option<usize> {
        // A valid Excel grid can exceed usize on wasm32. Saturating the
        // observation still proves the existing one-million-cell crossing.
        let total = reference_cell_count(range).unwrap_or(usize::MAX);
        reference_limit(
            HardResourceLimitKind::WorksheetReferenceCells,
            MAX_REFERENCE_CELLS,
            total,
        )?;
        reference_limit(
            HardResourceLimitKind::WorksheetReferenceCells,
            MAX_TOTAL_REFERENCE_CELLS,
            (MAX_TOTAL_REFERENCE_CELLS - self.remaining_cells).saturating_add(total),
        )?;
        Some(total)
    }

    fn consume_result(&mut self, cell_count: usize, string_bytes: usize) {
        self.remaining_cells -= cell_count;
        self.remaining_string_bytes -= string_bytes;
    }

    fn consume_new_index(&mut self, worksheet: &WorksheetCellLookup) {
        self.remaining_indexed_cells -= worksheet.resource_units();
        self.remaining_indexed_string_bytes -= worksheet.string_bytes;
        self.remaining_physical_indexed_cells -= worksheet.resource_units();
        self.remaining_physical_indexed_string_bytes -= worksheet.string_bytes;
    }

    fn index_materialized_rows(
        &self,
        rows: &[Row],
        hidden_columns: Option<&BTreeMap<u32, bool>>,
        shared_strings: &[SharedString],
    ) -> LookupBuildResult {
        let mut builder = WorksheetCellLookupBuilder::new(
            self.remaining_indexed_cells
                .min(self.remaining_physical_indexed_cells),
            self.remaining_indexed_string_bytes
                .min(self.remaining_physical_indexed_string_bytes),
        );
        for row in rows {
            if builder.push_row(row, shared_strings).is_none() {
                return LookupBuildResult::LimitExceeded(
                    builder
                        .limit_failure()
                        .expect("bounded builder records its refusal"),
                );
            }
        }
        if hidden_columns.is_some_and(|columns| builder.mark_hidden_columns(columns).is_none()) {
            return LookupBuildResult::LimitExceeded(
                builder
                    .limit_failure()
                    .expect("bounded builder records its refusal"),
            );
        }
        LookupBuildResult::Built(builder.finish())
    }

    fn resolve_current_sheet(
        &mut self,
        sheet_name: &str,
        rows: Option<&[Row]>,
        hidden_columns: Option<&BTreeMap<u32, bool>>,
        shared_strings: &[SharedString],
    ) {
        if self.sheets.contains_key(sheet_name) {
            return;
        }
        let state = std::mem::replace(&mut self.current_sheet, CurrentSheetLookupEntry::Unseeded);
        let lookup = match state {
            CurrentSheetLookupEntry::PrebuiltUncharged { name, lookup } if name == sheet_name => {
                if lookup.resource_units() <= self.remaining_indexed_cells
                    && lookup.string_bytes <= self.remaining_indexed_string_bytes
                {
                    self.remaining_indexed_cells -= lookup.resource_units();
                    self.remaining_indexed_string_bytes -= lookup.string_bytes;
                    Some(lookup)
                } else {
                    let _ = reference_limit(
                        HardResourceLimitKind::WorksheetReferenceIndexCells,
                        self.remaining_indexed_cells,
                        lookup.resource_units(),
                    );
                    let _ = reference_limit(
                        HardResourceLimitKind::WorksheetReferenceIndexStringBytes,
                        self.remaining_indexed_string_bytes,
                        lookup.string_bytes,
                    );
                    self.remaining_physical_indexed_cells += lookup.resource_units();
                    self.remaining_physical_indexed_string_bytes += lookup.string_bytes;
                    None
                }
            }
            CurrentSheetLookupEntry::Unavailable {
                name,
                limit_failure,
            } if name == sheet_name => {
                if let Some(failure) = limit_failure {
                    report_index_limit(failure);
                }
                None
            }
            other => {
                self.current_sheet = other;
                match rows
                    .map(|rows| self.index_materialized_rows(rows, hidden_columns, shared_strings))
                {
                    Some(LookupBuildResult::Built(lookup)) => {
                        self.consume_new_index(&lookup);
                        Some(lookup)
                    }
                    Some(LookupBuildResult::LimitExceeded(failure)) => {
                        report_index_limit(failure);
                        None
                    }
                    _ => None,
                }
            }
        };
        self.sheets.insert(sheet_name.to_string(), lookup);
    }

    fn take_unreferenced_prebuilt(&mut self) -> Option<(String, WorksheetCellLookup)> {
        let state = std::mem::replace(&mut self.current_sheet, CurrentSheetLookupEntry::Unseeded);
        match state {
            CurrentSheetLookupEntry::PrebuiltUncharged { name, lookup } => {
                self.remaining_physical_indexed_cells += lookup.resource_units();
                self.remaining_physical_indexed_string_bytes += lookup.string_bytes;
                Some((name, lookup))
            }
            other => {
                self.current_sheet = other;
                None
            }
        }
    }

    fn restore_unreferenced_prebuilt(&mut self, name: String, lookup: WorksheetCellLookup) {
        self.remaining_physical_indexed_cells = self
            .remaining_physical_indexed_cells
            .checked_sub(lookup.resource_units())
            .expect("restored prebuilt cells were previously refunded");
        self.remaining_physical_indexed_string_bytes = self
            .remaining_physical_indexed_string_bytes
            .checked_sub(lookup.string_bytes)
            .expect("restored prebuilt strings were previously refunded");
        self.current_sheet = CurrentSheetLookupEntry::PrebuiltUncharged { name, lookup };
    }
}

enum LookupBuildResult {
    Built(WorksheetCellLookup),
    LimitExceeded(ReferenceIndexLimit),
    Unavailable,
}

pub(crate) fn split_sheet_ref(formula: &str) -> Option<(Option<String>, String)> {
    let formula = formula.trim();
    if formula.is_empty() {
        return None;
    }

    let (sheet_name, reference) = match formula.rfind('!') {
        Some(bang) => {
            let raw_sheet = formula[..bang].trim();
            let reference = formula[bang + 1..].trim();
            if raw_sheet.is_empty() || reference.is_empty() {
                return None;
            }
            let sheet = if raw_sheet.starts_with('\'') && raw_sheet.ends_with('\'') {
                if raw_sheet.len() < 2 {
                    return None;
                }
                raw_sheet[1..raw_sheet.len() - 1].replace("''", "'")
            } else {
                if raw_sheet.contains(['\'', '!', '[', ']']) {
                    return None;
                }
                raw_sheet.to_string()
            };
            (Some(sheet), reference)
        }
        None => (None, formula),
    };

    Some((sheet_name, reference.replace('$', "")))
}

fn parse_a1_cell(reference: &str) -> Option<(u32, u32)> {
    let split = reference
        .find(|c: char| !c.is_ascii_alphabetic())
        .unwrap_or(reference.len());
    let (col_text, row_text) = reference.split_at(split);
    if col_text.is_empty() || row_text.is_empty() || !row_text.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }

    let col = col_text.chars().try_fold(0u32, |value, c| {
        value
            .checked_mul(26)?
            .checked_add(c.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
    })?;
    let row = row_text.parse::<u32>().ok()?;
    if !(1..=MAX_COL).contains(&col) || !(1..=MAX_ROW).contains(&row) {
        return None;
    }
    Some((col, row))
}

pub(crate) fn parse_a1_range(reference: &str) -> Option<CellRange> {
    let mut parts = reference.trim().split(':');
    let first = parts.next()?;
    let second = parts.next();
    if parts.next().is_some() {
        return None;
    }
    let (col_a, row_a) = parse_a1_cell(first)?;
    let (col_b, row_b) = match second {
        Some(cell) => parse_a1_cell(cell)?,
        None => (col_a, row_a),
    };
    Some(CellRange {
        top: row_a.min(row_b),
        left: col_a.min(col_b),
        bottom: row_a.max(row_b),
        right: col_a.max(col_b),
    })
}

#[cfg(test)]
fn parse_worksheet_cells(
    sheet_xml: &str,
    shared_strings: &[SharedString],
    max_cells: usize,
    max_string_bytes: usize,
) -> Option<WorksheetCellLookup> {
    let mut builder = WorksheetCellLookupBuilder::new(max_cells, max_string_bytes);
    crate::stream_worksheet_rows(sheet_xml, shared_strings, &[], |row| {
        builder
            .push_row(&row, shared_strings)
            .ok_or_else(|| "worksheet reference index resource limit".to_string())
    })
    .ok()?;
    Some(builder.finish())
}

fn parse_package_worksheet_cells(
    archive: &mut XlsxZip,
    part: &str,
    shared_strings: &[SharedString],
    max_cells: usize,
    max_string_bytes: usize,
) -> LookupBuildResult {
    let Ok(operation) = archive.operation() else {
        return LookupBuildResult::Unavailable;
    };
    let Ok(entry) = operation.open_entry(part) else {
        return LookupBuildResult::Unavailable;
    };
    let Ok(mut projector) =
        WorksheetRowProjector::from_package_entry(entry, shared_strings, &[] as &[String])
    else {
        return LookupBuildResult::Unavailable;
    };
    let mut builder = WorksheetCellLookupBuilder::new(max_cells, max_string_bytes);
    let mut refusal = None;
    loop {
        match projector.next_item() {
            Ok(WorksheetProjectorItem::Row(row)) => {
                if refusal.is_none() && builder.push_row(&row.row, shared_strings).is_none() {
                    // Stop growing the index, but finish source validation.
                    // A malformed/truncated tail makes this source unavailable;
                    // its provisional index refusal must not poison the package.
                    refusal = builder.limit_failure();
                }
            }
            Ok(WorksheetProjectorItem::Finished(tail)) => {
                if let Some(failure) = refusal {
                    return LookupBuildResult::LimitExceeded(failure);
                }
                if builder
                    .mark_hidden_columns(&hidden_columns_from_shell(&tail.shell_xml))
                    .is_none()
                {
                    return LookupBuildResult::LimitExceeded(
                        builder
                            .limit_failure()
                            .expect("bounded builder records its refusal"),
                    );
                }
                return LookupBuildResult::Built(builder.finish());
            }
            Err(_) => return LookupBuildResult::Unavailable,
        }
    }
}

fn hidden_columns_from_shell(shell_xml: &str) -> BTreeMap<u32, bool> {
    let Ok(document) = parse_guarded(shell_xml) else {
        return BTreeMap::new();
    };
    // Merge the authored ranges before expanding them: every hidden `<col>`
    // may span all 16 384 columns, so per-declaration expansion would be
    // quadratic in the number of declarations.
    let mut ranges = Vec::new();
    for column in document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().name() == "col"
            && is_x_ns(node.tag_name().namespace())
            && attr_bool(node, "hidden") == Some(true)
    }) {
        let Some(min) = column
            .attribute("min")
            .and_then(|value| value.parse::<u32>().ok())
            .filter(|value| (1..=MAX_COL).contains(value))
        else {
            continue;
        };
        let max = column
            .attribute("max")
            .and_then(|value| value.parse::<u32>().ok())
            .filter(|value| *value >= min)
            .unwrap_or(min)
            .min(MAX_COL);
        ranges.push((min, max));
    }
    ranges.sort_unstable();
    let mut hidden = BTreeMap::new();
    let mut next = 1;
    for (min, max) in ranges {
        for index in min.max(next)..=max {
            hidden.insert(index, true);
        }
        next = next.max(max.saturating_add(1));
    }
    hidden
}

fn reference_cell_count(range: &CellRange) -> Option<usize> {
    let row_count = range
        .bottom
        .checked_sub(range.top)
        .and_then(|value| value.checked_add(1))?;
    let col_count = range
        .right
        .checked_sub(range.left)
        .and_then(|value| value.checked_add(1))?;
    let total = (row_count as usize).checked_mul(col_count as usize)?;
    Some(total)
}

fn range_contains(ranges: &[(u32, u32)], value: u32) -> bool {
    let insertion = ranges.partition_point(|(_, end)| *end < value);
    ranges
        .get(insertion)
        .is_some_and(|(start, end)| *start <= value && value <= *end)
}

fn extract_from_sparse_cells(
    worksheet: &WorksheetCellLookup,
    range: &CellRange,
) -> Option<(Vec<ReferencedCellValue>, Vec<bool>)> {
    let total = reference_cell_count(range)?;
    if total > MAX_REFERENCE_CELLS {
        return None;
    }
    let col_count = (range.right - range.left + 1) as usize;
    let mut values = vec![ReferencedCellValue::Empty; total];
    let mut hidden = vec![false; total];
    for row in range.top..=range.bottom {
        for col in range.left..=range.right {
            let index = (row - range.top) as usize * col_count + (col - range.left) as usize;
            hidden[index] = range_contains(&worksheet.hidden_row_ranges, row)
                || range_contains(&worksheet.hidden_column_ranges, col);
            let Some(value) = worksheet.cells.get(&(row, col)) else {
                continue;
            };
            values[index] = value.value.clone();
        }
    }
    Some((values, hidden))
}

fn extract_visibility_from_sparse_cells(
    worksheet: &WorksheetCellLookup,
    range: &CellRange,
) -> Option<Vec<bool>> {
    let total = reference_cell_count(range)?;
    if total > MAX_REFERENCE_CELLS {
        return None;
    }
    let col_count = (range.right - range.left + 1) as usize;
    let mut hidden = vec![false; total];
    for row in range.top..=range.bottom {
        let row_hidden = range_contains(&worksheet.hidden_row_ranges, row);
        for col in range.left..=range.right {
            let index = (row - range.top) as usize * col_count + (col - range.left) as usize;
            hidden[index] = row_hidden || range_contains(&worksheet.hidden_column_ranges, col);
        }
    }
    Some(hidden)
}

#[cfg(test)]
pub(crate) fn extract_reference_cells(
    sheet_xml: &str,
    range: &CellRange,
    shared_strings: &[SharedString],
) -> Vec<ReferencedCellValue> {
    if reference_cell_count(range).is_none_or(|total| total > MAX_REFERENCE_CELLS) {
        return Vec::new();
    }
    parse_worksheet_cells(
        sheet_xml,
        shared_strings,
        MAX_TOTAL_INDEXED_CELLS,
        MAX_TOTAL_INDEXED_STRING_BYTES,
    )
    .as_ref()
    .and_then(|worksheet| extract_from_sparse_cells(worksheet, range))
    .map(|(values, _)| values)
    .unwrap_or_default()
}

pub(crate) struct ResolvedWorksheetReference {
    pub(crate) values: Vec<ReferencedCellValue>,
    pub(crate) hidden: Vec<bool>,
}

#[allow(clippy::too_many_arguments)]
fn resolve_worksheet_reference_projected(
    archive: &mut XlsxZip,
    formula: &str,
    current_rows: Option<&[Row]>,
    current_hidden_columns: Option<&BTreeMap<u32, bool>>,
    current_sheet_name: &str,
    sheets: &[SheetMeta],
    workbook_rels: &roxmltree::Document<'_>,
    shared_strings: &[SharedString],
    session: &mut WorksheetReferenceSession,
    include_values: bool,
) -> Option<ResolvedWorksheetReference> {
    let (source_sheet, reference) = split_sheet_ref(formula)?;
    let range = parse_a1_range(&reference)?;
    let sheet_name = source_sheet.as_deref().unwrap_or(current_sheet_name);
    if !session.sheets.contains_key(sheet_name) {
        if sheet_name == current_sheet_name {
            session.resolve_current_sheet(
                sheet_name,
                current_rows,
                current_hidden_columns,
                shared_strings,
            );
        } else {
            let part = sheets
                .iter()
                .find(|sheet| sheet.name == sheet_name)
                .and_then(|sheet| resolve_sheet_path(workbook_rels, &sheet.r_id))
                .map(|path| format!("xl/{path}"));
            let mut build = part
                .as_deref()
                .map_or(LookupBuildResult::Unavailable, |part| {
                    parse_package_worksheet_cells(
                        archive,
                        part,
                        shared_strings,
                        session
                            .remaining_indexed_cells
                            .min(session.remaining_physical_indexed_cells),
                        session
                            .remaining_indexed_string_bytes
                            .min(session.remaining_physical_indexed_string_bytes),
                    )
                });
            if matches!(build, LookupBuildResult::LimitExceeded(_)) {
                if let Some((current_name, prebuilt)) = session.take_unreferenced_prebuilt() {
                    let retried = part
                        .as_deref()
                        .map_or(LookupBuildResult::Unavailable, |part| {
                            parse_package_worksheet_cells(
                                archive,
                                part,
                                shared_strings,
                                session
                                    .remaining_indexed_cells
                                    .min(session.remaining_physical_indexed_cells),
                                session
                                    .remaining_indexed_string_bytes
                                    .min(session.remaining_physical_indexed_string_bytes),
                            )
                        });
                    if let LookupBuildResult::Built(worksheet) = &retried {
                        let remaining_cells =
                            session.remaining_indexed_cells - worksheet.resource_units();
                        let remaining_bytes =
                            session.remaining_indexed_string_bytes - worksheet.string_bytes;
                        let failure = if prebuilt.resource_units() > remaining_cells {
                            Some((
                                HardResourceLimitKind::WorksheetReferenceIndexCells,
                                remaining_cells,
                                prebuilt.resource_units(),
                            ))
                        } else if prebuilt.string_bytes > remaining_bytes {
                            Some((
                                HardResourceLimitKind::WorksheetReferenceIndexStringBytes,
                                remaining_bytes,
                                prebuilt.string_bytes,
                            ))
                        } else {
                            None
                        };
                        // A streamed index can peak above its final size when
                        // duplicate cells replace/remove earlier values. Retry
                        // admission alone does not prove both final indexes
                        // cannot coexist; retain the current sheet when they fit.
                        if let Some(failure) = failure {
                            session.current_sheet = CurrentSheetLookupEntry::Unavailable {
                                name: current_name,
                                limit_failure: Some(failure),
                            };
                        } else {
                            session.restore_unreferenced_prebuilt(current_name, prebuilt);
                        }
                    } else {
                        session.restore_unreferenced_prebuilt(current_name, prebuilt);
                    }
                    build = retried;
                }
            }
            let worksheet = match build {
                LookupBuildResult::Built(worksheet) => {
                    session.consume_new_index(&worksheet);
                    Some(worksheet)
                }
                LookupBuildResult::LimitExceeded(failure) => {
                    report_index_limit(failure);
                    None
                }
                LookupBuildResult::Unavailable => None,
            };
            session.sheets.insert(sheet_name.to_string(), worksheet);
        }
    }
    // Only a resolved source can produce dense output. This is library resource
    // policy, not formula validation (ECMA-376 CT_NumRef/CT_StrRef f is a string).
    // Missing sheets/relationships and unreadable parts keep content fallback,
    // even for oversized ranges or after earlier results exhausted the budget.
    let worksheet = session.sheets.get(sheet_name).and_then(Option::as_ref)?;
    let cell_count = session.reservable_cell_count(&range)?;
    // Count only resolved cells in the admitted range, without cloning strings
    // or scanning unrelated index entries. Both output budgets are admitted
    // before allocating the dense vectors or any duplicated string payload.
    let mut string_bytes = 0usize;
    if include_values {
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                if let Some(cell) = worksheet.cells.get(&(row, col)) {
                    string_bytes = string_bytes.saturating_add(cell.value.string_bytes());
                }
            }
        }
    }
    reference_limit(
        HardResourceLimitKind::WorksheetReferenceStringBytes,
        MAX_TOTAL_REFERENCE_STRING_BYTES,
        (MAX_TOTAL_REFERENCE_STRING_BYTES - session.remaining_string_bytes)
            .saturating_add(string_bytes),
    )?;
    let (values, hidden) = if include_values {
        extract_from_sparse_cells(worksheet, &range)?
    } else {
        (
            Vec::new(),
            extract_visibility_from_sparse_cells(worksheet, &range)?,
        )
    };
    session.consume_result(cell_count, string_bytes);
    Some(ResolvedWorksheetReference { values, hidden })
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn resolve_worksheet_reference_with_visibility(
    archive: &mut XlsxZip,
    formula: &str,
    current_rows: Option<&[Row]>,
    current_hidden_columns: Option<&BTreeMap<u32, bool>>,
    current_sheet_name: &str,
    sheets: &[SheetMeta],
    workbook_rels: &roxmltree::Document<'_>,
    shared_strings: &[SharedString],
    session: &mut WorksheetReferenceSession,
) -> Option<ResolvedWorksheetReference> {
    resolve_worksheet_reference_projected(
        archive,
        formula,
        current_rows,
        current_hidden_columns,
        current_sheet_name,
        sheets,
        workbook_rels,
        shared_strings,
        session,
        true,
    )
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn resolve_worksheet_visibility(
    archive: &mut XlsxZip,
    formula: &str,
    current_rows: Option<&[Row]>,
    current_hidden_columns: Option<&BTreeMap<u32, bool>>,
    current_sheet_name: &str,
    sheets: &[SheetMeta],
    workbook_rels: &roxmltree::Document<'_>,
    shared_strings: &[SharedString],
    session: &mut WorksheetReferenceSession,
) -> Option<Vec<bool>> {
    resolve_worksheet_reference_projected(
        archive,
        formula,
        current_rows,
        current_hidden_columns,
        current_sheet_name,
        sheets,
        workbook_rels,
        shared_strings,
        session,
        false,
    )
    .map(|resolved| resolved.hidden)
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn resolve_worksheet_reference(
    archive: &mut XlsxZip,
    formula: &str,
    current_rows: Option<&[Row]>,
    current_sheet_name: &str,
    sheets: &[SheetMeta],
    workbook_rels: &roxmltree::Document<'_>,
    shared_strings: &[SharedString],
    session: &mut WorksheetReferenceSession,
) -> Option<Vec<ReferencedCellValue>> {
    resolve_worksheet_reference_with_visibility(
        archive,
        formula,
        current_rows,
        None,
        current_sheet_name,
        sheets,
        workbook_rels,
        shared_strings,
        session,
    )
    .map(|resolved| resolved.values)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{parse_worksheet, CellRange, SharedString, SheetVisibility};
    use std::{fmt::Write as _, io::Write as _};
    use zip::write::SimpleFileOptions;

    fn reduced_session(index_cells: usize, string_bytes: usize) -> WorksheetReferenceSession {
        WorksheetReferenceSession {
            remaining_indexed_cells: index_cells,
            remaining_indexed_string_bytes: string_bytes,
            remaining_physical_indexed_cells: index_cells,
            remaining_physical_indexed_string_bytes: string_bytes,
            ..Default::default()
        }
    }

    fn one_sheet_reference_archive(
        values: &[f64],
    ) -> (XlsxZip, roxmltree::Document<'static>, Vec<SheetMeta>) {
        let cells = values
            .iter()
            .enumerate()
            .map(|(index, value)| {
                format!(
                    r#"<c r="{}1"><v>{value}</v></c>"#,
                    (b'A' + index as u8) as char
                )
            })
            .collect::<String>();
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">{cells}</row></sheetData></worksheet>"#
        );
        one_sheet_xml_reference_archive(&xml)
    }

    fn one_sheet_xml_reference_archive(
        xml: &str,
    ) -> (XlsxZip, roxmltree::Document<'static>, Vec<SheetMeta>) {
        let cursor = {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            writer
                .start_file("xl/worksheets/source.xml", SimpleFileOptions::default())
                .unwrap();
            writer.write_all(xml.as_bytes()).unwrap();
            writer.finish().unwrap()
        };
        let rels = ooxml_common::depth::parse_guarded(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSource" Target="worksheets/source.xml"/></Relationships>"#,
        ).unwrap();
        let sheets = vec![SheetMeta {
            name: "Source".into(),
            sheet_id: 2,
            r_id: "rSource".into(),
            tab_color: None,
            visibility: SheetVisibility::Visible,
        }];
        (XlsxZip::new(cursor).unwrap(), rels, sheets)
    }

    fn one_sheet_text_reference_archive(
        text: &str,
    ) -> (XlsxZip, roxmltree::Document<'static>, Vec<SheetMeta>) {
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{text}</t></is></c></row></sheetData></worksheet>"#
        );
        let cursor = {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            writer
                .start_file("xl/worksheets/source.xml", SimpleFileOptions::default())
                .unwrap();
            writer.write_all(xml.as_bytes()).unwrap();
            writer.finish().unwrap()
        };
        let rels = ooxml_common::depth::parse_guarded(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSource" Target="worksheets/source.xml"/></Relationships>"#,
        ).unwrap();
        let sheets = vec![SheetMeta {
            name: "Source".into(),
            sheet_id: 2,
            r_id: "rSource".into(),
            tab_color: None,
            visibility: SheetVisibility::Visible,
        }];
        (XlsxZip::new(cursor).unwrap(), rels, sheets)
    }

    #[test]
    fn quoted_unicode_sheet_reference_is_split_and_unescaped() {
        assert_eq!(
            split_sheet_ref("'التقرير'!$A$2:$A$5"),
            Some((Some("التقرير".into()), "A2:A5".into())),
        );
        assert_eq!(
            split_sheet_ref("'Bob''s data'!C1"),
            Some((Some("Bob's data".into()), "C1".into())),
        );
        assert_eq!(split_sheet_ref("A1:A3"), Some((None, "A1:A3".into())));
    }

    #[test]
    fn direct_a1_range_is_normalized() {
        let range = parse_a1_range("C5:A2").expect("direct A1 range parses");
        assert_eq!(
            (range.top, range.left, range.bottom, range.right),
            (2, 1, 5, 3),
        );
        assert!(parse_a1_range("Sales").is_none());
        assert!(parse_a1_range("A1+1").is_none());
    }

    #[test]
    fn worksheet_cells_resolve_inline_shared_formula_string_and_numbers() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Inline</t></is></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="str"><f>UPPER(&quot;x&quot;)</f><v>X</v></c><c r="D1"><v>42</v></c></row></sheetData></worksheet>"#;
        let shared = vec![SharedString {
            text: "Shared".into(),
            ..Default::default()
        }];
        let range = CellRange {
            top: 1,
            left: 1,
            bottom: 1,
            right: 4,
        };
        assert_eq!(
            extract_reference_cells(xml, &range, &shared),
            vec![
                ReferencedCellValue::Text("Inline".into()),
                ReferencedCellValue::Text("Shared".into()),
                ReferencedCellValue::Text("X".into()),
                ReferencedCellValue::Number(42.0),
            ],
        );
    }

    #[test]
    fn parsed_worksheet_lazy_index_matches_streamed_projection_without_reparsing() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="2" max="2" hidden="1"/></cols><sheetData>
          <row r="2"><c r="A2" t="inlineStr"><is><t>Inline</t><rPh sb="0" eb="1"><t>reading</t></rPh></is></c><c r="B2" t="s"><v>0</v></c><c r="C2"><v>42</v></c><c r="D2" t="b"><v>1</v></c></row>
          <row r="3" hidden="1"><c t="str"><v>cached</v></c><c t="e"><v>#N/A</v></c></row>
        </sheetData></worksheet>"#;
        let shared = vec![SharedString {
            text: "Shared".into(),
            runs: None,
            phonetic_runs: Vec::new(),
            phonetic_pr: None,
        }];
        let (worksheet, _) =
            parse_worksheet(xml, &shared, &[], "Sheet1").expect("worksheet parses");
        let expected = parse_worksheet_cells(
            xml,
            &shared,
            MAX_TOTAL_INDEXED_CELLS,
            MAX_TOTAL_INDEXED_STRING_BYTES,
        )
        .expect("streamed index parses");

        let session = WorksheetReferenceSession::default();
        assert!(session.sheets.is_empty(), "no eager index is retained");
        let LookupBuildResult::Built(actual) =
            session.index_materialized_rows(&worksheet.rows, Some(&worksheet.col_hidden), &shared)
        else {
            panic!("parsed worksheet index fits");
        };
        assert!(
            session.sheets.is_empty(),
            "building an index does not cache it before a formula needs it"
        );

        assert_eq!(actual.cells, expected.cells);
        assert_eq!(actual.string_bytes, expected.string_bytes);

        let zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()))
            .finish()
            .expect("empty zip finishes");
        let mut archive = XlsxZip::new(zip).expect("empty zip opens");
        let rels = ooxml_common::depth::parse_guarded(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
        )
        .expect("empty workbook relationships parse");
        let mut lazy_session = WorksheetReferenceSession::default();
        assert!(lazy_session.sheets.is_empty());
        lazy_session.seed_current_sheet("Sheet1", Some(actual));
        let resolved = resolve_worksheet_reference_with_visibility(
            &mut archive,
            "A2:C3",
            Some(&worksheet.rows),
            Some(&worksheet.col_hidden),
            "Sheet1",
            &[],
            &rels,
            &shared,
            &mut lazy_session,
        )
        .expect("current-sheet reference resolves from parsed worksheet");
        assert_eq!(
            resolved.values,
            vec![
                ReferencedCellValue::Text("Inline".into()),
                ReferencedCellValue::Text("Shared".into()),
                ReferencedCellValue::Number(42.0),
                ReferencedCellValue::Text("cached".into()),
                ReferencedCellValue::Empty,
                ReferencedCellValue::Empty,
            ]
        );
        assert_eq!(
            resolved.hidden,
            vec![false, true, false, true, true, true],
            "materialized hidden columns and an empty hidden row retain provenance"
        );
        assert_eq!(
            lazy_session.sheets.len(),
            1,
            "the index appears only after a formula is resolved"
        );
    }

    #[test]
    fn oversized_range_is_rejected_before_allocation() {
        let range = CellRange {
            top: 1,
            left: 1,
            bottom: 1_048_576,
            right: 16_384,
        };
        assert!(extract_reference_cells("<worksheet/>", &range, &[]).is_empty());
    }

    #[test]
    fn reference_cell_and_string_crossings_are_typed_package_failures() {
        for strings in [false, true] {
            let (mut archive, rels, sheets) = one_sheet_text_reference_archive("abc");
            let mut session = WorksheetReferenceSession {
                remaining_cells: if strings {
                    MAX_TOTAL_REFERENCE_CELLS
                } else {
                    0
                },
                remaining_string_bytes: 2,
                ..Default::default()
            };
            let error = archive
                .run_operation("reference-test", |archive| {
                    let _provisional = resolve_worksheet_reference(
                        archive,
                        "Source!A1",
                        None,
                        "Dashboard",
                        &sheets,
                        &rels,
                        &[],
                        &mut session,
                    );
                    Ok(())
                })
                .unwrap_err();
            let payload: serde_json::Value =
                serde_json::from_str(error.strip_prefix("OOXML_RESOURCE_LIMIT:").expect(&error))
                    .unwrap();
            assert_eq!(
                payload["details"]["violation"]["resource"],
                "worksheet-reference"
            );
            assert_eq!(
                payload["details"]["violation"]["metric"],
                if strings { "bytes" } else { "cells" }
            );
        }
    }

    #[test]
    fn unavailable_sources_do_not_admit_dense_output_even_after_budget_exhaustion() {
        for formula in ["Missing!A1:A1000001", "Source!A1:A1000001", "Source!A1"] {
            let (mut archive, rels, sheets) = one_sheet_xml_reference_archive("<worksheet>");
            let mut session = WorksheetReferenceSession {
                remaining_cells: 0,
                remaining_string_bytes: 0,
                ..Default::default()
            };
            archive
                .run_operation("unavailable-source-test", |archive| {
                    for include_values in [false, true] {
                        assert!(resolve_worksheet_reference_projected(
                            archive,
                            formula,
                            None,
                            None,
                            "Dashboard",
                            &sheets,
                            &rels,
                            &[],
                            &mut session,
                            include_values,
                        )
                        .is_none());
                    }
                    Ok(())
                })
                .expect("missing and unreadable sources have no dense output to admit");
        }
    }

    #[test]
    fn malformed_source_tail_does_not_turn_an_index_refusal_into_a_package_failure() {
        // Cross the 512-row projection batch so index admission runs before
        // the projector discovers the malformed tail in the next batch.
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{}</sheetData><bad / ></worksheet>"#,
            r#"<row><c><v>1</v></c><c><v>2</v></c></row>"#.repeat(513),
        );
        let (mut archive, rels, sheets) = one_sheet_xml_reference_archive(&xml);
        let mut session = reduced_session(1, 8);
        archive
            .run_operation("malformed-source-test", |archive| {
                assert!(resolve_worksheet_reference(
                    archive,
                    "Source!A1:B1",
                    None,
                    "Dashboard",
                    &sheets,
                    &rels,
                    &[],
                    &mut session,
                )
                .is_none());
                Ok(())
            })
            .expect("a failed source parse keeps its existing unavailable fallback");
    }

    #[test]
    fn unused_prebuilt_index_can_be_evicted_without_poisoning_complete_references() {
        for strings in [false, true] {
            for request_evicted_sheet in [false, true] {
                let (mut archive, rels, sheets) = if strings {
                    one_sheet_text_reference_archive("yz")
                } else {
                    one_sheet_reference_archive(&[7.0, 8.0])
                };
                let mut builder = WorksheetCellLookupBuilder::new(2, 2);
                builder
                    .push_cell(
                        1,
                        1,
                        if strings {
                            ReferencedCellValue::Text("x".into())
                        } else {
                            ReferencedCellValue::Number(1.0)
                        },
                    )
                    .unwrap();
                let mut session = reduced_session(2, 2);
                session.seed_current_sheet("Dashboard", Some(builder.finish()));
                let result = archive.run_operation("index-eviction-test", |archive| {
                    let resolved = resolve_worksheet_reference(
                        archive,
                        if strings { "Source!A1" } else { "Source!A1:B1" },
                        None,
                        "Dashboard",
                        &sheets,
                        &rels,
                        &[],
                        &mut session,
                    );
                    assert!(
                        resolved.is_some(),
                        "eviction admits the complete cross-sheet reference"
                    );
                    if request_evicted_sheet {
                        let _provisional = resolve_worksheet_reference(
                            archive,
                            "Dashboard!A1",
                            None,
                            "Dashboard",
                            &sheets,
                            &rels,
                            &[],
                            &mut session,
                        );
                    }
                    Ok(())
                });
                if request_evicted_sheet {
                    let error = result.unwrap_err();
                    assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"), "{error}");
                } else {
                    result.unwrap();
                }
            }
        }
    }

    #[test]
    fn successful_retry_restores_current_index_when_source_shrinks_to_fit() {
        for strings in [false, true] {
            let cells = if strings {
                r#"<c r="A1" t="inlineStr"><is><t>x</t></is></c><c r="B1" t="inlineStr"><is><t>y</t></is></c><c r="A1"/>"#
            } else {
                r#"<c r="A1"><v>7</v></c><c r="B1"><v>8</v></c><c r="A1"/>"#
            };
            let xml = format!(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">{cells}</row></sheetData></worksheet>"#
            );
            let (mut archive, rels, sheets) = one_sheet_xml_reference_archive(&xml);
            let mut current = WorksheetCellLookupBuilder::new(1, 1);
            let value = if strings {
                ReferencedCellValue::Text("z".into())
            } else {
                ReferencedCellValue::Number(1.0)
            };
            current.push_cell(1, 1, value.clone()).unwrap();
            // Isolate cell and string-byte admission: the source peaks at two
            // units but retains one after the duplicate A1 removes its value.
            let mut session = reduced_session(if strings { 3 } else { 2 }, 2);
            session.seed_current_sheet("Dashboard", Some(current.finish()));
            archive
                .run_operation("index-retry-test", |archive| {
                    assert_eq!(
                        resolve_worksheet_reference(
                            archive,
                            "Source!A1:B1",
                            None,
                            "Dashboard",
                            &sheets,
                            &rels,
                            &[],
                            &mut session,
                        ),
                        Some(vec![
                            ReferencedCellValue::Empty,
                            if strings {
                                ReferencedCellValue::Text("y".into())
                            } else {
                                ReferencedCellValue::Number(8.0)
                            }
                        ])
                    );
                    assert_eq!(
                        resolve_worksheet_reference(
                            archive,
                            "Dashboard!A1",
                            None,
                            "Dashboard",
                            &sheets,
                            &rels,
                            &[],
                            &mut session,
                        ),
                        Some(vec![value])
                    );
                    Ok(())
                })
                .expect("both final indexes fit without a resource crossing");
        }
    }

    #[test]
    fn failed_streamed_index_is_terminal_only_when_a_reference_needs_it() {
        let (mut archive, rels, sheets) = one_sheet_reference_archive(&[1.0]);
        let mut builder = WorksheetCellLookupBuilder::new(0, 0);
        assert!(builder
            .push_cell(1, 1, ReferencedCellValue::Number(1.0))
            .is_none());
        let mut session = WorksheetReferenceSession::default();
        session.seed_current_sheet_failure("Dashboard", builder.limit_failure().unwrap());
        let error = archive
            .run_operation("streamed-index-test", |archive| {
                let _provisional = resolve_worksheet_reference(
                    archive,
                    "Dashboard!A1",
                    None,
                    "Dashboard",
                    &sheets,
                    &rels,
                    &[],
                    &mut session,
                );
                Ok(())
            })
            .unwrap_err();
        assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"), "{error}");
    }

    #[test]
    fn reference_index_cell_and_string_crossings_are_typed_package_failures() {
        for strings in [false, true] {
            let (mut archive, rels, sheets) = one_sheet_text_reference_archive("abc");
            let mut session =
                reduced_session(if strings { 1 } else { 0 }, if strings { 2 } else { 3 });
            let error = archive
                .run_operation("reference-index-test", |archive| {
                    let _provisional = resolve_worksheet_reference(
                        archive,
                        "Source!A1",
                        None,
                        "Dashboard",
                        &sheets,
                        &rels,
                        &[],
                        &mut session,
                    );
                    Ok(())
                })
                .unwrap_err();
            let payload: serde_json::Value =
                serde_json::from_str(error.strip_prefix("OOXML_RESOURCE_LIMIT:").expect(&error))
                    .unwrap();
            assert_eq!(
                payload["details"]["violation"]["resource"],
                "worksheet-reference-index"
            );
            assert_eq!(
                payload["details"]["violation"]["metric"],
                if strings { "bytes" } else { "cells" }
            );
        }
    }

    #[test]
    fn cumulative_budget_rejects_individually_valid_ranges() {
        let mut session = WorksheetReferenceSession {
            remaining_cells: 4,
            ..Default::default()
        };
        let two_cells = CellRange {
            top: 1,
            left: 1,
            bottom: 1,
            right: 2,
        };
        let three_cells = CellRange {
            top: 1,
            left: 1,
            bottom: 1,
            right: 3,
        };

        let first_count = session.reservable_cell_count(&two_cells).unwrap();
        session.consume_result(first_count, 0);
        assert!(session.reservable_cell_count(&three_cells).is_none());
        assert_eq!(session.remaining_cells, 2);
    }

    #[test]
    fn sparse_index_has_cell_and_string_budgets() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"/><c r="B1" t="inlineStr"><is><t>alpha</t></is></c><c r="C1"><v>42</v></c></row></sheetData></worksheet>"#;

        let indexed = parse_worksheet_cells(xml, &[], 2, 5).expect("two non-empty cells fit");
        assert_eq!(indexed.cells.len(), 2);
        assert_eq!(indexed.string_bytes, 5);
        assert!(!indexed.cells.contains_key(&(1, 1)));
        assert!(parse_worksheet_cells(xml, &[], 1, 5).is_none());
        assert!(parse_worksheet_cells(xml, &[], 2, 4).is_none());
    }

    #[test]
    fn hidden_row_ranges_are_normalized_even_when_row_records_are_out_of_order() {
        let mut builder = WorksheetCellLookupBuilder::new(4, 0);
        builder.push_hidden_row(10).unwrap();
        builder.push_hidden_row(3).unwrap();
        builder.push_hidden_row(2).unwrap();
        let lookup = builder.finish();
        assert_eq!(lookup.hidden_row_ranges, vec![(2, 3), (10, 10)]);
    }

    #[test]
    fn duplicate_coordinate_empty_removes_previous_value_and_budget() {
        let mut builder = WorksheetCellLookupBuilder::new(1, 5);
        builder
            .push_cell(1, 1, ReferencedCellValue::Text("alpha".into()))
            .unwrap();
        builder.push_cell(1, 1, ReferencedCellValue::Empty).unwrap();
        let lookup = builder.finish();
        assert!(lookup.cells.is_empty());
        assert_eq!(lookup.string_bytes, 0);
    }

    #[test]
    fn duplicate_coordinate_nonempty_replaces_value_and_budget() {
        let mut builder = WorksheetCellLookupBuilder::new(1, 4);
        builder
            .push_cell(1, 1, ReferencedCellValue::Text("long".into()))
            .unwrap();
        builder
            .push_cell(1, 1, ReferencedCellValue::Text("x".into()))
            .unwrap();
        let lookup = builder.finish();
        assert_eq!(
            lookup.cells.get(&(1, 1)).map(|cell| &cell.value),
            Some(&ReferencedCellValue::Text("x".into()))
        );
        assert_eq!(lookup.string_bytes, 1);
    }

    #[test]
    fn lookup_cap_discards_the_partial_transaction() {
        let rows = vec![Row {
            index: 1,
            height: None,
            custom_height: false,
            cells: vec![
                crate::Cell {
                    col: 1,
                    row: 1,
                    value: CellValue::Number { number: 1.0 },
                    ..Default::default()
                },
                crate::Cell {
                    col: 2,
                    row: 1,
                    value: CellValue::Number { number: 2.0 },
                    ..Default::default()
                },
            ],
            outline_level: 0,
            collapsed: false,
            hidden: false,
        }];
        let mut lookup = Some(WorksheetCellLookupBuilder::new(1, 100));
        let _ = extend_lookup_transactionally(&mut lookup, &rows, &[]);
        assert!(lookup.is_none());
    }

    #[test]
    fn uncharged_prebuilt_current_lookup_does_not_starve_cross_sheet_only_formula() {
        let (mut archive, rels, sheets) = one_sheet_reference_archive(&[7.0]);
        let mut prebuilt = WorksheetCellLookupBuilder::new(1, 8);
        prebuilt
            .push_cell(1, 1, ReferencedCellValue::Number(1.0))
            .unwrap();
        let mut session = reduced_session(1, 8);
        session.seed_current_sheet("Dashboard", Some(prebuilt.finish()));
        let values = resolve_worksheet_reference(
            &mut archive,
            "Source!A1",
            None,
            "Dashboard",
            &sheets,
            &rels,
            &[],
            &mut session,
        )
        .expect("exclusive cross-sheet reference gets the physical budget");
        assert_eq!(values, vec![ReferencedCellValue::Number(7.0)]);
        assert!(matches!(
            session.current_sheet,
            CurrentSheetLookupEntry::Unavailable { .. }
        ));
    }

    #[test]
    fn mixed_current_cross_order_is_deterministic_under_combined_physical_cap() {
        let make_prebuilt = || {
            let mut builder = WorksheetCellLookupBuilder::new(1, 8);
            builder
                .push_cell(1, 1, ReferencedCellValue::Number(1.0))
                .unwrap();
            builder.finish()
        };

        let (mut cross_first_archive, rels, sheets) = one_sheet_reference_archive(&[7.0, 8.0]);
        let mut cross_first = reduced_session(2, 8);
        cross_first.seed_current_sheet("Dashboard", Some(make_prebuilt()));
        assert!(resolve_worksheet_reference(
            &mut cross_first_archive,
            "Source!A1:B1",
            None,
            "Dashboard",
            &sheets,
            &rels,
            &[],
            &mut cross_first,
        )
        .is_some());
        assert!(resolve_worksheet_reference(
            &mut cross_first_archive,
            "Dashboard!A1",
            None,
            "Dashboard",
            &sheets,
            &rels,
            &[],
            &mut cross_first,
        )
        .is_none());

        let (mut current_first_archive, rels, sheets) = one_sheet_reference_archive(&[7.0, 8.0]);
        let mut current_first = reduced_session(2, 8);
        current_first.seed_current_sheet("Dashboard", Some(make_prebuilt()));
        assert!(resolve_worksheet_reference(
            &mut current_first_archive,
            "Dashboard!A1",
            None,
            "Dashboard",
            &sheets,
            &rels,
            &[],
            &mut current_first,
        )
        .is_some());
        assert!(resolve_worksheet_reference(
            &mut current_first_archive,
            "Source!A1:B1",
            None,
            "Dashboard",
            &sheets,
            &rels,
            &[],
            &mut current_first,
        )
        .is_none());
        assert_eq!(current_first.remaining_physical_indexed_cells, 1);
    }

    #[test]
    fn failed_cross_cell_retry_restores_prebuilt_current_and_exact_budgets() {
        let (mut archive, rels, sheets) = one_sheet_reference_archive(&[7.0, 8.0]);
        let mut current = WorksheetCellLookupBuilder::new(1, 8);
        current
            .push_cell(1, 1, ReferencedCellValue::Number(1.0))
            .unwrap();
        let mut prior = WorksheetCellLookupBuilder::new(9, 8);
        for col in 1..=9 {
            prior
                .push_cell(1, col, ReferencedCellValue::Number(col.into()))
                .unwrap();
        }

        let mut session = reduced_session(10, 8);
        session.seed_current_sheet("Dashboard", Some(current.finish()));
        let prior = prior.finish();
        session.consume_new_index(&prior);
        session.sheets.insert("Prior".into(), Some(prior));
        let before = (
            session.remaining_indexed_cells,
            session.remaining_indexed_string_bytes,
            session.remaining_physical_indexed_cells,
            session.remaining_physical_indexed_string_bytes,
        );

        assert!(resolve_worksheet_reference(
            &mut archive,
            "Source!A1:B1",
            None,
            "Dashboard",
            &sheets,
            &rels,
            &[],
            &mut session,
        )
        .is_none());
        assert_eq!(
            (
                session.remaining_indexed_cells,
                session.remaining_indexed_string_bytes,
                session.remaining_physical_indexed_cells,
                session.remaining_physical_indexed_string_bytes,
            ),
            before,
            "a failed retry must restore both logical and physical accounting exactly"
        );
        assert_eq!(
            resolve_worksheet_reference(
                &mut archive,
                "Dashboard!A1",
                None,
                "Dashboard",
                &sheets,
                &rels,
                &[],
                &mut session,
            ),
            Some(vec![ReferencedCellValue::Number(1.0)])
        );
        assert_eq!(session.remaining_indexed_cells, 0);
        assert_eq!(session.remaining_physical_indexed_cells, 0);
    }

    #[test]
    fn failed_cross_string_retry_restores_prebuilt_current_and_exact_budgets() {
        let (mut archive, rels, sheets) = one_sheet_text_reference_archive("yz");
        let mut current = WorksheetCellLookupBuilder::new(1, 1);
        current
            .push_cell(1, 1, ReferencedCellValue::Text("x".into()))
            .unwrap();
        let mut prior = WorksheetCellLookupBuilder::new(1, 9);
        prior
            .push_cell(1, 1, ReferencedCellValue::Text("123456789".into()))
            .unwrap();

        let mut session = reduced_session(10, 10);
        session.seed_current_sheet("Dashboard", Some(current.finish()));
        let prior = prior.finish();
        session.consume_new_index(&prior);
        session.sheets.insert("Prior".into(), Some(prior));
        let before = (
            session.remaining_indexed_cells,
            session.remaining_indexed_string_bytes,
            session.remaining_physical_indexed_cells,
            session.remaining_physical_indexed_string_bytes,
        );

        assert!(resolve_worksheet_reference(
            &mut archive,
            "Source!A1",
            None,
            "Dashboard",
            &sheets,
            &rels,
            &[],
            &mut session,
        )
        .is_none());
        assert_eq!(
            (
                session.remaining_indexed_cells,
                session.remaining_indexed_string_bytes,
                session.remaining_physical_indexed_cells,
                session.remaining_physical_indexed_string_bytes,
            ),
            before,
            "string-byte retry failure must not double-charge or underflow"
        );
        assert_eq!(
            resolve_worksheet_reference(
                &mut archive,
                "Dashboard!A1",
                None,
                "Dashboard",
                &sheets,
                &rels,
                &[],
                &mut session,
            ),
            Some(vec![ReferencedCellValue::Text("x".into())])
        );
        assert_eq!(session.remaining_indexed_string_bytes, 0);
        assert_eq!(session.remaining_physical_indexed_string_bytes, 0);
    }

    #[test]
    fn inline_string_excludes_phonetic_guidance() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh></is></c></row></sheetData></worksheet>"#;
        let range = CellRange {
            top: 1,
            left: 1,
            bottom: 1,
            right: 1,
        };

        assert_eq!(
            extract_reference_cells(xml, &range, &[]),
            vec![ReferencedCellValue::Text("漢字".into())]
        );
    }

    #[test]
    fn streamed_projection_matches_primary_mce_row_selection() {
        let xml = r#"
          <x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:not-understood">
            <x:sheetData>
              <mc:AlternateContent>
                <mc:Choice Requires="future">
                  <x:row r="1"><x:c r="A1"><x:v>999</x:v></x:c></x:row>
                </mc:Choice>
                <mc:Fallback>
                  <x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>fallback</x:t></x:is></x:c></x:row>
                </mc:Fallback>
              </mc:AlternateContent>
              <x:row><x:c t="str"><x:v>implicit</x:v></x:c></x:row>
            </x:sheetData>
          </x:worksheet>
        "#;
        let (worksheet, _) =
            parse_worksheet(xml, &[], &[], "Sheet1").expect("primary worksheet parses MCE");
        let session = WorksheetReferenceSession::default();
        let LookupBuildResult::Built(expected) =
            session.index_materialized_rows(&worksheet.rows, None, &[])
        else {
            panic!("primary model index fits");
        };
        let actual = parse_worksheet_cells(
            xml,
            &[],
            MAX_TOTAL_INDEXED_CELLS,
            MAX_TOTAL_INDEXED_STRING_BYTES,
        )
        .expect("streamed reference projection parses MCE");

        assert_eq!(actual.cells, expected.cells);
        assert_eq!(actual.string_bytes, expected.string_bytes);
        assert_eq!(
            actual.cells.get(&(1, 1)).map(|cell| &cell.value),
            Some(&ReferencedCellValue::Text("fallback".into()))
        );
        assert_eq!(
            actual.cells.get(&(2, 1)).map(|cell| &cell.value),
            Some(&ReferencedCellValue::Text("implicit".into()))
        );
    }

    #[test]
    fn cross_sheet_dense_reference_uses_bounded_streamed_projection() {
        const ROWS: usize = 4_000;
        const COLS: usize = 39;
        let mut source = String::with_capacity(ROWS * COLS * 18);
        source.push_str(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
        );
        for row in 0..ROWS {
            write!(source, r#"<row r="{}">"#, row + 1).unwrap();
            for col in 0..COLS {
                write!(source, "<c><v>{}</v></c>", row * COLS + col).unwrap();
            }
            source.push_str("</row>");
        }
        source.push_str("</sheetData></worksheet>");

        let cursor = {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            writer
                .start_file("xl/worksheets/source.xml", SimpleFileOptions::default())
                .unwrap();
            writer.write_all(source.as_bytes()).unwrap();
            writer.finish().unwrap()
        };
        let mut archive = XlsxZip::new(cursor).expect("synthetic archive opens");
        let rels = ooxml_common::depth::parse_guarded(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSource" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/source.xml"/></Relationships>"#,
        )
        .expect("workbook relationships parse");
        let sheets = vec![SheetMeta {
            name: "Source".into(),
            sheet_id: 2,
            r_id: "rSource".into(),
            tab_color: None,
            visibility: SheetVisibility::Visible,
        }];
        let mut session = WorksheetReferenceSession::default();

        let values = resolve_worksheet_reference(
            &mut archive,
            "'Source'!A1:AM4000",
            None,
            "Dashboard",
            &sheets,
            &rels,
            &[],
            &mut session,
        )
        .expect("dense cross-sheet range resolves");

        assert_eq!(values.len(), ROWS * COLS);
        assert_eq!(values.first(), Some(&ReferencedCellValue::Number(0.0)));
        assert_eq!(
            values.last(),
            Some(&ReferencedCellValue::Number((ROWS * COLS - 1) as f64))
        );
        let indexed = session
            .sheets
            .get("Source")
            .and_then(Option::as_ref)
            .expect("source index is cached");
        assert_eq!(indexed.cells.len(), ROWS * COLS);
    }
}

#[cfg(test)]
mod hidden_column_shell_tests {
    use super::hidden_columns_from_shell;

    #[test]
    fn overlapping_hidden_declarations_mark_each_column_once() {
        let shell = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>{}<col min="20" max="30" hidden="1"/><col min="5" max="9" hidden="true"/><col min="40" hidden="1"/><col min="50" max="60"/></cols><sheetData/></worksheet>"#,
            r#"<col min="1" max="3" hidden="1"/>"#.repeat(3)
        );
        let hidden: Vec<u32> = hidden_columns_from_shell(&shell).into_keys().collect();
        let expected: Vec<u32> = (1..=3).chain(5..=9).chain(20..=30).chain([40]).collect();
        assert_eq!(hidden, expected);
    }
}
