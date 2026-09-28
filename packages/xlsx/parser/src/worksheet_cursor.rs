//! Production lifecycle for resumable worksheet-row projection.
//!
//! The cursor owns the worksheet entry stream and projector between pulls. The
//! surrounding [`XlsxZip`](crate::XlsxZip) continues to own the single logical
//! package operation so workbook dependencies, every row pull, and the sheet's
//! ancillary parts are charged to the same operation.

use std::borrow::Cow;
#[cfg(test)]
use std::collections::BTreeMap;
use std::io::{BufRead, BufReader, Cursor, Read};
use std::rc::Rc;

#[cfg(test)]
use ooxml_common::bounded_xml::BoundedXmlReader;
use ooxml_common::bounded_xml::MCE_NS;
use ooxml_common::ns::is_x_ns;
#[cfg(test)]
use quick_xml::events::{BytesStart, Event};

use crate::worksheet_projector::{
    authored_row_height, AuthoredRowGeometry, ProjectedWorksheetRow, WorksheetProjectorItem,
    WorksheetRowProjector,
};
use crate::{
    parse_cell_ref_checked, resolve_implicit_ordinal, xml_bool_value, Row, SharedString,
    SpreadsheetOrdinal, XlsxZip,
};

/// Default semantic credit for one production pull. Rows are indivisible: the
/// cursor never splits a row to satisfy this credit.
pub(super) const WORKSHEET_CURSOR_PULL_ROWS: usize = 128;
pub(super) const WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES: usize = 1024 * 1024;

type ProductionProjector =
    WorksheetRowProjector<std::io::BufReader<Box<dyn Read>>, Rc<[SharedString]>, Rc<[String]>>;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum WorksheetCursorState {
    Open,
    Finished,
    Failed,
    Canceled,
    Closed,
}

#[derive(Debug)]
pub(super) struct WorksheetCursorTail {
    pub(super) shell_xml: String,
    pub(super) row_geometry: AuthoredRowGeometry,
}

/// Facts that the viewer needs before accepting an exact provisional frame.
/// The first inflation retains only the row-free shell and row geometry; the
/// ordinary cursor opens the compressed entry again for bounded row pulls.
pub(super) struct WorksheetCursorPreview {
    pub(super) tail: Option<WorksheetCursorTail>,
    pub(super) max_row: u32,
    pub(super) max_col: u32,
    pub(super) has_row_outline: bool,
    pub(super) ordered_rows: bool,
}

#[derive(Debug)]
pub(super) enum WorksheetCursorPull {
    Rows {
        rows: Vec<Row>,
        /// Sum of internal standalone row-projection bytes. This is not a
        /// serialized wire size and therefore is not protocol byte credit.
        projected_bytes: usize,
    },
    Finished(WorksheetCursorTail),
}

/// An owned worksheet-entry cursor whose projector survives across pulls.
///
/// Row batches are provisional until [`WorksheetCursorPull::Finished`]. A
/// caller assembling retained state must not commit earlier batches: the ZIP
/// CRC and well-formed worksheet tail are validated only when the entry reaches
/// EOF. `cancel` and `close` are deliberately idempotent and immediately drop
/// the entry stream/decoder lease.
pub(super) struct WorksheetCursor {
    projector: Option<ProductionProjector>,
    pending_row: Option<ProjectedWorksheetRow>,
    pending_tail: Option<WorksheetCursorTail>,
    state: WorksheetCursorState,
}

impl WorksheetCursor {
    fn open_under_active_operation(
        archive: &mut XlsxZip,
        part: &str,
        shared_strings: Rc<[SharedString]>,
        theme_colors: Rc<[String]>,
    ) -> Result<Self, String> {
        let entry = archive.active_operation()?.open_entry(part)?;
        let reporter = entry.limit_reporter()?;
        let projector = WorksheetRowProjector::from_owned_reader(
            Box::new(entry),
            part.to_string(),
            reporter,
            shared_strings,
            theme_colors,
        );
        Ok(Self {
            projector: Some(projector),
            pending_row: None,
            pending_tail: None,
            state: WorksheetCursorState::Open,
        })
    }

    pub(super) fn pull(
        &mut self,
        max_rows: usize,
        target_projected_bytes: usize,
    ) -> Result<WorksheetCursorPull, String> {
        // `max_rows` is a hard semantic-unit limit and is clamped below. The
        // projected-byte argument is intentionally a soft internal batching
        // target: one indivisible row may cross it after passing the separate
        // 8 MiB hard row-projection cap. A future wire adapter must measure its
        // serialized payload independently and obey protocol `byteCredit`.
        if max_rows == 0 || target_projected_bytes == 0 {
            return Err("worksheet cursor pull limits must be greater than zero".to_string());
        }
        if self.state != WorksheetCursorState::Open {
            return Err(self.inactive_error());
        }
        if let Some(tail) = self.pending_tail.take() {
            self.state = WorksheetCursorState::Finished;
            self.projector.take();
            return Ok(WorksheetCursorPull::Finished(tail));
        }

        let row_limit = max_rows.min(WORKSHEET_CURSOR_PULL_ROWS);
        let mut rows = Vec::with_capacity(row_limit);
        let mut projected_bytes = 0usize;
        loop {
            let item = match self.pending_row.take() {
                Some(row) => Ok(WorksheetProjectorItem::Row(row)),
                None => self
                    .projector
                    .as_mut()
                    .expect("open worksheet cursor owns its projector")
                    .next_item(),
            };
            match item {
                Ok(WorksheetProjectorItem::Row(row)) => {
                    let next_bytes = projected_bytes.saturating_add(row.projected_bytes);
                    if !rows.is_empty() && next_bytes > target_projected_bytes {
                        self.pending_row = Some(row);
                        return Ok(WorksheetCursorPull::Rows {
                            rows,
                            projected_bytes,
                        });
                    }
                    projected_bytes = next_bytes;
                    rows.push(row.row);
                    if rows.len() == row_limit {
                        return Ok(WorksheetCursorPull::Rows {
                            rows,
                            projected_bytes,
                        });
                    }
                }
                Ok(WorksheetProjectorItem::Finished(tail)) => {
                    let tail = WorksheetCursorTail {
                        shell_xml: tail.shell_xml,
                        row_geometry: tail.row_geometry,
                    };
                    if rows.is_empty() {
                        self.state = WorksheetCursorState::Finished;
                        self.projector.take();
                        return Ok(WorksheetCursorPull::Finished(tail));
                    }
                    self.pending_tail = Some(tail);
                    return Ok(WorksheetCursorPull::Rows {
                        rows,
                        projected_bytes,
                    });
                }
                Err(error) => {
                    self.state = WorksheetCursorState::Failed;
                    self.projector.take();
                    self.pending_tail = None;
                    return Err(error.to_string());
                }
            }
        }
    }

    pub(super) fn cancel(&mut self) {
        if matches!(
            self.state,
            WorksheetCursorState::Canceled | WorksheetCursorState::Closed
        ) {
            return;
        }
        self.projector.take();
        self.pending_row = None;
        self.pending_tail = None;
        self.state = WorksheetCursorState::Canceled;
    }

    pub(super) fn close(&mut self) {
        if self.state == WorksheetCursorState::Closed {
            return;
        }
        self.projector.take();
        self.pending_row = None;
        self.pending_tail = None;
        self.state = WorksheetCursorState::Closed;
    }

    fn inactive_error(&self) -> String {
        let state = match self.state {
            WorksheetCursorState::Open => "open",
            WorksheetCursorState::Finished => "finished",
            WorksheetCursorState::Failed => "failed",
            WorksheetCursorState::Canceled => "canceled",
            WorksheetCursorState::Closed => "closed",
        };
        format!("worksheet cursor is {state}")
    }
}

impl XlsxZip {
    pub(super) fn scan_worksheet_preview(
        &mut self,
        part: &str,
        shared_strings: Rc<[SharedString]>,
        theme_colors: Rc<[String]>,
    ) -> Result<WorksheetCursorPreview, String> {
        let entry = self.active_operation()?.open_entry(part)?;
        let scanned = lexical_scan_worksheet_preview(entry);
        let mut preview = match scanned {
            Ok(preview) => preview,
            Err(_) => {
                return Ok(WorksheetCursorPreview {
                    tail: None,
                    max_row: 0,
                    max_col: 0,
                    has_row_outline: false,
                    ordered_rows: false,
                })
            }
        };
        if let Some(tail) = preview
            .tail
            .as_mut()
            .filter(|tail| tail.shell_xml.contains(MCE_NS))
        {
            // The shell may carry MCE attributes. Apply the same projection as
            // the terminal cursor, but only to the small, row-free shell.
            let shell = tail.shell_xml.as_bytes().to_vec();
            let reporter = self.active_operation()?.limit_reporter()?;
            let mut projector = WorksheetRowProjector::from_owned_reader(
                Box::new(Cursor::new(shell)),
                part.to_string(),
                reporter,
                shared_strings,
                theme_colors,
            );
            match projector.next_item().map_err(|error| error.to_string())? {
                WorksheetProjectorItem::Finished(projected) => {
                    tail.shell_xml = projected.shell_xml;
                }
                WorksheetProjectorItem::Row(_) => {
                    return Err("row-free worksheet shell contained a row".to_string());
                }
            }
        }
        Ok(preview)
    }

    /// Open a persistent worksheet cursor only inside an explicitly-started
    /// package operation. This makes operation ownership visible at the factory
    /// boundary and prevents the lazy compatibility operation from escaping
    /// across production pulls.
    pub(super) fn open_worksheet_cursor(
        &mut self,
        part: &str,
        shared_strings: Rc<[SharedString]>,
        theme_colors: Rc<[String]>,
    ) -> Result<WorksheetCursor, String> {
        WorksheetCursor::open_under_active_operation(self, part, shared_strings, theme_colors)
    }
}

#[cfg(test)]
fn numeric_attribute(start: &BytesStart<'_>, name: &[u8]) -> Result<Option<String>, String> {
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| error.to_string())?;
        if attribute.key.as_ref() == name {
            return std::str::from_utf8(attribute.value.as_ref())
                .map(str::to_string)
                .map(Some)
                .map_err(|error| error.to_string());
        }
    }
    Ok(None)
}

/// Read the tail and row coordinates without projecting cell bodies. The
/// ordinary cursor later validates and projects the entry a second time. MCE
/// worksheets stay on that cursor's complete, MCE-aware path.
#[cfg(test)]
fn fast_scan_worksheet_preview(raw: Rc<[u8]>) -> Result<WorksheetCursorPreview, String> {
    let mut reader = BoundedXmlReader::new(
        std::io::BufReader::new(Cursor::new(Rc::clone(&raw))),
        1024 * 1024,
        "worksheet preview",
    );
    let mut sheet_start = None;
    let mut sheet_end = None;
    let mut inside_sheet = false;
    let mut previous_row = 0;
    let mut previous_col = 0;
    let mut max_row = 0;
    let mut max_col = 0;
    let mut has_row_outline = false;
    let mut ordered_rows = true;
    let mut row_geometry = AuthoredRowGeometry::default();
    loop {
        let read = reader.read_event().map_err(|error| format!("{error:?}"))?;
        let x = is_x_ns(read.namespace.as_deref());
        let empty = matches!(read.event, Event::Empty(_));
        match read.event {
            Event::Start(start) | Event::Empty(start)
                if x && start.local_name().as_ref() == b"sheetData" =>
            {
                if sheet_start.is_some() {
                    return Err("worksheet has repeated sheetData".to_string());
                }
                sheet_start = Some(read.span.end as usize);
                if empty {
                    sheet_end = sheet_start;
                } else {
                    inside_sheet = true;
                }
            }
            Event::End(end) if x && end.local_name().as_ref() == b"sheetData" => {
                sheet_end = Some(read.span.start as usize);
                inside_sheet = false;
            }
            Event::Start(start) | Event::Empty(start)
                if inside_sheet && x && start.local_name().as_ref() == b"row" =>
            {
                let explicit = numeric_attribute(&start, b"r")?
                    .map(|value| {
                        value
                            .parse::<u32>()
                            .map_err(|_| format!("invalid row ordinal: {value}"))
                    })
                    .transpose()?;
                let previous = previous_row;
                let row =
                    resolve_implicit_ordinal(explicit, &mut previous_row, SpreadsheetOrdinal::Row)?;
                ordered_rows &= row > previous;
                max_row = max_row.max(row);
                previous_col = 0;
                let hidden = numeric_attribute(&start, b"hidden")?
                    .is_some_and(|value| xml_bool_value(&value));
                let ht = numeric_attribute(&start, b"ht")?;
                row_geometry.record(row, authored_row_height(hidden, ht.as_deref()));
                let outline = numeric_attribute(&start, b"outlineLevel")?
                    .and_then(|value| value.parse::<u8>().ok())
                    .unwrap_or(0);
                let collapsed = numeric_attribute(&start, b"collapsed")?
                    .is_some_and(|value| xml_bool_value(&value));
                has_row_outline |= outline != 0 || collapsed;
            }
            Event::Start(start) | Event::Empty(start)
                if inside_sheet && x && start.local_name().as_ref() == b"c" =>
            {
                let explicit = numeric_attribute(&start, b"r")?
                    .map(|reference| parse_cell_ref_checked(&reference).map(|(col, _)| col))
                    .transpose()?;
                let col = resolve_implicit_ordinal(
                    explicit,
                    &mut previous_col,
                    SpreadsheetOrdinal::Column,
                )?;
                max_col = max_col.max(col);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let (Some(start), Some(end)) = (sheet_start, sheet_end) else {
        return Err("worksheet has no complete sheetData".to_string());
    };
    if end < start || end > raw.len() {
        return Err("worksheet sheetData boundary is invalid".to_string());
    }
    if raw[start..end]
        .windows(b"AlternateContent".len())
        .any(|window| window == b"AlternateContent")
        || raw
            .windows(b"ProcessContent".len())
            .any(|window| window == b"ProcessContent")
    {
        return Err("worksheet row MCE requires the complete projector".to_string());
    }
    let mut shell = Vec::with_capacity(raw.len() - (end - start));
    shell.extend_from_slice(&raw[..start]);
    shell.extend_from_slice(&raw[end..]);
    let shell_xml = String::from_utf8(shell).map_err(|error| error.to_string())?;
    Ok(WorksheetCursorPreview {
        tail: Some(WorksheetCursorTail {
            shell_xml,
            row_geometry,
        }),
        max_row,
        max_col,
        has_row_outline,
        ordered_rows,
    })
}

const PREVIEW_TAG_BYTES: usize = 1024 * 1024;
const PREVIEW_SHELL_BYTES: usize = 16 * 1024 * 1024;

fn append_shell(shell: &mut Vec<u8>, bytes: &[u8]) -> Result<(), String> {
    if bytes.len() > PREVIEW_SHELL_BYTES.saturating_sub(shell.len()) {
        return Err("worksheet preview shell exceeds retained limit".to_string());
    }
    shell.extend_from_slice(bytes);
    Ok(())
}

/// Skip text in buffered slices and retain only one XML tag at a time. Quoted
/// `>` is not a terminator. Special row markup is rejected below, so CDATA or
/// comments containing a fake sheetData boundary can never authorize paint.
fn next_markup<R: Read>(
    reader: &mut BufReader<R>,
    mut shell: Option<&mut Vec<u8>>,
    tag: &mut Vec<u8>,
) -> Result<bool, String> {
    tag.clear();
    loop {
        let available = reader.fill_buf().map_err(|error| error.to_string())?;
        if available.is_empty() {
            return Ok(false);
        }
        let available_len = available.len();
        let count = available
            .iter()
            .position(|byte| *byte == b'<')
            .unwrap_or(available_len);
        if let Some(shell) = shell.as_deref_mut() {
            append_shell(shell, &available[..count])?;
        }
        reader.consume(count);
        if count == available_len {
            continue;
        }
        break;
    }
    let mut quote = 0;
    loop {
        let available = reader.fill_buf().map_err(|error| error.to_string())?;
        if available.is_empty() {
            return Err("worksheet preview tag is unclosed".to_string());
        }
        let mut end = available.len();
        for (index, byte) in available.iter().enumerate() {
            if quote == 0 {
                match byte {
                    b'\'' | b'"' => quote = *byte,
                    b'>' => {
                        end = index + 1;
                        break;
                    }
                    _ => {}
                }
            } else if *byte == quote {
                quote = 0;
            }
        }
        if end > PREVIEW_TAG_BYTES.saturating_sub(tag.len()) {
            return Err("worksheet preview tag exceeds event limit".to_string());
        }
        tag.extend_from_slice(&available[..end]);
        reader.consume(end);
        if tag.last() == Some(&b'>') && quote == 0 {
            return Ok(true);
        }
    }
}

fn drain_entry<R: Read>(reader: &mut R) -> Result<(), String> {
    std::io::copy(reader, &mut std::io::sink())
        .map(|_| ())
        .map_err(|error| error.to_string())
}

fn append_remaining<R: Read>(reader: &mut R, shell: &mut Vec<u8>) -> Result<(), String> {
    let mut buffer = [0; 32 * 1024];
    let mut overflow = false;
    loop {
        let count = reader
            .read(&mut buffer)
            .map_err(|error| error.to_string())?;
        if count == 0 {
            break;
        }
        if !overflow {
            overflow = append_shell(shell, &buffer[..count]).is_err();
        }
    }
    if overflow {
        Err("worksheet preview shell exceeds retained limit".to_string())
    } else {
        Ok(())
    }
}

fn tag_attribute<'a>(tag: &'a [u8], name: &[u8]) -> Result<Option<Cow<'a, str>>, String> {
    let mut i = 1;
    while i < tag.len() && !tag[i].is_ascii_whitespace() && tag[i] != b'>' && tag[i] != b'/' {
        i += 1;
    }
    while i < tag.len() {
        while i < tag.len() && tag[i].is_ascii_whitespace() {
            i += 1;
        }
        if i == tag.len() || tag[i] == b'/' || tag[i] == b'>' {
            return Ok(None);
        }
        let key_start = i;
        while i < tag.len() && !tag[i].is_ascii_whitespace() && tag[i] != b'=' {
            i += 1;
        }
        let key = &tag[key_start..i];
        while i < tag.len() && tag[i].is_ascii_whitespace() {
            i += 1;
        }
        if i == tag.len() || tag[i] != b'=' {
            return Err("worksheet preview attribute is malformed".to_string());
        }
        i += 1;
        while i < tag.len() && tag[i].is_ascii_whitespace() {
            i += 1;
        }
        if i == tag.len() || (tag[i] != b'\'' && tag[i] != b'"') {
            return Err("worksheet preview attribute is unquoted".to_string());
        }
        let quote = tag[i];
        i += 1;
        let value_start = i;
        while i < tag.len() && tag[i] != quote {
            i += 1;
        }
        if i == tag.len() {
            return Err("worksheet preview attribute is unclosed".to_string());
        }
        let value = &tag[value_start..i];
        i += 1;
        if key == name {
            let encoded = std::str::from_utf8(value).map_err(|error| error.to_string())?;
            if encoded.contains('<') {
                return Err("worksheet preview attribute contains unescaped markup".to_string());
            }
            if !encoded.contains('&') {
                return Ok(Some(Cow::Borrowed(encoded)));
            }
            // The terminal row parser reads roxmltree's decoded attribute
            // values. Use that same XML decoder for the uncommon entity path;
            // a malformed or unsupported entity disables the preview.
            let delimiter = quote as char;
            let xml = format!("<v a={delimiter}{encoded}{delimiter}/>");
            let document = roxmltree::Document::parse(&xml).map_err(|error| error.to_string())?;
            let decoded = document
                .root_element()
                .attribute("a")
                .ok_or_else(|| "worksheet preview attribute disappeared".to_string())?;
            return Ok(Some(Cow::Owned(decoded.to_string())));
        }
    }
    Ok(None)
}

fn lexical_scan_worksheet_preview<R: Read>(source: R) -> Result<WorksheetCursorPreview, String> {
    let mut reader = BufReader::with_capacity(32 * 1024, source);
    let result = lexical_scan_worksheet_preview_inner(&mut reader);
    if result.is_err() {
        // Complete the first inflation even for an ineligible sheet. This
        // validates the entry CRC and charges actual work to the operation.
        drain_entry(&mut reader)?;
    }
    result
}

/// The lexical scan identifies SpreadsheetML elements by their unprefixed
/// names. That is sound only while the in-scope default namespace is
/// SpreadsheetML (transitional or strict), exactly as the row projector checks
/// with a namespace-aware reader. A default-namespace declaration naming any
/// other vocabulary disables the preview instead of letting foreign `row`
/// elements contribute geometry the terminal model would not have.
fn foreign_default_namespace(tag: &[u8]) -> Result<bool, String> {
    if !tag.windows(b"xmlns".len()).any(|part| part == b"xmlns") {
        return Ok(false);
    }
    Ok(tag_attribute(tag, b"xmlns")?.is_some_and(|namespace| !is_x_ns(Some(&namespace))))
}

fn lexical_scan_worksheet_preview_inner<R: Read>(
    reader: &mut BufReader<R>,
) -> Result<WorksheetCursorPreview, String> {
    let mut shell = Vec::new();
    let mut tag = Vec::with_capacity(128);
    // Element depth of the head. `sheetData` must be a direct child of the
    // SpreadsheetML `worksheet` root (§18.3.1.99), which the row projector
    // also enforces; a nested look-alike never authorizes a preview.
    let mut depth = 0usize;
    let self_closing = loop {
        if !next_markup(reader, Some(&mut shell), &mut tag)? {
            return Err("worksheet has no unprefixed sheetData".to_string());
        }
        if tag.starts_with(b"<!") || (tag.starts_with(b"<?") && !tag.starts_with(b"<?xml ")) {
            return Err("worksheet preview has special markup".to_string());
        }
        if tag
            .windows(b"ProcessContent".len())
            .any(|part| part == b"ProcessContent")
        {
            return Err("worksheet requires the complete XML projector".to_string());
        }
        append_shell(&mut shell, &tag)?;
        if tag.starts_with(b"<?") {
            continue;
        }
        if tag.starts_with(b"</") {
            depth = depth
                .checked_sub(1)
                .ok_or_else(|| "worksheet preview head is unbalanced".to_string())?;
            continue;
        }
        if depth == 0 {
            let root = tag.starts_with(b"<worksheet")
                && tag
                    .get(b"<worksheet".len())
                    .is_some_and(|byte| byte.is_ascii_whitespace() || *byte == b'>');
            let namespace = tag_attribute(&tag, b"xmlns")?;
            if !root || !namespace.is_some_and(|namespace| is_x_ns(Some(&namespace))) {
                return Err("worksheet preview needs an unprefixed SpreadsheetML root".to_string());
            }
        } else if foreign_default_namespace(&tag)? {
            return Err("worksheet preview has a foreign default namespace".to_string());
        }
        if tag.starts_with(b"<sheetData")
            && tag
                .get(b"<sheetData".len())
                .is_some_and(|byte| byte.is_ascii_whitespace() || *byte == b'/' || *byte == b'>')
        {
            if depth != 1 {
                return Err("worksheet preview sheetData is not a root child".to_string());
            }
            break tag.ends_with(b"/>");
        }
        if !tag.ends_with(b"/>") {
            depth += 1;
        }
    };
    let mut previous_row = 0;
    let mut previous_col = 0;
    let mut max_row = 0;
    let mut max_col = 0;
    let mut has_row_outline = false;
    let mut ordered_rows = true;
    let mut row_geometry = AuthoredRowGeometry::default();
    if !self_closing {
        loop {
            if !next_markup(reader, None, &mut tag)? {
                return Err("worksheet has no closing sheetData".to_string());
            }
            if tag == b"</sheetData>" {
                append_shell(&mut shell, &tag)?;
                break;
            }
            if tag.get(1).is_some_and(|byte| *byte == b'/') {
                continue;
            }
            if tag
                .get(1)
                .is_some_and(|byte| *byte == b'!' || *byte == b'?')
            {
                return Err("worksheet preview has special row markup".to_string());
            }
            if tag
                .windows(b"AlternateContent".len())
                .any(|part| part == b"AlternateContent")
                || tag
                    .windows(b"ProcessContent".len())
                    .any(|part| part == b"ProcessContent")
            {
                return Err("worksheet row MCE requires the complete projector".to_string());
            }
            let mut name_end = 1;
            while name_end < tag.len()
                && !tag[name_end].is_ascii_whitespace()
                && tag[name_end] != b'/'
                && tag[name_end] != b'>'
            {
                name_end += 1;
            }
            let name = &tag[1..name_end];
            if name.contains(&b':') {
                return Err("worksheet preview has prefixed row markup".to_string());
            }
            if foreign_default_namespace(&tag)? {
                return Err("worksheet preview has a foreign default namespace".to_string());
            }
            if name == b"row" {
                let explicit = tag_attribute(&tag, b"r")?
                    .map(|value| value.parse::<u32>().map_err(|error| error.to_string()))
                    .transpose()?;
                let prior = previous_row;
                let row =
                    resolve_implicit_ordinal(explicit, &mut previous_row, SpreadsheetOrdinal::Row)?;
                ordered_rows &= row > prior;
                max_row = max_row.max(row);
                previous_col = 0;
                // Same readers as the row projector's `parse_row_node`: the
                // tail parse resolves sheet-level rules such as zeroHeight from
                // these recorded facts on both paths.
                let hidden =
                    tag_attribute(&tag, b"hidden")?.is_some_and(|value| xml_bool_value(&value));
                let ht = tag_attribute(&tag, b"ht")?;
                row_geometry.record(row, authored_row_height(hidden, ht.as_deref()));
                let outline = tag_attribute(&tag, b"outlineLevel")?
                    .and_then(|value| value.parse::<u8>().ok())
                    .unwrap_or(0);
                let collapsed =
                    tag_attribute(&tag, b"collapsed")?.is_some_and(|value| xml_bool_value(&value));
                has_row_outline |= outline != 0 || collapsed;
            } else if name == b"c" {
                let explicit = tag_attribute(&tag, b"r")?
                    .map(|reference| parse_cell_ref_checked(&reference).map(|(col, _)| col))
                    .transpose()?;
                let col = resolve_implicit_ordinal(
                    explicit,
                    &mut previous_col,
                    SpreadsheetOrdinal::Column,
                )?;
                max_col = max_col.max(col);
            }
        }
    }
    append_remaining(reader, &mut shell)?;
    let shell_xml = String::from_utf8(shell).map_err(|error| error.to_string())?;
    Ok(WorksheetCursorPreview {
        tail: Some(WorksheetCursorTail {
            shell_xml,
            row_geometry,
        }),
        max_row,
        max_col,
        has_row_outline,
        ordered_rows,
    })
}

impl Drop for WorksheetCursor {
    fn drop(&mut self) {
        self.close();
    }
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use zip::write::SimpleFileOptions;

    use super::*;
    use crate::open_zip;

    const PART: &str = "xl/worksheets/sheet1.xml";

    fn worksheet(row_count: usize, tail: &str) -> String {
        let mut xml = String::from(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
        );
        for row in 1..=row_count {
            xml.push_str(&format!(
                r#"<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>row {row}</t></is></c></row>"#
            ));
        }
        xml.push_str(tail);
        xml
    }

    fn package(xml: &str) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    PART,
                    SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored),
                )
                .unwrap();
            writer.write_all(xml.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        bytes
    }

    fn corrupt_crc_consistently(bytes: &mut [u8]) {
        let wrong_crc = u32::from_le_bytes(bytes[14..18].try_into().unwrap()) ^ 0xffff_ffff;
        bytes[14..18].copy_from_slice(&wrong_crc.to_le_bytes());
        let central = bytes
            .windows(4)
            .position(|window| window == 0x0201_4b50u32.to_le_bytes())
            .expect("central directory header");
        bytes[central + 16..central + 20].copy_from_slice(&wrong_crc.to_le_bytes());
    }

    #[test]
    fn lexical_tail_scan_matches_xml_scan_for_implicit_and_explicit_coordinates() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2" ht="21"><c r="B2"><v>1</v></c><c><v>2</v></c></row><row hidden="1"><c r="D3"/></row><row r="8" outlineLevel="2"><c r="AA8"><v>3</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="B2:C2"/></mergeCells></worksheet>"#;
        let raw: Rc<[u8]> = xml.as_bytes().into();
        let lexical = lexical_scan_worksheet_preview(Cursor::new(Rc::clone(&raw))).unwrap();
        let parsed = fast_scan_worksheet_preview(raw).unwrap();
        assert_eq!(lexical.max_row, parsed.max_row);
        assert_eq!(lexical.max_col, parsed.max_col);
        assert_eq!(lexical.has_row_outline, parsed.has_row_outline);
        assert_eq!(lexical.ordered_rows, parsed.ordered_rows);
        assert_eq!(
            lexical.tail.unwrap().row_geometry,
            parsed.tail.unwrap().row_geometry
        );
    }

    #[test]
    fn lexical_scan_decodes_row_attributes_like_the_terminal_xml_parser() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="&#50;" ht="&#51;0"><c r="B&#50;"><v>1</v></c></row><row r="3" ht="3&amp;0"><c r="C3"><v>2</v></c></row></sheetData></worksheet>"#;
        let preview = lexical_scan_worksheet_preview(xml.as_bytes()).unwrap();
        assert_eq!(preview.max_row, 3);
        assert_eq!(preview.max_col, 3);
        assert_eq!(
            preview.tail.unwrap().row_geometry.heights,
            BTreeMap::from([(2, 30.0)])
        );

        // The main row projector uses roxmltree attribute values. Compare the
        // same decoded values, including a predefined named entity whose
        // result is not a valid row height.
        let document = roxmltree::Document::parse(xml).unwrap();
        let rows: Vec<_> = document
            .descendants()
            .filter(|node| node.has_tag_name("row"))
            .collect();
        assert_eq!(rows[0].attribute("ht"), Some("30"));
        assert_eq!(rows[1].attribute("ht"), Some("3&0"));

        let unsupported = xml.replace("3&amp;0", "3&unknown;0");
        assert!(lexical_scan_worksheet_preview(unsupported.as_bytes()).is_err());
    }

    #[test]
    fn special_row_markup_cannot_forge_the_tail_boundary() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"/><![CDATA[</sheetData><mergeCells/>]]><row r="2"/></sheetData></worksheet>"#;
        assert!(lexical_scan_worksheet_preview(xml.as_bytes()).is_err());

        let xml = r#"<?pi <sheetData><row r="1"/></sheetData> ?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"/></sheetData></worksheet>"#;
        assert!(lexical_scan_worksheet_preview(xml.as_bytes()).is_err());

        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1" ht="&#51;0"/></sheetData><mergeCells count="0"/></worksheet>"#;
        let reader = xml
            .as_bytes()
            .chunks(3)
            .flat_map(|chunk| chunk.iter().copied());
        // A short-read source exercises boundaries in the XML tag and entity.
        struct ShortRead<I>(I);
        impl<I: Iterator<Item = u8>> Read for ShortRead<I> {
            fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
                if let Some(byte) = self.0.next() {
                    buf[0] = byte;
                    Ok(1)
                } else {
                    Ok(0)
                }
            }
        }
        let preview = lexical_scan_worksheet_preview(ShortRead(reader)).unwrap();
        assert_eq!(
            preview.tail.unwrap().row_geometry.heights.get(&1),
            Some(&30.0)
        );
    }

    #[test]
    fn preview_and_cursor_charge_both_inflations_without_double_counting_distinct_bytes() {
        let xml = worksheet(70_000, "</sheetData></worksheet>");
        let bytes = xml.len() as u64;
        let mut archive = crate::open_zip_with_limits(package(&xml), Some(bytes), Some(bytes))
            .expect("one inflated entry fits both configured limits");
        archive.begin_operation("parse-sheet").unwrap();
        let preview = archive
            .scan_worksheet_preview(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        assert_eq!(preview.max_row, 70_000);
        assert!(preview.tail.is_some());
        assert_eq!(archive.usage().distinct_inflated_bytes, bytes);
        assert_eq!(
            archive
                .operation
                .active()
                .unwrap()
                .usage()
                .unwrap()
                .operation_inflated_bytes,
            bytes
        );
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        let mut rows = 0;
        while let WorksheetCursorPull::Rows { rows: batch, .. } = cursor
            .pull(128, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES)
            .unwrap()
        {
            rows += batch.len();
        }
        assert_eq!(rows, 70_000);
        assert_eq!(archive.usage().distinct_inflated_bytes, bytes);
        assert_eq!(
            archive
                .operation
                .active()
                .unwrap()
                .usage()
                .unwrap()
                .operation_inflated_bytes,
            bytes * 2
        );
        archive.finish_operation().unwrap();
    }

    #[test]
    fn production_cursor_keeps_one_operation_across_multiple_atomic_pulls() {
        let xml = worksheet(600, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).expect("package opens");
        archive
            .begin_operation("worksheet-cursor")
            .expect("operation starts");
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .expect("cursor opens");

        let mut batches = Vec::new();
        let mut inflated_snapshots = Vec::new();
        let (batches, tail) = loop {
            match cursor
                .pull(64, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES)
                .expect("pull succeeds")
            {
                WorksheetCursorPull::Rows { rows, .. } => {
                    assert!(!rows.is_empty());
                    assert!(rows.len() <= 64);
                    batches.push(rows);
                    inflated_snapshots.push(
                        archive
                            .operation
                            .active()
                            .unwrap()
                            .usage()
                            .unwrap()
                            .operation_inflated_bytes,
                    );
                }
                WorksheetCursorPull::Finished(tail) => {
                    break (batches, tail);
                }
            }
        };

        assert_eq!(batches.iter().map(Vec::len).sum::<usize>(), 600);
        assert_eq!(
            batches
                .iter()
                .flatten()
                .map(|row| row.index)
                .collect::<Vec<_>>(),
            (1..=600).collect::<Vec<_>>()
        );
        assert!(
            inflated_snapshots.windows(2).any(|pair| pair[0] < pair[1]),
            "the same operation continues inflating after earlier row pulls"
        );
        assert!(tail.shell_xml.contains("<sheetData></sheetData>"));
        assert!(archive.operation.active().unwrap().usage().is_some());
        archive.finish_operation().expect("same operation finishes");
    }

    #[test]
    fn late_malformed_tail_never_produces_a_committable_transaction() {
        let xml = worksheet(600, "</sheetData><broken>");
        let mut archive = open_zip(package(&xml)).expect("package opens");
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        let mut provisional = Vec::new();
        let mut finished = false;

        loop {
            match cursor.pull(7, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES) {
                Ok(WorksheetCursorPull::Rows { rows, .. }) => provisional.extend(rows),
                Ok(WorksheetCursorPull::Finished(_)) => {
                    finished = true;
                    break;
                }
                Err(error) => {
                    assert!(error.contains("EOF") || error.contains("closed"), "{error}");
                    break;
                }
            }
        }

        assert!(
            !provisional.is_empty(),
            "earlier pulls are intentionally provisional"
        );
        assert!(!finished, "malformed tail must prevent commit");
        archive.cancel_operation();
    }

    #[test]
    fn crc_failure_after_row_pulls_never_produces_finished() {
        let xml = worksheet(600, "</sheetData></worksheet>");
        let mut bytes = package(&xml);
        corrupt_crc_consistently(&mut bytes);
        let mut archive = open_zip(bytes).expect("matching forged metadata passes preflight");
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        let mut provisional_rows = 0;

        let error = loop {
            match cursor.pull(7, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES) {
                Ok(WorksheetCursorPull::Rows { rows, .. }) => provisional_rows += rows.len(),
                Ok(WorksheetCursorPull::Finished(_)) => panic!("CRC failure must prevent commit"),
                Err(error) => break error,
            }
        };

        assert!(provisional_rows > 0);
        assert!(error.contains("CRC"), "{error}");
        archive.cancel_operation();
    }

    #[test]
    fn close_and_cancel_are_idempotent_and_release_the_entry() {
        let xml = worksheet(2, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        archive.begin_operation("parse-sheet").unwrap();

        let mut closed = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        closed.close();
        closed.close();
        assert_eq!(closed.pull(1, 1).unwrap_err(), "worksheet cursor is closed");

        let mut canceled = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        canceled.cancel();
        canceled.cancel();
        assert_eq!(
            canceled.pull(1, 1).unwrap_err(),
            "worksheet cursor is canceled"
        );
        archive
            .finish_operation()
            .expect("released readers allow finish");
    }

    #[test]
    fn pull_clamps_rows_and_stages_the_first_soft_projection_overrun() {
        let xml = worksheet(600, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();

        let first = cursor.pull(usize::MAX, usize::MAX).unwrap();
        let WorksheetCursorPull::Rows { rows, .. } = first else {
            panic!("large worksheet must yield rows");
        };
        assert_eq!(rows.len(), WORKSHEET_CURSOR_PULL_ROWS);

        let mut observed = rows.into_iter().map(|row| row.index).collect::<Vec<_>>();
        let mut saw_multi_row_projection_limited_batch = false;
        while let WorksheetCursorPull::Rows {
            rows,
            projected_bytes,
        } = cursor.pull(usize::MAX, 500).unwrap()
        {
            assert!(!rows.is_empty());
            if rows.len() > 1 {
                saw_multi_row_projection_limited_batch = true;
                assert!(projected_bytes <= 500);
            }
            observed.extend(rows.into_iter().map(|row| row.index));
        }
        assert!(saw_multi_row_projection_limited_batch);
        assert_eq!(observed, (1..=600).collect::<Vec<_>>());
        archive.finish_operation().unwrap();
    }

    #[test]
    fn indivisible_row_may_cross_soft_projection_target_but_not_hard_row_cap() {
        let xml = worksheet(2, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();

        for expected in 1..=2 {
            let WorksheetCursorPull::Rows {
                rows,
                projected_bytes,
            } = cursor.pull(128, 1).unwrap()
            else {
                panic!("row is returned atomically");
            };
            assert_eq!(rows.len(), 1);
            assert_eq!(rows[0].index, expected);
            assert!(projected_bytes > 1);
            assert!(projected_bytes <= crate::worksheet_projector::STREAMED_ROW_PROJECTION_BYTES);
        }
        assert!(matches!(
            cursor.pull(128, 1).unwrap(),
            WorksheetCursorPull::Finished(_)
        ));
        archive.finish_operation().unwrap();
    }

    #[test]
    fn cursor_factory_requires_an_explicit_active_operation() {
        let xml = worksheet(1, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        let error = match archive.open_worksheet_cursor(PART, Rc::from([]), Rc::from([])) {
            Ok(_) => panic!("cursor factory must not create a compatibility operation"),
            Err(error) => error,
        };
        assert_eq!(error, "xlsx package operation is not active");
        assert!(!archive.operation.is_active());
    }
}
