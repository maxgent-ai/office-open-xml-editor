//! Bounded, resumable projection of SpreadsheetML worksheet rows.
//!
//! `sheetData` is the one unbounded worksheet child, so this module owns the
//! complete streaming boundary: XML/MCE preprocessing, hard parser ceilings,
//! bounded row batches, and the compatibility materializer used by `lib.rs`.

use std::collections::{BTreeMap, VecDeque};
use std::fmt::Write as _;
#[cfg(test)]
use std::io::Cursor;
use std::io::{BufRead, BufReader, Read};
use std::rc::Rc;

use ooxml_common::bounded_xml::{
    self, BoundedXmlError, BoundedXmlReadError, BoundedXmlReader, MceScope as StreamedMceScope,
    NamespaceContext as StreamedNamespaceContext, MCE_NS,
};
use ooxml_common::depth::{parse_guarded, MAX_XML_DEPTH};
use ooxml_common::ns::is_x_ns;
use ooxml_common::package_session::{PackageEntryStream, PackageLimitReporter};
use ooxml_common::resource::HardResourceLimitKind;
use quick_xml::events::Event;
use quick_xml::NsReader;

use crate::{
    attr_bool, parse_row_cells, resolve_implicit_ordinal, xlsx_understands_ns, Row, SharedString,
    SpreadsheetOrdinal,
};

/// The memory-bounded portion of a worksheet parse.
///
/// ECMA-376 Part 1 §18.3.1.99 / `CT_Worksheet` defines exactly one
/// `sheetData` (§18.3.1.80), whose `CT_SheetData` content is an unbounded
/// sequence of `CT_Row` / `row` (§18.3.1.73). Building one roxmltree node arena
/// for that sequence
/// multiplies the already-inflated XML cost for large sheets. We therefore
/// stream only `sheetData`, parse bounded batches of complete rows with the
/// existing row/cell implementation, and return a shell XML with the
/// `sheetData` interior removed for the worksheet-level roxmltree pass.
///
/// The public `Worksheet` model is deliberately unchanged. Its rows (and the
/// later serde JSON buffer) still scale with the cell count; this removes the
/// additional full-sheet XML tree rather than claiming constant-memory parsing.
#[derive(Debug)]
pub(super) struct StreamedSheetData {
    pub(super) shell_xml: String,
    pub(super) rows: Vec<Row>,
    pub(super) row_geometry: AuthoredRowGeometry,
}

/// Row-band facts authored on `<row>` elements (ECMA-376 §18.3.1.73), recorded
/// while `sheetData` streams so the shell parse can resolve sheet-level rules
/// (`sheetFormatPr@zeroHeight`, §18.3.1.81) after the rows themselves have been
/// emitted and dropped. Both the bounded row projector and the lexical
/// worksheet preview record through [`AuthoredRowGeometry::record`] with a
/// height from [`authored_row_height`], so the two paths cannot diverge.
#[derive(Clone, Debug, Default, PartialEq)]
pub(super) struct AuthoredRowGeometry {
    /// Hidden rows (zero) and rows with a valid `@ht`, last authored row wins.
    pub(super) heights: BTreeMap<u32, f64>,
    /// Visible rows without a valid `@ht`, as inclusive index runs in
    /// document order. Contiguous authored rows collapse into one run, so this
    /// stays small for ordinary data and never exceeds the row count.
    pub(super) unspecified_visible: Vec<(u32, u32)>,
}

impl AuthoredRowGeometry {
    pub(super) fn record(&mut self, index: u32, height: Option<f64>) {
        match height {
            Some(height) => {
                self.heights.insert(index, height);
            }
            None => match self.unspecified_visible.last_mut() {
                Some((_, last)) if last.checked_add(1) == Some(index) => *last = index,
                _ => self.unspecified_visible.push((index, index)),
            },
        }
    }
}

/// Authored band of one row. `hidden` (already read as an XML boolean) forces
/// a zero band; otherwise `@ht` is the height in points when it is a finite,
/// nonnegative number. `customHeight` records how the height was set and does
/// not gate the value. `None` means the sheet default applies.
pub(super) fn authored_row_height(hidden: bool, ht: Option<&str>) -> Option<f64> {
    if hidden {
        return Some(0.0);
    }
    ht.and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite() && *value >= 0.0)
}

/// Hidden safety ceiling for one lexical XML event. `quick_xml` otherwise keeps
/// extending the caller-owned event buffer until it finds the token delimiter,
/// so one hostile text node or start tag could temporarily grow to the complete
/// worksheet entry size before any row-level limit can run.
const STREAMED_XML_EVENT_BYTES: usize = 1024 * 1024;

/// A complete row remains the indivisible roxmltree unit. Eight MiB is a
/// conservative browser/WASM safety ceiling that still permits unusually rich
/// authored rows while preventing many individually-small XML events from
/// accumulating without bound.
pub(super) const STREAMED_ROW_PROJECTION_BYTES: usize = 8 * 1024 * 1024;

/// The retained worksheet shell contains metadata outside `sheetData`. Keep it
/// large enough for substantial validation/formatting metadata, but bounded so
/// a worksheet cannot bypass the event ceiling with many small retained nodes.
const STREAMED_WORKSHEET_SHELL_BYTES: usize = 16 * 1024 * 1024;

/// Maximum retained namespace/MCE context along the active XML path. Context
/// is persistent and shared between frames; this limit accounts only newly
/// declared state, so inherited declarations are not multiplied by depth.
const STREAMED_XML_CONTEXT_BYTES: usize = 4 * 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum WorksheetProjectorLimitKind {
    XmlEventBytes,
    XmlContextBytes,
    XmlDepth,
    RowProjectionBytes,
    WorksheetShellBytes,
}

impl WorksheetProjectorLimitKind {
    const fn resource_name(self) -> &'static str {
        match self {
            Self::XmlEventBytes => "xml-event-bytes",
            Self::XmlContextBytes => "xml-context-bytes",
            Self::XmlDepth => "xml-depth",
            Self::RowProjectionBytes => "row-projection-bytes",
            Self::WorksheetShellBytes => "worksheet-shell-bytes",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum WorksheetProjectorError {
    Xml(String),
    Limit {
        kind: WorksheetProjectorLimitKind,
        part: Option<String>,
        limit: usize,
        observed: usize,
    },
}

impl std::fmt::Display for WorksheetProjectorError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Xml(message) => formatter.write_str(message),
            Self::Limit {
                kind,
                part,
                limit,
                observed,
            } => {
                let resource = kind.resource_name();
                if let Some(part) = part {
                    write!(
                        formatter,
                        "worksheet projector limit exceeded ({resource}, {part}): {observed} > {limit}"
                    )
                } else {
                    write!(
                        formatter,
                        "worksheet projector limit exceeded ({resource}): {observed} > {limit}"
                    )
                }
            }
        }
    }
}

impl std::error::Error for WorksheetProjectorError {}

impl From<String> for WorksheetProjectorError {
    fn from(error: String) -> Self {
        Self::Xml(error)
    }
}

/// Single mapping point for XLSX hard limits. An active package governor gets
/// the canonical wire envelope and poison latch; compatibility calls without a
/// governor retain the local structured error used by focused parser tests.
fn worksheet_projector_limit(
    kind: WorksheetProjectorLimitKind,
    part: Option<&str>,
    limit: usize,
    observed: usize,
) -> WorksheetProjectorError {
    WorksheetProjectorError::Limit {
        kind,
        part: part.map(str::to_string),
        limit,
        observed,
    }
}

fn map_bounded_xml_error(
    error: BoundedXmlError,
    kind: WorksheetProjectorLimitKind,
    part: Option<&str>,
) -> WorksheetProjectorError {
    match error {
        BoundedXmlError::Xml(message) => WorksheetProjectorError::Xml(message),
        BoundedXmlError::Limit { limit, observed } => {
            worksheet_projector_limit(kind, part, limit, observed)
        }
    }
}

/// Map the projecter's internal hard-limit vocabulary exactly once at its
/// state-machine boundary. Package-backed callers poison their owning package
/// operation through the explicit capability; compatibility callers retain the
/// historical TLS scope used by the existing WASM entry points.
fn report_projector_limit(
    reporter: Option<&PackageLimitReporter>,
    error: WorksheetProjectorError,
) -> WorksheetProjectorError {
    let WorksheetProjectorError::Limit {
        kind,
        ref part,
        limit,
        observed,
    } = error
    else {
        return error;
    };
    let common_kind = match kind {
        WorksheetProjectorLimitKind::XmlEventBytes => HardResourceLimitKind::XmlEventBytes,
        WorksheetProjectorLimitKind::XmlContextBytes => HardResourceLimitKind::XmlContextBytes,
        WorksheetProjectorLimitKind::XmlDepth => HardResourceLimitKind::XmlNestingDepth,
        WorksheetProjectorLimitKind::RowProjectionBytes => {
            HardResourceLimitKind::WorksheetRowProjectionBytes
        }
        WorksheetProjectorLimitKind::WorksheetShellBytes => {
            HardResourceLimitKind::WorksheetShellProjectionBytes
        }
    };
    let report =
        bounded_xml::report_hard_limit(reporter, common_kind, part.as_deref(), limit, observed);
    if let Err(message) = report {
        return WorksheetProjectorError::Xml(message);
    }
    error
}

/// Keep each temporary roxmltree arena small while amortizing its allocation
/// across many ordinary rows. A single unusually large `CT_Row` remains the
/// indivisible upper bound because its rich-string and formula descendants must
/// be interpreted together.
const STREAMED_ROW_BATCH_BYTES: usize = 1024 * 1024;
const STREAMED_ROW_BATCH_ROWS: usize = 512;
const STREAMED_ROWS_ARENA_OVERHEAD: usize = "<streamed-rows></streamed-rows>".len();

struct PendingStreamedRow {
    xml: Vec<u8>,
    namespaces: Vec<(Option<String>, String)>,
}

impl PendingStreamedRow {
    fn projected_arena_bytes(&self) -> usize {
        let row_bytes = self.xml.len();
        let namespace_bytes = self.namespaces.iter().fold(0usize, |total, (prefix, uri)| {
            let declaration_bytes = match prefix {
                // ` xmlns:{prefix}="{escaped-uri}"`
                Some(prefix) => 10usize
                    .saturating_add(prefix.len())
                    .saturating_add(quick_xml::escape::escape(uri).len()),
                // ` xmlns="{escaped-uri}"`
                None => 9usize.saturating_add(quick_xml::escape::escape(uri).len()),
            };
            total.saturating_add(declaration_bytes)
        });
        row_bytes
            .saturating_add(namespace_bytes)
            .saturating_add("<streamed-row></streamed-row>".len())
    }
}

struct ActiveStreamedRow {
    depth: usize,
    namespaces: Vec<(Option<String>, String)>,
    projection_overhead: usize,
    xml: Vec<u8>,
}

enum StreamedElementKind {
    Retained {
        namespace: Option<String>,
        local_name: String,
        opaque: bool,
    },
    Unwrapped,
    Ignored,
    AlternateContent {
        selected_branch: bool,
        seen_choice: bool,
        seen_fallback: bool,
    },
    AlternateBranch {
        selected: bool,
    },
}

struct StreamedElementFrame {
    kind: StreamedElementKind,
    scope: Rc<StreamedMceScope>,
    namespace_context: Rc<StreamedNamespaceContext>,
    visible: bool,
}

impl StreamedElementFrame {
    fn children_are_visible(&self) -> bool {
        self.visible
            && !matches!(
                self.kind,
                StreamedElementKind::Ignored
                    | StreamedElementKind::AlternateBranch { selected: false }
            )
    }

    fn is_opaque(&self) -> bool {
        matches!(
            self.kind,
            StreamedElementKind::Retained { opaque: true, .. }
        )
    }
}

fn worksheet_understands_ns(namespace: &str) -> bool {
    is_x_ns(Some(namespace)) || xlsx_understands_ns(namespace)
}

/// SpreadsheetML designates `extLst` as an application-defined extension
/// element (Part 1 §10 and §18.2.10). Per Part 3 §§8 and 9.1, MCE processing is
/// suspended for the complete contents of this element. Ordinary worksheet,
/// row, and cell content is not an extension boundary and remains processed.
fn worksheet_is_application_defined_extension_element(
    namespace: Option<&str>,
    local_name: &str,
) -> bool {
    namespace.is_some_and(|namespace| is_x_ns(Some(namespace))) && local_name == "extLst"
}

fn inject_moved_element_namespaces(
    element: &mut quick_xml::events::BytesStart<'static>,
    element_context: &StreamedNamespaceContext,
    processed_parent_context: Option<&StreamedNamespaceContext>,
    part: Option<&str>,
) -> Result<(), WorksheetProjectorError> {
    bounded_xml::inject_missing_namespaces(
        element,
        element_context,
        processed_parent_context,
        STREAMED_XML_CONTEXT_BYTES,
        "worksheet",
    )
    .map_err(|error| {
        map_bounded_xml_error(error, WorksheetProjectorLimitKind::XmlContextBytes, part)
    })
}

fn append_projected_event(
    target: &mut Vec<u8>,
    event: &Event<'_>,
    retained_overhead: usize,
    limit: usize,
    kind: WorksheetProjectorLimitKind,
    part: Option<&str>,
) -> Result<(), WorksheetProjectorError> {
    bounded_xml::append_projected_event(target, event, retained_overhead, limit, "worksheet")
        .map_err(|error| map_bounded_xml_error(error, kind, part))
}

/// Apply the same application-configuration predicate as the shared DOM MCE
/// selector without materializing the potentially large `sheetData` subtree.
///
/// ECMA-376 Part 3 §9.3 and §7.6 define `Requires` as namespace prefixes:
/// every prefix must resolve to a namespace understood by this XLSX parser.
fn streamed_choice_is_understood(
    choice: &quick_xml::events::BytesStart<'_>,
    context: &StreamedNamespaceContext,
    part: Option<&str>,
) -> Result<bool, String> {
    use ooxml_common::mce::ChoiceRequiresClassification;

    let classification = bounded_xml::classify_bounded_mce_choice_requires(
        choice,
        context,
        &worksheet_understands_ns,
        STREAMED_XML_CONTEXT_BYTES,
        "worksheet",
    )
    .map_err(|error| {
        map_bounded_xml_error(error, WorksheetProjectorLimitKind::XmlContextBytes, part).to_string()
    })?;
    match classification {
        ChoiceRequiresClassification::Understood => Ok(true),
        ChoiceRequiresClassification::Unsupported => Ok(false),
        ChoiceRequiresClassification::Missing | ChoiceRequiresClassification::Blank => {
            Err("worksheet MCE Choice must have a non-empty Requires attribute".to_string())
        }
        ChoiceRequiresClassification::Unresolved => {
            Err("worksheet MCE Choice Requires uses an unbound namespace prefix".to_string())
        }
    }
}

/// Append a tiny namespace-preserving wrapper around one raw `<row>`.
///
/// Namespace declarations are inherited in XML, so a raw row slice is not
/// necessarily a standalone document (`<x:row>` commonly inherits `xmlns:x`
/// from `<x:worksheet>`). `NsReader::prefixes` exposes the effective bindings at
/// the row start. Re-declaring those bindings on a neutral wrapper preserves the
/// row's expanded names, including Strict/Transitional SpreadsheetML, MCE
/// wrappers, extension prefixes, and `xml:space` semantics, without copying any
/// sibling row.
fn append_wrapped_row(
    fragment: &mut String,
    row_xml: &str,
    namespaces: &[(Option<String>, String)],
) -> Result<(), String> {
    fragment.push_str("<streamed-row");
    for (prefix, namespace) in namespaces {
        match prefix {
            Some(prefix) => write!(
                fragment,
                " xmlns:{}=\"{}\"",
                prefix,
                quick_xml::escape::escape(namespace)
            ),
            None => write!(
                fragment,
                " xmlns=\"{}\"",
                quick_xml::escape::escape(namespace)
            ),
        }
        .map_err(|e| e.to_string())?;
    }
    fragment.push('>');
    fragment.push_str(row_xml);
    fragment.push_str("</streamed-row>");
    Ok(())
}

fn flush_streamed_rows(
    pending: &mut Vec<PendingStreamedRow>,
    prev_row_idx: &mut u32,
    shared_strings: &[SharedString],
    theme_colors: &[String],
) -> Result<Vec<ProjectedWorksheetRow>, String> {
    if pending.is_empty() {
        return Ok(Vec::new());
    }
    let content_bytes = pending.iter().fold(0usize, |total, row| {
        total.saturating_add(row.projected_arena_bytes())
    });
    let mut fragment =
        String::with_capacity(content_bytes.saturating_add(STREAMED_ROWS_ARENA_OVERHEAD));
    fragment.push_str("<streamed-rows>");
    for row in pending.iter() {
        let row_xml = std::str::from_utf8(&row.xml)
            .map_err(|error| format!("streamed worksheet row is not UTF-8: {error}"))?;
        append_wrapped_row(&mut fragment, row_xml, &row.namespaces)?;
    }
    fragment.push_str("</streamed-rows>");
    let doc = parse_guarded(&fragment).map_err(|e| e.to_string())?;
    let mut rows = Vec::with_capacity(pending.len());
    for (wrapper, source) in doc
        .root_element()
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "streamed-row")
        .zip(pending.iter())
    {
        let row_node = wrapper
            .children()
            .find(|node| {
                node.is_element()
                    && node.tag_name().name() == "row"
                    && is_x_ns(node.tag_name().namespace())
            })
            .ok_or_else(|| "streamed worksheet row lost its SpreadsheetML namespace".to_string())?;
        let row = parse_row_node(&row_node, prev_row_idx, shared_strings, theme_colors)?;
        rows.push(ProjectedWorksheetRow {
            row,
            projected_bytes: source.projected_arena_bytes(),
        });
    }
    if rows.len() != pending.len() {
        return Err("streamed worksheet row batch changed row cardinality".to_string());
    }
    pending.clear();
    Ok(rows)
}

fn parse_row_node(
    node: &roxmltree::Node<'_, '_>,
    prev_row_idx: &mut u32,
    shared_strings: &[SharedString],
    theme_colors: &[String],
) -> Result<Row, String> {
    // ECMA-376 §18.3.1.73 makes `@r` optional; honor an explicit value when
    // present. When omitted, take the running previous row + 1 (implicit
    // sequential numbering — the de-facto consumer convention; the spec only
    // grants the optionality). An explicit value also re-anchors the counter.
    let explicit_row = node
        .attribute("r")
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| format!("worksheet row has invalid r ordinal: {value}"))
        })
        .transpose()?;
    let row_idx = resolve_implicit_ordinal(explicit_row, prev_row_idx, SpreadsheetOrdinal::Row)?;
    let hidden = attr_bool(node, "hidden").unwrap_or(false);
    // ECMA-376 §18.3.1.73 `<row>@ht` is the row height in points.
    let height = authored_row_height(hidden, node.attribute("ht"));
    let custom_height = attr_bool(node, "customHeight").unwrap_or(false);
    let outline_level = node
        .attribute("outlineLevel")
        .and_then(|s| s.parse::<u8>().ok())
        .unwrap_or(0)
        .min(7);
    let collapsed = attr_bool(node, "collapsed").unwrap_or(false);
    let row_ph = attr_bool(node, "ph").unwrap_or(false);
    let cells = parse_row_cells(node, row_idx, row_ph, shared_strings, theme_colors)?;
    Ok(Row {
        index: row_idx,
        height,
        custom_height,
        cells,
        outline_level,
        collapsed,
        hidden,
    })
}

fn derive_streamed_namespace_context(
    element: &quick_xml::events::BytesStart<'_>,
    inherited: &Rc<StreamedNamespaceContext>,
    part: Option<&str>,
) -> Result<Rc<StreamedNamespaceContext>, WorksheetProjectorError> {
    StreamedNamespaceContext::derive(element, inherited, STREAMED_XML_CONTEXT_BYTES, "worksheet")
        .map_err(|error| {
            map_bounded_xml_error(error, WorksheetProjectorLimitKind::XmlContextBytes, part)
        })
}

struct StreamedElementIdentity<'a> {
    namespace: Option<&'a str>,
    local_name: &'a str,
    visible: bool,
    part: Option<&'a str>,
}

fn classify_streamed_element<R>(
    reader: &NsReader<R>,
    element: &quick_xml::events::BytesStart<'_>,
    identity: StreamedElementIdentity<'_>,
    frames: &mut [StreamedElementFrame],
    namespace_context: &Rc<StreamedNamespaceContext>,
) -> Result<(StreamedElementKind, Rc<StreamedMceScope>), WorksheetProjectorError> {
    let StreamedElementIdentity {
        namespace,
        local_name,
        visible,
        part,
    } = identity;
    let inherited = frames
        .last()
        .map(|frame| Rc::clone(&frame.scope))
        .unwrap_or_else(StreamedMceScope::root);
    if !visible {
        return Ok((StreamedElementKind::Ignored, inherited));
    }
    if frames.last().is_some_and(StreamedElementFrame::is_opaque) {
        return Ok((
            StreamedElementKind::Retained {
                namespace: namespace.map(str::to_string),
                local_name: local_name.to_string(),
                opaque: true,
            },
            inherited,
        ));
    }

    let is_mc = namespace == Some(MCE_NS);
    let parent_is_alternate_content = frames
        .last()
        .is_some_and(|frame| matches!(frame.kind, StreamedElementKind::AlternateContent { .. }));
    if parent_is_alternate_content
        && worksheet_is_application_defined_extension_element(namespace, local_name)
    {
        return Err(WorksheetProjectorError::Xml(
            "worksheet MCE AlternateContent may contain only Choice/Fallback children after Ignorable processing"
                .to_string(),
        ));
    }
    if worksheet_is_application_defined_extension_element(namespace, local_name)
        && !parent_is_alternate_content
    {
        // Part 3 §§8 and 9.1: the extension element itself, including its
        // attributes, is passed through without MCE processing.
        return Ok((
            StreamedElementKind::Retained {
                namespace: namespace.map(str::to_string),
                local_name: local_name.to_string(),
                opaque: true,
            },
            inherited,
        ));
    }

    let attributes = bounded_xml::derive_mce_attributes(
        reader,
        element,
        &inherited,
        namespace_context,
        STREAMED_XML_CONTEXT_BYTES,
        "worksheet",
    )
    .map_err(|error| {
        map_bounded_xml_error(error, WorksheetProjectorLimitKind::XmlContextBytes, part)
    })?;
    if is_mc && matches!(local_name, "AlternateContent" | "Choice" | "Fallback") {
        bounded_xml::validate_mce_alternate_element_attributes(
            reader,
            element,
            local_name,
            &attributes.scope,
            "worksheet",
        )?;
    }
    if parent_is_alternate_content {
        if !is_mc || !matches!(local_name, "Choice" | "Fallback") {
            let ignored = namespace.is_some_and(|namespace| {
                attributes.scope.is_ignorable(namespace)
                    && !worksheet_understands_ns(namespace)
                    && !attributes.scope.processes_content(namespace, local_name)
            });
            if ignored {
                return Ok((StreamedElementKind::Ignored, attributes.scope));
            }
            return Err(WorksheetProjectorError::Xml(
                "worksheet MCE AlternateContent may contain only Choice/Fallback children after Ignorable processing"
                    .to_string(),
            ));
        }

        let alternate = frames.last_mut().expect("AlternateContent parent checked");
        let StreamedElementKind::AlternateContent {
            selected_branch,
            seen_choice,
            seen_fallback,
        } = &mut alternate.kind
        else {
            unreachable!("AlternateContent parent checked")
        };
        let selected = if local_name == "Choice" {
            if *seen_fallback {
                return Err(
                    "worksheet MCE Choice cannot follow Fallback in AlternateContent"
                        .to_string()
                        .into(),
                );
            }
            *seen_choice = true;
            !*selected_branch && streamed_choice_is_understood(element, namespace_context, part)?
        } else {
            if !*seen_choice {
                return Err(WorksheetProjectorError::Xml(
                    "worksheet MCE AlternateContent must contain at least one Choice before Fallback"
                        .to_string(),
                ));
            }
            if *seen_fallback {
                return Err(
                    "worksheet MCE AlternateContent may contain at most one Fallback"
                        .to_string()
                        .into(),
                );
            }
            *seen_fallback = true;
            !*selected_branch
        };
        if selected {
            *selected_branch = true;
            bounded_xml::validate_mce_must_understand(
                &attributes.must_understand,
                &worksheet_understands_ns,
                "worksheet",
            )?;
        }
        return Ok((
            StreamedElementKind::AlternateBranch { selected },
            attributes.scope,
        ));
    }

    if is_mc && local_name == "AlternateContent" {
        bounded_xml::validate_mce_must_understand(
            &attributes.must_understand,
            &worksheet_understands_ns,
            "worksheet",
        )?;
        return Ok((
            StreamedElementKind::AlternateContent {
                selected_branch: false,
                seen_choice: false,
                seen_fallback: false,
            },
            attributes.scope,
        ));
    }

    let ignored_namespace = namespace.filter(|namespace| {
        attributes.scope.is_ignorable(namespace) && !worksheet_understands_ns(namespace)
    });
    if let Some(namespace) = ignored_namespace {
        if attributes.scope.processes_content(namespace, local_name) {
            bounded_xml::validate_mce_process_content_element(reader, element, "worksheet")?;
            bounded_xml::validate_mce_must_understand(
                &attributes.must_understand,
                &worksheet_understands_ns,
                "worksheet",
            )?;
            return Ok((StreamedElementKind::Unwrapped, attributes.scope));
        }
        return Ok((StreamedElementKind::Ignored, attributes.scope));
    }

    bounded_xml::validate_mce_must_understand(
        &attributes.must_understand,
        &worksheet_understands_ns,
        "worksheet",
    )?;
    Ok((
        StreamedElementKind::Retained {
            namespace: namespace.map(str::to_string),
            local_name: local_name.to_string(),
            opaque: false,
        },
        attributes.scope,
    ))
}

struct StreamedProcessedParent {
    depth: usize,
    namespace: Option<String>,
    local_name: String,
    namespace_context: Rc<StreamedNamespaceContext>,
}

fn retained_streamed_parent(frames: &[StreamedElementFrame]) -> Option<StreamedProcessedParent> {
    frames
        .iter()
        .enumerate()
        .rev()
        .find_map(|(index, frame)| match &frame.kind {
            StreamedElementKind::Retained {
                namespace,
                local_name,
                ..
            } if frame.visible => Some(StreamedProcessedParent {
                depth: index + 1,
                namespace: namespace.clone(),
                local_name: local_name.clone(),
                namespace_context: Rc::clone(&frame.namespace_context),
            }),
            _ => None,
        })
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum StreamedHostRole {
    Worksheet,
    SheetData,
    Row,
    Other,
}

/// Validate one retained element against the worksheet host schema after MCE
/// logical reparenting. Start and empty-element events share this decision;
/// callers only differ in how they stream a non-empty sheetData or row.
/// Descendants of an application-defined extension element are opaque payload,
/// not host-schema children; the `extLst` boundary itself is still validated.
fn validate_streamed_host_element(
    namespace: Option<&str>,
    local_name: &str,
    opaque: &mut bool,
    inside_opaque_extension: bool,
    parent: Option<&StreamedProcessedParent>,
    processed_root_count: &mut usize,
    sheet_data_count: &mut usize,
) -> Result<StreamedHostRole, String> {
    if inside_opaque_extension {
        return Ok(StreamedHostRole::Other);
    }
    let is_spreadsheetml = namespace.is_some_and(|namespace| is_x_ns(Some(namespace)));
    *opaque = *opaque || worksheet_is_application_defined_extension_element(namespace, local_name);

    match parent {
        None => {
            *processed_root_count += 1;
            if *processed_root_count > 1 || !is_spreadsheetml || local_name != "worksheet" {
                return Err(
                    "MCE-processed worksheet root must be exactly one SpreadsheetML worksheet"
                        .to_string(),
                );
            }
            Ok(StreamedHostRole::Worksheet)
        }
        Some(parent)
            if parent
                .namespace
                .as_deref()
                .is_some_and(|namespace| is_x_ns(Some(namespace)))
                && parent.local_name == "worksheet" =>
        {
            if is_spreadsheetml && local_name == "sheetData" {
                *sheet_data_count += 1;
                if *sheet_data_count > 1 {
                    return Err(
                        "MCE-processed worksheet must contain exactly one non-nested sheetData"
                            .to_string(),
                    );
                }
                Ok(StreamedHostRole::SheetData)
            } else {
                Ok(StreamedHostRole::Other)
            }
        }
        Some(parent)
            if parent
                .namespace
                .as_deref()
                .is_some_and(|namespace| is_x_ns(Some(namespace)))
                && parent.local_name == "sheetData" =>
        {
            if is_spreadsheetml && local_name == "sheetData" {
                return Err(
                    "MCE-processed worksheet sheetData must be a direct child of worksheet"
                        .to_string(),
                );
            }
            if !is_spreadsheetml || local_name != "row" {
                return Err(
                    "MCE-processed sheetData may contain only direct SpreadsheetML row children"
                        .to_string(),
                );
            }
            Ok(StreamedHostRole::Row)
        }
        _ if is_spreadsheetml && local_name == "sheetData" => {
            Err("MCE-processed worksheet sheetData must be a direct child of worksheet".to_string())
        }
        _ => Ok(StreamedHostRole::Other),
    }
}

#[derive(Default)]
struct StreamedRowBatch {
    pending: Vec<PendingStreamedRow>,
    pending_bytes: usize,
    max_row_arena_bytes: usize,
    max_single_row_projection_bytes: usize,
    previous_index: u32,
    geometry: AuthoredRowGeometry,
}

struct RowProjectionInputs<'a> {
    shared_strings: &'a [SharedString],
    theme_colors: &'a [String],
}

impl StreamedRowBatch {
    fn push(&mut self, row: PendingStreamedRow) {
        if self.pending.is_empty() {
            self.pending_bytes = STREAMED_ROWS_ARENA_OVERHEAD;
        }
        self.pending_bytes = self
            .pending_bytes
            .saturating_add(row.projected_arena_bytes());
        self.max_single_row_projection_bytes = self
            .max_single_row_projection_bytes
            .max(row.projected_arena_bytes());
        self.max_row_arena_bytes = self.max_row_arena_bytes.max(self.pending_bytes);
        self.pending.push(row);
    }

    fn would_exceed_byte_ceiling(&self, row: &PendingStreamedRow) -> bool {
        !self.pending.is_empty()
            && self
                .pending_bytes
                .saturating_add(row.projected_arena_bytes())
                > STREAMED_ROW_BATCH_BYTES
    }

    fn should_dispatch(&self) -> bool {
        self.pending_bytes >= STREAMED_ROW_BATCH_BYTES
            || self.pending.len() >= STREAMED_ROW_BATCH_ROWS
    }

    fn dispatch(
        &mut self,
        shared_strings: &[SharedString],
        theme_colors: &[String],
        ready_rows: &mut VecDeque<ProjectedWorksheetRow>,
    ) -> Result<(), WorksheetProjectorError> {
        for row in flush_streamed_rows(
            &mut self.pending,
            &mut self.previous_index,
            shared_strings,
            theme_colors,
        )? {
            self.geometry.record(row.row.index, row.row.height);
            ready_rows.push_back(row);
        }
        self.pending_bytes = 0;
        Ok(())
    }

    fn push_bounded(
        &mut self,
        row: PendingStreamedRow,
        row_projection_limit: usize,
        part: Option<&str>,
        inputs: RowProjectionInputs<'_>,
        ready_rows: &mut VecDeque<ProjectedWorksheetRow>,
    ) -> Result<(), WorksheetProjectorError> {
        let projected_bytes = row.projected_arena_bytes();
        if projected_bytes > row_projection_limit {
            return Err(worksheet_projector_limit(
                WorksheetProjectorLimitKind::RowProjectionBytes,
                part,
                row_projection_limit,
                projected_bytes,
            ));
        }
        // Account for the generated projection and inherited namespace wrapper,
        // not just raw source bytes. Dispatch before crossing the deterministic
        // batch-arena ceiling. A single row may exceed that batching target but
        // has already passed the separate hard row-projection invariant above.
        if self.would_exceed_byte_ceiling(&row) {
            self.dispatch(inputs.shared_strings, inputs.theme_colors, ready_rows)?;
        }
        self.push(row);
        if self.should_dispatch() {
            self.dispatch(inputs.shared_strings, inputs.theme_colors, ready_rows)?;
        }
        Ok(())
    }
}

#[derive(Debug)]
pub(super) struct StreamedWorksheetRows {
    pub(super) shell_xml: String,
    pub(super) row_geometry: AuthoredRowGeometry,
    #[cfg(test)]
    max_row_arena_bytes: usize,
    #[cfg(test)]
    max_single_row_projection_bytes: usize,
}

#[derive(Debug)]
pub(super) enum WorksheetProjectorItem {
    Row(ProjectedWorksheetRow),
    Finished(StreamedWorksheetRows),
}

#[derive(Debug)]
pub(super) struct ProjectedWorksheetRow {
    pub(super) row: Row,
    /// Exact bytes retained for the row's standalone internal projection,
    /// including inherited namespace wrapper overhead.
    pub(super) projected_bytes: usize,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum WorksheetProjectorState {
    Reading,
    Draining,
    Finished,
    Failed,
}

/// Resumable projection of one worksheet XML stream into parsed rows and a
/// processed worksheet shell. The same state machine backs both bounded ZIP
/// entry readers and the historical `&str` compatibility adapter.
pub(super) struct WorksheetRowProjector<R, S, T>
where
    R: BufRead,
    S: AsRef<[SharedString]>,
    T: AsRef<[String]>,
{
    reader: Option<BoundedXmlReader<R>>,
    frames: Vec<StreamedElementFrame>,
    processed_root_count: usize,
    sheet_data_count: usize,
    sheet_data_depth: Option<usize>,
    shell_xml: Vec<u8>,
    active_row: Option<ActiveStreamedRow>,
    row_batch: StreamedRowBatch,
    ready_rows: VecDeque<ProjectedWorksheetRow>,
    finished_tail: Option<StreamedWorksheetRows>,
    shared_strings: Option<S>,
    theme_colors: Option<T>,
    part: Option<String>,
    limit_reporter: Option<PackageLimitReporter>,
    row_projection_limit: usize,
    shell_projection_limit: usize,
    state: WorksheetProjectorState,
    #[cfg(test)]
    max_event_bytes_observed: usize,
    #[cfg(test)]
    max_context_bytes_observed: usize,
}

impl<R, S, T> WorksheetRowProjector<R, S, T>
where
    R: BufRead,
    S: AsRef<[SharedString]>,
    T: AsRef<[String]>,
{
    #[cfg(test)]
    fn new(source: R, shared_strings: S, theme_colors: T) -> Self {
        Self::with_limits(
            source,
            shared_strings,
            theme_colors,
            STREAMED_XML_EVENT_BYTES,
            STREAMED_ROW_PROJECTION_BYTES,
            STREAMED_WORKSHEET_SHELL_BYTES,
            None,
        )
    }

    #[cfg(test)]
    fn with_limits(
        source: R,
        shared_strings: S,
        theme_colors: T,
        event_limit: usize,
        row_projection_limit: usize,
        shell_projection_limit: usize,
        part: Option<String>,
    ) -> Self {
        Self::with_limits_and_reporter(
            source,
            shared_strings,
            theme_colors,
            event_limit,
            row_projection_limit,
            shell_projection_limit,
            part,
            None,
        )
    }

    #[allow(clippy::too_many_arguments)]
    fn with_limits_and_reporter(
        source: R,
        shared_strings: S,
        theme_colors: T,
        event_limit: usize,
        row_projection_limit: usize,
        shell_projection_limit: usize,
        part: Option<String>,
        limit_reporter: Option<PackageLimitReporter>,
    ) -> Self {
        let reader = BoundedXmlReader::new(source, event_limit, "worksheet");
        Self {
            reader: Some(reader),
            frames: Vec::new(),
            processed_root_count: 0,
            sheet_data_count: 0,
            sheet_data_depth: None,
            shell_xml: Vec::new(),
            active_row: None,
            row_batch: StreamedRowBatch::default(),
            ready_rows: VecDeque::new(),
            finished_tail: None,
            shared_strings: Some(shared_strings),
            theme_colors: Some(theme_colors),
            part,
            limit_reporter,
            row_projection_limit,
            shell_projection_limit,
            state: WorksheetProjectorState::Reading,
            #[cfg(test)]
            max_event_bytes_observed: 0,
            #[cfg(test)]
            max_context_bytes_observed: 0,
        }
    }

    fn append_to_projection(&mut self, event: &Event<'_>) -> Result<(), WorksheetProjectorError> {
        if let Some(row) = self.active_row.as_mut() {
            append_projected_event(
                &mut row.xml,
                event,
                row.projection_overhead,
                self.row_projection_limit,
                WorksheetProjectorLimitKind::RowProjectionBytes,
                self.part.as_deref(),
            )
        } else if self.sheet_data_depth.is_none() {
            append_projected_event(
                &mut self.shell_xml,
                event,
                0,
                self.shell_projection_limit,
                WorksheetProjectorLimitKind::WorksheetShellBytes,
                self.part.as_deref(),
            )
        } else {
            Ok(())
        }
    }

    fn admit_row(&mut self, row: PendingStreamedRow) -> Result<(), WorksheetProjectorError> {
        self.row_batch.push_bounded(
            row,
            self.row_projection_limit,
            self.part.as_deref(),
            RowProjectionInputs {
                shared_strings: self
                    .shared_strings
                    .as_ref()
                    .expect("active projector owns shared strings")
                    .as_ref(),
                theme_colors: self
                    .theme_colors
                    .as_ref()
                    .expect("active projector owns theme colors")
                    .as_ref(),
            },
            &mut self.ready_rows,
        )
    }

    fn read_event(&mut self) -> Result<(Option<String>, Event<'static>), WorksheetProjectorError> {
        let reader = self
            .reader
            .as_mut()
            .expect("reading projector owns an XML reader");
        let read = match reader.read_event() {
            Ok(read) => read,
            Err(BoundedXmlReadError::EventLimit { limit, observed }) => {
                #[cfg(test)]
                {
                    self.max_event_bytes_observed = reader.max_event_bytes();
                }
                return Err(worksheet_projector_limit(
                    WorksheetProjectorLimitKind::XmlEventBytes,
                    self.part.as_deref(),
                    limit,
                    observed,
                ));
            }
            Err(BoundedXmlReadError::Xml(error)) => {
                #[cfg(test)]
                {
                    self.max_event_bytes_observed = reader.max_event_bytes();
                }
                return Err(WorksheetProjectorError::Xml(error));
            }
        };
        #[cfg(test)]
        {
            self.max_event_bytes_observed = reader.max_event_bytes();
        }
        Ok((read.namespace, read.event))
    }

    fn process_start(
        &mut self,
        namespace: Option<String>,
        mut start: quick_xml::events::BytesStart<'static>,
    ) -> Result<(), WorksheetProjectorError> {
        let depth = self.frames.len() + 1;
        if depth > MAX_XML_DEPTH as usize {
            return Err(worksheet_projector_limit(
                WorksheetProjectorLimitKind::XmlDepth,
                self.part.as_deref(),
                MAX_XML_DEPTH as usize,
                depth,
            ));
        }
        let local_name = std::str::from_utf8(start.local_name().as_ref())
            .map_err(|error| error.to_string())?
            .to_string();
        let visible = self
            .frames
            .last()
            .is_none_or(StreamedElementFrame::children_are_visible);
        let inside_opaque_extension = self
            .frames
            .last()
            .is_some_and(StreamedElementFrame::is_opaque);
        let physical_parent_depth = self.frames.len();
        let processed_parent = retained_streamed_parent(&self.frames);
        let inherited_namespace_context = self
            .frames
            .last()
            .map(|frame| Rc::clone(&frame.namespace_context))
            .unwrap_or_else(|| Rc::new(StreamedNamespaceContext::default()));
        let namespace_context = derive_streamed_namespace_context(
            &start,
            &inherited_namespace_context,
            self.part.as_deref(),
        )?;
        let reader = self
            .reader
            .as_ref()
            .expect("reading projector owns an XML reader")
            .reader();
        let (mut kind, scope) = classify_streamed_element(
            reader,
            &start,
            StreamedElementIdentity {
                namespace: namespace.as_deref(),
                local_name: &local_name,
                visible,
                part: self.part.as_deref(),
            },
            &mut self.frames,
            &namespace_context,
        )?;
        #[cfg(test)]
        {
            self.max_context_bytes_observed = self.max_context_bytes_observed.max(
                namespace_context
                    .active_bytes()
                    .saturating_add(scope.active_bytes()),
            );
        }
        let mut role = None;

        if visible {
            if let StreamedElementKind::Retained {
                namespace: retained_namespace,
                local_name: retained_local_name,
                opaque,
            } = &mut kind
            {
                let retained_role = validate_streamed_host_element(
                    retained_namespace.as_deref(),
                    retained_local_name,
                    opaque,
                    inside_opaque_extension,
                    processed_parent.as_ref(),
                    &mut self.processed_root_count,
                    &mut self.sheet_data_count,
                )?;
                let processed_parent_depth =
                    processed_parent.as_ref().map_or(0, |parent| parent.depth);
                if processed_parent_depth != physical_parent_depth {
                    inject_moved_element_namespaces(
                        &mut start,
                        &namespace_context,
                        processed_parent
                            .as_ref()
                            .map(|parent| parent.namespace_context.as_ref()),
                        self.part.as_deref(),
                    )?;
                }
                if retained_role == StreamedHostRole::Row {
                    let element_namespaces = namespace_context.effective_bindings();
                    if self.active_row.is_some() {
                        return Err(WorksheetProjectorError::Xml(
                            "streamed worksheet encountered a nested active row".to_string(),
                        ));
                    }
                    self.active_row = Some(ActiveStreamedRow {
                        depth,
                        namespaces: element_namespaces.clone(),
                        projection_overhead: PendingStreamedRow {
                            xml: Vec::new(),
                            namespaces: element_namespaces.clone(),
                        }
                        .projected_arena_bytes(),
                        xml: Vec::new(),
                    });
                }
                if !*opaque {
                    bounded_xml::strip_processed_mce_attributes(
                        reader,
                        &mut start,
                        &scope,
                        &worksheet_understands_ns,
                        "worksheet",
                    )?;
                }
                role = Some(retained_role);
            }
        }

        if matches!(kind, StreamedElementKind::Retained { .. }) && visible {
            self.append_to_projection(&Event::Start(start))?;
            if role == Some(StreamedHostRole::SheetData) {
                self.sheet_data_depth = Some(depth);
            }
        }
        self.frames.push(StreamedElementFrame {
            kind,
            scope,
            namespace_context,
            visible,
        });
        Ok(())
    }

    fn process_empty(
        &mut self,
        namespace: Option<String>,
        mut empty: quick_xml::events::BytesStart<'static>,
    ) -> Result<(), WorksheetProjectorError> {
        let local_name = std::str::from_utf8(empty.local_name().as_ref())
            .map_err(|error| error.to_string())?
            .to_string();
        let visible = self
            .frames
            .last()
            .is_none_or(StreamedElementFrame::children_are_visible);
        let inside_opaque_extension = self
            .frames
            .last()
            .is_some_and(StreamedElementFrame::is_opaque);
        let physical_parent_depth = self.frames.len();
        let processed_parent = retained_streamed_parent(&self.frames);
        let inherited_namespace_context = self
            .frames
            .last()
            .map(|frame| Rc::clone(&frame.namespace_context))
            .unwrap_or_else(|| Rc::new(StreamedNamespaceContext::default()));
        let namespace_context = derive_streamed_namespace_context(
            &empty,
            &inherited_namespace_context,
            self.part.as_deref(),
        )?;
        let reader = self
            .reader
            .as_ref()
            .expect("reading projector owns an XML reader")
            .reader();
        let (mut kind, _scope) = classify_streamed_element(
            reader,
            &empty,
            StreamedElementIdentity {
                namespace: namespace.as_deref(),
                local_name: &local_name,
                visible,
                part: self.part.as_deref(),
            },
            &mut self.frames,
            &namespace_context,
        )?;
        #[cfg(test)]
        {
            self.max_context_bytes_observed = self.max_context_bytes_observed.max(
                namespace_context
                    .active_bytes()
                    .saturating_add(_scope.active_bytes()),
            );
        }

        match &mut kind {
            StreamedElementKind::Retained {
                namespace: retained_namespace,
                local_name: retained_local_name,
                opaque,
            } if visible => {
                let role = validate_streamed_host_element(
                    retained_namespace.as_deref(),
                    retained_local_name,
                    opaque,
                    inside_opaque_extension,
                    processed_parent.as_ref(),
                    &mut self.processed_root_count,
                    &mut self.sheet_data_count,
                )?;
                let processed_parent_depth =
                    processed_parent.as_ref().map_or(0, |parent| parent.depth);
                if processed_parent_depth != physical_parent_depth {
                    inject_moved_element_namespaces(
                        &mut empty,
                        &namespace_context,
                        processed_parent
                            .as_ref()
                            .map(|parent| parent.namespace_context.as_ref()),
                        self.part.as_deref(),
                    )?;
                }
                if !*opaque {
                    bounded_xml::strip_processed_mce_attributes(
                        reader,
                        &mut empty,
                        &_scope,
                        &worksheet_understands_ns,
                        "worksheet",
                    )?;
                }
                if role == StreamedHostRole::Row {
                    let element_namespaces = namespace_context.effective_bindings();
                    let mut xml = Vec::new();
                    let projection_overhead = PendingStreamedRow {
                        xml: Vec::new(),
                        namespaces: element_namespaces.clone(),
                    }
                    .projected_arena_bytes();
                    append_projected_event(
                        &mut xml,
                        &Event::Empty(empty),
                        projection_overhead,
                        self.row_projection_limit,
                        WorksheetProjectorLimitKind::RowProjectionBytes,
                        self.part.as_deref(),
                    )?;
                    self.admit_row(PendingStreamedRow {
                        xml,
                        namespaces: element_namespaces,
                    })?;
                } else {
                    self.append_to_projection(&Event::Empty(empty))?;
                }
            }
            StreamedElementKind::Retained { .. } => {}
            StreamedElementKind::AlternateContent {
                seen_choice: false, ..
            } => {
                return Err(WorksheetProjectorError::Xml(
                    "worksheet MCE AlternateContent must contain at least one Choice".to_string(),
                ));
            }
            _ => {}
        }
        Ok(())
    }

    fn process_end(
        &mut self,
        end: quick_xml::events::BytesEnd<'static>,
    ) -> Result<(), WorksheetProjectorError> {
        let depth = self.frames.len();
        let frame = self.frames.pop().ok_or_else(|| {
            WorksheetProjectorError::Xml(
                "worksheet XML stream ended an element without a matching start".to_string(),
            )
        })?;
        let completed_row = self.active_row.as_ref().is_some_and(|row| {
            row.depth == depth
                && matches!(
                    &frame.kind,
                    StreamedElementKind::Retained {
                        namespace,
                        local_name,
                        ..
                    } if namespace
                        .as_deref()
                        .is_some_and(|namespace| is_x_ns(Some(namespace)))
                        && local_name == "row"
                )
        });
        let completed_sheet_data = self.sheet_data_depth == Some(depth)
            && matches!(
                &frame.kind,
                StreamedElementKind::Retained {
                    namespace,
                    local_name,
                    ..
                } if namespace
                    .as_deref()
                    .is_some_and(|namespace| is_x_ns(Some(namespace)))
                    && local_name == "sheetData"
            );

        match frame.kind {
            StreamedElementKind::Retained { .. } => {
                if completed_sheet_data {
                    self.sheet_data_depth = None;
                }
                self.append_to_projection(&Event::End(end))?;
            }
            StreamedElementKind::AlternateContent {
                seen_choice: false, ..
            } => {
                return Err(WorksheetProjectorError::Xml(
                    "worksheet MCE AlternateContent must contain at least one Choice".to_string(),
                ));
            }
            StreamedElementKind::Unwrapped
            | StreamedElementKind::Ignored
            | StreamedElementKind::AlternateBranch { .. }
            | StreamedElementKind::AlternateContent {
                seen_choice: true, ..
            } => {}
        }

        if completed_row {
            let row = self.active_row.take().expect("completed row is active");
            self.admit_row(PendingStreamedRow {
                xml: row.xml,
                namespaces: row.namespaces,
            })?;
        }
        Ok(())
    }

    fn finalize(&mut self) -> Result<(), WorksheetProjectorError> {
        if !self.frames.is_empty() {
            return Err(WorksheetProjectorError::Xml(
                "worksheet XML stream ended with unclosed elements".to_string(),
            ));
        }
        if self.active_row.is_some() || self.sheet_data_depth.is_some() {
            return Err(WorksheetProjectorError::Xml(
                "worksheet XML stream ended with unfinished streamed content".to_string(),
            ));
        }
        if self.processed_root_count != 1 {
            return Err(WorksheetProjectorError::Xml(
                "MCE-processed worksheet root must be exactly one SpreadsheetML worksheet"
                    .to_string(),
            ));
        }
        if self.sheet_data_count != 1 {
            return Err(WorksheetProjectorError::Xml(
                "MCE-processed worksheet must contain exactly one sheetData".to_string(),
            ));
        }
        self.row_batch.dispatch(
            self.shared_strings
                .as_ref()
                .expect("active projector owns shared strings")
                .as_ref(),
            self.theme_colors
                .as_ref()
                .expect("active projector owns theme colors")
                .as_ref(),
            &mut self.ready_rows,
        )?;
        self.reader.take();
        self.shared_strings.take();
        self.theme_colors.take();
        let shell_xml = String::from_utf8(std::mem::take(&mut self.shell_xml))
            .map_err(|error| format!("streamed worksheet shell is not UTF-8: {error}"))?;
        #[cfg(test)]
        let max_row_arena_bytes = self.row_batch.max_row_arena_bytes;
        #[cfg(test)]
        let max_single_row_projection_bytes = self.row_batch.max_single_row_projection_bytes;
        self.finished_tail = Some(StreamedWorksheetRows {
            shell_xml,
            row_geometry: std::mem::take(&mut self.row_batch.geometry),
            #[cfg(test)]
            max_row_arena_bytes,
            #[cfg(test)]
            max_single_row_projection_bytes,
        });
        self.state = WorksheetProjectorState::Draining;
        Ok(())
    }

    fn next_item_inner(&mut self) -> Result<WorksheetProjectorItem, WorksheetProjectorError> {
        if let Some(row) = self.ready_rows.pop_front() {
            return Ok(WorksheetProjectorItem::Row(row));
        }
        if self.state == WorksheetProjectorState::Draining {
            self.state = WorksheetProjectorState::Finished;
            return Ok(WorksheetProjectorItem::Finished(
                self.finished_tail
                    .take()
                    .expect("draining projector has a finished tail"),
            ));
        }
        if self.state == WorksheetProjectorState::Finished {
            return Err(WorksheetProjectorError::Xml(
                "worksheet projector is already finished".to_string(),
            ));
        }
        if self.state == WorksheetProjectorState::Failed {
            return Err(WorksheetProjectorError::Xml(
                "worksheet projector transaction has already failed".to_string(),
            ));
        }

        loop {
            let (namespace, event) = self.read_event()?;
            match event {
                Event::Start(start) => self.process_start(namespace, start)?,
                Event::Empty(empty) => self.process_empty(namespace, empty)?,
                Event::End(end) => self.process_end(end)?,
                Event::Eof => {
                    self.finalize()?;
                    if let Some(row) = self.ready_rows.pop_front() {
                        return Ok(WorksheetProjectorItem::Row(row));
                    }
                    self.state = WorksheetProjectorState::Finished;
                    return Ok(WorksheetProjectorItem::Finished(
                        self.finished_tail
                            .take()
                            .expect("finalized projector has a finished tail"),
                    ));
                }
                other => {
                    let visible = self
                        .frames
                        .last()
                        .is_none_or(StreamedElementFrame::children_are_visible);
                    if visible {
                        self.append_to_projection(&other)?;
                    }
                }
            }
            if let Some(row) = self.ready_rows.pop_front() {
                return Ok(WorksheetProjectorItem::Row(row));
            }
        }
    }

    fn discard_failed_transaction(&mut self) {
        // Taking the reader drops a future PackageEntryStream immediately, so
        // its decoder/lease is released even when the failed projector handle
        // itself remains stored for a later diagnostic call.
        self.reader.take();
        self.frames = Vec::new();
        self.active_row = None;
        self.sheet_data_depth = None;
        self.shell_xml = Vec::new();
        self.row_batch = StreamedRowBatch::default();
        self.ready_rows = VecDeque::new();
        self.finished_tail = None;
        self.shared_strings.take();
        self.theme_colors.take();
    }

    pub(super) fn next_item(&mut self) -> Result<WorksheetProjectorItem, WorksheetProjectorError> {
        let result = self
            .next_item_inner()
            .map_err(|error| report_projector_limit(self.limit_reporter.as_ref(), error));
        if result.is_err() {
            self.state = WorksheetProjectorState::Failed;
            self.discard_failed_transaction();
        }
        result
    }

    #[cfg(test)]
    fn max_event_bytes(&self) -> usize {
        self.max_event_bytes_observed
    }

    #[cfg(test)]
    fn max_context_bytes(&self) -> usize {
        self.max_context_bytes_observed
    }
}

impl<S, T> WorksheetRowProjector<BufReader<PackageEntryStream>, S, T>
where
    S: AsRef<[SharedString]>,
    T: AsRef<[String]>,
{
    /// Construct the production projector from one operation-bound package
    /// entry. The reporter is derived before the stream moves into the reader,
    /// so parser limits and inflation can never be charged to different
    /// operations.
    pub(super) fn from_package_entry(
        entry: PackageEntryStream,
        shared_strings: S,
        theme_colors: T,
    ) -> Result<Self, String> {
        let part = entry.part_name().to_string();
        let reporter = entry.limit_reporter()?;
        Ok(Self::with_limits_and_reporter(
            BufReader::new(entry),
            shared_strings,
            theme_colors,
            STREAMED_XML_EVENT_BYTES,
            STREAMED_ROW_PROJECTION_BYTES,
            STREAMED_WORKSHEET_SHELL_BYTES,
            Some(part),
            Some(reporter),
        ))
    }
}

impl<S, T> WorksheetRowProjector<BufReader<Box<dyn Read>>, S, T>
where
    S: AsRef<[SharedString]>,
    T: AsRef<[String]>,
{
    pub(super) fn from_owned_reader(
        source: Box<dyn Read>,
        part: String,
        reporter: PackageLimitReporter,
        shared_strings: S,
        theme_colors: T,
    ) -> Self {
        Self::with_limits_and_reporter(
            BufReader::new(source),
            shared_strings,
            theme_colors,
            STREAMED_XML_EVENT_BYTES,
            STREAMED_ROW_PROJECTION_BYTES,
            STREAMED_WORKSHEET_SHELL_BYTES,
            Some(part),
            Some(reporter),
        )
    }
}

/// Stream the MCE-processed `CT_SheetData` rows through a bounded batch parser.
///
/// ECMA-376 Part 3 §§9.2–9.4 define the processed infoset against which the
/// host schema is validated. The visitor therefore sees only rows whose
/// `AlternateContent` branch is selected and rows exposed by an ignorable
/// `ProcessContent` wrapper. Rows are delivered immediately and are provisional
/// until this function returns `Ok`; stateful callers must commit transactionally.
#[cfg(test)]
pub(super) fn stream_worksheet_rows(
    xml: &str,
    shared_strings: &[SharedString],
    theme_colors: &[String],
    visit_row: impl FnMut(Row) -> Result<(), String>,
) -> Result<StreamedWorksheetRows, String> {
    stream_worksheet_rows_with_part(xml, shared_strings, theme_colors, None, visit_row)
}

#[cfg(test)]
fn stream_worksheet_rows_with_part(
    xml: &str,
    shared_strings: &[SharedString],
    theme_colors: &[String],
    part: Option<&str>,
    visit_row: impl FnMut(Row) -> Result<(), String>,
) -> Result<StreamedWorksheetRows, String> {
    let projector = WorksheetRowProjector::with_limits(
        Cursor::new(xml.as_bytes()),
        shared_strings,
        theme_colors,
        STREAMED_XML_EVENT_BYTES,
        STREAMED_ROW_PROJECTION_BYTES,
        STREAMED_WORKSHEET_SHELL_BYTES,
        part.map(str::to_string),
    );
    drain_worksheet_projector(projector, visit_row)
}

#[cfg(test)]
fn drain_worksheet_projector<R, S, T>(
    mut projector: WorksheetRowProjector<R, S, T>,
    mut visit_row: impl FnMut(Row) -> Result<(), String>,
) -> Result<StreamedWorksheetRows, String>
where
    R: BufRead,
    S: AsRef<[SharedString]>,
    T: AsRef<[String]>,
{
    // Rows are intentionally delivered as soon as a bounded batch is parsed.
    // They remain provisional until `Finished`; callers that retain state must
    // commit transactionally only after this function returns `Ok`.
    loop {
        match projector.next_item().map_err(|error| error.to_string())? {
            WorksheetProjectorItem::Row(row) => {
                // A callback failure drops the projector here, immediately
                // releasing the source and preventing later callbacks.
                visit_row(row.row)?
            }
            WorksheetProjectorItem::Finished(tail) => return Ok(tail),
        }
    }
}

#[cfg(test)]
pub(super) fn stream_sheet_data(
    xml: &str,
    shared_strings: &[SharedString],
    theme_colors: &[String],
) -> Result<StreamedSheetData, String> {
    let mut rows = Vec::new();
    let streamed =
        stream_worksheet_rows_with_part(xml, shared_strings, theme_colors, None, |row| {
            rows.push(row);
            Ok(())
        })?;
    Ok(StreamedSheetData {
        shell_xml: streamed.shell_xml,
        rows,
        row_geometry: streamed.row_geometry,
    })
}
#[cfg(test)]
mod worksheet_streaming_tests {
    use super::*;
    use crate::{extract_reference_cells, parse_worksheet, CellRange, CellValue};
    use ooxml_common::package_session::PackageSessionHandle;
    use ooxml_common::resource::OoxmlFormat;
    use std::cell::Cell;
    use std::io::BufReader;
    use std::io::Write as _;
    use std::rc::Rc;
    use zip::write::SimpleFileOptions;

    fn project_reader<R: BufRead>(source: R) -> Result<StreamedSheetData, WorksheetProjectorError> {
        let mut projector = WorksheetRowProjector::new(source, &[], &[]);
        let mut rows = Vec::new();
        loop {
            match projector.next_item()? {
                WorksheetProjectorItem::Row(row) => rows.push(row.row),
                WorksheetProjectorItem::Finished(tail) => {
                    return Ok(StreamedSheetData {
                        shell_xml: tail.shell_xml,
                        rows,
                        row_geometry: tail.row_geometry,
                    });
                }
            }
        }
    }

    #[test]
    fn row_height_presence_distinguishes_authored_height_from_auto_fit() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1"><c r="A1" t="inlineStr"><is><t>automatic</t></is></c></row>
            <row r="2" ht="30" customHeight="1"><c r="A2" t="inlineStr"><is><t>manual</t></is></c></row>
            <row r="3" ht="24" customHeight="0"><c r="A3" t="inlineStr"><is><t>authored ht still wins</t></is></c></row>
            <row r="4" customHeight="1"><c r="A4" t="inlineStr"><is><t>manual default</t></is></c></row>
          </sheetData>
        </worksheet>"#;
        let streamed = stream_sheet_data(xml, &[], &[]).expect("rows project");
        assert_eq!(streamed.rows[0].height, None);
        assert_eq!(streamed.rows[1].height, Some(30.0));
        // §18.3.1.73: customHeight describes how the height was set; it does
        // not erase an authored ht value when false.
        assert_eq!(streamed.rows[2].height, Some(24.0));
        assert!(!streamed.rows[0].custom_height);
        assert!(streamed.rows[1].custom_height);
        assert!(!streamed.rows[2].custom_height);
        assert_eq!(streamed.rows[3].height, None);
        assert!(streamed.rows[3].custom_height);
    }

    struct CountingRead<'a> {
        cursor: Cursor<&'a [u8]>,
        consumed: Rc<Cell<usize>>,
    }

    impl Read for CountingRead<'_> {
        fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
            let count = self.cursor.read(output)?;
            self.consumed.set(self.consumed.get() + count);
            Ok(count)
        }
    }

    fn dom_rows(
        xml: &str,
        shared_strings: &[SharedString],
        theme_colors: &[String],
    ) -> (Vec<Row>, AuthoredRowGeometry) {
        let doc = parse_guarded(xml).expect("legacy full worksheet DOM parses");
        let mut previous_row = 0;
        let mut heights = AuthoredRowGeometry::default();
        let rows = doc
            .descendants()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().name() == "row"
                    && is_x_ns(node.tag_name().namespace())
            })
            .map(|node| {
                let row = parse_row_node(&node, &mut previous_row, shared_strings, theme_colors)
                    .expect("reference worksheet row parses");
                heights.record(row.index, row.height);
                row
            })
            .collect();
        (rows, heights)
    }

    fn assert_streamed_rows_match_dom(xml: &str) {
        let streamed = stream_sheet_data(xml, &[], &[]).expect("streaming parse succeeds");
        let (dom_rows, dom_heights) = dom_rows(xml, &[], &[]);
        assert_eq!(
            serde_json::to_value(&streamed.rows).expect("streamed rows serialize"),
            serde_json::to_value(&dom_rows).expect("DOM rows serialize")
        );
        assert_eq!(streamed.row_geometry, dom_heights);
        let shell = parse_guarded(&streamed.shell_xml).expect("worksheet shell stays valid");
        assert!(!shell
            .descendants()
            .any(|node| node.tag_name().name() == "row"));
    }

    #[test]
    fn one_byte_reader_matches_compatibility_projection_across_mce_and_strict_xml() {
        let xml = r#"<x:worksheet
            xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
            xmlns:future="urn:example:future"
            mc:Ignorable="future"
            mc:ProcessContent="future:cells">
          <x:sheetData>
            <x:row r="4" ht="22.5">
              <mc:AlternateContent>
                <mc:Choice Requires="x14">
                  <x:c r="A4" t="inlineStr"><x:is><x:t xml:space="preserve"> A &amp; 漢字 </x:t></x:is></x:c>
                </mc:Choice>
                <mc:Fallback><x:c r="B4"><x:v>99</x:v></x:c></mc:Fallback>
              </mc:AlternateContent>
              <future:cells><x:c r="C4"><x:v><![CDATA[7]]></x:v></x:c></future:cells>
            </x:row>
          </x:sheetData>
          <x:mergeCells count="1">
            <mc:AlternateContent>
              <mc:Choice Requires="x14"><x:mergeCell ref="A4:C4"/></mc:Choice>
              <mc:Fallback><x:mergeCell ref="D4:E4"/></mc:Fallback>
            </mc:AlternateContent>
          </x:mergeCells>
        </x:worksheet>"#;
        let expected = stream_sheet_data(xml, &[], &[]).expect("compatibility projection");
        let actual = project_reader(BufReader::with_capacity(1, Cursor::new(xml.as_bytes())))
            .expect("one-byte reader projection");

        assert_eq!(
            serde_json::to_value(actual.rows).unwrap(),
            serde_json::to_value(expected.rows).unwrap()
        );
        assert_eq!(actual.row_geometry, expected.row_geometry);
        assert_eq!(actual.shell_xml, expected.shell_xml);
    }

    #[test]
    fn resumable_projector_yields_rows_before_reading_entry_eof() {
        const ROWS: usize = STREAMED_ROW_BATCH_ROWS + 64;
        let mut xml = String::from(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
        );
        for row in 1..=ROWS {
            write!(
                xml,
                r#"<row r="{row}"><c r="A{row}"><v>{row}</v></c></row>"#
            )
            .unwrap();
        }
        xml.push_str("</sheetData><mergeCells count=\"0\"/></worksheet>");

        let consumed = Rc::new(Cell::new(0));
        let reader = CountingRead {
            cursor: Cursor::new(xml.as_bytes()),
            consumed: consumed.clone(),
        };
        let mut projector =
            WorksheetRowProjector::new(BufReader::with_capacity(31, reader), &[], &[]);
        let first = projector.next_item().expect("first projected item");
        assert!(matches!(first, WorksheetProjectorItem::Row(_)));
        assert!(
            consumed.get() < xml.len(),
            "projector consumed the complete entry before yielding a row"
        );
    }

    #[test]
    fn oversized_xml_event_returns_structured_limit_before_buffer_growth() {
        let oversized_text = "x".repeat(STREAMED_XML_EVENT_BYTES + 64);
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c t="inlineStr"><is><t>{oversized_text}</t></is></c></row></sheetData></worksheet>"#
        );
        let mut projector = WorksheetRowProjector::new(
            BufReader::with_capacity(4096, Cursor::new(xml.as_bytes())),
            &[],
            &[],
        );
        let error = projector
            .next_item()
            .expect_err("oversized text event must be rejected");
        assert_eq!(
            error,
            WorksheetProjectorError::Limit {
                kind: WorksheetProjectorLimitKind::XmlEventBytes,
                part: None,
                limit: STREAMED_XML_EVENT_BYTES,
                observed: STREAMED_XML_EVENT_BYTES + 1,
            }
        );
        assert_eq!(projector.max_event_bytes(), STREAMED_XML_EVENT_BYTES);
    }

    #[test]
    fn oversized_start_tag_uses_the_same_structured_event_limit() {
        let oversized_attribute = "x".repeat(STREAMED_XML_EVENT_BYTES + 64);
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" oversized="{oversized_attribute}"><sheetData/></worksheet>"#
        );
        let mut projector = WorksheetRowProjector::new(
            BufReader::with_capacity(4096, Cursor::new(xml.as_bytes())),
            &[],
            &[],
        );
        assert_eq!(
            projector
                .next_item()
                .expect_err("oversized start tag must be rejected"),
            WorksheetProjectorError::Limit {
                kind: WorksheetProjectorLimitKind::XmlEventBytes,
                part: None,
                limit: STREAMED_XML_EVENT_BYTES,
                observed: STREAMED_XML_EVENT_BYTES + 1,
            }
        );
        assert_eq!(projector.max_event_bytes(), STREAMED_XML_EVENT_BYTES);
    }

    #[test]
    fn repeated_process_content_qnames_are_charged_before_uri_cloning() {
        let namespace = format!("urn:{}", "n".repeat(32 * 1024));
        let qnames = (0..256)
            .map(|index| format!("p:e{index}"))
            .collect::<Vec<_>>()
            .join(" ");
        let xml = format!(
            r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="{MCE_NS}" xmlns:p="{namespace}" mc:Ignorable="p" mc:ProcessContent="{qnames}"><x:sheetData/></x:worksheet>"#
        );
        let mut projector = WorksheetRowProjector::new(Cursor::new(xml.as_bytes()), &[], &[]);
        let error = projector
            .next_item()
            .expect_err("derived QName strings must obey the active-context cap");
        assert!(matches!(
            error,
            WorksheetProjectorError::Limit {
                kind: WorksheetProjectorLimitKind::XmlContextBytes,
                limit: STREAMED_XML_CONTEXT_BYTES,
                observed,
                ..
            } if observed > STREAMED_XML_CONTEXT_BYTES
        ));
    }

    #[test]
    fn inherited_namespace_and_mce_context_is_not_multiplied_by_depth() {
        let mut xml = String::from(
            r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#,
        );
        let mut ignorable = Vec::new();
        for index in 0..128 {
            write!(xml, r#" xmlns:n{index}="urn:example:namespace:{index:04}""#).unwrap();
            ignorable.push(format!("n{index}"));
        }
        write!(
            xml,
            r#" mc:Ignorable="{}"><x:sheetData/><x:extLst>"#,
            ignorable.join(" ")
        )
        .unwrap();
        for _ in 0..128 {
            xml.push_str("<x:ext>");
        }
        for _ in 0..128 {
            xml.push_str("</x:ext>");
        }
        xml.push_str("</x:extLst></x:worksheet>");

        let mut projector = WorksheetRowProjector::new(Cursor::new(xml.as_bytes()), &[], &[]);
        loop {
            if matches!(
                projector.next_item().expect("deep shared context projects"),
                WorksheetProjectorItem::Finished(_)
            ) {
                break;
            }
        }
        assert!(projector.max_context_bytes() > 0);
        assert!(projector.max_context_bytes() < 32 * 1024);
    }

    #[test]
    fn active_governor_receives_canonical_projector_limit_and_part() {
        let governor = ooxml_common::resource::ResourceGovernor::from_wasm(
            ooxml_common::resource::OoxmlFormat::Xlsx,
            None,
            None,
            None,
        );
        let _scope = governor.scope("parse-sheet");
        let error = report_projector_limit(
            None,
            worksheet_projector_limit(
                WorksheetProjectorLimitKind::XmlContextBytes,
                Some("xl/worksheets/sheet9.xml"),
                4,
                5,
            ),
        );
        let message = error.to_string();
        assert!(message.starts_with("OOXML_RESOURCE_LIMIT:"));
        assert!(message.contains("xl/worksheets/sheet9.xml"));
        assert!(message.contains("xml-context"));
    }

    #[test]
    fn package_backed_projector_limit_poisons_sibling_operation() {
        const PART: &str = "xl/worksheets/sheet1.xml";
        let oversized_text = "x".repeat(STREAMED_XML_EVENT_BYTES + 64);
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c t="inlineStr"><is><t>{oversized_text}</t></is></c></row></sheetData></worksheet>"#
        );
        let cursor = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(cursor);
        writer
            .start_file(PART, SimpleFileOptions::default())
            .expect("worksheet entry starts");
        writer
            .write_all(xml.as_bytes())
            .expect("worksheet is written");
        let bytes = writer.finish().expect("package finishes").into_inner();

        let package = PackageSessionHandle::open(
            bytes,
            OoxmlFormat::Xlsx,
            Some(2 * 1024 * 1024),
            Some(2 * 1024 * 1024),
            None,
        )
        .expect("package opens");
        let operation = package
            .begin_operation("parse-sheet")
            .expect("primary operation starts");
        let sibling = package
            .begin_operation("inspect")
            .expect("sibling operation starts");
        let entry = operation.open_entry(PART).expect("worksheet stream opens");
        let mut projector = WorksheetRowProjector::from_package_entry(
            entry,
            &[] as &[SharedString],
            &[] as &[String],
        )
        .expect("package projector starts");

        let error = projector
            .next_item()
            .expect_err("projector hard limit must poison the package")
            .to_string();
        assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"));
        assert!(error.contains("\"operation\":\"parse-sheet\""));
        assert!(error.contains("\"resource\":\"xml-event\""));
        assert!(error.contains("\"part\":\"xl/worksheets/sheet1.xml\""));
        assert_eq!(package.assert_healthy().unwrap_err(), error);
        assert_eq!(sibling.read_head(PART, 1).unwrap_err(), error);
    }

    #[test]
    fn exact_event_limit_at_eof_is_xml_but_limit_plus_one_is_structured() {
        const EVENT_LIMIT: usize = 64;
        let exact = format!("<{}", "x".repeat(EVENT_LIMIT - 1));
        let mut exact_projector = WorksheetRowProjector::with_limits(
            Cursor::new(exact.as_bytes()),
            &[] as &[SharedString],
            &[] as &[String],
            EVENT_LIMIT,
            1024,
            1024,
            None,
        );
        let exact_error = exact_projector
            .next_item()
            .expect_err("unterminated exact-limit token is malformed XML");
        assert!(matches!(exact_error, WorksheetProjectorError::Xml(_)));

        let over = format!("<{}", "x".repeat(EVENT_LIMIT));
        let mut over_projector = WorksheetRowProjector::with_limits(
            Cursor::new(over.as_bytes()),
            &[] as &[SharedString],
            &[] as &[String],
            EVENT_LIMIT,
            1024,
            1024,
            None,
        );
        assert_eq!(
            over_projector
                .next_item()
                .expect_err("one proven byte beyond the event limit"),
            WorksheetProjectorError::Limit {
                kind: WorksheetProjectorLimitKind::XmlEventBytes,
                part: None,
                limit: EVENT_LIMIT,
                observed: EVENT_LIMIT + 1,
            }
        );
    }

    fn comment_with_projected_bytes(bytes: usize) -> String {
        assert!(bytes >= "<!---->".len());
        format!("<!--{}-->", "x".repeat(bytes - "<!---->".len()))
    }

    fn row_projection_overhead() -> usize {
        PendingStreamedRow {
            xml: Vec::new(),
            namespaces: vec![(
                None,
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            )],
        }
        .projected_arena_bytes()
    }

    #[test]
    fn projected_row_exact_limit_succeeds_and_plus_one_fails_before_growth() {
        const COMMENT_BYTES: usize = 64;
        let row_limit = row_projection_overhead() + "<row></row>".len() + COMMENT_BYTES;
        let document = |comment_bytes| {
            format!(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row>{}</row></sheetData></worksheet>"#,
                comment_with_projected_bytes(comment_bytes)
            )
        };

        let exact = document(COMMENT_BYTES);
        let mut exact_projector = WorksheetRowProjector::with_limits(
            Cursor::new(exact.as_bytes()),
            &[] as &[SharedString],
            &[] as &[String],
            1024,
            row_limit,
            1024,
            None,
        );
        assert!(matches!(
            exact_projector.next_item().expect("exact row limit"),
            WorksheetProjectorItem::Row(_)
        ));
        assert!(matches!(
            exact_projector.next_item().expect("exact row tail"),
            WorksheetProjectorItem::Finished(_)
        ));

        let over = document(COMMENT_BYTES + 1);
        let mut over_projector = WorksheetRowProjector::with_limits(
            Cursor::new(over.as_bytes()),
            &[] as &[SharedString],
            &[] as &[String],
            1024,
            row_limit,
            1024,
            Some("xl/worksheets/sheet1.xml".to_string()),
        );
        assert_eq!(
            over_projector
                .next_item()
                .expect_err("row limit + 1 must fail"),
            WorksheetProjectorError::Limit {
                kind: WorksheetProjectorLimitKind::RowProjectionBytes,
                part: Some("xl/worksheets/sheet1.xml".to_string()),
                limit: row_limit,
                observed: row_limit + 1,
            }
        );
        assert!(over_projector.active_row.is_none());
        assert!(over_projector.row_batch.pending.is_empty());

        let mut batch = StreamedRowBatch::default();
        let direct_error = batch
            .push_bounded(
                PendingStreamedRow {
                    xml: vec![b'x'; row_limit + 1],
                    namespaces: Vec::new(),
                },
                row_limit,
                None,
                RowProjectionInputs {
                    shared_strings: &[],
                    theme_colors: &[],
                },
                &mut VecDeque::new(),
            )
            .expect_err("row batch independently enforces the row cap");
        assert!(matches!(
            direct_error,
            WorksheetProjectorError::Limit {
                kind: WorksheetProjectorLimitKind::RowProjectionBytes,
                ..
            }
        ));
        assert!(batch.pending.is_empty());
    }

    #[test]
    fn retained_shell_exact_limit_succeeds_and_plus_one_fails_before_growth() {
        const COMMENT_BYTES: usize = 64;
        let document = |comment_bytes| {
            format!(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{}<sheetData/></worksheet>"#,
                comment_with_projected_bytes(comment_bytes)
            )
        };
        let base = document(7);
        let shell_limit = base.len() - 7 + COMMENT_BYTES;

        let exact = document(COMMENT_BYTES);
        let mut exact_projector = WorksheetRowProjector::with_limits(
            Cursor::new(exact.as_bytes()),
            &[] as &[SharedString],
            &[] as &[String],
            1024,
            1024,
            shell_limit,
            None,
        );
        let exact_tail = match exact_projector.next_item().expect("exact shell limit") {
            WorksheetProjectorItem::Finished(tail) => tail,
            WorksheetProjectorItem::Row(_) => panic!("empty sheet unexpectedly produced a row"),
        };
        assert_eq!(exact_tail.shell_xml.len(), shell_limit);

        let over = document(COMMENT_BYTES + 1);
        let mut over_projector = WorksheetRowProjector::with_limits(
            Cursor::new(over.as_bytes()),
            &[] as &[SharedString],
            &[] as &[String],
            1024,
            1024,
            shell_limit,
            None,
        );
        assert_eq!(
            over_projector
                .next_item()
                .expect_err("shell limit + 1 must fail"),
            WorksheetProjectorError::Limit {
                kind: WorksheetProjectorLimitKind::WorksheetShellBytes,
                part: None,
                limit: shell_limit,
                observed: shell_limit + 1,
            }
        );
        assert!(over_projector.shell_xml.is_empty());
    }

    #[test]
    fn malformed_xml_after_rows_marks_the_projection_transaction_failed() {
        let mut xml = String::from(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
        );
        for row in 1..=STREAMED_ROW_BATCH_ROWS {
            write!(xml, r#"<row r="{row}"/>"#).unwrap();
        }
        xml.push_str("</sheetData><mergeCells>");

        let mut projector = WorksheetRowProjector::new(Cursor::new(xml.as_bytes()), &[], &[]);
        assert!(matches!(
            projector.next_item().expect("provisional row"),
            WorksheetProjectorItem::Row(_)
        ));

        let mut provisional_rows = 1usize;
        let error = loop {
            match projector.next_item() {
                Ok(WorksheetProjectorItem::Row(_)) => provisional_rows += 1,
                Ok(WorksheetProjectorItem::Finished(_)) => {
                    panic!("malformed worksheet must not commit")
                }
                Err(error) => break error,
            }
        };
        assert_eq!(provisional_rows, STREAMED_ROW_BATCH_ROWS);
        assert!(error.to_string().contains("unclosed elements"));
        assert_eq!(projector.state, WorksheetProjectorState::Failed);
        assert!(projector.ready_rows.is_empty());
        assert!(projector.finished_tail.is_none());
        assert!(projector
            .next_item()
            .expect_err("failed transaction stays failed")
            .to_string()
            .contains("already failed"));
    }

    #[test]
    fn compatibility_callback_receives_provisional_rows_before_malformed_tail() {
        let mut xml = String::from(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
        );
        for row in 1..=STREAMED_ROW_BATCH_ROWS {
            write!(xml, r#"<row r="{row}"/>"#).unwrap();
        }
        xml.push_str("</sheetData><mergeCells>");

        let mut visited = 0usize;
        let error = stream_worksheet_rows(&xml, &[], &[], |_| {
            visited += 1;
            Ok(())
        })
        .expect_err("malformed tail invalidates provisional callbacks");
        assert!(error.contains("unclosed elements"));
        assert_eq!(visited, STREAMED_ROW_BATCH_ROWS);

        // The actual worksheet-reference caller owns its index transaction and
        // exposes no cells when the same malformed tail prevents `Finished`.
        let cells = extract_reference_cells(
            &xml,
            &CellRange {
                top: 1,
                left: 1,
                bottom: STREAMED_ROW_BATCH_ROWS as u32,
                right: 1,
            },
            &[],
        );
        assert!(cells.is_empty());
    }

    #[test]
    fn callback_failure_short_circuits_source_and_later_callbacks() {
        const ROWS: usize = STREAMED_ROW_BATCH_ROWS + 128;
        let mut xml = String::from(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
        );
        for row in 1..=ROWS {
            write!(xml, r#"<row r="{row}"/>"#).unwrap();
        }
        xml.push_str("</sheetData></worksheet>");
        let consumed = Rc::new(Cell::new(0));
        let source = CountingRead {
            cursor: Cursor::new(xml.as_bytes()),
            consumed: Rc::clone(&consumed),
        };
        let projector = WorksheetRowProjector::new(
            BufReader::with_capacity(31, source),
            &[] as &[SharedString],
            &[] as &[String],
        );
        let mut callbacks = 0usize;
        let error = drain_worksheet_projector(projector, |_| {
            callbacks += 1;
            Err("stop after first callback".to_string())
        })
        .expect_err("callback failure stops the projector");
        assert_eq!(error, "stop after first callback");
        assert_eq!(callbacks, 1);
        assert!(consumed.get() < xml.len());
    }

    #[test]
    fn xml_depth_is_rejected_before_the_next_frame_is_retained() {
        let mut xml = String::from(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row>"#,
        );
        for _ in 3..=MAX_XML_DEPTH {
            xml.push_str("<a>");
        }
        xml.push_str("<tooDeep>");

        let mut projector = WorksheetRowProjector::new(
            Cursor::new(xml.as_bytes()),
            &[] as &[SharedString],
            &[] as &[String],
        );
        assert_eq!(
            projector
                .next_item()
                .expect_err("depth limit must reject before frame push"),
            WorksheetProjectorError::Limit {
                kind: WorksheetProjectorLimitKind::XmlDepth,
                part: None,
                limit: MAX_XML_DEPTH as usize,
                observed: MAX_XML_DEPTH as usize + 1,
            }
        );
        assert!(projector.frames.is_empty());
    }

    #[test]
    fn failed_persistent_projector_drops_reader_buffers_and_owned_dependencies() {
        let shared: Rc<[SharedString]> = Rc::from(vec![SharedString::default()].into_boxed_slice());
        let colors: Rc<[String]> = Rc::from(vec!["FFFFFF".to_string()].into_boxed_slice());
        let shared_weak = Rc::downgrade(&shared);
        let colors_weak = Rc::downgrade(&colors);
        let malformed = b"<worksheet><sheetData><row>";
        let mut projector =
            WorksheetRowProjector::new(Cursor::new(malformed.as_slice()), shared, colors);

        projector
            .next_item()
            .expect_err("malformed persistent projector must fail");
        assert!(projector.reader.is_none());
        assert!(projector.frames.is_empty());
        assert!(projector.active_row.is_none());
        assert!(projector.shell_xml.is_empty());
        assert!(projector.row_batch.pending.is_empty());
        assert_eq!(projector.row_batch.geometry, AuthoredRowGeometry::default());
        assert!(projector.ready_rows.is_empty());
        assert!(projector.finished_tail.is_none());
        assert!(shared_weak.upgrade().is_none());
        assert!(colors_weak.upgrade().is_none());
    }

    #[test]
    fn stream_matches_dom_for_rich_inline_cells_and_implicit_references() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
              <row r="3" ht="21.5" outlineLevel="2" collapsed="1" ph="1">
                <c t="inlineStr"><is>
                  <r><rPr><b/><color rgb="FF112233"/></rPr><t xml:space="preserve"> A &amp; </t></r>
                  <r><t><![CDATA[B < C]]></t></r>
                  <rPh sb="0" eb="1"><t>えー</t></rPh>
                  <phoneticPr fontId="2" type="Hiragana" alignment="center"/>
                </is></c>
                <c r="C3" t="str" ph="0"><f>CONCAT(&quot;x&quot;,&quot;y&quot;)</f><v>x&amp;y</v></c>
                <c><v>42.25</v></c>
              </row>
              <row hidden="true"><c t="b"><v>true</v></c><c t="e"><v>#N/A</v></c></row>
              <row/>
            </sheetData>
            <mergeCells count="1"><mergeCell ref="A3:D3"/></mergeCells>
          </worksheet>"#;
        assert_streamed_rows_match_dom(xml);
    }

    #[test]
    fn strict_prefixed_stream_selects_first_understood_mce_choice() {
        // Part 3 §9.3: skip the unknown first Choice, select the first Choice
        // whose complete Requires set is understood, then ignore later choices
        // and Fallback. The raw DOM intentionally is not the oracle here because
        // it contains all branches before MCE preprocessing.
        let xml = r#"<x:worksheet
            xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
            xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
            mc:Ignorable="x15 x14">
          <x:sheetData>
            <mc:AlternateContent>
              <mc:Choice Requires="x15">
                <x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>unknown</x:t></x:is></x:c></x:row>
              </mc:Choice>
              <mc:Choice Requires="x14">
                <x:row r="4"><x:c r="A4" t="inlineStr"><x:is><x:t>selected</x:t></x:is></x:c></x:row>
              </mc:Choice>
              <mc:Choice Requires="x14">
                <x:row r="5"><x:c r="A5" t="inlineStr"><x:is><x:t>later</x:t></x:is></x:c></x:row>
              </mc:Choice>
              <mc:Fallback>
                <x:row><x:c t="str"><x:v>fallback</x:v></x:c></x:row>
              </mc:Fallback>
            </mc:AlternateContent>
          </x:sheetData>
          <x:hyperlinks><x:hyperlink ref="A1" location="Target!A1"/></x:hyperlinks>
        </x:worksheet>"#;
        let streamed = stream_sheet_data(xml, &[], &[]).expect("MCE worksheet streams");
        assert_eq!(streamed.rows.len(), 1);
        assert_eq!(streamed.rows[0].index, 4);
        assert!(matches!(
            &streamed.rows[0].cells[0].value,
            CellValue::Text { text, .. } if text == "selected"
        ));
    }

    #[test]
    fn mce_unknown_choice_selects_fallback_only() {
        let xml = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future">
          <x:sheetData>
            <mc:AlternateContent>
              <mc:Choice Requires="future">
                <x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>
              </mc:Choice>
              <mc:Fallback>
                <x:row r="2"><x:c r="A2"><x:v>2</x:v></x:c></x:row>
              </mc:Fallback>
            </mc:AlternateContent>
          </x:sheetData>
        </x:worksheet>"#;
        let streamed = stream_sheet_data(xml, &[], &[]).expect("MCE fallback streams");
        assert_eq!(streamed.rows.len(), 1);
        assert_eq!(streamed.rows[0].index, 2);
    }

    #[test]
    fn mce_alternate_content_can_supply_the_processed_document_root() {
        // Part 3 §§7.5, 9.3–9.4: AlternateContent is replaced by the selected
        // branch content before the host schema sees the document element.
        // All namespace declarations intentionally live on the removed wrapper.
        let xml = r#"<mc:AlternateContent
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:future="urn:example:future">
          <mc:Choice Requires="future">
            <x:notAWorksheet/>
          </mc:Choice>
          <mc:Fallback>
            <x:worksheet>
              <x:sheetData>
                <x:row r="7"><x:c r="A7"><x:v>42</x:v></x:c></x:row>
              </x:sheetData>
            </x:worksheet>
          </mc:Fallback>
        </mc:AlternateContent>"#;
        let streamed = stream_sheet_data(xml, &[], &[]).expect("selected root is processed");
        assert_eq!(streamed.rows.len(), 1);
        assert_eq!(streamed.rows[0].index, 7);
        let shell = parse_guarded(&streamed.shell_xml).expect("processed shell is namespace-valid");
        assert_eq!(shell.root_element().tag_name().name(), "worksheet");
        assert!(is_x_ns(shell.root_element().tag_name().namespace()));
    }

    #[test]
    fn moved_fallback_root_has_a_frozen_namespace_repaired_projection() {
        let xml = r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:future="urn:example:future"><mc:Choice Requires="future"><x:notAWorksheet/></mc:Choice><mc:Fallback><x:worksheet><x:sheetData><x:row r="7"><x:c r="A7"><x:v>42</x:v></x:c></x:row></x:sheetData></x:worksheet></mc:Fallback></mc:AlternateContent>"#;
        let actual = project_reader(BufReader::with_capacity(1, Cursor::new(xml.as_bytes())))
            .expect("one-byte fallback-root projection");
        assert_eq!(actual.rows.len(), 1);
        assert_eq!(actual.rows[0].index, 7);
        assert_eq!(actual.rows[0].cells.len(), 1);
        assert_eq!(actual.rows[0].cells[0].col, 1);
        assert!(matches!(
            actual.rows[0].cells[0].value,
            CellValue::Number { number } if number == 42.0
        ));
        assert_eq!(
            actual.shell_xml,
            r#"<x:worksheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:future="urn:example:future"><x:sheetData></x:sheetData></x:worksheet>"#
        );
    }

    #[test]
    fn moved_root_repairs_escaped_namespace_uri_exactly_once() {
        let xml = r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:f="urn:a&amp;b"><mc:Choice Requires="f"><x:notAWorksheet/></mc:Choice><mc:Fallback><x:worksheet><x:sheetData/></x:worksheet></mc:Fallback></mc:AlternateContent>"#;
        let streamed = stream_sheet_data(xml, &[], &[]).expect("escaped binding is repaired");
        assert_eq!(
            streamed.shell_xml,
            r#"<x:worksheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:f="urn:a&amp;b"><x:sheetData/></x:worksheet>"#
        );
    }

    #[test]
    fn moved_namespace_injection_limit_counts_escaped_serialized_uri() {
        let namespace = "urn:a&b".to_string();
        let root = StreamedNamespaceContext::root();
        let declaration =
            quick_xml::events::BytesStart::from_content(r#"r xmlns:f="urn:a&amp;b""#, 1);
        let context = StreamedNamespaceContext::derive(
            &declaration,
            &root,
            STREAMED_XML_CONTEXT_BYTES,
            "worksheet",
        )
        .expect("escaped namespace declaration is valid");
        let injected_bytes = 10 + "f".len() + quick_xml::escape::escape(&namespace).len();

        let exact_base = quick_xml::events::BytesStart::new("x:worksheet");
        let exact_padding = STREAMED_XML_CONTEXT_BYTES
            .checked_sub(injected_bytes)
            .and_then(|remaining| remaining.checked_sub(exact_base.len() + 7))
            .expect("context ceiling leaves padding room");
        let padding = "p".repeat(exact_padding);
        let mut exact = exact_base;
        exact.push_attribute(("pad", padding.as_str()));
        assert_eq!(exact.len() + injected_bytes, STREAMED_XML_CONTEXT_BYTES);
        inject_moved_element_namespaces(&mut exact, &context, None, None)
            .expect("exact escaped injection limit succeeds");
        assert_eq!(exact.len(), STREAMED_XML_CONTEXT_BYTES);
        assert!(String::from_utf8_lossy(exact.as_ref()).contains(r#"xmlns:f="urn:a&amp;b""#));

        let over_padding = "p".repeat(exact_padding + 1);
        let mut over = quick_xml::events::BytesStart::new("x:worksheet");
        over.push_attribute(("pad", over_padding.as_str()));
        let original_len = over.len();
        assert_eq!(
            original_len + injected_bytes,
            STREAMED_XML_CONTEXT_BYTES + 1
        );
        assert_eq!(
            inject_moved_element_namespaces(&mut over, &context, None, None)
                .expect_err("escaped injection limit + 1 fails before mutation"),
            WorksheetProjectorError::Limit {
                kind: WorksheetProjectorLimitKind::XmlContextBytes,
                part: None,
                limit: STREAMED_XML_CONTEXT_BYTES,
                observed: STREAMED_XML_CONTEXT_BYTES + 1,
            }
        );
        assert_eq!(over.len(), original_len);
    }

    #[test]
    fn mce_alternate_content_can_supply_direct_sheet_data() {
        // The physical parent is Choice, but §9.4 removes both
        // AlternateContent and the selected branch wrapper. CT_Worksheet
        // therefore sees sheetData as a direct child.
        let xml = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
          <mc:AlternateContent>
            <mc:Choice Requires="x14">
              <x:sheetData>
                <x:row r="3"><x:c r="A3" t="str"><x:v>selected</x:v></x:c></x:row>
              </x:sheetData>
            </mc:Choice>
            <mc:Fallback><x:sheetData/></mc:Fallback>
          </mc:AlternateContent>
        </x:worksheet>"#;
        let streamed = stream_sheet_data(xml, &[], &[]).expect("selected sheetData is direct");
        assert_eq!(streamed.rows.len(), 1);
        assert_eq!(streamed.rows[0].index, 3);
        let shell = parse_guarded(&streamed.shell_xml).expect("processed shell parses");
        let root = shell.root_element();
        let sheet_data = root
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "sheetData")
            .expect("processed direct sheetData");
        assert_eq!(sheet_data.parent_element(), Some(root));
    }

    #[test]
    fn mce_process_content_unwraps_ignorable_row_container() {
        // Part 3 §9.2 marks future:rows as unwrapped (rather than ignored)
        // because its expanded name appears in the inherited ProcessContent
        // set. CT_SheetData consequently validates and consumes its row child.
        let xml = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future"
            mc:Ignorable="future"
            mc:ProcessContent="future:*">
          <x:sheetData>
            <future:empty/>
            <future:rows>
              <x:row r="9"><x:c r="B9"><x:v>9</x:v></x:c></x:row>
            </future:rows>
          </x:sheetData>
        </x:worksheet>"#;
        let streamed = stream_sheet_data(xml, &[], &[]).expect("ProcessContent exposes direct row");
        assert_eq!(streamed.rows.len(), 1);
        assert_eq!(streamed.rows[0].index, 9);
        assert_eq!(streamed.rows[0].cells[0].col, 2);
    }

    #[test]
    fn escaped_namespace_uri_matches_mce_ignorable_and_process_content() {
        // Namespace names compare after XML entity expansion. The resolver and
        // persistent context must therefore both represent this as `urn:a&b`.
        let xml = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:f="urn:a&amp;b"
            mc:Ignorable="f"
            mc:ProcessContent="f:rows">
          <x:sheetData>
            <f:rows><x:row r="11"><x:c r="A11"><x:v>11</x:v></x:c></x:row></f:rows>
          </x:sheetData>
        </x:worksheet>"#;
        let streamed = stream_sheet_data(xml, &[], &[])
            .expect("escaped namespace URI participates in MCE matching");
        assert_eq!(streamed.rows.len(), 1);
        assert_eq!(streamed.rows[0].index, 11);
        assert!(matches!(
            streamed.rows[0].cells[0].value,
            CellValue::Number { number } if number == 11.0
        ));
    }

    #[test]
    fn mce_inside_row_selects_direct_cell_before_row_parsing() {
        let xml = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
          <x:sheetData>
            <x:row r="4">
              <mc:AlternateContent>
                <mc:Choice Requires="x14">
                  <x:c r="C4" t="str"><x:v>selected-cell</x:v></x:c>
                </mc:Choice>
                <mc:Fallback>
                  <x:c r="D4" t="str"><x:v>fallback-cell</x:v></x:c>
                </mc:Fallback>
              </mc:AlternateContent>
            </x:row>
          </x:sheetData>
        </x:worksheet>"#;
        let streamed = stream_sheet_data(xml, &[], &[]).expect("row MCE is preprocessed");
        assert_eq!(streamed.rows.len(), 1);
        assert_eq!(streamed.rows[0].cells.len(), 1);
        assert_eq!(streamed.rows[0].cells[0].col, 3);
        assert!(matches!(
            &streamed.rows[0].cells[0].value,
            CellValue::Text { text, .. } if text == "selected-cell"
        ));
    }

    #[test]
    fn mce_process_content_inside_row_exposes_direct_cell() {
        let xml = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future"
            mc:Ignorable="future"
            mc:ProcessContent="future:cells">
          <x:sheetData>
            <x:row r="6">
              <future:cells>
                <x:c r="B6"><x:v>6</x:v></x:c>
              </future:cells>
            </x:row>
          </x:sheetData>
        </x:worksheet>"#;
        let streamed =
            stream_sheet_data(xml, &[], &[]).expect("row ProcessContent is preprocessed");
        assert_eq!(streamed.rows.len(), 1);
        assert_eq!(streamed.rows[0].cells.len(), 1);
        assert_eq!(streamed.rows[0].cells[0].col, 2);
        assert!(matches!(
            streamed.rows[0].cells[0].value,
            CellValue::Number { number } if number == 6.0
        ));
    }

    #[test]
    fn mce_is_processed_inside_ordinary_worksheet_subtrees() {
        let xml = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
          <x:sheetData/>
          <x:mergeCells count="1">
            <mc:AlternateContent>
              <mc:Choice Requires="x14"><x:mergeCell ref="A1:B1"/></mc:Choice>
              <mc:Fallback><x:mergeCell ref="C1:D1"/></mc:Fallback>
            </mc:AlternateContent>
          </x:mergeCells>
        </x:worksheet>"#;
        let (worksheet, _) =
            parse_worksheet(xml, &[], &[], "MCE").expect("ordinary subtree is preprocessed");
        assert_eq!(worksheet.merge_cells.len(), 1);
        assert_eq!(worksheet.merge_cells[0].left, 1);
        assert_eq!(worksheet.merge_cells[0].right, 2);
    }

    #[test]
    fn processed_ordinary_elements_strip_mce_and_ignored_attributes() {
        // ECMA-376 Part 3 §9.4 step 5(a). The adjacent extLst test covers the
        // separate opaque application-extension rule.
        let xml = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future"
            mc:Ignorable="future"
            mc:ProcessContent="future:wrapper"
            future:discard="root">
          <x:sheetData/>
          <x:mergeCells future:discard="collection">
            <x:mergeCell ref="A1:B1" future:discard="item"/>
          </x:mergeCells>
        </x:worksheet>"#;

        let streamed = stream_sheet_data(xml, &[], &[]).expect("worksheet is processed");
        assert!(!streamed.shell_xml.contains("mc:Ignorable"));
        assert!(!streamed.shell_xml.contains("mc:ProcessContent"));
        assert!(!streamed.shell_xml.contains("future:discard"));
        assert!(streamed.shell_xml.contains("xmlns:future"));
        assert!(streamed.shell_xml.contains(r#"ref="A1:B1""#));
    }

    #[test]
    fn spreadsheetml_ext_lst_is_the_configured_opaque_boundary() {
        // Part 1 §10 designates extLst as the application-defined extension
        // element. Its contents pass through unchanged even when they contain
        // MCE markup that would be nonconformant in an ordinary worksheet
        // subtree.
        const PAYLOAD: &str = r#"<x:sheetData><x:row r="99"><mc:AlternateContent><mc:Fallback><x:c r="A99"/></mc:Fallback></mc:AlternateContent></x:row></x:sheetData><x:sheetData/><x:row r="100"/>"#;
        let xml = format!(
            r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future">
          <x:sheetData/>
          <x:extLst mc:MustUnderstand="future">
            <x:ext uri="urn:example">
              {PAYLOAD}
            </x:ext>
          </x:extLst>
        </x:worksheet>"#
        );
        let streamed =
            stream_sheet_data(&xml, &[], &[]).expect("extLst contents bypass MCE processing");
        assert!(streamed.rows.is_empty());
        assert!(streamed.shell_xml.contains(PAYLOAD));
        assert!(streamed.shell_xml.contains("<mc:AlternateContent>"));
        assert!(streamed.shell_xml.contains("<mc:Fallback>"));
        assert!(streamed.shell_xml.contains(r#"mc:MustUnderstand="future""#));

        let invalid_direct_child = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future">
          <x:sheetData/>
          <mc:AlternateContent>
            <x:extLst mc:MustUnderstand="future"/>
          </mc:AlternateContent>
        </x:worksheet>"#;
        let error = stream_sheet_data(invalid_direct_child, &[], &[])
            .expect_err("extLst does not bypass AlternateContent child grammar");
        assert!(error.contains("only Choice/Fallback children"));
    }

    #[test]
    fn generated_row_projection_bytes_obey_the_batch_ceiling() {
        const NAMESPACES: usize = 192;
        const ROWS: usize = 300;
        let mut xml = String::from(
            r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:example:future" mc:Ignorable="future" mc:ProcessContent="future:cells""#,
        );
        for index in 0..NAMESPACES {
            write!(xml, r#" xmlns:n{index}="urn:example:namespace:{index:04}""#)
                .expect("namespace declaration");
        }
        xml.push_str("><x:sheetData>");
        for row in 1..=ROWS {
            write!(
                xml,
                r#"<x:row r="{row}"><future:cells><x:c r="A{row}"><x:v>{row}</x:v></x:c></future:cells></x:row>"#
            )
            .expect("projected row");
        }
        xml.push_str("</x:sheetData></x:worksheet>");

        let mut row_count = 0usize;
        let streamed = stream_worksheet_rows(&xml, &[], &[], |row| {
            row_count += 1;
            assert_eq!(row.cells.len(), 1);
            Ok(())
        })
        .expect("namespace-heavy projected rows stream");
        assert_eq!(row_count, ROWS);
        assert!(
            streamed.max_row_arena_bytes
                <= STREAMED_ROWS_ARENA_OVERHEAD
                    + STREAMED_ROW_BATCH_BYTES.max(streamed.max_single_row_projection_bytes),
            "generated row arena used {} bytes; target {}, largest row {}",
            streamed.max_row_arena_bytes,
            STREAMED_ROW_BATCH_BYTES,
            streamed.max_single_row_projection_bytes
        );
    }

    #[test]
    fn normal_worksheet_shell_remains_byte_exact_outside_sheet_data() {
        let xml = r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr/><sheetData>
  <row r="1"><c r="A1"><v>1</v></c></row>
</sheetData><mergeCells count="0"/></worksheet>"#;
        let streamed = stream_sheet_data(xml, &[], &[]).expect("ordinary worksheet streams");
        assert_eq!(
            streamed.shell_xml,
            r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr/><sheetData></sheetData><mergeCells count="0"/></worksheet>"#
        );
    }

    #[test]
    fn malformed_mce_processed_root_and_sheet_data_are_rejected() {
        let no_root = r#"<mc:AlternateContent
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:future="urn:example:future">
          <mc:Choice Requires="future"><x:worksheet><x:sheetData/></x:worksheet></mc:Choice>
        </mc:AlternateContent>"#;
        let error = stream_sheet_data(no_root, &[], &[])
            .expect_err("no selected branch leaves no worksheet root");
        assert!(error.contains("root must be exactly one"));

        let duplicate_sheet_data = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future">
          <x:sheetData/>
          <mc:AlternateContent>
            <mc:Choice Requires="future"><x:sheetData/></mc:Choice>
            <mc:Fallback><x:sheetData/></mc:Fallback>
          </mc:AlternateContent>
        </x:worksheet>"#;
        let error = stream_sheet_data(duplicate_sheet_data, &[], &[])
            .expect_err("processed CT_Worksheet has duplicate sheetData");
        assert!(error.contains("exactly one non-nested sheetData"));

        let non_row = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future"
            mc:Ignorable="future"
            mc:ProcessContent="future:rows">
          <x:sheetData><future:rows><x:c r="A1"/></future:rows></x:sheetData>
        </x:worksheet>"#;
        let error = stream_sheet_data(non_row, &[], &[])
            .expect_err("processed CT_SheetData has a non-row child");
        assert!(error.contains("only direct SpreadsheetML row children"));
    }

    #[test]
    fn malformed_mce_declarations_and_alternate_content_are_rejected() {
        let unbound_process_content = r#"<worksheet
            xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:ProcessContent="missing:rows"><sheetData/></worksheet>"#;
        let error = stream_sheet_data(unbound_process_content, &[], &[])
            .expect_err("ProcessContent QName must resolve");
        assert!(error.contains("unbound namespace prefix"));

        let non_ignorable_process_content = r#"<worksheet
            xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future"
            mc:ProcessContent="future:rows"><sheetData/></worksheet>"#;
        let error = stream_sheet_data(non_ignorable_process_content, &[], &[])
            .expect_err("ProcessContent namespace must be ignorable");
        assert!(error.contains("must be declared Ignorable"));

        let xml_space_on_unwrapped = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future"
            mc:Ignorable="future"
            mc:ProcessContent="future:rows">
          <x:sheetData><future:rows xml:space="preserve"><x:row/></future:rows></x:sheetData>
        </x:worksheet>"#;
        let error = stream_sheet_data(xml_space_on_unwrapped, &[], &[])
            .expect_err("unwrapped element cannot alter XML context");
        assert!(error.contains("cannot carry xml:base, xml:lang, or xml:space"));

        let unknown_must_understand = r#"<worksheet
            xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:future="urn:example:future"
            mc:MustUnderstand="future"><sheetData/></worksheet>"#;
        let error = stream_sheet_data(unknown_must_understand, &[], &[])
            .expect_err("unknown MustUnderstand signals mismatch");
        assert!(error.contains("MustUnderstand namespace is not understood"));

        let invalid_alternate_child = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <mc:AlternateContent><x:sheetData/></mc:AlternateContent>
        </x:worksheet>"#;
        let error = stream_sheet_data(invalid_alternate_child, &[], &[])
            .expect_err("AlternateContent child grammar is enforced");
        assert!(error.contains("only Choice/Fallback children"));

        let fallback_without_choice = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <mc:AlternateContent><mc:Fallback><x:sheetData/></mc:Fallback></mc:AlternateContent>
        </x:worksheet>"#;
        let error = stream_sheet_data(fallback_without_choice, &[], &[])
            .expect_err("Part 3 §7.5 requires one or more Choice children");
        assert!(error.contains("at least one Choice before Fallback"));

        for invalid_choice in [
            r#"<mc:Choice><x:sheetData/></mc:Choice>"#,
            r#"<mc:Choice Requires="   "><x:sheetData/></mc:Choice>"#,
        ] {
            let xml = format!(
                r#"<x:worksheet
                    xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                  <mc:AlternateContent>{invalid_choice}</mc:AlternateContent>
                </x:worksheet>"#
            );
            let error = stream_sheet_data(&xml, &[], &[])
                .expect_err("Choice requires a non-empty Requires attribute");
            assert!(error.contains("Choice must have a non-empty Requires attribute"));
        }
    }

    #[test]
    fn unbound_choice_requires_is_nonconformant() {
        let xml = r#"<x:worksheet
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <mc:AlternateContent>
            <mc:Choice Requires="missing"><x:sheetData/></mc:Choice>
            <mc:Fallback><x:sheetData><x:row r="7"/></x:sheetData></mc:Fallback>
          </mc:AlternateContent>
        </x:worksheet>"#;
        let error = stream_sheet_data(xml, &[], &[])
            .expect_err("Part 3 section 7.6 requires every Requires prefix to be bound");
        assert!(error.contains("unbound namespace prefix"), "{error}");
    }

    #[test]
    fn mce_control_prefixes_must_not_resolve_to_the_mce_namespace() {
        let cases = [
            r#"mc:Ignorable="self"><x:sheetData/>"#,
            r#"mc:MustUnderstand="self"><x:sheetData/>"#,
            r#"><mc:AlternateContent><mc:Choice Requires="self"><x:sheetData/></mc:Choice></mc:AlternateContent>"#,
        ];
        for suffix in cases {
            let xml = format!(
                r#"<x:worksheet
                    xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:mc="{MCE_NS}"
                    xmlns:self="{MCE_NS}" {suffix}</x:worksheet>"#
            );
            let error = stream_sheet_data(&xml, &[], &[])
                .expect_err("Part 3 forbids control prefixes bound to the MCE namespace");
            assert!(
                error.contains("must not name the Markup Compatibility namespace"),
                "{error}"
            );
        }
    }

    #[test]
    fn alternate_content_elements_enforce_part3_attribute_grammar() {
        let cases = [
            r#"<mc:AlternateContent extra="no"><mc:Choice Requires="x"><x:sheetData/></mc:Choice></mc:AlternateContent>"#,
            r#"<mc:AlternateContent><mc:Choice Requires="x" extra="no"><x:sheetData/></mc:Choice></mc:AlternateContent>"#,
            r#"<mc:AlternateContent><mc:Choice Requires="future"><x:sheetData/></mc:Choice><mc:Fallback extra="no"><x:sheetData/></mc:Fallback></mc:AlternateContent>"#,
            r#"<mc:AlternateContent><mc:Choice Requires="x" future:allowed="yes" x:forbidden="no"><x:sheetData/></mc:Choice></mc:AlternateContent>"#,
        ];
        for alternate in cases {
            let xml = format!(
                r#"<x:worksheet
                    xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                    xmlns:future="urn:example:future"
                    mc:Ignorable="future">{alternate}</x:worksheet>"#
            );
            let error = stream_sheet_data(&xml, &[], &[])
                .expect_err("Part 3 sections 7.5 through 7.7 constrain MCE attributes");
            assert!(error.contains("attribute"), "{error}");
        }
    }

    #[test]
    fn empty_sheet_data_keeps_a_valid_empty_shell() {
        let xml = r#"<x:worksheet xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main"><x:sheetData/><x:mergeCells><x:mergeCell ref="A1:B1"/></x:mergeCells></x:worksheet>"#;
        let streamed = stream_sheet_data(xml, &[], &[]).expect("empty sheetData streams");
        assert!(streamed.rows.is_empty());
        assert_eq!(streamed.shell_xml, xml);
        let (worksheet, _) =
            parse_worksheet(xml, &[], &[], "Empty").expect("empty worksheet parses");
        assert!(worksheet.rows.is_empty());
        assert_eq!(worksheet.merge_cells.len(), 1);
    }

    #[test]
    fn malformed_unclosed_row_is_rejected_before_shell_parse() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></sheetData></worksheet>"#;
        let error = stream_sheet_data(xml, &[], &[]).expect_err("unclosed row must fail");
        assert!(
            error.contains("worksheet XML stream"),
            "unexpected error: {error}"
        );
    }

    #[test]
    fn missing_or_duplicate_sheet_data_is_rejected_by_ct_worksheet_contract() {
        let missing = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><mergeCells/></worksheet>"#;
        let missing_error =
            stream_sheet_data(missing, &[], &[]).expect_err("sheetData is required");
        assert!(missing_error.contains("exactly one sheetData"));

        let duplicate = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><sheetData/></worksheet>"#;
        let duplicate_error =
            stream_sheet_data(duplicate, &[], &[]).expect_err("sheetData is singular");
        assert!(duplicate_error.contains("exactly one non-nested sheetData"));

        let nested = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><sheetData/></sheetData></worksheet>"#;
        let nested_error = stream_sheet_data(nested, &[], &[]).expect_err("sheetData cannot nest");
        assert!(nested_error.contains("direct child of worksheet"));
    }

    #[test]
    fn sheet_data_rejects_rows_below_non_mce_wrappers() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><wrapper><row r="1"/></wrapper></sheetData></worksheet>"#;
        let error =
            stream_sheet_data(xml, &[], &[]).expect_err("CT_SheetData has direct rows only");
        assert!(error.contains("direct SpreadsheetML row children"));
    }

    #[test]
    fn high_density_sheet_keeps_dom_shell_small_and_preserves_all_cells() {
        const ROWS: u32 = 4_000;
        const COLS: u32 = 39;
        let mut xml = String::with_capacity(ROWS as usize * COLS as usize * 36);
        xml.push_str(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
        );
        for row in 1..=ROWS {
            write!(xml, r#"<row r="{row}">"#).expect("write row");
            for col in 1..=COLS {
                write!(
                    xml,
                    r#"<c r="{}{}" t="inlineStr"><is><t>v{}-{}</t></is></c>"#,
                    column_name(col),
                    row,
                    row,
                    col
                )
                .expect("write cell");
            }
            xml.push_str("</row>");
        }
        xml.push_str(
            r#"</sheetData><mergeCells><mergeCell ref="A1:B1"/></mergeCells></worksheet>"#,
        );

        let streamed = stream_sheet_data(&xml, &[], &[]).expect("large sheet streams");
        assert_eq!(streamed.rows.len(), ROWS as usize);
        assert_eq!(
            streamed
                .rows
                .iter()
                .map(|row| row.cells.len())
                .sum::<usize>(),
            (ROWS * COLS) as usize
        );
        assert!(
            streamed.shell_xml.len() < 256,
            "sheetData cell XML must not survive in the roxmltree shell"
        );
        let (worksheet, _) =
            parse_worksheet(&xml, &[], &[], "Dense").expect("public worksheet parse succeeds");
        assert_eq!(worksheet.rows.len(), ROWS as usize);
        assert_eq!(worksheet.merge_cells.len(), 1);
    }

    fn column_name(mut col: u32) -> String {
        let mut chars = Vec::new();
        while col > 0 {
            col -= 1;
            chars.push((b'A' + (col % 26) as u8) as char);
            col /= 26;
        }
        chars.into_iter().rev().collect()
    }
}
