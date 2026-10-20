//! Shared recursion-depth guard for the wild OOXML parsers.
//!
//! Several OOXML grammars nest a container element inside its own descendants,
//! and the natural parser for each is directly recursive:
//!
//!   * DrawingML group shapes — `<p:grpSp>` (pptx `parse_sp_tree_node`),
//!     `<xdr:grpSp>` (xlsx `collect_shapes`);
//!   * WordprocessingML tables — a `<w:tbl>` may live inside a `<w:tc>`
//!     (docx `parse_table` ↔ `parse_table_cell`);
//!   * OMML math — every math container (`<m:f>`, `<m:sSup>`, `<m:d>` …) holds
//!     child math (`parse_omath_nodes`).
//!
//! A hand-crafted document can nest these tens of thousands deep. Because the
//! parser runs in WASM (a fixed, comparatively small linear-memory stack — the
//! default `wasm-pack` build reserves 1 MiB), unbounded recursion overflows the
//! stack and aborts the whole parse with an unrecoverable `RuntimeError:
//! unreachable` **trap** — the browser tab loses the document, not just the
//! pathological subtree.
//!
//! [`DepthGuard`] bounds that recursion. Crossing it latches the existing typed
//! hard-resource failure in the active package operation. Walkers stop recursing;
//! their provisional output must never be published after a crossing. Package
//! operations and cursor pulls check health before returning output. Malformed
//! XML remains a separate, recoverable content error.
//!
//! ## Why 64
//!
//! [`MAX_PARSE_DEPTH`] is 64. That is comfortably above anything a real authoring
//! tool produces yet far below the stack budget:
//!
//!   * Word's own UI caps interactive table nesting well inside this range, and
//!     even programmatically-authored decks/sheets rarely nest groups past a
//!     handful of levels; a genuine document never approaches 64.
//!   * Each recursion frame here is heavy (a `roxmltree::Node` plus several
//!     `&HashMap`/`&[..]` references and, for tables, `String`/`Vec` locals), so
//!     64 frames is a tiny fraction of the 1 MiB WASM stack — there is a wide
//!     safety margin between the guard firing and an actual overflow.
//!
//! The limit is intentionally a single shared constant so the three parsers +
//! the shared OMML grammar behave identically; a document that is "too deep" in
//! one format is "too deep" in all of them.

use quick_xml::events::Event;

/// Maximum nesting depth accepted by any guarded OOXML recursion (group shapes,
/// nested tables, OMML math). See the module docs for the rationale.
pub const MAX_PARSE_DEPTH: u32 = 64;

/// A cheap, copyable recursion-depth counter threaded through directly-recursive
/// OOXML parsing functions.
///
/// Callers create one at the top of a recursion with [`DepthGuard::root`], then
/// obtain a child guard for each descent with [`DepthGuard::descend`]. When
/// `descend` returns `None` the limit has been reached and the caller must stop
/// recursing. The package operation rejects the provisional result.
///
/// The type is `Copy` and holds a single `u32`, so threading it as a by-value
/// parameter adds no allocation and no borrow-checker friction.
///
/// ```
/// use ooxml_common::depth::{DepthGuard, MAX_PARSE_DEPTH};
///
/// fn walk(depth: DepthGuard, fanout: u32) -> u32 {
///     // Count this node, then descend into `fanout` children if allowed.
///     let mut n = 1;
///     if let Some(child) = depth.descend() {
///         for _ in 0..fanout {
///             n += walk(child, fanout);
///         }
///     }
///     n
/// }
///
/// // A linear chain (fanout 1) visits exactly MAX_PARSE_DEPTH + 1 nodes before
/// // the guard stops the descent — it never traps.
/// assert_eq!(walk(DepthGuard::root(), 1), MAX_PARSE_DEPTH + 1);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DepthGuard {
    depth: u32,
    limit: u32,
}

impl DepthGuard {
    /// A guard at depth 0 with the shared [`MAX_PARSE_DEPTH`] limit. Use this at
    /// every top-level (non-recursive) entry into a guarded recursion.
    #[inline]
    pub const fn root() -> Self {
        Self {
            depth: 0,
            limit: MAX_PARSE_DEPTH,
        }
    }

    /// A guard at depth 0 with an explicit `limit`. Intended for tests; the
    /// wild parsers should use [`DepthGuard::root`] so every format shares the
    /// same ceiling.
    #[inline]
    pub const fn with_limit(limit: u32) -> Self {
        Self { depth: 0, limit }
    }

    /// The current depth (0 at the root).
    #[inline]
    pub const fn depth(&self) -> u32 {
        self.depth
    }

    /// `true` when descending one more level would exceed the limit — i.e. the
    /// caller has hit the floor and must not recurse again.
    #[inline]
    pub const fn is_exhausted(&self) -> bool {
        self.depth >= self.limit
    }

    /// A child guard one level deeper, or `None` when the limit is reached.
    ///
    /// `Some(child)` ⇒ recurse with `child`; `None` ⇒ stop and return the
    /// provisional result; the package operation must fail. The parent guard is
    /// unchanged (`Copy`), so sibling descents each get their own fresh child.
    #[inline]
    pub fn descend(self) -> Option<Self> {
        if self.depth >= self.limit {
            report_depth_limit(self.limit);
            None
        } else {
            Some(Self {
                depth: self.depth + 1,
                limit: self.limit,
            })
        }
    }
}

/// Library resource policy, not a schema restriction. Shared walkers may be
/// infallible, so latch before returning a bounded provisional projection. This
/// uses the package-owned governor selected by the synchronous resource scope;
/// no caller compatibility/unsupported-content policy can clear the failure.
fn report_depth_limit(limit: u32) {
    let _ = crate::resource::observe_hard_limit(
        crate::resource::HardResourceLimitKind::XmlNestingDepth,
        None,
        u64::from(limit),
        u64::from(limit) + 1,
    );
}

impl Default for DepthGuard {
    #[inline]
    fn default() -> Self {
        Self::root()
    }
}

/// Maximum lexical XML content-frame nesting admitted by library resource
/// policy, not an OOXML schema restriction. Upstream roxmltree 0.20 recursively
/// parses nonempty elements; preflight must reject unsafe byte streams before
/// entering it, even when a separate content error would also reject the part.
/// Empty-element tags do not open recursive content frames.
pub const MAX_XML_DEPTH: u32 = 256;

/// Scan tag nesting in the byte stream, independently of XML well-formedness.
/// Completed nonempty start tags open a frame, end tags close one (saturating
/// at zero), and empty-element tags open none. Names, duplicate attributes,
/// entities, namespaces and matching end names are left to roxmltree. A lexical
/// crossing is a resource violation even if such a content error occurs first.
///
/// XML 1.0 §§2.3–2.8 and 3.1 delimit the lexical constructs: quotes shield tag
/// delimiters; comments, CDATA and PIs (including the XML declaration) shield
/// their bodies; a DOCTYPE shields its external identifier and internal subset,
/// including quoted declarations, comments and PIs. This is only a depth lexer,
/// not a second XML validator or an entity expander. The DOM disallows DTDs,
/// consistent with ECMA-376 Part 2 §6.2.5 for OPC XML.
///
/// Every cursor advances monotonically, with fixed-size delimiter comparisons:
/// O(bytes scanned), O(1) space, no recursion or allocation. Return immediately
/// on crossing. An unfinished token opens no frame; truncated input that has
/// not already crossed remains an ordinary DOM parse error.
pub fn xml_nesting_exceeds(bytes: &[u8], limit: u32) -> bool {
    let mut depth = 0u32;
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] != b'<' {
            i += 1;
            continue;
        }
        match bytes.get(i + 1) {
            Some(b'?') => i = skip_xml_delimiter(bytes, i + 2, b"?>"),
            Some(b'!') if bytes[i..].starts_with(b"<!--") => {
                i = skip_xml_delimiter(bytes, i + 4, b"-->");
            }
            Some(b'!') if bytes[i..].starts_with(b"<![CDATA[") => {
                i = skip_xml_delimiter(bytes, i + 9, b"]]>");
            }
            Some(b'!') => i = skip_xml_declaration(bytes, i + 2),
            Some(next) => {
                let closing = *next == b'/';
                let Some(end) = xml_tag_end(bytes, i + 1) else {
                    return false;
                };
                if closing {
                    depth = depth.saturating_sub(1);
                } else if bytes[end - 1] != b'/' {
                    // Check before incrementing, including when limit is u32::MAX.
                    if depth == limit {
                        return true;
                    }
                    depth += 1;
                }
                i = end + 1;
            }
            None => break,
        }
    }
    false
}

/// Index of the unquoted tag terminator. An incomplete tag has no terminator.
fn xml_tag_end(bytes: &[u8], mut i: usize) -> Option<usize> {
    let mut quote = None;
    while i < bytes.len() {
        let byte = bytes[i];
        match quote {
            Some(delimiter) if byte == delimiter => quote = None,
            Some(_) => {}
            None if matches!(byte, b'\'' | b'"') => quote = Some(byte),
            None if byte == b'>' => return Some(i),
            None => {}
        }
        i += 1;
    }
    None
}

/// Fixed delimiters only (?>, -->, ]]>), so each comparison costs O(1).
fn skip_xml_delimiter(bytes: &[u8], from: usize, delimiter: &[u8]) -> usize {
    bytes[from..]
        .windows(delimiter.len())
        .position(|w| w == delimiter)
        .map_or(bytes.len(), |offset| from + offset + delimiter.len())
}

/// Skip a markup declaration, including DOCTYPE's bracketed internal subset.
/// Brackets and '>' inside literals, comments and PIs are not delimiters.
fn skip_xml_declaration(bytes: &[u8], mut i: usize) -> usize {
    let mut quote = None;
    let mut brackets = 0usize;
    while i < bytes.len() {
        let byte = bytes[i];
        match quote {
            Some(delimiter) if byte == delimiter => quote = None,
            Some(_) => {}
            None if matches!(byte, b'\'' | b'"') => quote = Some(byte),
            None if bytes[i..].starts_with(b"<!--") => {
                i = skip_xml_delimiter(bytes, i + 4, b"-->");
                continue;
            }
            None if bytes[i..].starts_with(b"<?") => {
                i = skip_xml_delimiter(bytes, i + 2, b"?>");
                continue;
            }
            None if byte == b'[' => brackets += 1,
            None if byte == b']' => brackets = brackets.saturating_sub(1),
            None if byte == b'>' && brackets == 0 => return i + 1,
            None => {}
        }
        i += 1;
    }
    bytes.len()
}

/// Allocation-free lexical preflight for DOM-producing XML complexity.
///
/// A byte/depth bound alone is insufficient: millions of tiny siblings,
/// comments, processing instructions, text nodes, CDATA nodes, or attributes
/// can make `roxmltree` allocate millions of records while the XML stays shallow.
/// `quick_xml` borrows every event directly from `xml`; no event buffer or DOM is
/// allocated. Complexity counts the document node, every element, every
/// comment/CDATA/PI event, every nonempty text/general-reference event, and every
/// attribute (including namespace declarations).
///
/// The scan continues after `limit + 1` so a later lexical error remains an
/// ordinary malformed-XML result in the real parser rather than being promoted
/// to a resource violation. `false` means either within budget or lexically
/// malformed; only a complete valid lexical scan can return `true`.
pub fn xml_dom_complexity_exceeds(xml: &str, limit: u64) -> bool {
    fn add(complexity: &mut u64, amount: u64, limit: u64) {
        *complexity = complexity
            .saturating_add(amount)
            .min(limit.saturating_add(1));
    }

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut complexity = 1u64; // roxmltree's document node
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                add(&mut complexity, 1, limit);
                for attribute in element.attributes() {
                    if attribute.is_err() {
                        return false;
                    }
                    add(&mut complexity, 1, limit);
                }
            }
            Ok(Event::Text(node)) if !node.is_empty() => add(&mut complexity, 1, limit),
            Ok(Event::CData(_)) | Ok(Event::Comment(_)) | Ok(Event::PI(_)) => {
                add(&mut complexity, 1, limit)
            }
            Ok(Event::GeneralRef(node)) if !node.is_empty() => add(&mut complexity, 1, limit),
            Ok(Event::DocType(_)) => return false,
            Ok(Event::Eof) => return complexity > limit,
            Ok(_) => {}
            Err(_) => return false,
        }
    }
}

/// Convenience over [`xml_nesting_exceeds`] using [`MAX_XML_DEPTH`].
/// A lexical crossing counts regardless of other XML content errors.
#[inline]
pub fn xml_too_deep(bytes: &[u8]) -> bool {
    xml_nesting_exceeds(bytes, MAX_XML_DEPTH)
}

/// Error returned by [`parse_guarded`]: a lexical depth crossing or a
/// `roxmltree` parse error for input within the nesting ceiling.
///
/// It implements [`Display`](std::fmt::Display) so the many parse sites that map
/// a parse failure to a `String` (`.map_err(|e| e.to_string())`) keep working
/// unchanged, while sites that only care whether they got a `Document` can keep
/// using `.ok()` / `let Ok(..) else`.
#[derive(Debug)]
pub enum GuardedParseError {
    /// The raw XML nests deeper than [`MAX_XML_DEPTH`]; handing it to
    /// `roxmltree::Document::parse` would risk a stack-overflow trap, so the
    /// caller must reject the package operation instead.
    TooDeep,
    /// `roxmltree::Document::parse` rejected input within the lexical depth ceiling.
    Xml(roxmltree::Error),
}

impl std::fmt::Display for GuardedParseError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            GuardedParseError::TooDeep => {
                write!(f, "XML nesting depth exceeds the safe limit")
            }
            GuardedParseError::Xml(e) => std::fmt::Display::fmt(e, f),
        }
    }
}

impl std::error::Error for GuardedParseError {}

/// Parse XML only after a non-recursive lexical scan establishes compliance
/// with [`MAX_XML_DEPTH`]. Well-formedness is checked by roxmltree. All untrusted
/// package parts use this wrapper so tolerant optional-part handlers cannot
/// accidentally accept a latched resource crossing.
///
/// On success the returned document borrows from `xml`. A depth crossing also
/// poisons the active package operation, so callers may still ignore malformed
/// optional parts without accidentally accepting a resource-limit failure.
#[inline]
pub fn parse_guarded(xml: &str) -> Result<roxmltree::Document<'_>, GuardedParseError> {
    if xml_too_deep(xml.as_bytes()) {
        report_depth_limit(MAX_XML_DEPTH);
        return Err(GuardedParseError::TooDeep);
    }
    roxmltree::Document::parse(xml).map_err(GuardedParseError::Xml)
}

/// Parse depth-guarded XML with an explicit roxmltree node ceiling.
///
/// This is defense in depth for callers that have already performed their own
/// typed, attributable complexity preflight. It deliberately takes the ceiling
/// as an argument and is separate from [`parse_guarded`]: applying an uncalibrated
/// global node cap here would silently change every DOCX/XLSX compatibility path
/// that intentionally degrades ordinary parse failures.
#[inline]
pub fn parse_guarded_with_node_limit(
    xml: &str,
    nodes_limit: u32,
) -> Result<roxmltree::Document<'_>, GuardedParseError> {
    if xml_too_deep(xml.as_bytes()) {
        report_depth_limit(MAX_XML_DEPTH);
        return Err(GuardedParseError::TooDeep);
    }
    roxmltree::Document::parse_with_options(
        xml,
        roxmltree::ParsingOptions {
            allow_dtd: false,
            nodes_limit,
        },
    )
    .map_err(GuardedParseError::Xml)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn root_starts_at_zero_with_shared_limit() {
        let g = DepthGuard::root();
        assert_eq!(g.depth(), 0);
        assert!(!g.is_exhausted());
    }

    #[test]
    fn descend_increments_and_stops_at_limit() {
        // A guard limited to 3 permits descents 0→1→2→3, then refuses.
        let mut g = DepthGuard::with_limit(3);
        for expected in 1..=3 {
            g = g.descend().expect("within limit");
            assert_eq!(g.depth(), expected);
        }
        assert!(g.is_exhausted());
        assert!(g.descend().is_none(), "must refuse past the limit");
    }

    #[test]
    fn siblings_get_independent_children() {
        // `descend` consumes a *copy*, so two children of the same parent are
        // both at depth 1 and neither observes the other's descent.
        let parent = DepthGuard::with_limit(MAX_PARSE_DEPTH);
        let a = parent.descend().unwrap();
        let b = parent.descend().unwrap();
        assert_eq!(a.depth(), 1);
        assert_eq!(b.depth(), 1);
        assert_eq!(a.descend().unwrap().depth(), 2);
        // `b` is still at depth 1 — unaffected by `a`'s further descent.
        assert_eq!(b.depth(), 1);
    }

    #[test]
    fn limit_of_zero_refuses_immediately() {
        let g = DepthGuard::with_limit(0);
        assert!(g.is_exhausted());
        assert!(g.descend().is_none());
    }

    #[test]
    fn default_matches_root() {
        assert_eq!(DepthGuard::default(), DepthGuard::root());
    }

    /// A deep linear chain — the failure this guard exists to prevent — is bounded:
    /// the recursion runs at most `MAX_PARSE_DEPTH` descents and returns rather
    /// than trapping. This mirrors the shape of the real recursive parsers.
    #[test]
    fn deep_chain_is_bounded_not_trapping() {
        fn descend_chain(g: DepthGuard) -> u32 {
            match g.descend() {
                Some(child) => 1 + descend_chain(child),
                None => 0,
            }
        }
        assert_eq!(descend_chain(DepthGuard::root()), MAX_PARSE_DEPTH);
    }

    // ── Raw-XML depth pre-check (guards roxmltree's own recursion) ──────────

    /// `depth` levels of `<a>…</a>` wrapping a text leaf.
    fn nested_xml(depth: usize) -> Vec<u8> {
        let mut s = String::new();
        for _ in 0..depth {
            s.push_str("<a>");
        }
        s.push('x');
        for _ in 0..depth {
            s.push_str("</a>");
        }
        s.into_bytes()
    }

    #[test]
    fn shallow_xml_is_not_flagged() {
        assert!(!xml_nesting_exceeds(&nested_xml(10), 256));
        // A real-ish OOXML snippet nests well under the limit.
        let doc = br#"<w:document xmlns:w="x"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>"#;
        assert!(!xml_too_deep(doc));
    }

    #[test]
    fn exactly_at_the_limit_is_allowed_and_one_past_is_flagged() {
        // `limit` levels deep is fine; `limit + 1` trips the guard.
        assert!(!xml_nesting_exceeds(&nested_xml(32), 32));
        assert!(xml_nesting_exceeds(&nested_xml(33), 32));
    }

    #[test]
    fn pathologically_deep_xml_is_flagged() {
        // The case this pre-check exists for: thousands deep would otherwise
        // overflow roxmltree's recursive tree builder.
        assert!(xml_too_deep(&nested_xml(10_000)));
    }

    #[test]
    fn self_closing_tags_do_not_accumulate_depth() {
        // 1000 sibling self-closing tags are all leaves — net depth stays 1.
        let mut s = String::from("<root>");
        for _ in 0..1000 {
            s.push_str("<a/>");
        }
        s.push_str("</root>");
        assert!(!xml_nesting_exceeds(s.as_bytes(), 8));
        // With whitespace before the '/>', still a leaf.
        assert!(!xml_nesting_exceeds(b"<root><a b=\"1\" /></root>", 4));
    }

    #[test]
    fn quoted_tag_delimiters_do_not_turn_siblings_into_nesting() {
        for quote in ['\'', '"'] {
            let mut xml = String::from("<styles>");
            for id in 0..257 {
                xml.push_str(&format!("<style id={quote}s{id}>{quote}/>"));
            }
            xml.push_str("</styles>");
            assert!(!xml_too_deep(xml.as_bytes()));
            assert_eq!(
                parse_guarded(&xml)
                    .unwrap()
                    .root_element()
                    .children()
                    .count(),
                257
            );
        }
        // A quoted '/>' must not hide genuine nesting either.
        assert!(xml_nesting_exceeds(
            br#"<a x="/>"><b x='>'><c/></b></a>"#,
            1
        ));
    }

    #[test]
    fn comments_pis_and_cdata_do_not_affect_depth() {
        // Stray '<' and '>' inside comments / PIs / CDATA must not be counted as
        // tags (or the guard could be tricked either way).
        let xml =
            br#"<r><!-- <a><b><c> not tags --><?pi <x><y> ?><![CDATA[ </z><w><v> ]]><c/></r>"#;
        // Real nesting is only <r> … </r> plus a leaf <c/>: depth 1.
        assert!(!xml_nesting_exceeds(xml, 3));
    }

    #[test]
    fn unbalanced_end_tags_saturate_not_underflow() {
        // Extra end tags must not underflow (u32) and wrap to a huge depth.
        let xml = b"</a></a></a><b><c></c></b>";
        assert!(!xml_nesting_exceeds(xml, 4));
    }

    #[test]
    fn depth_scan_handles_large_trailing_text() {
        // Lexical confirmation stays bounded even with a large trailing text token.
        let mut v = b"<r>".to_vec();
        v.extend(nested_xml(300));
        v.extend(std::iter::repeat_n(b'x', 1_000_000));
        v.extend(b"</r>");
        assert!(xml_nesting_exceeds(&v, 256));
    }

    // ── parse_guarded (the wrapper every attacker-facing parse site must use) ──

    #[test]
    fn parse_guarded_accepts_shallow_xml() {
        let doc = parse_guarded("<r><a>x</a></r>").expect("shallow XML parses");
        assert_eq!(doc.root_element().tag_name().name(), "r");
    }

    #[test]
    fn parse_guarded_rejects_too_deep_before_dom_allocation() {
        // Reject an extreme lexical chain without constructing a DOM or
        // recursing per element.
        let deep = String::from_utf8(nested_xml(10_000)).unwrap();
        let err = parse_guarded(&deep).expect_err("over-deep XML is rejected");
        assert!(matches!(err, GuardedParseError::TooDeep));
        assert_eq!(err.to_string(), "XML nesting depth exceeds the safe limit");
    }

    #[test]
    fn parse_guarded_reports_roxmltree_errors_for_in_budget_input() {
        // In-depth-budget but malformed input surfaces roxmltree's own error,
        // Display-forwarded so `.map_err(|e| e.to_string())` sites are unchanged.
        let err = parse_guarded("<r><a></r>").expect_err("malformed XML is rejected");
        assert!(matches!(err, GuardedParseError::Xml(_)));
        // The message is roxmltree's, not the depth message.
        assert_ne!(err.to_string(), "XML nesting depth exceeds the safe limit");
    }

    #[test]
    fn lexical_crossings_ignore_content_errors_before_and_after_the_chain() {
        use crate::resource::{OoxmlFormat, ResourceGovernor};
        let deep = String::from_utf8(nested_xml(257)).unwrap();
        for error in ["<x a='1' a='2'/>", "&unknown;", "<missing:x/>", "</wrong>"] {
            for xml in [
                format!("<r>{error}{deep}</r>"),
                format!("<r>{deep}{error}</r>"),
            ] {
                let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, None, None, None);
                let _scope = governor.scope("lexical-depth");
                for result in [
                    parse_guarded(&xml),
                    parse_guarded_with_node_limit(&xml, 1000),
                ] {
                    assert!(matches!(result, Err(GuardedParseError::TooDeep)), "{xml}");
                }
                assert!(governor.first_error().is_some());
            }
        }
    }

    #[test]
    fn lexical_constructs_have_exact_content_frame_depth() {
        // Empty elements never open a content frame in roxmltree, so neither
        // form of empty-element tag contributes to this stack safety ceiling.
        for (xml, exact) in [
            ("<x/>", 0),
            ("<r><x/><x a='1' /></r>", 1),
            ("<r><a></a><b><c></c></b></r>", 3),
            (r#"<r a="< > / '" b='< > / "'><a x='/> >'></a></r>"#, 2),
            ("<r><!-- <x><x> </r> --><a></a></r>", 2),
            ("<r><![CDATA[ <x><x> </r> ]]><a></a></r>", 2),
            ("<?xml version='1.0'?><?pi '<x><x>' ?><r><a></a></r>", 2),
            (r#"<!DOCTYPE r SYSTEM 'a>b<c'><r><a></a></r>"#, 2),
            (
                r#"<!DOCTYPE r [<!ELEMENT r ANY><!ENTITY x '<a><b> ] >'><!ENTITY y "[ > <c>"><!-- ] > <x> --><?pi ] > <x> ?>]><r><a></a></r>"#,
                2,
            ),
            ("</extra></extra><r><a></wrong></r>", 2),
        ] {
            assert!(!xml_nesting_exceeds(xml.as_bytes(), exact), "{xml}");
            if exact > 0 {
                assert!(xml_nesting_exceeds(xml.as_bytes(), exact - 1), "{xml}");
            }
        }
    }

    #[test]
    fn truncation_only_crosses_after_complete_start_tags() {
        for suffix in [
            "<",
            "<a",
            "<a x='>",
            "</a",
            "<!-- <a>",
            "<![CDATA[ <a>",
            "<?pi <a>",
            "<!DOCTYPE r [<!ENTITY x '<a>'>",
        ] {
            let xml = format!("<r>{suffix}");
            assert!(!xml_nesting_exceeds(xml.as_bytes(), 1), "{xml}");
            assert!(matches!(
                parse_guarded(&xml),
                Err(GuardedParseError::Xml(_))
            ));
        }
        assert!(xml_nesting_exceeds(b"<r><a><", 1));
        // The scanner operates on bytes and has no UTF-8 validation dependency.
        assert!(xml_nesting_exceeds(b"<r>\xff<a></a></r>", 1));
    }

    #[test]
    fn extreme_lexical_depth_is_safe_on_a_small_stack() {
        std::thread::Builder::new()
            .stack_size(128 * 1024)
            .spawn(|| {
                let deep = String::from_utf8(nested_xml(120_000)).unwrap();
                for xml in [deep.clone(), format!("<r a='1' a='2'>{deep}</r>")] {
                    assert!(matches!(
                        parse_guarded(&xml),
                        Err(GuardedParseError::TooDeep)
                    ));
                }
            })
            .unwrap()
            .join()
            .unwrap();
    }

    #[test]
    fn parse_guarded_document_borrows_from_input() {
        // The returned Document borrows `xml`, just like Document::parse — so a
        // caller can read node text without an extra allocation.
        let xml = String::from("<r><t>hello</t></r>");
        let doc = parse_guarded(&xml).unwrap();
        let t = doc
            .descendants()
            .find(|n| n.tag_name().name() == "t")
            .unwrap();
        assert_eq!(t.text(), Some("hello"));
    }

    #[test]
    fn explicit_node_limit_does_not_change_shared_parse_guarded_compatibility() {
        let mut xml = String::from("<r>");
        for _ in 0..200_000 {
            xml.push_str("<a/>");
        }
        xml.push_str("</r>");

        let document = parse_guarded(&xml)
            .expect("shared DOCX/XLSX-compatible parser has no implicit node ceiling");
        assert_eq!(document.descendants().count(), 200_002);
        assert!(matches!(
            parse_guarded_with_node_limit(&xml, 200_001),
            Err(GuardedParseError::Xml(_))
        ));
    }

    #[test]
    fn xml_dom_complexity_counts_nodes_attributes_and_namespaces() {
        let xml = r#"<?xml version="1.0"?><r xmlns="urn:x" a="1"><!-- c --><![CDATA[d]]><?pi x?><a/>text</r>"#;
        // document + r + 2 attrs + comment + cdata + PI + a + text = 9
        assert!(!xml_dom_complexity_exceeds(xml, 9));
        assert!(xml_dom_complexity_exceeds(xml, 8));
    }

    #[test]
    fn xml_dom_complexity_exact_boundary_covers_each_dom_producing_unit() {
        fn assert_exact(xml: &str, exact: u64) {
            assert!(
                !xml_dom_complexity_exceeds(xml, exact),
                "exact limit must accept {xml}"
            );
            assert!(
                xml_dom_complexity_exceeds(xml, exact - 1),
                "limit+1 must reject {xml}"
            );
        }

        // Each expectation includes roxmltree's document node and root element.
        assert_exact("<r><a/><b/></r>", 4); // two child elements
        assert_exact("<r><!--a--><!--b--></r>", 4); // two comments
        assert_exact("<r><?a x?><?b y?></r>", 4); // two processing instructions
        assert_exact("<r><![CDATA[a]]><![CDATA[b]]></r>", 4); // two CDATA events
        assert_exact("<r>a<a/>b</r>", 5); // two text nodes + child element
        assert_exact(r#"<r xmlns="urn:x" a="1" b="2"/>"#, 5); // three attributes
    }

    #[test]
    fn malformed_xml_is_left_to_the_real_parser_even_after_crossing_budget() {
        assert!(!xml_dom_complexity_exceeds("<r><a/><a/><broken", 1));
    }
}
