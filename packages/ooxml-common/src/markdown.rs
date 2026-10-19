//! Markdown projection helpers shared by the DOCX, PPTX and XLSX converters.
//!
//! Markdown output is consumed by renderers and LLM chat UIs that may turn it
//! into HTML, so document data must not become live markup: link
//! destinations follow the viewer's hyperlink scheme policy (the TypeScript
//! `sanitizeHyperlinkUrl` in `packages/core/src/interaction/hyperlink.ts`),
//! and inline text cannot open a link, image or raw-HTML construct.

/// Schemes a projected link may carry; mirrors the viewer's
/// `DEFAULT_ALLOWED_HYPERLINK_SCHEMES`.
const ALLOWED_LINK_SCHEMES: [&str; 4] = ["http", "https", "mailto", "tel"];

/// Characters a URL consumer may strip or trim before resolving the scheme:
/// ASCII controls and space (browsers drop embedded TAB/LF/CR), Unicode
/// whitespace such as U+00A0 (Markdown renderers trim it), and U+FEFF.
fn is_ignorable(c: char) -> bool {
    c <= ' ' || c.is_whitespace() || c.is_control() || c == '\u{feff}'
}

/// The lowercased RFC 3986 scheme of `url`, ignoring the characters in
/// [`is_ignorable`] (so `java\tscript:` or a leading U+00A0 cannot hide it).
fn link_scheme(url: &str) -> Option<String> {
    let cleaned: String = url.chars().filter(|&c| !is_ignorable(c)).collect();
    let (scheme, _) = cleaned.split_once(':')?;
    let mut chars = scheme.chars();
    let first = chars.next()?;
    (first.is_ascii_alphabetic()
        && chars.all(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '-' | '.')))
    .then(|| scheme.to_ascii_lowercase())
}

/// A Markdown link destination for a document hyperlink target, or `None` when
/// the target must not become a link (an empty target or a scheme outside the
/// allowlist, such as `javascript:` or `data:`); the caller then emits the link
/// text alone. A scheme-less (relative) reference cannot name a dangerous
/// scheme and is kept. Characters that would end or restructure the
/// destination (any whitespace, controls, U+FEFF, `(`, `)`, `<`, `>`, `\`) are
/// percent-encoded so the target cannot break out into further Markdown, and
/// `&` is written as `&amp;` so CommonMark entity decoding (`&colon;`,
/// `&#58;`) reproduces the checked target instead of forming a new scheme.
pub fn link_destination(url: &str) -> Option<String> {
    if url.is_empty() {
        return None;
    }
    if let Some(scheme) = link_scheme(url) {
        if !ALLOWED_LINK_SCHEMES.contains(&scheme.as_str()) {
            return None;
        }
    }
    let mut out = String::with_capacity(url.len());
    for c in url.chars() {
        if is_ignorable(c) || matches!(c, '(' | ')' | '<' | '>' | '\\') {
            let mut buf = [0u8; 4];
            for byte in c.encode_utf8(&mut buf).bytes() {
                out.push_str(&format!("%{byte:02X}"));
            }
        } else if c == '&' {
            out.push_str("&amp;");
        } else {
            out.push(c);
        }
    }
    Some(out)
}

/// True for characters escaped wherever document text reaches Markdown.
/// CommonMark §2.4 lets any ASCII punctuation be backslash-escaped with the
/// rendered text unchanged, so the set errs towards escaping:
///
/// - formatting: `\ * _ \`` and the GFM strikethrough `~` (escaping every
///   backtick and tilde also means no code fence can open);
/// - `&`, which would start an entity or numeric character reference;
/// - `[`, `]`, `<` and `!`, which would open a link, image, raw HTML or a
///   CommonMark autolink (`!` would turn a following projected link into an
///   image);
/// - `|`, which would build a GFM table row or end a table cell;
/// - `:` and `@`, which every GFM extended scheme (`https:`, `mailto:`) and
///   email autolink needs.
pub fn is_inline_metachar(c: char) -> bool {
    matches!(
        c,
        '\\' | '*' | '_' | '`' | '~' | '&' | '[' | ']' | '<' | '!' | '|' | ':' | '@'
    )
}

/// Where projected document text is placed; selects the extra handling on top
/// of the escaping every context shares.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum TextContext {
    /// Running text (paragraphs, list items, notes, quoted comment lines).
    /// Line breaks are kept, and the start of the text and of every line is
    /// treated as a line start (see [`escape_text`]).
    Inline,
    /// An ATX heading label: line breaks become spaces (a heading is one
    /// line) and `#` is escaped so the text cannot form a closing sequence.
    Heading,
    /// A GFM table cell: line breaks become `<br>` so the text stays inside
    /// its cell and row.
    TableCell,
}

/// Spacing the line-start block-marker scan looks past: all Unicode
/// whitespace plus U+FEFF, a superset of what `str::trim` removes and of what
/// Markdown renderers strip, so trimming or stripping can never expose a
/// marker. Only its ASCII spaces and tabs are ever dropped.
fn is_indentation(c: char) -> bool {
    c.is_whitespace() || c == '\u{feff}'
}

/// Escape document text for the Markdown projection so it reads as literal
/// text in every context and across fragment boundaries:
///
/// - every [`is_inline_metachar`] is backslash-escaped;
/// - a `.` is escaped when it follows `www` (case-insensitive) or when only
///   `w`s precede it in this fragment, because the rest of `www.` may end the
///   previous fragment; this defuses GFM `www.` autolinks;
/// - in [`TextContext::Inline`], a block marker at a possible line start (the
///   start of the text and after every line break, past any
///   [`is_indentation`]) is escaped, so the text cannot open an ATX heading
///   (`#`), block quote (`>`), list item (`-`, `+`, `1.`, `1)`), setext
///   underline or thematic break (`=`, `-`);
/// - in [`TextContext::Inline`], ASCII spaces and tabs at a projected line
///   start are dropped: four or more spaces (or a tab) would open an indented
///   code block, in which the escapes would render literally, and CommonMark
///   ignores up to three leading spaces of a paragraph line anyway. A BOM
///   there is invisible noise that renderers strip, so it is dropped together
///   with the indentation it would otherwise hide. Other Unicode spacing
///   (NBSP, U+3000, EM SPACE) is visible text that cannot indent, so it is
///   kept, as is a BOM in mid-line; a block marker after it is still escaped;
/// - [`TextContext::Heading`] also escapes `#` and folds line breaks into
///   spaces; [`TextContext::TableCell`] turns line breaks into `<br>`.
///
/// The text is taken to start a line. Writers that join fragments within a
/// line use [`escape_fragment`] instead. Callers trim document text before
/// escaping it, never the escaped output. Markdown the converter emits itself
/// (heading markers, quote prefixes, `<br>` between paragraphs, link syntax
/// for an allowed destination) is written around the result, never through
/// it. The output is at most four bytes per input byte (a lone `\n` becoming
/// `<br>`; every other rule at most doubles), which projection size bounds
/// rely on.
pub fn escape_text(text: &str, context: TextContext) -> String {
    escape_fragment(text, context, &mut true)
}

/// [`escape_text`] for one fragment of a line that the caller assembles from
/// several pieces (runs, wrappers, breaks). `line_start` says whether the
/// output so far ends at a line start (nothing but dropped indentation since
/// the last line break) and is updated for the next piece: indentation is
/// dropped only at a real line start, so whitespace between runs in mid-line
/// is kept, while indentation that arrives in its own run before a line's
/// first character is still dropped. Block markers are escaped at every
/// fragment start regardless, because a marker may be split across runs
/// (`1` + `) x`).
pub fn escape_fragment(text: &str, context: TextContext, line_start: &mut bool) -> String {
    let inline = context == TextContext::Inline;
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    // A block-marker scan is pending, and whether indentation is dropped.
    let mut scan = inline;
    let mut drop_indent = inline && *line_start;
    // Consecutive `w`s just written, and whether nothing else has been.
    let mut w_run = 0usize;
    let mut only_ws = true;
    while let Some(c) = rest.chars().next() {
        if scan {
            scan = false;
            if drop_indent {
                // CommonMark indentation is ASCII spaces and tabs only; other
                // Unicode spacing is visible text and is kept. A BOM here is
                // invisible noise that renderers strip (CommonMark drops a
                // document-initial one), which would expose the indentation
                // after it, so it is dropped along with that indentation.
                rest = rest.trim_start_matches([' ', '\t', '\u{feff}']);
            }
            let body = rest.trim_start_matches(|c| is_indentation(c) && !matches!(c, '\r' | '\n'));
            if body.len() != rest.len() {
                out.push_str(&rest[..rest.len() - body.len()]);
                (w_run, only_ws) = (0, false);
                // The line has begun with visible spacing: later ASCII
                // spaces are no longer indentation.
                drop_indent = false;
            }
            if body.is_empty() {
                // Indentation only: the line start (if any) carries over.
                break;
            }
            drop_indent = false;
            rest = body;
            if let Some(marker) = block_marker_offset(rest) {
                // Digits before an ordered-list delimiter, then the escaped
                // marker itself (ASCII punctuation no other rule escapes).
                out.push_str(&rest[..marker]);
                out.push('\\');
                out.push_str(&rest[marker..=marker]);
                rest = &rest[marker + 1..];
                (w_run, only_ws) = (0, false);
            }
            continue;
        }
        rest = &rest[c.len_utf8()..];
        if matches!(c, '\r' | '\n') {
            let crlf = c == '\r' && rest.starts_with('\n');
            match context {
                TextContext::Inline => {
                    out.push_str(if crlf {
                        "\r\n"
                    } else if c == '\r' {
                        "\r"
                    } else {
                        "\n"
                    });
                    (scan, drop_indent) = (true, true);
                }
                TextContext::Heading => out.push(' '),
                TextContext::TableCell => out.push_str("<br>"),
            }
            if crlf {
                rest = &rest[1..];
            }
            (w_run, only_ws) = (0, false);
            continue;
        }
        if is_inline_metachar(c)
            || (c == '.' && (w_run >= 3 || only_ws))
            || (context == TextContext::Heading && c == '#')
        {
            out.push('\\');
        }
        out.push(c);
        if c.eq_ignore_ascii_case(&'w') {
            w_run += 1;
        } else {
            (w_run, only_ws) = (0, false);
        }
    }
    if inline {
        // Still at a line start after a fragment of dropped indentation, or
        // after a trailing line break.
        *line_start = drop_indent;
    } else if !out.is_empty() {
        *line_start = false;
    }
    out
}

/// Whether projected output that ends with `markup` (Markdown the converter
/// emits itself, written between escaped fragments) is at a line start;
/// `line_start` is the state before it.
pub fn line_start_after_markup(markup: &str, line_start: bool) -> bool {
    if markup.is_empty() {
        line_start
    } else {
        markup.ends_with(['\n', '\r'])
    }
}

/// Byte offset of the character to escape when `line` (with its indentation
/// removed) would begin a CommonMark block: an ATX heading, block quote,
/// bullet or ordered list item, setext underline or thematic break. Markers
/// that need a following space are escaped when any [`is_indentation`]
/// character or the end of the text (which another fragment may continue, or
/// a trim may produce) follows. A bare `.`/`)` counts as an ordered-list
/// delimiter because the digits may end a preceding fragment. `*`, `_`, `|`,
/// `` ` `` and `~` never reach here unescaped.
fn block_marker_offset(line: &str) -> Option<usize> {
    let bytes = line.as_bytes();
    let ends_marker = |index: usize| line[index..].chars().next().is_none_or(is_indentation);
    match *bytes.first()? {
        b'>' => Some(0),
        b'#' => {
            let hashes = bytes.iter().take_while(|&&b| b == b'#').count();
            (hashes <= 6 && ends_marker(hashes)).then_some(0)
        }
        b'-' => (ends_marker(1) || bytes.get(1) == Some(&b'-')).then_some(0),
        b'=' => (ends_marker(1) || bytes.get(1) == Some(&b'=')).then_some(0),
        b'+' => ends_marker(1).then_some(0),
        _ => {
            let digits = bytes.iter().take_while(|b| b.is_ascii_digit()).count();
            (digits <= 9
                && matches!(bytes.get(digits), Some(b'.' | b')'))
                && ends_marker(digits + 1))
            .then_some(digits)
        }
    }
}

/// Split `text` into lines at CRLF, CR and LF alike. CommonMark ends a line at
/// each of them, so a quote prefix must be written after every one; a final
/// line break does not start an extra empty line.
pub fn split_lines(text: &str) -> impl Iterator<Item = &str> {
    let mut rest = Some(text);
    std::iter::from_fn(move || {
        let current = rest?;
        match current.find(['\r', '\n']) {
            Some(index) => {
                let next = if current[index..].starts_with("\r\n") {
                    index + 2
                } else {
                    index + 1
                };
                rest = (next < current.len()).then(|| &current[next..]);
                Some(&current[..index])
            }
            None => {
                rest = None;
                Some(current)
            }
        }
    })
}

/// Upper bound on the lines [`split_lines`] yields for `text`.
pub fn line_count_bound(text: &str) -> u64 {
    text.bytes().filter(|&b| b == b'\r' || b == b'\n').count() as u64 + 1
}

/// Test support: a GFM structure summary of projected Markdown, parsed with
/// tables, strikethrough, extended autolinks, task lists and footnotes on, so
/// regression tests can prove document text formed no Markdown structure.
#[cfg(feature = "gfm-check")]
pub mod gfm_check {
    use comrak::nodes::NodeValue;
    use comrak::{parse_document, Arena, Options};

    /// Counts of the constructs document text must never create, plus the
    /// plain text of every heading and table cell for structure assertions.
    #[derive(Debug, Default, PartialEq, Eq)]
    pub struct GfmSummary {
        /// Rendered text of the whole document.
        pub text: String,
        pub headings: Vec<String>,
        pub table_cells: Vec<String>,
        pub tables: usize,
        pub links: usize,
        pub images: usize,
        pub html: usize,
        pub strikethrough: usize,
        pub code: usize,
        pub block_quotes: usize,
        pub list_items: usize,
        pub thematic_breaks: usize,
    }

    pub fn summarize(markdown: &str) -> GfmSummary {
        let arena = Arena::new();
        let mut options = Options::default();
        options.extension.table = true;
        options.extension.strikethrough = true;
        options.extension.autolink = true;
        options.extension.tasklist = true;
        options.extension.footnotes = true;
        let root = parse_document(&arena, markdown, &options);
        let mut summary = GfmSummary::default();
        for node in root.descendants() {
            match &node.data.borrow().value {
                NodeValue::Heading(_) => summary.headings.push(plain_text(node)),
                NodeValue::TableCell => summary.table_cells.push(plain_text(node)),
                NodeValue::Table(_) => summary.tables += 1,
                NodeValue::Link(_) => summary.links += 1,
                NodeValue::Image(_) => summary.images += 1,
                NodeValue::HtmlInline(html) if html == "<br>" => {}
                NodeValue::HtmlInline(_) | NodeValue::HtmlBlock(_) => summary.html += 1,
                NodeValue::Strikethrough => summary.strikethrough += 1,
                NodeValue::Code(_) | NodeValue::CodeBlock(_) => summary.code += 1,
                NodeValue::BlockQuote => summary.block_quotes += 1,
                NodeValue::Item(_) | NodeValue::TaskItem(_) => summary.list_items += 1,
                NodeValue::ThematicBreak => summary.thematic_breaks += 1,
                _ => {}
            }
        }
        summary.text = plain_text(root);
        summary
    }

    /// Rendered text below `node`: line breaks read as spaces, and nested
    /// paragraphs, headings and cells start on a new line.
    fn plain_text<'a>(node: &'a comrak::nodes::AstNode<'a>) -> String {
        let mut text = String::new();
        for child in node.descendants() {
            match &child.data.borrow().value {
                NodeValue::Text(value) => text.push_str(value),
                NodeValue::SoftBreak | NodeValue::LineBreak => text.push(' '),
                NodeValue::Paragraph | NodeValue::Heading(_) | NodeValue::TableCell
                    if !std::ptr::eq(child, node) =>
                {
                    text.push('\n')
                }
                _ => {}
            }
        }
        text
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn document_text_cannot_form_links_images_or_html() {
        let attack = "<img src=x onerror=alert(1)> [x](javascript:alert(1)) ![i](y)";
        let escaped = r"\<img src=x onerror=alert(1)> \[x\](javascript\:alert(1)) \!\[i\](y)";
        assert_eq!(escape_text(attack, TextContext::Inline), escaped);
        assert_eq!(escape_text(attack, TextContext::Heading), escaped);
        assert_eq!(escape_text(attack, TextContext::TableCell), escaped);
        // Ordinary punctuation is untouched, including mid-line block markers.
        let plain = "Q3 revenue $1,200 (+5%) - see p. 4; a/b, -5 and =SUM(A1) 2026.10";
        for context in [
            TextContext::Inline,
            TextContext::Heading,
            TextContext::TableCell,
        ] {
            assert_eq!(escape_text(plain, context), plain);
        }
    }

    #[test]
    fn tables_and_autolinks_cannot_form() {
        let text =
            "| a | b |\n| :-- | --: |\nforged | row https://e.example/ www.e.example a@e.example";
        let escaped = "\\| a \\| b \\|\n\\| \\:-- \\| --\\: \\|\nforged \\| row https\\://e.example/ www\\.e.example a\\@e.example";
        assert_eq!(escape_text(text, TextContext::Inline), escaped);
        // `www.` split across fragments: any `.` after only `w`s is escaped.
        for (fragment, escaped) in [
            (".e.example", r"\.e.example"),
            ("w.e", r"w\.e"),
            ("WW.e", r"WW\.e"),
            ("xwww.e", r"xwww\.e"),
            ("aw.e", "aw.e"),
        ] {
            assert_eq!(escape_text(fragment, TextContext::TableCell), escaped);
        }
    }

    #[test]
    fn unicode_indentation_cannot_hide_a_line_start_marker() {
        for (line, escaped) in [
            ("\u{a0}# forged", "\u{a0}\\# forged"),
            ("\u{2003}> forged", "\u{2003}\\> forged"),
            // A line-start BOM is dropped with the indentation after it; a
            // BOM in mid-line or after visible spacing stays.
            ("\u{feff}# forged", "\\# forged"),
            ("\u{feff}    Cost: 1", "Cost\\: 1"),
            ("\u{feff}\tCost: 1", "Cost\\: 1"),
            ("x\u{feff}y", "x\u{feff}y"),
            ("\u{a0}\u{feff}y", "\u{a0}\u{feff}y"),
            ("  \u{3000}  - x", "\u{3000}  \\- x"),
            ("#\u{a0}forged", "\\#\u{a0}forged"),
            ("-\u{3000}item", "\\-\u{3000}item"),
        ] {
            assert_eq!(escape_text(line, TextContext::Inline), escaped, "{line:?}");
        }
    }

    #[test]
    fn line_start_indentation_cannot_open_a_code_block() {
        assert_eq!(
            escape_text(
                "    Cost: 1\n\tTab\r\n  \u{a0}x\r    y",
                TextContext::Inline
            ),
            "Cost\\: 1\nTab\r\n\u{a0}x\ry"
        );
        // Cells and headings are inline-only: nothing is dropped there.
        assert_eq!(escape_text("    a", TextContext::TableCell), "    a");

        // Fragments: indentation in its own run before the line's first
        // character is dropped; whitespace between runs in mid-line is kept.
        let mut line_start = true;
        let pieces: Vec<String> = ["  ", "\t ", "Cost", " ", "x", "  \n", "    ", "y"]
            .iter()
            .map(|piece| escape_fragment(piece, TextContext::Inline, &mut line_start))
            .collect();
        assert_eq!(pieces.concat(), "Cost x  \ny");
        assert!(!line_start);
        let mut line_start = false;
        assert_eq!(
            escape_fragment("  mid", TextContext::Inline, &mut line_start),
            "  mid"
        );
        assert!(line_start_after_markup("  \n", false));
        assert!(!line_start_after_markup("**", true));
        assert!(line_start_after_markup("", true));
    }

    #[test]
    fn entities_strikethrough_and_fences_stay_literal() {
        assert_eq!(
            escape_text("&copy; &#60; ~~gone~~", TextContext::Inline),
            r"\&copy; \&#60; \~\~gone\~\~"
        );
        assert_eq!(
            escape_text("x\n```\n~~~", TextContext::Inline),
            "x\n\\`\\`\\`\n\\~\\~\\~"
        );
    }

    #[test]
    fn line_start_block_markers_are_escaped_in_running_text() {
        for (line, escaped) in [
            ("# forged", r"\# forged"),
            ("###", r"\###"),
            ("  > quote", r"\> quote"),
            (">quote", r"\>quote"),
            ("- item", r"\- item"),
            ("---", r"\---"),
            ("+ item", r"\+ item"),
            ("===", r"\==="),
            ("1. item", r"1\. item"),
            ("12) item", r"12\) item"),
            (". rest of a split 1.", r"\. rest of a split 1."),
            ("-", r"\-"),
            // Not block markers: left alone.
            ("#tag", "#tag"),
            ("-5%", "-5%"),
            ("+1", "+1"),
            ("=SUM(A1)", "=SUM(A1)"),
            ("2026.10", "2026.10"),
            ("1234567890. big", "1234567890. big"),
        ] {
            assert_eq!(escape_text(line, TextContext::Inline), escaped, "{line:?}");
        }
        assert_eq!(
            escape_text("a\r# b\r\n> c\n- d", TextContext::Inline),
            "a\r\\# b\r\n\\> c\n\\- d"
        );
        // Headings and cells are inline-only constructs.
        assert_eq!(escape_text("- a", TextContext::TableCell), "- a");
        assert_eq!(escape_text("> a", TextContext::Heading), "> a");
    }

    #[test]
    fn contexts_keep_text_inside_its_construct() {
        let text = "a|b #1\r\nc\rd\ne";
        assert_eq!(
            escape_text(text, TextContext::Inline),
            "a\\|b #1\r\nc\rd\ne"
        );
        assert_eq!(escape_text(text, TextContext::Heading), r"a\|b \#1 c d e");
        assert_eq!(
            escape_text(text, TextContext::TableCell),
            r"a\|b #1<br>c<br>d<br>e"
        );
    }

    #[test]
    fn lines_split_at_cr_lf_and_crlf() {
        let text = "a\r\rb\r\nc\nd\n";
        assert_eq!(
            split_lines(text).collect::<Vec<_>>(),
            ["a", "", "b", "c", "d"]
        );
        assert!(split_lines(text).count() as u64 <= line_count_bound(text));
        assert_eq!(split_lines("").collect::<Vec<_>>(), [""]);
    }

    #[test]
    fn link_destinations_follow_the_viewer_scheme_policy() {
        assert_eq!(
            link_destination("https://example.com/a?b=c#d").as_deref(),
            Some("https://example.com/a?b=c#d")
        );
        assert_eq!(
            link_destination("MAILTO:x@example.com").as_deref(),
            Some("MAILTO:x@example.com")
        );
        assert_eq!(
            link_destination("docs/page.html").as_deref(),
            Some("docs/page.html")
        );
        for blocked in [
            "javascript:alert(1)",
            " java\tscript:alert(1)",
            "data:text/html,x",
            "vbscript:x",
            "file:///etc/passwd",
            "\u{a0}vscode://example.com/x",
            "\u{feff}javascript:alert(1)",
            "",
        ] {
            assert_eq!(link_destination(blocked), None, "{blocked:?}");
        }
        // Entity forms stay literal after CommonMark decoding.
        assert_eq!(
            link_destination("javascript&colon;alert(1)").as_deref(),
            Some("javascript&amp;colon;alert%281%29")
        );
        assert_eq!(
            link_destination("https://e.example/a\u{a0}b").as_deref(),
            Some("https://e.example/a%C2%A0b")
        );
        assert_eq!(
            link_destination("https://e.example/?a=1&b=2").as_deref(),
            Some("https://e.example/?a=1&amp;b=2")
        );
        assert_eq!(
            link_destination("https://e.example/x) [y](https://evil.example/").as_deref(),
            Some("https://e.example/x%29%20[y]%28https://evil.example/")
        );
    }
}
