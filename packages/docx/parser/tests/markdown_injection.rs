//! Markdown output is rendered by downstream Markdown/HTML consumers, so
//! hyperlink targets and run text from an untrusted document must not become
//! script links, injected links or raw HTML.

use docx_parser::to_markdown_native;
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn docx_with_links(links: &[(&str, &str)], body: &str) -> Vec<u8> {
    let relationships: String = links
        .iter()
        .map(|(id, target)| {
            format!(
                r#"<Relationship Id="{id}" Type="{R_NS}/hyperlink" Target="{target}" TargetMode="External"/>"#
            )
        })
        .collect();
    let document = format!(
        r#"<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    );
    let rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#
    );
    package(&[
        ("word/document.xml", document),
        ("word/_rels/document.xml.rels", rels),
    ])
}

fn package(parts: &[(&str, String)]) -> Vec<u8> {
    let mut bytes = Vec::new();
    {
        let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
        zip.start_file("[Content_Types].xml", SimpleFileOptions::default())
            .unwrap();
        zip.write_all(
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#,
        )
        .unwrap();
        for (name, text) in parts {
            zip.start_file(*name, SimpleFileOptions::default()).unwrap();
            zip.write_all(text.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    bytes
}

fn link(id: &str, text: &str) -> String {
    format!(
        r#"<w:hyperlink r:id="{id}"><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:hyperlink>"#
    )
}

#[test]
fn hyperlinks_and_text_cannot_inject_markdown_links_or_html() {
    let body = format!(
        "<w:p>{}{}{}</w:p><w:p><w:r><w:t>&lt;img src=x onerror=alert(1)&gt; ![p](https://t.example/p.png)</w:t></w:r></w:p><w:p><w:r><w:t>!</w:t></w:r>{}</w:p><w:p>{}</w:p>",
        link("rJs", "script"),
        link("rBreak", "safe"),
        link("rOk", "ok"),
        link("rOk", "pic"),
        link("rEntity", "entity"),
    );
    let bytes = docx_with_links(
        &[
            ("rJs", "javascript:alert(document.domain)"),
            ("rBreak", "https://e.example/x) [y](https://evil.example/"),
            ("rOk", "https://example.com/a?b=c"),
            ("rEntity", "javascript&amp;colon;alert(1)"),
        ],
        &body,
    );
    let markdown = to_markdown_native(&bytes).expect("markdown projects");
    assert!(!markdown.contains("javascript:"), "{markdown}");
    assert!(markdown.contains("script"), "link text is kept: {markdown}");
    assert!(
        markdown.contains("[safe](https://e.example/x%29%20[y]%28https://evil.example/)"),
        "{markdown}"
    );
    assert!(
        markdown.contains("[ok](https://example.com/a?b=c)"),
        "{markdown}"
    );
    assert!(markdown.contains(r"\<img src=x"), "{markdown}");
    // A preceding `!` run cannot turn the projected link into an image.
    assert!(
        markdown.contains(r"\![pic](https://example.com/a?b=c)"),
        "{markdown}"
    );
    // CommonMark decodes entities in destinations; `&` stays literal.
    assert!(
        markdown.contains("[entity](javascript&amp;colon;alert%281%29)"),
        "{markdown}"
    );
    assert!(
        markdown.contains(r"\!\[p\](https\://t.example/p.png)"),
        "{markdown}"
    );
}

/// Inline document text that must render literally wherever it lands: raw
/// HTML, a script link, scheme / `www.` / email autolinks, strikethrough, an
/// entity and a code span (XML-escaped for embedding).
const INLINE: &str = "&lt;img src=x onerror=alert(1)&gt; [x](javascript:alert(1)) https://evil.example/ www.evil.example a@evil.example ~~del~~ &amp;copy; `c`";
const INLINE_TEXT: &str = "<img src=x onerror=alert(1)> [x](javascript:alert(1)) https://evil.example/ www.evil.example a@evil.example ~~del~~ &copy; `c`";
/// Line-level document text: a code fence, a GFM table with a
/// colon-aligned delimiter row, and line-start block markers, separated by
/// lone CRs (XML character references survive parsing).
const LINES: &str = "alpha&#13;&#13;```&#13;| a | b |&#13;| :-- | --: |&#13;forged | row&#13;# h&#13;&#160;# nbsp&#13;&#xFEFF;# bom&#13;&gt; q&#13;- item&#13;1. item&#13;===&#13;    Cost: 50% &amp; tax&#13;&#9;Tabbed: 2&#13;&#160;&#160;Kept nbsp&#13;&#x3000;Kept ideographic&#13;&#x2003;# Kept em";

/// The Markdown, parsed as GFM (tables, strikethrough, extended autolinks),
/// contains only the structure the converter emits: document text in
/// headings, paragraphs, table cells, fields, equations, notes and comments
/// stays literal text and forms no table, link, autolink, heading,
/// strikethrough, code or HTML node.
#[test]
fn document_text_forms_no_markdown_structure() {
    let p = |runs: &str| format!("<w:p>{runs}</w:p>");
    let r = |text: &str| format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#);
    let body = [
        format!(
            r#"<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr>{}<w:r><w:br/></w:r>{}</w:p>"#,
            r("Topic ###"),
            r(INLINE)
        ),
        p(&r(INLINE)),
        p(&format!("{}<w:r><w:br/></w:r>{}", r("```"), r("&gt; quoted"))),
        p(&[
            r("| a | b |"),
            "<w:r><w:br/></w:r>".to_string(),
            r("| :-- | --: |"),
            "<w:r><w:br/></w:r>".to_string(),
            r("forged | row"),
        ]
        .concat()),
        p(&r("&#160;# forged nbsp")),
        p(&r("&#x2003;# forged em")),
        p(&r("&#xFEFF;# forged bom")),
        p(&format!("{}{}", r("1"), r(". split"))),
        p(&format!(
            "{}<w:r><w:br/></w:r><w:r><w:br/></w:r>{}",
            r("intro"),
            r("    Body cost: 50% &amp; tax")
        )),
        p(&format!(
            "{}<w:r><w:br/></w:r>{}{}",
            r("intro"),
            r("  "),
            r("  Split indent: 1")
        )),
        p(&format!(
            "{}<w:r><w:br/></w:r><w:r><w:tab/></w:r>{}",
            r("intro"),
            r("Body tab: 2")
        )),
        p(&format!(
            "{}<w:r><w:br/></w:r>{}",
            r("First"),
            r("&#160;&#160;Second nbsp")
        )),
        p(&format!(
            "{}<w:r><w:br/></w:r>{}{}",
            r("First"),
            r("&#x3000;"),
            r("# Second ideographic")
        )),
        p(&format!(
            "{}<w:r><w:br/></w:r>{}{}",
            r("First"),
            r("  "),
            r("&#x2003;Second em")
        )),
        p(&format!("{}{}", r("ww"), r("w.split.example"))),
        p(r#"<w:fldSimple w:instr=" PAGE "><w:r><w:t xml:space="preserve">&#160;# forged field</w:t></w:r></w:fldSimple>"#),
        p(r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t xml:space="preserve"># forged math https://evil.example/</m:t></m:r></m:oMath>"#),
        format!(
            "<w:tbl><w:tr><w:tc>{}</w:tc><w:tc>{}</w:tc></w:tr></w:tbl>",
            p(&format!("{}<w:r><w:br/></w:r>{}", r("first"), r("# forged"))),
            p(&r(&format!("a|b {INLINE}"))),
        ),
    ]
    .concat();
    let document =
        format!(r#"<w:document xmlns:w="{W_NS}"><w:body>{body}<w:sectPr/></w:body></w:document>"#);
    let rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="c" Type="{R_NS}/comments" Target="comments.xml"/><Relationship Id="f" Type="{R_NS}/footnotes" Target="footnotes.xml"/></Relationships>"#
    );
    let comments = format!(
        r#"<w:comments xmlns:w="{W_NS}"><w:comment w:id="&lt;b&gt;#" w:author="[a](javascript:alert(1)) a@evil.example"><w:p><w:r><w:t xml:space="preserve">{INLINE}&#13;{LINES}</w:t></w:r></w:p></w:comment></w:comments>"#
    );
    let footnotes = format!(
        r#"<w:footnotes xmlns:w="{W_NS}"><w:footnote w:id="]&lt;img src=x onerror=alert(1)&gt;"><w:p><w:r><w:t xml:space="preserve"># {INLINE}</w:t></w:r></w:p></w:footnote></w:footnotes>"#
    );
    let bytes = package(&[
        ("word/document.xml", document),
        ("word/_rels/document.xml.rels", rels),
        ("word/comments.xml", comments),
        ("word/footnotes.xml", footnotes),
    ]);

    let markdown = to_markdown_native(&bytes).expect("markdown projects");
    let gfm = ooxml_common::markdown::gfm_check::summarize(&markdown);

    assert_eq!(
        gfm.headings,
        [
            format!("Topic ### {INLINE_TEXT}"),
            "Footnotes".to_string(),
            "Review comments".to_string(),
            "Comment <b>#".to_string(),
        ],
        "{markdown}"
    );
    assert_eq!(gfm.tables, 1, "{markdown}");
    assert_eq!(
        gfm.table_cells,
        ["first# forged".to_string(), format!("a|b {INLINE_TEXT}")],
        "{markdown}"
    );
    assert_eq!(gfm.block_quotes, 1, "{markdown}");
    assert_eq!(
        (gfm.links, gfm.images, gfm.html, gfm.strikethrough, gfm.code),
        (0, 0, 0, 0, 0),
        "{markdown}"
    );
    assert_eq!((gfm.list_items, gfm.thematic_breaks), (0, 0), "{markdown}");
    for literal in [
        INLINE_TEXT,
        "```",
        "| a | b |",
        "| :-- | --: |",
        "forged | row",
        "# forged nbsp",
        "# forged em",
        "# forged bom",
        "1. split",
        "www.split.example",
        "# forged field",
        "# forged math https://evil.example/",
        "> quoted",
        "[a](javascript:alert(1)) a@evil.example",
        "- item",
        "1. item",
        "===",
        "Cost: 50% & tax",
        "Tabbed: 2",
        "\u{a0}\u{a0}Kept nbsp",
        "\u{3000}Kept ideographic",
        "\u{2003}# Kept em",
        "Body cost: 50% & tax",
        "Split indent: 1",
        "Body tab: 2",
        "\u{a0}\u{a0}Second nbsp",
        "\u{3000}# Second ideographic",
        "\u{2003}Second em",
    ] {
        assert!(gfm.text.contains(literal), "{literal:?} in {markdown}");
    }
}

/// CommonMark strips a document-initial BOM, so indentation after one at
/// the very start of the document must not survive to open a code block
/// (which would show the escapes literally). Covers a BOM followed by four
/// spaces in one run, the BOM and spaces split across runs, and a tab.
#[test]
fn document_initial_bom_cannot_expose_an_indented_code_block() {
    let r = |text: &str| format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#);
    for runs in [
        r("&#xFEFF;    Cost: 1"),
        [r("&#xFEFF;"), r("  "), r("  Cost: 1")].concat(),
        [
            r("&#xFEFF;"),
            "<w:r><w:tab/></w:r>".to_string(),
            r("Cost: 1"),
        ]
        .concat(),
        r("&#xFEFF;&#9;Cost: 1"),
    ] {
        let document = format!(
            r#"<w:document xmlns:w="{W_NS}"><w:body><w:p>{runs}</w:p><w:sectPr/></w:body></w:document>"#
        );
        let bytes = package(&[("word/document.xml", document)]);

        let markdown = to_markdown_native(&bytes).expect("markdown projects");
        let gfm = ooxml_common::markdown::gfm_check::summarize(&markdown);

        assert_eq!(gfm.code, 0, "{markdown:?}");
        assert_eq!(gfm.text.trim(), "Cost: 1", "{markdown:?}");
    }
}

/// The GFM checker itself sees every construct the regressions rule out, so
/// a zero count above is a real absence.
#[test]
fn gfm_checker_detects_the_constructs_it_rules_out() {
    let gfm = ooxml_common::markdown::gfm_check::summarize(
        "# h\n\n[x](https://e.example) https://e.example www.e.example a@e.example ~~d~~ `c` <b>x</b><br>\n\n> q\n\n- i\n\n| a | b |\n| :-- | --: |\n| c | d |\n",
    );
    assert_eq!(gfm.headings, ["h"]);
    assert_eq!(gfm.links, 4);
    assert_eq!(
        (gfm.html, gfm.strikethrough, gfm.code, gfm.block_quotes),
        (2, 1, 1, 1)
    );
    assert_eq!(
        (gfm.list_items, gfm.tables, gfm.table_cells.len()),
        (1, 1, 4)
    );
}
