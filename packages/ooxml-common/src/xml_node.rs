//! Small `roxmltree` lookups shared by the OOXML parsers.
//!
//! Element lookups match by local name only; attribute lookups match only
//! unqualified (no-namespace) attributes, which is how DrawingML and the
//! relationship/chart grammars declare their attributes.

use roxmltree::Node;

/// First child element of `node` whose local name is `local`.
pub fn child<'a, 'i>(node: Node<'a, 'i>, local: &str) -> Option<Node<'a, 'i>> {
    node.children()
        .find(|n| n.is_element() && n.tag_name().name() == local)
}

/// Unqualified attribute `local` of `node` as an owned string.
pub fn attr(node: &Node<'_, '_>, local: &str) -> Option<String> {
    node.attribute(local).map(str::to_owned)
}

/// Unqualified attribute `local` parsed as `i64`.
pub fn attr_i64(node: &Node<'_, '_>, local: &str) -> Option<i64> {
    node.attribute(local)?.parse().ok()
}

/// Unqualified attribute `local` parsed as `f64`. Rust's float grammar also
/// accepts `inf` / `NaN`; callers that feed geometry arithmetic should use
/// [`attr_finite_f64`] instead.
pub fn attr_f64(node: &Node<'_, '_>, local: &str) -> Option<f64> {
    node.attribute(local)?.parse().ok()
}

/// [`attr_f64`] restricted to finite values: `inf` / `NaN` read as absent.
pub fn attr_finite_f64(node: &Node<'_, '_>, local: &str) -> Option<f64> {
    attr_f64(node, local).filter(|value| value.is_finite())
}

/// ECMA-376 Part 1 §22.9.2.7 `ST_OnOff` lexical value: `1`/`true`/`on` and
/// `0`/`false`/`off`, matched exactly. Any other text is invalid and yields
/// `None` so the caller applies its own default.
pub fn parse_st_on_off(value: &str) -> Option<bool> {
    match value {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn attribute_lookups_ignore_qualified_names_and_finite_reader_rejects_inf() {
        let doc = roxmltree::Document::parse(
            r#"<a xmlns:x="urn:x" x:w="1" h="2.5" big="inf" n="-3"><b/><c/><b id="2"/></a>"#,
        )
        .unwrap();
        let root = doc.root_element();
        assert_eq!(attr(&root, "w"), None);
        assert_eq!(attr_f64(&root, "h"), Some(2.5));
        assert_eq!(attr_f64(&root, "big"), Some(f64::INFINITY));
        assert_eq!(attr_finite_f64(&root, "big"), None);
        assert_eq!(attr_i64(&root, "n"), Some(-3));
        assert_eq!(child(root, "b").and_then(|b| attr(&b, "id")), None);
    }

    #[test]
    fn st_on_off_accepts_exact_lexicals_only() {
        assert_eq!(parse_st_on_off("on"), Some(true));
        assert_eq!(parse_st_on_off("0"), Some(false));
        assert_eq!(parse_st_on_off("True"), None);
        assert_eq!(parse_st_on_off(" 1"), None);
    }
}
