//! DOCX resource compatibility policy, applied AFTER ordinary MCE selection.
//! ECMA-376 Part 3 §9.3 still owns Requires and branch ordering. This narrow
//! library policy substitutes only an authored picture fallback for a selected
//! single-drawing ChartEx payload that the package cannot render. It is not a
//! general MCE rule or an inference about Office's choice-selection algorithm.
//! Office's 2015 ChartEx capability namespace is recognized only while testing
//! `Requires` on a Choice that has this exact raw ChartEx drawing shape. It is
//! not added to the document's general MCE application configuration. The
//! resource verdict still rejects an unrenderable ChartEx part and substitutes
//! its authored fallback.
use ooxml_common::ns::{is_a_ns, is_c_ns, is_w_ns, is_wp_ns};
use ooxml_common::{bounded_xml::MCE_NS, mce::ChoiceRequiresClassification};
use std::collections::HashSet;

// [MS-ODRAWXML] §§2.1.5, 2.24.1.1 and 2.24.3.76 identify the
// ChartEx part family, chart element and its relationship-id attribute.
pub(crate) const CHARTEX_NS: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
pub(crate) const CHARTEX_CAPABILITY_NS: &str =
    "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
pub(crate) const CHARTEX_REL: &str =
    "http://schemas.microsoft.com/office/2014/relationships/chartEx";

pub(crate) struct PathLevel(fn(Option<&str>) -> bool, &'static [&'static str]);
impl PathLevel {
    pub(crate) fn matches(&self, namespace: Option<&str>, local: &str) -> bool {
        (self.0)(namespace) && self.1.contains(&local)
    }
}
fn chart_namespace(ns: Option<&str>) -> bool {
    is_c_ns(ns) || ns == Some(CHARTEX_NS)
}
// Only raw DIRECT children count; first matching child wins at each level.
// In particular, extension/ignorable payloads and nested ACs are never searched.
pub(crate) const PATH: [PathLevel; 5] = [
    PathLevel(is_w_ns, &["drawing"]),
    PathLevel(is_wp_ns, &["inline", "anchor"]),
    PathLevel(is_a_ns, &["graphic"]),
    PathLevel(is_a_ns, &["graphicData"]),
    PathLevel(chart_namespace, &["chart"]),
];
#[derive(Default)]
pub(crate) struct Facts {
    pub(crate) direct_children: u8,
    pub(crate) drawing: bool,
    pub(crate) chartex_uri: bool,
    pub(crate) renderable_rid: bool,
}
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum Verdict {
    Parent,
    Renderable,
    Unrenderable,
}
pub(crate) fn verdict(facts: &Facts) -> Verdict {
    if facts.direct_children != 1 || !facts.drawing || !facts.chartex_uri {
        Verdict::Parent
    } else if facts.renderable_rid {
        Verdict::Renderable
    } else {
        Verdict::Unrenderable
    }
}

pub(crate) fn native_verdict(choice: roxmltree::Node, rids: &HashSet<String>) -> Verdict {
    let mut facts = Facts::default();
    let mut elements = choice.children().filter(|node| node.is_element());
    let Some(mut node) = elements.next() else {
        return Verdict::Parent;
    };
    facts.direct_children = if elements.next().is_some() { 2 } else { 1 };
    facts.drawing = PATH[0].matches(node.tag_name().namespace(), node.tag_name().name());
    if !facts.drawing || facts.direct_children != 1 {
        return Verdict::Parent;
    }
    for (index, level) in PATH.iter().enumerate().skip(1) {
        let Some(child) = node.children().find(|child| {
            child.is_element()
                && level.matches(child.tag_name().namespace(), child.tag_name().name())
        }) else {
            break;
        };
        node = child;
        if index == 3 {
            facts.chartex_uri = node.attribute("uri") == Some(CHARTEX_NS);
        }
        if index == 4 {
            facts.renderable_rid = ooxml_common::ns::attr_ns(
                &node,
                ooxml_common::ns::relationships::TRANSITIONAL,
                ooxml_common::ns::relationships::STRICT,
                "id",
            )
            .is_some_and(|rid| rids.contains(rid));
        }
    }
    verdict(&facts)
}

/// Select with the parent's application configuration, except that the 2015
/// capability token is accepted for one Choice only after its raw direct-child
/// shape has been classified as the narrow ChartEx drawing path above.
pub(crate) fn select_native_alternate_content<'a, 'i>(
    alternate: roxmltree::Node<'a, 'i>,
    rids: &HashSet<String>,
    understood: &dyn Fn(&str) -> bool,
    mce_understood: &dyn Fn(&str) -> bool,
) -> Option<roxmltree::Node<'a, 'i>> {
    for choice in alternate.children().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MCE_NS)
            && node.tag_name().name() == "Choice"
    }) {
        let parent_classification =
            ooxml_common::mce::classify_choice_requires(choice.attribute("Requires"), |prefix| {
                match choice.lookup_namespace_uri(Some(prefix)) {
                    None => ooxml_common::mce::RequiredNamespaceSupport::Unresolved,
                    Some(namespace) if understood(namespace) => {
                        ooxml_common::mce::RequiredNamespaceSupport::Understood
                    }
                    Some(_) => ooxml_common::mce::RequiredNamespaceSupport::Unsupported,
                }
            });
        if parent_classification == ChoiceRequiresClassification::Understood {
            return Some(choice);
        }
        if parent_classification != ChoiceRequiresClassification::Unsupported
            || native_verdict(choice, rids) == Verdict::Parent
        {
            continue;
        }
        let local_classification =
            ooxml_common::mce::classify_choice_requires(choice.attribute("Requires"), |prefix| {
                match choice.lookup_namespace_uri(Some(prefix)) {
                    None => ooxml_common::mce::RequiredNamespaceSupport::Unresolved,
                    Some(namespace)
                        if understood(namespace) || namespace == CHARTEX_CAPABILITY_NS =>
                    {
                        ooxml_common::mce::RequiredNamespaceSupport::Understood
                    }
                    Some(_) => ooxml_common::mce::RequiredNamespaceSupport::Unsupported,
                }
            });
        if local_classification == ChoiceRequiresClassification::Understood
            && native_branch_must_understand(choice, mce_understood)
        {
            return Some(choice);
        }
    }
    alternate.children().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MCE_NS)
            && node.tag_name().name() == "Fallback"
    })
}

/// A resource override must not newly select a Fallback that the MCE processor
/// would reject. In that case retaining the parent's selected Choice preserves
/// native/streaming parity and the parent's fail-closed drawing result.
pub(crate) fn native_branch_must_understand(
    branch: roxmltree::Node,
    understood: &dyn Fn(&str) -> bool,
) -> bool {
    branch
        .attribute((MCE_NS, "MustUnderstand"))
        .is_none_or(|prefixes| {
            prefixes.split_whitespace().all(|prefix| {
                branch
                    .lookup_namespace_uri(Some(prefix))
                    .is_some_and(understood)
            })
        })
}
