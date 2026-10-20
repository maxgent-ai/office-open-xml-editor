//! Parts a chart part reaches through its own relationship part
//! (`…/charts/_rels/chartN.xml.rels`), shared by the DOCX, PPTX and XLSX
//! package adapters.
//!
//! The adapters own their archive and their XML preflight limits, so the
//! loaders here take a `read_part` callback (part name → UTF-8 text, `None`
//! when the part is absent or unreadable) and, for userShapes, an already
//! parsed chart document.
//!
//! Every target is resolved with [`crate::rels::RelTarget::resolve_part`]
//! (ECMA-376 Part 2 §6.5.2.3 / §9.3): a relationship whose target names no
//! package part, including any `TargetMode="External"` relationship, is never
//! read as a part.

use crate::rels::{parse_rels, relationship_part_path, RelTarget, TargetMode};

use super::{
    is_chart_style_relationship_type, ChartImageRelationships, ChartImageSource,
    CHART_COLOR_STYLE_REL_TYPE_SUFFIX,
};

/// Text the loaders store for a chartStyle / chartColorStyle relationship
/// whose part is missing or names no package part. It is distinct from
/// "no relationship" (`None`): the chart parser then treats the sidecar as
/// present but empty instead of falling back to the host's implicit style.
const MISSING_SIDECAR_PART: &str = "\0";

/// The theme override, chartStyle, chartColorStyle and image relationships a
/// chart part declares. `style_xml` / `color_style_xml` are `None` when the
/// chart has no such relationship and `Some("\0")` when the relationship
/// exists but its part cannot be read.
#[derive(Debug, Default)]
pub struct ChartRelatedParts {
    pub theme_override_xml: Option<String>,
    pub style_xml: Option<String>,
    pub color_style_xml: Option<String>,
    pub image_relationships: ChartImageRelationships,
}

/// Load the related parts of the chart part at `chart_path`
/// (e.g. `word/charts/chart6.xml`). The chartStyle relationship accepts the
/// Office 2011 and MS-ODRAWXML 2012 types; the chartStyle part's own
/// relationships contribute its image relationships.
pub fn load_chart_related_parts(
    mut read_part: impl FnMut(&str) -> Option<String>,
    chart_path: &str,
) -> ChartRelatedParts {
    let mut result = ChartRelatedParts::default();
    let Some(rels_xml) = read_part(&relationship_part_path(chart_path)) else {
        return result;
    };
    let relationships = parse_rels(&rels_xml);
    result.image_relationships.insert_parsed_relationships(
        ChartImageSource::Chart,
        chart_path,
        &relationships,
    );
    let internal_target = |accepts: &dyn Fn(&str) -> bool| -> Option<&RelTarget> {
        relationships.values().find(|relationship| {
            relationship.mode == TargetMode::Internal
                && relationship
                    .relationship_type
                    .as_deref()
                    .is_some_and(accepts)
        })
    };
    if let Some(relationship) = internal_target(&|kind| kind.ends_with("/themeOverride")) {
        if let Some(path) = relationship.resolve_part(chart_path) {
            result.theme_override_xml = read_part(&path);
        }
    }
    if let Some(relationship) = internal_target(&is_chart_style_relationship_type) {
        // A target naming no part reads like a missing chartStyle part, and
        // has no relationship part of its own to index.
        let style_path = relationship.resolve_part(chart_path);
        result.style_xml = Some(
            style_path
                .as_deref()
                .and_then(&mut read_part)
                .unwrap_or_else(|| MISSING_SIDECAR_PART.to_owned()),
        );
        if let Some(style_path) = style_path.as_deref() {
            if let Some(style_rels_xml) = read_part(&relationship_part_path(style_path)) {
                result.image_relationships.insert_parsed_relationships(
                    ChartImageSource::Style,
                    style_path,
                    &parse_rels(&style_rels_xml),
                );
            }
        }
    }
    if let Some(relationship) =
        internal_target(&|kind| kind.ends_with(CHART_COLOR_STYLE_REL_TYPE_SUFFIX))
    {
        result.color_style_xml = Some(
            relationship
                .resolve_part(chart_path)
                .as_deref()
                .and_then(&mut read_part)
                .unwrap_or_else(|| MISSING_SIDECAR_PART.to_owned()),
        );
    }
    result
}

/// Follow a legacy chart's first `<c:userShapes r:id>` relationship to its
/// Chart Drawing part (`cdr:CT_Drawing`, ECMA-376 Part 1 §21.2.2.211 and
/// dml-chartDrawing.xsd). `chart_doc` is the chart part at `chart_path`,
/// parsed under the caller's preflight limits.
pub fn load_chart_user_shapes_xml(
    mut read_part: impl FnMut(&str) -> Option<String>,
    chart_path: &str,
    chart_doc: &roxmltree::Document<'_>,
) -> Option<String> {
    let rid = chart_doc
        .root_element()
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "userShapes")?
        .attributes()
        .find(|attribute| attribute.name() == "id" && crate::ns::is_r_ns(attribute.namespace()))?
        .value();
    let rels_xml = read_part(&relationship_part_path(chart_path))?;
    let user_shapes_path = parse_rels(&rels_xml).get(rid)?.resolve_part(chart_path)?;
    read_part(&user_shapes_path)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    const RELS_OPEN: &str =
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#;

    fn reader(parts: &[(&str, &str)]) -> impl FnMut(&str) -> Option<String> {
        let parts: HashMap<String, String> = parts
            .iter()
            .map(|(name, text)| ((*name).to_owned(), (*text).to_owned()))
            .collect();
        move |name| parts.get(name).cloned()
    }

    #[test]
    fn missing_sidecar_parts_are_present_but_unreadable() {
        let rels = format!(
            r#"{RELS_OPEN}<Relationship Id="rStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="missing-style.xml"/><Relationship Id="rColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="missing-colors.xml"/></Relationships>"#
        );
        let related = load_chart_related_parts(
            reader(&[("word/charts/_rels/chart9.xml.rels", &rels)]),
            "word/charts/chart9.xml",
        );
        assert_eq!(related.style_xml.as_deref(), Some("\0"));
        assert_eq!(related.color_style_xml.as_deref(), Some("\0"));
        assert_eq!(related.theme_override_xml, None);
    }

    #[test]
    fn style_parts_resolve_root_nested_and_root_absolute_targets() {
        for (chart_path, style_target, expected_style_path, expected_image_path) in [
            ("chart1.xml", "style1.xml", "style1.xml", "media/marker.png"),
            (
                "chart2.xml",
                "/style2.xml",
                "style2.xml",
                "media/marker.png",
            ),
            (
                "xl/charts/chart3.xml",
                "style3.xml",
                "xl/charts/style3.xml",
                "xl/media/marker.png",
            ),
        ] {
            let image_target = if chart_path.contains('/') {
                "../media/marker.png"
            } else {
                "media/marker.png"
            };
            let chart_rels = format!(
                r#"{RELS_OPEN}<Relationship Id="aExternal" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="https://example.invalid/style.xml" TargetMode="External"/><Relationship Id="rStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="{style_target}"/></Relationships>"#
            );
            let style_rels = format!(
                r#"{RELS_OPEN}<Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{image_target}"/></Relationships>"#
            );
            let related = load_chart_related_parts(
                reader(&[
                    (&relationship_part_path(chart_path), &chart_rels),
                    (expected_style_path, "<style/>"),
                    (&relationship_part_path(expected_style_path), &style_rels),
                ]),
                chart_path,
            );
            assert_eq!(related.style_xml.as_deref(), Some("<style/>"));
            assert_eq!(related.color_style_xml, None);
            assert_eq!(
                super::super::ChartImageResolver::resolve_image(
                    &related.image_relationships,
                    ChartImageSource::Style,
                    "rImage",
                ),
                Some((expected_image_path.to_string(), "image/png".to_string())),
            );
        }
    }

    /// An internal chartStyle target that names no package part (here a URI
    /// with a scheme, ECMA-376 Part 2 §9.3) is read as a missing part; its
    /// relationship part is not looked up, so the package-root `_rels/.rels`
    /// never contributes chartStyle image relationships.
    #[test]
    fn unresolvable_style_target_indexes_no_style_images() {
        let chart_rels = format!(
            r#"{RELS_OPEN}<Relationship Id="rStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="https://example.invalid/style.xml"/></Relationships>"#
        );
        let root_rels = format!(
            r#"{RELS_OPEN}<Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/marker.png"/></Relationships>"#
        );
        let mut requested = Vec::new();
        let mut read = reader(&[
            ("xl/charts/_rels/chart1.xml.rels", &chart_rels),
            ("_rels/.rels", &root_rels),
            ("", "<style/>"),
        ]);
        let related = load_chart_related_parts(
            |name: &str| {
                requested.push(name.to_owned());
                read(name)
            },
            "xl/charts/chart1.xml",
        );
        assert_eq!(related.style_xml.as_deref(), Some("\0"));
        assert_eq!(
            super::super::ChartImageResolver::resolve_image(
                &related.image_relationships,
                ChartImageSource::Style,
                "rImage",
            ),
            None,
        );
        assert_eq!(requested, ["xl/charts/_rels/chart1.xml.rels"]);
    }

    /// `TargetMode="External"` names a resource outside the package
    /// (ECMA-376 Part 2 §9.3), so even a target spelled like a sibling part
    /// is not read as the Chart Drawing part.
    #[test]
    fn user_shapes_ignore_external_and_follow_internal_targets() {
        let chart_xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:userShapes r:id="rShapes"/></c:chartSpace>"#;
        let chart_doc = roxmltree::Document::parse(chart_xml).unwrap();
        for (mode, expected) in [
            ("", Some("<drawing/>")),
            (r#" TargetMode="External""#, None),
        ] {
            let rels = format!(
                r#"{RELS_OPEN}<Relationship Id="rShapes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" Target="../drawings/drawing1.xml"{mode}/></Relationships>"#
            );
            let shapes = load_chart_user_shapes_xml(
                reader(&[
                    ("ppt/charts/_rels/chart1.xml.rels", &rels),
                    ("ppt/drawings/drawing1.xml", "<drawing/>"),
                ]),
                "ppt/charts/chart1.xml",
                &chart_doc,
            );
            assert_eq!(shapes.as_deref(), expected, "TargetMode attribute {mode:?}");
        }
    }
}
