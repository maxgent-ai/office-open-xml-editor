//! DrawingML `CT_LineProperties` parsing shared by all three OOXML hosts.
//!
//! ECMA-376 §20.1.2.2.24 defines the same line grammar for PresentationML,
//! SpreadsheetDrawingML, and WordprocessingShape. The descriptor deliberately
//! preserves `Option` presence: inheritance needs to distinguish an absent
//! property from an explicit `<a:noFill>` or `<a:tailEnd type="none">`.

use crate::color::{parse_color_node, ThemeResolver, TintMode};
use crate::fill::{parse_grad_fill, parse_patt_fill, GradientFill, PatternFill};
use crate::ns::is_a_ns;
use roxmltree::Node;

/// Availability ceiling for one DrawingML custom dash list. The schema leaves
/// `<a:custDash><a:ds>` unbounded, while every consumer expands the stops into
/// a Canvas dash array. This bounds input work without changing authored
/// cadence for ordinary documents.
pub const MAX_LINE_DASH_STOPS: usize = 512;

#[derive(Debug, Clone, PartialEq)]
pub enum LinePaint {
    NoFill,
    Solid {
        color: Option<String>,
    },
    /// `None` retains an authored gradient whose stops could not be resolved.
    Gradient(Option<GradientFill>),
    Pattern(PatternFill),
}

#[derive(Debug, Clone, PartialEq)]
pub struct LineDashStop {
    /// Raw DrawingML percentage units (100000 = 100%). Strict OOXML decimal
    /// percentage lexical values are normalized to the same scale.
    pub dash: f64,
    /// Raw DrawingML percentage units (100000 = 100%).
    pub space: f64,
}

#[derive(Debug, Clone, PartialEq)]
pub enum LineDash {
    /// The optional `val` is preserved; hosts may apply their compatibility
    /// default only when converting to an effective renderer model.
    Preset(Option<String>),
    Custom(Vec<LineDashStop>),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LineJoin {
    Round,
    Bevel,
    Miter { limit: Option<i64> },
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct LineEnd {
    pub kind: Option<String>,
    pub width: Option<String>,
    pub length: Option<String>,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct LineProperties {
    pub paint: Option<LinePaint>,
    pub width: Option<i64>,
    pub cap: Option<String>,
    pub compound: Option<String>,
    pub alignment: Option<String>,
    pub dash: Option<LineDash>,
    pub join: Option<LineJoin>,
    pub head_end: Option<LineEnd>,
    pub tail_end: Option<LineEnd>,
}

impl LineProperties {
    /// Overlay authored local properties over a style-matrix recipe. Each
    /// property inherits independently, so a local width does not discard the
    /// recipe's paint/dash and a local `noFill` does suppress its paint.
    pub fn with_fallback(&self, fallback: &Self) -> Self {
        Self {
            paint: self.paint.clone().or_else(|| fallback.paint.clone()),
            width: self.width.or(fallback.width),
            cap: self.cap.clone().or_else(|| fallback.cap.clone()),
            compound: self.compound.clone().or_else(|| fallback.compound.clone()),
            alignment: self
                .alignment
                .clone()
                .or_else(|| fallback.alignment.clone()),
            dash: self.dash.clone().or_else(|| fallback.dash.clone()),
            join: self.join.clone().or_else(|| fallback.join.clone()),
            head_end: self.head_end.clone().or_else(|| fallback.head_end.clone()),
            tail_end: self.tail_end.clone().or_else(|| fallback.tail_end.clone()),
        }
    }
}

fn a_child<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Option<Node<'a, 'input>> {
    node.children().find(|child| {
        child.is_element()
            && child.tag_name().name() == name
            && is_a_ns(child.tag_name().namespace())
    })
}

fn parse_end(node: Node<'_, '_>) -> LineEnd {
    LineEnd {
        kind: node.attribute("type").map(str::to_owned),
        width: node.attribute("w").map(str::to_owned),
        length: node.attribute("len").map(str::to_owned),
    }
}

/// DrawingML percentages are integer thousandths of a percent in common
/// Transitional files, while Strict OOXML uses a decimal `%` lexical form.
/// Normalize both to the historical 100000-per-100% scale used by the shared
/// line model.
fn parse_positive_percentage(value: &str) -> Option<f64> {
    let thousandths = crate::units::drawingml_percentage_to_thousandths(value)?;
    (thousandths >= 0.0).then_some(thousandths)
}

/// Parse only the paint choice of one DrawingML `<a:ln>`. Chart Style palette
/// expansion uses this narrower path so a long custom dash recipe is not
/// reparsed once per palette color.
pub fn parse_line_paint(
    line: Node<'_, '_>,
    resolver: &dyn ThemeResolver,
    tint_mode: TintMode,
) -> Option<LinePaint> {
    line.children().find_map(|child| {
        if !child.is_element() || !is_a_ns(child.tag_name().namespace()) {
            return None;
        }
        match child.tag_name().name() {
            "noFill" => Some(LinePaint::NoFill),
            "solidFill" => Some(LinePaint::Solid {
                color: parse_color_node(child, resolver, tint_mode),
            }),
            "gradFill" => Some(LinePaint::Gradient(parse_grad_fill(
                child, resolver, tint_mode,
            ))),
            "pattFill" => Some(LinePaint::Pattern(parse_patt_fill(
                child, resolver, tint_mode,
            ))),
            _ => None,
        }
    })
}

/// Parse one DrawingML `<a:ln>` into a presence-preserving descriptor.
pub fn parse_line_properties(
    line: Node<'_, '_>,
    resolver: &dyn ThemeResolver,
    tint_mode: TintMode,
) -> LineProperties {
    let paint = parse_line_paint(line, resolver, tint_mode);

    let dash = if let Some(preset) = a_child(line, "prstDash") {
        Some(LineDash::Preset(preset.attribute("val").map(str::to_owned)))
    } else {
        a_child(line, "custDash").map(|custom| {
            LineDash::Custom(
                custom
                    .children()
                    .filter(|node| {
                        node.is_element()
                            && node.tag_name().name() == "ds"
                            && is_a_ns(node.tag_name().namespace())
                    })
                    .take(MAX_LINE_DASH_STOPS)
                    .filter_map(|stop| {
                        Some(LineDashStop {
                            dash: parse_positive_percentage(stop.attribute("d")?)?,
                            space: parse_positive_percentage(stop.attribute("sp")?)?,
                        })
                    })
                    .collect(),
            )
        })
    };

    let join = if a_child(line, "round").is_some() {
        Some(LineJoin::Round)
    } else if a_child(line, "bevel").is_some() {
        Some(LineJoin::Bevel)
    } else {
        a_child(line, "miter").map(|miter| LineJoin::Miter {
            limit: miter.attribute("lim").and_then(|value| value.parse().ok()),
        })
    };

    LineProperties {
        paint,
        width: line.attribute("w").and_then(|value| value.parse().ok()),
        cap: line.attribute("cap").map(str::to_owned),
        compound: line.attribute("cmpd").map(str::to_owned),
        alignment: line.attribute("algn").map(str::to_owned),
        dash,
        join,
        head_end: a_child(line, "headEnd").map(parse_end),
        tail_end: a_child(line, "tailEnd").map(parse_end),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::color::{ThemeResolver, TintMode};
    use roxmltree::Document;

    const NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    struct Resolver;
    impl ThemeResolver for Resolver {
        fn resolve_scheme_color(&self, name: &str) -> Option<String> {
            match name {
                "accent1" => Some("4472C4".to_owned()),
                _ => None,
            }
        }
    }

    #[test]
    fn parses_complete_line_properties_without_filling_absent_values() {
        let xml = format!(
            r#"<a:ln xmlns:a="{NS}" w="25400" cap="rnd" cmpd="dbl" algn="in">
              <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
              <a:custDash><a:ds d="125000" sp="75000"/></a:custDash>
              <a:miter lim="800000"/>
              <a:headEnd type="triangle" w="lg" len="sm"/>
              <a:tailEnd type="none"/>
            </a:ln>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let line = parse_line_properties(doc.root_element(), &Resolver, TintMode::WordLiteral);

        assert_eq!(line.width, Some(25400));
        assert_eq!(line.cap.as_deref(), Some("rnd"));
        assert_eq!(line.compound.as_deref(), Some("dbl"));
        assert_eq!(line.alignment.as_deref(), Some("in"));
        assert_eq!(
            line.paint,
            Some(LinePaint::Solid {
                color: Some("4472C4".into())
            })
        );
        assert_eq!(
            line.dash,
            Some(LineDash::Custom(vec![LineDashStop {
                dash: 125000.0,
                space: 75000.0
            }]))
        );
        assert_eq!(
            line.join,
            Some(LineJoin::Miter {
                limit: Some(800000)
            })
        );
        assert_eq!(
            line.head_end,
            Some(LineEnd {
                kind: Some("triangle".into()),
                width: Some("lg".into()),
                length: Some("sm".into()),
            })
        );
        assert_eq!(
            line.tail_end.as_ref().and_then(|end| end.kind.as_deref()),
            Some("none")
        );
    }

    #[test]
    fn explicit_no_fill_and_local_width_override_style_independently() {
        let style_xml = format!(
            r#"<a:ln xmlns:a="{NS}" w="12700" cap="sq"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:prstDash val="dash"/></a:ln>"#
        );
        let local_xml = format!(r#"<a:ln xmlns:a="{NS}" w="38100"><a:noFill/></a:ln>"#);
        let style_doc = Document::parse(&style_xml).unwrap();
        let local_doc = Document::parse(&local_xml).unwrap();
        let style =
            parse_line_properties(style_doc.root_element(), &Resolver, TintMode::WordLiteral);
        let local =
            parse_line_properties(local_doc.root_element(), &Resolver, TintMode::WordLiteral);
        let effective = local.with_fallback(&style);

        assert_eq!(effective.width, Some(38100));
        assert_eq!(effective.paint, Some(LinePaint::NoFill));
        assert_eq!(effective.cap.as_deref(), Some("sq"));
        assert_eq!(effective.dash, Some(LineDash::Preset(Some("dash".into()))));
    }

    #[test]
    fn fixed_style_color_resolves_without_a_placeholder_reference_color() {
        let xml = format!(
            r#"<a:ln xmlns:a="{NS}"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let style_resolver = crate::color::StyleMatrixColorResolver::new(&Resolver, None);
        let line = parse_line_properties(
            doc.root_element(),
            &style_resolver,
            TintMode::PowerPointLinear,
        );
        assert_eq!(
            line.paint,
            Some(LinePaint::Solid {
                color: Some("4472C4".into())
            })
        );
    }

    #[test]
    fn accepts_strict_namespace_but_ignores_same_named_foreign_children() {
        let xml = r#"<a:ln xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:x="urn:not-drawingml" w="9525">
          <x:noFill/><a:solidFill><a:srgbClr val="010203"/></a:solidFill>
        </a:ln>"#;
        let doc = Document::parse(xml).unwrap();
        let line = parse_line_properties(doc.root_element(), &Resolver, TintMode::WordLiteral);
        assert_eq!(
            line.paint,
            Some(LinePaint::Solid {
                color: Some("010203".into())
            })
        );
    }

    #[test]
    fn normalizes_strict_decimal_percent_custom_dash_values() {
        let xml = r#"<a:ln xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">
          <a:custDash><a:ds d="125.5%" sp="75.25%"/></a:custDash>
        </a:ln>"#;
        let doc = Document::parse(xml).unwrap();
        let line = parse_line_properties(doc.root_element(), &Resolver, TintMode::WordLiteral);
        assert_eq!(
            line.dash,
            Some(LineDash::Custom(vec![LineDashStop {
                dash: 125_500.0,
                space: 75_250.0,
            }]))
        );
    }

    #[test]
    fn rejects_custom_dash_percentages_outside_the_schema_lexical_space() {
        let xml = r#"<a:ln xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">
          <a:custDash>
            <a:ds d="1e2%" sp="75%"/>
            <a:ds d="+1%" sp="75%"/>
            <a:ds d="1.5" sp="75"/>
            <a:ds d="1e3" sp="75"/>
            <a:ds d="2147483648" sp="75"/>
            <a:ds d="100000" sp="75000"/>
          </a:custDash>
        </a:ln>"#;
        let doc = Document::parse(xml).unwrap();
        let line = parse_line_properties(doc.root_element(), &Resolver, TintMode::WordLiteral);
        assert_eq!(
            line.dash,
            Some(LineDash::Custom(vec![LineDashStop {
                dash: 100_000.0,
                space: 75_000.0,
            }]))
        );
    }

    #[test]
    fn retains_zero_length_dashes_allowed_by_positive_percentage() {
        let xml = r#"<a:ln xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">
          <a:custDash><a:ds d="0%" sp="75%"/></a:custDash>
        </a:ln>"#;
        let doc = Document::parse(xml).unwrap();
        let line = parse_line_properties(doc.root_element(), &Resolver, TintMode::WordLiteral);
        assert_eq!(
            line.dash,
            Some(LineDash::Custom(vec![LineDashStop {
                dash: 0.0,
                space: 75_000.0,
            }]))
        );
    }

    #[test]
    fn bounds_custom_dash_stops_before_building_the_wire_list() {
        let stops = (0..MAX_LINE_DASH_STOPS + 1)
            .map(|_| r#"<a:ds d="100000" sp="100000"/>"#)
            .collect::<String>();
        let xml = format!(r#"<a:ln xmlns:a="{NS}"><a:custDash>{stops}</a:custDash></a:ln>"#);
        let doc = Document::parse(&xml).unwrap();
        let line = parse_line_properties(doc.root_element(), &Resolver, TintMode::WordLiteral);
        let Some(LineDash::Custom(stops)) = line.dash else {
            panic!("expected bounded custom dash list");
        };
        assert_eq!(stops.len(), MAX_LINE_DASH_STOPS);
    }
}
