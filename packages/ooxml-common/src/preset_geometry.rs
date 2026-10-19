//! Shared `<a:prstGeom>` helpers (ECMA-376 Part 1 §20.1.9.18
//! `CT_PresetGeometry2D`).

use roxmltree::Node;

/// Most adjust handles any preset shape declares (`adj1`..`adj8`,
/// presetShapeDefinitions.xml).
pub const MAX_PRESET_ADJUST_HANDLES: usize = 8;

/// Parse the adjust handles from a `<a:prstGeom>`'s `<a:avLst>` into an
/// `adj1..adj8`-ordered vector for the shared preset engine.
///
/// ECMA-376 Part 1 §20.1.9.5 (`avLst`) / §20.1.9.11 (`gd`): each
/// `<a:gd name="adjN" fmla="val X"/>` supplies one handle. When any guide
/// carries a `name`, matching is strictly by name (`adj` or `adj1` for the
/// first handle), so a missing `adjN` leaves a `None` gap and later indices
/// stay aligned. Only a legacy list with no names at all is read by position.
/// A `fmla` that is not `val <number>` reads as `None`. Trailing `None`s are
/// trimmed, so a plain rect/ellipse yields an empty vector.
pub fn parse_preset_adj(prst_geom: Node) -> Vec<Option<f64>> {
    let Some(av) = prst_geom
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "avLst")
    else {
        return Vec::new();
    };
    let gd_nodes: Vec<Node> = av
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "gd")
        .collect();
    if gd_nodes.is_empty() {
        return Vec::new();
    }

    let parse_val = |gd: &Node| -> Option<f64> {
        gd.attribute("fmla")
            .and_then(|f| f.strip_prefix("val "))
            .and_then(|s| s.trim().parse::<f64>().ok())
    };

    let named = gd_nodes.iter().any(|n| n.attribute("name").is_some());
    let mut out: Vec<Option<f64>> = (1..=MAX_PRESET_ADJUST_HANDLES)
        .map(|i| {
            if named {
                let key = format!("adj{i}");
                gd_nodes
                    .iter()
                    .find(|n| {
                        let name = n.attribute("name");
                        if i == 1 {
                            matches!(name, Some("adj") | Some("adj1"))
                        } else {
                            name == Some(key.as_str())
                        }
                    })
                    .and_then(parse_val)
            } else {
                gd_nodes.get(i - 1).and_then(parse_val)
            }
        })
        .collect();

    while matches!(out.last(), Some(None)) {
        out.pop();
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn adj(av_lst: &str) -> Vec<Option<f64>> {
        let xml = format!(
            r#"<a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="wedgeRectCallout">{av_lst}</a:prstGeom>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        parse_preset_adj(doc.root_element())
    }

    #[test]
    fn absent_or_empty_av_lst_yields_no_handles() {
        assert!(adj("").is_empty());
        assert!(adj("<a:avLst/>").is_empty());
    }

    #[test]
    fn named_guides_match_by_name_and_keep_gaps() {
        assert_eq!(
            adj(r#"<a:avLst><a:gd name="adj" fmla="val 41667"/></a:avLst>"#),
            vec![Some(41667.0)]
        );
        // adj2 is missing: its slot stays None instead of taking adj3's value.
        assert_eq!(
            adj(
                r#"<a:avLst><a:gd name="adj3" fmla="val 55000"/><a:gd name="adj1" fmla="val -20000"/><a:gd name="adj9" fmla="val 1"/></a:avLst>"#
            ),
            vec![Some(-20000.0), None, Some(55000.0)]
        );
    }

    #[test]
    fn unnamed_guides_fall_back_to_position() {
        assert_eq!(
            adj(r#"<a:avLst><a:gd fmla="val 16667"/><a:gd fmla="val  5 "/></a:avLst>"#),
            vec![Some(16667.0), Some(5.0)]
        );
    }
}
