//! Parse one standalone PresentationML shape part into the presentation model.
//!
//! A shape part is an OPC part whose root element is a `p:sp` (ECMA-376
//! 19.3.1.43) or `p:cxnSp` (19.3.1.19) outside any slide. Theme references
//! resolve against the supplied theme part (20.1.6.9) and color map
//! (19.3.1.6). There is no layout or master: placeholder inheritance is not
//! resolved, so the result reports whether the shape is a placeholder and the
//! caller decides whether a partially inherited shape is usable.
use super::*;
use crate::theme::{apply_clr_map, parse_clr_map_node, PptxTheme};
use ooxml_common::ns::is_a_ns;
use ooxml_common::theme::StyleMatrixLookup;

/// A parsed standalone shape and the facts a caller needs to judge whether it
/// is complete without a slide context.
pub struct StandaloneShape {
    pub element: ShapeElement,
    /// The shape carries `p:nvPr/p:ph`; layout/master inheritance is absent.
    pub placeholder: bool,
    /// The shape or a selected theme style references a relationship, or a
    /// selected theme style could not be inspected. Theme relationships belong
    /// to the theme part, which this standalone API does not receive; callers
    /// must verify them separately.
    pub relationship_references: bool,
}

/// Parse the shape part `part` of the OPC `package`. Returns `Ok(None)` when
/// the root element is not a supported shape. `max_part_bytes` bounds each
/// inflated archive entry and `max_total_bytes` the archive as a whole.
pub fn parse_standalone_shape_part(
    package: &[u8],
    part: &str,
    theme_xml: &str,
    clr_map_xml: Option<&str>,
    max_part_bytes: u64,
    max_total_bytes: u64,
) -> Result<Option<StandaloneShape>, String> {
    let mut zip = open_zip_with_limits(
        package.to_vec(),
        Some(max_part_bytes),
        Some(max_total_bytes),
    )?;
    zip.run_operation("standalone-shape", |zip| {
        let xml = read_zip_str(zip, part).map_err(|e| e.to_string())?;
        let doc = parse_preflighted_pptx_xml(&xml).map_err(|e| e.to_string())?;
        let root = doc.root_element();
        let mut theme = PptxTheme::from_xml(theme_xml);
        if let Some(map_xml) = clr_map_xml {
            let map_doc = parse_preflighted_pptx_xml(map_xml).map_err(|e| e.to_string())?;
            let node = map_doc.root_element();
            // Slide masters carry p:clrMap (ECMA-376 §19.3.1.6), while
            // [MS-PPT] §2.11.9 round-trip color mappings use a:clrMap.
            if node.tag_name().name() != "clrMap"
                || !(is_p_ns(node.tag_name().namespace()) || is_a_ns(node.tag_name().namespace()))
            {
                return Err("color map part has no clrMap root".to_owned());
            }
            apply_clr_map(&mut theme, Some(&parse_clr_map_node(node)));
        }
        let rels_xml = read_zip_str(zip, &relationship_part_path(part)).unwrap_or_default();
        let rels = parse_rels(&rels_xml);
        // ECMA-376 §20.1.4.2: style-matrix references select DrawingML
        // fragments in the theme. A blip there owns a relationship in the
        // theme part, even when the shape XML contains no r: attribute. This
        // API has no theme part path or its relationships, so report that
        // dependency rather than presenting a fallback fill as complete.
        let relationship_references = root
            .descendants()
            .any(|n| n.attributes().any(|a| is_r_ns(a.namespace())))
            || theme_style_has_relationship(root, &theme);
        let placeholder = is_placeholder(root);
        let element = match (root.tag_name().name(), is_p_ns(root.tag_name().namespace())) {
            ("sp", true) => parse_shape(
                root,
                &LayoutPlaceholders::default(),
                &theme,
                &rels,
                part,
                None,
                zip,
            )
            .map(|mut shape| {
                // The slide-tree path promotes image-filled p:sp nodes to
                // PictureElement. This standalone API must retain ShapeElement,
                // so resolve the same blip relationship into its fill instead.
                // Relationship targets are relative to the source shape part,
                // not to its containing directory (ECMA-376 Part 2 §6.4.1).
                if let Some(blip_fill) =
                    child(root, "spPr").and_then(|node| child(node, "blipFill"))
                {
                    let mut resolve_blip = |relationship_id: &str| {
                        let target = rels.get(relationship_id)?;
                        let image_path = resolve_path(part, target);
                        zip.index_for_name(&image_path)?;
                        Some(image_path)
                    };
                    shape.fill = parse_blip_fill(blip_fill, theme.colors(), &mut resolve_blip);
                }
                shape
            }),
            ("cxnSp", true) => parse_connector(root, &theme, &rels),
            _ => None,
        };
        Ok(element.map(|element| StandaloneShape {
            element,
            placeholder,
            relationship_references,
        }))
    })
}

fn theme_style_has_relationship(root: roxmltree::Node<'_, '_>, theme: &PptxTheme) -> bool {
    // ECMA-376 §20.1.4.2: only the DrawingML references directly in the
    // shape's p:style select theme matrix entries. Extension descendants do
    // not select styles, even if their local names match a reference.
    let mut inspected = std::collections::HashMap::new();
    root.children()
        .find(|node| {
            node.is_element()
                && node.tag_name().name() == "style"
                && is_p_ns(node.tag_name().namespace())
        })
        .into_iter()
        .flat_map(|style| style.children())
        .any(|node| {
            if !node.is_element() || !is_a_ns(node.tag_name().namespace()) {
                return false;
            }
            let Some(index) = node.attribute("idx").and_then(|value| value.parse().ok()) else {
                return false;
            };
            let (kind, selected) = match node.tag_name().name() {
                "fillRef" => (0, theme.format_scheme.lookup_fill_ref(index)),
                "lnRef" => (1, theme.format_scheme.lookup_line_ref(index)),
                "effectRef" => (2, theme.format_scheme.lookup_effect_ref(index)),
                _ => return false,
            };
            if let Some(&has_relationship) = inspected.get(&(kind, index)) {
                return has_relationship;
            }
            let StyleMatrixLookup::Entry(entry) = selected else {
                return false;
            };
            // A failed fragment parse leaves the dependency unverifiable. Never
            // turn that state into a claim that this shape is self-contained.
            let has_relationship = theme_fragment_needs_relationship_verification(&entry.to_xml());
            inspected.insert((kind, index), has_relationship);
            has_relationship
        })
}

fn theme_fragment_needs_relationship_verification(xml: &str) -> bool {
    roxmltree::Document::parse(xml).map_or(true, |doc| {
        doc.descendants().any(|node| {
            node.attributes()
                .any(|attribute| is_r_ns(attribute.namespace()))
        })
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const STRICT_P: &str = "http://purl.oclc.org/ooxml/presentationml/main";
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn package(xml: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        zip.start_file("shape.xml", SimpleFileOptions::default())
            .unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
        zip.finish().unwrap().into_inner()
    }

    fn package_with_relationship(
        part: &str,
        xml: &str,
        relationship_target: &str,
        related_part: &str,
    ) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        zip.start_file(part, SimpleFileOptions::default()).unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
        zip.start_file(relationship_part_path(part), SimpleFileOptions::default())
            .unwrap();
        write!(
            zip,
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{relationship_target}"/></Relationships>"#
        )
        .unwrap();
        zip.start_file(related_part, SimpleFileOptions::default())
            .unwrap();
        zip.write_all(b"image").unwrap();
        zip.finish().unwrap().into_inner()
    }

    fn parse(xml: &str, theme: &str, map: Option<&str>) -> Result<Option<StandaloneShape>, String> {
        parse_standalone_shape_part(&package(xml), "shape.xml", theme, map, 100_000, 1_000_000)
    }

    fn shape(namespace: Option<&str>, kind: &str, extra: &str) -> String {
        let namespace = namespace.map_or(String::new(), |ns| format!(" xmlns:p=\"{ns}\""));
        let prefix = if namespace.is_empty() { "" } else { "p:" };
        format!(
            "<{prefix}{kind}{namespace} xmlns:a=\"{A}\" xmlns:r=\"{R}\"><{prefix}spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm></{prefix}spPr>{extra}</{prefix}{kind}>"
        )
    }

    #[test]
    fn only_presentationml_shape_roots_are_supported() {
        for namespace in [None, Some("urn:foreign")] {
            assert!(parse(&shape(namespace, "sp", ""), "", None)
                .unwrap()
                .is_none());
        }
        for namespace in [P, STRICT_P] {
            assert!(parse(&shape(Some(namespace), "sp", ""), "", None)
                .unwrap()
                .is_some());
            assert!(parse(&shape(Some(namespace), "cxnSp", ""), "", None)
                .unwrap()
                .is_some());
            assert!(parse(&shape(Some(namespace), "pic", ""), "", None)
                .unwrap()
                .is_none());
        }
    }

    #[test]
    fn standalone_shape_keeps_relationship_backed_image_fill() {
        use std::io::Write;
        let xml = format!(
            "<p:sp xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\"><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed=\"rId1\"/><a:srcRect l=\"35000\"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>"
        );
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        for (name, bytes) in [
            ("shape.xml", xml.as_bytes()),
            ("_rels/shape.xml.rels", b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image.png\"/></Relationships>".as_slice()),
            ("media/image.png", b"image".as_slice()),
        ] {
            zip.start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            zip.write_all(bytes).unwrap();
        }
        let package = zip.finish().unwrap().into_inner();
        let parsed =
            parse_standalone_shape_part(&package, "shape.xml", "", None, 100_000, 1_000_000)
                .unwrap()
                .unwrap();
        match parsed.element.fill {
            Some(Fill::Image {
                image_path,
                src_rect,
                ..
            }) => {
                assert_eq!(image_path, "media/image.png");
                assert_eq!(src_rect.unwrap().l, 0.35);
            }
            other => panic!("image fill lost: {other:?}"),
        }
    }

    #[test]
    fn color_map_requires_presentationml_or_drawingml_namespace() {
        let xml = shape(Some(P), "sp", "");
        for map in ["<clrMap/>", "<x:clrMap xmlns:x=\"urn:foreign\"/>"] {
            assert!(parse(&xml, "", Some(map)).is_err());
        }
        for namespace in [P, STRICT_P] {
            let map = format!("<p:clrMap xmlns:p=\"{namespace}\"/>");
            assert!(parse(&xml, "", Some(&map)).unwrap().is_some());
        }
        for namespace in [A, STRICT_A] {
            let map = format!("<a:clrMap xmlns:a=\"{namespace}\"/>");
            assert!(parse(&xml, "", Some(&map)).unwrap().is_some());
        }
    }

    #[test]
    fn selected_theme_image_relationship_is_reported() {
        let theme = format!("<a:theme xmlns:a=\"{A}\" xmlns:r=\"{R}\"><a:themeElements><a:fmtScheme name=\"x\"><a:fillStyleLst><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill><a:blipFill><a:blip r:embed=\"rIdImage\"/></a:blipFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>");
        for (index, expected) in [(1, false), (2, true)] {
            let xml = shape(
                Some(P),
                "sp",
                &format!("<p:style><a:fillRef idx=\"{index}\"/></p:style>"),
            );
            let parsed = parse(&xml, &theme, None).unwrap().unwrap();
            assert_eq!(parsed.relationship_references, expected);
        }
    }

    #[test]
    fn nested_standalone_shape_resolves_image_fill_from_source_part() {
        let part = "ppt/shapes/shape1.xml";
        let xml = format!(
            "<p:sp xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\"><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm><a:blipFill><a:blip r:embed=\"rIdImage\"/><a:stretch/></a:blipFill></p:spPr></p:sp>"
        );
        let package =
            package_with_relationship(part, &xml, "../media/image.png", "ppt/media/image.png");

        let parsed = parse_standalone_shape_part(&package, part, "", None, 100_000, 1_000_000)
            .unwrap()
            .unwrap();

        assert!(matches!(
            parsed.element.fill,
            Some(Fill::Image { ref image_path, .. }) if image_path == "ppt/media/image.png"
        ));
    }

    #[test]
    fn unicode_relationship_prefix_in_selected_theme_fill_is_reported() {
        let theme = format!("<a:theme xmlns:a=\"{A}\" xmlns:関係=\"{R}\"><a:themeElements><a:fmtScheme name=\"x\"><a:fillStyleLst><a:blipFill><a:blip 関係:embed=\"rIdImage\"/></a:blipFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>");
        let xml = shape(Some(P), "sp", "<p:style><a:fillRef idx=\"1\"/></p:style>");
        assert!(
            parse(&xml, &theme, None)
                .unwrap()
                .unwrap()
                .relationship_references
        );
    }

    #[test]
    fn unused_extension_style_reference_does_not_select_theme_fill() {
        let theme = format!("<a:theme xmlns:a=\"{A}\" xmlns:r=\"{R}\"><a:themeElements><a:fmtScheme name=\"x\"><a:fillStyleLst><a:blipFill><a:blip r:embed=\"rIdImage\"/></a:blipFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>");
        let xml = shape(
            Some(P),
            "sp",
            "<p:extLst><x:fillRef xmlns:x=\"urn:foreign\" idx=\"1\"/></p:extLst>",
        );
        assert!(
            !parse(&xml, &theme, None)
                .unwrap()
                .unwrap()
                .relationship_references
        );
        let xml = shape(
            Some(P),
            "sp",
            "<p:style><x:fillRef xmlns:x=\"urn:foreign\" idx=\"1\"/></p:style>",
        );
        assert!(
            !parse(&xml, &theme, None)
                .unwrap()
                .unwrap()
                .relationship_references
        );
    }

    #[test]
    fn unparseable_theme_fragment_remains_unverifiable() {
        assert!(theme_fragment_needs_relationship_verification(
            "<a:blip r:embed=\"rIdImage\"/>"
        ));
        assert!(!theme_fragment_needs_relationship_verification("<fill/>"));
    }
}
