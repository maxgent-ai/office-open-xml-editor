//! VML group coordinate adapter (ECMA-376 Part 4 §19.1.2.7/§19.1.2.19).
//! Group children are expressed in the parent's coordsize/coordorigin space,
//! not page points. The shared drawing transform maps their boxes and composes
//! clockwise rotations/reflections without Word-specific DrawingML heuristics.
use super::*;

const VML: &str = "urn:schemas-microsoft-com:vml";

fn unsupported(diagnostics: &mut Vec<PendingParseDiagnostic>) {
    push_pending_parse_diagnostic(
        diagnostics,
        PARSE_DIAGNOSTIC_CODE_UNSUPPORTED_VML_CONTENT,
        PARSE_DIAGNOSTIC_SEVERITY_UNSUPPORTED_VML_CONTENT,
    );
}

fn pair(raw: &str) -> Option<(f64, f64)> {
    let mut values = raw
        .split(|c: char| c == ',' || c.is_whitespace())
        .filter(|s| !s.is_empty());
    let x = values.next()?.parse::<f64>().ok()?;
    let y = values.next()?.parse::<f64>().ok()?;
    (x.is_finite() && y.is_finite() && values.next().is_none()).then_some((x, y))
}

// Visibility is explicitly excluded from shapetype inheritance (§19.1.2.19).
// Only the containing group supplies the inherited state, via traversal.
fn hidden(node: roxmltree::Node) -> bool {
    vml_css_str(node.attribute("style").unwrap_or(""), "visibility") == Some("hidden")
}

fn group_spec(group: roxmltree::Node, root: bool) -> Option<DrawingGroupSpec> {
    let style = group.attribute("style").unwrap_or("");
    let coordinates = VmlStyleCoordinates::for_node(group);
    let (cx, cy) = pair(group.attribute("coordsize").unwrap_or("1000,1000"))?;
    let (ox, oy) = pair(group.attribute("coordorigin").unwrap_or("0,0"))?;
    let width = coordinates.css_length(style, "width")?;
    let height = coordinates.css_length(style, "height")?;
    let flip = vml_css_str(style, "flip").unwrap_or_default();
    // Part 4 §§19.1.2.7/.19: combine nested-group numeric offsets in the
    // parent's coordinate space, using the same authority as leaf/page anchors.
    // The root's placement is owned separately by its retained story wrapper.
    let spec = DrawingGroupSpec {
        off_x: if root {
            0.0
        } else {
            coordinates.offset(style, "left", "margin-left")
        },
        off_y: if root {
            0.0
        } else {
            coordinates.offset(style, "top", "margin-top")
        },
        ext_x: width,
        ext_y: height,
        child_off_x: ox,
        child_off_y: oy,
        child_ext_x: cx,
        child_ext_y: cy,
        rotation_degrees: if root {
            0.0
        } else {
            vml_css_length_pt(style, "rotation").unwrap_or(0.0)
        },
        flip_h: flip.contains('x'),
        flip_v: flip.contains('y'),
    };
    (width.is_finite()
        && height.is_finite()
        && width > 0.0
        && height > 0.0
        && cx > 0.0
        && cy > 0.0
        && spec.off_x.is_finite()
        && spec.off_y.is_finite()
        && spec.rotation_degrees.is_finite())
    .then_some(spec)
}

const IDENTITY_EFFECT: [f64; 6] = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];

// CT_Group includes o:skew in EG_ShapeElements. Part 4 §19.2.2.31
// applies it to descendant vectors after this group's rotation, before its
// ancestors' effects. Conjugate from the group's parent coordinates into the
// root story's point space; never alter the image-data/brush transform.
fn group_effect(
    group: roxmltree::Node,
    spec: DrawingGroupSpec,
    parent: GroupTransform,
    outer: [f64; 6],
    diagnostics: &mut Vec<PendingParseDiagnostic>,
) -> [f64; 6] {
    if hidden(group) {
        return outer;
    }
    if vml_direct_child(group, "skew").is_some_and(|node| {
        node.attributes().any(|a| {
            !matches!(
                a.name(),
                "id" | "ext" | "on" | "matrix" | "offset" | "origin"
            )
        })
    }) {
        unsupported(diagnostics);
    }
    let Some(skew) = vml_geometry::skew(group, None) else {
        if vml_decoration_attr(group, None, "skew", "on").and_then(parse_vml_true_false)
            == Some(true)
        {
            unsupported(diagnostics);
        }
        return outer;
    };
    let (ex, ey) = parent.map_point(0.0, 0.0);
    let (ax, ay) = parent.map_point(1.0, 0.0);
    let (bx, by) = parent.map_point(0.0, 1.0);
    let (a, b, c, d) = (ax - ex, ay - ey, bx - ex, by - ey);
    let det = a * d - b * c;
    let ox = spec.off_x + spec.ext_x * (0.5 + skew.origin.0);
    let oy = spec.off_y + spec.ext_y * (0.5 + skew.origin.1);
    // Absolute offsets are physical points; fractional offsets follow the
    // group extent. Remove parent scaling before mapping them to story points.
    let dx = skew.offset.x_pt / parent.scale_x + skew.offset.x_fraction * spec.ext_x;
    let dy = skew.offset.y_pt / parent.scale_y + skew.offset.y_fraction * spec.ext_y;
    let project = |x: f64, y: f64| {
        let u = (d * (x - ex) - c * (y - ey)) / det - ox;
        let v = (-b * (x - ex) + a * (y - ey)) / det - oy;
        let px = ox + skew.matrix[0] * u + skew.matrix[2] * v + dx;
        let py = oy + skew.matrix[1] * u + skew.matrix[3] * v + dy;
        let (px, py) = parent.map_point(px, py);
        (
            outer[0] * px + outer[2] * py + outer[4],
            outer[1] * px + outer[3] * py + outer[5],
        )
    };
    let (x, y) = project(0.0, 0.0);
    let (ax, ay) = project(1.0, 0.0);
    let (bx, by) = project(0.0, 1.0);
    [ax - x, ay - y, bx - x, by - y, x, y]
}

// The root group's host transform belongs to the retained story wrapper.
// Conjugate its own post-rotation skew into that wrapper's pre-host frame;
// nested group effects already precede this outer host transform.
fn root_effect(group: roxmltree::Node, spec: DrawingGroupSpec, effect: [f64; 6]) -> [f64; 6] {
    if effect == IDENTITY_EFFECT {
        return effect;
    }
    let style = group.attribute("style").unwrap_or("");
    let radians = vml_css_length_pt(style, "rotation")
        .unwrap_or(0.0)
        .to_radians();
    let (sin, cos) = radians.sin_cos();
    // Unlike text in a leaf textbox, a group's descendants are vectors:
    // group_spec already baked root reflection into their coordinate map.
    // The retained story wrapper owns rotation only.
    let (fx, fy) = (1.0, 1.0);
    let (cx, cy) = (spec.ext_x / 2.0, spec.ext_y / 2.0);
    let project = |x: f64, y: f64| {
        let u = cx + cos * fx * (x - cx) - sin * fy * (y - cy);
        let v = cy + sin * fx * (x - cx) + cos * fy * (y - cy);
        let u2 = effect[0] * u + effect[2] * v + effect[4] - cx;
        let v2 = effect[1] * u + effect[3] * v + effect[5] - cy;
        (
            cx + fx * (cos * u2 + sin * v2),
            cy + fy * (-sin * u2 + cos * v2),
        )
    };
    let (x, y) = project(0.0, 0.0);
    let (ax, ay) = project(1.0, 0.0);
    let (bx, by) = project(0.0, 1.0);
    [ax - x, ay - y, bx - x, by - y, x, y]
}

#[allow(clippy::too_many_arguments)]
pub(super) fn parse_groups(
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    pict: roxmltree::Node,
    theme: &ThemeColors,
    media_map: &HashMap<String, String>,
    chart_map: &ChartMap,
    rel_map: &HashMap<String, String>,
    depth: DepthGuard,
    document_diagnostics: &mut Vec<PendingParseDiagnostic>,
) -> Option<Vec<DocRun>> {
    if !pict.children().any(|n| n.has_tag_name((VML, "group"))) {
        return None;
    }
    with_vml_wrap_polygon_memo(pict, || {
        parse_present_groups(
            style_map,
            num_map,
            pict,
            theme,
            media_map,
            chart_map,
            rel_map,
            depth,
            document_diagnostics,
        )
    })
}

#[allow(clippy::too_many_arguments)]
fn parse_present_groups(
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    pict: roxmltree::Node,
    theme: &ThemeColors,
    media_map: &HashMap<String, String>,
    chart_map: &ChartMap,
    rel_map: &HashMap<String, String>,
    depth: DepthGuard,
    document_diagnostics: &mut Vec<PendingParseDiagnostic>,
) -> Option<Vec<DocRun>> {
    // Resolve definitions once per pict, not once per leaf in large groups.
    let shape_types: HashMap<_, _> = pict
        .descendants()
        .filter(|n| n.has_tag_name((VML, "shapetype")))
        .filter_map(|n| n.attribute("id").map(|id| (id, n)))
        .collect();
    let mut result = Vec::new();
    let mut sibling_diagnostics = Vec::new();
    for group in pict.children().filter(|n| n.is_element()) {
        if group.has_tag_name((VML, "shapetype")) {
            continue;
        }
        if !group.has_tag_name((VML, "group")) {
            // CT_Group and CT_Pict use the same drawable leaves. A sibling's
            // coordinates are already points, so acquire with the identity map.
            if let Some(run) = acquire_leaf(
                style_map,
                num_map,
                group,
                &shape_types,
                GroupTransform::IDENTITY,
                false,
                result.len() as u32,
                theme,
                media_map,
                chart_map,
                rel_map,
                depth,
                &mut sibling_diagnostics,
            ) {
                result.push(run);
            }
            continue;
        }
        let Some(spec) = group_spec(group, true) else {
            // Part 4 §19.1.2.7 permits auto/relative sizing, but acquisition
            // currently needs a concrete point-space box. No wrapper survives
            // this gate, so preserve the omission as a document diagnostic in
            // every story rather than relying on retained unsupported markers.
            unsupported(document_diagnostics);
            continue;
        };
        let mut group_diagnostics = Vec::new();
        let diagnostics = &mut group_diagnostics;
        let mut children = Vec::new();
        let effect = group_effect(
            group,
            spec,
            GroupTransform::IDENTITY,
            IDENTITY_EFFECT,
            diagnostics,
        );
        let effect = root_effect(group, spec, effect);
        let mut pending = vec![(group, GroupTransform::from_group(spec), effect)];
        while let Some((parent, transform, effect)) = pending.pop() {
            // A stack frame per group avoids recursive parser calls. Reverse
            // insertion preserves authored painter order across nested groups.
            if parent.has_tag_name((VML, "group")) {
                // Part 4 §§19.1.2.7/.19: hidden is inherited from the
                // containing group; every other value means inherit, even
                // "visible". Keep the root wrapper's authored flow box below,
                // but acquire no descendant ink/resources for a hidden group.
                if hidden(parent) {
                    continue;
                }
                if vml_geometry::incomplete_host(parent, None) {
                    unsupported(diagnostics);
                }
                let mut ordered_children: Vec<_> =
                    parent.children().filter(|n| n.is_element()).collect();
                // §19.1.2.19: signed z-index controls overlap. Stable sorting
                // keeps source order for equal/default indices, and a nested
                // group remains one stacking unit before its children expand.
                ordered_children.sort_by_key(|child| {
                    let style = child.attribute("style").unwrap_or("");
                    match vml_css_str(style, "z-index").unwrap_or("0").parse::<i32>() {
                        Ok(index) => index,
                        Err(_) => {
                            unsupported(diagnostics);
                            0
                        }
                    }
                });
                for child in ordered_children.into_iter().rev() {
                    if child.has_tag_name((VML, "group")) {
                        if let Some(spec) = group_spec(child, false) {
                            let child_effect =
                                group_effect(child, spec, transform, effect, diagnostics);
                            pending.push((child, transform.compose_group(spec), child_effect));
                        } else {
                            unsupported(diagnostics);
                        }
                    } else if child.tag_name().namespace()
                        == Some("urn:schemas-microsoft-com:office:office")
                        && child.tag_name().name() == "skew"
                    {
                        // A group property has been consumed by group_effect.
                    } else {
                        pending.push((child, transform, effect));
                    }
                }
                // CT_Group (§19.1.2.7) does not declare shape-level opacity;
                // do not invent an extra group alpha multiplier. Leaf/template
                // fill alpha is resolved by vml_fill. Property children that
                // have no standalone geometry are disclosed by acquire_leaf.
                // The group itself is only a coordinate container.
                continue;
            }
            if let Some(mut run) = acquire_leaf(
                style_map,
                num_map,
                parent,
                &shape_types,
                transform,
                true,
                children.len() as u32,
                theme,
                media_map,
                chart_map,
                rel_map,
                depth,
                diagnostics,
            ) {
                if effect != IDENTITY_EFFECT {
                    if let DocRun::Shape(shape) = &mut run {
                        let (x, y) = (shape.anchor_x_pt, shape.anchor_y_pt);
                        shape.vml_group_transform = Some([
                            effect[0],
                            effect[1],
                            effect[2],
                            effect[3],
                            effect[0] * x + effect[2] * y + effect[4] - x,
                            effect[1] * x + effect[3] * y + effect[5] - y,
                        ]);
                    }
                }
                children.push(run);
            }
        }
        let style = group.attribute("style").unwrap_or("");
        let absolute = vml_css_str(style, "position") == Some("absolute");
        let mut wrapper = ShapeRun {
            inline: !absolute,
            width_pt: spec.ext_x,
            height_pt: spec.ext_y,
            rotation: vml_css_length_pt(style, "rotation").unwrap_or(0.0),
            flip_h: vml_css_str(style, "flip").is_some_and(|s| s.contains('x')),
            flip_v: vml_css_str(style, "flip").is_some_and(|s| s.contains('y')),
            text_box_content: vec![TextBoxBlockWire::Body(BodyElement::Paragraph(Box::new(
                DocParagraph {
                    alignment: "left".into(),
                    runs: children,
                    ..Default::default()
                },
            )))],
            ..Default::default()
        };
        // The group is represented as a retained text-box story, just like an
        // inline DrawingML group. Carry incomplete-content markers in that
        // story, so header/footer/note groups also reach DocumentLayout's
        // existing unsupported-block diagnostic collector (their paragraph
        // parsers do not publish body-only PendingParseDiagnostics).
        if !group_diagnostics.is_empty() {
            wrapper
                .text_box_content
                .push(TextBoxBlockWire::Unsupported {
                    kind: "unsupportedTextBoxBlock".into(),
                    q_name: "v:group".into(),
                    source_path: vec![1],
                });
        }
        if absolute {
            wrapper.anchor_acquisition = Some(vml_word_anchor_acquisition(
                group, None, style, spec.ext_x, spec.ext_y,
            ));
        }
        result.push(DocRun::Shape(Box::new(wrapper)));
    }
    if !sibling_diagnostics.is_empty() {
        // Paragraph parsers in headers/footers do not publish body parse
        // diagnostics. Bind sibling omissions to an existing retained group
        // story as well, so every story reports the same incomplete pict.
        if let Some(group) = result.iter_mut().find_map(|run| match run {
            DocRun::Shape(shape) if !shape.text_box_content.is_empty() => Some(shape),
            _ => None,
        }) {
            group.text_box_content.push(TextBoxBlockWire::Unsupported {
                kind: "unsupportedTextBoxBlock".into(),
                q_name: "v:pict".into(),
                source_path: vec![group.text_box_content.len()],
            });
        }
        document_diagnostics.extend(sibling_diagnostics);
    }
    Some(result)
}

// One acquisition boundary for CT_Group members and pict siblings (Part 4
// §19.1.2). All leaves retain the same geometry, resource and decoration facts.
#[allow(clippy::too_many_arguments)]
fn acquire_leaf<'a, 'input>(
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    child: roxmltree::Node<'a, 'input>,
    shape_types: &HashMap<&str, roxmltree::Node<'a, 'input>>,
    transform: GroupTransform,
    within_group: bool,
    z_order: u32,
    theme: &ThemeColors,
    media_map: &HashMap<String, String>,
    chart_map: &ChartMap,
    rel_map: &HashMap<String, String>,
    depth: DepthGuard,
    diagnostics: &mut Vec<PendingParseDiagnostic>,
) -> Option<DocRun> {
    if child.tag_name().namespace() != Some(VML) {
        // Lock/wrap/anchorlock and editor metadata have no independent
        // ink. Other Office extensions (diagram, skew, extrusion, etc.)
        // are not implemented and must be disclosed.
        if !matches!(
            child.tag_name().name(),
            "lock" | "wrap" | "anchorlock" | "ClientData"
        ) {
            unsupported(diagnostics);
        }
        return None;
    }
    match child.tag_name().name() {
        "shapetype" => return None,
        "shape" | "rect" | "roundrect" | "oval" | "line" | "image" | "arc" | "curve"
        | "polyline" => {}
        // Property elements have no standalone leaf geometry; invalid
        // drawable hosts and unavailable extensions remain explicit.
        _ => {
            unsupported(diagnostics);
            return None;
        }
    }
    let local_type = child
        .attribute("type")
        .and_then(|id| shape_types.get(id.trim_start_matches('#')))
        .copied();
    // Like an ungrouped pict shape, a definition Word wrote in an earlier
    // pict of the same story still applies (Part 4 §19.1.2.20).
    with_recorded_vml_shape_type(local_type, child, |shape_type| {
        acquire_typed_leaf(
            style_map,
            num_map,
            child,
            shape_type,
            transform,
            within_group,
            z_order,
            theme,
            media_map,
            chart_map,
            rel_map,
            depth,
            diagnostics,
        )
    })
}

#[allow(clippy::too_many_arguments)]
fn acquire_typed_leaf(
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    child: roxmltree::Node,
    shape_type: Option<roxmltree::Node>,
    transform: GroupTransform,
    within_group: bool,
    z_order: u32,
    theme: &ThemeColors,
    media_map: &HashMap<String, String>,
    chart_map: &ChartMap,
    rel_map: &HashMap<String, String>,
    depth: DepthGuard,
    diagnostics: &mut Vec<PendingParseDiagnostic>,
) -> Option<DocRun> {
    let style = child.attribute("style").unwrap_or("");
    let hidden = hidden(child);
    if !hidden && vml_geometry::incomplete_leaf(child, shape_type) {
        unsupported(diagnostics);
    }
    // Device/EMU formula values describe the final extent, whereas
    // path coordinates still belong to the shape's local coordsize.
    let unit = transform.apply_rect(DrawingRect {
        x: 0.0,
        y: 0.0,
        width: 1.0,
        height: 1.0,
        rotation_degrees: 0.0,
        flip_h: false,
        flip_v: false,
    });
    let physical_scale = (unit.width, unit.height);
    let Some(mut shape) = parse_vml_shape(
        style_map,
        num_map,
        child,
        shape_type,
        theme,
        media_map,
        chart_map,
        rel_map,
        physical_scale,
        depth,
    ) else {
        unsupported(diagnostics);
        return None;
    };
    if hidden {
        // A hidden leaf still owns its extent (in particular an inline pict
        // sibling). Retain placement only, so pictures, shadows and text can
        // never escape the visibility gate through another paint layer.
        shape = ShapeRun {
            inline: !within_group
                && vml_image::present(child, shape_type)
                && vml_css_str(style, "position") != Some("absolute"),
            text_box_content: Vec::new(),
            fill: None,
            stroke: None,
            preset_geometry: None,
            subpaths: Vec::new(),
            subpath_paint: Vec::new(),
            vml_shadow: None,
            vml_skew: None,
            vml_edge_strokes: Vec::new(),
            text_path: None,
            text_blocks: Vec::new(),
            ..shape
        };
    } else if vml_image::present(child, shape_type) {
        let Some(picture) = vml_image::acquire(child, shape_type, media_map) else {
            unsupported(diagnostics);
            return None;
        };
        // Part 4 §§19.1.2.10/.11: imagedata has an implied rectangle.
        // Preserve the host fill behind image transparency and its outline
        // above the image. A picture does not use the host's custom path.
        // Preserve the former ImageRun's inline ownership when a standalone
        // picture gains a backing brush; group placement belongs to its wrapper.
        shape.inline = !within_group && !picture.anchor;
        shape.vml_picture = Some(Box::new(picture));
    }
    if within_group {
        shape.anchor_acquisition = None;
        shape.anchor_x_relative_from = Some("column".into());
        shape.anchor_y_relative_from = Some("paragraph".into());
    }
    shape.z_order = z_order;
    let (x, y, width, height, rotation, flip_h, flip_v) = (
        &mut shape.anchor_x_pt,
        &mut shape.anchor_y_pt,
        &mut shape.width_pt,
        &mut shape.height_pt,
        &mut shape.rotation,
        &mut shape.flip_h,
        &mut shape.flip_v,
    );
    // parse_vml_shape already combines numeric offsets (or owns primitive bounds).
    // These are parent-group coordinates under Part 4 §19.1.2.19; preserve
    // that resolved placement and apply the containing transform exactly once.
    let rect = transform.apply_rect(DrawingRect {
        x: *x,
        y: *y,
        width: *width,
        height: *height,
        rotation_degrees: *rotation,
        flip_h: *flip_h,
        flip_v: *flip_v,
    });
    if ![
        rect.x,
        rect.y,
        rect.width,
        rect.height,
        rect.rotation_degrees,
    ]
    .iter()
    .all(|v| v.is_finite())
        || rect.width < 0.0
        || rect.height < 0.0
    {
        unsupported(diagnostics);
        return None;
    }
    *x = rect.x;
    *y = rect.y;
    *width = rect.width;
    *height = rect.height;
    *rotation = rect.rotation_degrees;
    *flip_h = rect.flip_h;
    *flip_v = rect.flip_v;
    Some(DocRun::Shape(Box::new(shape)))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> (ShapeRun, Vec<PendingParseDiagnostic>) {
        let xml = format!(
            r#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{xml}</w:pict>"#
        );
        let doc = XmlDoc::parse(&xml).unwrap();
        let mut diagnostics = Vec::new();
        let mut runs = parse_groups(
            &StyleMap::default(),
            &mut NumberingMap::default(),
            doc.root_element(),
            &ThemeColors::default(),
            &HashMap::from([("pic".to_string(), "word/media/picture.png".to_string())]),
            &ChartMap::default(),
            &HashMap::new(),
            DepthGuard::root(),
            &mut diagnostics,
        )
        .unwrap();
        let DocRun::Shape(shape) = runs.remove(0) else {
            panic!("group wrapper")
        };
        (*shape, diagnostics)
    }

    fn children(shape: &ShapeRun) -> &[DocRun] {
        let TextBoxBlockWire::Body(BodyElement::Paragraph(p)) = &shape.text_box_content[0] else {
            panic!("group story")
        };
        &p.runs
    }

    #[test]
    fn group_preserves_resolved_leaf_margins_in_parent_coordinates() {
        for (position, expected) in [
            (
                "left:0;top:0;margin-left:60pt;margin-top:50pt",
                (120.0, 150.0),
            ),
            ("left:60;top:50;margin-left:0;margin-top:0", (120.0, 150.0)),
            (
                "left:10;top:20;margin-left:60;margin-top:50",
                (140.0, 210.0),
            ),
            ("left:60;top:50", (120.0, 150.0)),
        ] {
            let (group, diagnostics) = parse(&format!(
                r#"<v:group style="width:200pt;height:300pt" coordsize="100,100">
                  <v:rect style="{position};width:10;height:20"/>
                  <v:shape style="{position};width:10;height:20"><v:imagedata r:id="pic"/></v:shape>
                  <v:image style="{position};width:10;height:20" src="media/picture.png"/>
                </v:group>"#
            ));
            assert!(diagnostics.is_empty());
            assert_eq!(children(&group).len(), 3);
            for run in children(&group) {
                let DocRun::Shape(shape) = run else {
                    panic!("shape or picture")
                };
                assert_eq!(
                    (shape.anchor_x_pt, shape.anchor_y_pt),
                    expected,
                    "{position}"
                );
                assert_eq!((shape.width_pt, shape.height_pt), (20.0, 60.0));
            }
        }
    }

    #[test]
    fn nested_group_margins_compose_with_origins_and_leaf_margins() {
        let (group, diagnostics) = parse(
            r#"<v:group style="width:150pt;height:150pt" coordsize="300,300" coordorigin="10,20">
              <v:group style="left:20;top:10;margin-left:100;margin-top:80;width:80;height:40" coordsize="40,20">
                <v:rect style="left:2;top:3;margin-left:10;margin-top:5;width:10;height:5"/>
                <v:shape style="left:2;top:3;margin-left:10;margin-top:5;width:10;height:5"><v:imagedata r:id="pic"/></v:shape>
              </v:group>
            </v:group>"#,
        );
        assert!(diagnostics.is_empty());
        assert_eq!(children(&group).len(), 2);
        for run in children(&group) {
            let DocRun::Shape(shape) = run else {
                panic!("shape or picture")
            };
            // Outer scale 1/2: ((20+100,10+80) + 2*(2+10,3+5) - (10,20))/2.
            assert_eq!((shape.anchor_x_pt, shape.anchor_y_pt), (67.0, 43.0));
            assert_eq!((shape.width_pt, shape.height_pt), (10.0, 5.0));
        }
    }

    #[test]
    fn grouped_style_units_are_parent_coordinates_at_every_depth() {
        for unit in ["", "pt", "px", "in", "cm", "mm", "pc", "em", "ex"] {
            let (group, diagnostics) = parse(&format!(
                r#"<v:group style="width:150pt;height:150pt" coordsize="300,300" coordorigin="10,20">
                  <v:group style="left:20{unit};top:10{unit};margin-left:100{unit};margin-top:80{unit};width:80{unit};height:40{unit}" coordsize="40,20">
                    <v:rect style="left:2{unit};top:3{unit};margin-left:10{unit};margin-top:5{unit};width:10{unit};height:5{unit}"/>
                    <v:shape style="left:2{unit};top:3{unit};margin-left:10{unit};margin-top:5{unit};width:10{unit};height:5{unit}"><v:imagedata r:id="pic"/></v:shape>
                  </v:group>
                </v:group>"#
            ));
            assert!(diagnostics.is_empty(), "{unit}");
            assert_eq!(children(&group).len(), 2, "{unit}");
            for run in children(&group) {
                let DocRun::Shape(shape) = run else {
                    panic!("shape or picture")
                };
                assert_eq!(
                    (
                        shape.anchor_x_pt,
                        shape.anchor_y_pt,
                        shape.width_pt,
                        shape.height_pt
                    ),
                    (67.0, 43.0, 10.0, 5.0),
                    "{unit}"
                );
            }
        }
        // A root group's dimensions still use physical units: 100px = 75pt.
        let (root, diagnostics) = parse(
            r#"<v:group style="width:100px;height:100px" coordsize="100,100"><v:rect style="left:10;top:20;width:30;height:40"/></v:group>"#,
        );
        assert!(diagnostics.is_empty());
        assert_eq!((root.width_pt, root.height_pt), (75.0, 75.0));
        let DocRun::Shape(shape) = &children(&root)[0] else {
            panic!("shape")
        };
        assert_eq!(
            (
                shape.anchor_x_pt,
                shape.anchor_y_pt,
                shape.width_pt,
                shape.height_pt
            ),
            (7.5, 15.0, 22.5, 30.0)
        );
    }

    #[test]
    fn text_rect_formulas_use_the_same_physical_frame_as_group_geometry() {
        let (group, _) = parse(
            r##"<v:group style="width:200pt;height:100pt" coordsize="100,100">
          <v:shapetype id="region" coordsize="1000,1000">
            <v:formulas><v:f eqn="prod emuwidth 1 2540"/></v:formulas>
            <v:path textboxrect="0,0,@0,1000"/>
          </v:shapetype>
          <v:rect type="#region" style="width:100;height:100"/>
        </v:group>"##,
        );
        let runs = children(&group);
        let DocRun::Shape(shape) = &runs[0] else {
            panic!("shape retained")
        };
        assert_eq!(shape.vml_text_rects, vec![[0.0, 0.0, 1.0, 1.0]]);
    }

    #[test]
    fn visibility_retains_flow_extent_without_inheriting_from_shapetype() {
        let (hidden, diagnostics) = parse(
            r##"<v:group style="width:100pt;height:80pt;visibility:hidden"><v:rect style="width:100;height:100"/></v:group>"##,
        );
        assert!(hidden.inline);
        assert_eq!((hidden.width_pt, hidden.height_pt), (100.0, 80.0));
        assert!(children(&hidden).is_empty());
        assert!(diagnostics.is_empty());
        let (group, _) = parse(
            r##"<v:group style="width:100pt;height:80pt"><v:shapetype id="hidden" style="visibility:hidden"/><v:rect type="#hidden" style="width:100;height:100" fillcolor="red"/></v:group>"##,
        );
        assert!(matches!(&children(&group)[0], DocRun::Shape(shape) if shape.fill.is_some()));
    }

    #[test]
    fn fill_opacity_arrow_dimensions_and_textpath_properties_cascade_individually() {
        let (group, _) = parse(
            r##"<v:group style="width:100pt;height:100pt" coordsize="100,100">
          <v:shapetype id="template" opacity=".5"><v:path arrowok="t" textpathok="t"/><v:stroke endarrow="block" endarrowwidth="wide"/><v:textpath string="inherited" on="t" style="font-family:serif;font-weight:bold;font-style:italic"/></v:shapetype>
          <v:shape type="#template" style="width:80;height:80" opacity=".25"><v:stroke endarrowlength="long"/><v:textpath style="font-style:normal"/></v:shape>
          <v:shape type="#template" style="width:80;height:80"><v:fill opacity="0"/></v:shape>
        </v:group>"##,
        );
        let DocRun::Shape(shape) = &children(&group)[0] else {
            panic!("shape")
        };
        assert_eq!(shape.fill_opacity, Some(0.25));
        let arrow = shape.tail_end.as_ref().unwrap();
        assert_eq!((arrow.w.as_str(), arrow.len.as_str()), ("lg", "lg"));
        let text = shape.text_path.as_ref().unwrap();
        assert_eq!(text.string, "inherited");
        assert!(text.bold);
        assert!(!text.italic);
        let DocRun::Shape(shape) = &children(&group)[1] else {
            panic!("shape")
        };
        assert_eq!(shape.fill_opacity, Some(0.0));
        assert_eq!(shape.text_path.as_ref().unwrap().string, "inherited");
    }

    #[test]
    fn effective_picture_outlines_and_stroke_alpha_are_retained() {
        let (group, _) = parse(
            r##"<v:group style="width:100pt;height:100pt">
          <v:shapetype id="border"><v:stroke color="red" weight="4pt" opacity="50%"/></v:shapetype>
          <v:shape style="width:20;height:20" stroked="t" strokecolor="red" strokeweight="4pt"><v:imagedata r:id="pic"/></v:shape>
          <v:shape type="#border" style="width:20;height:20"><v:imagedata r:id="pic"/></v:shape>
          <v:rect style="width:20;height:20" filled="f"><v:stroke color="red" opacity="0"/></v:rect>
        </v:group>"##,
        );
        let runs = children(&group);
        for (index, expected) in [(0, "ff0000"), (1, "ff000080"), (2, "ff000000")] {
            let DocRun::Shape(shape) = &runs[index] else {
                panic!("decoration lost")
            };
            assert_eq!(shape.stroke.as_deref(), Some(expected));
            assert!(!shape
                .text_box_content
                .iter()
                .any(|b| matches!(b, TextBoxBlockWire::Unsupported { .. })));
        }
    }

    #[test]
    fn leaf_property_audit_discloses_unacquired_paint_without_rejecting_metadata() {
        let (group, _) = parse(
            r#"<v:group style="width:100pt;height:100pt">
          <v:rect style="width:20;height:20"><v:stroke opacity=".5" title="editor metadata"/></v:rect>
          <v:rect style="width:20;height:20"><v:stroke src="media/pen.png" filltype="tile"/></v:rect>
          <v:shape style="width:20;height:20" stroked="f"><v:imagedata r:id="pic" gamma="2"/></v:shape>
          <v:rect style="width:20;height:20"><v:fill unavailable-property="1"/></v:rect>
        </v:group>"#,
        );
        for (i, run) in children(&group).iter().enumerate() {
            let DocRun::Shape(shape) = run else {
                panic!("retained leaf")
            };
            let incomplete = shape
                .text_box_content
                .iter()
                .any(|b| matches!(b, TextBoxBlockWire::Unsupported { .. }));
            assert_eq!(incomplete, i >= 2);
        }
    }

    #[test]
    fn property_sweep_discloses_active_unacquired_modes_and_honors_path_gates() {
        for (attributes, property, incomplete) in [
            ("insetpen=\"t\"", "", false),
            ("insetpen=\"t\"", "<v:path insetpenok=\"f\"/>", false),
            (
                "",
                "<v:fill type=\"gradientradial\"/><v:path gradientshapeok=\"t\"/>",
                false,
            ),
            (
                "",
                "<v:textpath string=\"text\" style=\"text-decoration:underline\"/>",
                true,
            ),
            (
                "",
                "<v:textpath string=\"text\" style=\"v-rotate-letters:true\"/>",
                true,
            ),
            (
                "",
                "<v:textpath string=\"text\" style=\"text-decoration:none\"/>",
                false,
            ),
            ("class=\"unresolved-css\"", "", true),
            ("", "<v:path shadowok=\"f\"/><v:shadow on=\"t\"/>", false),
        ] {
            let (group, _) = parse(&format!(
                r#"<v:group style="width:100pt;height:100pt"><v:rect style="width:100;height:100" {attributes}>{property}</v:rect></v:group>"#
            ));
            assert_eq!(
                group
                    .text_box_content
                    .iter()
                    .any(|b| matches!(b, TextBoxBlockWire::Unsupported { .. })),
                incomplete,
                "{attributes} {property}"
            );
        }
    }

    #[test]
    fn hasfill_formula_does_not_depend_on_picture_relationship_acquisition() {
        let (group, _) = parse(
            r#"<v:group style="width:100pt;height:100pt" coordsize="100,100">
          <v:shape style="width:100;height:100" coordsize="100,100" path="m0,0l@0,100e">
            <v:fill type="tile" r:id="pic"/><v:formulas><v:f eqn="val hasfill"/></v:formulas>
          </v:shape></v:group>"#,
        );
        let DocRun::Shape(shape) = &children(&group)[0] else {
            panic!("shape")
        };
        assert!(matches!(
            shape.subpaths[0][1],
            PathCmd::LineTo { x: 0.01, y: 1.0 }
        ));
        assert!(matches!(shape.fill, Some(ShapeFill::Image { .. })));
    }

    #[test]
    fn mixed_pict_siblings_use_the_group_leaf_dispatcher() {
        let xml = r#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><v:group style="width:10pt;height:10pt"/>
          <v:arc style="width:20pt;height:20pt"/><v:curve from="0,0" control1="0,20" control2="20,20" to="20,0"/>
          <v:polyline points="0,0 10,20 20,0"/><v:image src="media/picture.png" style="width:20pt;height:20pt"/>
        </w:pict>"#;
        let doc = XmlDoc::parse(xml).unwrap();
        let mut diagnostics = Vec::new();
        let runs = parse_groups(
            &StyleMap::default(),
            &mut NumberingMap::default(),
            doc.root_element(),
            &ThemeColors::default(),
            &HashMap::new(),
            &ChartMap::default(),
            &HashMap::new(),
            DepthGuard::root(),
            &mut diagnostics,
        )
        .unwrap();
        assert_eq!(runs.len(), 5);
        assert!(
            matches!(&runs[4], DocRun::Shape(picture) if picture.inline && picture.vml_picture.is_some())
        );
        assert!(diagnostics.is_empty());
    }

    #[test]
    fn group_maps_mixed_children_in_painter_order_with_nested_origin_rotation_and_flip() {
        let (shape, diagnostics) = parse(
            r#"<v:group style="width:200pt;height:100pt" coordsize="1000,500" coordorigin="100,50">
          <v:rect style="left:100;top:50;width:100;height:100" fillcolor="red"/>
          <v:group style="left:200;top:150;width:400;height:200;rotation:90;flip:x" coordsize="200,100" coordorigin="10,20">
            <v:shape style="left:10;top:20;width:100;height:50;rotation:30;flip:y"><v:imagedata r:id="pic"/></v:shape>
          </v:group>
          <v:oval style="left:600;top:250;width:100;height:50"/>
        </v:group>"#,
        );
        assert!(diagnostics.is_empty());
        assert!(shape.inline);
        assert_eq!((shape.width_pt, shape.height_pt), (200.0, 100.0));
        let runs = children(&shape);
        assert_eq!(runs.len(), 3);
        let DocRun::Shape(rect) = &runs[0] else {
            panic!("first rect")
        };
        assert_eq!(
            (
                rect.anchor_x_pt,
                rect.anchor_y_pt,
                rect.width_pt,
                rect.height_pt
            ),
            (0.0, 0.0, 20.0, 20.0)
        );
        let DocRun::Shape(image) = &runs[1] else {
            panic!("nested picture")
        };
        // Nested centre (60,45) -> reflected/rotated centre (450,350)
        // in the outer coordinate system -> (70,60)pt; extents 40x20pt.
        assert!((image.anchor_x_pt - 50.0).abs() < 1e-9);
        assert!((image.anchor_y_pt - 50.0).abs() < 1e-9);
        assert_eq!((image.width_pt, image.height_pt), (40.0, 20.0));
        assert_eq!(image.rotation, 60.0);
        assert!(image.flip_h && image.flip_v);
        assert!(
            matches!(&runs[2], DocRun::Shape(oval) if oval.preset_geometry.as_deref() == Some("ellipse"))
        );
    }

    #[test]
    fn nested_group_is_one_signed_z_index_stacking_unit() {
        let (shape, diagnostics) = parse(
            r#"<v:group style="width:100pt;height:100pt">
          <v:rect style="left:100;width:10;height:10;z-index:2"/>
          <v:group style="width:100;height:100;z-index:-1"><v:rect style="width:10;height:10;z-index:100"/></v:group>
          <v:rect style="left:50;width:10;height:10;z-index:1"/>
        </v:group>"#,
        );
        assert!(diagnostics.is_empty());
        let positions: Vec<_> = children(&shape)
            .iter()
            .map(|run| match run {
                DocRun::Shape(shape) => shape.anchor_x_pt,
                _ => panic!("shape"),
            })
            .collect();
        assert_eq!(positions, vec![0.0, 5.0, 10.0]);
    }

    #[test]
    fn group_default_space_and_absolute_host_keep_parent_and_child_placement_separate() {
        let (shape, diagnostics) = parse(
            r#"<v:group style="position:absolute;margin-left:30pt;margin-top:40pt;width:100pt;height:50pt;rotation:15;flip:xy">
          <v:shape style="left:500;top:200;width:100;height:100"><v:imagedata r:id="pic"/></v:shape>
        </v:group>"#,
        );
        assert!(diagnostics.is_empty());
        assert!(!shape.inline);
        assert_eq!(shape.rotation, 15.0);
        assert!(shape.flip_h && shape.flip_v);
        assert!(shape.anchor_acquisition.is_some());
        let DocRun::Shape(image) = &children(&shape)[0] else {
            panic!("picture")
        };
        assert_eq!(
            (
                image.anchor_x_pt,
                image.anchor_y_pt,
                image.width_pt,
                image.height_pt
            ),
            (40.0, 35.0, 10.0, 5.0)
        );
    }

    #[test]
    fn specified_geometry_and_crops_survive_group_acquisition() {
        let (group, _) = parse(
            r##"<v:group style="width:100pt;height:100pt" coordsize="100,100">
          <v:shape style="width:100;height:100" coordsize="100,100" path="m0,0l100,0,0,100xe"/>
          <v:shape style="width:100;height:100" coordsize="100,100" adj="40"><v:formulas><v:f eqn="sum #0 10 0"/></v:formulas><v:path v="m0,0l@0,0,0,100xe"/></v:shape>
          <v:arc style="width:100;height:100" startangle="90" endangle="270"/>
          <v:curve from="0,0" control1="20,100" control2="80,100" to="100,0"/>
          <v:polyline points="0,0 50,100 100,0"/>
          <v:image style="width:100;height:100" src="media/picture.png"/>
          <v:shape style="width:100;height:100"><v:imagedata r:id="pic" cropleft="32768f" croptop="25%"/></v:shape>
        </v:group>"##,
        );
        assert_eq!(children(&group).len(), 7);
        let DocRun::Shape(triangle) = &children(&group)[0] else {
            panic!("triangle")
        };
        assert!(triangle.preset_geometry.is_none());
        assert!(matches!(
            triangle.subpaths[0][1],
            PathCmd::LineTo { x: 1.0, y: 0.0 }
        ));
        let DocRun::Shape(formula) = &children(&group)[1] else {
            panic!("formula")
        };
        assert!(matches!(
            formula.subpaths[0][1],
            PathCmd::LineTo { x: 0.5, y: 0.0 }
        ));
        let DocRun::Shape(crop) = &children(&group)[6] else {
            panic!("crop")
        };
        let crop = crop
            .vml_picture
            .as_ref()
            .unwrap()
            .src_rect
            .as_ref()
            .unwrap();
        assert_eq!((crop.l, crop.t), (0.5, 0.25));
        assert!(!group
            .text_box_content
            .iter()
            .any(|block| matches!(block, TextBoxBlockWire::Unsupported { .. })));
    }

    #[test]
    fn leaves_preserve_inherited_shadow_and_disclose_unknown_decoration() {
        let (group, _) = parse(
            r##"<v:group style="width:100pt;height:100pt" coordsize="100,100">
          <v:shapetype id="decorated"><v:shadow on="t" color="red" offset="10pt,10pt"/></v:shapetype>
          <v:rect type="#decorated" style="width:40;height:40"/>
          <v:rect style="width:40;height:40"><v:unavailable-decoration/></v:rect>
        </v:group>"##,
        );
        let wire = serde_json::to_value(&children(&group)[0]).unwrap();
        assert_eq!(wire["vmlShadow"]["color"], "ff0000");
        assert!(group
            .text_box_content
            .iter()
            .any(|block| matches!(block, TextBoxBlockWire::Unsupported { .. })));
    }

    #[test]
    fn formula_physical_extents_follow_nested_group_scaling() {
        let (group, _) = parse(
            r#"<v:group style="width:20pt;height:10pt" coordsize="100,100">
          <v:shape style="width:100;height:100" coordsize="100,100">
            <v:formulas><v:f eqn="prod emuwidth 1 12700"/><v:f eqn="prod emuheight 1 12700"/></v:formulas>
            <v:path v="m0,0l@0,@1e"/>
          </v:shape>
        </v:group>"#,
        );
        let DocRun::Shape(shape) = &children(&group)[0] else {
            panic!("shape")
        };
        assert!(matches!(
            shape.subpaths[0][1],
            PathCmd::LineTo { x: 0.2, y: 0.1 }
        ));
    }

    #[test]
    fn metadata_only_shadow_override_keeps_inherited_off_state() {
        let (group, _) = parse(
            r##"<v:group style="width:100pt;height:100pt"><v:shapetype id="off"><v:shadow on="f"/></v:shapetype><v:rect type="#off" style="width:100;height:100"><v:shadow id="metadata"/></v:rect></v:group>"##,
        );
        let DocRun::Shape(shape) = &children(&group)[0] else {
            panic!("shape")
        };
        assert!(shape.vml_shadow.is_none());
        assert!(!group
            .text_box_content
            .iter()
            .any(|b| matches!(b, TextBoxBlockWire::Unsupported { .. })));
    }

    #[test]
    fn group_discloses_every_unrepresented_child_kind_and_unresolved_picture() {
        // CT_Group's complete drawable choice, including its Office extension.
        for child in [
            "arc",
            "image",
            "diagram",
            "shadow",
            "extrusion",
            "skew",
            "clippath",
            "textbox",
            "textpath",
            "imagedata",
        ] {
            let (shape, diagnostics) = parse(&format!(
                r#"<v:group style="width:100pt;height:50pt"><v:{child}/></v:group>"#
            ));
            assert!(children(&shape).is_empty(), "{child}");
            assert!(diagnostics.is_empty(), "{child}");
            assert!(
                shape
                    .text_box_content
                    .iter()
                    .any(|block| matches!(block, TextBoxBlockWire::Unsupported { .. })),
                "{child}"
            );
        }
        let (shape, diagnostics) = parse(
            r#"<v:group style="width:100pt;height:50pt"><v:shape style="width:10;height:10"><v:imagedata r:id="missing"/></v:shape></v:group>"#,
        );
        assert!(children(&shape).is_empty());
        assert!(diagnostics.is_empty());
        assert!(shape
            .text_box_content
            .iter()
            .any(|block| matches!(block, TextBoxBlockWire::Unsupported { .. })));
    }
}
