//! Isolated Excel compatibility rules for classic chart-space defaults.

use ooxml_common::chart::{ChartExElementStyle, ChartModel};
use std::collections::BTreeMap;

/// Materialize Excel's implicit classic-chart frame as an XLSX parser policy.
/// ECMA-376 constrains `<c:style>` to 1..48 but does not define the built-in
/// visual recipes. Excel-produced samples covering every style and the omitted
/// style showed 10pt rounded corners throughout.
/// Styles 1..40 use a white chart area with a 1pt #898989 outline; styles
/// 41..48 use a black chart area without an outline. Omitted style behaves as
/// style 2. Direct chart-space paint and a linked `chartArea` role remain
/// authoritative through the shared renderer's direct > linked precedence.
pub(crate) fn apply_excel_classic_chart_space_frame(chart: &mut ChartModel) {
    chart.rounded_corners.get_or_insert(true);

    let roles = chart.chart_style_roles.get_or_insert_with(BTreeMap::new);
    if roles.contains_key("chartArea") {
        return;
    }

    let style = if chart.legacy_chart_style.unwrap_or(2) <= 40 {
        ChartExElementStyle {
            fill_colors: Some(vec![Some("FFFFFF".to_string())]),
            line_colors: Some(vec![Some("898989".to_string())]),
            line_width_emu: Some(12_700),
            ..Default::default()
        }
    } else {
        ChartExElementStyle {
            fill_colors: Some(vec![Some("000000".to_string())]),
            line_hidden: Some(true),
            ..Default::default()
        }
    };
    roles.insert("chartArea".to_string(), style);
}

#[cfg(test)]
mod tests {
    use super::*;
    use ooxml_common::chart::ColorResolver;

    struct NoColors;

    impl ColorResolver for NoColors {
        fn resolve_solid_fill(&self, _node: roxmltree::Node<'_, '_>) -> Option<String> {
            None
        }
    }

    fn chart(style: Option<u8>, rounded: Option<bool>) -> ChartModel {
        let style = style
            .map(|value| format!(r#"<c:style val="{value}"/>"#))
            .unwrap_or_default();
        let rounded = rounded
            .map(|value| format!(r#"<c:roundedCorners val="{}"/>"#, u8::from(value)))
            .unwrap_or_default();
        let xml = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">{style}{rounded}<c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let document = roxmltree::Document::parse(&xml).expect("chart XML");
        let mut chart = ooxml_common::chart::parse_chart_part(document.root_element(), &NoColors)
            .expect("classic chart");
        apply_excel_classic_chart_space_frame(&mut chart);
        chart
    }

    #[test]
    fn style_boundary_selects_excel_light_and_dark_frames() {
        let light = chart(Some(40), None);
        let light_frame = &light.chart_style_roles.as_ref().unwrap()["chartArea"];
        assert_eq!(light.rounded_corners, Some(true));
        assert_eq!(
            light_frame.fill_colors.as_deref(),
            Some(&[Some("FFFFFF".to_string())][..])
        );
        assert_eq!(
            light_frame.line_colors.as_deref(),
            Some(&[Some("898989".to_string())][..])
        );
        assert_eq!(light_frame.line_width_emu, Some(12_700));

        let dark = chart(Some(41), None);
        let dark_frame = &dark.chart_style_roles.as_ref().unwrap()["chartArea"];
        assert_eq!(
            dark_frame.fill_colors.as_deref(),
            Some(&[Some("000000".to_string())][..])
        );
        assert_eq!(dark_frame.line_hidden, Some(true));
    }

    #[test]
    fn omitted_style_and_explicit_square_corners_keep_their_precedence() {
        let omitted = chart(None, None);
        let frame = &omitted.chart_style_roles.as_ref().unwrap()["chartArea"];
        assert_eq!(frame.line_width_emu, Some(12_700));
        assert_eq!(chart(Some(2), Some(false)).rounded_corners, Some(false));

        let mut linked = chart(Some(40), None);
        linked.chart_style_roles.as_mut().unwrap().insert(
            "chartArea".to_string(),
            ChartExElementStyle {
                line_colors: Some(vec![Some("FF0000".to_string())]),
                line_width_emu: Some(25_400),
                ..Default::default()
            },
        );
        apply_excel_classic_chart_space_frame(&mut linked);
        let linked_frame = &linked.chart_style_roles.as_ref().unwrap()["chartArea"];
        assert_eq!(
            linked_frame.line_colors.as_deref(),
            Some(&[Some("FF0000".to_string())][..])
        );
        assert_eq!(linked_frame.line_width_emu, Some(25_400));
    }
}
