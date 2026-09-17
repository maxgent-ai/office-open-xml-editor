//! Isolated PowerPoint compatibility rules for classic chart-space defaults.

use ooxml_common::chart::{ChartExElementStyle, ChartModel};
use std::collections::BTreeMap;

/// Materialize PowerPoint's implicit classic-chart frame as a PPTX parser
/// policy. ECMA-376 constrains `<c:style>` to 1..48 but does not define the
/// built-in visual recipes. PowerPoint-produced samples covering every style,
/// the omitted style, and direct-paint controls showed 10pt rounded corners
/// throughout. Styles 1..32 add no chart-area paint; styles 33..40 use white
/// with a 0.75pt #898989 outline; styles 41..48 use black without an outline.
/// Omitted style behaves as style 2. PowerPoint's PDF export also draws a
/// separate square 0.14pt graphic-frame path for every style; it is not a
/// chart-space style and therefore is intentionally not synthesized here.
/// Direct chart-space paint and a linked `chartArea` role remain authoritative
/// through the shared renderer's direct > linked precedence.
pub(crate) fn apply_powerpoint_classic_chart_space_frame(chart: &mut ChartModel) {
    chart.rounded_corners.get_or_insert(true);

    if chart
        .chart_style_roles
        .as_ref()
        .is_some_and(|roles| roles.contains_key("chartArea"))
    {
        return;
    }

    let style = match chart.legacy_chart_style.unwrap_or(2) {
        1..=32 => return,
        33..=40 => ChartExElementStyle {
            fill_colors: Some(vec![Some("FFFFFF".to_string())]),
            line_colors: Some(vec![Some("898989".to_string())]),
            line_width_emu: Some(9_525),
            ..Default::default()
        },
        _ => ChartExElementStyle {
            fill_colors: Some(vec![Some("000000".to_string())]),
            line_hidden: Some(true),
            ..Default::default()
        },
    };
    chart
        .chart_style_roles
        .get_or_insert_with(BTreeMap::new)
        .insert("chartArea".to_string(), style);
}
