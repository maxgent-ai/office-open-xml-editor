//! Isolated Word compatibility rules for classic chart-space defaults.

use ooxml_common::chart::{ChartExElementStyle, ChartModel};
use std::collections::BTreeMap;

/// Materialize Word's implicit classic-chart frame as a bounded DOCX parser
/// compatibility policy. ECMA-376 constrains `<c:style>` to 1..48 but does not
/// define Office's visual recipes. A complete Word-produced style matrix found
/// 10pt rounded chart-space corners for all 48 styles; omitted style and 1..40
/// use a 0.5pt #898989 outline, while 41..48 use no outline. A linked
/// `chartArea` role and direct chart-space paint remain authoritative through
/// the shared renderer's normal direct > linked precedence.
pub(crate) fn apply_word_classic_chart_space_frame(chart: &mut ChartModel) {
    chart.rounded_corners.get_or_insert(true);

    let roles = chart.chart_style_roles.get_or_insert_with(BTreeMap::new);
    if roles.contains_key("chartArea") {
        return;
    }

    let style = if chart.legacy_chart_style.unwrap_or(2) <= 40 {
        ChartExElementStyle {
            fill_no_style: Some(true),
            line_colors: Some(vec![Some("898989".to_string())]),
            line_width_emu: Some(6_350),
            ..Default::default()
        }
    } else {
        ChartExElementStyle {
            fill_no_style: Some(true),
            line_hidden: Some(true),
            ..Default::default()
        }
    };
    roles.insert("chartArea".to_string(), style);
}
