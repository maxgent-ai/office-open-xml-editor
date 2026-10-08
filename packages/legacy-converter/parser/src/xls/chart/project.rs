//! RawChart -> shared ChartModel. Colors come from verified ShapePropsStream
//! XML when present (MS-XLS 2.4.258), otherwise from the BIFF records: a
//! non-automatic AreaFormat/LineFormat uses its palette index, because Excel
//! ignores their rgb fields on load (MS-XLS 2.4.3 footnotes <23>/<24>).
//! Automatic formatting is left unset so the shared renderer applies its
//! ordinary automatic series colors.
use super::reader::{Cached, Format, GroupKind, RawChart};
use ooxml_common::chart::{ChartModel, ChartSeries};
use ooxml_common::color::{parse_color_node, ThemeResolver, TintMode};

pub(crate) struct Palette<'a> {
    /// Global Font record by FontX one-based index (2.4.123).
    pub global_font: &'a dyn Fn(u16) -> Option<super::super::styles::ChartFont>,
    /// Decode a chart-local Font record.
    pub decode_font: &'a dyn Fn(&[u8]) -> Option<super::super::styles::ChartFont>,
    pub global_font_count: usize,
    /// Resolved palette color for an Icv (MS-XLS 2.5.161), "#RRGGBB" or "RRGGBB".
    pub icv: &'a dyn Fn(u16) -> Option<String>,
    /// Theme colors in clrScheme order: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink.
    pub theme: [Option<String>; 12],
}

impl ThemeResolver for Palette<'_> {
    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        let slot = match ooxml_common::color::default_scheme_slot(name) {
            "dk1" => 0,
            "lt1" => 1,
            "dk2" => 2,
            "lt2" => 3,
            "accent1" => 4,
            "accent2" => 5,
            "accent3" => 6,
            "accent4" => 7,
            "accent5" => 8,
            "accent6" => 9,
            "hlink" => 10,
            "folHlink" => 11,
            _ => return None,
        };
        self.theme[slot]
            .as_deref()
            .map(|hex| hex.trim_start_matches('#').to_owned())
    }
}

impl Palette<'_> {
    /// FontX.iFont (2.4.123): 0 is the chart default font, indices up to the
    /// global count are one-based global Font records, later indices address
    /// this chart's own Font records.
    fn font(&self, raw: &RawChart, index: u16) -> Option<super::super::styles::ChartFont> {
        let index = if index == 0 { raw.default_font? } else { index };
        if index == 0 {
            return None;
        }
        if usize::from(index) <= self.global_font_count {
            (self.global_font)(index)
        } else {
            let local = usize::from(index) - self.global_font_count - 1;
            raw.local_fonts
                .get(local)
                .and_then(|data| (self.decode_font)(data))
        }
    }

    /// Theme accentN for an automatic series or varied point (N = i mod 6 + 1).
    fn accent(&self, index: usize) -> Option<String> {
        self.theme[4 + index % 6]
            .as_deref()
            .map(|hex| hex.trim_start_matches('#').to_uppercase())
    }
}

struct Paint {
    fill: Option<String>,
    fill_hidden: bool,
    line: Option<String>,
    line_hidden: bool,
    line_width_emu: Option<u32>,
}

/// The shared chart model carries RRGGBB without a leading '#'.
fn hex(value: String) -> String {
    value.trim_start_matches('#').to_uppercase()
}

fn child<'a, 'input>(
    parent: roxmltree::Node<'a, 'input>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'input>> {
    parent
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == name)
}

/// Shape XML (DrawingML spPr) fill/line, resolved with the workbook theme.
/// A verified but empty stream is an empty spPr: automatic formatting, which
/// supersedes the BIFF records (ShapePropsStream is their superset, 2.4.258).
/// Excel then draws the chart-style automatic colors, as for XLSX.
fn xml_paint(xml: &str, palette: &Palette<'_>) -> Option<Paint> {
    if xml.trim().is_empty() {
        return Some(Paint {
            fill: None,
            fill_hidden: false,
            line: None,
            line_hidden: false,
            line_width_emu: None,
        });
    }
    let document = roxmltree::Document::parse(xml).ok()?;
    let root = document.root_element();
    if root.tag_name().name() != "spPr" {
        return None;
    }
    let fill = child(root, "solidFill")
        .and_then(|node| parse_color_node(node, palette, TintMode::PowerPointLinear))
        .map(hex);
    let fill_hidden = child(root, "noFill").is_some();
    let line_node = child(root, "ln");
    let line = line_node
        .and_then(|ln| child(ln, "solidFill"))
        .and_then(|node| parse_color_node(node, palette, TintMode::PowerPointLinear))
        .map(hex);
    let line_hidden = line_node.and_then(|ln| child(ln, "noFill")).is_some();
    let line_width_emu = line_node
        .and_then(|ln| ln.attribute("w"))
        .and_then(|w| w.parse().ok());
    Some(Paint {
        fill,
        fill_hidden,
        line,
        line_hidden,
        line_width_emu,
    })
}

fn biff_paint(format: &Format, palette: &Palette<'_>) -> Paint {
    let u16_at = |bytes: &[u8], at: usize| u16::from_le_bytes([bytes[at], bytes[at + 1]]);
    let (mut fill, mut fill_hidden) = (None, false);
    if let Some(area) = format.area {
        let automatic = u16_at(&area, 10) & 1 != 0;
        let pattern = u16_at(&area, 8);
        if !automatic {
            if pattern == 0 {
                fill_hidden = true;
            } else {
                fill = (palette.icv)(u16_at(&area, 12)).map(hex);
            }
        }
    }
    let (mut line, mut line_hidden) = (None, false);
    let line_width_emu = None;
    if let Some(format) = format.line {
        let automatic = u16_at(&format, 8) & 1 != 0;
        let pattern = u16_at(&format, 4);
        if !automatic {
            if pattern == 5 {
                line_hidden = true;
            } else {
                line = (palette.icv)(u16_at(&format, 10)).map(hex);
                // LineFormat.we (2.4.156) names hairline/narrow/medium/wide
                // weights without a normative length; leave the width to the
                // renderer rather than inventing point values.
            }
        }
    }
    Paint {
        fill,
        fill_hidden,
        line,
        line_hidden,
        line_width_emu,
    }
}

fn paint(format: &Format, palette: &Palette<'_>) -> Paint {
    match format
        .shape_xml
        .get(&0)
        .and_then(|xml| xml_paint(xml, palette))
    {
        Some(xml) => xml,
        None => biff_paint(format, palette),
    }
}

fn number_text(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < 1e15 {
        format!("{}", value as i64)
    } else {
        format!("{value}")
    }
}

fn series_type(kind: GroupKind) -> &'static str {
    match kind {
        GroupKind::Bar { .. } => "bar",
        GroupKind::Line { .. } => "line",
        GroupKind::Area { .. } => "area",
        GroupKind::Pie { hole, .. } if hole > 0 => "doughnut",
        GroupKind::Pie { .. } | GroupKind::OfPie => "pie",
        GroupKind::Scatter { .. } => "scatter",
        GroupKind::Radar { .. } => "radar",
        GroupKind::Surface => "surface",
    }
}

fn chart_type(kind: GroupKind) -> String {
    match kind {
        GroupKind::Bar {
            horizontal,
            stacked,
            percent,
            ..
        } => ooxml_common::chart::canonical_chart_type(
            "bar",
            if horizontal { "bar" } else { "col" },
            grouping(stacked, percent),
        ),
        GroupKind::Line { stacked, percent } => {
            ooxml_common::chart::canonical_chart_type("line", "", grouping(stacked, percent))
        }
        GroupKind::Area { stacked, percent } => {
            ooxml_common::chart::canonical_chart_type("area", "", grouping(stacked, percent))
        }
        GroupKind::Pie { hole, .. } if hole > 0 => "doughnut".into(),
        GroupKind::Pie { .. } => "pie".into(),
        GroupKind::OfPie => "ofPie".into(),
        GroupKind::Scatter { bubbles: true } => "bubble".into(),
        GroupKind::Scatter { bubbles: false } => "scatter".into(),
        GroupKind::Radar { .. } => "radar".into(),
        GroupKind::Surface => "surface".into(),
    }
}

fn grouping(stacked: bool, percent: bool) -> &'static str {
    match (stacked, percent) {
        (true, true) => "percentStacked",
        (true, false) => "stacked",
        _ => "clustered",
    }
}

/// MarkerFormat.imk (MS-XLS 2.4.160) -> ST_MarkerStyle. The short and long
/// bar markers are the "dot" and "dash" styles of the same Office marker set.
fn marker_symbol(imk: u16) -> Option<&'static str> {
    Some(match imk {
        0 => "none",
        1 => "square",
        2 => "diamond",
        3 => "triangle",
        4 => "x",
        5 => "star",
        6 => "dot",
        7 => "dash",
        8 => "circle",
        9 => "plus",
        _ => return None,
    })
}

struct Marker {
    symbol: Option<&'static str>,
    size_pt: Option<f64>,
    fill: Option<String>,
    line: Option<String>,
}

fn marker(format: &Format, palette: &Palette<'_>) -> Option<Marker> {
    let data = format.marker?;
    let u16_at = |at: usize| u16::from_le_bytes([data[at], data[at + 1]]);
    let flags = u16_at(10);
    if flags & 1 != 0 {
        // fAuto: automatic marker formatting.
        return None;
    }
    let xml = format
        .shape_xml
        .get(&1)
        .and_then(|xml| xml_paint(xml, palette));
    let (fill, line) = match xml {
        Some(paint) => (paint.fill, paint.line),
        None => (
            (flags & 0x10 == 0)
                .then(|| (palette.icv)(u16_at(14)).map(hex))
                .flatten(),
            (flags & 0x20 == 0)
                .then(|| (palette.icv)(u16_at(12)).map(hex))
                .flatten(),
        ),
    };
    let size = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
    Some(Marker {
        symbol: marker_symbol(u16_at(8)),
        // miSize is in twips (1/20 point).
        size_pt: (size > 0).then(|| f64::from(size) / 20.0),
        fill,
        line,
    })
}

/// AttachedLabel (MS-XLS 2.4.5) flags -> series data labels.
fn data_labels(flags: u16) -> Option<ooxml_common::chart::ChartSeriesDataLabels> {
    let labels = ooxml_common::chart::ChartSeriesDataLabels {
        show_val: flags & 0x01 != 0,
        show_percent: flags & 0x02 != 0 || flags & 0x04 != 0,
        show_cat_name: flags & 0x10 != 0 || flags & 0x04 != 0,
        show_bubble_size: flags & 0x20 != 0,
        show_ser_name: flags & 0x40 != 0,
        ..Default::default()
    };
    (labels.show_val
        || labels.show_percent
        || labels.show_cat_name
        || labels.show_bubble_size
        || labels.show_ser_name)
        .then_some(labels)
}

/// Resolves a BRAI worksheet reference (rgce) to its cells in order.
pub(crate) type References<'a> = dyn Fn(&[u8]) -> Result<Option<Vec<Option<Cached>>>, String> + 'a;

/// Project a raw chart. Returns `None` when there is no drawable series.
/// The chart data cache is authoritative when present (MS-XLS 2.2.3.2);
/// otherwise series parts are read from the referenced worksheet cells.
#[cfg(test)]
type UnboundedReferences = dyn Fn(&[u8]) -> Option<Vec<Option<Cached>>>;

#[cfg(test)]
pub(crate) fn project(
    raw: &RawChart,
    palette: &Palette<'_>,
    references: &UnboundedReferences,
) -> Option<ChartModel> {
    project_bounded(
        raw,
        palette,
        &|bytes| Ok(references(bytes)),
        &std::cell::Cell::new(usize::MAX),
    )
    .unwrap()
}

fn charge(budget: &std::cell::Cell<usize>, bytes: usize) -> Result<(), String> {
    let left = budget
        .get()
        .checked_sub(bytes)
        .ok_or_else(|| super::super::unsupported("XLS chart model retention budget exceeded"))?;
    budget.set(left);
    Ok(())
}

/// Workbook aggregate admission includes typed slots, reference/cache copies,
/// and generated categories before each potentially amplified allocation.
pub(crate) fn project_bounded(
    raw: &RawChart,
    palette: &Palette<'_>,
    references: &References<'_>,
    budget: &std::cell::Cell<usize>,
) -> Result<Option<ChartModel>, String> {
    let Some(primary) = raw.groups.first() else {
        return Ok(None);
    };
    let values = raw.cache.get(&1);
    let categories = raw.cache.get(&2);
    let bubbles = raw.cache.get(&3);
    let point_count = |cache: Option<&std::collections::BTreeMap<(u16, u16), Cached>>,
                       series: u16| {
        cache.map_or(0, |cache| {
            cache
                .range((series, 0)..=(series, u16::MAX))
                .map(|((_, point), _)| usize::from(*point) + 1)
                .max()
                .unwrap_or(0)
        })
    };
    let mut model = ChartModel {
        chart_type: chart_type(primary.kind),
        title: raw.title.clone(),
        title_present: raw.title.is_some(),
        cat_axis_cross_between: "between".into(),
        val_axis_major_tick_mark: "out".into(),
        cat_axis_major_tick_mark: "out".into(),
        plot_visible_only: Some(raw.plot_visible_only),
        ..ChartModel::default()
    };
    let mut series_models = Vec::new();
    for (index, series) in raw.series.iter().enumerate() {
        if series.trend_or_error {
            continue;
        }
        let key = index as u16;
        let group = raw.groups.get(usize::from(series.group)).unwrap_or(primary);
        let part = |numindex: u16,
                    cache: Option<&std::collections::BTreeMap<(u16, u16), Cached>>|
         -> Result<Vec<Option<Cached>>, String> {
            let count = point_count(cache, key);
            if count > 0 {
                charge(
                    budget,
                    count
                        .checked_mul(std::mem::size_of::<Option<Cached>>())
                        .ok_or("XLS chart allocation overflow")?,
                )?;
                let mut values = Vec::with_capacity(count);
                for point in 0..count {
                    let value = cache.and_then(|c| c.get(&(key, point as u16)));
                    if let Some(Cached::Text(text)) = value {
                        charge(budget, text.len())?;
                    }
                    values.push(value.cloned());
                }
                return Ok(values);
            }
            match series.references[usize::from(numindex)].as_deref() {
                Some(bytes) => Ok(references(bytes)?.unwrap_or_default()),
                None => Ok(Vec::new()),
            }
        };
        let values = part(1, values)?;
        charge(
            budget,
            values
                .len()
                .checked_mul(std::mem::size_of::<Option<f64>>())
                .ok_or("XLS chart allocation overflow")?,
        )?;
        let series_values = values
            .into_iter()
            .map(|value| match value {
                Some(Cached::Number(number)) => Some(number),
                _ => None,
            })
            .collect::<Vec<_>>();
        let count = series_values.len();
        let categories = part(2, categories)?;
        let category_count = if categories.is_empty() {
            count
        } else {
            categories.len()
        };
        charge(
            budget,
            category_count
                .checked_mul(std::mem::size_of::<String>())
                .ok_or("XLS chart allocation overflow")?,
        )?;
        let mut series_categories = Vec::with_capacity(category_count);
        if categories.is_empty() {
            for point in 1..=count {
                let text = point.to_string();
                charge(budget, text.len())?;
                series_categories.push(text);
            }
        } else {
            for value in categories {
                let text = match value {
                    Some(Cached::Text(text)) => text,
                    Some(Cached::Number(number)) => {
                        let text = number_text(number);
                        charge(budget, text.len())?;
                        text
                    }
                    None => String::new(),
                };
                series_categories.push(text);
            }
        }
        let series_paint = series
            .series_format
            .as_ref()
            .or(group.default_format.as_ref())
            .map(|format| paint(format, palette));
        let mut model_series = ChartSeries {
            name: series.name.clone().unwrap_or_default(),
            values: series_values,
            series_type: Some(series_type(group.kind).into()),
            use_secondary_axis: (group.axis_group == 1).then_some(true),
            categories: Some(series_categories),
            ..ChartSeries::default()
        };
        // Line-drawn groups take the series color from its line (the fill
        // only paints markers or areas); filled groups use the fill.
        let line_drawn = matches!(
            group.kind,
            GroupKind::Line { .. }
                | GroupKind::Scatter { bubbles: false }
                | GroupKind::Radar { filled: false }
        );
        if let Some(paint) = series_paint {
            model_series.color = if line_drawn {
                paint.line.clone().or(paint.fill)
            } else {
                paint.fill
            };
            model_series.line_color = paint.line;
            model_series.line_width_emu = paint.line_width_emu;
            if paint.line_hidden {
                model_series.line_hidden = Some(true);
            }
        }
        // Automatic series formatting uses the theme accents in series order,
        // matching the XLSX parser's automatic series color.
        if model_series.color.is_none() && !group.varied_colors {
            model_series.color = palette.accent(index);
        }
        if let Some(marker) = series
            .series_format
            .as_ref()
            .and_then(|f| marker(f, palette))
        {
            model_series.show_marker = Some(marker.symbol != Some("none"));
            model_series.marker_symbol = marker.symbol.map(str::to_owned);
            model_series.marker_size = marker.size_pt;
            model_series.marker_fill = marker.fill;
            model_series.marker_line = marker.line;
        }
        if let Some(labels) = series
            .series_format
            .as_ref()
            .and_then(|f| f.data_labels)
            .and_then(data_labels)
        {
            model_series.series_data_labels = Some(labels);
        }
        if let Some(format) = series.series_format.as_ref() {
            // AreaFormat.fInvertNeg (2.4.3) swaps the foreground and
            // background colors for negative values, so a solid negative
            // point is filled with icvBack.
            if let Some(area) = format
                .area
                .filter(|area| u16::from_le_bytes([area[10], area[11]]) & 2 != 0)
            {
                model_series.invert_if_negative = Some(true);
                if let Some(color) = (palette.icv)(u16::from_le_bytes([area[14], area[15]])) {
                    model_series.inverted_fill =
                        Some(ooxml_common::chart::ChartStyleFill::Solid { color: hex(color) });
                }
            }
            model_series.explosion = format.explosion.map(u32::from);
            if format.smooth {
                model_series.smooth = Some(true);
            }
        }
        if !series.point_formats.is_empty() || group.varied_colors {
            let points = count.max(
                series
                    .point_formats
                    .keys()
                    .map(|p| usize::from(*p) + 1)
                    .max()
                    .unwrap_or(0),
            );
            charge(
                budget,
                points
                    .checked_mul(std::mem::size_of::<Option<String>>())
                    .ok_or("XLS chart allocation overflow")?,
            )?;
            let colors = (0..points)
                .map(|point| {
                    series
                        .point_formats
                        .get(&(point as u16))
                        .and_then(|format| paint(format, palette).fill)
                        .or_else(|| group.varied_colors.then(|| palette.accent(point)).flatten())
                })
                .collect::<Vec<_>>();
            if colors.iter().any(Option::is_some) {
                model_series.data_point_colors = Some(colors);
            }
        }
        let sizes = part(3, bubbles)?;
        charge(
            budget,
            sizes
                .len()
                .checked_mul(std::mem::size_of::<Option<f64>>())
                .ok_or("XLS chart allocation overflow")?,
        )?;
        let sizes = sizes
            .into_iter()
            .map(|value| match value {
                Some(Cached::Number(number)) => Some(number),
                _ => None,
            })
            .collect::<Vec<_>>();
        if sizes.iter().any(Option::is_some) {
            model_series.bubble_sizes = Some(sizes);
        }
        use ooxml_common::chart::RetainedBytes;
        charge(
            budget,
            std::mem::size_of::<ChartSeries>()
                .saturating_add(usize::try_from(model_series.heap_bytes()).unwrap_or(usize::MAX)),
        )?;
        series_models.push(model_series);
    }
    if series_models.is_empty() {
        // A chart with no Series record at all is authored empty: Excel draws
        // its chart area (see ChartModel::authored_without_series). Series
        // that exist but resolve to nothing are not projected.
        if raw.series.iter().any(|series| !series.trend_or_error) {
            return Ok(None);
        }
        model.authored_without_series = true;
    }
    model.categories = series_models
        .first()
        .and_then(|series| series.categories.clone())
        .unwrap_or_default();
    model.series = series_models;
    if let Some(legend) = raw.groups.iter().find_map(|g| g.legend) {
        model.show_legend = true;
        model.legend_pos = Some(
            match legend.position {
                0 => "b",
                1 => "tr",
                2 => "t",
                4 => "l",
                _ => "r",
            }
            .into(),
        );
    }
    if let Some(axis) = raw.axes.iter().find(|a| a.kind == 1 && a.axis_group == 0) {
        // MS-XLS Bar/Line/Area f100 (2.4.15 / 2.4.155 / 2.4.2) declares
        // percentage stacking. Office BIFF percent-stacked bar output stores a
        // 50%-step ValueRange interval as 50 percentage points, while the
        // canonical OOXML model stores 0.5; ValueRange (2.4.341) itself does not
        // spell out this format-unit distinction. All f100 families share that
        // percentage axis. Convert bounds and both authored intervals together
        // to ratios; the shared renderer performs the inverse x100 conversion.
        // Select by the authored group, never by the scalar's magnitude.
        let percent_axis = matches!(
            primary.kind,
            GroupKind::Bar {
                stacked: true,
                percent: true,
                ..
            } | GroupKind::Line {
                stacked: true,
                percent: true
            } | GroupKind::Area {
                stacked: true,
                percent: true
            }
        );
        let canonical_value = |value: f64| if percent_axis { value / 100.0 } else { value };
        model.val_min = axis.min.map(canonical_value);
        model.val_max = axis.max.map(canonical_value);
        // MS-XLS 2.4.341 ValueRange: fAutoMajor/fAutoMinor suppress the
        // stored intervals, and fReversed changes the value-axis direction.
        // These are existing shared renderer semantics, not Office tuning.
        model.val_axis_major_unit = axis.major.map(canonical_value);
        model.val_axis_minor_unit = axis.minor.map(canonical_value);
        model.val_axis_orientation = axis.reversed.then(|| "maxMin".into());
    }
    // MS-XLS 2.4.327: authored Tick locations belong to their enclosing axis,
    // not to the chart or the other axis group. Preserve them in the existing
    // primary-axis renderer slots; a secondary/series axis cannot overwrite
    // those slots. When Tick is absent, retain the existing model defaults.
    for axis in raw.axes.iter().filter(|a| a.axis_group == 0) {
        let Some(ticks) = axis.ticks else { continue };
        match axis.kind {
            0 => {
                model.cat_axis_major_tick_mark = ticks.major.into();
                model.cat_axis_minor_tick_mark = Some(ticks.minor.into());
                model.cat_axis_tick_label_pos = Some(ticks.labels.into());
            }
            1 => {
                model.val_axis_major_tick_mark = ticks.major.into();
                model.val_axis_minor_tick_mark = Some(ticks.minor.into());
                model.val_axis_tick_label_pos = Some(ticks.labels.into());
            }
            _ => {}
        }
    }
    // Font records carry twips; the shared model uses hundredths of a point.
    let hpt = |twips: u16| i32::from(twips) * 5;
    let color = |font: &super::super::styles::ChartFont| font.color.clone().map(hex);
    if let Some(font) = raw.title_font.and_then(|i| palette.font(raw, i)) {
        model.title_font_size_hpt = Some(hpt(font.size_twips));
        model.title_font_bold = Some(font.bold);
        model.title_font_color = color(&font);
        model.title_font_face = Some(font.name.clone());
    }
    if let Some(font) = raw.legend_font.and_then(|i| palette.font(raw, i)) {
        model.legend_font_size_hpt = Some(hpt(font.size_twips));
        model.legend_font_bold = Some(font.bold);
        model.legend_font_color = color(&font);
        model.legend_font_face = Some(font.name.clone());
    }
    for axis in raw.axes.iter().filter(|a| a.axis_group == 0) {
        let Some(font) = axis.font.and_then(|i| palette.font(raw, i)) else {
            continue;
        };
        let (size, bold, italic, font_color, face) = (
            Some(hpt(font.size_twips)),
            Some(font.bold),
            Some(font.italic),
            color(&font),
            Some(font.name.clone()),
        );
        if axis.kind == 1 {
            model.val_axis_font_size_hpt = size;
            model.val_axis_font_bold = bold;
            model.val_axis_font_italic = italic;
            model.val_axis_font_color = font_color;
            model.val_axis_font_face = face;
        } else if axis.kind == 0 {
            model.cat_axis_font_size_hpt = size;
            model.cat_axis_font_bold = bold;
            model.cat_axis_font_italic = italic;
            model.cat_axis_font_color = font_color;
            model.cat_axis_font_face = face;
        }
    }
    model.val_axis_title = raw.axis_titles.get(&2).cloned();
    model.cat_axis_title = raw.axis_titles.get(&3).cloned();
    if let Some(format) = raw.chart_format.as_ref() {
        // An empty verified ShapePropsStream on the chart area is automatic
        // formatting, and Excel writes that automatic chart area into the
        // BIFF records it keeps for older readers (a solid 0x808080 hairline
        // over a white AreaFormat). The shared chart model has no automatic
        // chart-area outline, so the chart area takes those BIFF records:
        // Excel's PDFs of such XLS files (21 corpus charts, all authored in
        // XLSX without a chartSpace spPr) draw that gray outline, where the
        // automatic projection drew none. Series keep their automatic
        // colours: Excel's PDFs draw theme accents there, not the BIFF
        // palette colours written beside an empty stream.
        let empty_stream = format
            .shape_xml
            .get(&0)
            .is_some_and(|xml| xml.trim().is_empty());
        let area = if empty_stream {
            biff_paint(format, palette)
        } else {
            paint(format, palette)
        };
        model.chart_bg = area.fill;
        if area.fill_hidden {
            model.chart_fill_hidden = Some(true);
        }
        model.chart_border_color = area.line;
        model.chart_border_width_emu = area.line_width_emu;
        if area.line_hidden {
            model.chart_border_hidden = Some(true);
        }
    }
    if let Some(format) = raw.plot_format.as_ref() {
        let area = paint(format, palette);
        model.plot_area_bg = area.fill;
        if area.fill_hidden {
            model.plot_area_fill_hidden = Some(true);
        }
        model.plot_area_line_color = area.line;
        model.plot_area_line_width_emu = area.line_width_emu;
        if area.line_hidden {
            model.plot_area_line_hidden = Some(true);
        }
    }
    // Resource policy counts both admitted expansion work and the complete
    // retained chart, including authored empty charts, titles and axes.
    // This aggregate quota is independent of Office rendering semantics.
    use ooxml_common::chart::RetainedBytes;
    charge(
        budget,
        std::mem::size_of::<ChartModel>()
            .saturating_add(usize::try_from(model.heap_bytes()).unwrap_or(usize::MAX)),
    )?;
    Ok(Some(model))
}
