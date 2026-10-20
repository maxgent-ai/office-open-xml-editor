//! Built-in SpreadsheetML PivotTable styles (ECMA-376 Part 1 §18.8.40,
//! Annex G presetTableStyles.xml) as typed differential formats.
//!
//! A workbook names a built-in style without defining it (pivotTableStyleInfo
//! in XLSX, SXAddl_SXCView_SXDTableStyleClient in XLS), so every spreadsheet
//! producer resolves it from this one table. Colors stay theme references
//! (§18.8.3); callers resolve them against the workbook theme, e.g. with
//! [`crate::spreadsheet_color::resolve_color`].

/// A theme color reference (`theme`, optional `tint`) of a preset format.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetColor {
    pub theme: u32,
    pub tint: Option<f64>,
}

/// `<font>`: the preset formats only set bold and a color.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetFont {
    pub bold: bool,
    pub color: Option<PresetColor>,
}

/// `<fill><patternFill>`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetFill {
    pub pattern_type: &'static str,
    pub fg: Option<PresetColor>,
    pub bg: Option<PresetColor>,
}

/// A border edge; `style` is `None` when the edge element has no style
/// attribute (an explicitly present, style-less edge).
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetEdge {
    pub style: Option<&'static str>,
    pub color: Option<PresetColor>,
}

/// `<border>` edges, including the region-interior `horizontal`/`vertical`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetBorder {
    pub left: Option<PresetEdge>,
    pub right: Option<PresetEdge>,
    pub top: Option<PresetEdge>,
    pub bottom: Option<PresetEdge>,
    pub horizontal: Option<PresetEdge>,
    pub vertical: Option<PresetEdge>,
}

/// One preset `<dxf>`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetDxf {
    pub font: Option<PresetFont>,
    pub fill: Option<PresetFill>,
    pub border: Option<PresetBorder>,
}

/// One `tableStyleElement` of a preset style.
#[derive(Clone, Copy, Debug)]
pub struct PresetElement {
    /// ST_TableStyleType (§18.18.82), e.g. `firstRowStripe`.
    pub kind: &'static str,
    /// Band size (`size`, default 1).
    pub size: u32,
    pub dxf: PresetDxf,
}

// Small component dictionaries avoid repeating Option/f64 and string-pointer
// padding in WASM for every Annex G Dxf. Zero denotes absence and
// positive indices denote present components (including empty, style-less
// edges). This preserves §18.8.6's explicit edge clearing and §18.8.41's
// element layering. Tints retain the original f64 values; no quantization or
// theme substitution occurs. Decode only the requested style, without a
// global decoded cache or an intermediate allocation.
type PresetStyleEntry = (&'static str, u16, u16);

include!("spreadsheet_style_presets.generated.rs");

fn component<T: Copy>(table: &[T], index: impl Into<usize>) -> Option<T> {
    index.into().checked_sub(1).map(|index| table[index])
}

fn preset_color(index: u8) -> Option<PresetColor> {
    let [theme, tint] = component(&PRESET_COLORS, index)?;
    Some(PresetColor {
        theme: u32::from(theme),
        tint: component(&PRESET_TINTS, tint),
    })
}

fn preset_edge(index: u8) -> Option<PresetEdge> {
    let [style, color] = component(&PRESET_EDGES, index)?;
    Some(PresetEdge {
        style: component(&PRESET_EDGE_STYLES, style),
        color: preset_color(color),
    })
}

fn decode_dxf(index: u16) -> PresetDxf {
    let [font, fill, border_lo, border_hi] = PRESET_DXFS[usize::from(index)];
    PresetDxf {
        font: component(&PRESET_FONTS, font).map(|[bold, color]| PresetFont {
            bold: bold != 0,
            color: preset_color(color),
        }),
        fill: component(&PRESET_FILLS, fill).map(|[pattern, fg, bg]| PresetFill {
            pattern_type: PRESET_PATTERNS[usize::from(pattern) - 1],
            fg: preset_color(fg),
            bg: preset_color(bg),
        }),
        border: component(&PRESET_BORDERS, u16::from_le_bytes([border_lo, border_hi])).map(
            |[left, right, top, bottom, horizontal, vertical]| PresetBorder {
                left: preset_edge(left),
                right: preset_edge(right),
                top: preset_edge(top),
                bottom: preset_edge(bottom),
                horizontal: preset_edge(horizontal),
                vertical: preset_edge(vertical),
            },
        ),
    }
}

fn style_elements(
    styles: &'static [PresetStyleEntry],
    name: &str,
) -> Option<impl ExactSizeIterator<Item = PresetElement>> {
    let &(_, start, end) = styles.iter().find(|(preset, _, _)| *preset == name)?;
    Some(
        PRESET_ELEMENTS[usize::from(start)..usize::from(end)]
            .iter()
            .map(|&[kind, size, dxf_lo, dxf_hi]| PresetElement {
                kind: PRESET_KINDS[usize::from(kind) - 1],
                size: PRESET_SIZES[usize::from(size) - 1],
                dxf: decode_dxf(u16::from_le_bytes([dxf_lo, dxf_hi])),
            }),
    )
}

/// The elements of the built-in PivotTable style `name` (e.g.
/// `PivotStyleLight16`), or `None` when `name` is not a built-in style.
/// Elements retain Annex G order and band sizes (§18.18.82 / §18.8.41).
pub fn pivot_style(name: &str) -> Option<impl ExactSizeIterator<Item = PresetElement>> {
    style_elements(&PIVOT_STYLES, name)
}

/// The elements of the built-in Table style `name` (e.g.
/// `TableStyleMedium2`), or `None` when `name` is not a built-in style.
/// Elements retain Annex G order and band sizes (§18.18.82 / §18.8.41).
pub fn table_style(name: &str) -> Option<impl ExactSizeIterator<Item = PresetElement>> {
    style_elements(&TABLE_STYLES, name)
}

#[cfg(test)]
mod tests {
    use super::*;

    // Captured from the pre-packing implementation at 41d29dbd6. Every preset
    // crosses all ST_TableStyleType values, including absent elements. This
    // catches lost presence/clearing, order, band sizes and exact tint values.
    #[test]
    fn packed_presets_match_original_resolution_exhaustively() {
        use std::fmt::Write;
        assert_eq!(PIVOT_STYLES.len(), 84);
        assert!(pivot_style("NoSuchStyle").is_none());
        let mut actual = String::new();
        for name in [
            "PivotStyleMedium28",
            "PivotStyleMedium27",
            "PivotStyleMedium26",
            "PivotStyleMedium25",
            "PivotStyleMedium24",
            "PivotStyleMedium23",
            "PivotStyleMedium22",
            "PivotStyleMedium21",
            "PivotStyleMedium20",
            "PivotStyleMedium19",
            "PivotStyleMedium18",
            "PivotStyleMedium17",
            "PivotStyleMedium16",
            "PivotStyleMedium15",
            "PivotStyleMedium14",
            "PivotStyleMedium13",
            "PivotStyleMedium12",
            "PivotStyleMedium11",
            "PivotStyleMedium10",
            "PivotStyleMedium9",
            "PivotStyleMedium8",
            "PivotStyleMedium7",
            "PivotStyleMedium6",
            "PivotStyleMedium5",
            "PivotStyleMedium4",
            "PivotStyleMedium3",
            "PivotStyleMedium2",
            "PivotStyleMedium1",
            "PivotStyleLight28",
            "PivotStyleLight27",
            "PivotStyleLight26",
            "PivotStyleLight25",
            "PivotStyleLight24",
            "PivotStyleLight23",
            "PivotStyleLight22",
            "PivotStyleLight21",
            "PivotStyleLight20",
            "PivotStyleLight19",
            "PivotStyleLight18",
            "PivotStyleLight17",
            "PivotStyleLight16",
            "PivotStyleLight15",
            "PivotStyleLight14",
            "PivotStyleLight13",
            "PivotStyleLight12",
            "PivotStyleLight11",
            "PivotStyleLight10",
            "PivotStyleLight9",
            "PivotStyleLight8",
            "PivotStyleLight7",
            "PivotStyleLight6",
            "PivotStyleLight5",
            "PivotStyleLight4",
            "PivotStyleLight3",
            "PivotStyleLight2",
            "PivotStyleLight1",
            "PivotStyleDark28",
            "PivotStyleDark27",
            "PivotStyleDark26",
            "PivotStyleDark25",
            "PivotStyleDark24",
            "PivotStyleDark23",
            "PivotStyleDark22",
            "PivotStyleDark21",
            "PivotStyleDark20",
            "PivotStyleDark19",
            "PivotStyleDark18",
            "PivotStyleDark17",
            "PivotStyleDark16",
            "PivotStyleDark15",
            "PivotStyleDark14",
            "PivotStyleDark13",
            "PivotStyleDark12",
            "PivotStyleDark11",
            "PivotStyleDark10",
            "PivotStyleDark9",
            "PivotStyleDark8",
            "PivotStyleDark7",
            "PivotStyleDark6",
            "PivotStyleDark5",
            "PivotStyleDark4",
            "PivotStyleDark3",
            "PivotStyleDark2",
            "PivotStyleDark1",
        ] {
            let elements = pivot_style(name).unwrap().collect::<Vec<_>>();
            writeln!(
                actual,
                "{} order {:?}",
                name,
                elements.iter().map(|e| e.kind).collect::<Vec<_>>()
            )
            .unwrap();
            for kind in [
                "wholeTable",
                "headerRow",
                "totalRow",
                "firstColumn",
                "lastColumn",
                "firstRowStripe",
                "secondRowStripe",
                "firstColumnStripe",
                "secondColumnStripe",
                "firstHeaderCell",
                "lastHeaderCell",
                "firstTotalCell",
                "lastTotalCell",
                "firstSubtotalColumn",
                "secondSubtotalColumn",
                "thirdSubtotalColumn",
                "firstSubtotalRow",
                "secondSubtotalRow",
                "thirdSubtotalRow",
                "blankRow",
                "firstColumnSubheading",
                "secondColumnSubheading",
                "thirdColumnSubheading",
                "firstRowSubheading",
                "secondRowSubheading",
                "thirdRowSubheading",
                "pageFieldLabels",
                "pageFieldValues",
            ] {
                writeln!(
                    actual,
                    "{} {} {:?}",
                    name,
                    kind,
                    elements
                        .iter()
                        .find(|e| e.kind == kind)
                        .map(|e| (e.size, e.dxf))
                )
                .unwrap();
            }
        }
        assert_eq!(actual, include_str!("../tests/fixtures/pivot-presets.txt"));
    }
}
