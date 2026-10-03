//! Alternative shape XML (`metroBlob`, MS-ODRAW 2.3.4.41, opid 0x3A9).
//!
//! PowerPoint 2007 and later store, per shape, an OPC package holding the
//! shape's DrawingML. MS-ODRAW says the property SHOULD be ignored and that
//! Office deletes it when the shape is modified; implementation note 32 says
//! Office 2007/2010 do not ignore it. Its text characters are masked
//! (letters replaced by spaces or underscores); the characters themselves
//! come from the binary.
//!
//! Evidence from PowerPoint 16 PDF exports:
//! - controls whose binary is byte-identical and whose alternative XML alone
//!   was edited (package rebuilt at the same length) render the edit: a fill
//!   color changed in the XML only and character spacing raised in the XML
//!   only are drawn;
//! - Office-saved decks render character spacing, shrink-to-fit scales and
//!   per-paragraph indents that exist only in the alternative XML;
//! - a control whose binary shape properties were edited (preset adjust and
//!   fill color) while the alternative XML was kept renders the binary, and
//!   a deck whose alternative states 10.5 pt text over a binary 10 pt run
//!   renders without the alternative's run properties: the alternative is
//!   not used once it disagrees with the binary.
//!
//! The alternative therefore carries display information the binary lacks,
//! and the binary projection alone does not reproduce a shape whose
//! alternative PowerPoint uses. Each shape resolves to one of three
//! outcomes (`adopt`):
//! - no alternative part (a package with only the `downRev` checksums), or
//!   an ordinary alternative that verifiably disagrees with the binary on a
//!   compared attribute: the binary projection;
//! - an alternative that agrees on every applicable compared attribute: the
//!   alternative, with the binary's transform, identifier and characters;
//!   a binary freeform keeps its outline, and a placeholder inherits locally
//!   omitted shape properties from the binary projection;
//! - anything else fails closed as unsupported: an oversized, over-budget,
//!   ambiguous or unreadable package or theme, a part that is not a shape or
//!   connector, unverified relationship references, and a compared attribute that the
//!   two forms state in ways this reader cannot equate. Placeholder shape
//!   properties omitted locally inherit from the binary projection; their
//!   transform always follows the binary anchor (PowerPoint 16 controls).
//!   A disagreement in a locally stated placeholder fill or geometry is also
//!   unsupported because its edited side cannot be determined.
//!
//! The compared attributes are those the evidence covers: geometry (preset
//! name and adjust values, or custom paths, except a binary freeform with an
//! XML preset), the untransformed position, size, rotation and flips for
//! ordinary shapes, the recorded fill when locally stated, run font size,
//! bold and italic where both state them, and the text structure (the same
//! paragraphs, runs and line breaks at the same UTF-16 lengths). Stroke,
//! effects and other text properties are not compared, so a binary edit that
//! preserves every compared attribute would still adopt a stale alternative.
//! The downrev checksums that PowerPoint stores beside the XML cannot be
//! recomputed (they are not a checksum of the binary records), so this
//! structural agreement stands in for them.
use super::*;
#[cfg(any(test, feature = "direct-ppt"))]
use pptx_model::{ShapeElement, TextRun};

/// Implementation resource policy, not a format limit.
#[cfg(feature = "direct-ppt")]
const MAX_BLOB_BYTES: usize = 4 * 1024 * 1024;
#[cfg(feature = "direct-ppt")]
const MAX_PART_BYTES: u64 = 1024 * 1024;
#[cfg(feature = "direct-ppt")]
const MAX_PACKAGE_BYTES: u64 = 4 * 1024 * 1024;
const MAX_THEME_BYTES: u64 = 2 * 1024 * 1024;
const MAX_THEME_ENTRIES: usize = 64;
/// One master unit is 1/576 inch.
#[cfg(any(test, feature = "direct-ppt"))]
const EMU_PER_MASTER_UNIT: f64 = 914_400.0 / 576.0;
/// One unit of the binary anchor's resolution, plus the EMU rounding of the
/// binary's master-unit conversion on each side.
#[cfg(any(test, feature = "direct-ppt"))]
const XFRM_TOLERANCE_UNITS: f64 = 1.0 + 1.0 / EMU_PER_MASTER_UNIT;
/// Normalized path coordinates are ratios of the same integer vertices and
/// extents on both sides; allow only floating-point rounding.
#[cfg(any(test, feature = "direct-ppt"))]
const PATH_TOLERANCE: f64 = 1e-9;

/// The main master's round-trip theme (MS-PPT RoundTripTheme12Atom, 0x040E)
/// and color map (RoundTripColorMapping12Atom, 0x040F), against which the
/// alternative XML's theme references resolve.
pub(in crate::ppt) enum Theme {
    // Read only by `adopt`.
    #[cfg_attr(not(feature = "direct-ppt"), allow(dead_code))]
    Readable {
        theme_xml: String,
        clr_map: Option<String>,
        /// A master theme is shared by many placeholders; resolve its style
        /// matrix once when checking explicit line/effect references.
        format_scheme: std::cell::OnceCell<ooxml_common::theme::ThemeFormatScheme>,
    },
    /// Malformed, oversized or ambiguous: no alternative XML on this
    /// master's slides can be resolved.
    Unreadable,
}

impl Theme {
    #[cfg(feature = "direct-ppt")]
    fn format_scheme(&self) -> Option<&ooxml_common::theme::ThemeFormatScheme> {
        let Theme::Readable {
            theme_xml,
            format_scheme,
            ..
        } = self
        else {
            return None;
        };
        Some(format_scheme.get_or_init(|| ooxml_common::theme::ThemeFormatScheme::parse(theme_xml)))
    }
    /// ECMA-376 §20.1.6.2 supplies hlink/folHlink in the theme color scheme.
    /// PowerPoint-saved PPTX controls retain those slots for legacy slides;
    /// an inherited binary text color must not hide them on linked text.
    pub(in crate::ppt) fn hyperlink_colors(&self) -> Option<(String, String)> {
        let Theme::Readable { theme_xml, .. } = self else {
            return None;
        };
        let colors = ooxml_common::theme::ThemeColorScheme::parse(theme_xml);
        let color = |slot| {
            colors
                .get(slot)
                .filter(|value| value.len() == 6 && value.bytes().all(|b| b.is_ascii_hexdigit()))
                .map(str::to_owned)
        };
        Some((color("hlink")?, color("folHlink")?))
    }
}

/// Read the round-trip theme of a main master's child records; `None` when
/// the master has none.
pub(in crate::ppt) fn master_theme(
    records: &[Record<'_>],
    decoded_budget: &mut usize,
) -> Result<Option<Theme>, String> {
    let mut themes = records.iter().filter(|r| r.kind == 0x040e);
    let Some(theme) = themes.next() else {
        return Ok(None);
    };
    let mut budget_exceeded = false;
    let mut readable = || -> Option<Theme> {
        if themes.next().is_some() {
            return None;
        }
        let mut maps = records.iter().filter(|r| r.kind == 0x040f);
        let clr_map = match (maps.next(), maps.next()) {
            (Some(map), None) => {
                if super::charge_text(decoded_budget, map.payload.len()).is_err() {
                    budget_exceeded = true;
                    return None;
                }
                Some(std::str::from_utf8(map.payload).ok()?.to_owned())
            }
            (None, None) => None,
            _ => return None,
        };
        Some(Theme::Readable {
            theme_xml: theme_part(theme.payload, decoded_budget, &mut budget_exceeded)?,
            clr_map,
            format_scheme: std::cell::OnceCell::new(),
        })
    };
    let theme = readable().unwrap_or(Theme::Unreadable);
    if budget_exceeded {
        return Err(unsupported("PowerPoint theme decoded byte budget exceeded"));
    }
    Ok(Some(theme))
}

fn theme_part(
    package: &[u8],
    decoded_budget: &mut usize,
    budget_exceeded: &mut bool,
) -> Option<String> {
    use std::io::Read;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(package)).ok()?;
    if archive.len() > MAX_THEME_ENTRIES {
        return None;
    }
    let mut read = |name: &str| -> Option<String> {
        let index = crate::opc_part::entry_index(&archive, name)?;
        let entry = archive.by_index(index).ok()?;
        if entry.size() > MAX_THEME_BYTES || entry.encrypted() {
            return None;
        }
        // Charge every inflated theme/relationship part before allocation.
        // All main masters share this budget; per-entry ZIP limits alone do
        // not bound retained themes from a highly compressed presentation.
        let declared_bytes = usize::try_from(entry.size()).ok()?;
        if super::charge_text(decoded_budget, declared_bytes).is_err() {
            *budget_exceeded = true;
            return None;
        }
        let mut text = String::new();
        entry
            .take(declared_bytes as u64 + 1)
            .read_to_string(&mut text)
            .ok()?;
        // ZIP metadata is untrusted. Never retain actual inflated bytes in
        // excess of the declared, precharged size (including forged sizes).
        (text.len() == declared_bytes).then_some(text)
    };
    // Follow the package relationships: root -> theme manager -> theme.
    let target = |rels: &str, source: &str, kind: &str| -> Option<String> {
        let doc = roxmltree::Document::parse(rels).ok()?;
        let mut found = doc.root_element().children().filter(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type").is_some_and(|t| t.ends_with(kind))
        });
        let first = found.next()?;
        if first
            .attribute("TargetMode")
            .is_some_and(|mode| mode != "Internal")
            || found.next().is_some()
        {
            return None;
        }
        ooxml_common::rels::resolve_part_name(source, first.attribute("Target")?)
    };
    let manager = target(&read("_rels/.rels")?, "", "/officeDocument")?;
    let (dir, name) = manager.rsplit_once('/').unwrap_or(("", &manager));
    let rels = if dir.is_empty() {
        format!("_rels/{name}.rels")
    } else {
        format!("{dir}/_rels/{name}.rels")
    };
    let theme = target(&read(&rels)?, &manager, "/theme")?;
    let xml = read(&theme)?;
    let doc = roxmltree::Document::parse(&xml).ok()?;
    let root = doc.root_element();
    (root.tag_name().name() == "theme" && ooxml_common::ns::is_a_ns(root.tag_name().namespace()))
        .then_some(xml)
}

/// What the binary shape records, for comparison with its alternative.
#[cfg(any(test, feature = "direct-ppt"))]
#[cfg_attr(not(feature = "direct-ppt"), allow(dead_code))]
pub(in crate::ppt) struct BinaryShape<'a> {
    /// The binary projection: geometry, adjust values and text runs.
    pub element: &'a ShapeElement,
    /// The untransformed (group-local) transform.
    pub leaf: &'a pptx_model::Transform,
    /// Whether the shape sits in a group's child coordinate space.
    pub nested: bool,
    /// The binary text characters, if any.
    pub text: Option<&'a str>,
    /// Direct CFStyle font-size bits, before master inheritance. Only used
    /// when an XML picture bullet needs inherited-size normalization.
    pub direct_size_authored: Option<&'a [bool]>,
    pub fill: RecordedFill,
    /// Authored per-path (fill, stroke) flags of custom geometry, before
    /// PowerPoint's open-path display rule (`officeart::geometry`).
    pub path_paint: Option<Vec<(bool, bool)>>,
    /// The range of each converted preset adjust value under the rounding
    /// of the binary anchor (`officeart::preset::adjustment_bounds`).
    pub adjust_bounds: [Option<(f64, f64)>; 8],
    /// Already admitted OfficeArt image bytes for a picture fill. The image
    /// path in `element.fill` identifies the same retained resource.
    pub image_bytes: Option<Vec<u8>>,
}

/// The fill a binary shape records (MS-ODRAW 2.3.7), as a model fill.
#[cfg(any(test, feature = "direct-ppt"))]
pub(in crate::ppt) enum RecordedFill {
    /// The geometry has no fill area (lines and connectors).
    NotDisplayed,
    /// No fill property is stated.
    Unstated,
    /// The stated fill; `None` when it is off.
    Stated(Option<Box<pptx_model::Fill>>),
    /// A stated fill this reader cannot restate as a model fill.
    Unknown,
}

/// Agreement of one attribute that both forms record.
#[cfg(any(test, feature = "direct-ppt"))]
#[derive(Clone, Copy, Debug, PartialEq)]
enum Verdict {
    Same,
    /// Both forms state the attribute, with different values.
    Differs(&'static str),
    /// Both forms state the attribute, but not in a form this reader can
    /// compare: agreement is unknown.
    Unverifiable(&'static str),
}

/// Adopt only when every compared attribute is the same. Any verified
/// difference keeps the binary projection: PowerPoint ignores an alternative
/// that disagrees with the binary shape. Otherwise, when some attribute
/// cannot be compared, the display is unknown and the shape fails closed.
#[cfg(any(test, feature = "direct-ppt"))]
fn decide(verdicts: &[Verdict]) -> Result<bool, String> {
    if verdicts.iter().any(|v| matches!(v, Verdict::Differs(_))) {
        return Ok(false);
    }
    match verdicts.iter().find_map(|v| match v {
        Verdict::Unverifiable(what) => Some(what),
        _ => None,
    }) {
        Some(what) => Err(unverifiable(what)),
        None => Ok(true),
    }
}

#[cfg(any(test, feature = "direct-ppt"))]
pub(in crate::ppt) fn unverifiable(what: &str) -> String {
    unsupported(format!(
        "PowerPoint alternative shape XML {what} cannot be compared with the binary shape"
    ))
}

/// The alternative part named by the package's root relationships. Office
/// writes one `downRev` relationship (the checksums) and at most one
/// relationship to the alternative DrawingML part; its type names the part's
/// kind (observed: shapeXml, connectorXml, groupShapeXml, pictureXml,
/// graphicFrameDoc, inkXml). A package with only `downRev` carries no
/// alternative.
#[cfg(feature = "direct-ppt")]
fn alternative_part(blob: &[u8]) -> Result<Option<(String, String)>, String> {
    use std::io::Read;
    let unreadable = || unsupported("unreadable PowerPoint alternative shape XML package");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).map_err(|_| unreadable())?;
    let index = crate::opc_part::entry_index(&archive, "_rels/.rels").ok_or_else(unreadable)?;
    let entry = archive.by_index(index).map_err(|_| unreadable())?;
    if entry.size() > MAX_PART_BYTES || entry.encrypted() {
        return Err(unreadable());
    }
    let mut rels = String::new();
    entry
        .take(MAX_PART_BYTES + 1)
        .read_to_string(&mut rels)
        .map_err(|_| unreadable())?;
    let doc = roxmltree::Document::parse(&rels).map_err(|_| unreadable())?;
    let mut found = None;
    for node in doc.root_element().children().filter(|n| n.is_element()) {
        if node.tag_name().name() != "Relationship" {
            return Err(unreadable());
        }
        let kind = node.attribute("Type").ok_or_else(unreadable)?;
        let kind = kind.rsplit_once('/').map_or(kind, |(_, kind)| kind);
        if kind == "downRev" {
            continue;
        }
        let target = ooxml_common::rels::resolve_part_name(
            "",
            node.attribute("Target").ok_or_else(unreadable)?,
        )
        .ok_or_else(|| unverifiable("package relationship"))?;
        if node
            .attribute("TargetMode")
            .is_some_and(|mode| mode != "Internal")
            || found.replace((kind.to_owned(), target)).is_some()
        {
            return Err(unverifiable("package relationship"));
        }
    }
    Ok(found)
}

/// Check the OPC references used by the alternative without fetching an
/// external target or treating a resolved target as evidence that the binary
/// and XML agree. ECMA-376 Part 2 §6.5.2.3 locates the relationship part
/// beside its source part; §6.5.3.4 keeps external targets outside the ZIP.
/// PowerPoint precedence for these references still needs its control PDFs.
#[cfg(feature = "direct-ppt")]
fn read_blob_reference_part<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    name: &str,
) -> Result<String, String> {
    use std::io::Read;
    let unreadable = || unverifiable("relationship");
    let index = crate::opc_part::entry_index(archive, name).ok_or_else(unreadable)?;
    let entry = archive.by_index(index).map_err(|_| unreadable())?;
    if entry.size() > MAX_PART_BYTES || entry.encrypted() {
        return Err(unreadable());
    }
    let mut text = String::new();
    entry
        .take(MAX_PART_BYTES + 1)
        .read_to_string(&mut text)
        .map_err(|_| unreadable())?;
    if text.len() as u64 > MAX_PART_BYTES {
        return Err(unreadable());
    }
    Ok(text)
}

#[cfg(feature = "direct-ppt")]
fn validate_blob_references(blob: &[u8], part: &str) -> Result<(), String> {
    use std::collections::BTreeSet;

    let unreadable = || unverifiable("relationship");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).map_err(|_| unreadable())?;
    // Resource policy: avoid a reference-by-entry quadratic scan of a ZIP
    // with a huge number of tiny entries. Ordinary Office shape blobs are
    // much smaller; this limit makes the passive check bounded.
    if archive.len() > 128 {
        return Err(unreadable());
    }
    let xml = read_blob_reference_part(&mut archive, part)?;
    let doc = ooxml_common::depth::parse_guarded(&xml).map_err(|_| unreadable())?;
    let ids: BTreeSet<&str> = doc
        .descendants()
        .flat_map(|node| node.attributes())
        .filter(|attribute| ooxml_common::ns::is_r_ns(attribute.namespace()))
        .map(|attribute| attribute.value())
        .collect();
    if ids.is_empty() {
        // The selected theme style may own the reference instead. It is not
        // in this blob and remains unverifiable at the adoption gate.
        return Ok(());
    }
    let rels_path = ooxml_common::rels::relationship_part_path(part);
    let rels = read_blob_reference_part(&mut archive, &rels_path)?;
    let rels = ooxml_common::depth::parse_guarded(&rels).map_err(|_| unreadable())?;
    let root = rels.root_element();
    if root.tag_name().name() != "Relationships"
        || root.tag_name().namespace()
            != Some("http://schemas.openxmlformats.org/package/2006/relationships")
    {
        return Err(unreadable());
    }
    let mut seen = BTreeSet::new();
    let mut matched = BTreeSet::new();
    for relationship in root.children().filter(|node| node.is_element()) {
        if relationship.tag_name().name() != "Relationship"
            || relationship.tag_name().namespace() != root.tag_name().namespace()
        {
            return Err(unreadable());
        }
        let id = relationship.attribute("Id").ok_or_else(unreadable)?;
        if !seen.insert(id) {
            return Err(unreadable());
        }
        if !ids.contains(id) {
            continue;
        }
        let target = relationship.attribute("Target").ok_or_else(unreadable)?;
        if target.is_empty() || relationship.attribute("Type").is_none_or(str::is_empty) {
            return Err(unreadable());
        }
        match relationship.attribute("TargetMode") {
            None | Some("Internal") => {
                let name =
                    ooxml_common::rels::resolve_part_name(part, target).ok_or_else(unreadable)?;
                if crate::opc_part::entry_index(&archive, &name).is_none() {
                    return Err(unreadable());
                }
            }
            Some("External") => {} // Opaque URI; never read or fetched.
            _ => return Err(unreadable()),
        }
        matched.insert(id);
    }
    if matched != ids {
        return Err(unreadable());
    }
    Ok(())
}

/// Read one image that is the only reference in the alternative shape part.
/// The ZIP and part size limits were checked at the adoption boundary. A
/// non-image or second reference is left to the typed relationship gate.
#[cfg(feature = "direct-ppt")]
fn sole_blob_image(blob: &[u8], part: &str, image_path: &str) -> Result<Vec<u8>, String> {
    use std::io::Read;
    let unreadable = || unverifiable("relationship");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).map_err(|_| unreadable())?;
    let xml = read_blob_reference_part(&mut archive, part)?;
    let doc = ooxml_common::depth::parse_guarded(&xml).map_err(|_| unreadable())?;
    let refs: Vec<_> = doc
        .descendants()
        .flat_map(|node| node.attributes().map(move |attribute| (node, attribute)))
        .filter(|(_, attribute)| ooxml_common::ns::is_r_ns(attribute.namespace()))
        .collect();
    if refs.len() != 1 || refs[0].0.tag_name().name() != "blip" || refs[0].1.name() != "embed" {
        return Err(unreadable());
    }
    let id = refs[0].1.value();
    let rels_path = ooxml_common::rels::relationship_part_path(part);
    let rels = read_blob_reference_part(&mut archive, &rels_path)?;
    let rels = ooxml_common::depth::parse_guarded(&rels).map_err(|_| unreadable())?;
    let mut matching = rels.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().name() == "Relationship"
            && node.attribute("Id") == Some(id)
    });
    let relationship = matching.next().ok_or_else(unreadable)?;
    if matching.next().is_some()
        || !relationship
            .attribute("Type")
            .is_some_and(|kind| kind.ends_with("/image"))
        || relationship
            .attribute("TargetMode")
            .is_some_and(|mode| mode != "Internal")
    {
        return Err(unreadable());
    }
    let target = ooxml_common::rels::resolve_part_name(
        part,
        relationship.attribute("Target").ok_or_else(unreadable)?,
    )
    .ok_or_else(unreadable)?;
    if target != image_path {
        return Err(unreadable());
    }
    let index = crate::opc_part::entry_index(&archive, &target).ok_or_else(unreadable)?;
    let entry = archive.by_index(index).map_err(|_| unreadable())?;
    if entry.size() > MAX_PART_BYTES || entry.encrypted() {
        return Err(unreadable());
    }
    let mut bytes = Vec::new();
    entry
        .take(MAX_PART_BYTES + 1)
        .read_to_end(&mut bytes)
        .map_err(|_| unreadable())?;
    if bytes.len() as u64 > MAX_PART_BYTES {
        return Err(unreadable());
    }
    Ok(bytes)
}

/// A shape whose only OPC dependencies are external click URLs can be
/// compared to [MS-PPT] InteractiveInfoAtom/ExHyperlinkContainer targets.
/// Other relationship classes still have no binary counterpart here.
#[cfg(feature = "direct-ppt")]
fn blob_hyperlinks_only(blob: &[u8], part: &str) -> Result<bool, String> {
    use std::collections::BTreeSet;
    let unreadable = || unverifiable("relationship");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).map_err(|_| unreadable())?;
    let xml = read_blob_reference_part(&mut archive, part)?;
    let doc = ooxml_common::depth::parse_guarded(&xml).map_err(|_| unreadable())?;
    let mut ids = BTreeSet::new();
    for node in doc.descendants() {
        for attr in node
            .attributes()
            .filter(|a| ooxml_common::ns::is_r_ns(a.namespace()))
        {
            if node.tag_name().name() != "hlinkClick" || attr.name() != "id" {
                return Ok(false);
            }
            ids.insert(attr.value().to_owned());
        }
    }
    if ids.is_empty() {
        return Ok(false);
    }
    let rels_path = ooxml_common::rels::relationship_part_path(part);
    let rels = read_blob_reference_part(&mut archive, &rels_path)?;
    let rels = ooxml_common::depth::parse_guarded(&rels).map_err(|_| unreadable())?;
    Ok(ids.iter().all(|id| {
        rels.descendants()
            .filter(|n| {
                n.is_element()
                    && n.tag_name().name() == "Relationship"
                    && n.attribute("Id") == Some(id)
            })
            .count()
            == 1
            && rels.descendants().any(|n| {
                n.is_element()
                    && n.tag_name().name() == "Relationship"
                    && n.attribute("Id") == Some(id)
                    && n.attribute("Type")
                        .is_some_and(|t| t.ends_with("/hyperlink"))
                    && n.attribute("TargetMode") == Some("External")
            })
    }))
}

/// Return only image references used by DrawingML picture bullets. The
/// presentation model already has a generic buBlip marker; this extracts its
/// passive OPC targets for the legacy resource provider. Unrelated links are
/// left at the typed gate.
#[cfg(any(test, feature = "direct-ppt"))]
type BlobBulletAsset = (String, &'static str, Vec<u8>);

#[cfg(feature = "direct-ppt")]
fn blob_bullet_images(blob: &[u8], part: &str) -> Result<Option<Vec<BlobBulletAsset>>, String> {
    use std::collections::BTreeMap;
    use std::io::Read;
    let unreadable = || unverifiable("relationship");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).map_err(|_| unreadable())?;
    let xml = read_blob_reference_part(&mut archive, part)?;
    let doc = ooxml_common::depth::parse_guarded(&xml).map_err(|_| unreadable())?;
    let mut ids = Vec::new();
    for node in doc.descendants() {
        for attr in node
            .attributes()
            .filter(|a| ooxml_common::ns::is_r_ns(a.namespace()))
        {
            if node.tag_name().name() != "blip"
                || node
                    .parent()
                    .is_none_or(|p| p.tag_name().name() != "buBlip")
                || attr.name() != "embed"
            {
                return Ok(None);
            }
            ids.push(attr.value().to_owned());
        }
    }
    if ids.is_empty() {
        return Ok(None);
    }
    let rels_path = ooxml_common::rels::relationship_part_path(part);
    let rels = read_blob_reference_part(&mut archive, &rels_path)?;
    let rels = ooxml_common::depth::parse_guarded(&rels).map_err(|_| unreadable())?;
    let mut images = BTreeMap::new();
    for id in ids {
        let mut matches = rels.descendants().filter(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Id") == Some(id.as_str())
        });
        let relation = matches.next().ok_or_else(unreadable)?;
        if matches.next().is_some()
            || !relation
                .attribute("Type")
                .is_some_and(|t| t.ends_with("/image"))
            || relation
                .attribute("TargetMode")
                .is_some_and(|m| m != "Internal")
        {
            return Err(unreadable());
        }
        let path = ooxml_common::rels::resolve_part_name(
            part,
            relation.attribute("Target").ok_or_else(unreadable)?,
        )
        .ok_or_else(unreadable)?;
        if images.contains_key(&path) {
            continue;
        }
        let extension = match path
            .rsplit_once('.')
            .map(|(_, ext)| ext.to_ascii_lowercase())
            .as_deref()
        {
            Some("gif") => "gif",
            Some("png") => "png",
            Some("jpg" | "jpeg") => "jpg",
            _ => return Err(unreadable()),
        };
        let index = crate::opc_part::entry_index(&archive, &path).ok_or_else(unreadable)?;
        let entry = archive.by_index(index).map_err(|_| unreadable())?;
        if entry.size() > MAX_PART_BYTES || entry.encrypted() {
            return Err(unreadable());
        }
        let mut bytes = Vec::new();
        entry
            .take(MAX_PART_BYTES + 1)
            .read_to_end(&mut bytes)
            .map_err(|_| unreadable())?;
        let signature = match extension {
            "gif" => bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a"),
            "png" => bytes.starts_with(b"\x89PNG\r\n\x1a\n"),
            _ => bytes.starts_with(&[0xff, 0xd8, 0xff]),
        };
        if !signature || bytes.len() as u64 > MAX_PART_BYTES {
            return Err(unreadable());
        }
        images.insert(path, (extension, bytes));
    }
    Ok(Some(
        images
            .into_iter()
            .map(|(path, (ext, bytes))| (path, ext, bytes))
            .collect(),
    ))
}

/// OfficeArt msofillPattern stores a 10x10 GIF but tiles only its upper-left
/// 8x8 cells. PowerPoint 16 controls of pct30 and ltUpDiag show that white
/// pixels map to the XML foreground and black to the background. Decode only
/// a single bounded GIF image; other encodings remain unverifiable.
#[cfg(any(test, feature = "direct-ppt"))]
fn binary_pattern_cells(bytes: &[u8]) -> Option<[u8; 8]> {
    if bytes.len() > 4096 || bytes.len() < 26 || !matches!(&bytes[..6], b"GIF87a" | b"GIF89a") {
        return None;
    }
    let word = |at: usize| -> Option<u16> {
        Some(u16::from_le_bytes(bytes.get(at..at + 2)?.try_into().ok()?))
    };
    if word(6)? != 10 || word(8)? != 10 {
        return None;
    }
    let mut at = 13usize;
    let table_len = |flag: u8| 3usize.checked_mul(1usize << (usize::from(flag & 7) + 1));
    let mut palette = if bytes[10] & 0x80 != 0 {
        let len = table_len(bytes[10])?;
        let table = bytes.get(at..at + len)?.to_vec();
        at += len;
        table
    } else {
        return None;
    };
    let mut transparent = None;
    let mut pixels = None;
    while at < bytes.len() {
        let kind = *bytes.get(at)?;
        at += 1;
        match kind {
            0x21 => {
                let label = *bytes.get(at)?;
                at += 1;
                if label == 0xf9 {
                    if *bytes.get(at)? != 4 {
                        return None;
                    }
                    let packed = *bytes.get(at + 1)?;
                    transparent = (packed & 1 != 0).then_some(*bytes.get(at + 4)?);
                }
                loop {
                    let size = usize::from(*bytes.get(at)?);
                    at += 1;
                    if size == 0 {
                        break;
                    }
                    at = at.checked_add(size)?;
                    bytes.get(..at)?;
                }
            }
            0x2c if pixels.is_none() => {
                if word(at)? != 0
                    || word(at + 2)? != 0
                    || word(at + 4)? != 10
                    || word(at + 6)? != 10
                {
                    return None;
                }
                let flags = *bytes.get(at + 8)?;
                if flags & 0x40 != 0 {
                    return None; // interlaced cells need a different row order
                }
                at += 9;
                if flags & 0x80 != 0 {
                    let len = table_len(flags)?;
                    palette = bytes.get(at..at + len)?.to_vec();
                    at += len;
                }
                let min_code = *bytes.get(at)?;
                at += 1;
                if !(2..=8).contains(&min_code) {
                    return None;
                }
                let mut data = Vec::new();
                loop {
                    let size = usize::from(*bytes.get(at)?);
                    at += 1;
                    if size == 0 {
                        break;
                    }
                    data.extend_from_slice(bytes.get(at..at + size)?);
                    at += size;
                }
                pixels = Some(decode_pattern_lzw(&data, min_code)?);
            }
            0x3b => break,
            _ => return None,
        }
    }
    let pixels = pixels?;
    let mut rows = [0u8; 8];
    for y in 0..8 {
        for x in 0..8 {
            let index = pixels[y * 10 + x];
            if transparent == Some(index) {
                return None;
            }
            let color = palette.get(usize::from(index) * 3..usize::from(index) * 3 + 3)?;
            match color {
                [255, 255, 255] => rows[y] |= 0x80 >> x,
                [0, 0, 0] => {}
                _ => return None,
            }
        }
    }
    Some(rows)
}

/// GIF LZW streams are little-endian bit packed. The 10x10 admission bound
/// caps both output and dictionary work, independent of compressed input.
#[cfg(any(test, feature = "direct-ppt"))]
fn decode_pattern_lzw(data: &[u8], min_code: u8) -> Option<Vec<u8>> {
    let clear = 1usize << min_code;
    let end = clear + 1;
    let initial = || {
        let mut table: Vec<Vec<u8>> = (0..clear).map(|value| vec![value as u8]).collect();
        table.extend([Vec::new(), Vec::new()]);
        table
    };
    let mut table = initial();
    let mut width = usize::from(min_code) + 1;
    let mut bit = 0usize;
    let mut previous: Option<Vec<u8>> = None;
    let mut output = Vec::with_capacity(100);
    for _ in 0..512 {
        if bit + width > data.len() * 8 {
            return None;
        }
        let mut code = 0usize;
        for shift in 0..width {
            let position = bit + shift;
            code |= usize::from((data[position / 8] >> (position % 8)) & 1) << shift;
        }
        bit += width;
        if code == clear {
            table = initial();
            width = usize::from(min_code) + 1;
            previous = None;
            continue;
        }
        if code == end {
            return (output.len() == 100).then_some(output);
        }
        let entry = if code < table.len() && !table[code].is_empty() {
            table[code].clone()
        } else if code == table.len() {
            let mut entry = previous.clone()?;
            entry.push(*entry.first()?);
            entry
        } else {
            return None;
        };
        if output.len().checked_add(entry.len())? > 100 {
            return None;
        }
        output.extend_from_slice(&entry);
        if let Some(mut prev) = previous {
            if table.len() < 4096 {
                prev.push(*entry.first()?);
                table.push(prev);
                if table.len() == 1 << width && width < 12 {
                    width += 1;
                }
            }
        }
        previous = Some(entry);
    }
    None
}

#[cfg(any(test, feature = "direct-ppt"))]
fn same_binary_pattern(binary: &BinaryShape<'_>, alternative: &ShapeElement) -> Option<Verdict> {
    use pptx_model::Fill;
    let (
        Some(Fill::Image {
            duotone: Some(colors),
            ..
        }),
        Some(Fill::Pattern { fg, bg, preset }),
    ) = (&binary.element.fill, &alternative.fill)
    else {
        return None;
    };
    let expected = match preset.as_str() {
        "pct30" => [0xaa, 0x44, 0xaa, 0x11, 0xaa, 0x44, 0xaa, 0x11],
        "ltUpDiag" => [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88],
        _ => return Some(Verdict::Unverifiable("fill")),
    };
    let Some(cells) = binary.image_bytes.as_deref().and_then(binary_pattern_cells) else {
        return Some(Verdict::Unverifiable("fill"));
    };
    // MS-ODRAW binary pattern BLIP: black is clr1/background; white is
    // clr2/foreground. The XML pattern preset names the same 8x8 cells.
    Some(
        if cells == expected
            && same_color(&colors.clr1, bg) == Some(true)
            && same_color(&colors.clr2, fg) == Some(true)
        {
            Verdict::Same
        } else {
            Verdict::Differs("pattern fill")
        },
    )
}

/// Resolve the alternative shape XML of a binary shape: `Ok(Some)` adopts
/// it; `Ok(None)` keeps the binary projection, because the package carries
/// no alternative part or the alternative verifiably disagrees with the
/// binary; an error means agreement cannot be established (oversized, over
/// budget, unreadable, or not comparable), and the shape fails closed. The
/// adopted shape is charged to `model_budget` at its serialized size.
#[cfg(feature = "direct-ppt")]
pub(in crate::ppt) fn adopt(
    binary: &BinaryShape<'_>,
    blob: &[u8],
    theme: &Theme,
    media: &mut super::media::SpanStore,
    work_budget: &mut usize,
    text_budget: &mut usize,
    model_budget: &mut usize,
) -> Result<Option<ShapeElement>, String> {
    if blob.len() > MAX_BLOB_BYTES {
        return Err(unsupported(
            "PowerPoint alternative shape XML exceeds its size limit",
        ));
    }
    // One parse attempt of work, and the package's declared inflated size
    // against the session's decoded-text byte budget.
    *work_budget = work_budget
        .checked_sub(1)
        .ok_or_else(|| unsupported("PowerPoint alternative shape XML work budget exceeded"))?;
    super::charge_text(text_budget, inflated_size(blob)?)?;
    let Some((kind, part)) = alternative_part(blob)? else {
        return Ok(None);
    };
    // Only shapes and connectors are parts this projection can compare
    // with a binary shape.
    if !matches!(kind.as_str(), "shapeXml" | "connectorXml") {
        return Err(unverifiable("part"));
    }
    let Theme::Readable {
        theme_xml, clr_map, ..
    } = theme
    else {
        return Err(unverifiable("theme"));
    };
    // The modern parser can tolerate an absent/malformed style matrix. An
    // optional binary alternative has a stronger contract: its theme facts
    // must be readable before claiming equivalence with the binary shape.
    let theme_doc = roxmltree::Document::parse(theme_xml).map_err(|_| unverifiable("theme"))?;
    let theme_root = theme_doc.root_element();
    if theme_root.tag_name().name() != "theme"
        || !ooxml_common::ns::is_a_ns(theme_root.tag_name().namespace())
    {
        return Err(unverifiable("theme"));
    }
    let parsed = pptx_parser::parse_standalone_shape_part(
        blob,
        &part,
        theme_xml,
        clr_map.as_deref(),
        MAX_PART_BYTES,
        MAX_PACKAGE_BYTES,
    );
    let parsed = parsed
        .map_err(|error| {
            unsupported(format!(
                "unreadable PowerPoint alternative shape XML: {error}"
            ))
        })?
        .ok_or_else(|| unverifiable("part"))?;
    // PowerPoint 16 controls retain XML-only picture crop edits when the
    // relationship-backed image bytes equal the binary BLIP, and switch to
    // binary after a BLIP-index edit changes those bytes. Compare the passive
    // package target with the already admitted binary resource. On adoption
    // the XML crop uses that resource's path; no blob media is retained.
    // External click hyperlinks use the binary text-action catalog. Other
    // references, including theme styles, remain at the typed gate.
    let (image_comparison, hyperlink_comparison, bullet_assets) = if parsed.relationship_references
    {
        validate_blob_references(blob, &part)?;
        let bullet_assets = blob_bullet_images(blob, &part)?;
        match (
            &parsed.element.fill,
            &binary.element.fill,
            binary.image_bytes.as_deref(),
        ) {
            (
                Some(pptx_model::Fill::Image { image_path, .. }),
                Some(pptx_model::Fill::Image { .. }),
                Some(binary_bytes),
            ) => {
                let blob_bytes = sole_blob_image(blob, &part, image_path)?;
                (
                    Some(if blob_bytes == binary_bytes {
                        Verdict::Same
                    } else {
                        Verdict::Differs("image content")
                    }),
                    None,
                    None,
                )
            }
            _ if blob_hyperlinks_only(blob, &part)? => (
                None,
                Some(same_hyperlinks(binary.element, &parsed.element)),
                None,
            ),
            _ if bullet_assets.is_some() => (None, None, bullet_assets),
            _ => return Err(unverifiable("relationship")),
        }
    } else {
        (None, None, None)
    };
    // Validate all selected style references, including ordinary shapes.
    // Missing matrix entries must not adopt the modern renderer's fallback.
    let locals = placeholder_locals(blob, &part, theme)?;
    let local = if parsed.placeholder {
        Some(locals)
    } else {
        None
    };
    let mut shape = parsed.element;
    let pattern_comparison = same_binary_pattern(binary, &shape);
    // PowerPoint 16 controls with a binary OfficeArt freeform and an XML
    // preset (frame and trapezoid, including visibly matching outlines)
    // retain the binary outline when the XML preset alone changes. An XML
    // fill edit on the same freeform is visible. This is geometry-specific
    // precedence, not rejection of the whole alternative. The converse
    // (binary preset, XML freeform) has not been established and stays closed.
    let binary_freeform_preset =
        binary.element.geometry == "custGeom" && shape.geometry != "custGeom";
    let substituted = {
        let mut candidate = shape.clone();
        substitute_text_bounded(&mut candidate, binary.text, model_budget).map(|()| candidate)
    };
    // A local placeholder override is ambiguous when it disagrees: XML-only
    // geometry/fill edits render from the alternative, whereas binary-only
    // edits on the same controls invalidate it and render from the binary.
    // The proprietary checksum cannot be recomputed here to identify which
    // side was edited. Reject that pair rather than guessing its provenance.
    let local_override = |verdict, stated, attribute| {
        if stated && matches!(verdict, Verdict::Differs(_)) {
            Verdict::Unverifiable(attribute)
        } else {
            verdict
        }
    };
    let verdicts = [
        hyperlink_comparison.unwrap_or(Verdict::Same),
        bullet_assets.as_ref().map_or(Verdict::Same, |assets| {
            same_bullet_fallback(binary.element, &shape, assets)
        }),
        if binary_freeform_preset || local.is_some_and(|p| !p.geometry) {
            Verdict::Same
        } else {
            local_override(
                same_geometry(binary, &shape),
                parsed.placeholder,
                "placeholder geometry precedence",
            )
        },
        if parsed.placeholder {
            Verdict::Same
        } else {
            same_transform(binary.leaf, binary.nested, &shape)
        },
        if local.is_some_and(|p| !p.fill) {
            Verdict::Same
        } else {
            local_override(
                image_comparison
                    .or(pattern_comparison)
                    .unwrap_or_else(|| same_fill(&binary.fill, &shape, binary.element.rotation)),
                parsed.placeholder,
                "placeholder fill precedence",
            )
        },
        same_run_formatting(
            binary.element,
            &shape,
            bullet_assets.as_ref().and(binary.direct_size_authored),
        ),
        match &substituted {
            Ok(_) => Verdict::Same,
            Err(verdict) => *verdict,
        },
    ];
    if !decide(&verdicts)? {
        return Ok(None);
    }
    shape = substituted.unwrap_or_else(|_| unreachable!("text agreement was decided"));
    if let Some(assets) = bullet_assets {
        let mut retained = std::collections::BTreeMap::new();
        for (path, extension, bytes) in assets {
            let id = media.admit_blob_image(extension, bytes)?;
            retained.insert(path, format!("legacy-ppt/image/{id}"));
        }
        let body = shape
            .text_body
            .as_mut()
            .ok_or_else(|| unverifiable("picture bullet"))?;
        for paragraph in &mut body.paragraphs {
            if let pptx_model::Bullet::Blip { image_path, .. } = &mut paragraph.bullet {
                *image_path = retained
                    .get(image_path)
                    .ok_or_else(|| unverifiable("picture bullet"))?
                    .clone();
            }
        }
    }
    let element = binary.element;
    if image_comparison == Some(Verdict::Same) {
        if let (
            Some(pptx_model::Fill::Image {
                image_path,
                mime_type,
                ..
            }),
            Some(pptx_model::Fill::Image {
                image_path: binary_path,
                mime_type: binary_mime,
                ..
            }),
        ) = (&mut shape.fill, &element.fill)
        {
            *image_path = binary_path.clone();
            *mime_type = binary_mime.clone();
        }
    }
    if binary_freeform_preset || local.is_some_and(|p| !p.geometry) {
        shape.geometry = element.geometry.clone();
        shape.cust_geom = element.cust_geom.clone();
        shape.cust_geom_paint = element.cust_geom_paint.clone();
        shape.adj = element.adj;
        shape.adj2 = element.adj2;
        shape.adj3 = element.adj3;
        shape.adj4 = element.adj4;
        shape.adj5 = element.adj5;
        shape.adj6 = element.adj6;
        shape.adj7 = element.adj7;
        shape.adj8 = element.adj8;
    } else if shape.geometry != "custGeom" {
        // PowerPoint 16 binary-only controls for an adjusted upArrow (the
        // alternative has an empty avLst) write the binary guide on save;
        // changing that binary value changes the PDF. An omitted XML guide
        // leaves that guide to the binary shape, even when the preset's
        // DrawingML default is unknown or the binary guide is non-default.
        for (xml, binary) in [
            (&mut shape.adj, element.adj),
            (&mut shape.adj2, element.adj2),
            (&mut shape.adj3, element.adj3),
            (&mut shape.adj4, element.adj4),
            (&mut shape.adj5, element.adj5),
            (&mut shape.adj6, element.adj6),
            (&mut shape.adj7, element.adj7),
            (&mut shape.adj8, element.adj8),
        ] {
            if xml.is_none() {
                *xml = binary;
            }
        }
    }
    if let Some(local) = local {
        // ECMA-376 Part 1 Annex L.3.2.3: a placeholder takes absent shape
        // properties from its layout/master. The standalone DrawingML part
        // has no layout; the binary projection already holds those values.
        // PowerPoint 16 controls with title/body/subtitle placeholders show
        // XML-only font and bold edits but binary-anchor position/size. An
        // XML-only local fill or geometry edit is visible; a binary edit can
        // invalidate the blob, so explicit values still pass the ordinary
        // agreement gate above. Do not infer missing local values from the
        // standalone parser's defaults.
        if !local.fill {
            shape.fill = element.fill.clone();
        }
        if !local.stroke {
            shape.stroke = element.stroke.clone();
        }
        if !local.effects {
            shape.shadow = element.shadow.clone();
            shape.inner_shadow = element.inner_shadow.clone();
            shape.glow = element.glow.clone();
            shape.soft_edge = element.soft_edge.clone();
            shape.reflection = element.reflection.clone();
        }
    }
    shape.x = element.x;
    shape.y = element.y;
    shape.width = element.width;
    shape.height = element.height;
    shape.rotation = element.rotation;
    shape.flip_h = element.flip_h;
    shape.flip_v = element.flip_v;
    shape.id = element.id.clone();
    let retained = usize::try_from(
        ooxml_common::json_measurement::measure_json(&shape)
            .map_err(unsupported)?
            .json_bytes,
    )
    .map_err(|_| unsupported("PowerPoint direct slide model budget exceeded"))?;
    *model_budget = model_budget
        .checked_sub(retained)
        .ok_or_else(|| unsupported("PowerPoint direct slide model budget exceeded"))?;
    Ok(Some(shape))
}

#[cfg(feature = "direct-ppt")]
#[derive(Clone, Copy)]
struct PlaceholderLocals {
    geometry: bool,
    fill: bool,
    stroke: bool,
    effects: bool,
}

/// A standalone placeholder parser supplies schema defaults where the actual
/// shape relies on a missing layout. Inspect direct `p:spPr` components and
/// authored `p:style` references to distinguish local style from inherited
/// layout properties (ECMA-376 Annex L.3.2.3). The archive and part sizes
/// have already passed `inflated_size` and the PPTX
/// standalone parser's bounds; this second read is limited to placeholders.
#[cfg(feature = "direct-ppt")]
fn placeholder_locals(blob: &[u8], part: &str, theme: &Theme) -> Result<PlaceholderLocals, String> {
    use ooxml_common::ns::{is_a_ns, is_p_ns};
    use std::io::Read;
    let unreadable = || unsupported("unreadable PowerPoint alternative shape XML placeholder");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).map_err(|_| unreadable())?;
    let index = crate::opc_part::entry_index(&archive, part).ok_or_else(unreadable)?;
    let entry = archive.by_index(index).map_err(|_| unreadable())?;
    if entry.size() > MAX_PART_BYTES || entry.encrypted() {
        return Err(unreadable());
    }
    let mut xml = String::new();
    entry
        .take(MAX_PART_BYTES + 1)
        .read_to_string(&mut xml)
        .map_err(|_| unreadable())?;
    if xml.len() as u64 > MAX_PART_BYTES {
        return Err(unreadable());
    }
    let doc = roxmltree::Document::parse(&xml).map_err(|_| unreadable())?;
    let sp_pr = doc.root_element().children().find(|n| {
        n.is_element() && n.tag_name().name() == "spPr" && is_p_ns(n.tag_name().namespace())
    });
    let style = doc.root_element().children().find(|n| {
        n.is_element() && n.tag_name().name() == "style" && is_p_ns(n.tag_name().namespace())
    });
    let has = |names: &[&str]| {
        sp_pr.is_some_and(|sp_pr| {
            sp_pr.children().any(|n| {
                n.is_element()
                    && is_a_ns(n.tag_name().namespace())
                    && names.contains(&n.tag_name().name())
            })
        })
    };
    let style_has = |name| {
        style.is_some_and(|style| {
            style.children().any(|n| {
                n.is_element() && is_a_ns(n.tag_name().namespace()) && n.tag_name().name() == name
            })
        })
    };
    if let Some(style) = style {
        let format = theme
            .format_scheme()
            .ok_or_else(|| unverifiable("placeholder style theme"))?;
        for reference in style
            .children()
            .filter(|node| node.is_element() && is_a_ns(node.tag_name().namespace()))
        {
            let selected = match reference.tag_name().name() {
                "fillRef" => format.lookup_fill_ref(
                    reference
                        .attribute("idx")
                        .and_then(|idx| idx.parse().ok())
                        .ok_or_else(|| unverifiable("fill style"))?,
                ),
                "lnRef" => format.lookup_line_ref(
                    reference
                        .attribute("idx")
                        .and_then(|idx| idx.parse().ok())
                        .ok_or_else(|| unverifiable("placeholder line style"))?,
                ),
                "effectRef" => format.lookup_effect_ref(
                    reference
                        .attribute("idx")
                        .and_then(|idx| idx.parse().ok())
                        .ok_or_else(|| unverifiable("placeholder effect style"))?,
                ),
                _ => continue,
            };
            if matches!(selected, ooxml_common::theme::StyleMatrixLookup::Missing) {
                return Err(unverifiable("placeholder style reference"));
            }
        }
    }
    Ok(PlaceholderLocals {
        geometry: has(&["prstGeom", "custGeom"]),
        fill: has(&[
            "noFill",
            "solidFill",
            "gradFill",
            "pattFill",
            "blipFill",
            "grpFill",
        ]),
        stroke: has(&["ln"]) || style_has("lnRef"),
        // PowerPoint-saved placeholders with only an effectRef retain their
        // theme-resolved shadow. The standalone parser resolves that style;
        // replacing it with a shadowless binary fallback drops that shadow.
        effects: has(&["effectLst", "effectDag"]) || style_has("effectRef"),
    })
}

/// Without the direct presentation model feature there is no DrawingML
/// parser, so no alternative can be verified (only the feature-less unit
/// tests compile the direct model).
#[cfg(all(test, not(feature = "direct-ppt")))]
pub(in crate::ppt) fn adopt(
    _binary: &BinaryShape<'_>,
    _blob: &[u8],
    _theme: &Theme,
    _media: &mut super::media::SpanStore,
    _work_budget: &mut usize,
    _text_budget: &mut usize,
    _model_budget: &mut usize,
) -> Result<Option<ShapeElement>, String> {
    Err(unverifiable("part"))
}

/// Sum of the package entries' declared uncompressed sizes, bounded by the
/// package policy. The PPTX package reader enforces the same limits while
/// inflating.
#[cfg(feature = "direct-ppt")]
fn inflated_size(blob: &[u8]) -> Result<usize, String> {
    let unreadable = || unsupported("unreadable PowerPoint alternative shape XML package");
    let oversized = || unsupported("PowerPoint alternative shape XML exceeds its size limit");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).map_err(|_| unreadable())?;
    let mut total = 0u64;
    for index in 0..archive.len() {
        let entry = archive.by_index_raw(index).map_err(|_| unreadable())?;
        if entry.size() > MAX_PART_BYTES {
            return Err(oversized());
        }
        total = total.checked_add(entry.size()).ok_or_else(oversized)?;
    }
    if total > MAX_PACKAGE_BYTES {
        return Err(oversized());
    }
    usize::try_from(total).map_err(|_| oversized())
}

#[cfg(any(test, feature = "direct-ppt"))]
fn adjusts(shape: &ShapeElement) -> [Option<f64>; 8] {
    [
        shape.adj, shape.adj2, shape.adj3, shape.adj4, shape.adj5, shape.adj6, shape.adj7,
        shape.adj8,
    ]
}

/// Presets compare by name and adjust values (both in ECMA-376 1/100000
/// units), custom geometry path by path. A binary shape type maps to one
/// preset (`officeart::preset`), so two different preset names are
/// different shapes. PowerPoint saves a preset that has no MS-ODRAW shape
/// type as a binary freeform: a preset on one side and custom geometry on
/// the other would need the preset's formulas evaluated to compare. `adopt`
/// handles the observed binary-freeform/XML-preset precedence before calling
/// this comparator; the converse still reaches its unverifiable result.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_geometry(binary: &BinaryShape<'_>, alternative: &ShapeElement) -> Verdict {
    let element = binary.element;
    match (
        element.geometry == "custGeom",
        alternative.geometry == "custGeom",
    ) {
        (true, true) => return same_paths(binary, alternative),
        (false, false) if element.geometry != alternative.geometry => {
            return Verdict::Differs("preset geometry");
        }
        (false, false) => {}
        _ => {
            return Verdict::Unverifiable("preset geometry");
        }
    }
    // PowerPoint 16 saves/PDFs of an upArrow whose XML avLst is empty retain
    // its adjusted binary guide, including after a binary-only edit. Omission
    // therefore does not assert the DrawingML default against a stated binary
    // guide. Two stated guides are compared within the rounding of the
    // binary anchor; when only the XML states a guide, the normative preset
    // default is the binary side of the comparison when defined.
    let defaults = crate::officeart::preset_defaults::defaults(&element.geometry);
    let mut verdict = Verdict::Same;
    for (index, (bounds, value)) in binary
        .adjust_bounds
        .iter()
        .zip(adjusts(alternative))
        .enumerate()
    {
        let default = defaults.and_then(|d| d.get(index)).map(|&v| f64::from(v));
        match (bounds, value) {
            (_, None) => {}
            (Some((low, high)), Some(value)) => {
                if value < low - 1.0 || high + 1.0 < value {
                    return Verdict::Differs("preset adjust values");
                }
            }
            (None, Some(value)) => match default {
                Some(default) if (value - default).abs() <= 1.0 => {}
                Some(_) => return Verdict::Differs("preset adjust values"),
                None => verdict = Verdict::Unverifiable("preset adjust values"),
            },
        }
    }
    verdict
}

/// Custom geometry: both sides hold ECMA-376 20.1.9.14 path commands
/// normalized by their path extents (the binary's MS-ODRAW geoLeft..geoRight
/// and geoTop..geoBottom, the alternative's `a:path` w and h). A closing line
/// back to a subpath's start point immediately before its close draws
/// nothing that the close does not (20.1.9.3), and MS-ODRAW freeforms store
/// it explicitly while the alternative omits it, so it is dropped before
/// comparing. Paths, their paint and their command kinds must then match
/// one to one; a different vertex is a different shape, while a different
/// command structure may be another encoding of the same outline.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_paths(binary: &BinaryShape<'_>, alternative: &ShapeElement) -> Verdict {
    use pptx_model::PathCmd;
    let (Some(b), Some(a), Some(b_paint)) = (
        &binary.element.cust_geom,
        &alternative.cust_geom,
        &binary.path_paint,
    ) else {
        return Verdict::Unverifiable("custom geometry");
    };
    if b.len() != a.len() || b_paint.len() != b.len() {
        return Verdict::Unverifiable("custom geometry");
    }
    let mut verdict = Verdict::Same;
    for (index, (bp, ap)) in b.iter().zip(a).enumerate() {
        let a_paint = match alternative.cust_geom_paint.as_ref().map(|p| p.get(index)) {
            None => (true, true),
            Some(Some(p)) => match p.fill.as_deref() {
                None => (true, p.stroke),
                Some("none") => (false, p.stroke),
                // A lightened or darkened fill has no MS-ODRAW path flag.
                Some(_) => return Verdict::Unverifiable("custom geometry path fill"),
            },
            Some(None) => return Verdict::Unverifiable("custom geometry"),
        };
        if b_paint[index] != a_paint {
            verdict = Verdict::Differs("custom geometry path paint");
        }
        let (bp, ap) = (without_closing_lines(bp), without_closing_lines(ap));
        if bp.len() != ap.len() {
            return Verdict::Unverifiable("custom geometry");
        }
        for (bc, ac) in bp.iter().zip(&ap) {
            let pairs = match (bc, ac) {
                (PathCmd::MoveTo { x, y }, PathCmd::MoveTo { x: ax, y: ay })
                | (PathCmd::LineTo { x, y }, PathCmd::LineTo { x: ax, y: ay }) => {
                    vec![(x, ax), (y, ay)]
                }
                (
                    PathCmd::CubicBezTo {
                        x1,
                        y1,
                        x2,
                        y2,
                        x,
                        y,
                    },
                    PathCmd::CubicBezTo {
                        x1: ax1,
                        y1: ay1,
                        x2: ax2,
                        y2: ay2,
                        x: ax,
                        y: ay,
                    },
                ) => vec![(x1, ax1), (y1, ay1), (x2, ax2), (y2, ay2), (x, ax), (y, ay)],
                (
                    PathCmd::ArcTo {
                        wr,
                        hr,
                        st_ang,
                        sw_ang,
                    },
                    PathCmd::ArcTo {
                        wr: awr,
                        hr: ahr,
                        st_ang: ast_ang,
                        sw_ang: asw_ang,
                    },
                ) => vec![(wr, awr), (hr, ahr), (st_ang, ast_ang), (sw_ang, asw_ang)],
                (PathCmd::Close, PathCmd::Close) => Vec::new(),
                _ => return Verdict::Unverifiable("custom geometry"),
            };
            if pairs.iter().any(|(p, q)| (*p - *q).abs() > PATH_TOLERANCE) {
                verdict = Verdict::Differs("custom geometry vertices");
            }
        }
    }
    verdict
}

/// The path without a line to its subpath's start point directly before a
/// close.
#[cfg(any(test, feature = "direct-ppt"))]
fn without_closing_lines(path: &[pptx_model::PathCmd]) -> Vec<&pptx_model::PathCmd> {
    use pptx_model::PathCmd;
    let mut result: Vec<&PathCmd> = Vec::with_capacity(path.len());
    let mut start = None;
    for command in path {
        match command {
            PathCmd::MoveTo { x, y } => start = Some((*x, *y)),
            PathCmd::Close => {
                if let (Some(PathCmd::LineTo { x, y }), Some((sx, sy))) = (result.last(), start) {
                    if (x - sx).abs() <= PATH_TOLERANCE && (y - sy).abs() <= PATH_TOLERANCE {
                        result.pop();
                    }
                }
            }
            _ => {}
        }
        result.push(command);
    }
    result
}

/// Both sides are compared in the binary anchor's own unit, where one unit
/// is its resolution: master units (1/576 inch; the alternative states EMU)
/// for a top-level shape, and the group's child coordinate units for a group
/// child (MS-ODRAW OfficeArtChildAnchor, which the alternative states as the
/// same unscaled `a:chOff`/`a:chExt` space values). The binary transform was
/// converted to EMU with the master-unit scale either way.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_transform(
    leaf: &pptx_model::Transform,
    nested: bool,
    alternative: &ShapeElement,
) -> Verdict {
    let alternative_unit = if nested { 1.0 } else { EMU_PER_MASTER_UNIT };
    let close = |a: i64, b: i64| {
        (a as f64 / EMU_PER_MASTER_UNIT - b as f64 / alternative_unit).abs() <= XFRM_TOLERANCE_UNITS
    };
    let turn = (leaf.rot - alternative.rotation).rem_euclid(360.0);
    let same = close(leaf.x, alternative.x)
        && close(leaf.y, alternative.y)
        && close(leaf.cx, alternative.width)
        && close(leaf.cy, alternative.height)
        && turn.min(360.0 - turn) < 0.01
        && leaf.flip_h == alternative.flip_h
        && leaf.flip_v == alternative.flip_v;
    if same {
        Verdict::Same
    } else {
        Verdict::Differs("transform")
    }
}

/// Solid colors compare exactly (RGB with any alpha suffix). A fill that is
/// off on one side and solid on the other is a different paint. Other
/// fill kinds on either side have no common model representation this
/// reader can equate (PowerPoint may restate a binary fill as another
/// DrawingML kind), except comparable gradient and pattern projections.
/// PowerPoint 16 controls show XML pattern edits and binary image edits both
/// affect PDF output. `same_binary_pattern` compares the image/pattern pair
/// using its decoded 8x8 bitmap and duotone endpoints. Rotated-gradient
/// omission is handled by `same_paint` below.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_fill(binary: &RecordedFill, alternative: &ShapeElement, rotation: f64) -> Verdict {
    use pptx_model::Fill;
    let paint = |fill: Option<&Fill>| match fill {
        None | Some(Fill::None) => None,
        Some(fill) => Some(fill.clone()),
    };
    let alternative = paint(alternative.fill.as_ref());
    let binary = match binary {
        RecordedFill::NotDisplayed => return Verdict::Same,
        RecordedFill::Unstated if alternative.is_none() => return Verdict::Same,
        RecordedFill::Unstated | RecordedFill::Unknown => return Verdict::Unverifiable("fill"),
        RecordedFill::Stated(fill) => paint(fill.as_deref()),
    };
    match (&binary, &alternative) {
        (None, None) => Verdict::Same,
        (Some(Fill::Solid { color: a }), Some(Fill::Solid { color: b })) => {
            match same_color(a, b) {
                Some(true) => Verdict::Same,
                Some(false) => Verdict::Differs("solid fill color"),
                None => Verdict::Unverifiable("fill"),
            }
        }
        (None, Some(Fill::Solid { .. })) | (Some(Fill::Solid { .. }), None) => {
            Verdict::Differs("fill")
        }
        (Some(a), Some(b)) if same_paint(a, b, rotation) => Verdict::Same,
        _ => Verdict::Unverifiable("fill"),
    }
}

/// Model colors are hex RGB with an optional alpha byte. The binary stores
/// the displayed RGB, while an alternative color may be a theme color with
/// transforms (ECMA-376 20.1.2.3) that this reader resolves itself; two
/// roundings of one computed channel differ by at most one unit (a corpus
/// alternative's 50% gray resolves to 0x80 where the binary stores 0x7F).
/// Alpha converts from 16.16 (MS-ODRAW) and 1/100000 (DrawingML) alike.
/// `None`: not a model color.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_color(binary: &str, alternative: &str) -> Option<bool> {
    let channels = |color: &str| -> Option<Vec<i16>> {
        if !matches!(color.len(), 6 | 8) || !color.is_ascii() {
            return None;
        }
        (0..color.len())
            .step_by(2)
            .map(|at| i16::from_str_radix(&color[at..at + 2], 16).ok())
            .collect()
    };
    let (mut b, mut a) = (channels(binary)?, channels(alternative)?);
    // An absent alpha byte is opaque.
    b.resize(4, 0xff);
    a.resize(4, 0xff);
    Some(b.iter().zip(&a).all(|(x, y)| (x - y).abs() <= 1))
}

/// Comparable gradient or pattern paint. Gradient stop positions are 16.16
/// fractions in the binary (MS-ODRAW 2.3.7.17 fillShadeColors) and
/// 1/100000 in DrawingML (ECMA-376 20.1.8.36), so each may round by one
/// step of either unit; angles are whole degrees on both sides. PowerPoint 16
/// uses the XML in Office-saved multi-stop gradients whose binary duplicates
/// the first colour at position zero and synthesizes a final stop from the
/// scalar back colour. The XML omits the redundant first stop and retains the
/// terminal colour at a position that the binary does not encode. Controls
/// with three and five binary stops, two and four XML stops, and changed
/// XML/binary colours establish this restatement; compare the authored
/// interior positions and colours, and the terminal colour, without inventing
/// a terminal position for the binary. Other unequal layouts stay unverifiable.
/// ECMA-376 CT_GradientFillProperties gives `flip` the default `none`; a
/// `tileRect` with all zero offsets covers the whole shape, like no tileRect.
/// PowerPoint 16 PDF controls on 90- and 180-degree rotations show that an
/// omitted XML `rotWithShape` renders like `1`, while the corresponding binary
/// shapes carry no authored rotation flag and project the binary default `0`.
/// This paired omission is a PowerPoint restatement, not visual equivalence of
/// the two projections. An XML-only `rotWithShape="0"` edit changes both PDFs.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_paint(binary: &pptx_model::Fill, alternative: &pptx_model::Fill, rotation: f64) -> bool {
    use pptx_model::Fill;
    const POSITION: f64 = 1.0 / 65536.0 + 1.0 / 100_000.0;
    fn sorted(stops: &[pptx_model::GradStop]) -> Vec<(f64, &str)> {
        let mut stops: Vec<(f64, &str)> = stops
            .iter()
            .map(|s| (s.position, s.color.as_str()))
            .collect();
        stops.sort_by(|a, b| a.0.total_cmp(&b.0));
        stops
    }
    fn comparable_stops(binary: &[(f64, &str)], alternative: &[(f64, &str)]) -> bool {
        let matching = |b: &(f64, &str), a: &(f64, &str)| {
            (b.0 - a.0).abs() <= POSITION && same_color(b.1, a.1) == Some(true)
        };
        if binary.len() == alternative.len() {
            return binary.iter().zip(alternative).all(|(b, a)| matching(b, a));
        }
        // The 0 stop duplicates the first authored colour; the final binary
        // stop is synthesized from fillBackColor at 1, not authored in
        // fillShadeColors. PowerPoint's XML retains its independent terminal
        // position, which must follow the last binary-authored stop.
        if binary.len() != alternative.len() + 1 || binary.len() < 3 {
            return false;
        }
        let redundant_start =
            binary[0].0.abs() <= POSITION && same_color(binary[0].1, binary[1].1) == Some(true);
        let authored = binary[1..binary.len() - 1]
            .iter()
            .zip(&alternative[..alternative.len() - 1])
            .all(|(b, a)| matching(b, a));
        let end = (binary.last().unwrap(), alternative.last().unwrap());
        redundant_start
            && authored
            && (end.0 .0 - 1.0).abs() <= POSITION
            && end.1 .0 + POSITION >= binary[binary.len() - 2].0
            && end.1 .0 <= 1.0 + POSITION
            && same_color(end.0 .1, end.1 .1) == Some(true)
    }
    let rect = |r: &Option<ooxml_common::fill::FillRect>| r.as_ref().map(|r| [r.l, r.t, r.r, r.b]);
    let tile_rect =
        |r: &Option<ooxml_common::fill::FillRect>| rect(r).unwrap_or([0.0, 0.0, 0.0, 0.0]);
    let rotation_irrelevant = rotation.rem_euclid(360.0) == 0.0;
    match (binary, alternative) {
        (
            Fill::Gradient {
                stops: bs,
                angle: ba,
                grad_type: bt,
                scaled: bsc,
                path: bp,
                fill_to_rect: bf,
                tile_rect: btr,
                flip: bfl,
                rot_with_shape: br,
            },
            Fill::Gradient {
                stops: as_,
                angle: aa,
                grad_type: at,
                scaled: asc,
                path: ap,
                fill_to_rect: af,
                tile_rect: atr,
                flip: afl,
                rot_with_shape: ar,
            },
        ) => {
            let (bs, as_) = (sorted(bs), sorted(as_));
            let turn = (ba - aa).rem_euclid(360.0);
            comparable_stops(&bs, &as_)
                && turn.min(360.0 - turn) < 0.01
                && bt == at
                && bsc == asc
                && bp == ap
                && rect(bf) == rect(af)
                && tile_rect(btr) == tile_rect(atr)
                && bfl.as_deref().unwrap_or("none") == afl.as_deref().unwrap_or("none")
                && (br == ar || ar.is_none() && *br == Some(false) || rotation_irrelevant)
        }
        (
            Fill::Pattern {
                fg: bf,
                bg: bb,
                preset: bp,
            },
            Fill::Pattern {
                fg: af,
                bg: ab,
                preset: ap,
            },
        ) => same_color(bf, af) == Some(true) && same_color(bb, ab) == Some(true) && bp == ap,
        _ => false,
    }
}

/// PowerPoint 16 picture-bullet controls with two distinct embedded GIFs
/// saved a plain U+2022 OfficeArt bullet as their down-revision fallback.
/// XML-only image replacement changed the PDF, while changing that binary
/// fallback to U+25A0 produced exactly the binary-only PDF. Thus only the
/// observed U+2022 fallback is compatible with an XML buBlip; a different
/// binary marker invalidates the alternative. This applies to picture bullet
/// references only, with every target resolved and retained above.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_bullet_fallback(
    binary: &ShapeElement,
    alternative: &ShapeElement,
    assets: &[BlobBulletAsset],
) -> Verdict {
    let (Some(b), Some(a)) = (&binary.text_body, &alternative.text_body) else {
        return Verdict::Unverifiable("picture bullet");
    };
    if b.paragraphs.len() != a.paragraphs.len() {
        return Verdict::Unverifiable("picture bullet");
    }
    let mut found = false;
    for (bp, ap) in b.paragraphs.iter().zip(&a.paragraphs) {
        if let pptx_model::Bullet::Blip { image_path, .. } = &ap.bullet {
            found = true;
            if !assets.iter().any(|(path, _, _)| path == image_path) {
                return Verdict::Unverifiable("picture bullet resource");
            }
            match &bp.bullet {
                pptx_model::Bullet::Char { ch, .. } if ch == "•" => {}
                _ => return Verdict::Differs("picture bullet fallback"),
            }
        }
    }
    if found {
        Verdict::Same
    } else {
        Verdict::Unverifiable("picture bullet")
    }
}

/// [MS-PPT] 2.6.10/2.9.57 text click ranges refer to the document's
/// ExHyperlinkContainer targets. PowerPoint 16 ignores an XML-only r:id URL
/// edit in a saved legacy deck, so a differing target selects binary. Compare
/// per UTF-16 position because the encodings may split style runs differently.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_hyperlinks(binary: &ShapeElement, alternative: &ShapeElement) -> Verdict {
    if binary.hyperlink != alternative.hyperlink
        || binary.hyperlink_action != alternative.hyperlink_action
    {
        return Verdict::Differs("shape hyperlink");
    }
    type Link<'a> = (Option<&'a str>, Option<&'a str>);
    let (Some(b), Some(a)) = (&binary.text_body, &alternative.text_body) else {
        return if binary.text_body.is_none() && alternative.text_body.is_none() {
            Verdict::Same
        } else {
            Verdict::Unverifiable("text hyperlink")
        };
    };
    if b.paragraphs.len() != a.paragraphs.len() {
        return Verdict::Differs("text hyperlink");
    }
    fn spans(runs: &[TextRun]) -> Option<impl Iterator<Item = (usize, Link<'_>)>> {
        if runs.iter().any(|run| matches!(run, TextRun::Math { .. })) {
            return None;
        }
        Some(
            runs.iter()
                .map(|run| match run {
                    TextRun::Text(data) => (
                        data.text.encode_utf16().count(),
                        (data.hyperlink.as_deref(), data.hyperlink_action.as_deref()),
                    ),
                    _ => (1, (None, None)),
                })
                .filter(|(len, _)| *len > 0),
        )
    }
    for (bp, ap) in b.paragraphs.iter().zip(&a.paragraphs) {
        let (Some(mut bx), Some(mut ax)) = (spans(&bp.runs), spans(&ap.runs)) else {
            return Verdict::Unverifiable("text hyperlink");
        };
        let (mut bs, mut aspan) = (bx.next(), ax.next());
        // Compare each run overlap once. Expanding per UTF-16 position would
        // allocate large scratch arrays and repeatedly compare long URLs.
        while let (Some((bn, bl)), Some((an, al))) = (bs, aspan) {
            if bl != al {
                return Verdict::Differs("text hyperlink");
            }
            let common = bn.min(an);
            bs = if bn == common {
                bx.next()
            } else {
                Some((bn - common, bl))
            };
            aspan = if an == common {
                ax.next()
            } else {
                Some((an - common, al))
            };
        }
        if bs.is_some() || aspan.is_some() {
            return Verdict::Differs("text hyperlink");
        }
    }
    Verdict::Same
}

/// Run formatting both forms state must agree character by character: font
/// size, bold and italic, where each side specifies them. The binary stores
/// whole points only; a corpus deck whose alternative says 10.5 pt over a
/// binary 10 pt renders in PowerPoint without the alternative's other run
/// properties (its character spacing), while a deck whose sizes agree
/// renders them. A field compares as one unit, because the binary
/// projection substitutes its displayed text.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_run_formatting(
    binary: &ShapeElement,
    alternative: &ShapeElement,
    direct_size_authored: Option<&[bool]>,
) -> Verdict {
    #[derive(Clone, Copy, PartialEq)]
    enum Unit {
        Character,
        Field,
        Break,
    }
    type Format = (Option<f64>, Option<bool>, Option<bool>);
    let (Some(b), Some(a)) = (&binary.text_body, &alternative.text_body) else {
        return Verdict::Same;
    };
    if b.paragraphs.len() != a.paragraphs.len() {
        return Verdict::Differs("paragraph structure");
    }
    fn units(runs: &[TextRun]) -> Option<impl Iterator<Item = (Unit, Format)> + Clone + '_> {
        if runs.iter().any(|run| matches!(run, TextRun::Math { .. })) {
            return None;
        }
        // Stream the logical positions rather than allocating a multi-word
        // tuple for every character. The text budget bounds traversal work.
        Some(runs.iter().flat_map(|run| {
            let (unit, format, count) = match run {
                TextRun::Text(d) => (
                    (if d.field_type.is_some() {
                        Unit::Field
                    } else {
                        Unit::Character
                    }),
                    (d.font_size, d.bold, d.italic),
                    if d.field_type.is_some() {
                        1
                    } else {
                        d.text.encode_utf16().count()
                    },
                ),
                _ => (Unit::Break, (None, None, None), 1),
            };
            std::iter::repeat_n((unit, format), count)
        }))
    }
    let agree = |x: Option<f64>, y: Option<f64>| match (x, y) {
        (Some(x), Some(y)) => (x - y).abs() < 1e-6,
        _ => true,
    };
    let agree_bool = |x: Option<bool>, y: Option<bool>| x.zip(y).is_none_or(|(x, y)| x == y);
    let mut verdict = Verdict::Same;
    let mut position = 0;
    for (bp, ap) in b.paragraphs.iter().zip(&a.paragraphs) {
        let (Some(bx), Some(ax)) = (units(&bp.runs), units(&ap.runs)) else {
            return Verdict::Unverifiable("equation");
        };
        let binary_count = bx.clone().count();
        if direct_size_authored.is_some_and(|mask| position + binary_count > mask.len()) {
            return Verdict::Unverifiable("run size origin");
        }
        // A masked alternative states a line break as one text character
        // (see `substitute_text`); its formatting still aligns.
        let kind = |unit: Unit| {
            if unit == Unit::Break {
                Unit::Character
            } else {
                unit
            }
        };
        if binary_count != ax.clone().count()
            || bx
                .clone()
                .zip(ax.clone())
                .any(|(x, y)| kind(x.0) != kind(y.0))
        {
            // `substitute_text` decides the text structure.
            position += binary_count + 1;
            continue;
        }
        if !bx.zip(ax).enumerate().all(|(index, (x, y))| {
            let size_stated = direct_size_authored.is_none_or(|mask| mask[position + index]);
            (!size_stated || agree(x.1 .0, y.1 .0))
                && agree_bool(x.1 .1, y.1 .1)
                && agree_bool(x.1 .2, y.1 .2)
        }) {
            verdict = Verdict::Differs("run formatting");
        }
        position += binary_count + 1;
    }
    verdict
}

/// Replace the masked characters of the alternative text with the binary
/// characters. Every paragraph, run and line break must line up exactly at
/// UTF-16 lengths; a run boundary inside a surrogate pair does not.
///
/// XML 1.0 cannot carry the vertical tab (U+000B) that breaks a line inside
/// a binary paragraph (MS-PPT 2.9.43), and the masked alternative states each
/// one as a masked character of the enclosing run, not as `a:br` (every
/// corpus alternative; both runs of one character and breaks leading a
/// longer run). Each such character becomes a line break that splits its run
/// into runs of the same formatting. Other break characters inside a run
/// have no observed form and are not compared. An equation has no binary
/// characters to compare.
#[cfg(any(test, feature = "direct-ppt"))]
#[cfg(test)]
fn substitute_text(shape: &mut ShapeElement, text: Option<&str>) -> Result<(), Verdict> {
    let mut budget = usize::MAX;
    substitute_text_bounded(shape, text, &mut budget)
}

fn substitute_text_bounded(
    shape: &mut ShapeElement,
    text: Option<&str>,
    model_budget: &mut usize,
) -> Result<(), Verdict> {
    const DIFFERS: Verdict = Verdict::Differs("text structure");
    let Some(body) = shape.text_body.as_mut() else {
        return if text.is_none_or(str::is_empty) {
            Ok(())
        } else {
            Err(DIFFERS)
        };
    };
    let text = text.unwrap_or("");
    let paragraphs: Vec<&str> = text.split('\r').collect();
    if paragraphs.len() != body.paragraphs.len() {
        // A body without text still has one empty paragraph in both forms.
        return Err(DIFFERS);
    }
    for (paragraph, source) in body.paragraphs.iter_mut().zip(paragraphs) {
        // Reserve scratch against the same cumulative model budget before
        // expanding text into UTF-16, independently of final JSON charging.
        let unit_count = source.encode_utf16().count();
        let scratch = unit_count
            .checked_mul(std::mem::size_of::<u16>())
            .ok_or(Verdict::Unverifiable("text model budget"))?;
        *model_budget = model_budget
            .checked_sub(scratch)
            .ok_or(Verdict::Unverifiable("text model budget"))?;
        let units: Vec<u16> = source.encode_utf16().collect();
        let mut at = 0usize;
        let mut runs = Vec::new();
        for run in std::mem::take(&mut paragraph.runs) {
            match run {
                TextRun::Text(mut data) => {
                    let len = data.text.encode_utf16().count();
                    let slice = at
                        .checked_add(len)
                        .and_then(|end| units.get(at..end))
                        .ok_or(DIFFERS)?;
                    at += len;
                    if slice.iter().any(|&u| matches!(u, 0x0a | 0x2028))
                        || (data.field_type.is_some() && slice.contains(&0x0b))
                    {
                        return Err(Verdict::Unverifiable("line break"));
                    }
                    // Clear masked text before cloning style so each fragment
                    // does not copy the entire original run. Charge retained
                    // style/character storage and run capacity before cloning.
                    data.text.clear();
                    let style_bytes = usize::try_from(
                        ooxml_common::json_measurement::measure_json(&data)
                            .map_err(|_| Verdict::Unverifiable("text model budget"))?
                            .json_bytes,
                    )
                    .map_err(|_| Verdict::Unverifiable("text model budget"))?;
                    let mut pieces = slice.split(|&u| u == 0x0b).peekable();
                    while let Some(piece) = pieces.next() {
                        if !piece.is_empty() {
                            let owned = piece
                                .len()
                                .checked_mul(3)
                                .and_then(|bytes| bytes.checked_add(style_bytes))
                                .ok_or(Verdict::Unverifiable("text model budget"))?;
                            *model_budget = model_budget
                                .checked_sub(owned)
                                .ok_or(Verdict::Unverifiable("text model budget"))?;
                            text_style::direct_model::reserve_run_slot(&mut runs, model_budget)
                                .map_err(|_| Verdict::Unverifiable("text model budget"))?;
                            let mut part = data.clone();
                            part.text = String::from_utf16(piece).map_err(|_| DIFFERS)?;
                            runs.push(TextRun::Text(part));
                        }
                        if pieces.peek().is_some() {
                            *model_budget = model_budget
                                .checked_sub(style_bytes)
                                .ok_or(Verdict::Unverifiable("text model budget"))?;
                            text_style::direct_model::reserve_run_slot(&mut runs, model_budget)
                                .map_err(|_| Verdict::Unverifiable("text model budget"))?;
                            runs.push(text_style::direct_model::line_break(Some(&data)));
                        }
                    }
                    if slice.is_empty() {
                        text_style::direct_model::reserve_run_slot(&mut runs, model_budget)
                            .map_err(|_| Verdict::Unverifiable("text model budget"))?;
                        runs.push(TextRun::Text(data));
                    }
                }
                line_break @ TextRun::Break { .. } => {
                    if !matches!(units.get(at), Some(0x0b | 0x0a | 0x2028)) {
                        return Err(DIFFERS);
                    }
                    at += 1;
                    text_style::direct_model::reserve_run_slot(&mut runs, model_budget)
                        .map_err(|_| Verdict::Unverifiable("text model budget"))?;
                    runs.push(line_break);
                }
                TextRun::Math { .. } => return Err(Verdict::Unverifiable("equation")),
            }
        }
        if at != units.len() {
            return Err(DIFFERS);
        }
        paragraph.runs = runs;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use pptx_model::{Fill, PathCmd};

    #[test]
    fn round_trip_theme_exposes_authored_hyperlink_colors() {
        let theme = Theme::Readable {
            theme_xml: "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:themeElements><a:clrScheme name=\"Office\"><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>".into(),
            clr_map: None,
            format_scheme: std::cell::OnceCell::new(),
        };
        assert_eq!(
            theme.hyperlink_colors(),
            Some(("0000FF".into(), "800080".into()))
        );
    }

    fn shape(geometry: &str) -> ShapeElement {
        serde_json::from_value(serde_json::json!({
            "x": 0, "y": 0, "width": 100, "height": 50, "rotation": 0.0,
            "flipH": false, "flipV": false, "geometry": geometry,
            "fill": null, "stroke": null, "textBody": null,
            "defaultTextColor": null, "custGeom": null,
            "adj": null, "adj2": null, "adj3": null, "adj4": null,
            "adj5": null, "adj6": null, "adj7": null, "adj8": null,
            "id": null, "name": null, "hyperlink": null,
            "placeholderType": null, "placeholderIdx": null
        }))
        .unwrap()
    }

    fn binary<'a>(element: &'a ShapeElement, leaf: &'a pptx_model::Transform) -> BinaryShape<'a> {
        BinaryShape {
            element,
            leaf,
            nested: false,
            text: None,
            direct_size_authored: None,
            fill: RecordedFill::Stated(element.fill.clone().map(Box::new)),
            path_paint: element
                .cust_geom
                .as_ref()
                .map(|paths| vec![(true, true); paths.len()]),
            adjust_bounds: adjusts(element).map(|v| v.map(|v| (v, v))),
            image_bytes: None,
        }
    }

    fn solid(color: &str) -> Option<Fill> {
        Some(Fill::Solid {
            color: color.to_owned(),
        })
    }

    #[test]
    fn a_verified_difference_keeps_the_binary_and_an_unknown_one_fails_closed() {
        use Verdict::*;
        assert_eq!(decide(&[Same, Same]), Ok(true));
        assert_eq!(
            decide(&[Same, Differs("fill"), Unverifiable("geometry")]),
            Ok(false)
        );
        let error = decide(&[Same, Unverifiable("fill")]).unwrap_err();
        assert!(
            error.starts_with("UNSUPPORTED:") && error.contains("fill"),
            "{error}"
        );
    }

    #[test]
    fn preset_adjusts_compare_within_the_rounding_of_their_binary_inputs() {
        let leaf = pptx_model::Transform::default();
        let element = shape("roundRect");
        let mut alternative = shape("roundRect");
        let mut shape = binary(&element, &leaf);
        assert_eq!(same_geometry(&shape, &alternative), Verdict::Same);
        // An omitted XML guide leaves the binary adjusted guide in force.
        shape.adjust_bounds[0] = Some((16660.0, 16670.0));
        assert_eq!(same_geometry(&shape, &alternative), Verdict::Same);
        shape.adjust_bounds[0] = Some((29000.0, 31000.0));
        assert_eq!(same_geometry(&shape, &alternative), Verdict::Same);
        alternative.adj = Some(31001.0);
        assert_eq!(same_geometry(&shape, &alternative), Verdict::Same);
        alternative.adj = Some(31002.0);
        assert_ne!(same_geometry(&shape, &alternative), Verdict::Same);
        // Different presets are different shapes; a preset against a
        // freeform would need the preset formulas evaluated.
        let other = self::shape("ellipse");
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &other),
            Verdict::Differs("preset geometry")
        );
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &self::shape("custGeom")),
            Verdict::Unverifiable("preset geometry")
        );
        // The preset definitions state no upArrow defaults (ECMA-376 lists
        // upDownArrow twice instead), but XML omission is still verifiable
        // because PowerPoint's binary-only control supplies the guide.
        let up = self::shape("upArrow");
        let mut shape = binary(&up, &leaf);
        shape.adjust_bounds[1] = Some((50000.0, 50000.0));
        assert_eq!(
            same_geometry(&shape, &self::shape("upArrow")),
            Verdict::Same
        );
    }

    #[test]
    fn custom_paths_compare_vertices_paint_and_structure() {
        let leaf = pptx_model::Transform::default();
        let square = |closing_line: bool, corner: f64| {
            let mut commands = vec![
                PathCmd::MoveTo { x: 0.0, y: 0.0 },
                PathCmd::LineTo { x: corner, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 1.0 },
            ];
            if closing_line {
                commands.push(PathCmd::LineTo { x: 0.0, y: 0.0 });
            }
            commands.push(PathCmd::Close);
            let mut element = shape("custGeom");
            element.cust_geom = Some(vec![commands]);
            element
        };
        // MS-ODRAW stores the line back to the start point explicitly.
        let element = square(true, 1.0);
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &square(false, 1.0)),
            Verdict::Same
        );
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &square(false, 0.9)),
            Verdict::Differs("custom geometry vertices")
        );
        let mut unfilled = binary(&element, &leaf);
        unfilled.path_paint = Some(vec![(false, true)]);
        assert_eq!(
            same_geometry(&unfilled, &square(false, 1.0)),
            Verdict::Differs("custom geometry path paint")
        );
        let mut curved = square(false, 1.0);
        if let Some(paths) = curved.cust_geom.as_mut() {
            paths[0][1] = PathCmd::QuadBezTo {
                x1: 0.5,
                y1: 0.0,
                x: 1.0,
                y: 0.0,
            };
        }
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &curved),
            Verdict::Unverifiable("custom geometry")
        );
    }

    #[test]
    fn transforms_compare_in_the_binary_anchor_unit() {
        let master = |units: f64| (units * EMU_PER_MASTER_UNIT).round() as i64;
        let leaf = pptx_model::Transform {
            x: master(100.0),
            y: master(200.0),
            cx: master(300.0),
            cy: master(400.0),
            ..Default::default()
        };
        let mut alternative = shape("rect");
        alternative.x = master(101.0);
        alternative.y = master(200.0);
        alternative.width = master(300.0);
        alternative.height = master(399.5);
        assert_eq!(same_transform(&leaf, false, &alternative), Verdict::Same);
        alternative.x = master(101.0) + 2;
        assert_eq!(
            same_transform(&leaf, false, &alternative),
            Verdict::Differs("transform")
        );
        alternative.x = master(100.0);
        alternative.rotation = 90.0;
        assert_ne!(same_transform(&leaf, false, &alternative), Verdict::Same);
        // Angles compare modulo a full turn.
        let turned = pptx_model::Transform {
            rot: -90.0,
            ..leaf.clone()
        };
        alternative.rotation = 270.0;
        assert_eq!(same_transform(&turned, false, &alternative), Verdict::Same);
        // A group child's alternative states the unscaled child units: one
        // child unit apart is within rounding, two are not.
        let mut child = shape("rect");
        (child.x, child.y, child.width, child.height) = (101, 200, 300, 400);
        assert_eq!(same_transform(&leaf, true, &child), Verdict::Same);
        child.x = 102;
        assert_eq!(
            same_transform(&leaf, true, &child),
            Verdict::Differs("transform")
        );
        // The same child values read as EMU are far off.
        child.x = 100;
        assert_eq!(
            same_transform(&leaf, false, &child),
            Verdict::Differs("transform")
        );
    }

    #[test]
    fn fills_compare_the_recorded_paint() {
        let mut alternative = shape("rect");
        assert_eq!(
            same_fill(&RecordedFill::Unstated, &alternative, 90.0),
            Verdict::Same
        );
        alternative.fill = solid("FFFFFF");
        assert_eq!(
            same_fill(&RecordedFill::Unstated, &alternative, 90.0),
            Verdict::Unverifiable("fill")
        );
        assert_eq!(
            same_fill(&RecordedFill::NotDisplayed, &alternative, 90.0),
            Verdict::Same
        );
        assert_eq!(
            same_fill(&RecordedFill::Stated(None), &alternative, 90.0),
            Verdict::Differs("fill")
        );
        // A color transform resolved independently may round one unit away.
        alternative.fill = solid("808080");
        assert_eq!(
            same_fill(
                &RecordedFill::Stated(solid("7f7f7f").map(Box::new)),
                &alternative,
                90.0
            ),
            Verdict::Same
        );
        assert_eq!(
            same_fill(
                &RecordedFill::Stated(solid("7E7F7F").map(Box::new)),
                &alternative,
                90.0
            ),
            Verdict::Differs("solid fill color")
        );
        assert_eq!(
            same_fill(
                &RecordedFill::Stated(solid("80808080").map(Box::new)),
                &alternative,
                90.0
            ),
            Verdict::Differs("solid fill color")
        );
        let gradient = |stops: &[(f64, &str)], rotate: Option<bool>| {
            Some(Fill::Gradient {
                stops: stops
                    .iter()
                    .map(|&(position, color)| pptx_model::GradStop {
                        position,
                        color: color.to_owned(),
                    })
                    .collect(),
                angle: 90.0,
                grad_type: "linear".to_owned(),
                scaled: Some(false),
                path: None,
                fill_to_rect: None,
                tile_rect: None,
                flip: None,
                rot_with_shape: rotate,
            })
        };
        let recorded = RecordedFill::Stated(
            gradient(&[(0.0, "000000"), (0.3099975, "FFFFFF")], Some(true)).map(Box::new),
        );
        alternative.fill = gradient(&[(0.31, "FFFFFF"), (0.0, "000000")], Some(true));
        assert_eq!(same_fill(&recorded, &alternative, 90.0), Verdict::Same);
        // PowerPoint's binary 0 stop may duplicate the first authored shade;
        // its scalar back colour has no authored stop position. The XML keeps
        // the latter position, including values short of 100%.
        let resampled = RecordedFill::Stated(
            gradient(
                &[
                    (0.0, "000000"),
                    (0.18, "000000"),
                    (0.39, "172DA6"),
                    (0.61, "00B0F0"),
                    (1.0, "20A472"),
                ],
                Some(false),
            )
            .map(Box::new),
        );
        alternative.fill = gradient(
            &[
                (0.18, "000000"),
                (0.39, "172DA6"),
                (0.61, "00B0F0"),
                (0.92, "20A472"),
            ],
            None,
        );
        assert_eq!(same_fill(&resampled, &alternative, 180.0), Verdict::Same);
        // The same restatement also occurs with three binary stops and two
        // XML stops. A distinct colour at zero is not the redundant stop.
        let short = RecordedFill::Stated(
            gradient(
                &[(0.0, "000000"), (0.31, "000000"), (1.0, "FFFFFF")],
                Some(false),
            )
            .map(Box::new),
        );
        alternative.fill = gradient(&[(0.31, "000000"), (0.88, "FFFFFF")], None);
        assert_eq!(same_fill(&short, &alternative, 90.0), Verdict::Same);
        let nonredundant = RecordedFill::Stated(
            gradient(
                &[(0.0, "172DA6"), (0.31, "000000"), (1.0, "FFFFFF")],
                Some(false),
            )
            .map(Box::new),
        );
        assert_eq!(
            same_fill(&nonredundant, &alternative, 90.0),
            Verdict::Unverifiable("fill")
        );
        alternative.fill = gradient(
            &[
                (0.18, "000000"),
                (0.39, "172DA6"),
                (0.61, "00B0F0"),
                (0.92, "20A472"),
            ],
            None,
        );
        if let Some(Fill::Gradient { stops, .. }) = alternative.fill.as_mut() {
            stops[1].position = 0.42;
        }
        assert_eq!(
            same_fill(&resampled, &alternative, 180.0),
            Verdict::Unverifiable("fill")
        );
        if let Some(Fill::Gradient { stops, .. }) = alternative.fill.as_mut() {
            stops[1].position = 0.39;
            stops[3].color = "FFFFFF".into();
        }
        assert_eq!(
            same_fill(&resampled, &alternative, 180.0),
            Verdict::Unverifiable("fill")
        );
        alternative.fill = gradient(&[(0.31, "FFFFFF"), (0.0, "000000")], Some(true));
        // ECMA-376 CT_GradientFillProperties defaults flip to "none";
        // tileRect with all-zero edges covers the same whole shape as none.
        if let Some(Fill::Gradient {
            tile_rect, flip, ..
        }) = alternative.fill.as_mut()
        {
            *tile_rect = Some(ooxml_common::fill::FillRect::default());
            *flip = Some("none".into());
        }
        assert_eq!(same_fill(&recorded, &alternative, 90.0), Verdict::Same);
        let no_rotation = RecordedFill::Stated(
            gradient(&[(0.0, "000000"), (0.3099975, "FFFFFF")], Some(false)).map(Box::new),
        );
        alternative.fill = gradient(&[(0.31, "FFFFFF"), (0.0, "000000")], None);
        assert_eq!(same_fill(&no_rotation, &alternative, 0.0), Verdict::Same);
        assert_eq!(same_fill(&no_rotation, &alternative, 90.0), Verdict::Same);
        // Another stop layout cannot be equated by this bounded rule.
        alternative.fill = gradient(&[(0.0, "000000"), (0.31, "FFFFFF")], None);
        assert_eq!(
            same_fill(&recorded, &alternative, 90.0),
            Verdict::Unverifiable("fill")
        );
        assert_eq!(
            same_fill(&RecordedFill::Unknown, &alternative, 90.0),
            Verdict::Unverifiable("fill")
        );
    }

    fn body(paragraphs: &[&[serde_json::Value]]) -> Option<pptx_model::TextBody> {
        Some(
            serde_json::from_value(serde_json::json!({
                "verticalAnchor": "t", "defaultFontSize": null,
                "defaultBold": null, "defaultItalic": null,
                "lIns": 0, "rIns": 0, "tIns": 0, "bIns": 0,
                "wrap": "square", "vert": "horz", "autoFit": "none",
                "paragraphs": paragraphs.iter().map(|runs| para(runs)).collect::<Vec<_>>()
            }))
            .unwrap(),
        )
    }

    fn texts(shape: &ShapeElement) -> Vec<Vec<String>> {
        shape
            .text_body
            .iter()
            .flat_map(|body| &body.paragraphs)
            .map(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .map(|run| match run {
                        TextRun::Text(d) => d.text.clone(),
                        TextRun::Break { .. } => "<br>".to_owned(),
                        TextRun::Math { .. } => "<math>".to_owned(),
                    })
                    .collect()
            })
            .collect()
    }

    #[test]
    fn masked_break_style_fanout_is_charged_before_projection() {
        let mut alternative = shape("rect");
        let mut styled = run("__________");
        styled["fontFamily"] = serde_json::json!("x".repeat(512));
        alternative.text_body = body(&[&[styled]]);
        assert_eq!(
            substitute_text_bounded(
                &mut alternative,
                Some("A\u{b}A\u{b}A\u{b}A\u{b}A\u{b}"),
                &mut 1024
            ),
            Err(Verdict::Unverifiable("text model budget"))
        );
    }

    #[test]
    fn masked_text_retains_line_break_formatting() {
        let mut alternative = shape("rect");
        let mut styled = run("___");
        styled["fontSize"] = serde_json::json!(24.0);
        styled["bold"] = serde_json::json!(true);
        alternative.text_body = body(&[&[
            styled,
            serde_json::json!({"type":"break", "fontSize":36.0, "fontFamily":"Arial", "italic":true}),
        ]]);
        substitute_text(&mut alternative, Some("A\u{b}B\u{b}")).unwrap();
        let runs = &alternative.text_body.as_ref().unwrap().paragraphs[0].runs;
        assert!(matches!(
            &runs[1],
            TextRun::Break {
                font_size: Some(24.0),
                bold: Some(true),
                ..
            }
        ));
        assert!(
            matches!(&runs[3], TextRun::Break { font_size: Some(36.0), font_family: Some(face), italic: Some(true), .. } if face == "Arial")
        );
    }

    #[test]
    fn masked_text_takes_binary_characters_run_by_run() {
        let mut alternative = shape("rect");
        alternative.text_body = body(&[
            &[run("__"), serde_json::json!({"type": "break"}), run(" _ ")],
            &[run("____")],
        ]);
        let mut adopted = alternative.clone();
        assert_eq!(
            substitute_text(&mut adopted, Some("AB\u{b}C😀\rDEFG")),
            Ok(())
        );
        assert_eq!(texts(&adopted), [vec!["AB", "<br>", "C😀"], vec!["DEFG"]]);
        // Length, paragraph and break mismatches are different text.
        for text in [
            "AB\u{b}C😀\rDEF",
            "AB\u{b}C😀",
            "ABXC😀\rDEFG",
            "AB\u{b}C😀\u{b}\rDEFG",
        ] {
            assert_eq!(
                substitute_text(&mut alternative.clone(), Some(text)),
                Err(Verdict::Differs("text structure")),
                "{text:?}"
            );
        }
        // A run boundary inside a surrogate pair is rejected.
        let mut split = shape("rect");
        split.text_body = body(&[&[run(" "), run(" ")]]);
        assert!(substitute_text(&mut split, Some("😀")).is_err());
    }

    #[test]
    fn picture_bullet_requires_the_office_downrevision_marker_and_live_target() {
        let mut binary = shape("rect");
        binary.text_body = body(&[&[run("A")]]);
        binary.text_body.as_mut().unwrap().paragraphs[0].bullet = pptx_model::Bullet::Char {
            ch: "•".to_owned(),
            color: None,
            size_pct: None,
            size_pts: None,
            font_family: None,
        };
        let mut alternative = binary.clone();
        alternative.text_body.as_mut().unwrap().paragraphs[0].bullet = pptx_model::Bullet::Blip {
            image_path: "media/bullet.gif".to_owned(),
            mime_type: "image/gif".to_owned(),
            size_pct: None,
            size_pts: None,
        };
        let assets = vec![("media/bullet.gif".to_owned(), "gif", b"GIF89a".to_vec())];
        assert_eq!(
            same_bullet_fallback(&binary, &alternative, &assets),
            Verdict::Same
        );
        if let pptx_model::Bullet::Char { ch, .. } =
            &mut binary.text_body.as_mut().unwrap().paragraphs[0].bullet
        {
            *ch = "■".to_owned();
        }
        assert_eq!(
            same_bullet_fallback(&binary, &alternative, &assets),
            Verdict::Differs("picture bullet fallback")
        );
        assert_eq!(
            same_bullet_fallback(&binary, &alternative, &[]),
            Verdict::Unverifiable("picture bullet resource")
        );
    }

    #[test]
    fn binary_pattern_cells_match_measured_drawingml_presets() {
        // A synthetic single-frame GIF: clear before every literal keeps its
        // LZW code width at three bits, independent of dictionary growth.
        fn gif(rows: [u8; 8]) -> Vec<u8> {
            let mut pixels = [0u8; 100];
            for y in 0..8 {
                for x in 0..8 {
                    pixels[y * 10 + x] = u8::from(rows[y] & (0x80 >> x) != 0);
                }
            }
            let mut codes = Vec::new();
            for pixel in pixels {
                codes.extend([4u8, pixel]);
            }
            codes.push(5);
            let mut compressed = vec![0u8; (codes.len() * 3).div_ceil(8)];
            for (i, code) in codes.into_iter().enumerate() {
                for bit in 0..3 {
                    let at = i * 3 + bit;
                    compressed[at / 8] |= ((code >> bit) & 1) << (at % 8);
                }
            }
            let mut gif = b"GIF87a\x0a\x00\x0a\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff\x2c\x00\x00\x00\x00\x0a\x00\x0a\x00\x00\x02".to_vec();
            gif.push(compressed.len() as u8);
            gif.extend(compressed);
            gif.extend([0, 0x3b]);
            gif
        }
        let pct30 = [0xaa, 0x44, 0xaa, 0x11, 0xaa, 0x44, 0xaa, 0x11];
        let diagonal = [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88];
        assert_eq!(binary_pattern_cells(&gif(pct30)), Some(pct30));
        assert_eq!(binary_pattern_cells(&gif(diagonal)), Some(diagonal));
        let mut changed = pct30;
        changed[0] ^= 1;
        assert_ne!(binary_pattern_cells(&gif(changed)), Some(pct30));
        let mut truncated = gif(pct30);
        truncated.truncate(32);
        assert!(binary_pattern_cells(&truncated).is_none());
    }

    #[test]
    fn a_masked_vertical_tab_becomes_a_line_break_inside_its_run() {
        let mut alternative = shape("rect");
        alternative.text_body = body(&[&[run("__"), run("_"), run("___"), run("__")]]);
        let mut adopted = alternative.clone();
        assert_eq!(
            substitute_text(&mut adopted, Some("AB\u{b}\u{b}CDEF")),
            Ok(())
        );
        assert_eq!(texts(&adopted), [vec!["AB", "<br>", "<br>", "CD", "EF"]]);
        // Other break characters inside a run have no observed form.
        assert_eq!(
            substitute_text(&mut alternative.clone(), Some("AB\u{2028}\u{b}CDEF")),
            Err(Verdict::Unverifiable("line break"))
        );
    }

    #[test]
    fn run_formatting_must_agree_where_both_sides_state_it() {
        let sized = |size: serde_json::Value| {
            let mut r = run("ABCD");
            r["fontSize"] = size;
            body(&[&[r]])
        };
        let mut binary = shape("rect");
        let mut alternative = shape("rect");
        binary.text_body = sized(serde_json::json!(10.0));
        alternative.text_body = sized(serde_json::json!(10.0));
        assert_eq!(
            same_run_formatting(&binary, &alternative, None),
            Verdict::Same
        );
        // Whole-point binary sizes cannot state 10.5 pt: not the same shape.
        alternative.text_body = sized(serde_json::json!(10.5));
        assert_eq!(
            same_run_formatting(&binary, &alternative, None),
            Verdict::Differs("run formatting")
        );
        // A picture-bullet alternative may restate an inherited master size;
        // a directly authored CFStyle size still has to agree.
        assert_eq!(
            same_run_formatting(&binary, &alternative, Some(&[false; 5])),
            Verdict::Same
        );
        assert_eq!(
            same_run_formatting(&binary, &alternative, Some(&[true; 5])),
            Verdict::Differs("run formatting")
        );
        // An unstated size on either side is not a disagreement.
        binary.text_body = sized(serde_json::Value::Null);
        assert_eq!(
            same_run_formatting(&binary, &alternative, None),
            Verdict::Same
        );
        // A field compares as one unit whatever its displayed text.
        let mut field = run("12");
        field["fieldType"] = serde_json::json!("slidenum");
        field["fontSize"] = serde_json::json!(20.0);
        binary.text_body = body(&[&[run("A"), field.clone()]]);
        field["text"] = serde_json::json!("<#>");
        field["fontSize"] = serde_json::json!(18.0);
        alternative.text_body = body(&[&[run("A"), field]]);
        assert_eq!(
            same_run_formatting(&binary, &alternative, None),
            Verdict::Differs("run formatting")
        );
    }

    #[cfg(feature = "direct-ppt")]
    mod adoption {
        use super::*;
        use std::io::Write;

        const SHAPE_XML: &str = r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="9" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1587500" cy="793750"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>"#;

        fn package(parts: &[(&str, &str)], relationships: &[(&str, &str)]) -> Vec<u8> {
            package_named_rels("_rels/.rels", parts, relationships)
        }

        fn package_named_rels(
            rels_name: &str,
            parts: &[(&str, &str)],
            relationships: &[(&str, &str)],
        ) -> Vec<u8> {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            let rels: String = relationships
                .iter()
                .enumerate()
                .map(|(i, (kind, target))| {
                    format!(r#"<Relationship Id="rId{i}" Type="http://schemas.microsoft.com/office/2006/relationships/{kind}" Target="{target}"/>"#)
                })
                .collect();
            let rels = format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
            );
            for (name, body) in [(rels_name, rels.as_str())].iter().chain(parts) {
                writer.start_file(*name, options).unwrap();
                writer.write_all(body.as_bytes()).unwrap();
            }
            writer.finish().unwrap().into_inner()
        }

        fn blob(shape_xml: &str) -> Vec<u8> {
            package(
                &[("drs/shapexml.xml", shape_xml), ("drs/downrev.xml", "<x/>")],
                &[
                    ("downRev", "drs/downrev.xml"),
                    ("shapeXml", "drs/shapexml.xml"),
                ],
            )
        }

        #[test]
        fn alternative_relationship_uses_opc_resolution_and_equivalent_part_lookup() {
            let package = package_named_rels(
                "_RELS/%2Erels",
                &[("drs/ShapeXML.xml", SHAPE_XML)],
                &[("shapeXml", "DRS/%73hapeXML.xml#shape")],
            );
            assert_eq!(
                alternative_part(&package).unwrap(),
                Some(("shapeXml".into(), "DRS/shapeXML.xml".into()))
            );
            assert!(placeholder_locals(&package, "DRS/shapeXML.xml", &theme()).is_ok());
        }

        #[test]
        fn placeholder_style_references_are_local_effects_and_lines() {
            let shape = SHAPE_XML.replace(
                "</p:spPr>",
                "</p:spPr><p:style><a:lnRef idx=\"1\"/><a:effectRef idx=\"1\"/></p:style>",
            );
            let local = placeholder_locals(&blob(&shape), "drs/shapexml.xml", &theme()).unwrap();
            assert!(local.stroke);
            assert!(local.effects);
            assert!(local.fill);
            let foreign = SHAPE_XML.replace(
                "</p:spPr>",
                "</p:spPr><x:style xmlns:x=\"urn:foreign\"><x:lnRef/><x:effectRef/></x:style>",
            );
            let foreign =
                placeholder_locals(&blob(&foreign), "drs/shapexml.xml", &theme()).unwrap();
            assert!(!foreign.stroke);
            assert!(!foreign.effects);
            let missing = shape.replace("idx=\"1\"", "idx=\"99\"");
            assert!(
                placeholder_locals(&blob(&missing), "drs/shapexml.xml", &theme())
                    .err()
                    .unwrap()
                    .contains("placeholder style reference")
            );
        }

        #[test]
        fn blob_relationship_targets_are_checked_passively() {
            let shape = SHAPE_XML.replace(
                "</p:spPr>",
                "<a:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rIdImage\"/></a:blipFill></p:spPr>",
            );
            let rels = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image.jpeg\"/></Relationships>";
            let parts = &[
                ("drs/shapexml.xml", shape.as_str()),
                ("drs/_rels/shapexml.xml.rels", rels),
                ("drs/media/image.jpeg", "passive: never decode"),
            ];
            let present = package(parts, &[]);
            assert!(validate_blob_references(&present, "drs/shapexml.xml").is_ok());
            assert_eq!(
                sole_blob_image(&present, "drs/shapexml.xml", "drs/media/image.jpeg").unwrap(),
                b"passive: never decode"
            );
            assert!(sole_blob_image(&present, "drs/shapexml.xml", "drs/media/other.jpeg").is_err());
            let missing = package(&parts[..2], &[]);
            assert!(validate_blob_references(&missing, "drs/shapexml.xml").is_err());

            let external = rels.replace(
                "Target=\"media/image.jpeg\"",
                "Target=\"https://example.test/image\" TargetMode=\"External\"",
            );
            let external = package(
                &[
                    ("drs/shapexml.xml", &shape),
                    ("drs/_rels/shapexml.xml.rels", &external),
                ],
                &[],
            );
            assert!(validate_blob_references(&external, "drs/shapexml.xml").is_ok());
            assert!(
                sole_blob_image(&external, "drs/shapexml.xml", "drs/media/image.jpeg").is_err()
            );

            let duplicate = rels.replacen(
                "<Relationship ",
                "<Relationship Id=\"rIdImage\" Type=\"image\" Target=\"media/image.jpeg\"/><Relationship ",
                1,
            );
            let duplicate = package(
                &[
                    ("drs/shapexml.xml", &shape),
                    ("drs/_rels/shapexml.xml.rels", &duplicate),
                    ("drs/media/image.jpeg", "passive"),
                ],
                &[],
            );
            assert!(validate_blob_references(&duplicate, "drs/shapexml.xml").is_err());
        }

        fn theme() -> Theme {
            Theme::Readable {
                theme_xml: crate::ppt::theme(),
                clr_map: None,
                format_scheme: std::cell::OnceCell::new(),
            }
        }

        fn element(color: &str) -> ShapeElement {
            let mut element = shape("rect");
            (element.width, element.height) = (1587500, 793750);
            element.fill = solid(color);
            element.id = Some("3".to_owned());
            element
        }

        fn run_adopt(
            element: &ShapeElement,
            blob: &[u8],
            theme: &Theme,
            budgets: (usize, usize, usize),
        ) -> (Result<Option<ShapeElement>, String>, usize) {
            let leaf = pptx_model::Transform {
                cx: element.width,
                cy: element.height,
                ..Default::default()
            };
            let shape = binary(element, &leaf);
            let (mut work, mut text, mut model) = budgets;
            let result = adopt(
                &shape,
                blob,
                theme,
                &mut super::media::SpanStore::new(Vec::new()),
                &mut work,
                &mut text,
                &mut model,
            );
            (result, model)
        }

        const AMPLE: (usize, usize, usize) = (usize::MAX, usize::MAX, usize::MAX);

        #[test]
        fn ordinary_alternative_requires_readable_and_resolved_theme_styles() {
            let styled = SHAPE_XML.replace("</p:sp>", r#"<p:style><a:lnRef idx="999"><a:srgbClr val="FF0000"/></a:lnRef></p:style></p:sp>"#);
            assert!(
                run_adopt(&element("FF0000"), &blob(&styled), &theme(), AMPLE)
                    .0
                    .is_err()
            );
            let unreadable = Theme::Readable {
                theme_xml: "<broken".to_owned(),
                clr_map: None,
                format_scheme: std::cell::OnceCell::new(),
            };
            assert!(
                run_adopt(&element("FF0000"), &blob(SHAPE_XML), &unreadable, AMPLE)
                    .0
                    .is_err()
            );
        }

        #[test]
        fn forged_theme_expanded_size_is_not_retained() {
            let theme_xml = format!("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><!--{}--></a:theme>", "x".repeat(65536));
            let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            for (name, bytes) in [
                (
                    "_rels/.rels",
                    r#"<Relationships><Relationship Id="m" Type="http://schemas.microsoft.com/office/2006/relationships/officeDocument" Target="manager.xml"/></Relationships>"#,
                ),
                (
                    "_rels/manager.xml.rels",
                    r#"<Relationships><Relationship Id="t" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme.xml"/></Relationships>"#,
                ),
                ("theme.xml", theme_xml.as_str()),
            ] {
                zip.start_file(name, options).unwrap();
                zip.write_all(bytes.as_bytes()).unwrap();
            }
            let mut blob = zip.finish().unwrap().into_inner();
            fn record(payload: &[u8]) -> Record<'_> {
                Record {
                    version: 0,
                    instance: 0,
                    kind: 0x040e,
                    payload,
                }
            }
            assert!(matches!(
                master_theme(&[record(&blob)], &mut 100_000).unwrap(),
                Some(Theme::Readable { .. })
            ));
            // Central-directory uncompressed size, with real deflate data and
            // CRC intact. The reader must bound actual reads by the precharge.
            let offset = blob
                .windows(4)
                .enumerate()
                .find_map(|(offset, signature)| {
                    (signature == b"PK\x01\x02"
                        && blob.get(offset + 46..offset + 55) == Some(b"theme.xml"))
                    .then_some(offset)
                })
                .unwrap();
            blob[offset + 24..offset + 28].copy_from_slice(&0u32.to_le_bytes());
            assert!(matches!(
                master_theme(&[record(&blob)], &mut 100_000).unwrap(),
                Some(Theme::Unreadable)
            ));
        }

        #[test]
        fn master_themes_share_a_decoded_byte_budget() {
            let manager_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme.xml"/></Relationships>"#;
            let blob = package(
                &[
                    ("manager.xml", "<x/>"),
                    ("_rels/manager.xml.rels", manager_rels),
                    ("theme.xml", &crate::ppt::theme()),
                ],
                &[("officeDocument", "manager.xml")],
            );
            let record = Record {
                version: 0,
                instance: 0,
                kind: 0x040e,
                payload: &blob,
            };
            let mut budget = 4096;
            assert!(matches!(
                master_theme(&[record], &mut budget).unwrap(),
                Some(Theme::Readable { .. })
            ));
            let decoded_bytes = 4096 - budget;
            assert!(decoded_bytes > 0);
            let mut budget = decoded_bytes;
            assert!(matches!(
                master_theme(&[record], &mut budget).unwrap(),
                Some(Theme::Readable { .. })
            ));
            assert!(master_theme(&[record], &mut budget).is_err());
        }

        #[test]
        fn a_consistent_alternative_is_adopted_and_charged() {
            let (result, model) = run_adopt(&element("FF0000"), &blob(SHAPE_XML), &theme(), AMPLE);
            let adopted = result.unwrap().expect("consistent alternative");
            assert_eq!(adopted.id.as_deref(), Some("3"));
            let charged = ooxml_common::json_measurement::measure_json(&adopted).unwrap();
            assert_eq!(usize::MAX - model, charged.json_bytes as usize);
            // A model budget below the adopted shape fails closed.
            let short = (usize::MAX, usize::MAX, charged.json_bytes as usize - 1);
            assert!(
                run_adopt(&element("FF0000"), &blob(SHAPE_XML), &theme(), short)
                    .0
                    .is_err()
            );
        }

        #[test]
        fn an_absent_or_disagreeing_alternative_keeps_the_binary() {
            let downrev_only = package(
                &[("drs/downrev.xml", "<x/>")],
                &[("downRev", "drs/downrev.xml")],
            );
            assert!(matches!(
                run_adopt(&element("FF0000"), &downrev_only, &theme(), AMPLE).0,
                Ok(None)
            ));
            assert!(matches!(
                run_adopt(&element("00FF00"), &blob(SHAPE_XML), &theme(), AMPLE).0,
                Ok(None)
            ));
        }

        #[test]
        fn placeholder_inherits_missing_shape_properties_from_the_binary() {
            let placeholder = SHAPE_XML
                .replace("<p:nvPr/>", r#"<p:nvPr><p:ph type="title"/></p:nvPr>"#)
                .replace(
                    r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1587500" cy="793750"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>"#,
                    "<p:spPr/>",
                );
            let adopted = run_adopt(&element("FF0000"), &blob(&placeholder), &theme(), AMPLE)
                .0
                .unwrap()
                .expect("placeholder alternative");
            assert_eq!(adopted.geometry, "rect");
            assert!(matches!(adopted.fill, Some(Fill::Solid { color }) if color == "FF0000"));
            assert_eq!((adopted.width, adopted.height), (1587500, 793750));
        }

        #[test]
        fn placeholder_uses_binary_transform_even_with_a_local_xml_transform() {
            let placeholder = SHAPE_XML
                .replace("<p:nvPr/>", r#"<p:nvPr><p:ph type="body"/></p:nvPr>"#)
                .replace(r#"<a:off x="0" y="0"/>"#, r#"<a:off x="635000" y="0"/>"#);
            let adopted = run_adopt(&element("FF0000"), &blob(&placeholder), &theme(), AMPLE)
                .0
                .unwrap()
                .expect("binary anchor wins");
            assert_eq!(adopted.x, 0);
            assert!(matches!(adopted.fill, Some(Fill::Solid { color }) if color == "FF0000"));
        }

        #[test]
        fn disagreeing_local_placeholder_overrides_have_unknown_provenance() {
            let placeholder =
                SHAPE_XML.replace("<p:nvPr/>", r#"<p:nvPr><p:ph type="body"/></p:nvPr>"#);
            let geometry = placeholder.replace("prst=\"rect\"", "prst=\"ellipse\"");
            let fill = placeholder.replace("val=\"FF0000\"", "val=\"00FF00\"");
            for (xml, attribute) in [
                (geometry.as_str(), "placeholder geometry precedence"),
                (fill.as_str(), "placeholder fill precedence"),
            ] {
                let error = run_adopt(&element("FF0000"), &blob(xml), &theme(), AMPLE)
                    .0
                    .unwrap_err();
                assert!(error.starts_with("UNSUPPORTED:") && error.contains(attribute));
            }
        }

        #[test]
        fn placeholder_keeps_alternative_text_formatting_with_binary_characters() {
            let xml = SHAPE_XML
                .replace("<p:nvPr/>", r#"<p:nvPr><p:ph type="title"/></p:nvPr>"#)
                .replace(
                    "</p:sp>",
                    "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz=\"4400\" b=\"1\"/><a:t>_____</a:t></a:r></a:p></p:txBody></p:sp>",
                );
            let mut element = element("FF0000");
            element.text_body = body(&[&[run("HELLO")]]);
            let leaf = pptx_model::Transform {
                cx: element.width,
                cy: element.height,
                ..Default::default()
            };
            let mut binary = binary(&element, &leaf);
            binary.text = Some("HELLO");
            let (mut work, mut text, mut model) = AMPLE;
            let adopted = adopt(
                &binary,
                &blob(&xml),
                &theme(),
                &mut super::media::SpanStore::new(Vec::new()),
                &mut work,
                &mut text,
                &mut model,
            )
            .unwrap()
            .expect("placeholder alternative");
            let run = &adopted.text_body.unwrap().paragraphs[0].runs[0];
            assert!(
                matches!(run, TextRun::Text(data) if data.text == "HELLO" && data.font_size == Some(44.0) && data.bold == Some(true))
            );
        }

        #[test]
        fn freeform_binary_keeps_its_outline_with_a_preset_alternative() {
            let mut element = element("FF0000");
            element.geometry = "custGeom".into();
            element.cust_geom = Some(vec![vec![
                pptx_model::PathCmd::MoveTo { x: 0.0, y: 0.0 },
                pptx_model::PathCmd::LineTo { x: 1.0, y: 0.0 },
                pptx_model::PathCmd::LineTo { x: 1.0, y: 1.0 },
                pptx_model::PathCmd::Close,
            ]]);
            let adopted = run_adopt(&element, &blob(SHAPE_XML), &theme(), AMPLE)
                .0
                .unwrap()
                .expect("alternative formatting with binary outline");
            assert_eq!(adopted.geometry, "custGeom");
            assert!(adopted.cust_geom.is_some());
            assert!(matches!(adopted.fill, Some(Fill::Solid { color }) if color == "FF0000"));
        }

        #[test]
        fn unverifiable_alternatives_fail_closed() {
            let element = element("FF0000");
            let unsupported = |result: (Result<Option<ShapeElement>, String>, usize)| {
                let error = result.0.expect_err("must fail closed");
                assert!(error.starts_with("UNSUPPORTED:"), "{error}");
            };
            // Oversized, over budget, unreadable.
            unsupported(run_adopt(
                &element,
                &vec![0; MAX_BLOB_BYTES + 1],
                &theme(),
                AMPLE,
            ));
            unsupported(run_adopt(
                &element,
                &blob(SHAPE_XML),
                &theme(),
                (0, usize::MAX, usize::MAX),
            ));
            unsupported(run_adopt(
                &element,
                &blob(SHAPE_XML),
                &theme(),
                (usize::MAX, 10, usize::MAX),
            ));
            unsupported(run_adopt(&element, b"not a package", &theme(), AMPLE));
            unsupported(run_adopt(&element, &blob("<p:sp"), &theme(), AMPLE));
            unsupported(run_adopt(
                &element,
                &blob(SHAPE_XML),
                &Theme::Unreadable,
                AMPLE,
            ));
            // Non-shape alternatives.
            let group = package(
                &[("drs/groupshapexml.xml", "<x/>")],
                &[
                    ("downRev", "drs/downrev.xml"),
                    ("groupShapeXml", "drs/groupshapexml.xml"),
                ],
            );
            unsupported(run_adopt(&element, &group, &theme(), AMPLE));
            let two = package(
                &[("drs/shapexml.xml", SHAPE_XML)],
                &[
                    ("shapeXml", "drs/shapexml.xml"),
                    ("connectorXml", "drs/shapexml.xml"),
                ],
            );
            unsupported(run_adopt(&element, &two, &theme(), AMPLE));
            // A common attribute that cannot be compared.
            let gradient = SHAPE_XML.replace(
                r#"<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>"#,
                r#"<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>"#,
            );
            let mut patterned = element.clone();
            patterned.fill = Some(Fill::Pattern {
                fg: "FF0000".into(),
                bg: "0000FF".into(),
                preset: "pct5".into(),
            });
            unsupported(run_adopt(&patterned, &blob(&gradient), &theme(), AMPLE));
        }

        #[test]
        fn non_presentation_root_and_theme_image_dependency_fail_closed() {
            let element = element("FF0000");
            for root in [
                SHAPE_XML.replace(
                    "http://schemas.openxmlformats.org/presentationml/2006/main",
                    "urn:foreign",
                ),
                SHAPE_XML.replace("p:sp", "sp"),
            ] {
                let error = run_adopt(&element, &blob(&root), &theme(), AMPLE)
                    .0
                    .expect_err("foreign root cannot be adopted");
                assert!(error.starts_with("UNSUPPORTED:"), "{error}");
            }

            let xml = SHAPE_XML
                .replace("<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>", "")
                .replace(
                    "</p:sp>",
                    "<p:style><a:fillRef idx=\"1\"/></p:style></p:sp>",
                );
            for prefix in ["r", "関係"] {
                let theme = Theme::Readable {
                    theme_xml: format!(
                        r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:{prefix}="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:themeElements><a:fmtScheme name="x"><a:fillStyleLst><a:blipFill><a:blip {prefix}:embed="rIdImage"/></a:blipFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>"#
                    ),
                    clr_map: None,
                    format_scheme: std::cell::OnceCell::new(),
                };
                let error = run_adopt(&element, &blob(&xml), &theme, AMPLE)
                    .0
                    .expect_err("theme image relationship cannot be verified");
                assert!(
                    error.starts_with("UNSUPPORTED:") && error.contains("relationship"),
                    "{error}"
                );
            }
        }
    }

    fn run(text: &str) -> serde_json::Value {
        serde_json::json!({
            "type": "text", "text": text, "bold": null, "italic": null,
            "underline": false, "strikethrough": false, "strikeDouble": false,
            "fontSize": null, "color": null, "fontFamily": null, "fieldType": null, "hyperlinkUsesTextFill": false
        })
    }

    fn para(runs: &[serde_json::Value]) -> serde_json::Value {
        serde_json::json!({
            "alignment": "l", "marL": 0, "marR": 0, "indent": 0,
            "spaceBefore": null, "spaceAfter": null, "spaceLine": null, "lvl": 0,
            "bullet": {"type": "none"}, "defFontSize": null, "defColor": null,
            "defBold": null, "defItalic": null, "defFontFamily": null,
            "tabStops": [], "rtl": false, "eaLnBrk": true, "runs": runs
        })
    }
}
