//! Shared OOXML chart model and parsers used by DOCX, XLSX, and PPTX.
//!
//! The format parsers read a chart `<c:chartSpace>` (and the modern
//! `<cx:chartSpace>` for waterfall / treemap / box-and-whisker etc.) but
//! historically did so with two near-identical bodies sitting in
//! `packages/xlsx/parser/src/lib.rs` and `packages/pptx/parser/src/lib.rs`.
//! The result was that fields added on one side stayed missing on the other
//! until a host-specific regression exposed the drift (for example, one
//! adapter once discarded `legendPos` while another preserved it).
//!
//! This module owns the shared chart wire model and XML parsing. Package
//! adapters supply theme, image, style, and formula-resolution resources
//! through [`ChartParseContext`]; host-specific frame geometry stays local.
//!
//! ## Namespace handling
//!
//! All helpers match elements by local name only. Real chart documents put
//! everything under either the `c:` (chart 2006) or `cx:` (chartEx 2014)
//! namespace and never mix non-chart elements at these paths, so the strict
//! `tag_name().namespace() == Some(c_ns)` check in xlsx adds nothing in
//! practice — this module drops it for symmetry with the pptx side and to
//! keep the API simple. If a future format wedges a non-chart element into
//! `<c:plotArea>` the caller can pre-filter before delegating here.
//!
//! All field references are to ECMA-376 / ISO-29500 part 1 §21.2 (DrawingML
//! Charts) unless stated otherwise.

use roxmltree::Node;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};

use crate::text::{parse_body_pr, BodyPrDefaults};
use crate::units::coordinate32_to_emu;

mod classic_style;

/// Resource ceiling for the expanded Chart Colors total set. Typical Office
/// parts contain 6 base colors × at most 9 variations; this bound prevents an
/// adversarial colors×variations product from amplifying a bounded XML tree.
const MAX_CHART_COLOR_STYLE_ENTRIES: usize = 4096;
/// A legend is a bounded UI list even when its backing series cache is much
/// larger. Bound authored entry overrides before allocating the wire model.
const MAX_CHART_LEGEND_ENTRIES: usize = 4096;
/// Aggregate structured-fill components retained after expanding one Chart
/// Style role across the linked Chart Colors palette. A gradient contributes
/// one component per stop; solid and pattern fills contribute one. This bounds
/// the otherwise multiplicative `palette entries × gradient stops` wire model.
const MAX_CHART_STYLE_PAINT_COMPONENTS: usize = 1_048_576;
/// A single structured paint recipe can be replayed for every visible chart
/// datum. Keep it bounded before `parse_grad_fill` allocates/sorts its stop
/// list; the aggregate budgets below then bound all retained recipes.
const MAX_CHART_PAINT_RECIPE_COMPONENTS: usize = 4096;
const MAX_CHART_MARKER_GRADIENT_STOPS: usize = MAX_CHART_PAINT_RECIPE_COMPONENTS;
/// Aggregate direct marker-paint components retained by one chart. This is an
/// availability boundary, not a visual compatibility rule.
const MAX_CHART_MARKER_PAINT_COMPONENTS: usize = MAX_CHART_STYLE_PAINT_COMPONENTS;
/// Direct data/trendline label shapes can replay one gradient for every label.
/// Bound both a single recipe and the aggregate retained wire components before
/// parsing stop lists. These are availability ceilings, not visual rules.
const MAX_CHART_LABEL_GRADIENT_STOPS: usize = MAX_CHART_PAINT_RECIPE_COMPONENTS;
const MAX_CHART_LABEL_PAINT_COMPONENTS: usize = MAX_CHART_STYLE_PAINT_COMPONENTS;
/// Aggregate role×palette slots retained from one linked Chart Style. Typical
/// Office parts use 30 paint roles × 6–54 colors. Reject the entire fallback
/// table before expansion rather than retaining an arbitrary role or palette
/// prefix when a hostile colors part exceeds this bounded wire budget.
const MAX_CHART_STYLE_ROLE_SLOTS: usize = 8_192;
/// Maximum cache width accepted from `<c:ptCount>` / `<cx:lvl ptCount>`.
/// Chart data originates in worksheet ranges, whose largest single dimension
/// is 1,048,576 rows. Rejecting wider sparse caches prevents an XML attribute
/// or point index from requesting an unbounded WASM allocation.
const MAX_CHART_CACHE_POINTS: usize = 1_048_576;
/// Classic plot groups are renderer work items even when they contain no
/// series. Keep public/wire planning aligned with the synchronous Canvas
/// availability ceiling before allocating the ordered group vector.
const MAX_CHART_PLOT_GROUPS: usize = 10_000;
const MAX_CHART_PLOT_SERIES: usize = 10_000;
const MAX_CHART_PLOT_AXES: usize = 4;
const MAX_CHART_GROUP_AXIS_IDS: usize = 3;

fn claim_plot_group_axis_slot(
    axis_id: Option<&String>,
    known_axis_ids: &BTreeMap<String, String>,
    primary: &mut Option<String>,
    secondary: &mut Option<String>,
) -> String {
    let slot = match axis_id {
        None => "unresolved",
        Some(axis_id) if !known_axis_ids.contains_key(axis_id) => "unresolved",
        Some(axis_id) if primary.as_ref().is_none_or(|current| current == axis_id) => {
            primary.get_or_insert_with(|| axis_id.clone());
            "primary"
        }
        Some(axis_id) if secondary.as_ref().is_none_or(|current| current == axis_id) => {
            secondary.get_or_insert_with(|| axis_id.clone());
            "secondary"
        }
        Some(_) => "unresolved",
    };
    slot.to_string()
}

mod axis;
mod cache;
mod chartex;
mod classic;
mod labels;
mod model;
mod related_parts;
mod retained;
mod style;

pub use axis::*;
pub use cache::*;
use chartex::*;
use classic::*;
pub use labels::*;
pub use model::*;
pub use related_parts::{load_chart_related_parts, load_chart_user_shapes_xml, ChartRelatedParts};
pub use retained::RetainedBytes;
use style::*;

#[cfg(test)]
mod tests;

/// Host application whose ChartEx layout policy a chart part follows. Office
/// applications lay out the same ChartEx part differently (for example Pareto
/// series retention and ordering), so the parsing host must be named.
#[derive(Default, Clone, Copy, Debug, PartialEq, Eq)]
pub enum ChartHost {
    /// The caller did not name a host. Host-neutral: it applies the non-Excel
    /// behaviour (all columns retained, PowerPoint-style ordering), but none of
    /// the extras gated on a positive PowerPoint identity (currently the
    /// unpaired percentage axis). Every shipped package sets its own host, so
    /// this is only for hosts-agnostic callers and tests.
    #[default]
    Unspecified,
    PowerPoint,
    Excel,
    /// Word follows the PowerPoint policy until Office-produced Word controls
    /// are measured; `parse_chartex_impl` maps it in one place.
    Word,
}

/// Identity of one independently retained ChartEx model within an OPC package.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ChartRetentionKey<'a> {
    pub source_part: &'a str,
    pub site: &'a str,
}

/// Package-owned sidecars and lookup hooks for one chart part. All fields are
/// optional so callers only supply resources present in the host package.
/// A color resolver is required to parse; omission returns `None`.
/// Formula lookup uses a cell because the host may memoize range resolution
/// while the parse entry point accepts a shared context reference.
#[derive(Default)]
pub struct ChartParseContext<'a> {
    pub host: ChartHost,
    /// Package-operation capability: hard failures survive Option-based adapters.
    pub limit_reporter: Option<&'a crate::package_session::PackageLimitReporter>,
    /// Stable source-part and retention-site identity for package accounting.
    pub retention_key: Option<ChartRetentionKey<'a>>,
    pub color_resolver: Option<&'a dyn ColorResolver>,
    pub style_xml: Option<&'a str>,
    pub color_style_xml: Option<&'a str>,
    pub images: Option<&'a dyn ChartImageResolver>,
    pub references: std::cell::Cell<Option<&'a mut dyn ChartReferenceResolver>>,
}

impl<'a> ChartParseContext<'a> {
    /// Construct a context from the host's optional linked parts and resources.
    #[inline(never)]
    pub fn new(
        color_resolver: &'a dyn ColorResolver,
        style_xml: Option<&'a str>,
        color_style_xml: Option<&'a str>,
        images: Option<&'a dyn ChartImageResolver>,
        references: Option<&'a mut dyn ChartReferenceResolver>,
    ) -> Self {
        Self {
            host: ChartHost::Unspecified,
            limit_reporter: None,
            retention_key: None,
            color_resolver: Some(color_resolver),
            style_xml,
            color_style_xml,
            images,
            references: std::cell::Cell::new(references),
        }
    }
}

/// Parse an ECMA-376 DrawingML chart part into the shared wire model.
pub fn parse_chart_part(root: Node, context: &ChartParseContext<'_>) -> Option<ChartModel> {
    if let Err((limit, observed)) = classic_cache_preparse(root) {
        // Fail closed with a typed hard-limit violation (the host turns the
        // latched violation into its OOXML_RESOURCE_LIMIT error) instead of
        // letting `vec![_; ptCount]` request an unbounded allocation.
        use crate::resource::{observe_hard_limit, HardResourceLimitKind};
        let kind = HardResourceLimitKind::ChartCachePoints;
        if let Some(reporter) = context.limit_reporter {
            let _ = reporter.observe_hard_limit(kind, None, limit, observed);
        } else {
            let _ = observe_hard_limit(kind, None, limit, observed);
        }
        return None;
    }
    parse_part(root, context, parse_classic_impl)
}

/// Total cache slots one classic chart part may declare: four full-height
/// worksheet columns of [`MAX_CHART_CACHE_POINTS`] each. Library resource
/// policy, not an ECMA-376 rule.
const MAX_CLASSIC_CHART_CACHE_SLOTS: usize = 4 * MAX_CHART_CACHE_POINTS;

/// Bound the dense vectors the classic parser sizes from `<c:ptCount>` and
/// `<c:pt idx>` before any of them is allocated.
///
/// The cache collectors size one vector per cache from whichever `ptCount`
/// they meet (some take the first child, others the last descendant of a
/// series container) or from the largest `pt idx`, so admission does not
/// mirror one collector. Every element's width is the largest `ptCount` value,
/// `pt idx + 1` or number of `pt` among its children; each width is bounded by
/// [`MAX_CHART_CACHE_POINTS`] (the worksheet row limit chart data comes from)
/// and their sum over the part by [`MAX_CLASSIC_CHART_CACHE_SLOTS`]. Every
/// `ptCount` and `pt` is some element's child, so this covers every vector the
/// collectors can size (a `multiLvlStrCache` counts each `lvl`; its own
/// collector separately bounds width × levels), and error bars beyond the
/// schema's two per series are charged at their plus/minus vector cost.
/// `Err` carries the crossed limit and the observed (saturated) count.
fn classic_cache_preparse(root: Node) -> Result<(), (u64, u64)> {
    let single = MAX_CHART_CACHE_POINTS as u64;
    let total_limit = MAX_CLASSIC_CHART_CACHE_SLOTS as u64;
    let parse = |value: Option<&str>| value.and_then(|v| v.trim().parse::<u64>().ok());
    let mut total = 0u64;
    for node in root.descendants().filter(Node::is_element) {
        // Collectors without a usable `ptCount` fall back to the number of
        // `pt` elements, so that count is a width too.
        let (mut width, mut point_elements) = (0u64, 0u64);
        for child in node.children().filter(Node::is_element) {
            let child_width = match child.tag_name().name() {
                "ptCount" => parse(child.attribute("val")),
                "pt" => {
                    point_elements += 1;
                    parse(child.attribute("idx")).map(|index| index.saturating_add(1))
                }
                _ => None,
            };
            width = width.max(child_width.unwrap_or(0));
        }
        let width = width.max(point_elements);
        if width > single {
            return Err((single, width));
        }
        total = total.saturating_add(width);
        // Each `<c:errBars>` of a series materializes plus/minus vectors as
        // long as the series. The schema allows at most two per series
        // (CT_ScatterSer x and y); charge every further one at that cost so
        // repeated error bars cannot multiply a bounded cache.
        let extra_error_bars = node
            .children()
            .filter(|child| child.is_element() && child.tag_name().name() == "errBars")
            .count()
            .saturating_sub(2) as u64;
        if extra_error_bars > 0 {
            let series_width = node
                .descendants()
                .filter(Node::is_element)
                .filter_map(|child| match child.tag_name().name() {
                    "ptCount" => parse(child.attribute("val")),
                    "pt" => parse(child.attribute("idx")).map(|index| index.saturating_add(1)),
                    _ => None,
                })
                .max()
                .unwrap_or(0);
            total = total
                .saturating_add(extra_error_bars.saturating_mul(series_width.saturating_mul(2)));
        }
        if total > total_limit {
            return Err((total_limit, total));
        }
    }
    Ok(())
}

/// Parse a Microsoft chartEx part into the shared wire model.
pub fn parse_chartex_part(root: Node, context: &ChartParseContext<'_>) -> Option<ChartModel> {
    use crate::resource::{
        observe_hard_limit, HardResourceLimitKind, HARD_MAX_CHARTEX_ALLOCATION_BYTES,
        HARD_MAX_CHARTEX_ALLOCATION_ELEMENTS,
    };

    match chartex_preparse_lower_bound(root) {
        Ok(_) => {}
        Err(violation) => {
            let (kind, limit, observed) = match violation {
                ChartexPreparseViolation::Elements(observed) => (
                    HardResourceLimitKind::ChartexAllocationElements,
                    HARD_MAX_CHARTEX_ALLOCATION_ELEMENTS,
                    observed,
                ),
                ChartexPreparseViolation::Bytes(observed) => (
                    HardResourceLimitKind::ChartexAllocationBytes,
                    HARD_MAX_CHARTEX_ALLOCATION_BYTES,
                    observed,
                ),
            };
            if let Some(reporter) = context.limit_reporter {
                let _ = reporter.observe_hard_limit(kind, None, limit, observed);
            } else {
                let _ = observe_hard_limit(kind, None, limit, observed);
            }
            return None;
        }
    }

    let canonical = parse_part_value(
        root,
        context,
        |root, resolver, style, colors, refs, images| {
            Some(build_chartex_canonical(
                root, resolver, style, colors, refs, images,
            ))
        },
    )?;
    if let Some(observed) = canonical.checkpoint_violation() {
        if let Some(reporter) = context.limit_reporter {
            let _ = reporter.observe_hard_limit(
                HardResourceLimitKind::ChartexAllocationBytes,
                None,
                HARD_MAX_CHARTEX_ALLOCATION_BYTES,
                observed,
            );
        } else {
            let _ = observe_hard_limit(
                HardResourceLimitKind::ChartexAllocationBytes,
                None,
                HARD_MAX_CHARTEX_ALLOCATION_BYTES,
                observed,
            );
        }
        return None;
    }
    let retained = canonical.retained_bytes();
    let result =
        if let (Some(reporter), Some(key)) = (context.limit_reporter, context.retention_key) {
            reporter.retain_instance(
                HardResourceLimitKind::ChartexAllocationBytes,
                key.source_part,
                key.site,
                HARD_MAX_CHARTEX_ALLOCATION_BYTES,
                retained,
            )
        } else if let Some(reporter) = context.limit_reporter {
            reporter.observe_hard_limit(
                HardResourceLimitKind::ChartexAllocationBytes,
                None,
                HARD_MAX_CHARTEX_ALLOCATION_BYTES,
                retained,
            )
        } else {
            observe_hard_limit(
                HardResourceLimitKind::ChartexAllocationBytes,
                None,
                HARD_MAX_CHARTEX_ALLOCATION_BYTES,
                retained,
            )
        };
    if result.is_err() || retained > HARD_MAX_CHARTEX_ALLOCATION_BYTES {
        return None;
    }
    canonical.project(context.host)
}

fn parse_part(
    root: Node,
    context: &ChartParseContext<'_>,
    parser: impl FnOnce(
        Node<'_, '_>,
        &dyn ColorResolver,
        Option<&str>,
        Option<&str>,
        &mut dyn ChartReferenceResolver,
        &dyn ChartImageResolver,
    ) -> Option<ChartModel>,
) -> Option<ChartModel> {
    parse_part_value(root, context, parser)
}

fn parse_part_value<T>(
    root: Node,
    context: &ChartParseContext<'_>,
    parser: impl FnOnce(
        Node<'_, '_>,
        &dyn ColorResolver,
        Option<&str>,
        Option<&str>,
        &mut dyn ChartReferenceResolver,
        &dyn ChartImageResolver,
    ) -> Option<T>,
) -> Option<T> {
    let color_resolver = context.color_resolver?;
    let images = context.images.unwrap_or(&EmptyChartImageResolver);
    let mut references = context.references.take();
    let mut empty_references = EmptyChartReferenceResolver;
    let active_references = references.as_deref_mut().unwrap_or(&mut empty_references);
    let chart = parser(
        root,
        color_resolver,
        context.style_xml,
        context.color_style_xml,
        active_references,
        images,
    );
    context.references.set(references);
    chart
}
