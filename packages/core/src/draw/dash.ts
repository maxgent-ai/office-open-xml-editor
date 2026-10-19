import type { DrawingMLCustomDashSegment } from '../types/common.js';
import { ownValue } from '../internal/own-key.js';

/**
 * Shared border / line dash-pattern core.
 *
 * The three OOXML formats each define their own dash vocabulary in different
 * parts of the standard, and — importantly — each renders the same *logical*
 * shape (a "dash", a "dot", a "dot-dash") with **different on/off lengths**:
 *
 *   - docx  §17.18.2  ST_Border            (dotted / dashed / dotDash / …)
 *   - xlsx  §18.18.3  ST_BorderStyle       (dashed / dotted / dashDot / hair / …)
 *   - pptx  §20.1.10.49 ST_PresetLineDashVal (dot / dash / lgDash / sysDash / …)
 *           — shape/line borders (`<a:ln><a:prstDash>`)
 *   - pptx  §20.1.10.82 ST_TextUnderlineType (dotted / dash / dotDash / …)
 *           — run underlines; a distinct enum that reuses some of the same shape
 *             names as the preset line dash but is NOT the same list.
 *
 * pptx therefore carries TWO relative tables here (preset line dash + text
 * underline), each keyed on its own enum's strings.
 *
 * The standard gives no normative pixel geometry for any of them, so each
 * renderer carries Word-/Excel-/PowerPoint-like approximations whose multipliers
 * genuinely differ (e.g. a "dot" is 1·unit in docx, 2 px in xlsx, 1.5·unit in
 * pptx; a "dash" is 3·unit / 4 px / 6·unit respectively). Those values are part
 * of each format's visual contract and MUST NOT be unified — doing so would
 * change pixel output.
 *
 * What *is* shared, and what this module deduplicates, is the **structure**:
 *   1. the `[on, off, …].map(x => x * unit)` array-generation helper
 *      (`dashArray`), previously re-derived inline in three renderers, and
 *   2. the per-format relative tables, co-located here so each format's cadence
 *      lives in one place. The tables key on each format's OWN enum string
 *      (`Record<string, RelativeDashPattern>`) — there is deliberately no shared
 *      "logical type" indirection, because the format enums are not 1:1 (e.g.
 *      xlsx `mediumDashDot` and pptx `dotDash` are the "same" shape but distinct
 *      enum members with different cadences). The shape correspondence lives in
 *      the table comments, not a type.
 *
 * Each format keeps its own multipliers (its own relative table) — the helper
 * is generic over the `unit`. docx/pptx pass `unit = lw` (the stroked width, so
 * dashes scale with thickness); xlsx passes `unit = 1` because Excel cell
 * borders use a fixed pixel cadence regardless of the (sub-pixel) hairline
 * width — see `xlsxBorderDashArray`.
 */

/** A relative dash pattern: `[on, off, on, off, …]` in multiples of `unit`. */
export type RelativeDashPattern = readonly number[];

/**
 * Generic dash-array generator: scales a relative `[on, off, …]` pattern by
 * `unit`. This is the single shared implementation of what the renderers used
 * to inline (`pattern.map(v => v * lineW)` in pptx, `[lw, lw * 2]` literals in
 * docx). For static-pixel formats (xlsx) the wrapper passes `unit = 1`.
 *
 * An empty input yields an empty array (a continuous / solid stroke).
 */
export function dashArray(relative: RelativeDashPattern, unit: number): number[] {
  return relative.map((x) => x * unit);
}

// ─────────────────────────────────────────────────────────────────────────────
// docx — ECMA-376 §17.18.2 ST_Border (lw-relative)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * §17.18.2 ST_Border dash/dot families → relative `[on, off, …]` pattern in
 * units of the stroked width `lw`. Word-like: a "dot" is one `lw` square, a
 * "dash" three `lw`, gaps two `lw` (one for the small-gap variant).
 *
 * `dashDotStroked` (alternating thin/thick strokes) cannot be expressed with a
 * single setLineDash, so it is approximated as `dotDash` — noted, not exact.
 */
const DOCX_BORDER_RELATIVE: Record<string, RelativeDashPattern> = {
  dotted: [1, 2],
  dashed: [3, 2],
  dashSmallGap: [3, 1],
  dotDash: [1, 2, 3, 2],
  dotDotDash: [1, 2, 1, 2, 3, 2],
  dashDotStroked: [1, 2, 3, 2],
};

/**
 * docx ST_Border style → `setLineDash` pattern scaled by the stroked width
 * `lw`. Returns `[]` for solid styles (single/thick/triple/wave/…), which then
 * stroke as a continuous line.
 */
export function docxBorderDashArray(style: string, lw: number): number[] {
  const relative = ownValue(DOCX_BORDER_RELATIVE, style);
  return relative ? dashArray(relative, lw) : [];
}

// ─────────────────────────────────────────────────────────────────────────────
// xlsx — ECMA-376 §18.18.3 ST_BorderStyle (static px)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * §18.18.3 ST_BorderStyle dash families → a fixed-pixel `[on, off, …]` cadence.
 * Excel cell borders use a constant pixel rhythm regardless of the (sub-pixel)
 * hairline width — unlike docx/pptx these do NOT scale with the stroked width.
 * So the relative table is already in *pixels* and the wrapper scales by 1.
 *
 * "hair" is the finest dashing in Excel's border picker — a 1-px on / 1-px off
 * cadence at the hair lineWidth reproduces it (without a pattern it would read
 * as a faint solid line). The medium* variants share the dotDash/dotDotDash
 * cadence of their thin counterparts (the medium*ness is in the stroke width,
 * not the dash rhythm).
 */
const XLSX_BORDER_PX: Record<string, RelativeDashPattern> = {
  hair: [1, 1],
  dashed: [4, 3],
  mediumDashed: [4, 3],
  dotted: [2, 2],
  dashDot: [4, 2, 1, 2],
  mediumDashDot: [4, 2, 1, 2],
  dashDotDot: [4, 2, 1, 2, 1, 2],
  mediumDashDotDot: [4, 2, 1, 2, 1, 2],
  slantDashDot: [5, 3, 1, 3],
};

/**
 * xlsx ST_BorderStyle style → `setLineDash` pattern in static pixels (unit = 1,
 * see the table doc above). Returns `[]` for solid styles.
 */
export function xlsxBorderDashArray(style: string): number[] {
  const px = ownValue(XLSX_BORDER_PX, style);
  return px ? dashArray(px, 1) : [];
}

// ─────────────────────────────────────────────────────────────────────────────
// pptx — ECMA-376 §20.1.10.49 ST_PresetLineDashVal (shape/line borders, lw-relative)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * §20.1.10.49 ST_PresetLineDashVal (`<a:ln><a:prstDash val>`) → relative
 * `[on, off, …]` pattern in units of the stroked width `lineW`, so the dashes
 * scale with line thickness. The enum has 11 members; `solid` is intentionally
 * absent (it maps to `[]` — a continuous line) so this table holds the other 10.
 *
 * The values are PowerPoint-like approximations and are DELIBERATELY not the
 * spec's binary bit representations (§20.1.10.49 documents dash/dot cadences as
 * bit patterns, not pixel lengths). They are part of pptx's visual contract and
 * MUST NOT be changed. The `sys*` family uses the tighter Windows cosmetic-pen
 * cadence; `lgDash*` is the "long dash" variant.
 */
const PPTX_PRESET_DASH_RELATIVE: Record<string, RelativeDashPattern> = {
  dash: [6, 3],
  dot: [1.5, 3],
  dashDot: [6, 3, 1.5, 3],
  lgDash: [10, 4],
  lgDashDot: [10, 4, 1.5, 4],
  lgDashDotDot: [10, 4, 1.5, 4, 1.5, 4],
  sysDash: [4, 2],
  sysDot: [1, 2],
  sysDashDot: [4, 2, 1, 2],
  sysDashDotDot: [4, 2, 1, 2, 1, 2],
};

/**
 * pptx ST_PresetLineDashVal (§20.1.10.49) style → `setLineDash` pattern scaled
 * by the stroked width `lineW`. Returns `[]` for solid / continuous styles and
 * for any unknown value (table miss ⇒ solid line).
 */
export function pptxPresetDashArray(style: string, lineW: number): number[] {
  const relative = ownValue(PPTX_PRESET_DASH_RELATIVE, style);
  return relative ? dashArray(relative, lineW) : [];
}

/** Synchronous Canvas availability ceiling, kept in parity with the shared
 * OOXML parser's `MAX_LINE_DASH_STOPS`. */
const MAX_DRAWINGML_CUSTOM_DASH_STOPS = 512;

/**
 * Resolve the mutually-exclusive DrawingML dash choice. ECMA-376
 * §20.1.8.22 defines each custom dash/space length relative to the rendered
 * line width, so the same function serves shapes and all chart outlines.
 * An authored empty custom list remains a continuous stroke and still wins
 * over an inherited preset.
 */
export function drawingmlLineDashArray(
  customDash: ReadonlyArray<DrawingMLCustomDashSegment> | null | undefined,
  preset: string | null | undefined,
  lineW: number,
): number[] {
  if (customDash != null) {
    const pattern: number[] = [];
    for (let index = 0;
      index < Math.min(customDash.length, MAX_DRAWINGML_CUSTOM_DASH_STOPS);
      index += 1) {
      const segment = customDash[index];
      if (!Number.isFinite(segment.dash) || !Number.isFinite(segment.space)
        || segment.dash < 0 || segment.space < 0
        || segment.dash === 0 && segment.space === 0) continue;
      pattern.push(segment.dash * lineW, segment.space * lineW);
    }
    return pattern;
  }
  return pptxPresetDashArray(preset ?? 'solid', lineW);
}

/**
 * Shared shape-stroke dash resolver. Parsers normalize symbolic vocabularies
 * (DrawingML presets, legacy VML names) to ST_PresetLineDashVal strings before
 * they reach the shared {@link Stroke} contract. VML additionally permits a
 * custom whitespace/comma-separated numeric sequence in line-width units
 * (Part 4 §19.1.2.21); an unmatched final value is explicitly discarded.
 */
export function shapeStrokeDashArray(style: string, lineW: number): number[] {
  const preset = pptxPresetDashArray(style, lineW);
  if (preset.length > 0) return preset;

  const numeric = style.trim().split(/[\s,]+/).map(Number);
  if (
    numeric.length >= 2 &&
    numeric.every((value) => Number.isFinite(value) && value >= 0) &&
    numeric.some((value) => value > 0)
  ) {
    if (numeric.length % 2 !== 0) numeric.pop();
    return dashArray(numeric, lineW);
  }
  return [];
}

// ─────────────────────────────────────────────────────────────────────────────
// pptx — ECMA-376 §20.1.10.82 ST_TextUnderlineType (run underlines, lineW-relative)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * §20.1.10.82 ST_TextUnderlineType → relative `[on, off, …]` pattern in units of
 * the underline stroke width `lineW`, so the dashes stay proportional at any font
 * size. This is the run-underline enum (`<a:rPr u="…">`), NOT the shape/line
 * preset dash of §20.1.10.49: it shares a few shape names (dash / dotDash / …)
 * but is a distinct enumeration with its own members (e.g. `dashLong`,
 * `dottedHeavy`, and the `*Heavy` variants, which share the cadence of their base
 * name — the heaviness is in the stroke width, not the dash rhythm). Underline
 * types handled elsewhere in the renderer (sng / dbl / wavy*) are absent here and
 * map to `[]` (a continuous rule).
 */
const PPTX_UNDERLINE_RELATIVE: Record<string, RelativeDashPattern> = {
  dotted: [1.5, 3],
  dottedHeavy: [1.5, 3],
  dash: [6, 3],
  dashHeavy: [6, 3],
  dashLong: [10, 4],
  dashLongHeavy: [10, 4],
  dotDash: [6, 3, 1.5, 3],
  dotDashHeavy: [6, 3, 1.5, 3],
  dotDotDash: [6, 3, 1.5, 3, 1.5, 3],
  dotDotDashHeavy: [6, 3, 1.5, 3, 1.5, 3],
};

/**
 * pptx ST_TextUnderlineType (§20.1.10.82) style → `setLineDash` pattern scaled by
 * the underline stroke width `lineW`. Returns `[]` for solid / continuous
 * underline types (sng / dbl / wavy* / unknown).
 */
export function pptxUnderlineDashArray(style: string, lineW: number): number[] {
  const relative = ownValue(PPTX_UNDERLINE_RELATIVE, style);
  return relative ? dashArray(relative, lineW) : [];
}
