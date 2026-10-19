/**
 * Ascent/descent split of a synthetic one-em line box, as fractions of the font
 * size (or, for PowerPoint, of its line height).
 *
 * Library fallback, not a spec value and not an Office measurement. ECMA-376
 * takes a line's ascent and descent from the font (§17.3.1.33 / §21.1.2.1.2
 * line spacing is defined relative to the font's single line), and every caller
 * prefers real font metrics: Canvas `fontBoundingBox*`, a parsed resource
 * metric, or the generated reference-font-metrics table. The split is used only
 * where none of those exist — an engine without `TextMetrics` font boxes, a
 * zero-height measurement, or as the minimum line box reserved around inline
 * math whose ink is shorter than a text line. It is deliberately one shared
 * pair rather than a per-font value.
 */
export const SYNTHETIC_ASCENT_EM = 0.8;
export const SYNTHETIC_DESCENT_EM = 0.2;
