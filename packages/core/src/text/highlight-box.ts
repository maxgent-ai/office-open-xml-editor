/**
 * Vertical extents of a text-highlight (marker) background box.
 *
 * Both OOXML highlight families paint a solid rectangle behind the glyphs:
 *   - WordprocessingML §17.3.2.15 `<w:highlight>` (fixed 16-name enum)
 *   - DrawingML       §21.1.2.3.4 `<a:highlight>` (any CT_Color)
 *
 * Neither ECMA-376 nor ISO/IEC 29500 fixes the box geometry — the marker's
 * vertical band is a rendering detail of Word / PowerPoint. The coefficients
 * below (top = baseline − 0.85·em, height = 1.1·em) are library policy for the
 * pptx renderer: they date from the initial docx renderer and no Office
 * measurement is recorded for them. The DOCX layout has since moved to the
 * font-metric rule (the run's selected face ascent + descent, see
 * layout/paragraph.ts highlightBounds); adopting that rule for DrawingML text
 * would change PowerPoint rendering and needs PowerPoint controls first.
 *
 * The horizontal extent (x, width) is the glyph advance and is owned by each
 * renderer's layout, so it stays at the call site.
 *
 * @param baseline Text baseline y, in px (already includes any baseline shift).
 * @param fontPx   Run font size in px (the em the band scales against).
 * @returns `top` (box y in px) and `height` (box height in px).
 */
export function highlightBox(
  baseline: number,
  fontPx: number,
): { top: number; height: number } {
  return { top: baseline - fontPx * 0.85, height: fontPx * 1.1 };
}
