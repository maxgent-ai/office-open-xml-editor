import { hex6ToRgb, type Stroke } from '@silurus/ooxml-core';

/**
 * DrawingML table (`<a:tbl>`) — adjacent-cell border conflict resolution.
 *
 * SPEC IS SILENT. ECMA-376 / ISO-29500 define the WordprocessingML analog
 * (§17.4.66 `tcBorders` conflict rules) but give NO conflict-resolution rule for
 * a DrawingML/PresentationML table: `<a:tcPr>` carries `<a:lnL>` / `<a:lnR>` /
 * `<a:lnT>` / `<a:lnB>` per cell, and when cell spacing is zero two neighbouring
 * cells each contribute a line for the shared interior gridline — but the spec
 * never says which one PowerPoint displays. The renderer previously drew BOTH,
 * so the later-painted cell won by paint order (and with a translucent line the
 * overlap doubled the ink density). This module is the pure kernel that picks a
 * single winner so each shared gridline is drawn exactly once.
 *
 * Because the spec is silent, the rules below are OUR OWN DETERMINISTIC choice
 * (they are NOT quoted from ECMA-376). They mirror the SHAPE of the docx
 * §17.4.66 resolver (see `docx/src/cell-border-conflict.ts`) but adapted to the
 * DrawingML data model, where a suppressed line (`<a:ln><a:noFill/></a:ln>`) is
 * parsed to `null` — there is no `nil`/`none` style marker to carry, so a
 * missing/suppressed edge is simply the `null` candidate.
 *
 * Rules, applied in order:
 *   0. A `null` candidate (absent OR `<a:noFill>` line) contributes nothing — the
 *      OTHER side is displayed. Both `null` ⇒ nothing (`null`).
 *   1. Both real ⇒ the WIDER line wins (larger `<a:ln w>` in EMU). PowerPoint has
 *      no style "weight" table like Word's; line width is the natural strength
 *      ordering for DrawingML strokes.
 *   2. Equal width ⇒ the DARKER colour wins (smaller perceived luminance), so the
 *      choice is deterministic and independent of paint order.
 *   3. Fully tied (equal width AND equal luminance) ⇒ the OWNING side (`a`, the
 *      cell first in reading order) is displayed. The caller passes the
 *      reading-order-first cell as `a` (interior vertical → the LEFT cell;
 *      interior horizontal → the ABOVE cell), matching the ownership convention
 *      that draws each line exactly once.
 */

/** Parse a 6- or 8-hex colour to (r,g,b). `null`/malformed ⇒ black (0,0,0). The
 *  DrawingML stroke colour is stored without a leading `#`; an 8-hex value's
 *  trailing alpha pair is ignored for the luminance comparison. */
function rgb(color: string | null | undefined): { r: number; g: number; b: number } {
  const parsed = color ? hex6ToRgb(color.replace(/^#/, '').slice(0, 6)) : null;
  if (!parsed) return { r: 0, g: 0, b: 0 };
  return { r: parsed[0], g: parsed[1], b: parsed[2] };
}

/** Perceived luminance (Rec. 601 weights). Smaller = darker. Only the ORDERING
 *  matters here, so the exact weights are not load-bearing — any monotonic
 *  darkness metric would make the tie-break deterministic. */
function luminance(color: string | null | undefined): number {
  const c = rgb(color);
  return 0.299 * c.r + 0.587 * c.g + 0.114 * c.b;
}

/**
 * Pick the winning stroke for a shared interior cell edge from the two
 * neighbouring cells' facing lines. `a` is the OWNING (reading-order-first) side;
 * it wins a total tie (rule #3). Either side may be `null` (that cell contributes
 * no line to this edge). Returns the winning stroke, or `null` when neither side
 * paints.
 *
 * NOTE: the spec is silent — see the module doc. This is a DEFINED deterministic
 * rule, not an ECMA-376 requirement.
 */
export function resolveTableBorderConflict(a: Stroke | null, b: Stroke | null): Stroke | null {
  // Rule #0 — a null candidate contributes nothing.
  if (!a && !b) return null;
  if (!a) return b;
  if (!b) return a;

  // Rule #1 — the wider line wins (EMU width).
  if (a.width !== b.width) return a.width > b.width ? a : b;

  // Rule #2 — equal width ⇒ the darker colour wins.
  const la = luminance(a.color);
  const lb = luminance(b.color);
  if (la !== lb) return la < lb ? a : b;

  // Rule #3 — fully tied ⇒ the owning (reading-order-first) side.
  return a;
}
