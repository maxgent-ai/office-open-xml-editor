// ECMA-376 §18.8.1 alignment `textRotation="255"` — Excel "stacked" vertical
// text: each character is stacked upright, one per line, top to bottom. Which
// glyphs stay literally upright and which must be substituted / rotated is decided
// by the Unicode UAX#50 Vertical_Orientation (vo) property (core
// `verticalOrientation`), the SAME classifier the pptx eaVert and docx tbRl paths
// use (issue #790) — only the geometry differs (xlsx has no page rotation).
//
//   • vo=U  (CJK / kana / Hangul / fullwidth forms & digits) → upright, drawn as
//     the caller set up (center / top). Excel stacks these upright.
//   • vo=R  (Latin letters, ASCII digits, Latin punctuation) → upright, unchanged.
//     Unlike pptx/docx eaVert (where Latin rotates), Excel stacks Latin UPRIGHT,
//     one letter per line.
//   • vo=Tu with a U+FE10–FE12 vertical form (、。，) → SUBSTITUTE that form
//     (core `verticalFormSubstitute`), drawn upright so the font hangs it in the
//     cell's upper-right corner; ！？ and small kana have no form and draw upright.
//   • vo=Tr with a U+FE1x/FE3x vertical form present in the substitute fonts — the
//     brackets （）「」〈〉【】… and the white lenticular brackets 〖〗 (#969) →
//     SUBSTITUTE that form (core `verticalBracketFormSubstitute`), drawn upright
//     (UAX#50 §5 "substitute a vertical glyph, rotate only as fallback"). XLSX has
//     no Excel ground-truth image, so 〖〗 follow the Word/PowerPoint verdict.
//   • vo=Tr with NO substituted form → a geometric fallback: plain ROTATE 90° CW
//     (quotes, colon, and unreachable long marks ー 〜 ～; a rotation turns the base
//     colon's stacked dots into FE13's side-by-side dots), or UPRIGHT
//     (the semicolon ；, whose FE14 form is an upright dot-over-comma, not a rotation;
//     core `verticalTrUprightFallback`). FE13/FE14 are absent from most render fonts,
//     so ：；take this geometric path rather than substitution (issue #969 follow-up).

import {
  verticalOrientation,
  verticalFormSubstitute,
  verticalBracketFormSubstitute,
  verticalTrUprightFallback,
  verticalTrLongMark,
  withVertFeature,
} from '@silurus/ooxml-core';

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/**
 * Draw one character of an Excel stacked (`textRotation="255"`) cell into its
 * vertical cell slot. The caller stacks the cells top-to-bottom and sets
 * `textAlign='center'` / `textBaseline='top'`; this routes each glyph through the
 * UAX#50 classifier to draw it upright, substituted, or rotated.
 *
 * Substitution and rotation are GLYPH-only: the caller's cell height / advance and
 * everything the model reports (selection, find) keep the ORIGINAL character.
 *
 * @param ctx       2D context; `ctx.font`/`ctx.fillStyle` and (for the upright
 *                  path) `textAlign='center'` / `textBaseline='top'` are the
 *                  caller's.
 * @param ch        One character (single code point) of the cell text.
 * @param centerX   Horizontal centre of the cell (px) — the upright glyph's x.
 * @param cellTopY  Top y of this character's cell slot (px) — the upright glyph's
 *                  `top`-baseline y.
 * @param charH     Height of one stacked cell slot (px) — used to centre a rotated
 *                  glyph within its slot.
 * @param paintText Optional cell painter, used after glyph substitution and
 *                  rotation so synthetic bold follows the same geometry.
 */
export function drawStackedVerticalChar(
  ctx: Ctx2D,
  ch: string,
  centerX: number,
  cellTopY: number,
  charH: number,
  vertCapable = false,
  paintText: (text: string, x: number, y: number) => void = (text, x, y) => ctx.fillText(text, x, y),
): void {
  const cp = ch.codePointAt(0) ?? 0;
  const vo = verticalOrientation(cp);

  if (vertCapable && verticalTrLongMark(cp)) {
    ctx.save();
    ctx.translate(centerX, cellTopY + charH / 2);
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    withVertFeature(ctx, () => paintText(ch, 0, 0));
    ctx.restore();
    return;
  }

  if (vo === 'Tr') {
    // Substitute-first: a fullwidth bracket or white lenticular bracket 〖〗 (#969)
    // has a U+FE3x/FE1x vertical form drawn UPRIGHT (UAX#50 §5). It fills the cell
    // like an ideograph, so the caller's center/top placement lands it correctly.
    const bracket = verticalBracketFormSubstitute(cp);
    if (bracket !== null) {
      paintText(String.fromCodePoint(bracket), centerX, cellTopY);
      return;
    }
    // Semicolon ；(FF1B): its FE14 vertical form is an upright dot-over-comma (not a
    // rotation), so draw it UPRIGHT like an ideograph — NOT rotated (issue #969
    // follow-up; FE13/FE14 were dropped from the core substitute map because those
    // forms are absent from most render fonts). The colon ：falls through to rotate,
    // reproducing FE13's side-by-side dots.
    if (verticalTrUprightFallback(cp)) {
      paintText(ch, centerX, cellTopY);
      return;
    }
    // Fallback: no vertical form (ー U+30FC, wave dash / tilde 〜 ～, quotes, colon ：)
    // → rotate 90° CW, centred in the cell slot. save/restore also restores the
    // caller's textAlign/textBaseline.
    //
    // Unreachable `ー〜～` use this plain rotation too. The worker/skia fallback
    // cannot verify the font's inaccessible vertical design, so it adds no mirror
    // or shear approximation.
    ctx.save();
    ctx.translate(centerX, cellTopY + charH / 2);
    ctx.rotate(Math.PI / 2);
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    paintText(ch, 0, 0);
    ctx.restore();
    return;
  }

  // vo=U / vo=R → upright unchanged; vo=Tu → substitute a vertical form when one
  // exists (、。，), else upright unchanged (！？, small kana).
  const sub = vo === 'Tu' ? verticalFormSubstitute(cp) : null;
  paintText(sub !== null ? String.fromCodePoint(sub) : ch, centerX, cellTopY);
}
