import { WORD_FIXED_PARAGRAPH_AUTO_SPACING_STORED_MARGINS } from './line-compatibility.js';
export interface ParagraphSpacingParticipant {
  readonly contextualSpacing?: boolean;
  readonly styleId?: string | null;
}

/**
 * ECMA-376 §17.3.1.33 paragraph gap, with the per-side
 * `w:contextualSpacing` suppression from §17.3.1.9.
 *
 * The returned value is the complete distance between the preceding line block
 * and the current line block. Callers use this one rule for ordinary flow,
 * table-cell block folds, and DrawingML/WPS text boxes.
 */
export function paragraphGapPt(
  previous: ParagraphSpacingParticipant | null,
  current: ParagraphSpacingParticipant,
  previousAfterPt: number,
  currentBeforePt: number,
): number {
  if (!previous) return currentBeforePt;
  const sameStyle = !!(previous.styleId && previous.styleId === current.styleId);
  const dropPrevious = !!(sameStyle && previous.contextualSpacing);
  const dropCurrent = !!(sameStyle && current.contextualSpacing);
  if (dropPrevious && dropCurrent) return 0;
  if (dropCurrent) return previousAfterPt;
  if (dropPrevious) return Math.max(currentBeforePt - previousAfterPt, 0);
  return Math.max(previousAfterPt, currentBeforePt);
}

/** Cursor adjustment expressed from the shared total-gap authority. */
export function paragraphGapAdjustment(
  previous: ParagraphSpacingParticipant | null,
  current: ParagraphSpacingParticipant,
  previousAfterPt: number,
  currentBeforePt: number,
): { readonly suppressBefore: boolean; readonly overlap: number } {
  const gapPt = paragraphGapPt(previous, current, previousAfterPt, currentBeforePt);
  const suppressBefore = gapPt <= previousAfterPt;
  return {
    suppressBefore,
    overlap: previousAfterPt + (suppressBefore ? 0 : currentBeforePt) - gapPt,
  };
}

/**
 * ECMA-376 §17.3.1.33 delegates automatic margins to the consumer's HTML
 * paragraph policy and gives an active automatic flag priority over stored
 * before/after values (including line-unit values). The library uses the
 * standard HTML p margin of 1em on each side, with the paragraph's resolved
 * base/mark font size as the em; changing an inline run must not change its
 * parent paragraph's margin. This is a consumer policy, not a fitted Office
 * amount or a font-dependent correction. Explicit false retains stored values.
 *
 * Part 4 §14.8.3.15 specifies 5pt/10pt under the fixed compatibility setting.
 * WORD_FIXED_PARAGRAPH_AUTO_SPACING_STORED_MARGINS records the approved Word
 * deviation: that setting preserves the stored resolved pair. Missing values
 * remain zero; line-unit interpretation stays at its existing parser owner.
 * Resolve at the immutable model boundary, before both pagination and line
 * acquisition consume the same pair; paint and paragraph-gap folding do not
 * need to know about automatic flags.
 */
export function resolveAutomaticParagraphMarginsPt(
  paragraph: Readonly<{
    spaceBefore?: number;
    spaceAfter?: number;
    beforeAutospacing?: boolean;
    afterAutospacing?: boolean;
  }>,
  baseFontSizePt: number,
  fixed: boolean,
): Readonly<{ spaceBefore: number; spaceAfter: number }> {
  void WORD_FIXED_PARAGRAPH_AUTO_SPACING_STORED_MARGINS;
  if (fixed) return { spaceBefore: paragraph.spaceBefore ?? 0, spaceAfter: paragraph.spaceAfter ?? 0 };
  return {
    spaceBefore: paragraph.beforeAutospacing === true
      ? baseFontSizePt : paragraph.spaceBefore ?? 0,
    spaceAfter: paragraph.afterAutospacing === true
      ? baseFontSizePt : paragraph.spaceAfter ?? 0,
  };
}
