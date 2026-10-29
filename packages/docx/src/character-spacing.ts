import { joiningType } from '@silurus/ooxml-core/internal/joining-type';
import { getDefaultBidiEngine } from '@silurus/ooxml-core';

/** ECMA-376 §17.3.2.35 adds pitch after each rendered character. For run
 * tracking, combining marks (\p{M}) get no w:spacing, even when
 * §17.3.2.26 font-slot selection splits a mark from its base. Unselected U+00AD
 * gets none either (§17.3.3.29); a selected break has its own ordinary '-' glyph.
 * Grid/fit allocation remains per scalar and independent of run tracking.
 * Share eligibility across advances and retained source/cluster geometry while
 * preserving each caller's established floating-point order.
 * Iterate without materializing or filtering a code-point array per prefix. */
export function textSpacingAdvance(
  text: string,
  characterGridDeltaPx: number,
  charSpacingPx: number,
  uniformPitchPx: number,
): number {
  return countedSpacingAdvance(
    countSpacingCharacters(text), characterGridDeltaPx, charSpacingPx, uniformPitchPx,
  );
}

export function countSpacingCharacters(text: string): Readonly<{ scalars: number; pitchBearing: number }> {
  let scalars = 0;
  let pitchBearing = 0;
  for (const character of text) {
    scalars += 1;
    if (character !== '\u00AD' && !COMBINING_MARK.test(character)) pitchBearing += 1;
  }
  return { scalars, pitchBearing };
}

/** Retained geometry historically multiplies the combined pitch by an integer
 * scalar count. Preserve that operation (including the caller's original pitch)
 * when nothing is excluded: splitting grid/tracking or repeatedly adding pitch
 * changes floating-point offsets and can change native Canvas translations.
 * Prefix callers must accumulate integer counts, never fractional advances. */
export function countedSpacingAdvance(
  { scalars, pitchBearing }: Readonly<{ scalars: number; pitchBearing: number }>,
  characterGridDeltaPx: number,
  charSpacingPx: number,
  uniformPitchPx: number,
): number {
  return scalars === pitchBearing
    ? scalars * uniformPitchPx
    : scalars * characterGridDeltaPx + pitchBearing * charSpacingPx;
}

/** Advance one cluster gains from the zero floor: the amount by which its
 * scaled glyph advance plus its counted grid/run pitch falls below zero. */
export function condensedClusterFloorPx(
  text: string,
  scaledNaturalPx: number,
  characterGridDeltaPx: number,
  charSpacingPx: number,
): number {
  const raw = scaledNaturalPx
    + countedSpacingAdvance(
      countSpacingCharacters(text),
      characterGridDeltaPx,
      charSpacingPx,
      characterGridDeltaPx + charSpacingPx,
    );
  return raw < 0 ? -raw : 0;
}

/** WORD_CONDENSED_SPACING_ADVANCE_FLOOR: preceding graphemes measured with a
 * cluster for its contextual advance (pair kerning, cursive joining and
 * ligature context), so floor probes stay linear in the text length. */
export const CONDENSED_CONTEXT_GRAPHEMES = 4;

/** Whether `text` contains a cursively joining letter (Unicode Joining_Type
 * D, R or L; join controls and transparent marks alone are not letters).
 * Canvas applies no tracking inside cursive scripts, so floored cursive
 * clusters stay in their contextual paint operation. */
export function hasCursiveJoining(text: string): boolean {
  for (const character of text) {
    const type = joiningType(character.codePointAt(0)!);
    if (type === 'D' || type === 'R' || type === 'L') return true;
  }
  return false;
}

const COMBINING_MARK = /^\p{M}$/u;

/** UAX#9 X9: the level the bidi engine reports for a removed control. */
const UAX9_REMOVED_LEVEL = 255;

/** UAX#9 levels of `text` as Canvas resolves it in `direction` (no higher-
 * level overrides). A code unit removed by X9 (join controls, soft hyphen,
 * embedding controls) takes the level of the preceding retained one, or the
 * following one at the start (UAX#9 §5.2), so it never breaks a run. */
export function resolvedBidiLevels(text: string, direction: 'ltr' | 'rtl'): Uint8Array {
  const { levels, paragraphLevel } = getDefaultBidiEngine().computeLevels(text, direction);
  const resolved = Uint8Array.from(levels);
  let previous = -1;
  let pendingFrom = -1;
  for (let index = 0; index < resolved.length; index += 1) {
    if (resolved[index] === UAX9_REMOVED_LEVEL) {
      if (previous >= 0) resolved[index] = previous;
      else if (pendingFrom < 0) pendingFrom = index;
      continue;
    }
    if (pendingFrom >= 0) {
      resolved.fill(resolved[index]!, pendingFrom, index);
      pendingFrom = -1;
    }
    previous = resolved[index]!;
  }
  if (pendingFrom >= 0) resolved.fill(paragraphLevel, pendingFrom);
  return resolved;
}

/** Whether a split paint operation must end before `text[index]`: its UAX#9
 * level differs from the previous code unit's in either paint direction, or
 * either unit is ordered by a directional override (whose order only the
 * whole sequence carries, so its clusters are placed one by one). */
export function bidiLevelChanges(text: string): (index: number) => boolean {
  const ltr = resolvedBidiLevels(text, 'ltr');
  const rtl = resolvedBidiLevels(text, 'rtl');
  // An override (LRO, RLO) and the embedding of the same direction (LRE,
  // RLE) open identical levels, overflow included (UAX#9 X2-X5); the code
  // units whose resolved level differs between the two are the ones the
  // override orders.
  const embedded = text.replace(/\u202D/gu, '\u202A').replace(/\u202E/gu, '\u202B');
  const ltrEmbedded = embedded === text ? ltr : resolvedBidiLevels(embedded, 'ltr');
  const rtlEmbedded = embedded === text ? rtl : resolvedBidiLevels(embedded, 'rtl');
  const overridden = (index: number): boolean =>
    ltr[index] !== ltrEmbedded[index] || rtl[index] !== rtlEmbedded[index];
  return (index) => index > 0 && index < text.length
    && (ltr[index] !== ltr[index - 1] || rtl[index] !== rtl[index - 1]
      || overridden(index) || overridden(index - 1));
}
