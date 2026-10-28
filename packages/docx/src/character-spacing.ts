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

const COMBINING_MARK = /^\p{M}$/u;
