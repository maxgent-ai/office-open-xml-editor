// Typed accessors over the generated UCD tables. The generated arrays are raw
// data; this file is the small, hand-written, typed surface used by the engine.

import type { BidiClass } from './types.js';
import {
  BIDI_CLASS_NAMES,
  BIDI_RANGE_STARTS,
  BIDI_RANGE_CLASS,
  MIRROR_FLAT,
  BRACKET_FLAT,
} from './char-data.generated.js';

export { UNICODE_VERSION } from './char-data.generated.js';

/** Bidi_Class index (into BIDI_CLASS_NAMES) for a code point. */
function bidiClassIndex(cp: number): number {
  // Binary search for the greatest range start <= cp. Ranges are gap-free and
  // cover [0, 0x110000), so a match always exists for valid code points.
  let lo = 0;
  let hi = BIDI_RANGE_STARTS.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (BIDI_RANGE_STARTS[mid] <= cp) lo = mid;
    else hi = mid - 1;
  }
  return BIDI_RANGE_CLASS[lo];
}

/** Bidi_Class for a code point (UAX#9 / UCD DerivedBidiClass). */
export function bidiClass(cp: number): BidiClass {
  return BIDI_CLASS_NAMES[bidiClassIndex(cp)] as BidiClass;
}

const mirrorMap: Map<number, number> = (() => {
  const m = new Map<number, number>();
  for (let i = 0; i < MIRROR_FLAT.length; i += 2) {
    m.set(MIRROR_FLAT[i], MIRROR_FLAT[i + 1]);
  }
  return m;
})();

/** Bidi_Mirroring_Glyph for a code point, or null if it has no mirror. */
export function mirror(cp: number): number | null {
  return mirrorMap.get(cp) ?? null;
}

export interface BracketInfo {
  /** The matching paired bracket code point. */
  pair: number;
  /** 'o' = opening bracket, 'c' = closing bracket (UAX#9 BD14/BD15). */
  type: 'o' | 'c';
}

const bracketMap: Map<number, BracketInfo> = (() => {
  const m = new Map<number, BracketInfo>();
  for (let i = 0; i < BRACKET_FLAT.length; i += 3) {
    m.set(BRACKET_FLAT[i], {
      pair: BRACKET_FLAT[i + 1],
      type: BRACKET_FLAT[i + 2] === 0 ? 'o' : 'c',
    });
  }
  return m;
})();

/** Bidi_Paired_Bracket info for a code point, or null if not a paired bracket. */
export function bracket(cp: number): BracketInfo | null {
  return bracketMap.get(cp) ?? null;
}
