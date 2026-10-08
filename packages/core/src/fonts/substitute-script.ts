/**
 * Script scope of a *visual* web-font substitute.
 *
 * Some registry entries map an Office face to a substitute that only
 * stands in for one script. For example, Sakkal Majalla maps to Noto Naskh
 * Arabic. The substitute also carries Latin glyphs, but those glyphs, and the
 * substitute's much taller line box, have nothing to do with the authored face:
 * the authored face has its own Latin glyphs. A scoped substitute may
 * therefore paint and measure only characters of its script. Every other
 * character requested from the same family falls back exactly as if no
 * substitute existed.
 *
 * This is library font-substitution policy (ECMA-376 §17.8.2 leaves
 * substitution implementation-defined), not an Office layout rule. It is shared
 * by every format that consults {@link GOOGLE_FONT_SUBSTITUTES}.
 */
import { graphemeClusterOffsets } from '../text/sea-break.js';

export type FontSubstituteScript = 'arabic';

const FORMAT_CONTROL = /^\p{Cf}$/u;
const WHITE_SPACE = /^\s$/u;
const MARK = /^\p{M}$/u;
const ARABIC_EXTENSIONS = /^\p{Script_Extensions=Arabic}$/u;
const ARABIC_LETTER = /^(?=\p{Script=Arabic})\p{L}$/u;

/** Class of one grapheme cluster for a script-scoped substitute, by its base
 * (the first code point that is not white space or a format control):
 * - `'proof'`: a letter of the script, or a combining mark whose UAX #24
 *   Script_Extensions include it (a lone Arabic vowel sign).
 * - `'extension'`: another character whose Script_Extensions include the
 *   script: tatweel, the Arabic comma, Arabic-Indic digits. It belongs to the
 *   script's text but never proves it.
 * - `'neutral'`: white space, invisible format controls (ZWNJ, ZWJ, LRM, RLM,
 *   ALM, U+FEFF and the other Cf characters, all Joining_Type=Transparent or
 *   non-joining controls), or a lone combining mark with no script of its
 *   own. It inherits the adjacent script (UAX #24 §5.2).
 * - `'other'`: anything else, such as Latin letters, Latin digits and
 *   punctuation.
 * Classifying whole clusters means a mark is never split from its base. */
export function fontSubstituteScriptClusterClass(
  script: FontSubstituteScript,
  cluster: string,
): 'proof' | 'extension' | 'neutral' | 'other' {
  for (const character of cluster) {
    if (WHITE_SPACE.test(character) || FORMAT_CONTROL.test(character)) continue;
    switch (script) {
      case 'arabic':
        if (ARABIC_LETTER.test(character)) return 'proof';
        if (ARABIC_EXTENSIONS.test(character)) return MARK.test(character) ? 'proof' : 'extension';
        return MARK.test(character) ? 'neutral' : 'other';
    }
  }
  return 'neutral';
}

export interface FontSubstituteScopeCluster {
  readonly start: number;
  readonly end: number;
  readonly cls: 'proof' | 'extension' | 'neutral' | 'other';
  /** Whether the scoped substitute may paint and measure this cluster. */
  readonly inScope: boolean;
}

/**
 * THE scope rule for a script-scoped substitute, shared by coverage checks and
 * the shaper so they cannot diverge.
 *
 * `text` is split into grapheme clusters. A contiguous context of the script is
 * a maximal run of clusters that are not `'other'` (a caller may mark more
 * clusters `'other'` with `eligible`, for example those whose font slot does
 * not use the scoped family). A context is in scope only when it contains a
 * `'proof'` cluster. Then:
 * - its proof and extension clusters are in scope;
 * - a neutral cluster is in scope when it follows an in-scope cluster of the
 *   same context, so joiners, bidi marks and spaces continue joined text.
 * Digits, the Arabic comma or tatweel alone therefore never enable the
 * substitute, while inside proven Arabic text they stay with it.
 */
export function fontSubstituteScriptScope(
  script: FontSubstituteScript,
  text: string,
  eligible?: (start: number, end: number) => boolean,
): FontSubstituteScopeCluster[] {
  const offsets = [...new Set([0, ...graphemeClusterOffsets(text), text.length])].sort((x, y) => x - y);
  const classes = offsets.slice(0, -1).map((start, index) => {
    const end = offsets[index + 1]!;
    const cls = fontSubstituteScriptClusterClass(script, text.slice(start, end));
    return {
      start,
      end,
      cls: cls !== 'neutral' && cls !== 'other' && eligible && !eligible(start, end) ? 'other' as const : cls,
    };
  });
  const result: FontSubstituteScopeCluster[] = [];
  let contextStart = 0;
  while (contextStart < classes.length) {
    if (classes[contextStart]!.cls === 'other') {
      result.push({ ...classes[contextStart]!, inScope: false });
      contextStart += 1;
      continue;
    }
    let contextEnd = contextStart;
    while (contextEnd < classes.length && classes[contextEnd]!.cls !== 'other') contextEnd += 1;
    const proven = classes.slice(contextStart, contextEnd).some((cluster) => cluster.cls === 'proof');
    let previousInScope = false;
    for (let index = contextStart; index < contextEnd; index += 1) {
      const cluster = classes[index]!;
      const inScope: boolean = proven && (cluster.cls === 'neutral' ? previousInScope : true);
      result.push({ ...cluster, inScope });
      previousInScope = inScope;
    }
    contextStart = contextEnd;
  }
  return result;
}

/**
 * Whether a scoped substitute may supply `text` as a whole, by
 * {@link fontSubstituteScriptScope}.
 * - `'exclusive'`: no cluster is `'other'` and the text is proven. Use this for
 *   a span that can mix scripts, such as an ECMA-376 §17.3.2.26 ascii-slot span.
 * - `'any'`: some cluster proves the script. Use this for a complex-script
 *   span that the script owns as a whole, including its neutral digits and
 *   punctuation.
 */
export function fontSubstituteScriptCoversText(
  script: FontSubstituteScript,
  text: string,
  mode: 'exclusive' | 'any',
): boolean {
  const scope = fontSubstituteScriptScope(script, text);
  const proven = scope.some((cluster) => cluster.cls === 'proof');
  return mode === 'any' ? proven : proven && scope.every((cluster) => cluster.cls !== 'other');
}

/** Whether a registry entry may supply `text`. Unscoped entries always may. */
export function substituteEntryCoversText(
  entry: Readonly<{ script?: FontSubstituteScript }> | undefined,
  text: string,
  mode: 'exclusive' | 'any',
): boolean {
  return !entry?.script || fontSubstituteScriptCoversText(entry.script, text, mode);
}
