import { describe, expect, it } from 'vitest';
import { mixedSpaceFitTextOutsideScope } from './line-breaker/segment-builder.js';
import { layoutStubParagraph } from './test-support/word-space-fit.test-support.js';

// Scope of WORD_COMPRESSED_SPACE_LINE_FIT where ECMA-376 §17.3.1.2
// autoSpaceDE / §17.3.1.3 autoSpaceDN automatic spacing applies. The renderer
// does not model that spacing, so these paragraphs keep the unchanged line
// breaker; no Word spacing amount is asserted here.
describe('WORD_COMPRESSED_SPACE_LINE_FIT automatic-spacing scope', () => {
  // The fixture faces cover only the Word controls' characters, so cluster
  // cases with other code points exercise the production scope predicate.
  it.each([
    ['an ideograph with a variation selector', ['葛\u{E0100}A']],
    ['a decomposed voiced kana', ['か\u3099A']],
    ['a half-width kana with its voiced sound mark', ['ﾃﾞA']],
    ['a decomposed Latin letter before an ideograph', ['e\u0301漢']],
    ['an extending character alone in a source run', ['葛', '\u{E0100}', 'A']],
  ] as const)('reads adjacency between grapheme clusters: %s', (_label, texts) => {
    expect(mixedSpaceFitTextOutsideScope(texts, undefined, undefined)).toBe(true);
    expect(mixedSpaceFitTextOutsideScope(texts, false, false)).toBe(false);
  });

  // These Han reading marks are also Grapheme_Extend. The extra cluster-base
  // check must retain every scalar adjacency the previous guard excluded.
  it.each([
    ['A\u{16FF0}'], ['\u{16FF1}A'], ['Ω\u{16FF0}A'],
  ])('preserves exclusions for script-bearing extending marks: %s', (text) => {
    expect(mixedSpaceFitTextOutsideScope([text], undefined, undefined)).toBe(true);
  });

  const run = { ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true } as const;
  const layout = (
    chunks: readonly string[],
    flags: { readonly autoSpaceDE: boolean; readonly autoSpaceDN: boolean },
    bandPt: number,
  ) => layoutStubParagraph({
    runs: chunks.map((text) => ({ ...run, text })),
    environment: { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation', ...flags },
    bandPt,
    justification: 'left',
  });

  it.each([
    ['Latin letter', '丙丙A', { autoSpaceDE: true, autoSpaceDN: false }, { autoSpaceDE: false, autoSpaceDN: true }],
    ['digit', '丙丙1', { autoSpaceDE: false, autoSpaceDN: true }, { autoSpaceDE: true, autoSpaceDN: false }],
  ] as const)('withdraws an ideograph-%s adjacency under its own flag only', (_label, tail, governing, other) => {
    const text = `甲甲甲甲 + 乙乙 + ${tail}`;
    const off = { autoSpaceDE: false, autoSpaceDN: false };
    const natural = layout([text], off, 1e7)[0]!.reduce((sum, segment) => sum + segment.width, 0);
    const bandPt = natural - 1;
    // The other flag leaves the paragraph in scope: its four spaces shrink.
    expect(layout([text], other, bandPt)).toHaveLength(1);
    // The governing flag withdraws it whether or not a source run starts at
    // the adjacency.
    for (const chunks of [[text], [text.slice(0, -1), text.slice(-1)]]) {
      const lines = layout(chunks, governing, bandPt);
      expect(lines.length, chunks.join('|')).toBeGreaterThan(1);
      expect(lines.every((line) => line.every((segment) => segment.compression === 0))).toBe(true);
    }
  });
});
