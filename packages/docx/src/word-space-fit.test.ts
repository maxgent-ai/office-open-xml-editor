import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import {
  mixedSpaceSummaryWork,
  setMixedSpaceSummaryAssertions,
} from './line-breaker/mixed-space-fit.js';
import {
  BAND_DEFICIT_PT, FONTS, VARIANTS, advancePt, layoutStubParagraph, stubMeasurements, type Variant,
} from './test-support/word-space-fit.test-support.js';

// WORD_COMPRESSED_SPACE_LINE_FIT. Every case is a synthetic Word control
// (issue #1660: 576 coarse, 595 fine and 602 round-3 cells) exported by Word
// for Mac 16.113.3 with Save As PDF; `wordWraps` is read from Word's own glyph
// origins. The stub measures each character with the hmtx advance of the face
// Word embedded (fonts[*].advances) and exposes that face's OS/2 xAvgCharWidth,
// so no Word value is restated as a renderer constant.
//
// Table geometry is not part of this rule. The controls' half-point bordered
// fixed cells end Word's line 1pt inside the authored content width (measured
// by the no-space anchors: wrap at C=85.95, fit at 86.00 for an 85pt natural
// line) and borderless cells end at the content width; the stub line width is
// that measured band.

function sourceChunks(variant: Variant): string[] {
  if (variant.sourceRuns === 'single') return [variant.text];
  return variant.text.match(/[　-鿿＀-￯]+|[^　-鿿＀-￯]+/gu) ?? [];
}

function kerning(variant: Variant): number | undefined {
  if (variant.kern === 'on') return 0;
  if (variant.kern === 'off') return variant.sizePt + 0.5;
  return undefined;
}

function lineCount(variant: Variant, widthTwips: number) {
  return layoutStubParagraph({
    runs: sourceChunks(variant).map((text) => ({
      text, ascii: variant.ascii, eastAsia: variant.eastAsia, sizePt: variant.sizePt,
      bold: variant.bold, kerning: kerning(variant),
    })),
    environment: {
      compatibilityMode: variant.compatibilityMode,
      ...(variant.characterSpacingControl
        ? { characterSpacingControl: variant.characterSpacingControl } : {}),
      enableOpenTypeFeatures: variant.enableOpenTypeFeatures,
      lineWrapLikeWord6: false,
      autoSpaceDE: variant.autoSpaceDE,
      autoSpaceDN: variant.autoSpaceDN,
    },
    bandPt: widthTwips / 20 - BAND_DEFICIT_PT[variant.borderEighths]!,
    justification: variant.justification,
  });
}

const inScope = VARIANTS.filter((variant) => variant.outOfScope === null);

// Every query and finalization compares the running line summary with a full
// recomputation from the committed items (throws on drift).
beforeAll(() => setMixedSpaceSummaryAssertions(true));
afterAll(() => setMixedSpaceSummaryAssertions(false));
const BIZ = { ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true } as const;
const MODE_14 = { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' } as const;

describe('WORD_COMPRESSED_SPACE_LINE_FIT controls', () => {
  it('covers every measured cell exactly once', () => {
    expect(VARIANTS.reduce((sum, variant) => sum + variant.widthsTwips.length, 0)).toBe(1773);
    expect(inScope.reduce((sum, variant) => sum + variant.widthsTwips.length, 0)).toBe(1443);
    // Every excluded class names a measured cause outside this rule.
    expect(new Set(VARIANTS.flatMap((variant) => variant.outOfScope
      ? [variant.outOfScope.split(':')[0]] : []))).toEqual(new Set([
      'autospace', 'kerning', 'punctuation', 'geometry', 'latin-only', 'unexplained',
    ]));
  });

  it.each(inScope.map((variant) => [
    `${variant.stage} ${variant.document} ${variant.variant}`, variant,
  ] as const))('%s matches every Word wrap outcome', (_label, variant) => {
    const actual = variant.widthsTwips
      .map((width) => (lineCount(variant, width).length > 1 ? '1' : '0'))
      .join('');
    expect(actual).toBe(variant.wordWraps);
  });

  it('fits the issue #1660 cell on one line with 3.5pt spaces', () => {
    // sample-36 A12 equivalent: 107.5pt band, OpenType features on; Word
    // paints four 3.5pt spaces.
    const [line, ...rest] = layoutStubParagraph({
      runs: [{ ...BIZ, text: '甲甲甲甲 + 乙乙 + 丙丙丙丙' }],
      environment: { ...MODE_14, enableOpenTypeFeatures: true },
      bandPt: 107.5,
      justification: 'left',
    });
    expect(rest).toHaveLength(0);
    const spaces = line!.filter((segment) => segment.text.endsWith(' '))
      .map((segment) => segment.width - advancePt(FONTS['BIZ UDGothic|700']!,
        segment.text.trimEnd(), 8.5));
    expect(spaces).toHaveLength(4);
    for (const space of spaces) expect(space).toBeCloseTo(3.5, 6);
  });

  it('fits identically when a source run starts inside the terminal cluster', () => {
    // Review probe: round-3 kana + closing-parenthesis class at its first fit
    // (102pt band). Splitting the run before the parenthesis, or before the
    // kana, must not change the partition or the retained space advances.
    const text = '甲甲甲甲  + 乙乙 +  丙丙ｱ）';
    const layout = (chunks: readonly string[]) => layoutStubParagraph({
      runs: chunks.map((chunk) => ({ ...BIZ, text: chunk })),
      environment: MODE_14,
      bandPt: 102,
      justification: 'left',
    });
    const joined = layout([text]);
    expect(joined).toHaveLength(1);
    const visible = (lines: ReturnType<typeof layout>) => lines.map((line) =>
      line.map((segment) => segment.text).join(''));
    const total = (lines: ReturnType<typeof layout>) => lines.map((line) =>
      line.reduce((sum, segment) => sum + segment.width, 0));
    for (const split of [[text.slice(0, -1), text.slice(-1)], [text.slice(0, -2), text.slice(-2)]]) {
      const lines = layout(split);
      expect(visible(lines)).toEqual(visible(joined));
      expect(total(lines)[0]).toBeCloseTo(total(joined)[0]!, 9);
    }
    // One pt narrower Word wraps; joined and split agree there too.
    const narrow = (chunks: readonly string[]) => layoutStubParagraph({
      runs: chunks.map((chunk) => ({ ...BIZ, text: chunk })),
      environment: MODE_14, bandPt: 101, justification: 'left',
    }).map((line) => line.map((segment) => segment.text).join(''));
    expect(narrow([text.slice(0, -1), text.slice(-1)])).toEqual(narrow([text]));
  });
});

// Deterministic PRNG (mulberry32) so property cases are reproducible.
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

/** Split points inside visible text only: a U+0020 at a source-run boundary
 * has its own registered seam rule (WORD_SOURCE_RUN_SPACE_SEQUENCE). */
function randomChunks(text: string, next: () => number): string[] {
  const characters = [...text];
  const legal = characters.flatMap((character, index) => index > 0
    && character !== ' ' && characters[index - 1] !== ' ' ? [index] : []);
  const cuts = [...new Set(Array.from({ length: 1 + Math.floor(next() * 3) },
    () => legal[Math.floor(next() * legal.length)]!))].sort((a, b) => a - b);
  return [0, ...cuts].map((start, index) =>
    characters.slice(start, cuts[index] ?? characters.length).join(''));
}

describe('WORD_COMPRESSED_SPACE_LINE_FIT properties', () => {
  const mixed = inScope.filter((variant) => variant.sourceRuns === 'single'
    && /[　-鿿＀-￯]/u.test(variant.text));
  const paragraph = (variant: Variant, chunks: readonly string[], widthTwips: number) =>
    layoutStubParagraph({
      runs: chunks.map((text) => ({
        text, ascii: variant.ascii, eastAsia: variant.eastAsia, sizePt: variant.sizePt,
        bold: variant.bold, kerning: kerning(variant),
      })),
      environment: {
        compatibilityMode: variant.compatibilityMode,
        ...(variant.characterSpacingControl
          ? { characterSpacingControl: variant.characterSpacingControl } : {}),
        enableOpenTypeFeatures: variant.enableOpenTypeFeatures,
        autoSpaceDE: variant.autoSpaceDE,
        autoSpaceDN: variant.autoSpaceDN,
      },
      bandPt: widthTwips / 20 - BAND_DEFICIT_PT[variant.borderEighths]!,
      justification: variant.justification,
    });
  const summary = (lines: ReturnType<typeof paragraph>) => lines.map((line) => ({
    text: line.map((segment) => segment.text).join(''),
    // Per-segment widths are rounded to 1e-6pt by the stub; compare sums at 1e-4.
    width: Number(line.reduce((sum, segment) => sum + segment.width, 0).toFixed(4)),
    compression: Number(line.reduce((sum, segment) => sum + segment.compression, 0).toFixed(4)),
  }));

  it('is invariant under source-run seams inside visible text', () => {
    const next = random(1660);
    let compared = 0;
    for (const variant of mixed) {
      for (const width of variant.widthsTwips) {
        const joined = summary(paragraph(variant, [variant.text], width));
        for (let trial = 0; trial < 3; trial += 1) {
          const chunks = randomChunks(variant.text, next);
          expect(summary(paragraph(variant, chunks, width)), `${variant.variant} ${width} ${chunks.join('|')}`)
            .toEqual(joined);
          compared += 1;
        }
      }
    }
    expect(compared).toBeGreaterThan(3000);
  }, 120_000);

  it('is invariant under seams between and beside terminal closing marks', () => {
    // Review round 5: the terminal cluster (closing marks after the final
    // core character) is excluded alike whether committed or in the candidate.
    const run = { ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true } as const;
    const lines = (chunks: readonly string[], bandPt: number) => layoutStubParagraph({
      runs: chunks.map((text) => ({ ...run, text })),
      environment: { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' },
      bandPt, justification: 'left',
    });
    const view = (result: ReturnType<typeof lines>) => result.map((line) => ({
      text: line.map((segment) => segment.text).join(''),
      width: Number(line.reduce((sum, segment) => sum + segment.width, 0).toFixed(4)),
    }));
    let compared = 0;
    for (const tail of ['ｱ）', 'ｱ））', 'ｱ）））', '丙）', '丙））', '丙）））']) {
      const text = `甲甲甲甲  + 乙乙 +  丙丙${tail}`;
      const characters = [...text];
      for (let bandPt = 94; bandPt <= 112; bandPt += 0.25) {
        const joined = view(lines([text], bandPt));
        for (let cut = characters.length - tail.length - 1; cut < characters.length; cut += 1) {
          const split = [characters.slice(0, cut).join(''), characters.slice(cut).join('')];
          expect(view(lines(split, bandPt)), `${tail} ${bandPt} ${split.join('|')}`).toEqual(joined);
          compared += 1;
        }
      }
    }
    expect(compared).toBeGreaterThan(1000);
  }, 120_000);

  it('decides on the joined character sequence for random seams in mixed punctuation', () => {
    // Review round 6: opening brackets, U+3000, small kana, prolonged sound
    // marks and full-/half-width punctuation, split at random positions over
    // the whole text. Wherever the rule takes part (either form shrinks a
    // space) the partition and every line width must not depend on the seam.
    const run = { ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true } as const;
    const alphabet = [...'丙ｱ（「『【）」』】　ぁゃッー、。，．()A1ｰｧ'];
    const next = random(6060);
    const layout = (chunks: readonly string[], bandPt: number) => layoutStubParagraph({
      runs: chunks.map((text) => ({ ...run, text })),
      environment: { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' },
      bandPt, justification: 'left',
    });
    const view = (lines: ReturnType<typeof layout>) => lines.map((line) => ({
      text: line.map((segment) => segment.text).join(''),
      width: Number(line.reduce((sum, segment) => sum + segment.width, 0).toFixed(4)),
    }));
    const shrinks = (lines: ReturnType<typeof layout>) =>
      lines.some((line) => line.some((segment) => segment.compression > 0));
    let engaged = 0;
    for (let trial = 0; trial < 2500; trial += 1) {
      const tail = Array.from({ length: 2 + Math.floor(next() * 5) },
        () => alphabet[Math.floor(next() * alphabet.length)]!).join('');
      const text = `甲甲甲甲  + 乙乙 +  丙丙${tail}`;
      const characters = [...text];
      const cut = 1 + Math.floor(next() * (characters.length - 1));
      const split = [characters.slice(0, cut).join(''), characters.slice(cut).join('')];
      const bandPt = 92 + Math.round(next() * 88) / 4;
      const joined = layout([text], bandPt);
      const separate = layout(split, bandPt);
      if (!shrinks(joined) && !shrinks(separate)) continue;
      engaged += 1;
      expect(view(separate), `${bandPt} ${split.join('|')}`).toEqual(view(joined));
    }
    expect(engaged).toBeGreaterThan(300);
  }, 120_000);

  it('keeps the measured Latin-terminal control seam-invariant', () => {
    // Review round 2: `ABCD` split as `A` / `BCD` at the 106.25pt band.
    const variant = mixed.find((item) => item.variant === 'cap-latin-word')!;
    const text = variant.text;
    const split = [text.slice(0, -3), text.slice(-3)];
    for (const width of variant.widthsTwips) {
      expect(summary(paragraph(variant, split, width))).toEqual(
        summary(paragraph(variant, [text], width)));
    }
    // Word's first fit (authored 107.25pt, the measured 106.25pt band).
    expect(paragraph(variant, split, 2145)).toHaveLength(1);
  });

  it('shrinks exactly the overflow, equally per space, never below the floor', () => {
    let compressedLines = 0;
    for (const variant of mixed) {
      for (const width of variant.widthsTwips) {
        const band = width / 20 - BAND_DEFICIT_PT[variant.borderEighths]!;
        for (const line of paragraph(variant, [variant.text], width)) {
          const placed = line.reduce((sum, segment) => sum + segment.width, 0);
          const reduction = line.reduce((sum, segment) => sum + segment.compression, 0);
          if (reduction === 0) {
            expect(placed).toBeLessThanOrEqual(band + 1e-9);
            continue;
          }
          compressedLines += 1;
          // Placed advances end exactly at the band: no excess contraction.
          expect(placed).toBeCloseTo(band, 5);
          const perSpace = line.filter((segment) => segment.compression > 0)
            .map((segment) => segment.compression / (segment.text.length - segment.text.trimEnd().length));
          for (const value of perSpace) expect(value).toBeCloseTo(perSpace[0]!, 5);
        }
      }
    }
    expect(compressedLines).toBeGreaterThan(300);
  }, 120_000);

  it('accounts spaces committed after the line was first shrunk', () => {
    // Review round 2: a gap added after an admission must not restore a
    // reduction that was never applied to it. 甲 admits by shrinking, then
    // `+ ` (a new gap) and 乙 follow on the same run.
    const run = { ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true } as const;
    const text = '甲甲甲甲 + 乙乙 + 丙丙丙丙';
    for (const chunks of [
      [text], ['甲甲甲甲 + 乙乙 ', '+ 丙丙丙丙'], ['甲甲甲甲 +', ' 乙乙 + 丙丙丙丙'],
      // Space-only runs committed after the admitting ideograph.
      [text, ' ', ' '], [`${text} `, '丙'],
    ]) {
      const lines = layoutStubParagraph({
        runs: chunks.map((chunk) => ({ ...run, text: chunk })),
        environment: { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' },
        bandPt: 106.25, justification: 'left',
      });
      // Line one holds the whole text; its visible advance ends exactly at the
      // band, a line-end space keeps its natural advance, and every inner
      // space shrinks by the same 1.0625pt (4.25pt over four spaces).
      const first = lines[0]!;
      expect(first.map((segment) => segment.text).join('').trimEnd()).toBe(text);
      const end = first.at(-1)!;
      const lineEndSpace = end.compression === 0
        ? 4.25 * (end.text.length - end.text.trimEnd().length) : 0;
      const placed = first.reduce((sum, segment) => sum + segment.width, 0) - lineEndSpace;
      expect(placed).toBeCloseTo(106.25, 5);
      const inner = first.filter((segment) => segment.compression > 0);
      expect(inner).toHaveLength(4);
      for (const segment of inner) expect(segment.compression).toBeCloseTo(1.0625, 9);
    }
  });
});

describe('WORD_COMPRESSED_SPACE_LINE_FIT work', () => {
  it('reads the committed line a bounded number of times per commit', () => {
    // Review round 3 stress shape: `甲 ` + n × `AB ` + `AB` on one line whose
    // band is the natural advance minus 85% of the quarter-em space capacity,
    // so every late candidate is admitted by shrinking.
    setMixedSpaceSummaryAssertions(false);
    try {
      const run = { ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true } as const;
      const work = (count: number) => {
        const text = `甲 ${'AB '.repeat(count)}AB`;
        const runs = [{ ...run, text }];
        const environment = { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' };
        const natural = layoutStubParagraph({ runs, environment, bandPt: 1e7, justification: 'left' })[0]!
          .reduce((sum, segment) => sum + segment.width, 0);
        const capacity = (count + 1) * (4.25 - 8.5 / 4);
        mixedSpaceSummaryWork(true);
        const lines = layoutStubParagraph({
          runs, environment, bandPt: natural - 0.85 * capacity, justification: 'left',
        });
        expect(lines).toHaveLength(1);
        return mixedSpaceSummaryWork(true);
      };
      const small = work(250);
      const large = work(1000);
      // Linear: 4n words need about 4x the reads (a full rescan per query
      // would need about 16x).
      expect(large / small).toBeLessThan(5);
      expect(large).toBeLessThan(40 * 1000);

      // Review round 4: whitespace-heavy lines of separate space-only runs,
      // with and without East Asian text, fitting naturally or by shrinking.
      const spaced = (count: number, lead: string, shrink: boolean) => {
        const runs = [
          ...(lead ? [{ ...run, text: lead }] : []),
          ...Array.from({ length: count }, (_, index) => ({ ...run, text: index % 2 ? ' ' : 'AB' })),
          { ...run, text: 'AB' },
        ];
        const environment = { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' };
        const natural = layoutStubParagraph({ runs, environment, bandPt: 1e7, justification: 'left' })[0]!
          .reduce((sum, segment) => sum + segment.width, 0);
        mixedSpaceSummaryWork(true);
        layoutStubParagraph({
          runs, environment, justification: 'left',
          bandPt: shrink ? natural - 0.85 * (count / 2) * (4.25 - 8.5 / 4) : natural + 10,
        });
        return mixedSpaceSummaryWork(true);
      };
      // Leading whitespace: a line holding only separate space runs.
      const leading = (count: number, lead: string) => {
        const runs = [
          ...Array.from({ length: count }, () => ({ ...run, text: ' ' })),
          { ...run, text: `${lead}AB` },
        ];
        mixedSpaceSummaryWork(true);
        layoutStubParagraph({
          runs, justification: 'left', bandPt: 50,
          environment: { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' },
        });
        return mixedSpaceSummaryWork(true);
      };
      for (const lead of ['甲', '']) {
        expect(leading(2000, lead) / leading(500, lead), `leading ${lead || 'latin'}`).toBeLessThan(5);
      }

      // Review round 5: runs of closing punctuation. Outside the rule's scope
      // (mode 15, omitted mode, doNotCompress) the projection does no work;
      // inside it the extra work grows linearly with the text.
      const closing = (count: number, lead: string, environment: Record<string, unknown>) => {
        mixedSpaceSummaryWork(true);
        layoutStubParagraph({
          runs: [{ ...run, text: `${lead}AB ` }, { ...run, text: '）'.repeat(count) }],
          environment, bandPt: 100, justification: 'left',
        });
        return mixedSpaceSummaryWork(true);
      };
      for (const environment of [
        { compatibilityMode: 15, characterSpacingControl: 'compressPunctuation' },
        { characterSpacingControl: 'compressPunctuation' },
        { compatibilityMode: 14, characterSpacingControl: 'doNotCompress' },
      ]) {
        expect(closing(1000, '甲', environment), JSON.stringify(environment)).toBe(0);
      }
      for (const lead of ['甲', '']) {
        const inScope14 = { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' };
        const n = closing(250, lead, inScope14);
        const n4 = closing(1000, lead, inScope14);
        expect(n4 / n, `closing ${lead || 'latin'}`).toBeLessThan(6);
      }
      for (const lead of ['甲', '']) {
        for (const shrink of [true, false]) {
          const n = spaced(500, lead, shrink);
          const n4 = spaced(2000, lead, shrink);
          expect(n4 / n, `${lead || 'latin'} shrink=${shrink}`).toBeLessThan(5);
          expect(n4, `${lead || 'latin'} shrink=${shrink}`).toBeLessThan(20 * 2000);
        }
      }
    } finally {
      setMixedSpaceSummaryAssertions(true);
    }
  }, 120_000);
});

describe('WORD_COMPRESSED_SPACE_LINE_FIT scope cost', () => {
  it('measures Latin-only paragraphs exactly as often as outside the rule', () => {
    // Review round 6: Latin-only text in scope (mode 14 + compressPunctuation)
    // follows main's path. Mode 15 keeps the same Latin projection and has no
    // mixed-script rule, so its glyph-measurement count is main's.
    const run = { ascii: 'Arial', eastAsia: 'Arial', sizePt: 8.5, bold: true } as const;
    const measure = (mode: number, runs: readonly string[], bandPt: number) => {
      stubMeasurements(true);
      layoutStubParagraph({
        runs: runs.map((text) => ({ ...run, text })),
        environment: { compatibilityMode: mode, characterSpacingControl: 'compressPunctuation' },
        bandPt, justification: 'left', freshServices: true,
      });
      return stubMeasurements(true);
    };
    for (const runs of [
      [`${'AV To '.repeat(400)}AV`],
      Array.from({ length: 400 }, (_, index) => (index % 2 ? 'To ' : 'A')),
      Array.from({ length: 400 }, (_, index) => (index % 3 ? ' ' : 'AV')),
    ]) {
      for (const bandPt of [100, 233.3, 468]) {
        const latin14 = measure(14, runs, bandPt);
        expect(latin14).toBeGreaterThan(0);
        expect(latin14, `${runs.length} runs, ${bandPt}pt`).toBe(measure(15, runs, bandPt));
      }
    }
  }, 120_000);
});

describe('WORD_COMPRESSED_SPACE_LINE_FIT and U+3000', () => {
  // Review round 7: no Word control measured U+3000, so a paragraph holding
  // one keeps compression disabled. Joined expectations were produced by
  // origin/main 4a387ebcb (BIZ UDGothic bold 8.5pt, mode 14,
  // compressPunctuation); formatting-only seams use that same joined input.
  const MAIN: readonly Readonly<{
    chunks: readonly string[];
    band: number;
    main: readonly (readonly [string, number])[];
  }>[] = [{"chunks": ["甲甲甲甲  + 乙乙 +  丙丙、）　　("], "band": 108.5, "main": [["甲甲甲甲  + 乙乙 +  丙", 93.5], ["丙、）　　(", 38.25]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙、）　", "　("], "band": 108.5, "main": [["甲甲甲甲  + 乙乙 +  丙", 93.5], ["丙、）　　(", 38.25]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙　(，　【ーｱ，」　"], "band": 109.25, "main": [["甲甲甲甲  + 乙乙 +  丙丙　", 110.5], ["(，　【ーｱ，」　", 55.25]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙　(，　", "【ーｱ，」　"], "band": 109.25, "main": [["甲甲甲甲  + 乙乙 +  丙丙　", 110.5], ["(，　【ーｱ，」　", 55.25]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙　(，　【", "ーｱ，」　"], "band": 109.25, "main": [["甲甲甲甲  + 乙乙 +  丙丙　", 110.5], ["(，　【ーｱ，」　", 55.25]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙　(，　【ー", "ｱ，」　"], "band": 109.25, "main": [["甲甲甲甲  + 乙乙 +  丙丙　", 110.5], ["(，　【ーｱ，」　", 55.25]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙　(，　【ーｱ", "，」　"], "band": 109.25, "main": [["甲甲甲甲  + 乙乙 +  丙丙　", 110.5], ["(，　【ーｱ，」　", 55.25]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙　(，　【ーｱ，", "」　"], "band": 109.25, "main": [["甲甲甲甲  + 乙乙 +  丙丙　", 110.5], ["(，　【ーｱ，」　", 55.25]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙　(，　【ーｱ，」", "　"], "band": 109.25, "main": [["甲甲甲甲  + 乙乙 +  丙丙　", 110.5], ["(，　【ーｱ，」　", 55.25]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙．1，ｱ。　丙ゃ　"], "band": 120, "main": [["甲甲甲甲  + 乙乙 +  丙丙．1，", 114.75], ["ｱ。　丙ゃ　", 42.5]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙．1，ｱ。　", "丙ゃ　"], "band": 120, "main": [["甲甲甲甲  + 乙乙 +  丙丙．1，", 114.75], ["ｱ。　丙ゃ　", 42.5]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙．1，ｱ。　丙", "ゃ　"], "band": 120, "main": [["甲甲甲甲  + 乙乙 +  丙丙．1，", 114.75], ["ｱ。　丙ゃ　", 42.5]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙．1，ｱ。　丙ゃ", "　"], "band": 120, "main": [["甲甲甲甲  + 乙乙 +  丙丙．1，", 114.75], ["ｱ。　丙ゃ　", 42.5]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙】1　　『"], "band": 107.75, "main": [["甲甲甲甲  + 乙乙 +  丙丙】", 106.25], ["1　　『", 29.75]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙】1　", "　『"], "band": 107.75, "main": [["甲甲甲甲  + 乙乙 +  丙丙】", 106.25], ["1　　『", 29.75]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙　　("], "band": 108.5, "main": [["甲甲甲甲  + 乙乙 +  丙丙　　", 119], ["(", 4.25]]}, {"chunks": ["甲甲甲甲  + 乙乙 +  丙丙　", "　("], "band": 108.5, "main": [["甲甲甲甲  + 乙乙 +  丙丙　", 110.5], ["　(", 12.75]]}];
  const run = { ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true } as const;
  const groups = [...new Map(MAIN.map(item => [item.chunks.join(''), item])).keys()];
  it.each(groups)('%s keeps compression disabled with transparent source seams', text => {
    const whole = MAIN.find(item => item.chunks.length === 1 && item.chunks[0] === text)!;
    for (const item of MAIN.filter(item => item.chunks.join('') === text)) {
      const lines = layoutStubParagraph({
        runs: item.chunks.map((text) => ({ ...run, text })),
        environment: { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' },
        bandPt: item.band, justification: 'left',
      });
      expect(lines.map((line) => [
        line.map((segment) => segment.text).join(''),
        Number(line.reduce((sum, segment) => sum + segment.width, 0).toFixed(4)),
      ])).toEqual(whole.main);
      expect(lines.every((line) => line.every((segment) => segment.compression === 0))).toBe(true);
    }
  });
});

describe('WORD_COMPRESSED_SPACE_LINE_FIT outside its measured inputs', () => {
  // Targeted VRT: the 25 out-of-scope cells where the rule moved away from
  // Word. With autospace enabled beside an ideograph, or a closing mark
  // before U+0020, the paragraph keeps main's line breaker; Word and main
  // both wrap every one of these cells.
  const CELLS: readonly (readonly [string, string, string, readonly number[]])[] = [
    ['coarse', 'controls-1660-c14-coarse.docx', 'punct-medial-close', [2082, 2125]],
    ['coarse', 'controls-1660-c14-coarse.docx', 'latin-adjacent-on', [2082]],
    ['coarse', 'controls-1660-c14-coarse.docx', 'digit-adjacent-on', [2082]],
    ['fine', 'controls-1660-c14-fine-compressPunctuation.docx', 'latin-adjacent-on',
      Array.from({ length: 21 }, (_, index) => 2082 + index)],
  ];
  it.each(CELLS.flatMap(([stage, document, name, widths]) => widths.map((width) => [
    `${stage} ${name} ${width}`, stage, document, name, width,
  ] as const)))('%s wraps like main and Word', (_label, stage, document, name, width) => {
    const variant = VARIANTS.find((item) =>
      item.stage === stage && item.document === document && item.variant === name)!;
    const index = variant.widthsTwips.indexOf(width);
    expect(index).toBeGreaterThanOrEqual(0);
    expect(variant.wordWraps[index]).toBe('1');
    const lines = lineCount(variant, width);
    expect(lines.length).toBeGreaterThan(1);
    expect(lines.every((line) => line.every((segment) => segment.compression === 0))).toBe(true);
  });
});
