import { describe, expect, it } from 'vitest';
import type { TextPlacement } from './layout/types.js';
import { assignEastAsianAutoSpace } from './line-breaker/east-asian-autospace.js';
import type { LayoutSeg, LayoutTextSeg } from './line-breaker/model.js';
import {
  FONTS, advancePt, layoutStubParagraph, retainStubParagraph,
} from './test-support/word-space-fit.test-support.js';

// WORD_EAST_ASIAN_AUTO_SPACE (§17.3.1.2 autoSpaceDE / §17.3.1.3 autoSpaceDN):
// production segment building, line breaking and retained paint over the
// fixture faces' own hmtx advances.
const FACE = FONTS['BIZ UDGothic|700']!;
const run = (text: string, sizePt = 8.5) =>
  ({ text, ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt, bold: true }) as const;
const ON = { autoSpaceDE: true, autoSpaceDN: true } as const;
const MODE_15 = { compatibilityMode: 15, characterSpacingControl: 'doNotCompress', ...ON } as const;
const MODE_14 = { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation', ...ON } as const;
const adv = (text: string, sizePt = 8.5) => advancePt(FACE, text, sizePt);
const lineTexts = (lines: ReturnType<typeof layoutStubParagraph>) =>
  lines.map((line) => line.map((segment) => segment.text).join(''));
const lineWidths = (lines: ReturnType<typeof layoutStubParagraph>) =>
  lines.map((line) => line.reduce((sum, segment) => sum + segment.width, 0));
const wide = (runs: readonly ReturnType<typeof run>[], environment: object = MODE_15) =>
  layoutStubParagraph({ runs, environment, bandPt: 1000, justification: 'left' });

describe('WORD_EAST_ASIAN_AUTO_SPACE', () => {
  it('adds a quarter em of the East Asian side on both sides of a Latin letter or digit', () => {
    for (const [text, flags] of [
      ['甲A甲', { autoSpaceDE: true, autoSpaceDN: false }],
      ['甲1甲', { autoSpaceDE: false, autoSpaceDN: true }],
      ['ぁBｱ', { autoSpaceDE: true, autoSpaceDN: false }],
    ] as const) {
      const [line] = wide([run(text)], { ...MODE_15, ...flags });
      expect(line!.map((segment) => segment.width)).toEqual([...text].map((character, index) =>
        expect.closeTo(adv(character) + (index > 0 ? 8.5 / 4 : 0), 6)));
    }
    // The East Asian side's size defines the em, whichever side follows.
    const [mixed] = wide([run('甲', 10), run('A', 20), run('甲', 12)]);
    expect(mixed!.map((segment) => segment.width)).toEqual([
      expect.closeTo(adv('甲', 10), 6),
      expect.closeTo(adv('A', 20) + 10 / 4, 6),
      expect.closeTo(adv('甲', 12) + 12 / 4, 6),
    ]);
  });

  it('applies each flag to its own class and leaves neutral neighbours unspaced', () => {
    const natural = (text: string) => adv(text);
    for (const [text, flags] of [
      ['甲A甲', { autoSpaceDE: false, autoSpaceDN: true }],
      ['甲1甲', { autoSpaceDE: true, autoSpaceDN: false }],
      // East Asian punctuation, ASCII symbols and an authored space are neutral.
      ['甲、A', ON], ['「A」', ON], ['甲+1', ON], ['甲 A', ON], ['（A）', ON],
    ] as const) {
      expect(lineWidths(wide([run(text)], { ...MODE_15, ...flags }))[0], text)
        .toBeCloseTo(natural(text), 6);
    }
  });

  it('keeps the gap only between partners on the same line', () => {
    const text = '甲甲甲甲A甲甲';
    const lines = layoutStubParagraph({
      runs: [run(text)], environment: MODE_15, bandPt: adv('甲甲甲甲') + 1, justification: 'left',
    });
    expect(lineTexts(lines)).toEqual(['甲甲甲甲', 'A甲甲']);
    // Neither the line end before A nor the line start at A carries it.
    expect(lineWidths(lines)[0]).toBeCloseTo(adv('甲甲甲甲'), 6);
    expect(lines[1]![0]!.width).toBeCloseTo(adv('A'), 6);
    expect(lineWidths(lines)[1]).toBeCloseTo(adv('A甲甲') + 8.5 / 4, 6);
  });

  it('fits a mode-15 line with its natural gap and a mode-14 compressing line down to an eighth em', () => {
    const text = '甲A甲甲';
    const natural = adv(text) + 2 * 8.5 / 4;
    const lines = (environment: object, bandPt: number) => lineTexts(layoutStubParagraph({
      runs: [run(text)], environment, bandPt, justification: 'left',
    }));
    expect(lines(MODE_15, natural)).toEqual([text]);
    expect(lines(MODE_15, natural - 0.01)).toHaveLength(2);
    // Two gaps may each give up 1/8 em: 2.125pt in total at 8.5pt.
    const compressed = layoutStubParagraph({
      runs: [run(text)], environment: MODE_14, bandPt: natural - 2, justification: 'left',
    });
    expect(lineTexts(compressed)).toEqual([text]);
    expect(lineWidths(compressed)[0]).toBeCloseTo(natural - 2, 6);
    expect(lines(MODE_14, natural - 2.2)).toHaveLength(2);
    expect(lines({ ...MODE_14, characterSpacingControl: 'doNotCompress' }, natural - 0.01)).toHaveLength(2);
  });

  const placementsOf = (node: ReturnType<typeof retainStubParagraph>, line: number) =>
    node.lines[line]!.placements.filter((placement): placement is TextPlacement => placement.kind === 'text');

  it('paints the glyph after the gap and keeps the retained box at its measured advance', () => {
    const node = retainStubParagraph({
      runs: [run('甲A甲')], environment: MODE_15, bandPt: 1000, justification: 'left',
    });
    const [ea, latin, last] = placementsOf(node, 0);
    expect(latin!.bounds.xPt).toBeCloseTo(ea!.bounds.xPt + adv('甲'), 9);
    expect(latin!.advancePt).toBeCloseTo(adv('A') + 8.5 / 4, 9);
    expect(latin!.clusters[0]!.offset.xPt).toBeCloseTo(8.5 / 4, 9);
    expect(latin!.paintOps[0]!.offset.xPt).toBeCloseTo(8.5 / 4, 9);
    expect(last!.bounds.xPt).toBeCloseTo(latin!.bounds.xPt + latin!.advancePt, 9);
    expect(last!.clusters[0]!.offset.xPt).toBeCloseTo(8.5 / 4, 9);
  });

  it('widens the gaps first, up to half an em, when justifying', () => {
    const text = '甲A甲甲甲甲甲甲甲甲';
    const natural = adv(text) + 2 * 8.5 / 4;
    const justified = (slackPt: number) => {
      const node = retainStubParagraph({
        runs: [run(`${text}${text}`)], environment: MODE_15, bandPt: natural + slackPt,
        justification: 'both',
      });
      const placements = placementsOf(node, 0);
      expect(placements.map((placement) => placement.text).join('')).toBe(text);
      return placements;
    };
    // 3pt of slack fits in the two gaps' room (2 × 2.125pt): ideographs keep
    // their natural pitch.
    const small = justified(3);
    expect(small[2]!.bounds.xPt - small[1]!.bounds.xPt).toBeCloseTo(adv('A') + 8.5 / 4 + 1.5, 6);
    expect(small[2]!.clusters[1]!.offset.xPt - small[2]!.clusters[0]!.offset.xPt).toBeCloseTo(adv('甲'), 6);
    // Beyond that room the rest reaches the ordinary opportunities.
    const large = justified(6);
    const latinToNext = large[2]!.bounds.xPt - large[1]!.bounds.xPt;
    expect(latinToNext).toBeGreaterThan(adv('A') + 8.5 / 2);
    const pitch = large[2]!.clusters[1]!.offset.xPt - large[2]!.clusters[0]!.offset.xPt;
    expect(pitch).toBeGreaterThan(adv('甲'));
  });

  it('admits a justified ideograph by compressing the gaps in mode 15', () => {
    const text = '甲A甲A甲甲甲甲甲甲甲甲';
    // Four gaps (S = 8.5pt); the tenth character overflows by 1pt <= S/4.
    const head = [...text].slice(0, 10).join('');
    const band = adv(head) + 4 * 8.5 / 4 - 1;
    const first = (justification: 'both' | 'left') => placementsOf(retainStubParagraph({
      runs: [run(text)], environment: MODE_15, bandPt: band, justification,
    }), 0).map((placement) => placement.text).join('');
    expect(first('both')).toBe(head);
    expect(first('left')).toBe([...head].slice(0, 9).join(''));
  });

  it('shifts punctuation-compressed glyph groups by the gap once', () => {
    const node = retainStubParagraph({
      runs: [run('A甲、甲')], environment: MODE_14, bandPt: 1000, justification: 'left',
    });
    const ea = placementsOf(node, 0).find((placement) => placement.text.startsWith('甲'))!;
    expect(ea.clusters[0]!.offset.xPt).toBeCloseTo(8.5 / 4, 9);
    expect(ea.paintOps[0]!.offset.xPt).toBeCloseTo(ea.clusters[0]!.offset.xPt, 9);
  });

  it('leaves justified lines without an automatic gap on their established fit', () => {
    const lines = retainStubParagraph({
      runs: [run('甲甲 乙乙乙乙')], environment: { ...MODE_15, autoSpaceDE: false, autoSpaceDN: false },
      bandPt: 29, justification: 'both',
    }).lines;
    expect(lines).toHaveLength(3);
  });

  it('keeps a compressed justified line consistent when a kinsoku leader is forced onto it', () => {
    // The line is admitted by compressing its gaps; the closing mark that may
    // not start a line then joins it. Planning must follow the closed line.
    const node = retainStubParagraph({
      runs: [run('甲A甲 ）')], environment: MODE_15, bandPt: 24.5, justification: 'both',
    });
    // Gaps shrink by the accepted amount only, never below zero.
    for (const line of node.lines) {
      const placements = placementsOf(node, node.lines.indexOf(line));
      for (const [index, placement] of placements.entries()) {
        for (const cluster of placement.clusters) expect(cluster.advancePt).toBeGreaterThan(0);
        const next = placements[index + 1];
        // Visible glyphs never overlap (a space cluster keeps its natural
        // advance; the pen carries its justified delta).
        const end = placement.clusters[[...placement.text.trimEnd()].length - 1];
        if (next && end) {
          expect(next.bounds.xPt + next.clusters[0]!.offset.xPt)
            .toBeGreaterThanOrEqual(placement.bounds.xPt + end.offset.xPt + end.advancePt - 1e-9);
        }
      }
    }
  });

  it('counts the gap inside an aligned tab cell', () => {
    const node = retainStubParagraph({
      runs: [run('\t甲A甲')], environment: MODE_15, bandPt: 200, justification: 'left',
      tabStops: [{ pos: 80, alignment: 'right', leader: 'none' }],
    });
    const last = placementsOf(node, 0).at(-1)!;
    expect(last.bounds.xPt + last.advancePt).toBeCloseTo(80, 6);
    expect(last.clusters.at(-1)!.offset.xPt + last.clusters.at(-1)!.advancePt).toBeCloseTo(last.advancePt, 6);
  });

  it('reads adjacency between grapheme-cluster bases, across and inside segments', () => {
    const segments = (...texts: string[]) => texts.map((text) =>
      ({ text, fontSize: 10, measuredWidth: 0 }) as unknown as LayoutTextSeg);
    const gaps = (segs: LayoutSeg[]) => {
      assignEastAsianAutoSpace(segs, {});
      return segs.map((segment) => [(segment as LayoutTextSeg).text, (segment as LayoutTextSeg).autoSpaceBeforePt]);
    };
    // A variation selector, a combining voiced mark and a half-width voiced
    // mark belong to their East Asian base; a decomposed Latin letter to its base.
    expect(gaps(segments('葛\u{E0100}', 'A'))).toEqual([['葛\u{E0100}', undefined], ['A', 2.5]]);
    expect(gaps(segments('か\u3099', 'A'))).toEqual([['か\u3099', undefined], ['A', 2.5]]);
    // An extender authored in its own run still belongs to the base before it.
    expect(gaps(segments('葛', '\u{E0100}', 'A'))).toEqual([['葛', undefined], ['\u{E0100}', undefined], ['A', 2.5]]);
    expect(gaps(segments('か', '\u3099', 'A'))).toEqual([['か', undefined], ['\u3099', undefined], ['A', 2.5]]);
    // A run that starts with its predecessor's mark splits after the mark.
    expect(gaps(segments('甲', '\u0301é'))).toEqual([['甲', undefined], ['\u0301', undefined], ['é', 2.5]]);
    expect(gaps(segments('ﾃﾞ', 'A'))).toEqual([['ﾃﾞ', undefined], ['A', 2.5]]);
    expect(gaps(segments('e\u0301', '漢'))).toEqual([['e\u0301', undefined], ['漢', 2.5]]);
    // A boundary inside one segment (a Latin letter on the East Asian face)
    // becomes a segment seam that owns the gap.
    expect(gaps(segments('漢é漢'))).toEqual([['漢', undefined], ['é', 2.5], ['漢', 2.5]]);
    // Full-width Latin, Greek and an empty metric host do not form a boundary;
    // the empty host does not interrupt one either.
    expect(gaps(segments('漢', 'Ａ', 'Ω', '漢'))).toEqual([
      ['漢', undefined], ['Ａ', undefined], ['Ω', undefined], ['漢', undefined]]);
    expect(gaps(segments('漢', '', 'A'))).toEqual([['漢', undefined], ['', undefined], ['A', 2.5]]);
  });
});
