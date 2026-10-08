import { describe, expect, it } from 'vitest';
import { layoutStackedText, type StackedParagraph, type StackedLayoutOptions } from './stacked.js';
import {
  stackedCellBoxOverride, stackedFaceBox, stackedSidewaysCharacter, stackedVerticalGlyph, STACKED_CELL_FACTOR,
} from './stacked-faces.js';

// Arial 24 pt at 1 px/pt: cell = 7/6 × usWin (1.1172 em) × 24 = 31.28.
const CELL = STACKED_CELL_FACTOR * (2288 / 2048) * 24;
const RECT = { left: 0, top: 0, width: 185.6, height: 172.8 };

type Style = { tag: string };
const S: Style = { tag: 'a' };
const para = (texts: string[], extra: Partial<StackedParagraph<Style>> = {}): StackedParagraph<Style> => ({
  runs: texts.map((text) => ({ type: 'text' as const, text, style: S })),
  emptyThickness: CELL,
  ...extra,
});
const options = (extra: Partial<StackedLayoutOptions<Style>> = {}): StackedLayoutOptions<Style> => ({
  direction: 'wordArtVert',
  rect: RECT,
  wrap: true,
  pxPerPt: 1,
  glyphs: (text, style) => [...text].map((ch) => ({
    text: ch, style, kind: 'upright' as const, advance: CELL, thickness: CELL, space: ch === ' ',
  })),
  ...extra,
});
const axes = (paragraphs: StackedParagraph<Style>[], extra: Partial<StackedLayoutOptions<Style>> = {}) =>
  layoutStackedText(paragraphs, options(extra)).columns.map((c) => c.axisX);

// Expected values follow the model; the comments give the PowerPoint export
// (issue #1626 controls, text y corrected for the export's 1.0015 scale).
describe('layoutStackedText (wordArtVert / wordArtVertRtl)', () => {
  it('stacks columns from the start side, half a cell from the edge', () => {
    // FV02 "ABCDEFGHIJKLMN" in a 172.8 pt column: 5 per column, axes 15.23 / 46.90 / 77.93.
    const l = layoutStackedText([para(['ABCDEFGHIJKLMN'])], options());
    expect(l.columns.map((c) => c.axisX)).toEqual([CELL / 2, CELL * 1.5, CELL * 2.5].map((v) => expect.closeTo(v, 6)));
    expect(l.glyphs.filter((g) => g.column === 0).map((g) => g.text).join('')).toBe('ABCDE');
    expect(l.glyphs[1].cellTop).toBeCloseTo(CELL, 6);
    // FR02: right to left from the right edge (15.65 / 46.65 / 78.30 from the right).
    expect(axes([para(['ABCDEFGHIJKLMN'])], { direction: 'wordArtVertRtl' }))
      .toEqual([185.6 - CELL / 2, 185.6 - CELL * 1.5, 185.6 - CELL * 2.5].map((v) => expect.closeTo(v, 6)));
  });

  it('wraps at spaces, keeps the trailing space in its column and starts a column per paragraph', () => {
    // FV05 "Hi all foo bar": columns "Hi␠" / "all␠" / "foo␠" / "bar".
    const l = layoutStackedText([para(['Hi all foo bar'])], options());
    const cols = [0, 1, 2, 3].map((c) => l.glyphs.filter((g) => g.column === c).map((g) => g.text).join(''));
    expect(cols).toEqual(['Hi ', 'all ', 'foo ', 'bar']);
    expect(axes([para(['AB']), para(['CDE']), para(['F'])])).toHaveLength(3);
    // wrap="none" never breaks (FV06 runs past the box bottom).
    expect(layoutStackedText([para(['ABCDEFGH'])], options({ wrap: false })).columns).toHaveLength(1);
  });

  it('puts lnSpc extra on the page-right side and paragraph spacing per direction', () => {
    const spaced = (extra: Partial<StackedParagraph<Style>>) => [para(['ABCD'], extra), para(['EFGH'], extra)];
    const pct150 = { lineSpacing: { type: 'pct' as const, val: 150000 }, spaceBefore: { type: 'pts' as const, val: 12 } };
    // FV07: 16.08, 62.89 — spcBef lies right of paragraph 2's column.
    expect(axes(spaced(pct150))).toEqual([CELL / 2, CELL * 1.5 + CELL / 2].map((v) => expect.closeTo(v, 6)));
    // FR07: 30.64, 89.71 from the right — spcBef between the paragraphs.
    const r = axes(spaced(pct150), { direction: 'wordArtVertRtl' }).map((x) => 185.6 - x);
    expect(r).toEqual([CELL, CELL * 2 + 12 + CELL / 2].map((v) => expect.closeTo(v, 6)));
    // SPV02 spcAft 12 on all three: 27.53, 70.67, 102.46 — the first paragraph's
    // spcAft lies left of its column, the last one's is suppressed.
    const aft = [para(['AB'], { spaceAfter: { type: 'pts', val: 12 } }), para(['CD'], { spaceAfter: { type: 'pts', val: 12 } }),
      para(['EF'], { spaceAfter: { type: 'pts', val: 12 } })];
    expect(axes(aft)).toEqual([12 + CELL / 2, 24 + CELL * 1.5, 24 + CELL * 2.5].map((v) => expect.closeTo(v, 6)));
    // spcFirstLastPara restores it (SPV06: 113.99 for the third column).
    expect(axes(aft, { spcFirstLastPara: true })[2]).toBeCloseTo(36 + CELL * 2.5, 6);
  });

  it('keeps the natural descent of the last line when spacing shortens the lines', () => {
    // LSVU80 lnSpc 80 %: pitches 25.22 / 24.48 / 25.22 / 26.64 (H 25.02; last H + (L - H) / 4).
    const H = CELL * 0.8;
    const five = [0, 1, 2, 3, 4].map(() => para(['HH'], { lineSpacing: { type: 'pct', val: 80000 } }));
    const v = axes(five, { rect: { ...RECT, width: 205.6 } });
    const pitches = v.slice(1).map((x, i) => x - v[i]);
    expect(pitches.slice(0, 3)).toEqual([H, H, H].map((p) => expect.closeTo(p, 6)));
    expect(pitches[3]).toBeCloseTo(H + (CELL - H) / 4, 6);
    // First column: the reduced descent 0.25 H + L / 4 (LSVU80 14.37).
    expect(v[0]).toBeCloseTo(0.25 * H + CELL / 4, 6);
    // wordArtVertRtl keeps every pitch at H; the last descent sits on the far side.
    const r = axes(five, { direction: 'wordArtVertRtl', rect: { ...RECT, width: 205.6 } });
    expect(r.slice(1).map((x, i) => r[i] - x)).toEqual([H, H, H, H].map((p) => expect.closeTo(p, 6)));
  });

  it('anchors the block across the columns and anchorCtr centres it along them', () => {
    const two = [para(['AB']), para(['CDE'])];
    // ANV3 anchor b: the last column's axis half a cell from the right.
    expect(axes(two, { anchor: 'b' })[1]).toBeCloseTo(185.6 - CELL / 2, 6);
    expect(axes(two, { anchor: 'ctr' })[0]).toBeCloseTo((185.6 - 2 * CELL) / 2 + CELL / 2, 6);
    // ANR3 anchor b in Rtl: the block sits at the left edge.
    expect(axes(two, { anchor: 'b', direction: 'wordArtVertRtl' })[1]).toBeCloseTo(CELL / 2, 6);
    // ANV4 anchorCtr: columns keep a common top, the longest one centred.
    const l = layoutStackedText(two, options({ anchorCtr: true }));
    expect(l.glyphs[0].cellTop).toBeCloseTo((172.8 - 3 * CELL) / 2, 6);
    expect(l.glyphs[2].cellTop).toBeCloseTo(l.glyphs[0].cellTop, 6);
  });

  it('aligns along each column; just stretches spaces, dist spreads glyphs', () => {
    const tops = (algn: string, text = 'BCD') =>
      layoutStackedText([para([text], { alignment: algn })], options({ rect: { ...RECT, height: 212.8 } }))
        .glyphs.filter((g) => !g.space).map((g) => g.cellTop);
    expect(tops('r')[2]).toBeCloseTo(212.8 - CELL, 6); // ALV3: last glyph at the bottom
    expect(tops('ctr')[0]).toBeCloseTo((212.8 - 3 * CELL) / 2, 6);
    expect(tops('dist')).toEqual([0, (212.8 - CELL) / 2, 212.8 - CELL].map((v) => expect.closeTo(v, 6)));
    expect(tops('dist', 'A')[0]).toBeCloseTo((212.8 - CELL) / 2, 6); // one glyph is centred
    // ALV6 just "AB CDE FG HIJ KL": the first column ends at the bottom through its space.
    const just = tops('just', 'AB CDE FG HIJ KL');
    expect(just[4]).toBeCloseTo(212.8 - CELL, 6);
    expect(just[1]).toBeCloseTo(CELL, 6);
  });
});

describe('layoutStackedText input size', () => {
  it('lays out a 150,000-glyph run without spreading it into an argument list', () => {
    const text = 'A'.repeat(150_000);
    const one = layoutStackedText([para([text])], options({ wrap: false }));
    expect(one.glyphs).toHaveLength(150_000);
    expect(one.columns[0].thickness).toBeCloseTo(CELL, 6);
    const centred = layoutStackedText([para([text])], options({ wrap: false, anchorCtr: true }));
    expect(centred.glyphs[0].cellTop).toBeCloseTo((172.8 - 150_000 * CELL) / 2, 3);
    // Wrapped into columns of five cells: 30,000 columns in one paragraph.
    const wrapped = layoutStackedText([para(['A '.repeat(75_000)])], options());
    expect(wrapped.columns.length).toBeGreaterThan(20_000);
  }, 60_000);
});

describe('stacked face and character tables', () => {
  it('uses measured boxes for the faces that do not follow their font tables', () => {
    expect(stackedCellBoxOverride('游ゴシック')).toBe(1.5084);
    expect(stackedCellBoxOverride('SimSun')).toBe(1.038);
    expect(stackedCellBoxOverride('Yu Mincho')).toBeUndefined();
    const light = stackedFaceBox('Yu Mincho Light', false, false);
    expect((light?.ascent ?? 0) + (light?.descent ?? 0)).toBeCloseTo(2636 / 2048, 6);
    expect(stackedFaceBox('No Such Face', false, false)).toBeUndefined();
  });

  it('turns only the measured set sideways and knows which faces have vertical glyphs', () => {
    const cp = (c: string) => c.codePointAt(0) as number;
    for (const c of '(（「、ー…∥') expect(stackedSidewaysCharacter(cp(c)), c).toBe(true);
    for (const c of '，．：-!—ぁ〜A日') expect(stackedSidewaysCharacter(cp(c)), c).toBe(false);
    expect(stackedVerticalGlyph('Yu Gothic', cp('∥'))).toBe(true);
    expect(stackedVerticalGlyph('Meiryo', cp('∥'))).toBe(false);
    expect(stackedVerticalGlyph('ＭＳ ゴシック', cp('（'))).toBe(true);
    expect(stackedVerticalGlyph('Yu Gothic', cp('('))).toBe(false);
    expect(stackedVerticalGlyph('Arial', cp('（'))).toBe(false);
  });
});

it.each(['wordArtVert', 'wordArtVertRtl'] as const)('%s wraps an upright inline object atomically using its column extent', (direction) => {
  const object = { type: 'object' as const, width: 48, style: S, payload: 'equation' };
  const result = layoutStackedText([para([], { runs: [
    { type: 'text', text: 'A', style: S }, object,
    { type: 'text', text: 'B', style: S },
  ] })], options({ direction, rect: { ...RECT, height: CELL + 48 },
    objectGlyph: (segment) => ({ text: String(segment.payload), style: segment.style,
      kind: 'upright', advance: segment.width, thickness: 72, space: false }),
  }));
  expect(result.glyphs.map((g) => [g.text, g.column, g.cellTop])).toEqual([
    ['A', 0, 0], ['equation', 0, CELL], ['B', 1, 0],
  ]);
  expect(result.columns[0].thickness).toBe(72);
  expect(Math.sign(result.columns[1].axisX - result.columns[0].axisX))
    .toBe(direction === 'wordArtVert' ? 1 : -1);
});

it.each(['wordArtVert', 'wordArtVertRtl'] as const)('%s preserves blank display columns and zero inline advance for unavailable objects', (direction) => {
  for (const display of [false, true]) {
    const result = layoutStackedText([para([], { runs: [
      { type: 'text', text: 'A', style: S },
      { type: 'object', width: 0, style: S, display },
      { type: 'text', text: 'B', style: S },
    ] })], options({ direction, objectGlyph: ({ style }) => ({
      text: '', style, kind: 'upright', advance: 0, thickness: 0, space: false,
    }) }));
    expect(result.columns.map((c) => c.thickness)).toEqual(display ? [CELL, CELL, CELL] : [CELL]);
    expect(result.blockWidth).toBeCloseTo((display ? 3 : 1) * CELL);
    const [a, b] = result.glyphs;
    expect(result.glyphs.map((g) => [g.text, g.column, g.cellTop])).toEqual([
      ['A', 0, 0], ['B', display ? 2 : 0, display ? 0 : CELL],
    ]);
    expect(b.axisX - a.axisX).toBeCloseTo(display ? (direction === 'wordArtVert' ? 2 : -2) * CELL : 0);
  }
});
