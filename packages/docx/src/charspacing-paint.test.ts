import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ─────────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.3.2.35 `<w:spacing>` / §17.3.2.43 `<w:w>` must reach the actual
// GLYPH DRAW on every paint branch, and the FOLLOWING run's pen position must
// coincide with the previous run's painted extent (measure==paint at the run
// boundary, the sample-34 overlap contract). The advance-model unit tests
// (charspacing-advance.test.ts) derive both sides from segAdvanceWidth, so they
// cannot catch a paint branch that forgets ctx.letterSpacing or advances glyphs
// by the natural width — these tests drive renderDocumentToCanvas end-to-end
// with a recording canvas and assert on the recorded fillText calls instead.
//
// Four paint branches (renderer.ts drawLine segment loop):
//   1. horizontal common single-fillText (no grid, no justify)
//   2. horizontal docGrid character-grid branch (grid pitch + char spacing)
//   3. horizontal §17.18.44 justify branch (distributed pitch + char spacing)
//   4. vertical tbRl branch (drawVerticalRun per-glyph advance)
// ─────────────────────────────────────────────────────────────────────────────

import { makeRecordingCanvas, type FillCall } from './test-support/recording-canvas.test-support.js';

const FONT_PX = 20;

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...extra,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(runs: DocxTextRun[], alignment = 'left'): BodyElement {
  const p: DocParagraph = {
    alignment, indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r }) as DocRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(extra: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 600, pageHeight: 600, marginTop: 0, marginRight: 0, marginBottom: 0,
    marginLeft: 0, headerDistance: 0, footerDistance: 0, titlePage: false,
    evenAndOddHeaders: false, ...extra,
  } as SectionProps;
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(
  body: BodyElement[],
  sec: SectionProps = section(),
): Promise<FillCall[]> {
  const { canvas, fills } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, { dpr: 1, width: 600 });
  return fills;
}

function drawOf(fills: FillCall[], text: string): FillCall {
  const f = fills.find((c) => c.text === text);
  expect(f, `a fillText drew ${JSON.stringify(text)}`).toBeDefined();
  return f as FillCall;
}

/** Effective x (page space) of the fills the vertical per-glyph draw produced:
 *  the accumulated translate at fill time. Constant page-level offsets cancel
 *  when the caller diffs consecutive values. */
function glyphCenters(fills: FillCall[], chars: string[]): number[] {
  return chars.map((ch) => drawOf(fills, ch).translateX);
}

describe('run charSpacing/charScale reach the painted glyphs on every branch', () => {
  it('paints full-width punctuation with the same compressed advance used by layout', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const model = {
      ...doc([para([
        textRun('甲乙．', { charSpacing: -0.2 }),
        textRun('次'),
      ])], section()),
      settings: { characterSpacingControl: 'compressPunctuation' },
    } as DocxDocumentModel;

    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 600 });

    const punctuation = drawOf(fills, '甲乙．');
    const following = drawOf(fills, '次');
    // The synthetic selected glyph exposes a tight 10pt trailing sidebearing:
    // (3 × 20) + (3 × -0.2) - 10 = 49.4pt. The compressed punctuation
    // stays joined to its preceding text so the browser preserves contextual
    // shaping at that boundary; only the retained trailing advance is trimmed.
    expect(punctuation.letterSpacing).toBe('-0.2px');
    expect(following.x - punctuation.x).toBeCloseTo(49.4, 5);
  });

  it('paints consecutive compressed kana and closing punctuation as one retained group', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const model = {
      ...doc([para([
        textRun('希望しない）'),
        textRun('次'),
      ])], section()),
      settings: {
        characterSpacingControl: 'compressPunctuationAndJapaneseKana',
      },
    } as DocxDocumentModel;

    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 600 });

    expect(drawOf(fills, 'しない）').letterSpacing).toBe('-10px');
    expect(fills.some(({ text }) => text === '）')).toBe(false);
    expect(drawOf(fills, '次').x).toBeCloseTo(80, 5);
  });

  it('paints a single internal punctuation compression at its retained cluster offsets', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const model = {
      ...doc([para([textRun('甲、乙')])], section()),
      settings: { characterSpacingControl: 'compressPunctuation' },
    } as DocxDocumentModel;

    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 600 });

    const first = drawOf(fills, '甲');
    const punctuation = drawOf(fills, '、');
    const following = drawOf(fills, '乙');
    expect(punctuation.letterSpacing).toBe('-10px');
    expect(punctuation.x - first.x).toBeCloseTo(20, 5);
    expect(following.x - first.x).toBeCloseTo(30, 5);
  });

  it('does not compress a contextually half-width closing parenthesis to zero advance', async () => {
    const followingGlyphLeftInsetPx = 3;
    const { canvas, fills } = makeRecordingCanvas({
      followingGlyphLeftInsetPx,
      closingPunctuationPairKerningPx: 10,
    });
    const model = {
      ...doc([para([textRun('本文。）次')])], section()),
      settings: { characterSpacingControl: 'compressPunctuation' },
    } as DocxDocumentModel;

    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 600 });

    const closingParenthesis = drawOf(fills, '）');
    const following = drawOf(fills, '次');
    const closingInkRightPx = closingParenthesis.x + FONT_PX / 2;
    const followingInkLeftPx = following.x + followingGlyphLeftInsetPx;

    expect(closingParenthesis.letterSpacing).toBe('0px');
    expect(following.x - closingParenthesis.x).toBeCloseTo(10, 5);
    expect(followingInkLeftPx).toBeGreaterThanOrEqual(closingInkRightPx);
  });

  it('preserves authored run spacing instead of layering punctuation compression over it', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const authoredText = '独自文面の配置を確認します）。';
    const model = {
      ...doc([para([
        textRun(authoredText, { charSpacing: 0.4 }),
        textRun('次'),
      ])], section()),
      settings: {
        characterSpacingControl: 'compressPunctuation',
      },
    } as DocxDocumentModel;

    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 600 });

    const authored = drawOf(fills, authoredText);
    const following = drawOf(fills, '次');
    expect(authored.letterSpacing).toBe('0.4px');
    expect(fills.some(({ text }) => text === '）')).toBe(false);
    expect(fills.some(({ text }) => text === '。')).toBe(false);
    expect(following.x - authored.x).toBeCloseTo([...authoredText].length * (FONT_PX + 0.4), 5);
  });

  it('does not collapse middle-punctuation advance into the following glyph', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const model = {
      ...doc([para([textRun('甲乙・丙丁')])], section()),
      settings: { characterSpacingControl: 'compressPunctuation' },
    } as DocxDocumentModel;

    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 600 });

    const textFills = fills.filter((fill) => /[甲乙・丙丁]/u.test(fill.text));
    // `characterSpacingControl` is not a license to remove half an em from
    // every punctuation cell. U+30FB remains outside the reviewed supported
    // subset because the public sources do not define an exhaustive character
    // set; splitting it without evidence would shift the following glyph left.
    expect(textFills).toEqual([
      expect.objectContaining({ text: '甲乙・丙丁', letterSpacing: '0px' }),
    ]);
  });

  it('retains a full-width middle-punctuation cell before a following Latin run', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const model = {
      ...doc([para([
        textRun('E-mail：'),
        textRun('kifu@example.test', { isLink: true }),
      ])], section()),
      settings: { characterSpacingControl: 'compressPunctuation' },
    } as DocxDocumentModel;

    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 600 });

    const label = drawOf(fills, 'E-mail');
    const address = drawOf(fills, 'kifu@example.test');
    expect(address.x - label.x).toBeCloseTo(7 * FONT_PX, 5);
  });

  it('horizontal common path: following run starts at the expanded right edge', async () => {
    // run1 4 cps × (20 natural + 2 spacing) = 88; run2 must start +88 from run1.
    const fills = await render([para([
      textRun('あいうえ', { charSpacing: 2 }),
      textRun('かき'),
    ])]);
    const r1 = drawOf(fills, 'あいうえ');
    const r2 = drawOf(fills, 'かき');
    expect(r1.letterSpacing).toBe('2px'); // §17.3.2.35 at the glyph draw
    expect(r2.letterSpacing).toBe('0px');
    expect(r2.x - r1.x).toBeCloseTo(4 * FONT_PX + 4 * 2, 5);
  });

  it('docGrid character-grid path: grid pitch and charSpacing combine at the draw', async () => {
    // §17.6.5 linesAndChars with charSpace 4096 units = 1 pt per EA glyph.
    // Combined per-glyph pitch = 1 (grid) + 2 (w:spacing) = 3px on run1; the
    // pure-EA run2 keeps its own 1px grid pitch. run2 starts at run1's celled
    // advance: 4 × 20 + 4 × 3 = 92.
    const fills = await render(
      [para([textRun('あいうえ', { charSpacing: 2 }), textRun('かき')])],
      section({ docGridType: 'linesAndChars', docGridLinePitch: 24, docGridCharSpace: 4096 }),
    );
    const r1 = drawOf(fills, 'あいうえ');
    const r2 = drawOf(fills, 'かき');
    expect(r1.letterSpacing).toBe('3px');
    expect(r2.letterSpacing).toBe('1px');
    expect(r2.x + r2.translateX - (r1.x + r1.translateX)).toBeCloseTo(4 * FONT_PX + 4 * 3, 5);
  });

  it('paints balanced SBCS grid pitch and advances across balanced authored spaces', async () => {
    const { canvas, fills } = makeRecordingCanvas({ spaceAdvancePx: 6 });
    const model = {
      ...doc([para([
        textRun('AB  '),
        textRun('Target'),
        textRun('日本', { fontFamilyEastAsia: 'NotInMetrics' }),
      ])], section({
        docGridType: 'linesAndChars',
        docGridLinePitch: 24,
        docGridCharSpace: -4096,
      })),
      settings: { balanceSingleByteDoubleByteWidth: true },
    } as DocxDocumentModel;

    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 600 });

    const spaced = drawOf(fills, 'AB  ');
    const target = drawOf(fills, 'Target');
    const cjk = drawOf(fills, '日本');
    // Word's §17.15.3.3 projection applies half of the -1pt grid delta to
    // SBCS and the full delta to DBCS. The two authored spaces each occupy
    // half of the selected 20pt East-Asian cell: 2×20 + 2×10 - 4×0.5 = 58.
    expect(spaced.letterSpacing).toBe('-0.5px');
    expect(target.x + target.translateX - (spaced.x + spaced.translateX)).toBeCloseTo(58, 5);
    expect(cjk.letterSpacing).toBe('-1px');
  });

  it('justify path: charSpacing adds to the distributed pitch at the draw', async () => {
    // A justified ('both') CJK paragraph that wraps: line 1 carries run1 (25 cps,
    // w:spacing 2) + a 2-cp prefix of run2, with 10px of slack distributed over
    // the inter-CJK gaps. Acquisition retains contextual text and a uniform
    // pitch. The two runs' recorded pitch differs by exactly the authored 2px.
    const fills = await render([para(
      [textRun('あ'.repeat(25), { charSpacing: 2 }), textRun('い'.repeat(10))],
      'both',
    )]);
    const r1 = fills.find((f) => f.text.startsWith('あ'));
    const r2 = fills.find((f) => f.text.startsWith('い'));
    expect(r1, 'run1 glyphs painted').toBeDefined();
    expect(r2, 'run2 first-line glyphs painted').toBeDefined();
    const p1 = parseFloat((r1 as FillCall).letterSpacing);
    const p2 = parseFloat((r2 as FillCall).letterSpacing);
    expect((r1 as FillCall).text).toBe('あ'.repeat(25));
    expect(p1 - p2).toBeCloseTo(2, 5); // the §17.3.2.35 component survives justify
  });

  it('vertical tbRl path: per-glyph advance includes charSpacing (no gap before the next run)', async () => {
    // ECMA-376 §17.6.20 rotate-layout: the logical line axis is the along-column
    // advance. layoutLines gives run1 measuredWidth = 4 × 20 + 4 × 2 = 88, so the
    // pen places run2 at 88. drawVerticalRun must advance each upright glyph by
    // measure + 2 (the §17.3.2.35 pitch); with the natural-only advance run1's
    // glyphs end at 80 and an 8px hole opens before run2.
    const fills = await render(
      [para([textRun('あいうえ', { charSpacing: 2 }), textRun('かき')])],
      section({ textDirection: 'tbRl' }),
    );
    const [a, i, u, e, ka] = glyphCenters(fills, ['あ', 'い', 'う', 'え', 'か']);
    // Cell centres of run1's glyphs are (20+2) apart.
    expect(i - a).toBeCloseTo(22, 5);
    expect(u - i).toBeCloseTo(22, 5);
    expect(e - u).toBeCloseTo(22, 5);
    // Run boundary: run2's first cell centre sits half-cells from run1's last —
    // (22 + 20) / 2 = 21 — i.e. flush, no hole (pen 88 == painted extent 88).
    expect(ka - e).toBeCloseTo((22 + 20) / 2, 5);
  });

  it('vertical tbRl path: over-condensed cells floor at zero and run2 stays flush (#1975)', async () => {
    // WORD_CONDENSED_SPACING_ADVANCE_FLOOR: 20px cells at -30 pitch advance 0,
    // so run1's cells stack and run2's pen (= run1 measuredWidth = 0) starts on
    // them. Measurement sums the same planned cells paint draws.
    const fills = await render(
      [para([textRun('あいうえ', { charSpacing: -30 }), textRun('かき')])],
      section({ textDirection: 'tbRl' }),
    );
    const [a, i, u, e, ka, ki] = glyphCenters(fills, ['あ', 'い', 'う', 'え', 'か', 'き']);
    expect(i - a).toBeCloseTo(0, 5);
    expect(u - i).toBeCloseTo(0, 5);
    expect(e - u).toBeCloseTo(0, 5);
    expect(ka - e).toBeCloseTo(10, 5);
    expect(ki - ka).toBeCloseTo(20, 5);
  });

  it('vertical tbRl path: balanced spaces floor with their width adjustment (#1975)', async () => {
    // §17.15.3.3 width balance is part of each floored cell's advance: two
    // over-condensed balanced spaces advance zero, never a negative total.
    const { canvas, fills } = makeRecordingCanvas();
    const model = {
      ...doc(
        [para([textRun('あ'), textRun('  ', { charSpacing: -30 }), textRun('い')])],
        section({ textDirection: 'tbRl' }),
      ),
      settings: { balanceSingleByteDoubleByteWidth: true },
    } as DocxDocumentModel;
    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 600 });
    const [a, i] = glyphCenters(fills, ['あ', 'い']);
    expect(i - a).toBeCloseTo(20, 5);
  });

  it('vertical tbRl path: w:w scales the along-column advance (measure==paint)', async () => {
    // §17.3.2.43 w:w in the rotate-layout vertical engine: the layout kernel is
    // direction-agnostic and has always scaled the LINE-AXIS advance by w:w
    // (segAdvanceWidth), which after the +90° page rotation IS the along-column
    // cell extent — so wrap points, run boxes, and find/selection already assume
    // it. Paint must follow measure: run1 (w:w=50%) cells are 10px, and run2's
    // pen (= run1 measuredWidth = 40) sits flush against the last cell.
    const fills = await render(
      [para([textRun('あいうえ', { charScale: 0.5 }), textRun('かき')])],
      section({ textDirection: 'tbRl' }),
    );
    const [a, i, u, e, ka] = glyphCenters(fills, ['あ', 'い', 'う', 'え', 'か']);
    expect(i - a).toBeCloseTo(10, 5);
    expect(u - i).toBeCloseTo(10, 5);
    expect(e - u).toBeCloseTo(10, 5);
    // Boundary: (10 + 20) / 2 = 15. With the unscaled per-glyph advance run1's
    // glyphs overrun their cells and run2 paints INSIDE run1's tail.
    expect(ka - e).toBeCloseTo((10 + 20) / 2, 5);
  });
});

// These literal call/advance expectations were also run against origin/main
// 4bcb91c487e38393aadb3f3b70a732adf2c7b708. They pin the pre-PR contextual
// paint boundary, independently of cluster geometry and the pitch helper.
describe('character spacing preserves main kerning and joining paint', () => {
  async function acquire(runs: DocxTextRun[]) {
    const { canvas, fills } = makeRecordingCanvas({ contextualShaping: true });
    const model = doc([para(runs)], section());
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
    }), { currentDateMs: 0 });
    const paragraph = layout.pages[0].layers.body.find(node => node.kind === 'paragraph');
    if (!paragraph || paragraph.kind !== 'paragraph') throw new Error('Expected paragraph');
    const advances = paragraph.lines.flatMap(line => line.placements)
      .filter(placement => placement.kind === 'text')
      .map(placement => [placement.text, placement.advancePt]);
    await paintLayoutPage(layout, 0, canvas, { scale: 1, dpr: 1 });
    return { advances, fills: fills.map(({ text, x, y, font, fontKerning, letterSpacing }) =>
      ({ text, x, y, font, fontKerning, letterSpacing })) };
  }

  function call(text: string, x: number, letterSpacing: number,
    fontKerning: CanvasFontKerning = 'none', style = '') {
    const family = ['"NotInMetrics"', '"Noto Sans"', '"Noto Sans Hebrew"',
      '"Noto Sans Thai"', '"Noto Sans Devanagari"', '"Arial"', '"Helvetica"',
      '"Liberation Sans"', '"Noto Sans JP"', '"Hiragino Sans"', '"Meiryo"',
      '"Noto Sans SC"', '"Noto Sans TC"', '"Noto Sans KR"', '"Noto Naskh Arabic"',
      '"Noto Sans Arabic"', 'sans-serif'].join(', ');
    const face = style === 'bold ' ? 'normal 700' : style === 'italic ' ? 'italic 400' : 'normal 400';
    return { text, x, y: 16, font: `${face} 20px ${family}`,
      fontKerning, letterSpacing: `${letterSpacing}px` };
  }

  // w:kern is a minimum size (§17.3.2.19): exercise equality, below threshold,
  // and absence. Distinct styles keep the two pairs as separate layout runs.
  for (const kerning of [undefined, 21, 20]) {
    it.each([0, 2])(`pins Latin pair calls and advances at kern=${kerning}, tracking=%s`, async charSpacing => {
      const { fills, advances } = await acquire([
        textRun('AV', { kerning, charSpacing }),
        textRun('To', { kerning, charSpacing, italic: true }),
        textRun('|', { bold: true }),
      ]);
      const active = kerning === 20;
      const advance = (active ? 36 : 40) + 2 * charSpacing;
      expect(advances).toEqual([['AV', advance], ['To', advance], ['|', 20]]);
      expect(fills).toEqual([
        call('AV', 0, charSpacing, active ? 'normal' : 'none'),
        call('To', advance, charSpacing, active ? 'normal' : 'none', 'italic '),
        call('|', advance * 2, 0, 'none', 'bold '),
      ]);
    });
  }

  it.each([0, 2])('pins Arabic joining calls and advances at tracking=%s pt', async charSpacing => {
    const { fills, advances } = await acquire([
      textRun('سلام', { cs: true, charSpacing }), textRun('|', { bold: true }),
    ]);
    const advance = 68 + 4 * charSpacing;
    expect(advances).toEqual([['سلام', advance], ['|', 20]]);
    // Windows Word kerns Arabic text without w:kern (WORD_ARABIC_PAIR_KERNING).
    expect(fills).toEqual([
      call('سلام', 0, charSpacing, 'normal'), call('|', advance, 0, 'none', 'bold '),
    ]);
  });

  it('pins mixed marks, SHY, pairs and Arabic to the main fillText partition', async () => {
    const { fills } = await acquire([
      textRun('AV', { kerning: 20, charSpacing: 2 }),
      textRun('a\u030Ab', { cs: true }),
      textRun('co', { bold: true }),
      // A parsed w:softHyphen is optional ink on both revisions. Literal
      // U+00AD has a deliberately changed invisible-call policy on this branch.
      { ...textRun('', { bold: true }), __optionalHyphen: true } as DocxTextRun,
      textRun('operate', { bold: true }),
      textRun('To', { kerning: 20, charSpacing: 2, italic: true }),
      textRun('سلام', { cs: true, charSpacing: 2 }),
    ]);
    expect(fills).toEqual([
      call('AV', 0, 2, 'normal'),
      call('a\u030Ab', 40, 0),
      call('cooperate', 80, 0, 'none', 'bold '),
      call('To', 260, 2, 'normal', 'italic '),
      call('سلام', 300, 2, 'normal'),
    ]);
  });
});
