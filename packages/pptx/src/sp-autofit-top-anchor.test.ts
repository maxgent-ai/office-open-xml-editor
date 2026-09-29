import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { renderTextBody } from './renderer.js';
import type { Paragraph, TextBody } from './types.js';

const SCALE = 1 / 12700; // 1 point => 1 canvas unit

function recordingContext(
  actualAscent: number,
  actualDescent: number,
  exposeFontMetrics = true,
  fontAscent = actualAscent,
  fontDescent = actualDescent,
) {
  const draws: Array<{ text: string; y: number }> = [];
  let font = '';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    get fillStyle() { return fillStyle; },
    set fillStyle(value: string) { fillStyle = value; },
    get direction() { return direction; },
    set direction(value: CanvasDirection) { direction = value; },
    measureText: (text: string) => ({
      width: [...text].length * 20,
      actualBoundingBoxAscent: actualAscent,
      actualBoundingBoxDescent: actualDescent,
      ...(exposeFontMetrics
        ? { fontBoundingBoxAscent: fontAscent, fontBoundingBoxDescent: fontDescent }
        : {}),
    }),
    fillText: (text: string, _x: number, y: number) => draws.push({ text, y }),
    fillRect: () => {},
    drawImage: () => {},
    save: () => {},
    restore: () => {},
    translate: () => {},
    rotate: () => {},
    scale: () => {},
    beginPath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    stroke: () => {},
    clip: () => {},
    rect: () => {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, draws };
}

function body(fontFamily: string, fontFamilyEa: string, fontSize: number): TextBody {
  const run: TextRunData = {
    type: 'text',
    text: '見出し',
    bold: true,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize,
    color: '000000',
    fontFamily,
    fontFamilyEa,
  };
  const paragraph: Paragraph = {
    alignment: 'l',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs: [run],
  } as Paragraph;
  return {
    verticalAnchor: 't',
    paragraphs: [paragraph],
    defaultFontSize: fontSize,
    defaultBold: null,
    defaultItalic: null,
    lIns: 0,
    rIns: 0,
    tIns: 0,
    bIns: 0,
    wrap: 'none',
    vert: 'horz',
    autoFit: 'sp',
  } as TextBody;
}

function twoLineBody(fontFamily: string, fontFamilyEa: string, fontSize: number): TextBody {
  const textBody = body(fontFamily, fontFamilyEa, fontSize);
  const first = textBody.paragraphs[0]!.runs[0] as TextRunData;
  textBody.paragraphs[0]!.runs = [
    { ...first, text: '一行目' },
    { type: 'break' } as unknown as TextRunData,
    { ...first, text: '二行目' },
  ];
  return textBody;
}


// PowerPoint's metric split (#1610, powerpoint-line-metrics.ts): usWin ascent
// shares of a 1.2 × size line.
const MEIRYO_SHARE = 2171 / (2171 + 901);
const ARIAL_SHARE = 1854 / (1854 + 434);
describe('pptx spAutoFit top anchoring', () => {
  it.each([
    ['Meiryo', 12], ['Meiryo', 18], ['Meiryo', 24], ['Meiryo', 32],
    ['Arial', 12], ['Arial', 18], ['Arial', 24], ['Arial', 32],
  ] as const)(
    'uses the PowerPoint 120%% implicit pitch for %s at %d pt with and without spAutoFit',
    (fontFamily, fontSize) => {
      for (const autoFit of ['none', 'sp'] as const) {
        const { ctx, draws } = recordingContext(
          fontSize * 0.78, fontSize * 0.18, true, fontSize * 0.98, fontSize * 0.37,
        );
        const textBody = { ...twoLineBody(fontFamily, fontFamily, fontSize), autoFit };
        renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);
        expect(draws[1]!.y - draws[0]!.y).toBeCloseTo(fontSize * 1.2, 5);
      }
    },
  );

  it('keeps implicit multi-line pitch separate from the taller resolved font box (#1473)', () => {
    const fontSize = 32;
    const { ctx, draws } = recordingContext(
      fontSize * 0.78, fontSize * 0.18, true, fontSize * 0.98, fontSize * 0.37,
    );
    const textBody = twoLineBody('Meiryo', 'Meiryo', fontSize);
    renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);

    expect(draws.map(({ text }) => text)).toEqual(['一行目', '二行目']);
    // #1610 controls (spAutoFit and fixed boxes alike): the usWin split of the
    // 1.2 × size line, independent of the Canvas font box and glyph ink.
    expect(draws[0]!.y).toBeCloseTo((fontSize * 1.2 * MEIRYO_SHARE), 5);
    expect(draws[1]!.y).toBeCloseTo((fontSize * 1.2 * (1 + MEIRYO_SHARE)), 5);
    const neededHeight = renderTextBody(
      ctx, textBody, 0, 0, 400, 46, SCALE,
      null, 0, false, false, '#000000', undefined, undefined, undefined, true,
    );
    // #1610 table controls: a row grows to the sum of the 1.2 × size lines.
    expect(neededHeight).toBeCloseTo(2 * fontSize * 1.2, 5);
  });

  it.each([
    ['theme-resolved heading face', 32, 46],
    ['explicit Meiryo face', 36, 50.9],
  ])('seats the %s first line by its metric split, not by glyph ink', (_label, fontSize, storedHeight) => {
    // #1610 H-Meiryo-*-t: the spAutoFit first baseline is 118 of 1/100 in for
    // 100 pt Meiryo, the usWin split, whatever Canvas reports for ink.
    const { ctx, draws } = recordingContext(
      fontSize * 0.78, fontSize * 0.18, true, fontSize * 0.98, fontSize * 0.37,
    );
    renderTextBody(ctx, body('Meiryo', 'Meiryo', fontSize), 0, 0, 400, storedHeight, SCALE);
    expect(draws).toHaveLength(1);
    expect(draws[0]!.y).toBeCloseTo((fontSize * 1.2 * MEIRYO_SHARE), 5);
  });

  it.each(['ctr', 'b'] as const)(
    'anchors %s spAutoFit text in the authored box without regrowing it',
    (verticalAnchor) => {
      // #1610 H-*-h60-ctr: a centred block overflows a too-small authored
      // box on both edges; PowerPoint does not regrow the shape.
      const fontSize = 32;
      const { ctx, draws } = recordingContext(fontSize * 0.78, fontSize * 0.18);
      const textBody = { ...twoLineBody('Meiryo', 'Meiryo', fontSize), verticalAnchor };
      const boxHeight = 30;
      renderTextBody(ctx, textBody, 0, 0, 400, boxHeight, SCALE);
      const block = 2 * fontSize * 1.2;
      const top = verticalAnchor === 'ctr' ? (boxHeight - block) / 2 : boxHeight - block;
      expect(draws[0]!.y).toBeCloseTo(top + (fontSize * 1.2 * MEIRYO_SHARE), 5);
    },
  );

  it('keeps the zero-height bottom-anchored shape growing upward', () => {
    const fontSize = 32;
    const { ctx, draws } = recordingContext(fontSize * 0.78, fontSize * 0.18);
    const textBody = { ...body('Meiryo', 'Meiryo', fontSize), verticalAnchor: 'b' as const };
    renderTextBody(ctx, textBody, 0, 100, 400, 0, SCALE);
    expect(draws[0]!.y).toBeCloseTo(100 - fontSize * 1.2 + (fontSize * 1.2 * MEIRYO_SHARE), 5);
  });

  it('places Arial by its own split in fixed and spAutoFit shapes alike', () => {
    const fontSize = 32;
    for (const autoFit of ['none', 'sp'] as const) {
      const { ctx, draws } = recordingContext(fontSize * 0.98, fontSize * 0.37, false);
      renderTextBody(ctx, { ...body('Arial', 'Arial', fontSize), autoFit }, 0, 0, 400, 46, SCALE);
      expect(draws[0]!.y).toBeCloseTo((fontSize * 1.2 * ARIAL_SHARE), 5);
    }
  });

  it('retains the resolved font-box baseline for percentage spacing in a top-anchored spAutoFit shape', () => {
    const fontSize = 32;
    const actualAscent = fontSize * 0.78;
    const fontAscent = fontSize * 0.98;
    const fontDescent = fontSize * 0.37;
    const { ctx, draws } = recordingContext(
      actualAscent,
      fontSize * 0.18,
      true,
      fontAscent,
      fontDescent,
    );
    const textBody = body('Unavailable CJK face', 'Unavailable CJK face', fontSize);
    textBody.paragraphs[0]!.spaceLine = { type: 'pct', val: 120000 };

    renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);

    expect(draws).toHaveLength(1);
    const authoredLineHeight = fontSize * 1.2;
    const resolvedFontHeight = fontAscent + fontDescent;
    expect(draws[0]!.y).toBeCloseTo(
      fontAscent + (authoredLineHeight - resolvedFontHeight) / 2,
      5,
    );
  });

  it.each(['ctr', 'b'] as const)(
    'does not apply the top-anchor percentage correction to %s anchoring',
    (verticalAnchor) => {
      const fontSize = 32;
      const actualAscent = fontSize * 0.78;
      const fontAscent = fontSize * 0.98;
      const fontDescent = fontSize * 0.37;
      const { ctx, draws } = recordingContext(
        actualAscent,
        fontSize * 0.18,
        true,
        fontAscent,
        fontDescent,
      );
      const textBody = { ...body('Unavailable CJK face', 'Unavailable CJK face', fontSize), verticalAnchor };
      textBody.paragraphs[0]!.spaceLine = { type: 'pct', val: 120000 };
      const lineHeight = fontSize * 1.2 * 1.2;

      renderTextBody(ctx, textBody, 0, 0, 400, lineHeight, SCALE);

      expect(draws).toHaveLength(1);
      const resolvedFontHeight = fontAscent + fontDescent;
      expect(draws[0]!.y).toBeCloseTo(
        fontAscent + Math.max(0, lineHeight - resolvedFontHeight) / 2,
        5,
      );
    },
  );

  it('does not apply the percentage correction to absolute point spacing', () => {
    const fontSize = 32;
    const actualAscent = fontSize * 0.78;
    const fontAscent = fontSize * 0.98;
    const fontDescent = fontSize * 0.37;
    const { ctx, draws } = recordingContext(
      actualAscent,
      fontSize * 0.18,
      true,
      fontAscent,
      fontDescent,
    );
    const textBody = body('Unavailable CJK face', 'Unavailable CJK face', fontSize);
    textBody.paragraphs[0]!.spaceLine = { type: 'pts', val: 36 };

    renderTextBody(ctx, textBody, 0, 0, 400, 60, SCALE);

    expect(draws).toHaveLength(1);
    expect(draws[0]!.y).toBeCloseTo(actualAscent, 5);
  });
});
