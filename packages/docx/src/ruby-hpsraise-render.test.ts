import { describe, expect, it } from 'vitest';
import { buildSegments, rubyAscentReservePx } from './line-layout.js';
import { layoutDocument } from './document-layout.js';
import { referenceFontLineMetrics } from './reference-font-line-metrics.js';
import { createLayoutServices } from './layout-runtime.js';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types';

interface GlyphCall {
  text: string;
  x: number;
  y: number;
  font: string;
  kerning: string;
}

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  glyphs: GlyphCall[];
} {
  let font = '10px serif';
  let letterSpacing = '0px';
  let fontKerning = 'auto';
  const glyphs: GlyphCall[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(value: string) { letterSpacing = value; },
    get fontKerning() { return fontKerning; },
    set fontKerning(value: string) { fontKerning = value; },
    measureText(text: string) {
      const size = px();
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    fillText(text: string, x: number, y: number) { glyphs.push({ text, x, y, font, kerning: fontKerning }); },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, setLineDash() {}, drawImage() {}, clearRect() {},
    arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection, globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, glyphs };
}

function textRun(hpsRaisePt?: number, guide = 'かん'): DocxTextRun {
  return {
    text: '漢', bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 12, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null,
    ruby: {
      text: guide,
      fontSizePt: 8,
      ...(hpsRaisePt != null ? { hpsRaisePt } : {}),
    },
  };
}

function documentWithRuby(
  hpsRaisePt?: number,
  lineSpacing: DocParagraph['lineSpacing'] = null,
  fontFamily = 'NotInMetrics',
  guide?: string,
): DocxDocumentModel {
  const paragraph: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing, numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(hpsRaisePt, guide), fontFamily }],
    defaultFontSize: 12, defaultFontFamily: fontFamily, widowControl: false,
  };
  const section: SectionProps = {
    pageWidth: 200, pageHeight: 200,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  };
  return {
    section,
    body: [{ type: 'paragraph', ...paragraph } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function bodyRubyBaselineDistance(hpsRaisePt?: number, scale = 2): Promise<number> {
  const { canvas, glyphs } = makeRecordingCanvas();
  await renderDocumentToCanvas(documentWithRuby(hpsRaisePt), canvas, 0, {
    dpr: 1,
    width: 200 * scale,
  });
  const base = glyphs.find((glyph) => glyph.text === '漢');
  const ruby = glyphs.find((glyph) => /[かん]/.test(glyph.text));
  expect(base, 'base glyph drawn').toBeDefined();
  expect(ruby, 'ruby glyph drawn').toBeDefined();
  return base!.y - ruby!.y;
}

describe('ruby line ascent (WORD_RUBY_LINE_ASCENT, #1885)', () => {
  // Windows Word 16.0: the guide's em box stands hpsRaise above the base
  // baseline; the line ascent is the larger of the base ascent and the guide
  // top, over the base descent. The stub face has a 0.8/0.2 em box.
  it.each([
    { raise: 9, ascent: 9 + 8 },
    { raise: 2, ascent: 12 * 0.8 + 0.4 },
    { raise: 0, ascent: 12 * 0.8 },
  ])('hpsRaise $raise pt with an 8pt guide over 12pt base text', ({ raise, ascent }) => {
    const model = documentWithRuby(raise);
    const { canvas } = makeRecordingCanvas();
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: canvas.getContext('2d') }));
    const retained = layout.pages[0].layers.body[0];
    if (retained.kind !== 'paragraph') throw new Error('expected ruby test paragraph');
    const line = retained.lines[0];
    const expectedAscent = Math.max(12 * 0.8, ascent);
    expect(line.advancePt).toBeCloseTo(expectedAscent + 12 * 0.2, 6);
    expect(line.baselinePt - line.bounds.yPt).toBeCloseTo(expectedAscent, 6);
  });

  it('leaves the unmeasured absent-hpsRaise class on its previous baseline', () => {
    const model = documentWithRuby(undefined, { value: 1.5, rule: 'auto', explicit: true });
    const { canvas } = makeRecordingCanvas();
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: canvas.getContext('2d') }));
    const retained = layout.pages[0].layers.body[0];
    if (retained.kind !== 'paragraph') throw new Error('expected ruby test paragraph');
    const line = retained.lines[0];
    // The values main produced before the hpsRaise class was measured.
    expect(line.advancePt).toBeCloseTo(34.8, 6);
    expect(line.baselinePt - line.bounds.yPt).toBeCloseTo(26.6, 6);
  });

  it('keeps measured sides for absent-hpsRaise ruby over a catalogued face', () => {
    const model = documentWithRuby(undefined, null, 'MS Mincho');
    const { canvas } = makeRecordingCanvas();
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: canvas.getContext('2d') }));
    const retained = layout.pages[0].layers.body[0];
    if (retained.kind !== 'paragraph') throw new Error('expected ruby test paragraph');
    const line = retained.lines[0];
    // The values main produced before the hpsRaise class was measured.
    expect(line.advancePt).toBeCloseTo(23.2, 6);
    expect(line.baselinePt - line.bounds.yPt).toBeCloseTo(20.8, 6);
  });

  it('paints a pair-kerned Arabic guide as one kerned contextual span', async () => {
    const { canvas, glyphs } = makeRecordingCanvas();
    await renderDocumentToCanvas(documentWithRuby(9, null, 'NotInMetrics', 'سلام'), canvas, 0, { dpr: 1, width: 400 });
    const guide = glyphs.filter((glyph) => /\p{Script=Arabic}/u.test(glyph.text));
    expect(guide.map((glyph) => [glyph.text, glyph.kerning])).toEqual([['سلام', 'normal']]);
    expect(glyphs.find((glyph) => glyph.text === '漢')?.kerning).not.toBe('normal');
  });
});

describe('§17.3.3.12 w:hpsRaise ruby geometry', () => {
  it('keeps the non-ruby grid cell count when the annotation fits below the design minimum', () => {
    const advances = [false, true].map((hasRuby) => {
      const model = documentWithRuby(2);
      model.section.docGridType = 'lines';
      model.section.docGridLinePitch = 18;
      const paragraph = model.body[0] as Extract<BodyElement, { type: 'paragraph' }>;
      paragraph.defaultFontFamily = 'Yu Mincho';
      const run = paragraph.runs[0] as DocxTextRun;
      run.fontFamily = 'Yu Mincho';
      run.ruby = hasRuby ? { text: 'かん', fontSizePt: 6, hpsRaisePt: 2 } : undefined;
      const { canvas } = makeRecordingCanvas();
      // Yu Mincho is present (its resource metric is the catalog usWin
      // projection), so its design minimum governs the grid cells.
      const yuMincho = referenceFontLineMetrics('Yu Mincho', 400, 'normal')!;
      const layout = layoutDocument(model, createLayoutServices(model, {
        measureContext: canvas.getContext('2d'),
        localMetrics: { 'yu mincho': {
          ...yuMincho, requestedFamily: 'Yu Mincho', family: 'Yu Mincho', weight: 400, style: 'normal',
          sourceIdentity: 'local:yu-mincho',
        } },
      }));
      const retained = layout.pages[0].layers.body[0];
      if (retained.kind !== 'paragraph') throw new Error('expected ruby test paragraph');
      return retained.lines[0].advancePt;
    });
    // 12pt Yu Mincho: 1.3 × usWin = 20.08pt, requiring two 18pt cells.
    // The measured base + guide is smaller; adding ruby cannot remove a cell.
    expect(advances).toEqual([36, 36]);
  });

  it('draws horizontal body ruby hpsRaise above the base baseline', async () => {
    expect(await bodyRubyBaselineDistance(14)).toBeCloseTo(14, 8);
  });

  it('touches retained base and guide ink when hpsRaise is absent', async () => {
    // The injected authoritative ink is 9.6pt above the base baseline and
    // 1.6pt below the 8pt guide baseline. Point-space geometry is scale-free.
    expect(await bodyRubyBaselineDistance(undefined, 2)).toBeCloseTo(11.2, 8);
  });

  it('treats explicit hpsRaise zero as a zero reservation and zero draw offset', async () => {
    expect(rubyAscentReservePx(8, 0, 2)).toBe(0);
    expect(await bodyRubyBaselineDistance(0)).toBe(0);
  });

  it('keeps absent hpsRaise off the emitted segment own-property set', () => {
    const environment = {
      defaultFontSize: 12,
      defaultFontFamily: 'NotInMetrics',
      fontFamilyClasses: {},
      pageIndex: 0,
      totalPages: 1,
    };
    const [absent] = buildSegments(
      [{ type: 'text', ...textRun() }],
      environment,
    );
    const [zero] = buildSegments(
      [{ type: 'text', ...textRun(0) }],
      environment,
    );
    expect(absent && 'ruby' in absent && absent.ruby).toBeDefined();
    expect(zero && 'ruby' in zero && zero.ruby).toBeDefined();
    if (!absent || !('ruby' in absent) || !zero || !('ruby' in zero)) {
      throw new Error('expected text segments');
    }
    expect(Object.hasOwn(absent.ruby!, 'hpsRaisePt')).toBe(false);
    expect(Object.hasOwn(zero.ruby!, 'hpsRaisePt')).toBe(true);
    expect(zero.ruby!.hpsRaisePt).toBe(0);
  });
});
