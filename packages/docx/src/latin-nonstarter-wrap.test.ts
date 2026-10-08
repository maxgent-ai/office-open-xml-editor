import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// UAX#14 LB13 / ECMA-376 §17.15.1.59 (行頭禁則 / line-start-forbidden): a comma
// (and other closing / mid punctuation — class IS/CL/CP/EX) has NO line-break
// opportunity before it, so it can never BEGIN a line. Word keeps such a mark
// with the word it follows even when the two live in different runs.
//
// Real-world trigger (sample-12 "Keywords"): the runs are
//   …"cyber security, "  |  "intrusion detection system"  |  ", metadata"…
// so the word "system" (end of one run, NO trailing space) is followed by a
// SEPARATE run that begins with a comma. Splitting on spaces yields the layout
// segments  … "detection " · "system" · ", " · "metadata" …  with no whitespace
// between "system" and ",". A naive wrap treats every segment boundary as a
// break opportunity and orphans the comma at the next line's start; Word wraps
// "system," together.

interface Call { text: string; x: number; y: number; px: number; }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) { calls.push({ text: s, x, y, px: px() }); },
    strokeText(s: string, x: number, y: number) { calls.push({ text: s, x, y, px: px() }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function run(text: string): DocParagraph['runs'][number] {
  return {
    type: 'text', text, bold: false, italic: false, underline: false,
    strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
    fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    smallCaps: false, allCaps: false,
  } as DocParagraph['runs'][number];
}

function multiRunPara(texts: string[]): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: texts.map(run),
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[], pageWidth: number): DocxDocumentModel {
  const section = {
    pageWidth, pageHeight: 600,
    marginTop: 5, marginRight: 5, marginBottom: 5, marginLeft: 5,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(body: BodyElement[], pageWidth: number): Promise<Call[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(body, pageWidth), canvas, 0, { dpr: 1, width: pageWidth });
  return calls;
}

describe('Latin line-start-forbidden wrap (UAX#14 LB13 / §17.15.1.59)', () => {
  it('does not orphan a comma at the start of a line when it lives in a separate run', async () => {
    // contentW = 90 - 5 - 5 = 80px (mock: 5px / char @ 10pt). Tuned so "system"
    // alone would fit at the end of line 1 but "system," (comma glued) does not —
    // the exact band where a naive wrap orphans the comma onto line 2.
    const calls = await render(
      [multiRunPara(['aaaa bbbb system', ', cc']) as unknown as BodyElement],
      90,
    );
    const system = calls.find((c) => c.text.includes('system'));
    const comma = calls.find((c) => c.text.includes(','));
    expect(system, 'painted "system"').toBeDefined();
    expect(comma, 'painted comma').toBeDefined();
    // The comma must ride on the SAME line as "system" (never lead a line).
    expect(comma!.y, '"system" and "," share a line').toBeCloseTo(system!.y, 3);
    // …and sit immediately AFTER it (to its right), not before.
    expect(comma!.text.indexOf(',')).toBeGreaterThan(comma!.text.indexOf('system'));
  });

  it('still wraps normally at real (whitespace) break opportunities', async () => {
    // A plain space before the comma DOES permit a break; nothing is glued.
    const calls = await render(
      [multiRunPara(['aaaa bbbb cccc dddd eeee ffff gggg']) as unknown as BodyElement],
      90,
    );
    // Sanity: the paragraph wrapped onto more than one line.
    const ys = new Set(calls.map((c) => Math.round(c.y)));
    expect(ys.size).toBeGreaterThan(1);
  });
});
