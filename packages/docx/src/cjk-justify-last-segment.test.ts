import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ECMA-376 §17.18.44 (ST_Jc `both`/`distribute`): a justified line spreads its
// slack EQUALLY across EVERY inter-CJK boundary on the line — Word adds the same
// inter-character pitch to all glyphs, not just to the earlier segments'.
//
// Regression guard: the docx renderer passed the line's VISUALLY-LAST segment
// index to the shared slack kernel as `lastDrawnSi`, which suppresses ALL of that
// segment's interior gaps (not only the final glyph's). When the last segment was
// a multi-character CJK run, it rendered at the un-stretched grid pitch while the
// earlier segment absorbed every bit of slack — two visibly different inter-
// character pitches on ONE line. This is exactly sample-10's "でご確認ください．"
// (loose) vs "使用言語は日本語または英語で" (tight). The fix excludes no segment
// on an LTR line (the content-span trim already keeps the final glyph on the
// margin), matching the pptx justifier. See packages/core/src/text/
// line-distribute.ts and packages/docx/src/renderer.ts (distributeLineSlack call).

const FONT_PX = 20; // glyph advance per CJK char in the stub (scale = 1)

/** Recording 2D context: glyph advance = charCount × fontPx, font box 0.8/0.2 em. */
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; x: number; y: number; letterSpacing: number }[];
} {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fillTextCalls: { text: string; x: number; y: number; letterSpacing: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) {
      fillTextCalls.push({ text, x, y, letterSpacing: parseFloat(letterSpacing) });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

function textRun(text: string, color: string | null): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
}

type DocRun = DocParagraph['runs'][number];

/** A `both`-justified paragraph carrying two adjacent CJK runs. The differing
 *  colour guarantees two separate layout segments (so the second is the line's
 *  visually-last segment), reproducing sample-10's run16|run17 split. */
function twoRunPara(a: string, b: string): BodyElement {
  const p: DocParagraph = {
    alignment: 'both',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [
      { type: 'text', ...textRun(a, null) } as DocRun,
      { type: 'text', ...textRun(b, 'FF0000') } as DocRun,
    ],
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics',
    widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 210, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    // Active character grid (§17.6.5): every full-width EA glyph is drawn
    // individually, so fillText reports a position PER glyph on every segment —
    // including the last — which lets us read the realised inter-glyph pitch.
    docGridType: 'linesAndChars', docGridLinePitch: 20, docGridCharSpace: -2048,
    ...overrides,
  };
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(body: BodyElement[], sec: SectionProps, compatibilityMode?: number) {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  await renderDocumentToCanvas({ ...doc(body, sec), settings: { compatibilityMode } } as DocxDocumentModel, canvas, 0, { dpr: 1, width: sec.pageWidth });
  return fillTextCalls;
}

describe('CJK justification — even slack distribution across the last segment (§17.18.44)', () => {
  it('preserves unweighted CJK expansion across an empty metric host in mode 15', async () => {
    const paragraph = twoRunPara('観', '察観察観察') as DocParagraph & { type: 'paragraph' };
    paragraph.runs.splice(1, 0, {
      type: 'anchorHost', fontSize: FONT_PX, fontFamily: 'NotInMetrics',
    } as unknown as DocRun);
    const sec = section({ pageWidth: 70, docGridType: 'default', docGridCharSpace: undefined });
    const previous = await render([paragraph], sec, 14);
    expect(await render([paragraph], sec, 15)).toEqual(previous);
  });

  it('gives the visually-last CJK segment the SAME inter-glyph pitch as the first', async () => {
    // Cell = 20 + (-2048/4096) = 19.5 px. availW 210 → 10 cells (195) fit with
    // 15 px slack; the 11th (214.5) overflows. So line 1 = "ああああ" (run A, 4) +
    // 6 of run B; line 2 holds the rest (and is the natural last line). Line 1 is
    // a justify candidate with two segments — the bug stretched only run A.
    const calls = await render(
      [twoRunPara('ああああ', 'いいいいいいいいいいいいいいいい')],
      section(),
    );
    expect(calls.length).toBeGreaterThan(0);

    // Group glyphs by baseline y; the first (top) line is the justified one.
    const byY = new Map<number, { text: string; x: number; letterSpacing: number }[]>();
    for (const c of calls) {
      const key = Math.round(c.y);
      (byY.get(key) ?? byY.set(key, []).get(key)!).push(c);
    }
    const firstY = Math.min(...byY.keys());
    const line = byY.get(firstY)!.slice().sort((p, q) => p.x - q.x);

    // Contextual paint remains split only at the authored run boundary: run A
    // (4) plus the first-line prefix of run B (6). Both retain the same uniform
    // inter-glyph pitch, including the visually-last segment.
    expect(line.map((operation) => operation.text)).toEqual([
      'ああああ', 'いいいいいい',
    ]);
    expect(line[0].letterSpacing).toBeCloseTo(line[1].letterSpacing, 9);
    expect(line[1].letterSpacing).toBeGreaterThan(-0.5);
  });
});

it.each(['ああああああああああああ', 'ああ ab ああああああああ'])('preserves mode-15 expansion for the unmeasured CJK gap family: %s', async text => {
  const body = [twoRunPara(text, 'いいいいいい')];
  const sec = section({ docGridType: undefined, docGridCharSpace: undefined });
  const previous = await render(body, sec, 14);
  const current = await render(body, sec, 15);
  expect(current).toEqual(previous);
});
