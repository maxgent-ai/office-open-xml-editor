import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ECMA-376 §17.6.5 docGrid CHARACTER grid (字詰め) — 約物半角 bracket overlap.
//
// Under an active character grid a PURE East-Asian segment is drawn so its
// glyphs occupy exactly `measuredWidth` (= natural + len·Δ). The bug: the draw
// loop painted EACH code point with a SEPARATE `fillText(cp, …)` (ISOLATED
// shaping) while positioning glyph i at `measureText(prefix_i) + i·Δ`
// (CONTEXTUAL shaping). An opening bracket "［" (U+FF3B) is collapsed to
// half-width by the browser's 約物半角 contextual shaping ONLY inside a
// multi-char string (`measureText("［分") < measureText("［")+measureText("分")`).
// Because the bracket was painted ISOLATED at FULL width but every later glyph
// was positioned by the COLLAPSED cumulative `measureText(prefix)`, the
// following glyphs were pulled ~half-em left and OVERLAPPED the bracket.
//
// The fix draws each CONTIGUOUS piece as ONE `fillText` with `ctx.letterSpacing`
// = Δ, so measure and draw both shape the SAME (contextual) way ⇒ no overlap and
// measure==draw preserved.

const FONT_PX = 20; // glyph advance per full-width CJK char in the stub (scale 1)

// East-Asian test: CJK Unified, kana, full-/half-width forms, CJK punctuation.
const EA_RE = /[　-〿぀-ヿ㐀-鿿＀-￯]/;

/** Simulate the browser's 約物半角 contextual half-width collapse of an opening
 *  bracket "［" when it is FOLLOWED by an East-Asian glyph WITHIN the measured
 *  string. This is the exact rule that makes the real bug reproduce: an isolated
 *  "［" measures FULL, but "［分" measures less than "［" + "分". The collapse
 *  must NOT include letterSpacing (the helper adds Δ via `from*Δ` and the canvas
 *  adds Δ between glyphs WITHIN a drawn piece). */
function ctxMeasure(s: string, fontPx: number): number {
  const cps = [...s];
  let w = 0;
  for (let i = 0; i < cps.length; i++) {
    const c = cps[i];
    const ea = EA_RE.test(cps[i + 1] ?? '');
    w += c === '［' && ea ? fontPx / 2 : fontPx;
  }
  return w;
}

/** Recording 2D context. measureText models 約物半角 contextual collapse; fillText
 *  records text + x + y AND the current letterSpacing at call time (so we can
 *  assert the contiguous-draw fix sets `ctx.letterSpacing = Δ`). The stub HONOURS
 *  letterSpacing in measureText only when a non-glued context would otherwise be
 *  measured — but per the helper contract, measure is always called BEFORE the
 *  renderer sets letterSpacing, so this never double-counts. */
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; x: number; y: number; letterSpacing: string }[];
} {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fillTextCalls: { text: string; x: number; y: number; letterSpacing: string }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      const w = ctxMeasure(s, p);
      return {
        width: w,
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
      fillTextCalls.push({ text, x, y, letterSpacing });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

function textRun(text: string, fontSize = FONT_PX): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(
  runs: DocxTextRun[],
  opts: { alignment?: DocParagraph['alignment']; autoSpace?: boolean } = {},
): BodyElement {
  const p: DocParagraph = {
    alignment: opts.alignment ?? 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r }) as DocRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics',
    widowControl: false,
    ...(opts.autoSpace === false ? { autoSpaceDE: false, autoSpaceDN: false } : {}),
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 600, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

/** linesAndChars grid with the given raw charSpace (active CHARACTER grid). */
function charGrid(charSpace: number): Partial<SectionProps> {
  return { docGridType: 'linesAndChars', docGridLinePitch: 20, docGridCharSpace: charSpace };
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
  sec: SectionProps,
): Promise<{
  runs: DocxTextRunInfo[];
  fillTextCalls: { text: string; x: number; y: number; letterSpacing: string }[];
}> {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, {
    dpr: 1,
    width: sec.pageWidth, // scale = 1 (px per pt)
    onTextRun: (r) => runs.push(r),
  });
  return { runs, fillTextCalls };
}

const EA_TEXT = '［分類］である'; // 7 EA code points, opening bracket first → one segment

describe('docGrid character grid — 約物半角 brackets are drawn contiguously, never overlap (§17.6.5)', () => {
  const CHARSPACE = -1024;
  const DELTA = CHARSPACE / 4096; // Δpt × scale(=1) = -0.25px per EA glyph

  // (1) RED→GREEN: the EA grid segment is drawn as ONE contiguous fillText whose
  //     text === the whole segment, at the segment's left x, with letterSpacing=Δ.
  //     Before the fix: 7 single-glyph fillText calls with letterSpacing '0px'.
  it('draws the whole EA grid segment as a single contiguous fillText with letterSpacing=Δ', async () => {
    const { runs, fillTextCalls } = await render(
      [para([textRun(EA_TEXT)])],
      section(charGrid(CHARSPACE)),
    );

    const seg = runs.find((r) => r.text === EA_TEXT);
    expect(seg, 'EA segment reported via onTextRun').toBeDefined();

    // The segment is painted by EXACTLY ONE fillText carrying the whole string —
    // not seven isolated per-code-point draws (the overlap source).
    const segCalls = fillTextCalls.filter((c) => c.text === EA_TEXT);
    expect(segCalls.length, 'one contiguous fillText for the EA segment').toBe(1);
    // No per-single-codepoint isolated draw exists for this segment's glyphs.
    const isolated = fillTextCalls.filter(
      (c) => [...c.text].length === 1 && [...EA_TEXT].includes(c.text),
    );
    expect(isolated.length, 'no per-code-point isolated EA draws').toBe(0);

    // Drawn at the segment's left edge…
    expect(segCalls[0].x).toBeCloseTo(seg!.x, 6);
    // …and the per-EA-glyph grid delta is applied via ctx.letterSpacing.
    expect(segCalls[0].letterSpacing).toBe(`${DELTA}px`);
  });

  // (2) No-overlap invariant (stronger): the following segment starts exactly at
  //     the EA box edge = x + measure(whole) + len·Δ. Because the draw is
  //     contiguous (browser shapes the whole piece) and the box uses the same
  //     contextual measure, measure==draw box is preserved and the next run never
  //     overlaps the bracket's tail.
  it('keeps measure==draw so the following Latin run abuts (no overlap)', async () => {
    const { runs, fillTextCalls } = await render(
      // The bracket box edge alone: no automatic East Asian/Latin gap.
      [para([textRun(EA_TEXT), textRun('A')], { autoSpace: false })],
      section(charGrid(CHARSPACE)),
    );

    const segEA = runs.find((r) => r.text === EA_TEXT)!;
    const segLat = runs.find((r) => r.text === 'A')!;

    // The EA segment is a single draw; the Latin run is a single draw.
    expect(fillTextCalls.filter((c) => c.text === EA_TEXT).length).toBe(1);

    // The box edge: measure(whole) is contextual (bracket collapsed), plus len·Δ.
    const len = [...EA_TEXT].length;
    const boxW = ctxMeasure(EA_TEXT, FONT_PX) + len * DELTA;
    expect(segEA.w).toBeCloseTo(boxW, 6);
    // The Latin run starts exactly at the EA box edge — no overlap, no gap.
    expect(segLat.x).toBeCloseTo(segEA.x + boxW, 6);
    expect(fillTextCalls.find((c) => c.text === 'A')!.x).toBeCloseTo(segEA.x + boxW, 6);
  });

  // (3a) Grid delta still applied: an active grid on a pure-EA run records
  //      letterSpacing == Δ (字詰め spacing not lost).
  it('records the grid delta as letterSpacing on an active char grid', async () => {
    const { fillTextCalls } = await render(
      [para([textRun(EA_TEXT)])],
      section(charGrid(CHARSPACE)),
    );
    const segCall = fillTextCalls.find((c) => c.text === EA_TEXT)!;
    expect(segCall.letterSpacing).toBe(`${DELTA}px`);
  });

  // (3b) Grid inactive (charSpace absent): the EA run is drawn by case 3 (single
  //      fillText) and letterSpacing is left at its default '0px' (restored).
  it('leaves letterSpacing at 0px when the char grid is inactive', async () => {
    const { fillTextCalls } = await render(
      [para([textRun(EA_TEXT)])],
      section({ docGridType: 'linesAndChars', docGridLinePitch: 20 }), // no charSpace
    );
    const segCall = fillTextCalls.find((c) => c.text === EA_TEXT)!;
    expect(segCall.letterSpacing).toBe('0px');
    // Single fillText (no per-glyph walk).
    expect(fillTextCalls.filter((c) => c.text === EA_TEXT).length).toBe(1);
  });

  // (4) Justify + grid: a `distribute` line retains one contextual EA operation
  //     whose uniform pitch combines the active grid delta and line slack.
  it('keeps the contextual EA segment under distribute with its uniform pitch', async () => {
    // Narrow page so the line is stretched and gets internal distribute gaps.
    const { fillTextCalls } = await render(
      [para([textRun('あいうえお')], { alignment: 'distribute' })],
      section({ ...charGrid(CHARSPACE), pageWidth: 200 }),
    );

    const contextual = fillTextCalls.filter((operation) => operation.text === 'あいうえお');
    expect(contextual).toHaveLength(1);
    const uniformPitch = parseFloat(contextual[0].letterSpacing);
    expect(uniformPitch).toBeGreaterThan(DELTA);
    const finalGlyphOrigin = contextual[0].x + 4 * (FONT_PX + uniformPitch);
    expect(finalGlyphOrigin).toBeGreaterThan(4 * FONT_PX);
  });
});
