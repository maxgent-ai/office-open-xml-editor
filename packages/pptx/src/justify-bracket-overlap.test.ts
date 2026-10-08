import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { TextBody, Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// ECMA-376 §20.1.10.59 (ST_TextAlignType `just`/`dist`) — opening-bracket overlap
// on the JUSTIFY draw path. This is the pptx leg of docx PR #630 (and siblings
// #626/#627/#629).
//
// A justified (`algn="just"`) or distributed (`algn="dist"`) pure-CJK line goes
// through the justify branch of the draw loop. justifyLine opens a gap at EVERY
// inter-CJK boundary on a pure-CJK line, so `splitBefore` lists a cut before each
// glyph (length === cps.length - 1 ⇒ FULLY distributed). The OLD code then drew
// ONE single-glyph piece per code point via `drawWithFont` — and because each
// piece has `text.length === 1`, drawWithFont's `ls > 0 && text.length > 1` guard
// is FALSE, so it always fell to the `else → paint(ch)` branch, drawing each glyph
// ISOLATED (even when ls > 0).
//
// The contextual collapse that bites here is JIS X 4051 約物連続 (consecutive-
// punctuation packing), NOT a bracket-next-to-kana collapse: a CLOSING-class glyph
// immediately followed by an OPENING bracket — "：［", "、［", "）（" — packs the
// pair ~half-em tighter in `measureText` (verified on real fonts: "［本" does NOT
// pack, "名：［" does). So the cumulative measure stepping INTO the bracket (after
// "：") is half-width, but the OLD draw painted the bracket ISOLATED at FULL width,
// so it overran its successor by ~half-em — the next glyph was pulled under the
// bracket (the "分 ⊂ ［" smashing of docx sample-16, real-font-verified in #630).
//
// The fix: a FULLY-distributed run (a gap at every internal boundary ⇒ uniform
// per-glyph pitch ls+segPerGap) is drawn as the whole CONTEXTUALLY-shaped run in
// ONE `fillText` with `ctx.letterSpacing = ls + segPerGap`. measure and draw then
// shape the SAME (contextual) way ⇒ the packing is honoured identically and
// nothing overlaps; glyph i lands at measure(prefix_i)+i·(ls+segPerGap) — the
// exact justified position — and the final glyph reaches the segment box edge.
//
// NOTE: the load-bearing guard below is STRUCTURAL and font-independent — the
// fully-distributed line is drawn as exactly ONE contiguous `fillText` with
// `ctx.letterSpacing` set (vs the OLD N isolated single-glyph draws). The
// `ctxMeasure` mock models the 約物連続 packing only to reconstruct the OLD path's
// overlap illustratively (it is an ILLUSTRATIVE model of the real-font mechanism).

const FONT_PX = 20; // full-width EA glyph advance in the stub (scale 1)
const SCALE = 1 / 12700; // emuToPx(emu, SCALE) = emu·SCALE; PT_TO_EMU=12700 ⇒ 1pt → 1px

// JIS X 4051 約物連続 (consecutive-punctuation packing): a CLOSING-class glyph
// immediately FOLLOWED by an OPENING bracket has its adjacent empty half-bodies
// merged, so the pair measures ~half-em tighter. (A bare opening bracket next to
// a kana/kanji does NOT collapse — verified on real fonts; only the punctuation
// PAIR packs.) These are the classes the browser's measureText compresses.
const CLOSE_PUNCT = '：。）、］」』';
const OPEN_BRACKET = '［（「『';

/** Model the browser's contextual width: full-em per glyph, MINUS a half-em at
 *  each adjacency where a closing-class punctuation is immediately followed by an
 *  opening bracket (約物連続 packing, e.g. "：［", "、［", "）（"). So
 *  `measure("名：［") < measure("名：") + measure("［")` while `measure("［本")
 *  === measure("［") + measure("本")` (a bracket next to a kanji does NOT pack).
 *  The collapse itself does not include letterSpacing; the recording context
 *  adds the active Canvas tracking separately. */
function ctxMeasure(s: string, fontPx: number): number {
  const cps = [...s];
  let w = 0;
  for (let i = 0; i < cps.length; i++) {
    w += fontPx;
    if (i > 0 && OPEN_BRACKET.includes(cps[i]) && CLOSE_PUNCT.includes(cps[i - 1])) {
      w -= fontPx / 2; // closing-punct + opening-bracket pair packs half-em tighter
    }
  }
  return w;
}

/** Recording 2D context. `measureText` models 約物連続 contextual packing plus
 * native Canvas letterSpacing. `fillText` records text + x + y AND the current
 * letterSpacing at call time, so a test can assert the contiguous-draw fix set
 * `ctx.letterSpacing = ls + segPerGap`. */
function mockCtx(): {
  ctx: CanvasRenderingContext2D;
  texts: { text: string; x: number; y: number; letterSpacing: string }[];
  getLetterSpacing: () => string;
} {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const px = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const texts: { text: string; x: number; y: number; letterSpacing: string }[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    get letterSpacing() { return letterSpacing; }, set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: ctxMeasure(s, p) + [...s].length * (parseFloat(letterSpacing) || 0),
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText: (t: string, x: number, y: number) => texts.push({ text: t, x, y, letterSpacing }),
    strokeText: () => {},
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter' as CanvasLineJoin,
  };
  return {
    ctx: ctx as unknown as CanvasRenderingContext2D,
    texts,
    getLetterSpacing: () => letterSpacing,
  };
}

/** `letterSpacing` arrives in points and is scaled to px by `ls·PT_TO_EMU·SCALE`;
 *  with SCALE = 1/12700 that is 1:1, so `letterSpacing: ls` yields `ls` px. A
 *  distinct `color` forces a SEPARATE layout segment (the layout merges adjacent
 *  same-style runs into one segment via `sameMeta`). */
function run(text: string, letterSpacing?: number, color = '000000'): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color, fontFamily: 'Serif',
    // One face for every slot keeps each probe a single-font run; an empty
    // East Asian slot would draw the CJK glyphs in PowerPoint's application
    // default instead (issue #1627).
    fontFamilyEa: 'Serif',
    ...(letterSpacing != null ? { letterSpacing } : {}),
  } as TextRunData;
}

function bodyWith(alignment: Paragraph['alignment'], runs: TextRunData[]): TextBody {
  const para: Paragraph = {
    alignment,
    marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], eaLnBrk: true, runs,
  } as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: 20,
    defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  };
}

type RunInfo = { text: string; inShapeX: number; w: number };

function render(
  body: TextBody,
  boxW = 600,
): {
  texts: { text: string; x: number; y: number; letterSpacing: string }[];
  runs: RunInfo[];
  finalLetterSpacing: string;
} {
  const { ctx, texts, getLetterSpacing } = mockCtx();
  const runs: RunInfo[] = [];
  renderTextBody(
    ctx, body, 0, 0, boxW, 400, SCALE,
    null, 0, false, false, '#000000', undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    (r) => runs.push({ text: r.text, inShapeX: r.inShapeX, w: r.w }),
  );
  return { texts, runs, finalLetterSpacing: getLetterSpacing() };
}

// A pure-CJK run carrying a real 約物連続 pair: "：［" (closing punctuation
// immediately followed by an opening bracket). NO run @spc (ls === 0). Wide enough
// (the box below holds them all on one `dist` line) that the line FULLY distributes
// — justifyLine sets splitBefore at every inter-glyph boundary.
const EA_JUST = 'スタイル名：［分類］。あいうえお';
// `dist` fills EVERY line (even a single/last line), so a single-line box is
// fully distributed. Box must be WIDER than the natural advance so there is slack.
const NATURAL = ctxMeasure(EA_JUST, FONT_PX); // contextual natural advance
const BOX_W = NATURAL + 60; // slack 60 → spread across all internal gaps

describe('pptx justify CJK — 約物連続 brackets are drawn contiguously, never overlap (§20.1.10.59)', () => {
  // (1) RED→GREEN core fix: a FULLY-distributed `dist` CJK line is painted by
  //     EXACTLY ONE contiguous fillText carrying the whole run text, with
  //     letterSpacing === `${segPerGap}px` (ls=0 ⇒ ls+segPerGap = segPerGap,
  //     non-zero). Before the fix: N single-code-point fillText calls at '0px'.
  it('draws the fully-distributed justify line as a single contiguous fillText with letterSpacing=segPerGap', () => {
    const { texts } = render(bodyWith('dist', [run(EA_JUST)]), BOX_W);

    const cps = [...EA_JUST];
    // GREEN: one contiguous fillText carrying the whole run.
    const segCalls = texts.filter((c) => c.text === EA_JUST);
    expect(segCalls.length, 'one contiguous fillText for the fully-distributed line').toBe(1);
    const drawn = segCalls[0];
    expect([...drawn.text].length, 'the contiguous draw spans many code points').toBeGreaterThan(1);
    expect(drawn.text).toContain('［');

    // No per-single-code-point isolated draw exists for this run's glyphs (the OLD
    // path emitted one fillText per code point).
    const isolated = texts.filter(
      (c) => [...c.text].length === 1 && cps.includes(c.text),
    );
    expect(isolated.length, 'no per-code-point isolated EA draws (the OLD path)').toBe(0);

    // letterSpacing carries the justify pitch (= ls + segPerGap; ls=0 here), a
    // non-zero positive px — NOT '0px'.
    expect(drawn.letterSpacing).toMatch(/^-?\d+(\.\d+)?px$/);
    expect(drawn.letterSpacing).not.toBe('0px');
    const pitch = parseFloat(drawn.letterSpacing);
    expect(pitch, 'justify pitch is a positive expansion (= segPerGap)').toBeGreaterThan(0);
    // slack 60 across (cps.length - 1) internal gaps.
    expect(pitch).toBeCloseTo(60 / (cps.length - 1), 6);
  });

  // (2) No-overlap (illustrative): reconstruct the OLD isolated per-glyph x
  //     positions (dx = ctxMeasure(prefix_i) + i·segPerGap, each glyph drawn
  //     ISOLATED at FULL FONT_PX) and assert the OLD path overlapped at the "：［"
  //     pair, while the fix's single contiguous draw makes overlap impossible.
  it('keeps the opening bracket from overlapping the next glyph (the 分⊂［ regression is gone)', () => {
    const { texts } = render(bodyWith('dist', [run(EA_JUST)]), BOX_W);

    const segCalls = texts.filter((c) => c.text === EA_JUST);
    expect(segCalls.length, 'one contiguous draw for the line').toBe(1);
    const drawn = segCalls[0];

    const cps = [...EA_JUST];
    const segPerGap = parseFloat(drawn.letterSpacing);
    expect(EA_JUST).toContain('［');

    // OLD-path x positions (the overlap source): isolated glyph i at
    // ctxMeasure(prefix_i) + i·segPerGap. The 約物連続 packing collapses the "：［"
    // pair, so the cumulative measure stepping past the bracket is half-em short,
    // pulling the glyph AFTER the bracket left of where the FULL-width isolated
    // bracket actually ends.
    const oldXs = cps.map((_c, i) =>
      ctxMeasure(cps.slice(0, i).join(''), FONT_PX) + i * segPerGap,
    );
    // Each glyph was painted ISOLATED, so it occupies a FULL FONT_PX advance from
    // its x. Scan adjacent pairs for a glyph whose isolated full-width end exceeds
    // the next glyph's start ⇒ OVERLAP (this is "分⊂［", real-font ~half-em).
    let worstOverlap = 0;
    for (let i = 1; i < cps.length; i++) {
      const prevIsolatedEnd = oldXs[i - 1] + FONT_PX;
      const overlap = prevIsolatedEnd - oldXs[i];
      if (overlap > worstOverlap) worstOverlap = overlap;
    }
    expect(
      worstOverlap,
      'OLD isolated per-code-point path overlapped at the 約物連続 "：［" pair',
    ).toBeGreaterThan(1); // ~half-em (≈ FONT_PX/2) minus one gap pitch

    // The fix: ONE contiguous fillText ⇒ the browser shapes the whole run the SAME
    // way it was measured (約物連続 packing honoured) ⇒ no isolated glyph can be
    // pulled left of its neighbour. The overlap is structurally GONE.
    expect(segCalls.length).toBe(1);
    expect(drawn.text).toBe(EA_JUST);
  });

  // (3) Partial-gap path unaffected: a justified line where splitBefore is a
  //     SUBSET of the inter-glyph boundaries (length < cps.length - 1) keeps the
  //     per-piece drawWithFont loop (multiple fillText). A mixed CJK+Latin run
  //     produces gaps only at inter-CJK boundaries (not inside the Latin word), so
  //     the run is NOT fully distributed and draws as multiple pieces.
  it('keeps the multi-glyph piece loop on a partial-gap (CJK+Latin) justified line', () => {
    // "あ［AB］い": gaps at every inter-CJK boundary; A|B (Latin|Latin) opens none,
    // so the pieces are [あ][［][AB][］][い] (5 pieces, splitBefore subset).
    const MIXED = 'あ［AB］い';
    const natural = ctxMeasure(MIXED, FONT_PX);
    const { texts } = render(bodyWith('dist', [run(MIXED)]), natural + 40);

    // The "AB" piece is one contiguous draw (a multi-glyph piece, drawWithFont).
    const abCalls = texts.filter((c) => c.text === 'AB');
    expect(abCalls.length, 'the AB piece is one contiguous draw').toBe(1);
    // It is NOT the whole-run contiguous fast-path: the whole run is never drawn as
    // one call (it is split into pieces), so multiple fillText calls exist.
    const wholeRun = texts.filter((c) => c.text === MIXED);
    expect(wholeRun.length, 'the fast-path did NOT fire (partial gaps)').toBe(0);
    expect(texts.length, 'multiple pieces drawn via the loop').toBeGreaterThan(1);
    // ls === 0 here, so each piece is drawn at letterSpacing '0px' (unchanged path).
    for (const c of texts) expect(c.letterSpacing).toBe('0px');
  });

  // (4) No-regression: a non-justified line and an @spc-only (ls>0, not fully
  //     distributed) line behave as before.
  it('leaves a non-justified line and an @spc-only line unchanged', () => {
    // (a) Left-aligned control: one contiguous draw at letterSpacing '0px'.
    const left = render(bodyWith('l', [run(EA_JUST)]), BOX_W);
    const leftSeg = left.texts.filter((c) => c.text === EA_JUST);
    expect(leftSeg.length, 'left-aligned line is one contiguous draw').toBe(1);
    expect(leftSeg[0].letterSpacing, 'left-aligned line not stretched').toBe('0px');

    // (b) @spc-only NON-justified line: ls>0 distributed via the existing
    //     drawWithFont letterSpacing path (text.length > 1), ONE contiguous draw at
    //     letterSpacing === `${ls}px` (NOT ls+segPerGap; this is not a justify line).
    const LS = 4;
    // Include the authored tracking in the box budget so this remains a
    // paint-path control rather than accidentally becoming a wrap test.
    const spcBoxW = BOX_W + ([...EA_JUST].length - 1) * LS;
    const spc = render(bodyWith('l', [run(EA_JUST, LS)]), spcBoxW);
    const spcSeg = spc.texts.filter((c) => c.text === EA_JUST);
    expect(spcSeg.length, '@spc-only line is one contiguous draw').toBe(1);
    expect(spcSeg[0].letterSpacing, '@spc-only line carries ls (not a justify pitch)').toBe(`${LS}px`);

    // letterSpacing is restored to its initial value after each render.
    expect(left.finalLetterSpacing).toBe('0px');
    expect(spc.finalLetterSpacing).toBe('0px');
  });
});
