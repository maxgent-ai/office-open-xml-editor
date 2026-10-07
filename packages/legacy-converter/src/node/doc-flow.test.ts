// Binary Word section text flow (MS-DOC 2.6.4 sprmSTextFlow) through the
// ordinary DOCX Canvas renderer.
//
// Every fixture is checked through the native source -> shared model ->
// production layout -> Canvas paint contract: the materialized section flow and
// paragraph text, the public onTextRun projection (source paragraph, exact text
// and order, page-contained geometry, overlay orientation, line advance) and the
// fillText calls that actually reach the returned page canvas. Raster ink is
// additionally asserted for the ASCII and mixed fixtures. For the Han-only
// fixture these checks establish retained text, physical placement and draw
// delivery only; visible CJK ink depends on the host's font supply and is NOT
// established here (a missing-ink diagnostic is printed instead). The native
// MSOTXFL tests below apply the same contract to every normative base flow,
// including glyph orientation read from the actual Canvas matrix.
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16, little32 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import {
  materializeDocxDocument, openDocxDocument, skia, skiaFactory,
  type NodeCanvasFactory, type NodeCanvasLike,
} from './node-facade.js';

type Pixels = { width: number; height: number; getContext(kind: '2d'): { getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray } } };
type Model = { section: { textDirection?: string | null }; body: { type: string; runs?: { type: string; text?: string }[] }[] };
type Run = { text: string; x: number; y: number; w: number; h: number; font: string; transform?: string; source?: { story: string; path: readonly number[] } };
type Box = { left: number; right: number; top: number; bottom: number };
type Direction = 'right' | 'left' | 'down' | 'up';
/** A fillText anchor in device pixels, with the physical directions of the
 * glyph advance (local +x) and glyph top (local -y) under the canvas matrix. */
type Draw = { text: string; font: string; x: number; y: number; advance: Direction; top: Direction };
type DrawLog = { draws: Draw[]; overflow: number };
type Recorded2D = {
  font: string;
  fillText(text: string, x: number, y: number, maxWidth?: number): void;
  getTransform(): { a: number; b: number; c: number; d: number; e: number; f: number };
};

const HAN = /\p{Script=Han}/u;
const han = (value: string) => [...value].filter(ch => HAN.test(ch));
/** Numerical tolerance for composed affine maps, not a layout distance. */
const EPS = 1e-6;
const MAX_DRAWS = 4096;
/** Dominant physical direction of a device-space vector. */
const directionOf = (x: number, y: number): Direction =>
  Math.abs(x) > Math.abs(y) ? (x > 0 ? 'right' : 'left') : (y > 0 ? 'down' : 'up');

/** skia-canvas factory whose canvases log each 2D fillText: its text, the
 * context font and the anchor mapped by the public getTransform() matrix to
 * device pixels. The native method runs with its own receiver and arguments,
 * so production paint is unchanged. */
function recordingSkiaFactory(): { factory: NodeCanvasFactory; logFor(canvas: object): DrawLog | undefined } {
  const base = skiaFactory();
  const logs = new WeakMap<object, DrawLog>();
  const patched = new WeakSet<object>();
  return {
    factory: {
      loadImage: base.loadImage,
      createCanvas(width, height) {
        const canvas = base.createCanvas(width, height);
        const log: DrawLog = { draws: [], overflow: 0 };
        logs.set(canvas, log);
        const getContext = canvas.getContext;
        canvas.getContext = function (this: unknown, ...args: unknown[]) {
          const context = Reflect.apply(getContext, this, args) as Recorded2D | null;
          if (args[0] === '2d' && context && !patched.has(context)) {
            patched.add(context);
            const fillText = context.fillText;
            context.fillText = function (this: Recorded2D, ...call: Parameters<Recorded2D['fillText']>) {
              const [text, x, y] = call;
              const m = this.getTransform();
              if (log.draws.length < MAX_DRAWS) {
                log.draws.push({
                  text: String(text), font: this.font, x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f,
                  advance: directionOf(m.a, m.b), top: directionOf(-m.c, -m.d),
                });
              } else log.overflow++;
              return Reflect.apply(fillText, this, call);
            };
          }
          return context;
        } as unknown as NodeCanvasLike['getContext'];
        return canvas;
      },
    },
    logFor: canvas => logs.get(canvas),
  };
}

/** Physical box of an overlay run: x/y are the physical top-left of the span,
 * which a vertical page rotates a quarter turn about that point (clockwise
 * for downward lines, counter-clockwise for upward lines). */
function physicalBox(run: Run): Box {
  if (run.transform === undefined) return { left: run.x, right: run.x + run.w, top: run.y, bottom: run.y + run.h };
  if (run.transform === 'rotate(90deg)') return { left: run.x - run.h, right: run.x, top: run.y, bottom: run.y + run.w };
  if (run.transform === 'rotate(-90deg)') return { left: run.x, right: run.x + run.h, top: run.y - run.w, bottom: run.y };
  throw new Error(`unexpected overlay transform ${run.transform}`);
}

const contains = (box: Box, point: { x: number; y: number }) => point.x >= box.left - EPS && point.x <= box.right + EPS
  && point.y >= box.top - EPS && point.y <= box.bottom + EPS;

/** Read-only, bounded font catalog summary for the missing-ink diagnostic. */
function fontCatalog() {
  const families = (skia as unknown as { FontLibrary?: { families?: readonly string[] } }).FontLibrary?.families;
  if (!Array.isArray(families)) return { available: false };
  return {
    count: families.length,
    cjkLike: families.filter(family => /\b(cjk|jp|sc|tc|kr)\b|source han|hiragino|mincho|gothic|meiryo|pingfang|songti|wenquanyi|droid sans fallback/i.test(family)).slice(0, 12),
  };
}

it.skipIf(!skia).each([
  '天地玄黄宇宙洪荒\r日月盈昃辰宿列張\r',
  'ABCDEFGH\rIJKLMNOP\r',
  '天地ABC123。\r玄黄DEF456、\r',
])('renders horizontal and vertical section writing direction: %j', async (text) => {
  const paragraphs = text.split('\r').slice(0, -1);
  const expectedHan = han(text).sort();
  const hanOnly = paragraphs.every(paragraph => han(paragraph).length === [...paragraph].length);
  const paragraphOfHan = new Map<string, number>();
  paragraphs.forEach((paragraph, index) => {
    for (const ch of han(paragraph)) {
      expect(paragraphOfHan.get(ch) ?? index).toBe(index);
      paragraphOfHan.set(ch, index);
    }
  });
  for (const flow of [0, 1]) {
    const bytes = buildDocFixture({ text, sectionProperties: concat(little16(0x5033), little16(flow)) });
    const model = await materializeDocxDocument(bytes, { modelSources: [testDocSource()] }) as unknown as Model;
    expect(model.section.textDirection ?? null).toBe(flow === 1 ? 'tbRl' : null);
    expect(model.body.map(block => block.type === 'paragraph'
      ? (block.runs ?? []).filter(run => run.type === 'text').map(run => run.text ?? '').join('')
      : block.type)).toEqual(paragraphs);

    const recording = recordingSkiaFactory();
    const session = await openDocxDocument(bytes, { factory: recording.factory, currentDate: 0, modelSources: [testDocSource()] });
    try {
      expect(session.pageCount).toBe(1);
      const runs: Run[] = [];
      const canvas = await session.renderPage(0, { dpr: 1, onTextRun: run => runs.push(run) }) as unknown as Pixels;
      expect([canvas.width, canvas.height]).toEqual([816, 1056]);

      // Retained text: body paragraphs in source order, each spelling its
      // exact text, with finite page-contained geometry and the page's overlay
      // orientation (none for horizontal, a quarter turn for tbRl).
      expect(runs.length).toBeGreaterThan(0);
      for (const run of runs) {
        expect(run.source?.story).toBe('body');
        expect(run.transform).toBe(flow === 1 ? 'rotate(90deg)' : undefined);
        for (const value of [run.x, run.y, run.w, run.h]) expect(Number.isFinite(value)).toBe(true);
        expect(run.w).toBeGreaterThanOrEqual(0);
        expect(run.h).toBeGreaterThan(0);
        const box = physicalBox(run);
        expect(box.left).toBeGreaterThanOrEqual(0);
        expect(box.top).toBeGreaterThanOrEqual(0);
        expect(box.right).toBeLessThanOrEqual(canvas.width);
        expect(box.bottom).toBeLessThanOrEqual(canvas.height);
      }
      const order = runs.map(run => run.source!.path[0]!);
      expect(order).toEqual([...order].sort((a, b) => a - b));
      const byParagraph = paragraphs.map((_, index) => runs.filter(run => run.source!.path[0] === index));
      for (const [index, group] of byParagraph.entries()) {
        for (const run of group) expect(run.source!.path).toEqual([index]);
      }
      expect(byParagraph.flat()).toHaveLength(runs.length);
      expect(byParagraph.map(group => group.map(run => run.text).join(''))).toEqual(paragraphs);
      // One line per paragraph: segments share the line box and progress along
      // the inline axis (non-strictly, so a zero font advance is not a failure).
      for (const group of byParagraph) {
        for (const [index, run] of group.entries()) {
          if (index === 0) continue;
          const previous = group[index - 1]!;
          if (flow === 1) {
            expect(run.x).toBeCloseTo(previous.x, 6);
            expect(run.y).toBeGreaterThanOrEqual(previous.y - EPS);
          } else {
            expect(run.y).toBeCloseTo(previous.y, 6);
            expect(run.x).toBeGreaterThanOrEqual(previous.x - EPS);
          }
        }
      }
      // The second paragraph's line follows the first: below it horizontally,
      // in the next column to the left for tbRl.
      const [first, second] = byParagraph.map(group => physicalBox(group[0]!));
      if (flow === 1) {
        expect(second!.right).toBeLessThanOrEqual(first!.left + EPS);
        expect(second!.top).toBeCloseTo(first!.top, 6);
      } else {
        expect(second!.top).toBeGreaterThanOrEqual(first!.bottom - EPS);
        expect(second!.left).toBeCloseTo(first!.left, 6);
      }

      // Draw delivery: fillText calls on the returned page canvas land within
      // the page, and every fixture Han character is drawn exactly once,
      // inside an onTextRun box of its own paragraph.
      const log = recording.logFor(canvas);
      expect(log).toBeDefined();
      expect(log!.overflow).toBe(0);
      expect(log!.draws.length).toBeGreaterThan(0);
      for (const draw of log!.draws) {
        expect(Number.isFinite(draw.x) && Number.isFinite(draw.y)).toBe(true);
        expect(contains({ left: 0, right: canvas.width, top: 0, bottom: canvas.height }, draw)).toBe(true);
      }
      const hanDraws = log!.draws.flatMap(draw => han(draw.text).map(ch => ({ ch, draw })));
      expect(hanDraws.map(entry => entry.ch).sort()).toEqual(expectedHan);
      for (const { ch, draw } of hanDraws) {
        const owner = byParagraph[paragraphOfHan.get(ch)!]!;
        expect(owner.some(run => run.text.includes(ch) && contains(physicalBox(run), draw)), ch).toBe(true);
      }
      if (expectedHan.length) {
        const anchors = paragraphs.map((_, index) => hanDraws.filter(entry => paragraphOfHan.get(entry.ch) === index).map(entry => entry.draw));
        if (flow === 1) expect(Math.min(...anchors[0]!.map(draw => draw.x))).toBeGreaterThan(Math.max(...anchors[1]!.map(draw => draw.x)));
        else expect(Math.max(...anchors[0]!.map(draw => draw.y))).toBeLessThan(Math.min(...anchors[1]!.map(draw => draw.y)));
      }

      const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
      let left = canvas.width, right = -1, top = canvas.height, bottom = -1;
      for (let y = 0; y < canvas.height; y++) {
        for (let x = 0; x < canvas.width; x++) {
          const i = (y * canvas.width + x) * 4;
          if (pixels[i + 3] > 200 && pixels[i] < 100 && pixels[i + 1] < 100 && pixels[i + 2] < 100) {
            left = Math.min(left, x); right = Math.max(right, x);
            top = Math.min(top, y); bottom = Math.max(bottom, y);
          }
        }
      }
      if (hanOnly) {
        // Visible Han ink requires a host font covering these characters; the
        // contract above does not establish it. Report missing ink, bounded,
        // so a zero-ink run can be diagnosed separately from the contract.
        if (right < 0) {
          console.warn('[doc-flow] no dark ink for the Han-only fixture; CJK ink is host/font dependent', JSON.stringify({
            flow,
            runs: runs.slice(0, 8).map(run => ({ text: run.text, font: run.font, x: run.x, y: run.y, w: run.w, h: run.h, transform: run.transform })),
            draws: log!.draws.filter(draw => han(draw.text).length).slice(0, 8),
            fonts: fontCatalog(),
          }));
        }
      } else {
        expect(right).toBeGreaterThan(left);
        expect(bottom).toBeGreaterThan(top);
        if (flow === 1) expect(bottom - top).toBeGreaterThan(right - left);
        else expect(right - left).toBeGreaterThan(bottom - top);
      }
    } finally { await session.close(); }
  }
});

// MS-DOC 2.6.4 sprmSTextFlow carries an MS-ODRAW 2.4.5 MSOTXFL. Its normative
// base flows: 0 and 4 advance rightward with later lines below and glyph tops
// up; 1, 3 and 5 advance downward with later lines to the left and tops right;
// 2 advances upward with later lines to the right and tops left. Value 1 keeps
// its established tbRl display (upright East Asian glyphs) and is covered
// above. Expectations below are physical directions taken from that
// description, not from renderer helpers. Word 2007/2010 place later VertN (5)
// lines to the right; the file does not record the displaying application, so
// the normative base flow is expected.
const textFlow = (value: number) => concat(little16(0x5033), little16(value));
const NATIVE_FLOWS: Record<number, Readonly<{ inline: Direction; block: Direction; top: Direction; overlay?: string }>> = {
  0: { inline: 'right', block: 'down', top: 'up' },
  2: { inline: 'up', block: 'right', top: 'left', overlay: 'rotate(-90deg)' },
  3: { inline: 'down', block: 'left', top: 'right', overlay: 'rotate(90deg)' },
  4: { inline: 'right', block: 'down', top: 'up' },
  5: { inline: 'down', block: 'left', top: 'right', overlay: 'rotate(90deg)' },
};
// Distinct physical margins (twips): top 1080, right 1800, bottom 2520, left 720.
const DISTINCT_MARGINS = concat(...([[0x9023, 1080], [0xb022, 1800], [0x9024, 2520], [0xb021, 720]] as const)
  .map(([sprm, value]) => concat(little16(sprm), little16(value))));
/** The box edge where content along `direction` starts, and where it ends. */
const startEdge = (box: Box, direction: Direction) => ({ right: box.left, left: box.right, down: box.top, up: box.bottom })[direction];
const endEdge = (box: Box, direction: Direction) => ({ right: box.right, left: box.left, down: box.bottom, up: box.top })[direction];
const forward = (direction: Direction) => (direction === 'right' || direction === 'down' ? 1 : -1);

it.skipIf(!skia).each([0, 2, 3, 4, 5])('renders native MSOTXFL %i in its normative physical flow', async (flow) => {
  const expected = NATIVE_FLOWS[flow]!;
  const text = '天地ABC123。\r玄黄DEF456、\r';
  const paragraphs = text.split('\r').slice(0, -1);
  const paragraphOfHan = new Map(paragraphs.flatMap((paragraph, index) => han(paragraph).map(ch => [ch, index] as const)));
  const bytes = buildDocFixture({ text, sectionProperties: concat(DISTINCT_MARGINS, textFlow(flow)) });

  const model = await materializeDocxDocument(bytes, { modelSources: [testDocSource()] }) as unknown as Model & {
    section: { pageWidth: number; pageHeight: number; marginTop: number; marginRight: number; marginBottom: number; marginLeft: number };
  };
  // Physical page facts stay physical; 3 and 5 share the established
  // all-rotated clockwise display family, 0 and 4 stay horizontal.
  const section = model.section;
  expect([section.marginTop, section.marginRight, section.marginBottom, section.marginLeft]).toEqual([54, 90, 126, 36]);
  if (flow !== 2) expect(section.textDirection ?? null).toBe(flow === 3 || flow === 5 ? 'btLr' : null);
  expect(model.body.map(block => (block.runs ?? []).filter(run => run.type === 'text').map(run => run.text ?? '').join(''))).toEqual(paragraphs);

  const recording = recordingSkiaFactory();
  const session = await openDocxDocument(bytes, { factory: recording.factory, currentDate: 0, modelSources: [testDocSource()] });
  try {
    expect(session.pageCount).toBe(1);
    const runs: Run[] = [];
    const canvas = await session.renderPage(0, { dpr: 1, onTextRun: run => runs.push(run) }) as unknown as Pixels;
    const scale = canvas.width / section.pageWidth;
    const content = {
      right: section.marginLeft * scale,
      left: canvas.width - section.marginRight * scale,
      down: section.marginTop * scale,
      up: canvas.height - section.marginBottom * scale,
    };

    // Retained text and ownership: body paragraphs in source order, exact text,
    // page-contained boxes, and the public overlay rotation of the flow.
    expect(runs.length).toBeGreaterThan(0);
    for (const run of runs) {
      expect(run.source?.story).toBe('body');
      expect(run.transform).toBe(expected.overlay);
      const box = physicalBox(run);
      expect(box.left >= -EPS && box.top >= -EPS && box.right <= canvas.width + EPS && box.bottom <= canvas.height + EPS).toBe(true);
    }
    const byParagraph = paragraphs.map((_, index) => runs.filter(run => run.source!.path[0] === index));
    expect(byParagraph.flat()).toHaveLength(runs.length);
    expect(byParagraph.map(group => group.map(run => run.text).join(''))).toEqual(paragraphs);

    // The first line starts at the physical margins where its inline and block
    // directions begin; segments advance along the inline direction on one line
    // (non-strictly: a zero font advance is not a failure) and the next
    // paragraph's line lies beyond the first in the block direction.
    const first = physicalBox(byParagraph[0]![0]!);
    expect(startEdge(first, expected.inline)).toBeCloseTo(content[expected.inline], 6);
    expect(startEdge(first, expected.block)).toBeCloseTo(content[expected.block], 6);
    for (const group of byParagraph) {
      for (const [index, run] of group.entries()) {
        if (index === 0) continue;
        const previous = physicalBox(group[index - 1]!), current = physicalBox(run);
        expect((startEdge(current, expected.inline) - startEdge(previous, expected.inline)) * forward(expected.inline)).toBeGreaterThanOrEqual(-EPS);
        expect(startEdge(current, expected.block)).toBeCloseTo(startEdge(previous, expected.block), 6);
      }
    }
    const second = physicalBox(byParagraph[1]![0]!);
    expect((startEdge(second, expected.block) - endEdge(first, expected.block)) * forward(expected.block)).toBeGreaterThanOrEqual(-EPS);
    expect(startEdge(second, expected.inline)).toBeCloseTo(startEdge(first, expected.inline), 6);

    // Actual Canvas draws on the returned page: every glyph advances along the
    // inline direction with its top toward the flow's top direction, and every
    // fixture Han character is drawn once inside a box of its own paragraph.
    const log = recording.logFor(canvas)!;
    expect(log.overflow).toBe(0);
    expect(log.draws.length).toBeGreaterThan(0);
    for (const draw of log.draws) {
      expect(contains({ left: 0, right: canvas.width, top: 0, bottom: canvas.height }, draw)).toBe(true);
      expect([draw.text, draw.advance, draw.top]).toEqual([draw.text, expected.inline, expected.top]);
    }
    const hanDraws = log.draws.flatMap(draw => han(draw.text).map(ch => ({ ch, draw })));
    expect(hanDraws.map(entry => entry.ch).sort()).toEqual(han(text).sort());
    for (const { ch, draw } of hanDraws) {
      expect(byParagraph[paragraphOfHan.get(ch)!]!.some(run => run.text.includes(ch) && contains(physicalBox(run), draw)), ch).toBe(true);
    }

    // ASCII ink is portable: its extent follows the inline axis.
    const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    let left = canvas.width, right = -1, top = canvas.height, bottom = -1;
    for (let y = 0; y < canvas.height; y++) {
      for (let x = 0; x < canvas.width; x++) {
        const i = (y * canvas.width + x) * 4;
        if (pixels[i + 3] > 200 && pixels[i] < 100 && pixels[i + 1] < 100 && pixels[i + 2] < 100) {
          left = Math.min(left, x); right = Math.max(right, x);
          top = Math.min(top, y); bottom = Math.max(bottom, y);
        }
      }
    }
    expect(right).toBeGreaterThan(left);
    expect(bottom).toBeGreaterThan(top);
    if (expected.inline === 'up' || expected.inline === 'down') expect(bottom - top).toBeGreaterThan(right - left);
    else expect(right - left).toBeGreaterThan(bottom - top);
  } finally { await session.close(); }
});

it.skipIf(!skia)('keeps a native MSOTXFL 2 section header upright in the physical page frame', async () => {
  // Page stories keep the existing physical header/footer placement: the
  // section's counter-clockwise frame applies to its body, not its header.
  const bytes = buildDocFixture({
    text: 'Body\r', defaultTabTwips: 720, sectionProperties: textFlow(2),
    headers: ['', 'Head\r', '', '', '', ''],
  });
  const session = await openDocxDocument(bytes, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
  try {
    expect(session.pageCount).toBe(1);
    const runs: Run[] = [];
    await session.renderPage(0, { dpr: 1, onTextRun: run => runs.push(run) });
    const header = runs.filter(run => run.source?.story === 'header');
    const body = runs.filter(run => run.source?.story === 'body');
    expect(header.map(run => run.text).join('')).toBe('Head');
    expect(body.map(run => run.text).join('')).toBe('Body');
    for (const run of header) expect(run.transform).toBeUndefined();
    for (const run of body) expect(run.transform).toBe('rotate(-90deg)');
    // The header lies above the body's physical extent.
    expect(Math.max(...header.map(run => physicalBox(run).bottom)))
      .toBeLessThanOrEqual(Math.min(...body.map(run => physicalBox(run).top)) + EPS);
  } finally { await session.close(); }
});

it.skipIf(!skia).each<{ flows: (number | null)[]; overlays: (string | undefined)[] }>([
  { flows: [2, 1], overlays: ['rotate(-90deg)', 'rotate(90deg)'] },
  { flows: [3, null], overlays: ['rotate(90deg)', undefined] },
])('keeps each native section flow on its own section pages: $flows', async ({ flows, overlays }) => {
  // An ending section's flow must neither be lost nor leak into the final one.
  const bytes = buildDocFixture({
    text: 'One\fTwo\r', sectionEnds: [4, 8],
    sectionProperties: flows.map(flow => (flow === null ? new Uint8Array() : textFlow(flow))),
  });
  const session = await openDocxDocument(bytes, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
  try {
    expect(session.pageCount).toBe(2);
    for (const [page, overlay] of overlays.entries()) {
      const runs: Run[] = [];
      await session.renderPage(page, { dpr: 1, onTextRun: run => runs.push(run) });
      expect(runs.map(run => run.text).join('')).toBe(['One', 'Two'][page]);
      for (const run of runs) expect(run.transform, `page ${page}`).toBe(overlay);
    }
  } finally { await session.close(); }
});

// Native raw 2 and raw 3 sections share the public btLr token but not their
// physical frame. Expected physical corners follow the normative flows and the
// distinct margins above (twips/20 = points on a Letter page): raw 3 lines
// start at the top margin and the first line lies at the right margin; raw 2
// lines start at the bottom margin and the first line lies at the left margin.
// The upright-table test is the runtime-reproduced regression for a later
// section reusing the first section's physical frame (both orders failed
// before the per-location frame fix).
const PAGE_PT = { width: 612, height: 792 };
const MARGINS_PT = { top: 54, right: 90, bottom: 126, left: 36 };
const flowStartCorner = (flow: number, extentPt: number): Box => (flow === 2
  ? {
      left: MARGINS_PT.left, right: MARGINS_PT.left + extentPt,
      top: PAGE_PT.height - MARGINS_PT.bottom - extentPt, bottom: PAGE_PT.height - MARGINS_PT.bottom,
    }
  : {
      left: PAGE_PT.width - MARGINS_PT.right - extentPt, right: PAGE_PT.width - MARGINS_PT.right,
      top: MARGINS_PT.top, bottom: MARGINS_PT.top + extentPt,
    });

it.skipIf(!skia).each([{ flows: [2, 3] }, { flows: [3, 2] }])(
  'places each native section upright table at its own flow start corner: $flows',
  async ({ flows }) => {
    // One two-cell row per section (sprmTInsert 2 x 1000 twips), as in the
    // table shading tests; an empty paragraph carries each section break.
    const cell = new Uint8Array([0, 0, 0x16, 0x24, 1]);
    const row = concat(cell, new Uint8Array([0x17, 0x24, 1, 0x21, 0x76, 0, 2, 0xe8, 3]));
    const tableWidthPt = 100;
    const bytes = buildDocFixture({
      text: 'A\x07B\x07\x07\fC\x07D\x07\x07\r', sectionEnds: [6, 12],
      sectionProperties: flows.map(flow => concat(DISTINCT_MARGINS, textFlow(flow))),
      paragraphMarks: [
        { end: 2, properties: cell }, { end: 4, properties: cell }, { end: 5, properties: row },
        { end: 6, properties: new Uint8Array(2) },
        { end: 8, properties: cell }, { end: 10, properties: cell }, { end: 11, properties: row },
      ],
    });
    const recording = recordingSkiaFactory();
    const session = await openDocxDocument(bytes, { factory: recording.factory, currentDate: 0, modelSources: [testDocSource()] });
    try {
      expect(session.pageCount).toBe(2);
      for (const [page, flow] of flows.entries()) {
        const runs: Run[] = [];
        const canvas = await session.renderPage(page, { dpr: 1, onTextRun: run => runs.push(run) }) as unknown as Pixels;
        const scale = canvas.width / PAGE_PT.width;
        expect(runs.map(run => run.text).join(''), `page ${page}`).toBe(page === 0 ? 'AB' : 'CD');
        // The table is upright in the physical page: no overlay rotation and
        // cells left to right, inside the corner where this section starts.
        const corner = flowStartCorner(flow, tableWidthPt);
        for (const run of runs) {
          expect(run.source?.story).toBe('body');
          expect(run.source!.path.length).toBeGreaterThan(1);
          expect(run.transform).toBeUndefined();
          const box = physicalBox(run);
          expect(box.left, `MSOTXFL ${flow}`).toBeGreaterThanOrEqual(corner.left * scale - 1);
          expect(box.right, `MSOTXFL ${flow}`).toBeLessThanOrEqual(corner.right * scale + 1);
          expect(box.top, `MSOTXFL ${flow}`).toBeGreaterThanOrEqual(corner.top * scale - 1);
          expect(box.bottom, `MSOTXFL ${flow}`).toBeLessThanOrEqual(corner.bottom * scale + 1);
        }
        expect(physicalBox(runs[0]!).left).toBeLessThan(physicalBox(runs[1]!).left);
        // Actual Canvas draws of the cell text are upright.
        const draws = recording.logFor(canvas)!.draws.filter(draw => /[A-D]/.test(draw.text));
        expect(draws.length).toBeGreaterThan(0);
        for (const draw of draws) expect([draw.text, draw.advance, draw.top]).toEqual([draw.text, 'right', 'up']);
      }
    } finally { await session.close(); }
  },
);

// A page-relative square-wrap picture on a native raw 2 page (after a raw 3
// section, so the page owns a different counter-clockwise frame). It protects
// the physical paint position of the picture, text conservation and overlay
// rotation, and that the wrap exclusion projected into the native frame keeps
// the section's text out of the picture. It did not reproduce the stale-frame
// defect at runtime (the mixed upright-table test owns that). Not protected
// here: where the first line lands after the picture (a production line-search
// policy, not a normative rule) and wrap-distance edge relabeling (this
// fixture's padding is equal on every edge, so a relabeled edge is invisible).
// The clockwise page of the same mixed document is not repeated: its anchor
// geometry is the established btLr/tbRl path covered by the docx vertical
// anchor tests.
it.skipIf(!skia)('wraps native MSOTXFL 2 text around its own page-relative picture', async () => {
  const flows = [3, 2];
  // In the left margin corner where the raw 2 section's first line starts.
  const picture = { left: 36, top: 594, right: 108, bottom: 630 };
  const record = (kind: number, options: number, payload: Uint8Array): Uint8Array =>
    concat(little16(options), little16(kind), little32(payload.length), payload);
  const { Canvas } = skia as NonNullable<typeof skia>;
  const raster = new Canvas(20, 10);
  const rasterContext = raster.getContext('2d');
  rasterContext.fillStyle = '#ff0000';
  rasterContext.fillRect(0, 0, 20, 10);
  const blip = record(0xf01e, 0x6e0 << 4, concat(new Uint8Array(17), await raster.toBuffer('png')));
  const shape = record(0xf004, 15, concat(
    record(0xf00a, (75 << 4) | 2, concat(little32(1027), little32(0xa00))),
    record(0xf00b, (1 << 4) | 3, concat(little16(0x4104), little32(1))),
    record(0xf010, 0, little32(0)),
  ));
  const drawingGroupData = concat(
    record(0xf000, 15, record(0xf001, (1 << 4) | 15, blip)),
    new Uint8Array([0]),
    record(0xf002, 15, record(0xf003, 15, shape)),
  );
  // MS-DOC Spa: page-relative position (fXaPage/fYaPage via bx/by) and
  // square wrap (wr 2), anchored by the special picture character of the
  // second section.
  const anchorCp = 2;
  const text = `A\f\u0008${'x'.repeat(240)}\r`;
  const floatingAnchors = concat(
    little32(anchorCp), little32(anchorCp + 1), little32(1027),
    little32(picture.left * 20), little32(picture.top * 20), little32(picture.right * 20), little32(picture.bottom * 20),
    little16((1 << 1) | (1 << 3) | (2 << 5)), little32(0),
  );
  const bytes = buildDocFixture({
    text, sectionEnds: [2, text.length], floatingAnchors, drawingGroupData,
    sectionProperties: flows.map(flow => concat(DISTINCT_MARGINS, textFlow(flow))),
    formattingRuns: [
      { end: anchorCp, properties: new Uint8Array() },
      { end: anchorCp + 1, properties: concat(little16(0x0855), new Uint8Array([1])) },
      { end: text.length, properties: new Uint8Array() },
    ],
  });
  const session = await openDocxDocument(bytes, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
  try {
    expect(session.pageCount).toBe(2);
    const runs: Run[] = [];
    const canvas = await session.renderPage(1, { dpr: 1, onTextRun: run => runs.push(run) }) as unknown as Pixels;
    const scale = canvas.width / PAGE_PT.width;
    const flow = flows[1]!;
    expect(runs.map(run => run.text).join('')).toBe('x'.repeat(240));
    for (const run of runs) expect(run.transform).toBe(NATIVE_FLOWS[flow]!.overlay);
    // The drawing layer is physical: the picture paints at its page position
    // whatever the section's text flow.
    const image = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
    let left = canvas.width, right = -1, top = canvas.height, bottom = -1;
    for (let y = 0; y < canvas.height; y++) for (let x = 0; x < canvas.width; x++) {
      const i = (y * canvas.width + x) * 4;
      if (image.data[i] > 240 && image.data[i + 1] < 15 && image.data[i + 2] < 15) {
        left = Math.min(left, x); right = Math.max(right, x); top = Math.min(top, y); bottom = Math.max(bottom, y);
      }
    }
    const expected = [picture.left * scale, picture.top * scale, picture.right * scale - 1, picture.bottom * scale - 1];
    [left, top, right, bottom].forEach((value, index) => expect(Math.abs(value - expected[index]!)).toBeLessThanOrEqual(1));
    // The section's text wraps around the picture instead of running through
    // it.
    const pictureBox: Box = {
      left: picture.left * scale, right: picture.right * scale, top: picture.top * scale, bottom: picture.bottom * scale,
    };
    for (const run of runs) {
      const box = physicalBox(run);
      const overlaps = box.left < pictureBox.right - EPS && box.right > pictureBox.left + EPS
        && box.top < pictureBox.bottom - EPS && box.bottom > pictureBox.top + EPS;
      expect(overlaps, JSON.stringify({ text: run.text, box })).toBe(false);
    }
    // The no-overlap check above is not vacuous: the section's text shares
    // the picture's inline (physical vertical) band, so text that ignored the
    // native-frame exclusion would run through the picture.
    expect(runs.some((run) => {
      const box = physicalBox(run);
      return box.top < pictureBox.bottom - EPS && box.bottom > pictureBox.top + EPS;
    })).toBe(true);
  } finally { await session.close(); }
});

// An inline picture stays upright in the physical page whatever the section
// flow: its local counter-turn must invert the owner section's own quarter
// turn (clockwise for 1 and 3, counter-clockwise for native 2). The picture
// is asymmetric (red left half, blue right half), so an upside-down or
// quarter-turned paint is visible in its actual canvas pixels.
/** MS-DOC PICFAndOfficeArtData with an OfficeArtBlipPNG (pib 1) of a 40 x 20
 * picture, red left half and blue right half, displayed at one inch by half
 * an inch (as in the doc-images and doc-notes tests). */
async function asymmetricInlinePicture(): Promise<Uint8Array> {
  const record = (kind: number, options: number, payload: Uint8Array): Uint8Array =>
    concat(little16(options), little16(kind), little32(payload.length), payload);
  const { Canvas } = skia as NonNullable<typeof skia>;
  const source = new Canvas(40, 20);
  const sourceContext = source.getContext('2d');
  sourceContext.fillStyle = '#ff0000'; sourceContext.fillRect(0, 0, 20, 20);
  sourceContext.fillStyle = '#0000ff'; sourceContext.fillRect(20, 0, 20, 20);
  const shape = record(0xf004, 15, concat(
    record(0xf00a, (75 << 4) | 2, concat(little32(1), little32(0x800))),
    record(0xf00b, 0x13, concat(little16(0x0104), little32(1))),
  ));
  const blip = record(0xf01e, 0x6e0 << 4, concat(new Uint8Array(17), await source.toBuffer('png')));
  const header = new Uint8Array(68);
  const view = new DataView(header.buffer);
  view.setUint32(0, header.length + shape.length + blip.length, true);
  for (const [offset, value] of [[4, 68], [6, 100], [28, 1440], [30, 720], [32, 1000], [34, 1000]]) view.setUint16(offset, value, true);
  return concat(header, shape, blip);
}

/** sprmCFSpec plus sprmCPicLocation 0: the picture character's properties. */
const PICTURE_CHARACTER = concat(little16(0x0855), new Uint8Array([1]), little16(0x6a03), little32(0));

/** Physical centroid offset of the painted blue half from the red half. */
function redToBlueOffset(canvas: Pixels): { dx: number; dy: number } {
  const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
  const red = { count: 0, x: 0, y: 0 }, blue = { count: 0, x: 0, y: 0 };
  for (let y = 0; y < canvas.height; y++) for (let x = 0; x < canvas.width; x++) {
    const i = (y * canvas.width + x) * 4;
    const target = pixels[i] > 240 && pixels[i + 1] < 15 && pixels[i + 2] < 15 ? red
      : pixels[i] < 15 && pixels[i + 1] < 15 && pixels[i + 2] > 240 ? blue : null;
    if (target) { target.count++; target.x += x; target.y += y; }
  }
  expect(red.count).toBeGreaterThan(100);
  expect(blue.count).toBeGreaterThan(100);
  return { dx: blue.x / blue.count - red.x / red.count, dy: blue.y / blue.count - red.y / red.count };
}

it.skipIf(!skia).each([0, 1, 3, 2])('keeps an inline picture upright in MSOTXFL %i', async (flow) => {
  const bytes = buildDocFixture({
    text: '\u0001', data: await asymmetricInlinePicture(),
    sectionProperties: concat(DISTINCT_MARGINS, textFlow(flow)),
    characterProperties: PICTURE_CHARACTER,
  });
  const session = await openDocxDocument(bytes, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
  try {
    expect(session.pageCount).toBe(1);
    const canvas = await session.renderPage(0, { dpr: 1 }) as unknown as Pixels;
    // Upright: the red half lies to the physical left of the blue half, along
    // the physical horizontal axis.
    const { dx, dy } = redToBlueOffset(canvas);
    expect(dx, `MSOTXFL ${flow}`).toBeGreaterThan(0);
    expect(Math.abs(dx), `MSOTXFL ${flow}`).toBeGreaterThan(Math.abs(dy));
  } finally { await session.close(); }
});

// The same asymmetric picture inline in a body table cell. On a vertical page
// the body table is laid out upright in the physical page (no rotation of its
// own), so the picture inside its cell is already in an upright frame and must
// not receive the section's counter-turn. The physical owner of the picture is
// the upright table, not the section's rotated text frame.
it.skipIf(!skia).each([0, 3, 2])('keeps an inline picture upright inside a body table cell in MSOTXFL %i', async (flow) => {
  // One cell holding the picture character, its row mark (sprmTInsert one
  // 2000-twip cell, as in the table shading tests) and a closing paragraph.
  const cell = new Uint8Array([0, 0, 0x16, 0x24, 1]);
  const row = concat(cell, new Uint8Array([0x17, 0x24, 1, 0x21, 0x76, 0, 1, 0xd0, 0x07]));
  const text = '\u0001\x07\x07\r';
  const bytes = buildDocFixture({
    text, data: await asymmetricInlinePicture(),
    sectionProperties: concat(DISTINCT_MARGINS, textFlow(flow)),
    paragraphMarks: [{ end: 2, properties: cell }, { end: 3, properties: row }],
    formattingRuns: [
      { end: 1, properties: PICTURE_CHARACTER },
      { end: text.length, properties: new Uint8Array() },
    ],
  });
  const session = await openDocxDocument(bytes, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
  try {
    expect(session.pageCount).toBe(1);
    const canvas = await session.renderPage(0, { dpr: 1 }) as unknown as Pixels;
    const { dx, dy } = redToBlueOffset(canvas);
    expect(dx, `MSOTXFL ${flow}`).toBeGreaterThan(0);
    expect(Math.abs(dx), `MSOTXFL ${flow}`).toBeGreaterThan(Math.abs(dy));
  } finally { await session.close(); }
});
