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
// established here (a missing-ink diagnostic is printed instead).
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import {
  materializeDocxDocument, openDocxDocument, skia, skiaFactory,
  type NodeCanvasFactory, type NodeCanvasLike,
} from './node-facade.js';

type Pixels = { width: number; height: number; getContext(kind: '2d'): { getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray } } };
type Model = { section: { textDirection?: string | null }; body: { type: string; runs?: { type: string; text?: string }[] }[] };
type Run = { text: string; x: number; y: number; w: number; h: number; font: string; transform?: string; source?: { story: string; path: readonly number[] } };
type Box = { left: number; right: number; top: number; bottom: number };
type Draw = { text: string; font: string; x: number; y: number };
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
                log.draws.push({ text: String(text), font: this.font, x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f });
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

/** Physical box of an overlay run: x/y are the physical top-left and a
 * vertical page rotates the span 90deg clockwise about it. */
function physicalBox(run: Run): Box {
  if (run.transform === undefined) return { left: run.x, right: run.x + run.w, top: run.y, bottom: run.y + run.h };
  if (run.transform === 'rotate(90deg)') return { left: run.x - run.h, right: run.x, top: run.y, bottom: run.y + run.w };
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
