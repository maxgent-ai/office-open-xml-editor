import { describe, expect, it } from 'vitest';
import { renderSlide } from './renderer';
import type { ShapeElement, Slide } from './types';

// An open-arrow line end puts its outer miter on the line end and the line
// stops at the arms' vertex (observed PowerPoint rule, core arrow-geometry.ts).
// Leaders that are not re-stroked from straight vertices must be trimmed too,
// or the shaft's flat end shows beside the pointed tip.

const EMU_PER_PX = 9_525; // 960 px across 9 144 000 EMU
const LINE_EMU = 114_300; // 9 pt → 12 px
// lg/lg open arrow: arm 4.5 × 2.5 line widths, so the vertex sits
// 6 px / sin(atan(2.5 / 4.5)) behind the tip.
const VERTEX_PX = 6 / Math.sin(Math.atan2(2.5, 4.5));

type Point = { x: number; y: number };

function shape(overrides: Partial<ShapeElement>): ShapeElement {
  return {
    type: 'shape', x: 100 * EMU_PER_PX, y: 100 * EMU_PER_PX, width: 300 * EMU_PER_PX, height: 200 * EMU_PER_PX,
    rotation: 0, flipH: false, flipV: false, geometry: 'line', fill: null,
    stroke: { color: '000000', width: LINE_EMU, tailEnd: { type: 'arrow', w: 'lg', len: 'lg' } },
    textBody: null, defaultTextColor: null, custGeom: null,
    adj: null, adj2: null, adj3: null, adj4: null, adj5: null, adj6: null, adj7: null, adj8: null,
    ...overrides,
  } as ShapeElement;
}

/** Render one shape and return the slide-space end point of every shaft stroke. */
async function shaftEnds(element: ShapeElement): Promise<Point[]> {
  let current: Point[] = [];
  let translated = false;
  const ends: Point[] = [];
  const state: Record<string, unknown> = {
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
    measureText: (text: string) => ({ width: text.length * 6 }),
    save: () => {}, restore: () => { translated = false; },
    translate: () => { translated = true; },
    beginPath: () => { current = []; },
    moveTo: (x: number, y: number) => current.push({ x, y }),
    lineTo: (x: number, y: number) => current.push({ x, y }),
    quadraticCurveTo: (_a: number, _b: number, x: number, y: number) => current.push({ x, y }),
    bezierCurveTo: (_a: number, _b: number, _c: number, _d: number, x: number, y: number) => current.push({ x, y }),
    ellipse: (cx: number, cy: number, rx: number, ry: number, rot: number, _from: number, to: number) => current.push({
      x: cx + rx * Math.cos(to) * Math.cos(rot) - ry * Math.sin(to) * Math.sin(rot),
      y: cy + rx * Math.cos(to) * Math.sin(rot) + ry * Math.sin(to) * Math.cos(rot),
    }),
    // Line-end heads are painted in a translated tip frame; keep shafts only.
    stroke: () => { if (!translated && current.length) ends.push(current[current.length - 1]); },
  };
  const ctx = new Proxy(state, {
    get: (target, property) => (property in target ? target[property as string] : () => undefined),
    set: (target, property, value) => { target[property as string] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const canvas = { width: 960, height: 720, style: {}, offsetWidth: 960, getContext: () => ctx };
  const slide: Slide = { index: 0, slideNumber: 1, background: null, elements: [element] };
  await renderSlide(canvas as unknown as HTMLCanvasElement, slide, 9_144_000, 6_858_000, { width: 960, dpr: 1 });
  return ends;
}

describe('PPTX line-end shaft trimming', () => {
  const expectStopsAtVertex = (ends: Point[], tip: Point) => {
    expect(ends.length).toBeGreaterThan(0);
    for (const end of ends) expect(Math.hypot(end.x - tip.x, end.y - tip.y)).toBeCloseTo(VERTEX_PX, 1);
  };

  it('trims a curved connector leader along its curve', async () => {
    expectStopsAtVertex(await shaftEnds(shape({ geometry: 'curvedConnector3' })), { x: 400, y: 300 });
  });

  it('trims a custom-geometry cubic path', async () => {
    expectStopsAtVertex(await shaftEnds(shape({
      geometry: 'custGeom',
      custGeom: [[{ cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'cubicBezTo', x1: 0.5, y1: 0, x2: 0.5, y2: 1, x: 1, y: 1 }]],
    })), { x: 400, y: 300 });
  });

  it('trims a custom-geometry arc and keeps it on its ellipse', async () => {
    // Quarter ellipse from the top-left corner to the bottom-right (400, 300).
    const ends = await shaftEnds(shape({
      geometry: 'custGeom',
      custGeom: [[{ cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'arcTo', wr: 1, hr: 1, stAng: -90, swAng: 90 }]],
    }));
    expectStopsAtVertex(ends, { x: 400, y: 300 });
    // Centre (100, 300), radii 300 × 200.
    for (const end of ends) expect(((end.x - 100) / 300) ** 2 + ((end.y - 300) / 200) ** 2).toBeCloseTo(1, 6);
  });

  it('trims a bent connector across a terminal leg shorter than the retraction', async () => {
    // bentConnector2: right along the top edge, then a 1 px leg down to (400, 101).
    expectStopsAtVertex(await shaftEnds(shape({ geometry: 'bentConnector2', height: EMU_PER_PX })), { x: 400, y: 101 });
  });

  it('paints no compound shaft when both ends consume a short connector', async () => {
    const arrow = { type: 'arrow', w: 'lg', len: 'lg' };
    expect(await shaftEnds(shape({
      geometry: 'straightConnector1', width: 15 * EMU_PER_PX, height: 0,
      stroke: { color: '000000', width: LINE_EMU, cmpd: 'dbl', headEnd: arrow, tailEnd: arrow },
    }))).toEqual([]);
  });

  it('trims both lines of a compound straight connector', async () => {
    const ends = await shaftEnds(shape({
      geometry: 'straightConnector1', height: 0,
      stroke: { color: '000000', width: LINE_EMU, cmpd: 'dbl', tailEnd: { type: 'arrow', w: 'lg', len: 'lg' } },
    }));
    expect(ends.length).toBeGreaterThan(1);
    for (const end of ends) expect(end.x).toBeCloseTo(400 - VERTEX_PX, 1);
  });
});
