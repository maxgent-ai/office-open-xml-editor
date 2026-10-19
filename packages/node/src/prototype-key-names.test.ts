import { describe, expect, it } from 'vitest';
import { renderSlideNode } from './render';
import type { NodeCanvasFactory } from './render';
import type { Presentation } from '@silurus/ooxml-pptx';
import { loadSkiaForTests } from './test-imports';

// Font, preset-geometry, text-warp and shadow-alignment names are document
// strings. Names that are also Object.prototype members ("constructor",
// "__proto__", "toString") used to resolve to inherited properties of the
// lookup tables and made the whole slide render reject. They must behave like
// any other unknown name.
const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;
const factory: NodeCanvasFactory | null = Canvas && loadImage
  ? {
      createCanvas: (w, h) => new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
      loadImage: (buf) => loadImage(buf as Buffer) as unknown as ReturnType<NodeCanvasFactory['loadImage']>,
    }
  : null;

const px = (n: number) => Math.round(n * 9525);

function slideWith(geometry: string, warp: string | null, font: string, shadowAlign: string | null = null): Presentation {
  const shape = {
    type: 'shape', x: px(40), y: px(40), width: px(320), height: px(160), rotation: 0, flipH: false, flipV: false,
    geometry, fill: { type: 'solid', color: '3366CC' }, stroke: null, defaultTextColor: null, custGeom: null,
    adj: null, adj2: null, adj3: null, adj4: null, adj5: null, adj6: null, adj7: null, adj8: null,
    shadow: shadowAlign === null ? null : {
      color: '000000', alpha: 0.5, blur: px(4), dist: px(4), dir: 45, sx: 0.9, sy: 0.9, kx: 10, ky: 0, algn: shadowAlign,
    },
    textBody: {
      verticalAnchor: 'ctr', defaultFontSize: 40, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'none', vert: 'horz', autoFit: 'none',
      ...(warp ? { textWarp: { preset: warp } } : {}),
      paragraphs: [{
        alignment: 'ctr', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
        bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
        defFontFamily: null, tabStops: [], eaLnBrk: true,
        runs: [{
          type: 'text', text: 'HI', bold: false, italic: null, underline: false, strikethrough: false,
          fontSize: 40, color: '000000', fontFamily: font,
        }],
      }],
    },
  };
  return {
    slideWidth: px(400), slideHeight: px(240), defaultTextColor: null, majorFont: font, minorFont: font,
    slides: [{ index: 0, slideNumber: 1, background: null, elements: [shape] }],
  } as unknown as Presentation;
}

describe.skipIf(!factory)('document names that collide with Object.prototype members', () => {
  for (const name of ['constructor', '__proto__', 'toString']) {
    it(`renders a slide whose font, geometry, warp and shadow alignment are named "${name}"`, async () => {
      const cases = [
        slideWith(name, null, 'Arial'),
        slideWith('rect', name, 'Arial'),
        slideWith('rect', null, name),
        slideWith('rect', null, 'Arial', name),
      ];
      for (const pres of cases) {
        await expect(renderSlideNode(new Canvas(400, 240), pres, 0, { width: 400, dpr: 1, factory: factory! }))
          .resolves.not.toThrow();
      }
    });
  }
});
