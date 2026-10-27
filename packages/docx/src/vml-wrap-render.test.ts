/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { Canvas, loadImage } from 'skia-canvas';
import { afterEach, beforeAll, expect, it, vi } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { renderDocumentToCanvas } from './renderer.js';

const encode = (s: string) => new TextEncoder().encode(s);

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});
afterEach(() => vi.unstubAllGlobals());

// One floating 144pt VML object at the column's left edge above three
// paragraphs.
// The page is 400pt wide with 20pt margins and is painted at 1px per point.
async function render(pict: string) {
  // Opaque blue left half, transparent right half: text that tight wrapping
  // lets into the polygon-free half stays visible.
  const source = new Canvas(144, 144);
  source.getContext('2d').fillStyle = '#0000ff';
  source.getContext('2d').fillRect(0, 0, 72, 144);
  const png = new Uint8Array(source.toBufferSync('png'));
  vi.stubGlobal('OffscreenCanvas', Canvas);
  vi.stubGlobal('createImageBitmap', async (s: Blob | Canvas) => loadImage(s instanceof Blob
    ? Buffer.from(await s.arrayBuffer()) : s.toBufferSync('png')));
  const text = '<w:p><w:r><w:t>Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</w:t></w:r></w:p>';
  const bytes = storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="pic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/blue.png"/></Relationships>')],
    ['word/media/blue.png', png],
    ['word/document.xml', encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:pict>${pict}</w:pict></w:r></w:p>${text.repeat(3)}<w:sectPr><w:pgSz w:w="8000" w:h="8000"/><w:pgMar w:top="400" w:right="400" w:bottom="400" w:left="400" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`)],
  ]));
  const archive = new DocxArchive(bytes);
  let model;
  try { model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document; }
  finally { archive.free(); }
  const canvas = new Canvas(400, 400);
  const diagnostics: unknown[] = [];
  await renderDocumentToCanvas(model, canvas as unknown as HTMLCanvasElement, 0, {
    width: 400, dpr: 1, fetchImage: async () => new Blob([png], { type: 'image/png' }),
    onDiagnostic: d => diagnostics.push(d),
  });
  const pixels = canvas.getContext('2d').getImageData(0, 0, 400, 400).data;
  const count = (x0: number, x1: number, y0: number, y1: number, match: (i: number) => boolean) => {
    let n = 0;
    for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) if (match((y * 400 + x) * 4)) n++;
    return n;
  };
  const layout = layoutDocument(model, createLayoutServices(model, {
    measureContext: canvas.getContext('2d') as unknown as CanvasRenderingContext2D,
  }));
  return {
    pixels,
    diagnostics: [...layout.diagnostics, ...diagnostics],
    blue: count(0, 400, 0, 400, i => pixels[i + 2]! > 180 && pixels[i]! < 60),
    // Dark text ink over the right half of the object, within its height.
    inkOverRightHalf: count(20 + 80, 20 + 140, 30, 150, i => pixels[i]! < 128 && pixels[i + 2]! < 128),
  };
}

// The picture is painted 1:1, so its blue left half is exactly 72 by 144px.
const expectLeftHalfPainted = (blue: number) => expect(blue).toBe(72 * 144);

const RELATIVE = ';mso-position-horizontal-relative:text;mso-position-vertical-relative:text';
const LEFT_HALF = '0 0 0 21600 10800 21600 10800 0 0 0';
const picture = (wrap: string, coords: string | null, extra = RELATIVE) =>
  `<v:shape type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:144pt;height:144pt;z-index:1${extra}"${coords === null ? '' : ` wrapcoords="${coords}"`}><v:imagedata r:id="pic" o:title=""/>${wrap}</v:shape>`;

// Issue #1959: tight/through VML anchors lost their polygon, so retained
// layout rejected the anchor and the object was never painted.
it.each(['tight', 'through', 'square', 'topAndBottom', 'none'])('paints a floating VML picture with %s wrapping', async type => {
  for (const extra of [RELATIVE, '']) {
    const actual = await render(picture(`<w10:wrap type="${type}"/>`, LEFT_HALF, extra));
    expect(actual.diagnostics).toEqual([]);
    expectLeftHalfPainted(actual.blue);
  }
});

it.each(['tight', 'through'])('wraps %s text around the wrapcoords polygon, not the frame', async type => {
  const square = await render(picture('<w10:wrap type="square"/>', LEFT_HALF));
  expect(square.inkOverRightHalf).toBe(0);
  // MS-OE376 §2.1.1692(uu-vv): space-delimited points in a 21600 square mapped
  // onto the object frame; Part 4's comma-delimited form means the same.
  const reference = await render(picture(`<w10:wrap type="${type}"/>`, LEFT_HALF));
  expectLeftHalfPainted(reference.blue);
  expect(reference.inkOverRightHalf).toBeGreaterThan(100);
  const template = (coords: string) => `<v:shapetype id="half" wrapcoords="${coords}"/>`;
  for (const content of [
    picture(`<w10:wrap type="${type}"/>`, '0,0,0,21600,10800,21600,10800,0,0,0'),
    // `wrapcoords` is also a shapetype core attribute; the instance overrides it.
    template(LEFT_HALF) + picture(`<w10:wrap type="${type}"/>`, null).replace('#_x0000_t75', '#half'),
    template('0 0 0 21600 21600 21600 21600 0') + picture(`<w10:wrap type="${type}"/>`, LEFT_HALF).replace('#_x0000_t75', '#half'),
  ]) {
    const actual = await render(content);
    expect(actual.diagnostics).toEqual([]);
    expect(Buffer.compare(Buffer.from(actual.pixels), Buffer.from(reference.pixels))).toBe(0);
  }
});

it('inherits wrapcoords from an earlier pict beside a group', async () => {
  // A pict holding a group acquires its sibling leaves separately; the
  // shapetype written in an earlier pict of the story must still apply.
  // Unfilled and unstroked, the leaf shows the text its polygon admits.
  const leaf = (coords: string | null) => picture('<w10:wrap type="tight"/>', coords)
    .replace('type="#_x0000_t75"', 'type="#half" filled="f" stroked="f"')
    + '<v:group style="position:absolute;width:1pt;height:1pt" coordsize="1,1"/>';
  const own = await render(`<v:shapetype id="half"/></w:pict><w:pict>${leaf(LEFT_HALF)}`);
  expectLeftHalfPainted(own.blue);
  expect(own.inkOverRightHalf).toBeGreaterThan(100);
  const inherited = await render(`<v:shapetype id="half" wrapcoords="${LEFT_HALF}"/></w:pict><w:pict>${leaf(null)}`);
  expect(inherited.diagnostics).toEqual([]);
  expect(Buffer.compare(Buffer.from(inherited.pixels), Buffer.from(own.pixels))).toBe(0);
});

it.each(['tight', 'through'])('wraps %s text around the frame when wrapcoords is absent or unusable', async type => {
  const frame = await render(picture(`<w10:wrap type="${type}"/>`, '0 0 0 21600 21600 21600 21600 0 0 0'));
  expectLeftHalfPainted(frame.blue);
  expect(frame.inkOverRightHalf).toBe(0);
  for (const coords of [null, '', '0 0 0 21600 10800', '0 0 0 21600 10800.5 21600 10800 0', '0 0 0 21600']) {
    const actual = await render(picture(`<w10:wrap type="${type}"/>`, coords));
    expect(actual.diagnostics).toEqual([]);
    expect(Buffer.compare(Buffer.from(actual.pixels), Buffer.from(frame.pixels))).toBe(0);
  }
});

it('paints tight-wrapped VML shapes and groups with their wrapcoords', async () => {
  const rect = (style: string, coords: string) =>
    `<v:rect style="position:absolute;${style}"${coords} filled="f" strokecolor="blue" strokeweight="4pt"><w10:wrap type="tight"/></v:rect>`;
  const frame = '0 0 0 21600 21600 21600 21600 0';
  for (const content of [
    (coords: string) => rect(`margin-left:0;margin-top:0;width:144pt;height:144pt${RELATIVE}`, ` wrapcoords="${coords}"`),
    (coords: string) => `<v:group style="position:absolute;margin-left:0;margin-top:0;width:144pt;height:144pt${RELATIVE}" coordsize="144,144" wrapcoords="${coords}">${rect('left:0;top:0;width:144;height:144', '')}<w10:wrap type="tight"/></v:group>`,
  ]) {
    // The unfilled outline leaves text visible, so the full-frame control
    // proves the exclusion is applied and the half polygon proves its shape.
    for (const [coords, ink] of [[frame, 0], [LEFT_HALF, 100]] as const) {
      const actual = await render(content(coords));
      expect(actual.diagnostics).toEqual([]);
      expect(actual.blue).toBeGreaterThan(4 * 144);
      if (ink === 0) expect(actual.inkOverRightHalf).toBe(0);
      else expect(actual.inkOverRightHalf).toBeGreaterThan(ink);
    }
  }
});

it('wraps around the bounding rectangle of a self-intersecting polygon beyond the vertex budget', async () => {
  // Crossing count, and with it exact compilation work, is quadratic in
  // vertices. 300 pseudo-random vertices plus the frame corners exceed the
  // budget and must wrap exactly like the full-frame polygon.
  let seed = 12345;
  const next = () => (seed = (Math.imul(1664525, seed) + 1013904223) >>> 0) % 21601;
  const scribble = ['0 0', '21600 21600', ...Array.from({ length: 300 }, () => `${next()} ${next()}`)].join(' ');
  const frame = await render(picture('<w10:wrap type="tight"/>', '0 0 0 21600 21600 21600 21600 0'));
  const actual = await render(picture('<w10:wrap type="tight"/>', scribble));
  expect(actual.diagnostics).toEqual([]);
  expect(Buffer.compare(Buffer.from(actual.pixels), Buffer.from(frame.pixels))).toBe(0);
});

it('discloses wrapcoords on a shape that keeps legacy placement without wrapping', async () => {
  // Without complete relative frames a VML shape keeps its legacy placement,
  // which applies no wrapping; its polygon must not vanish without a marker.
  const shape = (coords: string) => `<v:rect style="position:absolute;margin-left:0;margin-top:0;width:144pt;height:144pt"${coords} filled="f" strokecolor="blue" strokeweight="4pt"><w10:wrap type="tight"/></v:rect>`;
  expect((await render(shape(''))).diagnostics).toEqual([]);
  const legacy = await render(shape(` wrapcoords="${LEFT_HALF}"`));
  expect(legacy.blue).toBeGreaterThan(4 * 144);
  expect(legacy.diagnostics).not.toEqual([]);
});

it('wraps around the bounding rectangle of a polygon beyond the vertex budget', async () => {
  // 150,000 vertices along the left half's outline: their bounding rectangle
  // is the left half, so the result equals the four-point polygon.
  const edge = Array.from({ length: 50_000 }, (_, i) => `${Math.round(i * 10800 / 50_000)} 21600`);
  const outline = ['0 0', '0 21600', ...edge, ...edge.map((_, i) => `10800 ${21600 - Math.round(i * 21600 / 50_000)}`),
    ...edge.map((_, i) => `${10800 - Math.round(i * 10800 / 50_000)} 0`)].join(' ');
  const reference = await render(picture('<w10:wrap type="tight"/>', LEFT_HALF));
  const actual = await render(picture('<w10:wrap type="tight"/>', outline));
  expect(actual.diagnostics).toEqual([]);
  expect(Buffer.compare(Buffer.from(actual.pixels), Buffer.from(reference.pixels))).toBe(0);
});
