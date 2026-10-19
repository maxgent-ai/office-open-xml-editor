/**
 * End-to-end pixel probe for the DrawingML `<a:duotone>` recolour (§20.1.8.23)
 * and `<a:alphaModFix>` opacity (§20.1.8.6) on a docx inline picture. A parser
 * round-trip is not proof the effects draw, so this builds a synthetic `.docx`
 * (in-memory STORED zip, no file committed) with a mid-grey PNG picture, parses
 * it through the real docx WASM, renders it through the docx renderer (skia +
 * the node OffscreenCanvas / createImageBitmap shims), and measures the image
 * region against the expected recolour / alpha composite.
 */
import { describe, it, expect } from 'vitest';
import { crc32 } from 'node:zlib';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import { importForTests, loadDocxRendererForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;
const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');
const rendererMod = await loadDocxRendererForTests();

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  // skia's loadImage wants a Buffer; the shim may hand us an ArrayBuffer/Uint8Array.
  loadImage: ((buf: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.isBuffer(buf) ? buf : Buffer.from(buf as ArrayBuffer))) as unknown as NodeCanvasFactory['loadImage'],
};

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

/** Minimal STORED zip supporting binary entries (Uint8Array) as well as text. */
function storedZip(files: Record<string, string | Uint8Array>): Uint8Array {
  const enc = new TextEncoder();
  const chunks: number[] = [];
  const central: number[] = [];
  const u16 = (n: number) => [n & 0xff, (n >> 8) & 0xff];
  const u32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = [...enc.encode(name)];
    const data = [...(typeof content === 'string' ? enc.encode(content) : content)];
    const crc = crc32(Uint8Array.from(data)) >>> 0;
    const local = [
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...nameBytes, ...data,
    ];
    central.push(
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset),
      ...nameBytes,
    );
    chunks.push(...local);
    offset += local.length;
  }
  const centralOffset = offset;
  const end = [
    ...u32(0x06054b50), ...u16(0), ...u16(0),
    ...u16(Object.keys(files).length), ...u16(Object.keys(files).length),
    ...u32(central.length), ...u32(centralOffset), ...u16(0),
  ];
  return Uint8Array.from([...chunks, ...central, ...end]);
}

/** A `<w:drawing>` inline picture, optionally carrying a duotone and/or an
 *  alphaModFix on its blip. cx/cy in EMU (12700 EMU/pt). */
interface PictureOptions {
  effects?: string;
  sizesPt?: number[];
  sizePt?: number;
  duotone?: boolean;
  alpha?: string;
  bright?: string;
  contrast?: string;
  inHeader?: boolean;
  grouped?: boolean;
  rotation?: number;
}

function inlinePicXml(opts: PictureOptions): string {
  const cx = (opts.sizePt ?? 200) * 12700; // 200pt square
  const effects = opts.effects ?? (
    (opts.duotone
      ? '<a:duotone><a:srgbClr val="000000"/><a:srgbClr val="DAB6BA"/></a:duotone>'
      : '') + (opts.alpha ? `<a:alphaModFix amt="${opts.alpha}"/>` : '')
      + (opts.bright || opts.contrast
        ? `<a:lum bright="${opts.bright ?? '0'}" contrast="${opts.contrast ?? '0'}"/>` : ''));
  return (
    '<w:p><w:r><w:drawing>' +
    `<wp:inline><wp:extent cx="${cx}" cy="${cx}"/>` +
    `<a:graphic><a:graphicData${opts.grouped ? ' uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"' : ''}>` +
    (opts.grouped ? `<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cx}"/><a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="${cx}"/></a:xfrm></wpg:grpSpPr>` : '') +
    '<pic:pic><pic:blipFill>' +
    `<a:blip r:embed="rIdImg">${effects}</a:blip>` +
    '<a:stretch><a:fillRect/></a:stretch>' +
    '</pic:blipFill>' +
    '<pic:spPr><a:xfrm rot="' + (opts.rotation ?? 0) * 60000 + '"><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cx + '"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    (opts.grouped ? '</wpg:wgp>' : '') +
    '</a:graphicData></a:graphic></wp:inline>' +
    '</w:drawing></w:r></w:p>'
  );
}

function buildDocx(pngBytes: Uint8Array, opts: PictureOptions): Uint8Array {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    (opts.inHeader ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : '') +
    '</Types>';
  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const imageRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
    '</Relationships>';
  const docRels = opts.inHeader
    ? imageRels.replace('</Relationships>', '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>')
    : imageRels;
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
    (opts.inHeader ? '<w:p/>' : (opts.sizesPt ?? [opts.sizePt ?? 200]).map(sizePt => inlinePicXml({ ...opts, sizePt })).join('')) +
    '<w:sectPr>' +
    (opts.inHeader ? '<w:headerReference w:type="default" r:id="rIdHeader"/>' : '') +
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/>' +
    '</w:sectPr></w:body></w:document>';
  return storedZip({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'word/_rels/document.xml.rels': docRels,
    'word/document.xml': document,
    'word/media/image1.png': pngBytes,
    ...(opts.inHeader ? {
      'word/header1.xml': `<w:hdr ${NS}>${inlinePicXml(opts)}</w:hdr>`,
      'word/_rels/header1.xml.rels': imageRels,
    } : {}),
  });
}

async function flatPngBytes(w: number, h: number, color: string): Promise<Uint8Array> {
  const c = new Canvas(w, h);
  const ctx = c.getContext('2d');
  ctx.fillStyle = color;
  ctx.fillRect(0, 0, w, h);
  return new Uint8Array(await c.toBuffer('png'));
}

interface Rendered {
  data: Uint8ClampedArray;
  w: number;
  h: number;
}

async function render(docxBytes: Uint8Array, pngBytes: Uint8Array): Promise<Rendered> {
  const { materializeDocxDocument } = docxMod!;
  const doc = await materializeDocxDocument(docxBytes);
  const { renderDocumentToCanvas } = rendererMod!;
  const widthPx = doc.section.pageWidth;
  const heightPx = doc.section.pageHeight;
  const canvas = new Canvas(Math.round(widthPx), Math.round(heightPx));
  const restoreImg = installImageBitmapShim(factory);
  const restoreOff = installOffscreenCanvasShim(factory);
  try {
    await renderDocumentToCanvas(doc, canvas as unknown as OffscreenCanvas, 0, {
      dpr: 1,
      width: widthPx,
      fetchImage: async (_path: string, mime: string) =>
        new Blob([pngBytes as BlobPart], { type: mime }),
    });
  } finally {
    restoreOff();
    restoreImg();
  }
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  return { data: img.data, w: canvas.width, h: canvas.height };
}

function lum(r: number, g: number, b: number): number {
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
}

/** Sample the picture interior. The inline picture sits at the top-left content
 *  origin (~720 twip = 48px margin), 200pt square, so a point ~60px inside the
 *  margin is safely inside the image. */
function sampleInside(r: Rendered): [number, number, number] {
  const x = 48 + 40;
  const y = 48 + 40;
  const i = (y * r.w + x) * 4;
  return [r.data[i], r.data[i + 1], r.data[i + 2]];
}

describe.skipIf(!skia || !docxMod || !rendererMod)('docx duotone + alphaModFix picture render', () => {
  it('preserves host orientation through DOCX body and header resource acquisition', async () => {
    const png = await flatPngBytes(16, 16, '#FF0000');
    const paint = async (rotWithShape: boolean, inHeader: boolean) => {
      const effects = `<a:alphaMod><a:cont type="tree"><a:outerShdw dist="1270000" dir="0" rotWithShape="${rotWithShape}"><a:srgbClr val="000000"/></a:outerShdw></a:cont></a:alphaMod>`;
      const bytes = buildDocx(png, { effects, rotation: 90, inHeader });
      const result = await render(bytes, png);
      // Inside the east-shifted mask, outside the south-shifted mask.
      const i = ((48 + 40) * result.w + 48 + 160) * 4;
      return [...result.data.slice(i, i + 3)];
    };
    for (const inHeader of [false, true]) {
      expect(await paint(true, inHeader), `rotating shadow, header=${inHeader}`).toEqual([255, 255, 255]);
      expect(await paint(false, inHeader), `page-fixed shadow, header=${inHeader}`).toEqual([255, 0, 0]);
    }
  });
  it('renders authored brightness and contrast through DOCX parsing and paint (§20.1.8.42)', async () => {
    const black = await flatPngBytes(16, 16, '#000000');
    expect(sampleInside(await render(buildDocx(black, {}), black))).toEqual([0, 0, 0]);
    for (const inHeader of [false, true]) {
      const bright = sampleInside(await render(buildDocx(black, { bright: '90000', inHeader }), black));
      for (const channel of bright) expect(Math.abs(channel - 230)).toBeLessThanOrEqual(1);
    }
    const contrast = sampleInside(await render(buildDocx(black, { contrast: '-50000' }), black));
    for (const channel of contrast) expect(Math.abs(channel - 64)).toBeLessThanOrEqual(1);
  });

  it('preserves the existing grouped-picture transparent colour with alphaModFix', async () => {
    const png = await flatPngBytes(16, 16, '#000000');
    expect(sampleInside(await render(buildDocx(png, { grouped: true }), png))).toEqual([0, 0, 0]);
    const effects = '<a:alphaModFix amt="50000"/>' +
      '<a:clrChange useA="1"><a:clrFrom><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:clrFrom>' +
      '<a:clrTo><a:srgbClr val="000000"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>';
    expect(sampleInside(await render(buildDocx(png, { effects, grouped: true }), png)))
      .toEqual([255, 255, 255]);
  });

  it('keeps an inline alpha-sensitive colour nonmatch unchanged', async () => {
    const png = await flatPngBytes(16, 16, '#000000');
    const effects = '<a:alphaModFix amt="50000"/>' +
      '<a:clrChange useA="1"><a:clrFrom><a:srgbClr val="000000"/></a:clrFrom>' +
      '<a:clrTo><a:srgbClr val="FF0000"/></a:clrTo></a:clrChange>';
    const pixel = sampleInside(await render(buildDocx(png, { effects }), png));
    for (const channel of pixel) expect(Math.abs(channel - 128)).toBeLessThanOrEqual(1);
  });

  it('does not reapply the first duotone for every palette marker', async () => {
    const png = await flatPngBytes(16, 16, '#FFFFFF');
    // Both the established first-palette-only path and the authored two
    // palettes produce red. Reusing the first palette twice darkens it.
    const effects = '<a:grayscl/>' +
      '<a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FF0000"/></a:duotone>' +
      '<a:duotone><a:srgbClr val="FF0000"/><a:srgbClr val="FF0000"/></a:duotone>';
    expect(sampleInside(await render(buildDocx(png, { effects }), png))).toEqual([255, 0, 0]);
  });

  it('preserves alpha-sensitive replacement and independent palettes through body and header paint', async () => {
    const black = await flatPngBytes(16, 16, '#000000');
    const chains = [
      { effects: '<a:alphaModFix amt="50000"/><a:clrChange useA="1">'
        + '<a:clrFrom><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:clrFrom>'
        + '<a:clrTo><a:srgbClr val="FF0000"/></a:clrTo></a:clrChange>', expected: [255, 0, 0] },
      { effects: '<a:lum bright="90000"/>'
        + '<a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FF0000"/></a:duotone>'
        + '<a:duotone><a:srgbClr val="000000"/><a:srgbClr val="00FF00"/></a:duotone>', expected: [0, 69, 0] },
      { effects: '<a:lum bright="90000"/><a:clrRepl><a:srgbClr val="FF0000"/></a:clrRepl>', expected: [255, 0, 0] },
    ];
    for (const chain of chains) {
      for (const inHeader of [false, true]) {
        const actual = sampleInside(await render(buildDocx(black, { effects: chain.effects, inHeader }), black));
        actual.forEach((value, index) => expect(Math.abs(value - chain.expected[index])).toBeLessThanOrEqual(1));
      }
    }
  });

  it('renders recursive alpha, blur, HSL, tint and fill overlay in body and header pictures', async () => {
    const black = await flatPngBytes(16, 16, '#000000');
    const effects = '<a:clrRepl><a:srgbClr val="FF0000"/></a:clrRepl>'
      + '<a:hsl hue="7200000"/><a:tint hue="14400000" amt="50000"/>'
      + '<a:fillOverlay blend="over"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:fillOverlay>'
      + '<a:alphaMod><a:cont type="tree"><a:alphaRepl a="50%"/></a:cont></a:alphaMod>'
      + '<a:blur rad="12700" grow="false"/>';
    // Green hue rotation -> cyan tint -> half-red overlay -> half alpha over
    // white: independently computed byte result (191,159,159).
    for (const inHeader of [false, true]) {
      const actual = sampleInside(await render(buildDocx(black, { effects, inHeader }), black));
      actual.forEach((value, index) => expect(Math.abs(value - [191, 159, 159][index])).toBeLessThanOrEqual(1));
    }
  });

  it('duotone remaps a mid-grey picture along the clr1→clr2 ramp (§20.1.8.23)', async () => {
    const png = await flatPngBytes(16, 16, '#808080');
    const t = lum(128, 128, 128);
    const c1 = [0x00, 0x00, 0x00];
    const c2 = [0xda, 0xb6, 0xba];
    const expected = c1.map((d, i) => Math.round(d + (c2[i] - d) * t));

    const plain = sampleInside(await render(buildDocx(png, {}), png));
    const duo = sampleInside(await render(buildDocx(png, { duotone: true }), png));

    // Plain: untouched grey.
    expect(Math.abs(plain[0] - 128)).toBeLessThan(8);
    expect(Math.abs(plain[1] - 128)).toBeLessThan(8);
    // Duotone: the luminance-ramp interpolation (a light pink, R dominant).
    expect(Math.abs(duo[0] - expected[0])).toBeLessThan(12);
    expect(Math.abs(duo[1] - expected[1])).toBeLessThan(12);
    expect(Math.abs(duo[2] - expected[2])).toBeLessThan(12);
    expect(duo[0]).toBeGreaterThan(duo[1]);
    expect(duo[0]).toBeGreaterThan(duo[2]);
  });

  it('alphaModFix composites a semi-transparent picture over the white page (§20.1.8.6)', async () => {
    // amt=50000 → 0.5 opacity. A mid-grey (128) picture over white (255) at 0.5
    // composites to 0.5*128 + 0.5*255 ≈ 191.5 on every channel.
    const png = await flatPngBytes(16, 16, '#808080');
    const opaque = sampleInside(await render(buildDocx(png, {}), png));
    const semi = sampleInside(await render(buildDocx(png, { alpha: '50000' }), png));

    // Opaque baseline is the flat grey; the semi-transparent draw is lighter
    // (composited toward white) by a clear margin.
    expect(Math.abs(opaque[0] - 128)).toBeLessThan(8);
    const expectedComposite = Math.round(0.5 * 128 + 0.5 * 255);
    expect(Math.abs(semi[0] - expectedComposite)).toBeLessThan(14);
    expect(semi[0]).toBeGreaterThan(opaque[0] + 30);
  });
});


it.skipIf(!skia || !docxMod || !rendererMod)('keeps the DOCX 8pt picture unchanged beside an 80pt copy sharing its PNG and 1pt blur', async () => {
  const png = await flatPngBytes(16, 16, '#ff0000');
  const effects = '<a:blur rad="12700" grow="false"/>';
  const alone = await render(buildDocx(png, { effects, sizesPt: [8] }), png);
  const copies = await render(buildDocx(png, { effects, sizesPt: [8, 80] }), png);
  // The first paragraph and its full 8pt image occupy the same page region.
  const region = (r: Rendered) => Array.from({ length: 8 }, (_, y) =>
    [...r.data.slice(((36 + y) * r.w + 36) * 4, ((36 + y) * r.w + 44) * 4)]).flat();
  expect(region(copies)).toEqual(region(alone));
  expect(region(alone).some((v, i) => i % 4 === 1 && v > 0 && v < 255)).toBe(true);
});
