/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import type { DocxDocumentModel } from './types.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { storeZip } from './conformance/generate.js';

// Issue #1623 synthetic reductions of the Word for Mac controls behind
// word-mode14-column-line-start-origin, word-tight-wrap-bottom-edge,
// word-tight-wrap-line-step-advance, word-mode14-tight-anchor-line-rewrap,
// word-later-anchor-earlier-line-wrap and word-mode14-tight-anchor-top-touch.
// Geometry: US Letter, 1 in margins, one label per paragraph, exact 24 pt
// lines (paragraph n starts at 72 + 24n without floats), pictures with
// distL/distR 9 pt and allowOverlap=1. Every expected value is the Word
// placement of the corresponding control.

const EMU = 12700;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0xf0,
  0x1f, 0x00, 0x05, 0x00, 0x01, 0xff, 0x89, 0x99, 0x3d, 0x1d, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
  0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

interface Picture {
  readonly paragraph: number;
  readonly wrap: 'tight' | 'square' | 'none';
  readonly allowOverlap?: boolean;
  readonly hFrom: 'column' | 'margin';
  readonly xPt: number;
  readonly yPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
}

interface Control {
  readonly mode: 14 | 15;
  readonly pictures: readonly Picture[];
  readonly empty?: readonly number[];
  readonly lineHeightPt?: number;
}

const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);

function anchor(picture: Picture, index: number): string {
  const w = Math.round(picture.widthPt * EMU);
  const h = Math.round(picture.heightPt * EMU);
  const wrap = picture.wrap === 'none' ? '<wp:wrapNone/>' : picture.wrap === 'tight'
    ? '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/>'
      + '<wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/>'
      + '<wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'
    : '<wp:wrapSquare wrapText="bothSides"/>';
  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="${9 * EMU}" distR="${9 * EMU}" `
    + `simplePos="0" relativeHeight="${251659264 + index * 1024}" behindDoc="0" locked="0" `
    + `layoutInCell="1" allowOverlap="${picture.allowOverlap === false ? 0 : 1}"><wp:simplePos x="0" y="0"/>`
    + `<wp:positionH relativeFrom="${picture.hFrom}"><wp:posOffset>${Math.round(picture.xPt * EMU)}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${Math.round(picture.yPt * EMU)}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${w}" cy="${h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}`
    + `<wp:docPr id="${index + 1}" name="Picture ${index + 1}"/>`
    + `<a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="${index + 1}" name="p.png"/>`
    + `<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch>`
    + `</pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
    + '</wp:anchor></w:drawing></w:r>';
}

function docx(control: Control): Uint8Array {
  const paragraphs: string[] = [];
  for (let p = 0; p < 30; p += 1) {
    const label = control.empty?.includes(p) ? '' : `<w:r><w:t>P${String(p).padStart(3, '0')}</w:t></w:r>`;
    const drawings = control.pictures
      .map((picture, index) => (picture.paragraph === p ? anchor(picture, index) : ''))
      .join('');
    paragraphs.push('<w:p><w:pPr><w:widowControl w:val="0"/>'
      + `<w:spacing w:before="0" w:after="0" w:line="${(control.lineHeightPt ?? 24) * 20}" w:lineRule="exact"/></w:pPr>`
      + `${label}${drawings}</w:p>`);
  }
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}">`
    + `<w:body>${paragraphs.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>`
    + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    + '</w:sectPr></w:body></w:document>';
  const styles = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>`
    + '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>';
  const settings = `<w:settings xmlns:w="${W}"><w:compat><w:compatSetting w:name="compatibilityMode" `
    + `w:uri="http://schemas.microsoft.com/office/word" w:val="${control.mode}"/></w:compat></w:settings>`;
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return storeZip(new Map([
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + '</Types>')],
    ['_rels/.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rIdStyles" Type="${rel}/styles" Target="styles.xml"/>`
      + `<Relationship Id="rIdSettings" Type="${rel}/settings" Target="settings.xml"/>`
      + `<Relationship Id="rIdImg" Type="${rel}/image" Target="media/p.png"/></Relationships>`)],
    ['word/document.xml', xml(document)],
    ['word/styles.xml', xml(styles)],
    ['word/settings.xml', xml(settings)],
    ['word/media/p.png', PNG],
  ]));
}

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'normal',
    measureText: (text: string) => {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px * 0.5,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

interface Placed {
  readonly lines: ReadonlyMap<number, Readonly<{ yPt: number; xPt: number }>>;
  readonly pictures: readonly Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>[];
}

function layOut(control: Control): Placed {
  const archive = new DocxArchive(docx(control));
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(
      JSON.parse(new TextDecoder().decode(archive.parse())),
    ).document;
  } finally {
    archive.free();
  }
  const layout = layoutDocument(
    model,
    createLayoutServices(model, { measureContext: measureContext() }),
    { currentDateMs: 0 },
  );
  const lines = new Map<number, Readonly<{ yPt: number; xPt: number }>>();
  const pictures: Array<Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>> = [];
  const page = layout.pages[0]!;
  for (const node of page.layers.body) {
    if (node.kind !== 'paragraph') continue;
    const index = node.source.path[0] as number;
    const line = node.lines[0];
    if (line && !lines.has(index)) lines.set(index, { yPt: line.bounds.yPt, xPt: line.bounds.xPt });
    for (const drawing of node.drawings) {
      pictures.push({ ...drawing.flowBounds });
    }
  }
  return { lines, pictures };
}

// Picture A (P008) and B (P009) of the reduced corpus page: B overlaps A.
const pair = (
  wrap: Picture['wrap'],
  hFrom: Picture['hFrom'],
  ax = -15.2,
  bx = 7.05,
  by = 0,
  bParagraph = 9,
): readonly Picture[] => [
  { paragraph: 8, wrap, hFrom, xPt: ax, yPt: 20.85, widthPt: 238, heightPt: 178.4 },
  { paragraph: bParagraph, wrap, hFrom, xPt: bx, yPt: by, widthPt: 223.6, heightPt: 180 },
];

// One picture B in P009 at the margin edge.
const single = (wrap: Picture['wrap'], yPt: number, widthPt = 200, heightPt = 96): readonly Picture[] => [
  { paragraph: 9, wrap, hFrom: 'margin', xPt: 0, yPt, widthPt, heightPt },
];

const close = (actual: number | undefined, expected: number) => expect(actual).toBeCloseTo(expected, 1);

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('issue #1623 Word placement of floats from different paragraphs', () => {
  // Resource-policy regression, not a new Office compatibility claim: the
  // parser-backed layout must remain total after exhausting tight line steps.
  it('lays out separated tight regions bridged by a square after the step limit', () => {
    const placed = layOut({
      mode: 15,
      lineHeightPt: 0.5,
      pictures: [
        { paragraph: 0, wrap: 'tight', hFrom: 'margin', xPt: 0, yPt: 0, widthPt: 468, heightPt: 10_000 },
        { paragraph: 0, wrap: 'square', hFrom: 'margin', xPt: 0, yPt: 9000, widthPt: 468, heightPt: 21_000 },
        { paragraph: 0, wrap: 'tight', hFrom: 'margin', xPt: 0, yPt: 20_000, widthPt: 468, heightPt: 5000 },
      ],
    });
    expect(placed.lines.size).toBeGreaterThan(0);
    for (const line of placed.lines.values()) expect(Number.isFinite(line.yPt)).toBe(true);
  }, 30_000);

  it('keeps overlap-permitted pictures from different paragraphs at their resolved positions', () => {
    const placed = layOut({ mode: 15, pictures: pair('tight', 'column') });
    close(placed.pictures[0]?.xPt, 56.8);
    close(placed.pictures[1]?.xPt, 79.05);
    close(placed.pictures[1]?.yPt, 288);
    close(placed.lines.get(9)?.xPt, 311.65);
  });

  it('measures a mode-14 column offset from the anchor line start around earlier floats', () => {
    const right = layOut({ mode: 14, pictures: pair('square', 'column') });
    close(right.pictures[1]?.xPt, 303.8 + 7.05);
    close(right.pictures[1]?.yPt, 288);
    close(right.lines.get(9)?.yPt, 463.25);
    const left = layOut({ mode: 14, pictures: pair('square', 'column', 245.2, 237.35) });
    close(left.pictures[1]?.xPt, 72 + 237.35);
    const margin = layOut({ mode: 14, pictures: pair('square', 'margin') });
    close(margin.pictures[1]?.xPt, 79.05);
  });

  it('wraps a line that starts on the bottom edge of a tight polygon', () => {
    const tight = layOut({ mode: 15, pictures: single('tight', 0) });
    close(tight.lines.get(13)?.xPt, 281);
    close(tight.lines.get(14)?.xPt, 72);
    const square = layOut({ mode: 15, pictures: single('square', 0) });
    close(square.lines.get(13)?.xPt, 72);
  });

  it('moves a line beside a tight polygon down in whole line heights', () => {
    for (const [yPt, expected] of [[0, 408], [6, 408]] as const) {
      const tight = layOut({ mode: 14, pictures: single('tight', yPt, 468, 100) });
      close(tight.lines.get(8)?.yPt, 264);
      close(tight.lines.get(9)?.yPt, expected);
    }
    const square = layOut({ mode: 14, pictures: single('square', 0, 468, 100) });
    close(square.lines.get(9)?.yPt, 388);
  });

  it('lays a mode-14 tight anchor line out around its object only when the object meets its content', () => {
    const mode14 = layOut({ mode: 14, pictures: pair('tight', 'margin') });
    close(mode14.lines.get(9)?.xPt, 303.8);
    close(mode14.lines.get(10)?.xPt, 311.65);
    const mode15 = layOut({ mode: 15, pictures: pair('tight', 'margin') });
    close(mode15.lines.get(9)?.xPt, 311.65);
    const own = layOut({ mode: 14, pictures: single('tight', 0) });
    close(own.lines.get(9)?.xPt, 281);
  });

  it('wraps earlier lines around a later paragraph-relative picture that reaches up into them', () => {
    for (const control of [
      { mode: 14, pictures: single('tight', -6) },
      { mode: 15, pictures: single('tight', -6) },
      { mode: 14, pictures: single('square', -6) },
    ] as const) {
      const placed = layOut(control);
      close(placed.pictures[0]?.yPt, 282);
      close(placed.lines.get(8)?.xPt, 281);
      close(placed.lines.get(9)?.xPt, 281);
    }
  });

  it('blocks a mode-14 line touching the top of a tight picture whose anchor line had to move', () => {
    for (const empty of [[], [8, 9]]) {
      const placed = layOut({ mode: 14, pictures: pair('tight', 'column'), empty });
      close(placed.pictures[0]?.yPt, 284.85);
      close(placed.pictures[1]?.xPt, 310.85);
      close(placed.pictures[1]?.yPt, 288);
      close(placed.lines.get(10)?.yPt, 528);
    }
    const text = layOut({ mode: 14, pictures: pair('tight', 'column') });
    close(text.lines.get(8)?.yPt, 480);
    close(text.lines.get(9)?.yPt, 504);
    const later = layOut({ mode: 14, pictures: pair('tight', 'column', -15.2, 7.05, 0, 10) });
    close(later.lines.get(8)?.yPt, 264);
    close(later.lines.get(9)?.yPt, 480);
    close(later.lines.get(10)?.yPt, 504);
    const lower = layOut({ mode: 14, pictures: pair('tight', 'column', -15.2, 7.05, 1) });
    close(lower.lines.get(8)?.yPt, 264);
    close(lower.lines.get(9)?.yPt, 480);
    const mode15 = layOut({ mode: 15, pictures: pair('tight', 'margin', -15.2, 238.85) });
    close(mode15.lines.get(8)?.yPt, 264);
    close(mode15.lines.get(9)?.yPt, 480);
  });

  it('keeps allowOverlap=false separation when a carried picture moves an earlier paragraph', () => {
    const placed = layOut({
      mode: 15,
      pictures: [
        { paragraph: 8, wrap: 'none', hFrom: 'margin', xPt: 0, yPt: -200, widthPt: 100, heightPt: 100 },
        {
          paragraph: 9, wrap: 'tight', hFrom: 'margin', xPt: 0, yPt: -30,
          widthPt: 468, heightPt: 96, allowOverlap: false,
        },
      ],
    });
    const [first, second] = placed.pictures;
    expect(first && second).toBeTruthy();
    const overlapX = Math.min(first!.xPt + first!.widthPt, second!.xPt + second!.widthPt)
      - Math.max(first!.xPt, second!.xPt);
    const overlapY = Math.min(first!.yPt + first!.heightPt, second!.yPt + second!.heightPt)
      - Math.max(first!.yPt, second!.yPt);
    expect(overlapX <= 0.001 || overlapY <= 0.001).toBe(true);
  });

  it('carries a later tight picture to an earlier line that starts on its bottom edge', () => {
    const placed = layOut({
      mode: 15,
      // The picture spans 48-72 in the top margin; only P000 (72-96) touches it.
      pictures: [{ paragraph: 1, wrap: 'tight', hFrom: 'margin', xPt: 0, yPt: -48, widthPt: 200, heightPt: 24 }],
    });
    close(placed.pictures[0]?.yPt, 48);
    close(placed.lines.get(0)?.xPt, 281);
    close(placed.lines.get(1)?.xPt, 72);
  });
});
