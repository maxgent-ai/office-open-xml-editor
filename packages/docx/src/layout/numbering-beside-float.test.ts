/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { measureContext } from '../test-support/tab-fitting.test-support.js';
import { storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

// A 144pt square-wrapped picture at the 54pt margin, distR 9pt: wrap edge 207pt.
const float = '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="F"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="F"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdMissing"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';

function bytes(ind: string, levelTabs = '', styleOrigin?: string, variant: { empty?: boolean; rtl?: boolean; defaultsPPr?: string } = {}): Uint8Array {
  const enc = new TextEncoder();
  const p = (body: string, ppr = '') => `<w:p><w:pPr>${ppr}<w:spacing w:before="0" w:after="0"/></w:pPr>${body}</w:p>`;
  const item = p(variant.empty ? '' : '<w:r><w:t>word</w:t></w:r>', (styleOrigin === undefined
    ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : '<w:pStyle w:val="ListItem"/>')
    + (variant.rtl ? '<w:bidi/>' : ''));
  const floatXml = variant.rtl
    ? float.replace('<wp:posOffset>0</wp:posOffset></wp:positionH>', '<wp:align>right</wp:align></wp:positionH>')
    : float;
  const styles = `<w:styles ${W}><w:docDefaults><w:pPrDefault><w:pPr>${variant.defaultsPPr ?? ''}</w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="ListItem"><w:name w:val="List Item"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${styleOrigin ?? ''}</w:pPr></w:style></w:styles>`;
  const doc = `<w:document ${W}><w:body>${p(`${floatXml}<w:r><w:t>host</w:t></w:r>`)}${item}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const numbering = `<w:numbering ${W}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr>${levelTabs}${ind}</w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  return storeZip(new Map([
    ['[Content_Types].xml', enc.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')],
    ['_rels/.rels', enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')],
    ['word/document.xml', enc.encode(doc)],
    ['word/numbering.xml', enc.encode(numbering)],
    ['word/styles.xml', enc.encode(styles)],
  ]));
}

function firstItemLine(input: Uint8Array) {
  const archive = new DocxArchive(input);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(5) }), { currentDateMs: 0 });
  const paragraphs = layout.pages[0]!.layers.body.filter((node) => node.kind === 'paragraph');
  const line = paragraphs[1]!.kind === 'paragraph' ? paragraphs[1]!.lines[0]! : undefined;
  const texts = line!.placements.filter((node) => node.kind === 'text');
  const body = texts.find((node) => node.role !== 'numbering-marker')?.bounds;
  return {
    marker: texts.find((node) => node.role === 'numbering-marker')?.bounds.xPt,
    body: body?.xPt,
    ...(body ? { bodyEnd: body.xPt + body.widthPt } : {}),
  };
}

describe('numbered first line beside a square-wrapped float (#1883)', () => {
  // Windows Word 16.0 places the marker at max(max(indent, wrap) + firstLine,
  // wrap) and advances its suffix tab to the paragraph's absolute stops; the
  // fake measurer gives the "1." marker 10pt.
  it.each([
    // left 54 hanging 36: marker at the wrap edge, body at the 234pt default stop.
    { ind: '<w:ind w:left="1080" w:hanging="720"/>', tabs: '', marker: 207, body: 234 },
    // left 36 firstLine 18: the positive first-line indent stays an inset; the
    // marker ends at 235pt, past the 234pt stop.
    { ind: '<w:ind w:left="720" w:firstLine="360"/>', tabs: '', marker: 225, body: 270 },
    // A level w:ind without hanging/firstLine has no first-line indent.
    { ind: '<w:ind w:left="0"/>', tabs: '', marker: 207, body: 234 },
    // The level's own custom stop (216pt from the margin) wins over defaults.
    { ind: '<w:ind w:left="1080" w:hanging="720"/>', tabs: '<w:tabs><w:tab w:val="left" w:pos="4320"/></w:tabs>', marker: 207, body: 270 },
  ])('$ind $tabs', ({ ind, tabs, marker, body }) => {
    expect(firstItemLine(bytes(ind, tabs))).toMatchObject({ marker, body });
  });

  it('lets a paragraph style clear a style-origin level tab', () => {
    // §17.7.2: style-origin numbering sits beneath the paragraph style, whose
    // clear at 216pt removes the level's stop there; the suffix tab then uses
    // the default stops.
    const levelTab = '<w:tabs><w:tab w:val="left" w:pos="4320"/></w:tabs>';
    const ind = '<w:ind w:left="1080" w:hanging="720"/>';
    expect(firstItemLine(bytes(ind, levelTab, '')).body).toBe(270);
    expect(firstItemLine(bytes(ind, levelTab, '<w:tabs><w:tab w:val="clear" w:pos="4320"/></w:tabs>')).body).toBe(234);
  });

  it('places an empty item\'s marker at the moved edge', () => {
    expect(firstItemLine(bytes('<w:ind w:left="1080" w:hanging="720"/>', '', undefined, { empty: true })).marker).toBe(207);
  });

  it('mirrors the suffix tab for an RTL item beside a right float', () => {
    // The 612pt page mirrors the LTR geometry: the body ends where the LTR
    // body (234pt) starts, measured from the right page edge.
    const line = firstItemLine(bytes('<w:ind w:left="1080" w:hanging="720"/>', '', undefined, { rtl: true }));
    expect(line.marker! + 10).toBe(612 - 207);
    expect(line.bodyEnd).toBe(612 - 234);
  });

  it.each([
    // A document-default clear sits beneath style-origin numbering: the level stop survives.
    { defaults: '<w:tabs><w:tab w:val="clear" w:pos="4320"/></w:tabs>', level: '<w:tabs><w:tab w:val="left" w:pos="4320"/></w:tabs>', body: 270 },
    // A level clear removes a document-default stop beneath it.
    { defaults: '<w:tabs><w:tab w:val="left" w:pos="4320"/></w:tabs>', level: '<w:tabs><w:tab w:val="clear" w:pos="4320"/></w:tabs>', body: 234 },
  ])('orders numbering tabs above document defaults ($body)', ({ defaults, level, body }) => {
    const ind = '<w:ind w:left="1080" w:hanging="720"/>';
    expect(firstItemLine(bytes(ind, level, '', { defaultsPPr: defaults })).body).toBe(body);
  });
});
