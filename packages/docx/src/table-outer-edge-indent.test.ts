/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { Canvas } from 'skia-canvas';
import { beforeAll, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const enc = (text: string) => new TextEncoder().encode(text);

const FIXED = `<w:tbl><w:tblPr><w:tblW w:w="2880" w:type="dxa"/><w:tblInd w:w="1440" w:type="dxa"/>`
  + `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
// AutoFit with no table width: the leading-indent band decides the ceiling.
const AUTOFIT = `<w:tbl><w:tblPr><w:tblInd w:w="720" w:type="dxa"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${'word '.repeat(60)}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
// A fixed table whose every row skips the leading grid track.
const SKIPPED = `<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblInd w:w="1440" w:type="dxa"/><w:tblLayout w:type="fixed"/>`
  + `<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="2880"/></w:tblGrid><w:tr><w:trPr><w:gridBefore w:val="1"/><w:wBefore w:w="1440" w:type="dxa"/></w:trPr>`
  + `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

function model(table: string, mode: number | null) {
  const settings = mode === null ? `<w:settings ${W}/>` : `<w:settings ${W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat></w:settings>`;
  const bytes = storeZip(new Map([
    ['[Content_Types].xml', enc('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')],
    ['_rels/.rels', enc('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', enc('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')],
    ['word/settings.xml', enc(settings)],
    ['word/document.xml', enc(`<w:document ${W}><w:body><w:p/>${table}<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`)],
  ]));
  return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(new DocxArchive(bytes).parse()))).document;
}

function tableGeometry(table: string, mode: number | null, outerEdgeIndent: boolean) {
  const document = model(table, mode);
  const source = document.body.find((element) => element.type === 'table') as unknown as { __tableLayout: Record<string, unknown> };
  if (outerEdgeIndent) source.__tableLayout = { ...source.__tableLayout, outerEdgeIndent: true };
  const layout = layoutDocument(document, createLayoutServices(document, {
    measureContext: new Canvas(10, 10).getContext('2d') as unknown as CanvasRenderingContext2D,
  }), { currentDateMs: 0 });
  const node = layout.pages[0]!.layers.body.find((n) => n.kind === 'table')!;
  return { bounds: node.flowBounds, columns: 'columnWidthsPt' in node ? node.columnWidthsPt : null };
}

it('applies WORD_TABLE_ORIGIN_COMPATIBILITY to an OOXML indent but keeps a producer-supplied outer-edge indent', () => {
  // Modes 11/12/14 place an OOXML table's outer edge one left cell margin
  // before the authored indent. A producer-supplied outer-edge indent
  // already is that edge, so the table keeps the mode-absent contract.
  expect(tableGeometry(FIXED, 11, false).bounds.xPt).toBeCloseTo(72 + 72 - 5.4, 6);
  expect(tableGeometry(FIXED, 11, true).bounds.xPt).toBeCloseTo(72 + 72, 6);
});

it.each([['fixed', FIXED], ['AutoFit band', AUTOFIT], ['skipped leading grid', SKIPPED]])(
  'lays a producer-supplied outer-edge %s table out exactly as with no compatibility mode',
  (_label, table) => {
    expect(tableGeometry(table, 11, true)).toEqual(tableGeometry(table, null, true));
  },
);

it.each([['AutoFit band', AUTOFIT], ['skipped leading grid', SKIPPED]])(
  'still applies the mode to an OOXML %s table',
  (_label, table) => {
    expect(tableGeometry(table, 11, false)).not.toEqual(tableGeometry(table, null, false));
  },
);
