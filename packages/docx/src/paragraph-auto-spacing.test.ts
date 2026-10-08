import { readFile } from 'node:fs/promises';
import { beforeAll, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeDocxDocumentModel, bodyLayoutAcquisitionInput } from './parser-model.js';
import type { DocxDocumentModel, DocParagraph } from './types.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});
function document(fixed = false, containers = false): DocxDocumentModel {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const p = (props = '') => `<w:p><w:pPr><w:pStyle w:val="Auto"/>${props}</w:pPr><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>text</w:t></w:r></w:p>`;
  const table = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc>${p()}</w:tc></w:tr></w:tbl>`;
  const box = `<w:p><w:r><w:pict><v:shape id="box" type="#_x0000_t202" style="width:100pt;height:100pt"><v:textbox inset="0,0,0,0"><w:txbxContent>${p()}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;
  const body = (containers ? table + box : '') + p() + p('<w:spacing w:beforeAutospacing="0" w:afterAutospacing="0"/>')
    + p('<w:spacing w:before="800" w:beforeLines="500"/>');
  const files = new Map([
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>${body}<w:sectPr>${containers ? '<w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/>' : ''}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`],
    ['word/styles.xml', `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Auto"><w:pPr><w:spacing w:before="100" w:after="100" w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr></w:style></w:styles>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}"><w:compat>${fixed ? '<w:doNotUseHTMLParagraphAutoSpacing/>' : ''}</w:compat></w:settings>`],
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'],
  ]);
  if (containers) {
    files.set('word/header1.xml', `<w:hdr xmlns:w="${W}">${p()}</w:hdr>`);
    files.set('word/footer1.xml', `<w:ftr xmlns:w="${W}">${p()}</w:ftr>`);
    files.set('word/footnotes.xml', `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${p()}</w:footnote></w:footnotes>`);
    files.set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="footer" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`);
  }
  const O = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const relationships = [['styles', 'styles', 'styles.xml'], ['settings', 'settings', 'settings.xml'],
    ...(containers ? [['header', 'header', 'header1.xml'], ['footer', 'footer', 'footer1.xml'], ['notes', 'footnotes', 'footnotes.xml']] : [])];
  files.set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map(([id, type, target]) => `<Relationship Id="${id}" Type="${O}/${type}" Target="${target}"/>`).join('')}</Relationships>`);
  files.set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`);
  files.set('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, xml]) => [name, new TextEncoder().encode(xml)]))));
  try { return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel; }
  finally { archive.free(); }
}

it('uses an HTML paragraph em for inherited automatic margins while explicit false retains authored spacing', () => {
  const raw = document();
  const normalized = normalizeDocxDocumentModel(raw);
  const paragraphs = normalized.body.filter((e): e is DocParagraph & { type: 'paragraph' } => e.type === 'paragraph');
  expect(paragraphs.map(p => [p.spaceBefore, p.spaceAfter])).toEqual([[11, 11], [5, 5], [11, 11]]);
  // Projection is immutable and the pagination sequence uses the same resolved pair.
  expect((raw.body[0] as DocParagraph).spaceBefore).toBe(5);
  expect(normalizeDocxDocumentModel(normalized).body).toEqual(normalized.body);
  const sequence = bodyLayoutAcquisitionInput(normalized).sequence;
  expect(sequence.filter(e => e.kind === 'body-block').map(e => e.kind === 'body-block' && e.block.kind === 'paragraph' ? [e.block.spaceBeforePt, e.block.spaceAfterPt] : null)).toEqual([[11, 11], [5, 5], [11, 11]]);
});

it('retains stored margins under Word fixed automatic-spacing compatibility', () => {
  const normalized = normalizeDocxDocumentModel(document(true));
  const paragraphs = normalized.body.filter((e): e is DocParagraph & { type: 'paragraph' } => e.type === 'paragraph');
  expect(paragraphs.map(p => [p.spaceBefore, p.spaceAfter])).toEqual([[5, 5], [5, 5], [40, 5]]);
});

it.each([false, true])('propagates document fixed=%s into inherited cell and story margins', (fixed) => {
  const raw = document(fixed, true);
  const normalized = normalizeDocxDocumentModel(raw);
  const table = normalized.body.find(e => e.type === 'table');
  const cell = table?.type === 'table' ? table.rows[0]?.cells[0]?.content[0] : null;
  const paragraphs = [cell, normalized.headers.default?.body[0], normalized.footers.default?.body[0], normalized.footnotes?.[0]?.content?.[0]];
  for (const paragraph of paragraphs) {
    expect(paragraph?.type).toBe('paragraph');
    if (paragraph?.type === 'paragraph') expect([paragraph.spaceBefore, paragraph.spaceAfter]).toEqual(fixed ? [5, 5] : [11, 11]);
  }
  expect((raw.headers.default?.body[0] as DocParagraph).spaceBefore).toBe(5);
});

it('keeps parsed compatibility text-box margins independent of the first inline font', () => {
  const raw = document(false, true);
  const normalized = normalizeDocxDocumentModel(raw);
  const shapes = normalized.body.flatMap(e => e.type === 'paragraph' ? e.runs.filter(run => run.type === 'shape') : []);
  expect(shapes).toHaveLength(1);
  const block = shapes[0]?.type === 'shape' ? shapes[0].textBlocks?.[0] : undefined;
  expect(block?.fontSizePt).toBe(24);
  expect([block?.spaceBefore, block?.spaceAfter]).toEqual([11, 11]);
  expect(normalizeDocxDocumentModel(normalized).body).toEqual(normalized.body);
});
