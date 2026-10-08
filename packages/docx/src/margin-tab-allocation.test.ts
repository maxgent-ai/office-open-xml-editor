/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { measureContext } from './test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

function docxBytes(bodyXml: string, sectionXml: string, footerXml?: string): Uint8Array {
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${bodyXml}<w:sectPr>`
    + (footerXml ? '<w:footerReference w:type="default" r:id="rIdFooter"/>' : '')
    + `${sectionXml}</w:sectPr></w:body></w:document>`;
  const parts = new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + (footerXml ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : '')
      + '</Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode(`<Relationships xmlns="${PKG}">`
      + (footerXml ? '<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' : '')
      + '</Relationships>')],
    ['word/document.xml', encode(document)],
  ]);
  if (footerXml) parts.set('word/footer1.xml', encode(footerXml));
  return storeZip(parts);
}

function layout(bytes: Uint8Array, advancePt: number) {
  const archive = new DocxArchive(bytes);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  return layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(advancePt) }), { currentDateMs: 0 });
}

describe('footer right tab past the paragraph right indent', () => {
  // Synthetic equivalent of a private journal footer: an A4 page with 936
  // twip margins (501.7 pt text width), a right indent of 8108 twips (405.4
  // pt, leaving a 96.3 pt indent band) and an ordinary right stop at 9923
  // twips (496.15 pt), past the indent but inside the text margin.
  // ECMA-376 §17.3.1.37 positions custom stops relative to the page margins;
  // the right indent does not bound them. Word keeps such a footer's aligned
  // cell on the first line; containing it in the indent band instead grew
  // the footer from two lines to five and pushed body text to later pages.
  it('aligns the cell at its margin-relative stop and keeps two lines', () => {
    const footer = `<w:ftr xmlns:w="${W}"><w:p><w:pPr>`
      + '<w:tabs><w:tab w:val="right" w:pos="9923"/></w:tabs><w:ind w:right="8108"/>'
      + '<w:rPr><w:sz w:val="16"/></w:rPr></w:pPr>'
      + '<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>DOI: 10.2478/x</w:t></w:r>'
      + '<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:tab/></w:r>'
      + '<w:r><w:rPr><w:i/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve"> *Corresponding author: a@b.edu (F. Surname) </w:t></w:r>'
      + '<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:tab/></w:r>'
      + '</w:p></w:ftr>';
    const result = layout(docxBytes('<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
      '<w:pgSz w:w="11906" w:h="16838"/>'
      + '<w:pgMar w:top="1418" w:right="936" w:bottom="1418" w:left="936" w:header="431" w:footer="431" w:gutter="0"/>',
      footer), 4);
    const paragraph = result.pages[0]?.layers.footer.find((node) => node.kind === 'paragraph');
    if (paragraph?.kind !== 'paragraph') throw new Error('Missing footer paragraph');
    // Line 1: DOI, the right tab and its cell; line 2: the trailing tab.
    expect(paragraph.lines).toHaveLength(2);
    const texts = paragraph.lines[0].placements.filter((node) => node.kind === 'text');
    expect(texts.map((text) => text.text).join('').replace(/\s+/g, ' ').trim())
      .toBe('DOI: 10.2478/x *Corresponding author: a@b.edu (F. Surname)');
    const cell = texts.filter((text) => text.text.trim().length > 0).at(-1)!;
    // The stop is at 46.8 + 496.15 pt; the right-aligned cell (with its 4 pt
    // trailing space) ends there, outside the 46.8 + 96.3 pt indent band.
    expect(cell.text.endsWith(') ')).toBe(true);
    expect(cell.bounds.xPt + cell.bounds.widthPt).toBeCloseTo(46.8 + 496.15, 6);
  });
});

describe('Word body control: tabs beside a 360 pt right indent', () => {
  // Synthetic equivalent of a Word 16 export: Letter page, 72 pt margins,
  // right indent 7200 twips (indent band 72–180 pt), Courier New 10 pt "12.3"
  // (6 pt advances), with and without paragraph borders and shading.
  const run = (tab: string) => '<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="20"/></w:rPr>'
    + `${tab}<w:t>12.3</w:t></w:r>`;
  const decorations = '<w:pBdr><w:top w:val="single" w:sz="8" w:space="0"/><w:left w:val="single" w:sz="8" w:space="0"/>'
    + '<w:bottom w:val="single" w:sz="8" w:space="0"/><w:right w:val="single" w:sz="8" w:space="0"/></w:pBdr>'
    + '<w:shd w:val="clear" w:fill="FFFF00"/>';
  const cellBounds = (decorated: boolean, tabs: string, tab: string) => {
    const result = layout(docxBytes(
      `<w:p><w:pPr>${decorated ? decorations : ''}${tabs}<w:ind w:right="7200"/></w:pPr>${run(tab)}</w:p>`,
      '<w:pgSz w:w="12240" w:h="15840"/>'
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>',
    ), 6);
    const paragraph = result.pages[0]?.layers.body.find((node) => node.kind === 'paragraph');
    if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
    expect(paragraph.lines).toHaveLength(1);
    const texts = paragraph.lines[0].placements.filter((node) => node.kind === 'text');
    expect(texts.map((text) => text.text).join('')).toBe('12.3');
    return [texts[0].bounds.xPt, texts.at(-1)!.bounds.xPt + texts.at(-1)!.bounds.widthPt];
  };

  // ECMA-376 §17.3.3.23 selects the margin target independently of indents,
  // but Word ends the cell at the right indent (x 156–180 pt in its PDF).
  // Expected failure: positional tabs keep the #1675 containment policy, which
  // collapses the unreachable gap (x 72–96 pt). This difference predates the
  // ordinary-stop fix; one control does not establish a general rule, so a
  // follow-up needs a varied Word control set (indent and cell sizes, targets
  // inside versus past the band, center/right, tables, headers/footers, RTL).
  it.fails.each(['right', 'center'].flatMap((alignment) => [false, true].map((decorated) => ({ alignment, decorated }))))(
    'ends a margin $alignment ptab at the right indent, decorated=$decorated', ({ alignment, decorated }) => {
      expect(cellBounds(decorated, '', `<w:ptab w:alignment="${alignment}" w:relativeTo="margin" w:leader="none"/>`))
        .toEqual([156, 180]);
    });

  it.each([false, true])('ends an indent right ptab at the right indent, decorated=%s', (decorated) => {
    expect(cellBounds(decorated, '', '<w:ptab w:alignment="right" w:relativeTo="indent" w:leader="none"/>'))
      .toEqual([156, 180]);
  });

  // Ordinary stops at 7200 twips (432 pt) extend past the indent band in the
  // same export: right x 408–432, center and decimal x 420–444.
  it.each([
    { alignment: 'right', expected: [408, 432] },
    { alignment: 'center', expected: [420, 444] },
    { alignment: 'decimal', expected: [420, 444] },
  ].flatMap((entry) => [false, true].map((decorated) => ({ ...entry, decorated }))))(
    'aligns an ordinary $alignment stop past the right indent, decorated=$decorated', ({ alignment, expected, decorated }) => {
      expect(cellBounds(decorated, `<w:tabs><w:tab w:val="${alignment}" w:pos="7200"/></w:tabs>`, '<w:tab/>'))
        .toEqual(expected);
    });
});
