/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

it('lays the split paragraph mark at the next page top and displaces a near-page-height row', () => {
  const table = (text: string) => `<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr>
    <w:trPr><w:cantSplit/><w:trHeight w:val="12760" w:hRule="exact"/></w:trPr>
    <w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>
    </w:tr></w:tbl>`;
  for (const enabled of [false, true]) {
    const parts = new Map([
      ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'],
      ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${table('first')}
        <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="exact"/>
          <w:ind w:left="240"/><w:rPr><w:sz w:val="36"/></w:rPr></w:pPr>
          <w:r><w:br w:type="page"/></w:r></w:p>${table('second')}
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/>
        </w:sectPr></w:body></w:document>`],
      ['word/settings.xml', `<w:settings xmlns:w="${W}"><w:compat>
        <w:splitPgBreakAndParaMark w:val="${enabled ? 1 : 0}"/>
      </w:compat></w:settings>`],
    ]);
    const encoder = new TextEncoder();
    const archive = new DocxArchive(storeZip(new Map([...parts].map(([path, xml]) => [path, encoder.encode(xml)]))));
    try {
      const model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
      const context = {
        font: '10px serif', letterSpacing: '0px', fontKerning: 'none',
        measureText: (text: string) => ({ width: text.length * 6,
          fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 }),
      } as unknown as CanvasRenderingContext2D;
      const layout = layoutDocument(model, createLayoutServices(model, { measureContext: context }), { currentDateMs: 0 });
      expect(layout.pages).toHaveLength(enabled ? 3 : 2);
      expect(layout.pages[0]?.layers.body.map(block => block.kind)).toEqual(['table']);
      if (enabled) {
        const blocks = layout.pages[1]?.layers.body;
        expect(blocks?.map(block => block.kind)).toEqual(['paragraph']);
        expect(blocks?.[0]?.flowBounds).toMatchObject({ yPt: 72, heightPt: 18 });
      }
      expect(layout.pages[enabled ? 2 : 1]?.layers.body.map(block => block.kind)).toEqual(['table']);
    } finally { archive.free(); }
  }
});
