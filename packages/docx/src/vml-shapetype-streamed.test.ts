/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';

function docx(body: string): Uint8Array {
  const enc = new TextEncoder();
  return storeZip(new Map([
    ['[Content_Types].xml', enc.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', enc.encode(`<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`)],
  ]));
}

const shape = (id: string, text: string) =>
  `<v:shape id="${id}" type="#_x0000_t136" style="width:144pt;height:36pt" fillcolor="red" stroked="f"><v:textpath string="${text}"/></v:shape>`;

describe('VML shape types across streamed body blocks (Part 4 §14.1.2.20)', () => {
  it('resolves a shape type defined in an earlier paragraph', () => {
    const shapeType = '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" path="m@7,l@8,m@5,21600l@6,21600e"><v:path textpathok="t"/><v:textpath on="t" fitshape="t"/></v:shapetype>';
    const bytes = docx(
      `<w:p><w:r><w:pict>${shapeType}${shape('a', 'ONE')}</w:pict></w:r></w:p>`
      + `<w:p><w:r><w:pict>${shape('b', 'TWO')}</w:pict></w:r></w:p>`,
    );
    const model = JSON.parse(new TextDecoder().decode(new DocxArchive(bytes).parse()));
    const shapes: Array<{ textPath?: { string: string; on?: boolean }; inline?: boolean }> = [];
    JSON.stringify(model, (_key, value) => {
      if (value && value.type === 'shape') shapes.push(value);
      return value;
    });
    expect(shapes.map((s) => [s.textPath?.string, s.textPath?.on, s.inline])).toEqual([
      ['ONE', true, true], ['TWO', true, true],
    ]);
  });
});
