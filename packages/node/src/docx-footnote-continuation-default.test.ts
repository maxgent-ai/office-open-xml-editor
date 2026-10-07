import { Canvas, loadImage } from 'skia-canvas';
import { describe, expect, it } from 'vitest';
import type { ModelSource } from '@silurus/ooxml-core';
import { openDocxDocument, type DocxDocumentSession } from './docx.ts';
import type { NodeCanvasFactory } from './render.ts';
import { storedZip } from './test-ooxml-package.ts';

// Node counterpart of the public footnote continuation default: an omitted
// `allowFootnoteContinuation` continues notes like explicit true, and explicit
// false keeps whole notes, through both public `openDocxDocument` branches:
// - an authored DOCX package through the parser and production layout;
// - a claimed model source serving a self-authored model (no parser WASM).
// Both inputs are the same exact-line arithmetic: five 10pt filler lines, B01
// opened by reference 1 with text after it, a four-line note, 80pt body, no
// separator story (6pt band). Continued: 60 + 6 + one note line = 76 keeps
// note-0; beside B01. Whole: 60 + 46 > 80, so B01 and the note go to page 2.

const factory: NodeCanvasFactory = {
  createCanvas: (width, height) =>
    new Canvas(width, height) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async (buffer: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.from(buffer as Uint8Array))) as unknown as NodeCanvasFactory['loadImage'],
};

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006';
const paragraph = (inner: string) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/></w:pPr>${inner}</w:p>`;
const text = (value: string) => `<w:r><w:t>${value}</w:t></w:r>`;

/** Authored DOCX: 200x100pt page, 10pt margins, exact 10pt lines. */
const authoredDocx = () => storedZip({
  '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${PKG}/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>`,
  '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG}/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG}/relationships"><Relationship Id="rId1" Type="${R}/footnotes" Target="footnotes.xml"/></Relationships>`,
  'word/document.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body>${
  Array.from({ length: 5 }, (_, index) => paragraph(text(`F${index}`))).join('')
}${paragraph(`<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>${text('B01')}`)
}<w:sectPr><w:pgSz w:w="4000" w:h="2000"/><w:pgMar w:top="200" w:right="200" w:bottom="200" w:left="200" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  'word/footnotes.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${
  Array.from({ length: 4 }, (_, index) => paragraph(text(`note-${index};`))).join('')
}</w:footnote></w:footnotes>`,
});

const materializedSource: ModelSource<'docx'> = {
  target: 'docx',
  claim: () => true,
  beginLoad: () => ({
    module: {
      protocol: 'ooxml-model-source-module/v1',
      target: 'docx',
      moduleUrl: new URL('./test-fixtures/docx-materialized-model-source.mjs', import.meta.url).href,
      config: {},
    },
    release: () => undefined,
  }),
};

async function pageTexts(session: DocxDocumentSession): Promise<string[]> {
  const pages: string[] = [];
  for (let pageIndex = 0; pageIndex < session.pageCount; pageIndex += 1) {
    const runs: string[] = [];
    await session.renderPage(pageIndex, { onTextRun: run => runs.push(run.text) });
    pages.push(runs.join(''));
  }
  return pages;
}

const CONTINUED = ['F0F1F2F3F41B01note-0;', 'note-1;note-2;note-3;'];
const WHOLE = ['F0F1F2F3F4', '1B01note-0;note-1;note-2;note-3;'];

describe('openDocxDocument footnote continuation default', () => {
  it.each([
    { branch: 'authored DOCX', option: undefined, expected: CONTINUED },
    { branch: 'authored DOCX', option: true, expected: CONTINUED },
    { branch: 'authored DOCX', option: false, expected: WHOLE },
    { branch: 'claimed model source', option: undefined, expected: CONTINUED },
    { branch: 'claimed model source', option: true, expected: CONTINUED },
    { branch: 'claimed model source', option: false, expected: WHOLE },
  ])('lays out footnote continuation $option for an $branch', async ({ branch, option, expected }) => {
    const session = await openDocxDocument(
      branch === 'authored DOCX' ? authoredDocx() : new Uint8Array([1, 2, 3, 4]),
      {
        factory, currentDate: 0,
        ...(branch === 'authored DOCX' ? {} : { modelSources: [materializedSource] }),
        ...(option === undefined ? {} : { allowFootnoteContinuation: option }),
      },
    );
    try {
      expect(await pageTexts(session)).toEqual(expected);
    } finally {
      await session.close();
    }
  });
});
