import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { textRunsForPage } from './text-run-projection.js';
import type { DocxTextRunInfo } from './renderer.js';
import type { DocxDocumentModel } from './types.js';

// A text box nested in a page story (here a header) is acquired with ITS
// owner's section and destination page, not with the body's current location.
// The PAGE field inside the header text box is the independent observable:
// every page's header shows its own page number. The native variant feeds the
// same parsed model with a canonical native BtoT section (the producer's
// private raw flow 2 on its nominal btLr token), so the body runs in its
// counter-clockwise frame while the header story, and the text box inside it,
// stay horizontal and upright; a body text box exercises the upright-physical
// nested story on that native page.

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const V = 'urn:schemas-microsoft-com:vml';
const PAGE = { width: 612, height: 792 };

function textBox(content: string): string {
  return `<w:r><w:pict><v:shape id="box" type="#_x0000_t202" style="width:72pt;height:36pt">`
    + `<v:textbox inset="0,0,0,0"><w:txbxContent><w:p>${content}</w:p></w:txbxContent></v:textbox>`
    + '</v:shape></w:pict></w:r>';
}

function parsedDocument(): DocxDocumentModel {
  const body = `<w:p><w:r><w:t>One</w:t></w:r>${textBox('<w:r><w:t>Box</w:t></w:r>')}</w:p>`
    + '<w:p><w:r><w:br w:type="page"/><w:t>Two</w:t></w:r></w:p>';
  const header = `<w:p>${textBox('<w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple>')}</w:p>`;
  const files = new Map<string, string>([
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:v="${V}"><w:body>${body}`
      + '<w:sectPr><w:headerReference w:type="default" r:id="hdr"/><w:pgSz w:w="12240" w:h="15840"/>'
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
      + '</w:sectPr></w:body></w:document>'],
    ['word/header1.xml', `<w:hdr xmlns:w="${W}" xmlns:v="${V}">${header}</w:hdr>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="hdr" Type="${R}/header" Target="header1.xml"/></Relationships>`],
    ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
      + '</Types>'],
  ]);
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, xml]) => [name, new TextEncoder().encode(xml)]))));
  try { return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel; }
  finally { archive.free(); }
}

/** The canonical native producer facts of an MS-ODRAW BtoT section: nominal
 * btLr plus the private raw flow, exactly as the native reader emits them. */
function asNativeBottomToTop(doc: DocxDocumentModel): DocxDocumentModel {
  const section = doc.section as DocxDocumentModel['section'] & { __sectionPlacement?: Record<string, unknown> };
  section.textDirection = 'btLr';
  section.__sectionPlacement = { ...(section.__sectionPlacement ?? { sectionId: 'section:0' }), nativeTextFlow: 2 };
  return doc;
}

function context(): CanvasRenderingContext2D {
  const ctx = { font: '12px serif', fontKerning: 'auto', letterSpacing: '0px',
    measureText(text: string) {
      const size = Number(/([\d.]+)px/u.exec(ctx.font)?.[1] ?? 12);
      const width = [...text].length * size * 0.6;
      return { width, fontBoundingBoxAscent: size * .8, fontBoundingBoxDescent: size * .2,
        actualBoundingBoxLeft: 0, actualBoundingBoxRight: width,
        actualBoundingBoxAscent: size * .8, actualBoundingBoxDescent: size * .2 } as TextMetrics;
    } };
  return ctx as unknown as CanvasRenderingContext2D;
}

/** Physical box of an overlay run (x/y are its physical top-left). */
function physicalBox(run: DocxTextRunInfo) {
  if (run.transform === undefined) return { left: run.x, right: run.x + run.w, top: run.y, bottom: run.y + run.h };
  if (run.transform === 'rotate(90deg)') return { left: run.x - run.h, right: run.x, top: run.y, bottom: run.y + run.w };
  if (run.transform === 'rotate(-90deg)') return { left: run.x, right: run.x + run.h, top: run.y - run.w, bottom: run.y };
  throw new Error(`unexpected overlay transform ${run.transform}`);
}

describe('nested text-box story ownership', () => {
  it.each([
    { label: 'horizontal', native: false, bodyTransform: undefined },
    { label: 'native BtoT', native: true, bodyTransform: 'rotate(-90deg)' },
  ])('acquires a header text box with its own page and section ($label)', ({ native, bodyTransform }) => {
    const parsed = parsedDocument();
    const doc = native ? asNativeBottomToTop(parsed) : parsed;
    const layout = layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(2);
    for (const [pageIndex, bodyText] of ['One', 'Two'].entries()) {
      const runs = textRunsForPage(layout, pageIndex, { scale: 1 });
      const body = runs.filter(run => run.source?.story === 'body');
      const boxes = runs.filter(run => run.source?.story === 'textbox');
      expect(body.map(run => run.text).join(''), `page ${pageIndex}`).toBe(bodyText);
      for (const run of body) expect(run.transform).toBe(bodyTransform);
      // The header's own text box shows THIS page's number.
      const headerBox = boxes.filter(run => /^\d+$/u.test(run.text));
      expect(headerBox.map(run => run.text).join(''), `page ${pageIndex}`).toBe(String(pageIndex + 1));
      for (const run of headerBox) {
        // Horizontal and upright in the physical header band.
        expect(run.transform).toBeUndefined();
        expect(physicalBox(run).bottom).toBeLessThan(PAGE.height / 2);
      }
      // The body text box belongs to the first page only and stays on it.
      const bodyBox = boxes.filter(run => !/^\d+$/u.test(run.text));
      expect(bodyBox.map(run => run.text).join(''), `page ${pageIndex}`).toBe(pageIndex === 0 ? 'Box' : '');
      for (const run of bodyBox) {
        const box = physicalBox(run);
        expect(box.left >= 0 && box.top >= 0 && box.right <= PAGE.width && box.bottom <= PAGE.height).toBe(true);
      }
    }
  });
});
