/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import { storeZip } from './conformance/generate.js';
import type { DocumentLayout } from './layout/types.js';
import type { DocxDocumentModel } from './types.js';

// Owner-context eligibility of a cell-owner row carrier (a leading first-cell
// framePr paragraph of a table whose effective §17.4.57 positioning is null),
// reached through the REAL parser from minimal synthetic OOXML built here.
//
// Bounded Word observation (public synthetic controls, original-open and
// saved/reopened full pages exactly equal; no tblpPr added on save). Neither
// ECMA-376 §17.3.1.11 nor §17.4.57, nor MS-OI29500 2.1.43, defines row
// promotion; these cases record only the tested classes:
//   - Word keeps framePr, and renders differently from the frame-free control,
//     for one- and two-cell tables that are DIRECT roots of the body, header
//     or footer story (page/page notBeside; body also width-omitted
//     margin/text around).
//   - Word removes framePr on save, and renders exactly the frame-free
//     control, for one- and two-cell tables NESTED in a body or footer cell,
//     that are roots of a TEXT BOX story (unrotated, 90°, ±30°), or roots of a
//     FOOTNOTE or ENDNOTE story (page/page notBeside and width-omitted
//     margin/text around).
// So the cell count does not decide; the owning structural context does. In
// the inert contexts the parsed framePr remains a retained source fact and
// the paragraph remains ordinary cell content: layout must equal the
// frame-free control. Synthetic linear metrics (half-em advances) are a test
// measurement service only; expected geometry comes from authored operands
// (page 400×400pt, 40pt margins, header/footer distance 20pt, exact 24pt
// lines, 10pt top/bottom cell margins) or from the frame-free control, never
// from the implementation.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const EMU = 12700;
const TW = 20;

const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);

const exact24 = `<w:spacing w:before="0" w:after="0" w:line="${24 * TW}" w:lineRule="exact"/>`;
// Property order follows the Transitional WML schema sequences: CT_PPrBase
// places framePr before spacing; CT_TblPrBase places tblW, tblBorders,
// tblLayout, tblCellMar in that order.
const para = (text: string, pPr = '') => `<w:p><w:pPr>${pPr}${exact24}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

const FRAME_X = 100;
const GRID_PT = 100;
const CELL_TOP_PT = 10;

/** `page`: page/page notBeside at (100, y) with w = the 100pt grid.
 * `around`: width-omitted margin/text around, 100pt right of the margin. */
type FrameKind = 'page' | 'around';

function framePr(kind: FrameKind, yPt: number): string {
  return kind === 'page'
    ? `<w:framePr w:w="${GRID_PT * TW}" w:wrap="notBeside" w:hAnchor="page" w:vAnchor="page" `
      + `w:x="${FRAME_X * TW}" w:y="${yPt * TW}"/>`
    : `<w:framePr w:wrap="around" w:hAnchor="margin" w:vAnchor="text" w:x="${FRAME_X * TW}" w:y="0"/>`;
}

/** Fixed-layout table over `widthsPt`, one row of `cells`, top/bottom
 * margins `marginPt`, left/right 0, optional 1pt single rules. */
function table(cells: readonly string[], widthsPt: readonly number[], marginPt: number, ruled: boolean): string {
  const rule = (side: string) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`;
  const borders = ruled
    ? `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideV'].map(rule).join('')}</w:tblBorders>`
    : '';
  const total = widthsPt.reduce((sum, width) => sum + width, 0);
  return `<w:tbl><w:tblPr><w:tblW w:w="${total * TW}" w:type="dxa"/>${borders}<w:tblLayout w:type="fixed"/>`
    + `<w:tblCellMar><w:top w:w="${marginPt * TW}" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>`
    + `<w:bottom w:w="${marginPt * TW}" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>`
    + `<w:tblGrid>${widthsPt.map((width) => `<w:gridCol w:w="${width * TW}"/>`).join('')}</w:tblGrid><w:tr>`
    + cells.map((content, index) =>
      `<w:tc><w:tcPr><w:tcW w:w="${widthsPt[index]! * TW}" w:type="dxa"/></w:tcPr>${content}</w:tc>`).join('')
    + '</w:tr></w:tbl>';
}

/** The tested host table: HOST leads cell A (carrying the frame unless
 * `frame` is null); a two-cell table adds TWO in cell B over a 50/50 grid. */
function hostTable(
  cellCount: 1 | 2,
  frame: Readonly<{ kind: FrameKind; yPt: number }> | null,
  ruled = false,
): string {
  const host = para('HOST', frame ? framePr(frame.kind, frame.yPt) : '');
  return cellCount === 1
    ? table([host], [GRID_PT], CELL_TOP_PT, ruled)
    : table([host, para('TWO')], [GRID_PT / 2, GRID_PT / 2], CELL_TOP_PT, ruled);
}

/** A 320pt borderless margin-free outer cell holding PARENT, `inner`, and
 * PARENT-AFTER (the r2 containment topology). */
const nestedIn = (inner: string) => table([para('PARENT') + inner + para('PARENT-AFTER')], [320], 0, false);

type Stories = Readonly<{
  body: string;
  header?: string;
  footer?: string;
  footnote?: string;
  endnote?: string;
  /** §17.6.20 section `<w:textDirection>`; absent is horizontal. */
  textDirection?: 'tbRl';
}>;

function docx(stories: Stories): Uint8Array {
  const parts = [
    ['header', 'header1.xml', 'hdr', stories.header],
    ['footer', 'footer1.xml', 'ftr', stories.footer],
    ['footnotes', 'footnotes.xml', 'footnotes', stories.footnote],
    ['endnotes', 'endnotes.xml', 'endnotes', stories.endnote],
  ] as const;
  const present = parts.filter(([, , , content]) => content !== undefined);
  const references = (stories.header !== undefined ? '<w:headerReference w:type="default" r:id="rIdheader"/>' : '')
    + (stories.footer !== undefined ? '<w:footerReference w:type="default" r:id="rIdfooter"/>' : '');
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}">`
    + `<w:body>${stories.body}<w:sectPr>${references}<w:pgSz w:w="${400 * TW}" w:h="${400 * TW}"/>`
    + `<w:pgMar w:top="${40 * TW}" w:right="${40 * TW}" w:bottom="${40 * TW}" w:left="${40 * TW}" `
    + `w:header="${20 * TW}" w:footer="${20 * TW}" w:gutter="0"/>`
    + (stories.textDirection ? `<w:textDirection w:val="${stories.textDirection}"/>` : '')
    + '</w:sectPr></w:body></w:document>';
  const styles = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>`
    + '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>';
  const files = new Map<string, Uint8Array>([
    ['[Content_Types].xml', xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
      + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + `<Override PartName="/word/document.xml" ContentType="${CONTENT_TYPE}.document.main+xml"/>`
      + `<Override PartName="/word/styles.xml" ContentType="${CONTENT_TYPE}.styles+xml"/>`
      + present.map(([type, file]) =>
        `<Override PartName="/word/${file}" ContentType="${CONTENT_TYPE}.${type}+xml"/>`).join('')
      + '</Types>')],
    ['_rels/.rels', xml(`<Relationships xmlns="${PACKAGE_RELS}">`
      + `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', xml(`<Relationships xmlns="${PACKAGE_RELS}">`
      + `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>`
      + present.map(([type, file]) => `<Relationship Id="rId${type}" Type="${R}/${type}" Target="${file}"/>`).join('')
      + '</Relationships>')],
    ['word/document.xml', xml(document)],
    ['word/styles.xml', xml(styles)],
  ]);
  for (const [type, file, root, content] of present) {
    const inner = type === 'footnotes' ? `<w:footnote w:id="1">${content}</w:footnote>`
      : type === 'endnotes' ? `<w:endnote w:id="1">${content}</w:endnote>`
        : content;
    const element = type === 'footnotes' || type === 'endnotes' ? `w:${type}` : `w:${root}`;
    files.set(`word/${file}`, xml(`<${element} xmlns:w="${W}">${inner}</${element}>`));
  }
  return storeZip(files);
}

const noteReference = (kind: 'footnote' | 'endnote') =>
  `<w:p><w:pPr>${exact24}</w:pPr><w:r><w:t>REF</w:t></w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`;

/** A body paragraph anchoring a page-positioned (50, 100) 300×250 wrap-none
 * text box, turned `degrees`, whose story (`vert`) holds `content`. */
function textBox(degrees: number, vert: 'horz' | 'wordArtVert', content: string): string {
  const cx = 300 * EMU;
  const cy = 250 * EMU;
  return '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
    + 'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
    + `<wp:positionH relativeFrom="page"><wp:posOffset>${50 * EMU}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="page"><wp:posOffset>${100 * EMU}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`
    + `<wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/>`
    + `<wps:spPr><a:xfrm rot="${degrees * 60000}"><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
    + `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>`
    + `<wps:bodyPr rot="0" vert="${vert}" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t">`
    + '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>';
}

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * px() * 0.5,
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

type Parsed = Readonly<{ model: DocxDocumentModel; result: DocumentLayout }>;

function layout(stories: Stories): Parsed {
  const archive = new DocxArchive(docx(stories));
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
  const before = JSON.stringify(model);
  const result = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
    currentDateMs: 0,
  });
  // Layout never rewrites the immutable source model.
  expect(JSON.stringify(model)).toBe(before);
  return { model, result };
}

/** Every parsed framePr declaration, wherever the model retains it. */
function frameDeclarations(value: unknown): unknown[] {
  if (Array.isArray(value)) return value.flatMap(frameDeclarations);
  if (value === null || typeof value !== 'object') return [];
  return Object.entries(value).flatMap(([key, child]) =>
    key === 'framePr' ? (child == null ? [] : [child]) : frameDeclarations(child));
}

/** Everything a reader observes: pages, every placed run with its source and
 * page transform, and diagnostics. */
function observed(result: DocumentLayout) {
  return {
    pageCount: result.pages.length,
    runs: result.pages.flatMap((_, pageIndex) => textRunGeometryForPage(result, pageIndex).map((run) => ({
      pageIndex,
      text: run.placement.text,
      source: run.source,
      bounds: run.placement.bounds,
      pointToPage: run.pointToPage,
    }))),
    diagnostics: result.diagnostics,
  };
}

/** Page position of the top-left of a run placed once in the document. */
function placedOnce(result: DocumentLayout, text: string) {
  const matches = result.pages.flatMap((_, pageIndex) => textRunGeometryForPage(result, pageIndex))
    .filter((run) => run.placement.text === text);
  expect(matches, text).toHaveLength(1);
  const { pointToPage: m, placement: { bounds } } = matches[0]!;
  return { x: m.a * bounds.xPt + m.c * bounds.yPt + m.e, y: m.b * bounds.xPt + m.d * bounds.yPt + m.f };
}

/** Inert contract: the framed input retains its parsed framePr, places HOST
 * exactly once, and is observably identical to the frame-free control —
 * pages, every run's source and page transform, and diagnostics (so no
 * unsupported-host warning either). */
function expectInert(framed: Stories, plain: Stories) {
  const withFrame = layout(framed);
  const control = layout(plain);
  expect(frameDeclarations(control.model)).toEqual([]);
  expect(frameDeclarations(withFrame.model)).toHaveLength(1);
  placedOnce(withFrame.result, 'HOST');
  expect(withFrame.result.diagnostics.filter((entry) => entry.severity === 'error')).toEqual([]);
  expect(observed(withFrame.result)).toEqual(observed(control.result));
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const CELL_COUNTS = [1, 2] as const;

describe('cell-owner carriers that Word keeps: direct body, header and footer roots', () => {
  // Defect caught: an eligibility rule that also drops these (e.g. one keyed
  // on cell count, page/page anchoring, or a story-kind blanket for stories
  // other than the body) — Word keeps and honours the frame here.
  const hostAt = (yPt: number) => ({ x: FRAME_X, y: yPt + CELL_TOP_PT });

  // One- and two-cell: the counterexample to a one-cell rule.
  for (const cellCount of CELL_COUNTS) {
    it(`places a direct ${cellCount}-cell body-root host at its page frame`, () => {
      const stories = (framed: boolean) => ({
        body: para('BEFORE') + hostTable(cellCount, framed ? { kind: 'page', yPt: 200 } : null) + para('AFTER'),
      });
      const { model, result } = layout(stories(true));
      expect(frameDeclarations(model)).toHaveLength(1);
      expect(placedOnce(result, 'HOST')).toEqual(hostAt(200));
      if (cellCount === 2) expect(placedOnce(result, 'TWO')).toEqual({ x: FRAME_X + GRID_PT / 2, y: 200 + CELL_TOP_PT });
      expect(observed(result).runs).not.toEqual(observed(layout(stories(false)).result).runs);
    });
  }

  // The counterexample to a page/page or notBeside rule on the direct side.
  it('honours a width-omitted margin/text around frame on a direct two-cell body root', () => {
    const stories = (framed: boolean) => ({
      body: para('BEFORE') + hostTable(2, framed ? { kind: 'around', yPt: 0 } : null) + para('AFTER'),
    });
    const { result } = layout(stories(true));
    // Horizontal operand only: 40pt margin + 100pt; the plain control's
    // HOST starts at the 40pt margin.
    expect(placedOnce(result, 'HOST').x).toBe(40 + FRAME_X);
    expect(placedOnce(layout(stories(false)).result, 'HOST').x).toBe(40);
  });

  // Header and footer roots are wired apart from the body (story layout, and
  // the footer's extent-dependent band).
  for (const story of ['header', 'footer'] as const) {
    it(`places a direct ${story}-root host at its page frame`, () => {
      const stories = (framed: boolean) => ({
        body: para('BODY'),
        [story]: hostTable(1, framed ? { kind: 'page', yPt: 200 } : null) + para('STORY-AFTER'),
      });
      const { model, result } = layout(stories(true));
      expect(frameDeclarations(model)).toHaveLength(1);
      expect(placedOnce(result, 'HOST')).toEqual(hostAt(200));
      expect(observed(result).runs).not.toEqual(observed(layout(stories(false)).result).runs);
    });
  }
});

// A direct story-root host (Word keeps it, above) aligned on the page by
// yAlign (§17.3.1.11: the frame relative to its vAnchor band, here the page
// [0, 400]), so its position follows from its extent. Its cell B holds a
// §17.4.57 page-positioned child table and ANCH, the child's anchor paragraph,
// which wraps around the child. Whether ANCH's first line meets the child
// depends on where the host is placed, so the host's extent does too: the
// placed host, the frame box it is aligned by and the exclusion the following
// text wraps around must be those of ONE placement.
// Host: two 50pt cells, 10pt top/bottom cell margins, page x 100, w 100,
// notBeside. Child: 50pt wide at page x 150 (exactly over cell B's content,
// so ANCH cannot sit beside it), one exact 24pt line, no cell margins.
const pageChild = (yPt: number) => table([para('CHILD')], [50], 0, false).replace(
  '<w:tblPr>',
  '<w:tblPr><w:tblpPr w:leftFromText="0" w:rightFromText="0" w:topFromText="0" w:bottomFromText="0" '
    + `w:vertAnchor="page" w:horzAnchor="page" w:tblpX="${150 * TW}" w:tblpY="${yPt * TW}"/>`,
);
const alignedHost = (yAlign: 'center' | 'bottom', childYPt: number) => table([
  para('HOST', `<w:framePr w:w="${GRID_PT * TW}" w:wrap="notBeside" w:hAnchor="page" w:vAnchor="page" `
    + `w:x="${FRAME_X * TW}" w:yAlign="${yAlign}"/>`),
  pageChild(childYPt) + para('ANCH'),
], [GRID_PT / 2, GRID_PT / 2], CELL_TOP_PT, false);

describe('a page-aligned header-root host whose extent depends on its page position', () => {
  // Six 24pt fillers from the 20pt header distance end at 164; FOLLOW, the
  // seventh line, is the first header line the host's exclusion can reach.
  const fillers = ['F1', 'F2', 'F3', 'F4', 'F5', 'F6'].map((text) => para(text)).join('');

  it('centres the host by the extent it is painted with and wraps FOLLOW below that host', () => {
    // Child at page y 176. Laid out at the story cursor (page top) ANCH
    // misses the child and the host is 10 + 24 + 10 = 44pt; centred by that
    // extent its top would be 178, where ANCH (188) meets the child and drops
    // below it, so 44 is not the extent of the host placed there.
    // The one consistent placement, from the operands: with ANCH below the
    // child the host bottom is the child bottom 200 + 24 + 10 = 234, and
    // centring (top + bottom = 400) puts the top at 166. There ANCH's first
    // line would be 176–200, the child's span, so it is indeed displaced:
    // the host is 166–234 (68pt), centred, and holds ANCH.
    const { result } = layout({
      body: para('BODY'),
      header: alignedHost('center', 176) + fillers + para('FOLLOW'),
    });
    expect(placedOnce(result, 'CHILD')).toEqual({ x: 150, y: 176 });
    expect(placedOnce(result, 'HOST')).toEqual({ x: FRAME_X, y: 166 + CELL_TOP_PT });
    expect(placedOnce(result, 'ANCH')).toEqual({ x: 150, y: 200 });
    expect(placedOnce(result, 'F6')).toEqual({ x: 40, y: 140 });
    // notBeside: FOLLOW (164–188 meets 166–234) resumes on the first line
    // clear of the host the page shows, at its bottom 234.
    expect(placedOnce(result, 'FOLLOW')).toEqual({ x: 40, y: 234 });
  });

  it('fails closed when no placement is consistent with its own extent', () => {
    // Bottom-aligned, child at page y 370. A host whose ANCH misses the child
    // is 44pt and bottom-aligned at 356, where ANCH (366–390) meets the child
    // (370–394); one whose ANCH is displaced ends at 394 + 24 + 10 = 428,
    // which bottom alignment moves 28pt up, to where ANCH misses the child.
    // No placement reproduces itself, so layout must fail closed with the
    // library's exact-solve error instead of painting a host off its frame.
    expect(() => layout({
      body: para('BODY'),
      header: alignedHost('bottom', 370) + para('STORY-AFTER'),
    })).toThrow(/NON_CONVERGENCE/);
  });
});

describe('a page-aligned body-root host whose extent depends on its page position', () => {
  // The body counterpart of the header case above: the same host, child and
  // operands as the direct root of the body, which Word also keeps. The body
  // owner is framed by its yAlign on the page like the header one, so the
  // one consistent placement follows from the same operands: centred with
  // ANCH below the child, 166–234 (68pt, centre 200), HOST at 166 + 10 and
  // ANCH at the child bottom 200. A host framed by the 44pt extent it has
  // where ANCH misses the child is aligned at 178, where ANCH meets the child
  // (176–200) and the host is no longer 44pt: such a host is painted off the
  // frame it was aligned by.
  // Pagination does not decide: the whole host (68pt) fits the 400pt page it
  // is page-anchored on and the 320pt body band, so it is admitted whole on
  // page 1 and nothing continues.
  // Following body text: LEAD is one exact 12pt line, so the 24pt fillers
  // after it are offset from the header case: F4 ends at 148 and F5 spans
  // 148–172, which meets a host starting at 166 but not one starting at 178.
  // notBeside: F5 resumes on the first line clear of the host, at its bottom.
  const lead = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${12 * TW}" w:lineRule="exact"/></w:pPr>`
    + '<w:r><w:t>LEAD</w:t></w:r></w:p>';
  const fillers = ['F1', 'F2', 'F3', 'F4', 'F5'];

  it('centres the host by the extent it is painted with and wraps body text below that host', () => {
    const { result } = layout({
      body: alignedHost('center', 176) + lead + fillers.map((text) => para(text)).join('') + para('FOLLOW'),
    });
    expect(result.pages).toHaveLength(1);
    expect(result.diagnostics.filter((entry) => entry.severity === 'error')).toEqual([]);
    // The child keeps its authored page position; the host is the consistent
    // placement and holds ANCH, displaced below the child.
    expect(placedOnce(result, 'CHILD')).toEqual({ x: 150, y: 176 });
    expect(placedOnce(result, 'HOST')).toEqual({ x: FRAME_X, y: 166 + CELL_TOP_PT });
    expect(placedOnce(result, 'ANCH')).toEqual({ x: 150, y: 200 });
    // The host takes no body flow; the exclusion is the painted host's
    // 166–234: F4 (124–148) stays, F5 resumes at 234, FOLLOW below it.
    expect(placedOnce(result, 'LEAD')).toEqual({ x: 40, y: 40 });
    expect(placedOnce(result, 'F4')).toEqual({ x: 40, y: 124 });
    expect(placedOnce(result, 'F5')).toEqual({ x: 40, y: 234 });
    expect(placedOnce(result, 'FOLLOW')).toEqual({ x: 40, y: 258 });
    // Reading order is source order: the host row, then the body text.
    const order = textRunGeometryForPage(result, 0).map((run) => run.placement.text)
      .filter((text) => text !== 'CHILD');
    expect(order).toEqual(['HOST', 'ANCH', 'LEAD', ...fillers, 'FOLLOW']);
  });
});

describe('a page-aligned upright vertical-section host whose extent depends on its page position', () => {
  // The same direct body-root host and child in a §17.6.20 tbRl section. The
  // host is an upright physical table (library policy; vertical sections are
  // outside the native cell-owner controls): its frame
  // resolves on the physical page, which here is the same 400×400pt page
  // with 40pt margins, and its child is placed on that physical page. So the
  // operands of the horizontal body case decide it unchanged: laid out at
  // the frame of its 44pt unplaced extent (top 178) ANCH (188) meets the
  // child (176–200) and the host is no longer 44pt; the one consistent
  // placement is centred with ANCH below the child, 166–234 (68pt).
  // The host is atomic and takes no vertical flow, so pagination does not
  // decide either. Coordinates below are physical page points.

  it('centres the host by the extent it is painted with', () => {
    const { model, result } = layout({ body: alignedHost('center', 176), textDirection: 'tbRl' });
    // Preconditions: a vertical section whose host is an upright root.
    expect(result.pages).toHaveLength(1);
    expect(result.pages[0]!.section.textDirection).toBe('tbRl');
    const tableRoots = result.pages[0]!.layers.roots.filter((root) => root.node.kind === 'table');
    expect(tableRoots.length).toBeGreaterThan(0);
    expect(new Set(tableRoots.map((root) => root.coordinateSpace))).toEqual(new Set(['upright-physical']));
    expect(frameDeclarations(model)).toHaveLength(1);
    expect(result.diagnostics.filter((entry) => entry.severity === 'error')).toEqual([]);
    // Each run placed once: the child at its authored physical position,
    // ANCH displaced to the child's bottom, HOST in the consistent frame.
    const child = placedOnce(result, 'CHILD');
    const host = placedOnce(result, 'HOST');
    const anchor = placedOnce(result, 'ANCH');
    expect(child).toEqual({ x: 150, y: 176 });
    expect(anchor).toEqual({ x: 150, y: 200 });
    expect(host).toEqual({ x: FRAME_X, y: 166 + CELL_TOP_PT });
    // The painted extent (HOST's top cell margin to ANCH's exact line and
    // bottom cell margin) is the extent it is centred by on the 400pt page.
    const topPt = host.y - CELL_TOP_PT;
    const bottomPt = anchor.y + 24 + CELL_TOP_PT;
    expect(bottomPt - topPt).toBe(68);
    expect(topPt + bottomPt).toBe(400);
  });

  it('fails closed when no placement is consistent with its own extent', () => {
    // The header case's bottom-aligned operands on the upright physical
    // page: a 44pt host at 356 displaces ANCH and is 72pt, which bottom
    // alignment moves to where ANCH misses the child. An atomic upright host
    // has no other region to try, so layout must fail closed with the
    // library's exact-solve error instead of painting off its frame.
    expect(() => layout({ body: alignedHost('bottom', 370), textDirection: 'tbRl' }))
      .toThrow(/NON_CONVERGENCE/);
  });
});

describe('cell-owner carriers that Word imports as inert', () => {
  // Nested in a body or footer cell (r2 containment topology: body flow end
  // 360, footer flow end 380, page 400; frame y 290 inside / 370 beyond).
  // Defect caught: promoting the nested carrier into an out-of-flow cell
  // host — which grows/moves the outer row, maps the host from the page into
  // its cell, or fails the footer band solve (the two r2 footer errors) —
  // where Word renders the frame-free control.
  // One- and two-cell tables each time (Word drops both).
  for (const yPt of [290, 370]) {
    it(`keeps tables nested in a body cell, frame y ${yPt}, as plain content`, () => {
      for (const cellCount of CELL_COUNTS) {
        const stories = (framed: boolean) => ({
          body: para('BODY') + nestedIn(hostTable(cellCount, framed ? { kind: 'page', yPt } : null, true))
            + para('BODY-AFTER'),
        });
        expectInert(stories(true), stories(false));
      }
    });

    it(`keeps tables nested in a footer cell, frame y ${yPt}, as plain content`, () => {
      for (const cellCount of CELL_COUNTS) {
        const stories = (framed: boolean) => ({
          body: para('BODY'),
          footer: nestedIn(hostTable(cellCount, framed ? { kind: 'page', yPt } : null, true)) + para('STORY-AFTER'),
        });
        expectInert(stories(true), stories(false));
      }
    });
  }

  // Defect caught: the nested gate keyed on page/page or notBeside only —
  // Word drops width-omitted margin/text around nested carriers too.
  it('keeps nested width-omitted margin/text around carriers as plain content', () => {
    for (const cellCount of CELL_COUNTS) {
      const stories = (framed: boolean) => ({
        body: para('BODY') + nestedIn(hostTable(cellCount, framed ? { kind: 'around', yPt: 0 } : null, true))
          + para('BODY-AFTER'),
      });
      expectInert(stories(true), stories(false));
    }
  });

  // Text box story roots with rich content around the table. Defect caught:
  // promoting a text box story-root carrier — resolving it through the box
  // placement (moving HOST on a quarter turn, shifting AFTER by a zero host
  // advance) or recording an unmapped row-owner warning on a non-quarter
  // turn — where Word renders the frame-free control at every tested angle.
  // One representative per tested class: horizontal unrotated and +30°, and
  // a stacked (quarter-turned) story at 90° and −30°, alternating cell counts.
  for (const [degrees, vert, cellCount] of [
    [0, 'horz', 1], [30, 'horz', 2], [90, 'wordArtVert', 1], [-30, 'wordArtVert', 2],
  ] as const) {
    it(`keeps a ${cellCount}-cell ${vert} text box story-root carrier turned ${degrees}° as plain content`, () => {
      const stories = (framed: boolean) => ({
        body: textBox(degrees, vert, para('BEFORE')
          + hostTable(cellCount, framed ? { kind: 'page', yPt: 200 } : null) + para('AFTER')),
      });
      expectInert(stories(true), stories(false));
    });
  }

  // Footnote and endnote story roots are wired independently. Defect caught:
  // promoting a note story-root carrier — placing a page-frame host, adding a
  // host layer, or growing the note reserve that moves body text — where
  // Word drops the frame and renders the frame-free control.
  for (const kind of ['footnote', 'endnote'] as const) {
    it(`keeps ${kind} story-root carriers as plain content`, () => {
      for (const [frame, cellCount] of [
        [{ kind: 'page', yPt: 200 }, 1], [{ kind: 'around', yPt: 0 }, 2],
      ] as const) {
        const stories = (framed: boolean) => ({
          body: para('BODY') + noteReference(kind) + para('BODY-AFTER'),
          [kind]: para('NOTE') + hostTable(cellCount, framed ? frame : null) + para('NOTE-AFTER'),
        });
        expectInert(stories(true), stories(false));
      }
    });
  }
});
