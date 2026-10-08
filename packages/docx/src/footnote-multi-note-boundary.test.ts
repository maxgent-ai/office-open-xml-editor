import { expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';
import type { DocumentLayout, NoteLayout, PaintNode } from './layout/types.js';

// Public-model (parser-independent) boundary for the reference-line footnote
// plan with several notes. Exact 10pt lines isolate admission and ownership;
// these synthetic glyph metrics are not an Office pagination or font oracle.
const context = {
  font: '10px Arial', letterSpacing: '0px', fontKerning: 'none',
  measureText: (text: string) => ({
    width: [...text].length * 6,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
  }),
} as unknown as CanvasRenderingContext2D;

const label = (prefix: string, index: number) => prefix + String(index + 1).padStart(2, '0');

function textRun(text: string, extra: Record<string, unknown> = {}) {
  return {
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Arial', fontFamilyEastAsia: '', isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...extra,
  };
}

function paragraph(prefix: string, count: number, referenceIds: string[] = []): BodyElement {
  const runs = Array.from({ length: count }, (_, index) => [
    ...(index === 0 ? referenceIds.map(id => textRun(id, { noteRef: { kind: 'footnote', id } })) : []),
    textRun(label(prefix, index)),
    ...(index + 1 < count ? [{ type: 'break', breakType: 'line' }] : []),
  ]).flat();
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: { value: 10, rule: 'exact', explicit: true },
    numbering: null, tabStops: [], runs: runs as DocParagraph['runs'],
    defaultFontSize: 10, defaultFontFamily: 'Arial', widowControl: false,
  } as unknown as BodyElement;
}

/** Twelve preceding body lines, then a two-line paragraph whose first line
 * references a one-line and a four-line footnote; `remainingTwips` of body
 * region remain after the preceding lines (page 200pt wide, 10pt margins). */
function model(remainingTwips: number): DocxDocumentModel {
  return {
    section: {
      pageWidth: 200, pageHeight: (12 * 200 + remainingTwips + 400) / 20,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      sectionStart: 'nextPage',
    } as SectionProps,
    body: [paragraph('P', 12), paragraph('B', 2, ['1', '2'])],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
    footnotes: [
      { id: '1', content: [paragraph('N1L', 1)] },
      { id: '2', content: [paragraph('N2L', 4)] },
    ],
  } as unknown as DocxDocumentModel;
}

function text(block: PaintNode): string {
  return block.kind === 'paragraph'
    ? block.lines.flatMap(line => line.placements.flatMap(run => run.kind === 'text' ? [run.text] : [])).join('')
    : '';
}

const notesOnPage = (page: DocumentLayout['pages'][number]): NoteLayout[] =>
  page.layers.notes.filter((node): node is NoteLayout => node.kind === 'note');

it.each([
  // 1120 twips = 56pt: reference line 10 + first note (6pt missing-story band
  // + 10) 16 + three final-note lines 30. One twip less keeps two lines. B02
  // never fits beside the committed plan (66 > 56).
  { remainingTwips: 1120, firstPageTail: 'N2L01N2L02N2L03', nextPage: 'N2L04' },
  { remainingTwips: 1119, firstPageTail: 'N2L01N2L02', nextPage: 'N2L03N2L04' },
])('partitions the final referenced note at the reference line at the $remainingTwips twip boundary', ({
  remainingTwips, firstPageTail, nextPage,
}) => {
  const document = model(remainingTwips);
  const layout = layoutDocument(document, createLayoutServices(document, {
    measureContext: context, allowFootnoteContinuation: true,
  }), { currentDateMs: 0 });
  const preceding = Array.from({ length: 12 }, (_, index) => label('P', index)).join('');
  expect(layout.pages.map(page => page.layers.body.map(text))).toEqual([[preceding, '12B01'], ['B02']]);
  expect(layout.pages.map(page => notesOnPage(page).map(note => note.story.blocks.map(text).join(''))))
    .toEqual([['N1L01', firstPageTail], [nextPage]]);
  for (const [id, lines] of [['1', 'N1L01'], ['2', 'N2L01N2L02N2L03N2L04']] as const) {
    expect(layout.pages.flatMap(page => notesOnPage(page)
      .filter(note => note.source.storyInstance === id)
      .flatMap(note => note.story.blocks.map(text))).join('')).toBe(lines);
  }
  for (const page of layout.pages) {
    // Exactly one separator per page note band, owned by its first note.
    expect(notesOnPage(page).map(note => note.separator.length > 0))
      .toEqual(notesOnPage(page).map((_, index) => index === 0));
    for (const note of notesOnPage(page)) {
      expect(note.flowDomainId).toBe(`notes:page:${page.pageIndex}`);
    }
  }
});
