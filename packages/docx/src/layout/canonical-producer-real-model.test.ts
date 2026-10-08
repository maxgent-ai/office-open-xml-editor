import { describe, expect, it } from 'vitest';
import { buildBookmarkPageMap } from '../bookmark-nav.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { layoutDocumentInputAsync } from './document.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from '../types.js';
import type { NoteLayout } from './types.js';

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 5,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function paragraph(): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: '1', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'serif',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: 'super',
      hyperlink: null, noteRef: { kind: 'footnote', id: '7' },
    }],
    defaultFontSize: 10, defaultFontFamily: 'serif', widowControl: false,
  } as unknown as DocParagraph;
}

function ordinaryParagraph(text: string): DocParagraph {
  const result = paragraph();
  const run = result.runs[0];
  if (run?.type === 'text') {
    run.text = text;
    run.noteRef = undefined;
    run.vertAlign = null;
  }
  return result;
}

function ordinaryBodyParagraph(text: string): BodyElement {
  return ordinaryParagraph(text) as unknown as BodyElement;
}

function endnoteReferenceParagraph(id: string): BodyElement {
  const result = paragraph();
  const run = result.runs[0];
  if (run?.type === 'text') {
    run.noteRef = { kind: 'endnote', id };
  }
  return result as unknown as BodyElement;
}

function noteMarkerParagraph(kind: 'footnote' | 'endnote'): DocParagraph {
  const result = ordinaryParagraph('');
  const run = result.runs[0];
  if (run?.type === 'text') {
    run.noteRef = { kind, id: '' };
    run.vertAlign = 'super';
  }
  return result;
}

function pageOwnedWrappingParagraph(): BodyElement {
  const result = ordinaryParagraph('');
  result.runs = [{
    type: 'shape',
    widthPt: 30,
    heightPt: 20,
    anchorXPt: 0,
    anchorYPt: 0,
    anchorXFromMargin: true,
    anchorYFromPara: false,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'FFFFFF' },
    stroke: null,
    wrapMode: 'square',
    wrapSide: 'bothSides',
    distTop: 0,
    distBottom: 0,
    distLeft: 0,
    distRight: 0,
  }] as DocParagraph['runs'];
  return result as unknown as BodyElement;
}

function unsupportedTextPathParagraph(): BodyElement {
  const result = ordinaryParagraph('');
  result.runs = [{
    type: 'shape',
    widthPt: 30,
    heightPt: 20,
    anchorXPt: 0,
    anchorYPt: 0,
    anchorXFromMargin: false,
    anchorYFromPara: true,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'D9D9D9' },
    stroke: null,
    textPath: {
      string: 'DRAFT',
      fontFamily: 'Arial',
      bold: false,
      italic: false,
      textPathOk: true,
      on: true,
      fitShape: true,
      fitPath: true,
      trim: false,
      xScale: false,
      fontSizePt: 12,
    },
  }] as unknown as DocParagraph['runs'];
  return result as unknown as BodyElement;
}

function sectionBreak(): BodyElement {
  return {
    type: 'sectionBreak', kind: 'continuous', columns: null,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    titlePage: false,
  } as unknown as BodyElement;
}

function sectionBoundaryModel(
  before: BodyElement,
  startType: 'continuous' | 'nextPage',
): DocxDocumentModel {
  const section = {
    pageWidth: 200, pageHeight: 100,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 5, footerDistance: 5, titlePage: false,
    evenAndOddHeaders: false, sectionStart: startType, columns: null,
  } as SectionProps;
  return {
    section,
    body: [before, sectionBreak(), ordinaryBodyParagraph('after')],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [], endnotes: [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function floatingTable(): BodyElement {
  return {
    type: 'table', colWidths: [40],
    rows: [{
      cells: [{
        content: [ordinaryParagraph('1')], colSpan: 1, vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null, vAlign: 'top', widthPt: 40,
      }],
      rowHeight: null, rowHeightRule: 'auto', isHeader: false,
    }],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed', overlap: 'overlap',
    tblpPr: {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'text', horzSpecified: true, vertAnchor: 'text',
      tblpX: 20, tblpY: 10,
    },
  } as unknown as BodyElement;
}

function ordinaryTable(
  justification: 'left' | 'center' | 'right',
  indentPt: number,
  bidiVisual: boolean,
): BodyElement {
  const source = floatingTable() as Extract<BodyElement, { type: 'table' }>;
  return {
    ...source,
    jc: justification,
    tblInd: indentPt,
    bidiVisual,
    tblpPr: null,
  } as unknown as BodyElement;
}

function fragmentLineAdvancesPt(fragment: Extract<ReturnType<typeof layoutDocument>['pages'][number]['layers']['body'][number], { kind: 'paragraph' }>): number {
  return fragment.lines.reduce((sum, line) => sum + line.advancePt, 0);
}

describe('canonical producer with a real document model', () => {
  it('keeps valid surrounding body blocks when an optional drawing feature is unsupported', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [
        ordinaryBodyParagraph('before'),
        unsupportedTextPathParagraph(),
        ordinaryBodyParagraph('after'),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const paragraphs = layout.pages.flatMap((page) => page.layers.body)
      .filter((node) => node.kind === 'paragraph');
    const text = paragraphs.flatMap((paragraph) => paragraph.lines)
      .flatMap((line) => line.placements)
      .filter((placement) => placement.kind === 'text')
      .map((placement) => placement.text);
    const unsupported = paragraphs.find((paragraph) => paragraph.source.path[0] === 1);

    expect(text).toEqual(expect.arrayContaining(['before', 'after']));
    expect(unsupported?.drawings[0]?.commands).toEqual([{ kind: 'noop' }]);
    expect(layout.diagnostics).toContainEqual({
      code: 'UNSUPPORTED_FEATURE',
      severity: 'error',
      source: { story: 'body', storyInstance: 'body', path: [1, 0] },
      message: 'VML textPath fitPath=true is not rendered',
    });
    expect(Object.isFrozen(layout.diagnostics[0])).toBe(true);
  });

  it('lays out document-end notes through the retained story engine', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [endnoteReferenceParagraph('2')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [],
      endnotes: [
        { id: '1', content: [ordinaryParagraph('unreferenced')] },
        { id: '2', content: [noteMarkerParagraph('endnote')] },
      ],
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const endnote = page.layers.notes.find((node) => node.source.story === 'endnote');

    expect(layout.diagnostics).toEqual([]);
    expect(endnote).toMatchObject({
      kind: 'note',
      source: { story: 'endnote', storyInstance: '2' },
      flowDomainId: 'endnotes:page:0',
    });
    expect(endnote!.flowBounds.yPt).toBeGreaterThanOrEqual(
      page.layers.body[0]!.flowBounds.yPt + page.layers.body[0]!.flowBounds.heightPt,
    );
    expect(page.readingOrder).toEqual([page.layers.body[0]!.id, endnote!.id]);
    if (endnote?.kind !== 'note') throw new Error('Expected retained endnote');
    const storyParagraph = endnote.story.blocks[0];
    expect(storyParagraph?.kind).toBe('paragraph');
    if (storyParagraph?.kind !== 'paragraph') throw new Error('Expected retained endnote paragraph');
    expect(storyParagraph.lines.flatMap((line) => line.placements)
      .find((placement) => placement.kind === 'text' && placement.noteReference))
      .toMatchObject({ kind: 'text', text: '1' });
  });

  it('reports concrete document-end note overflow without hiding other layout failures', () => {
    const section = {
      pageWidth: 80, pageHeight: 50,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [endnoteReferenceParagraph('1')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [],
      endnotes: [{ id: '1', content: [ordinaryParagraph('endnote '.repeat(80))] }],
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages.flatMap((page) => page.layers.notes)).toEqual([]);
    expect(layout.diagnostics).toEqual([
      expect.objectContaining({
        code: 'UNSUPPORTED_FEATURE',
        severity: 'error',
        source: expect.objectContaining({ story: 'endnote', storyInstance: '1' }),
        message: expect.stringContaining('do not fit the retained terminal flow region'),
      }),
    ]);
  });

  it('surfaces unsupported authored note positions instead of silently treating them as defaults', () => {
    const section = {
      pageWidth: 200, pageHeight: 140,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [paragraph() as unknown as BodyElement, endnoteReferenceParagraph('1')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [ordinaryParagraph('footnote')] }],
      endnotes: [{ id: '1', content: [ordinaryParagraph('endnote')] }],
      fontFamilyClasses: {},
      __noteLayoutSettings: {
        footnotePosition: 'beneathText',
        endnotePosition: 'sectEnd',
      },
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages.flatMap((page) => page.layers.notes).map((note) => note.source.story))
      .toEqual(['endnote', 'footnote']);
    expect(layout.diagnostics.map((diagnostic) => diagnostic.message)).toEqual([
      expect.stringContaining('Unsupported footnote position "beneathText"'),
      expect.stringContaining('Unsupported endnote position "sectEnd"'),
    ]);
  });

  it('retains parser-model bookmark starts through production pagination', () => {
    const anchored = ordinaryParagraph('destination');
    anchored.bookmarks = ['destination', 'alias'];
    const model = sectionBoundaryModel(
      anchored as unknown as BodyElement,
      'nextPage',
    );
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages[0]?.bookmarkStarts.map(({ name }) => name))
      .toEqual(['destination', 'alias']);
    expect([...buildBookmarkPageMap(layout)]).toEqual([
      ['destination', 0],
      ['alias', 0],
    ]);
  });

  it('derives bookmark page metadata after header stories join the retained graph', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const headerDestination = ordinaryParagraph('header destination');
    headerDestination.bookmarks = ['header-destination'];
    const model = {
      section,
      body: [ordinaryBodyParagraph('body')],
      headers: {
        default: { body: [headerDestination as unknown as BodyElement] },
        first: null,
        even: null,
      },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages[0]?.bookmarkStarts).toEqual([
      expect.objectContaining({
        name: 'header-destination',
        sectionOccurrenceId: layout.pages[0]?.sectionOccurrenceId,
      }),
    ]);
    expect([...buildBookmarkPageMap(layout)]).toEqual([['header-destination', 0]]);
  });

  it('composes compatible continuous sections as disjoint regions on one physical page', () => {
    const model = sectionBoundaryModel(ordinaryBodyParagraph('before'), 'continuous');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const [outgoing, incoming] = page.sectionRegions;

    expect(layout.pages).toHaveLength(1);
    expect(page.sectionRegions).toHaveLength(2);
    expect(outgoing!.blockEndPt).toBe(incoming!.blockStartPt);
    expect(outgoing!.blockEndPt).toBeGreaterThan(outgoing!.blockStartPt);
    expect(page.layers.body.map((node) => node.flowDomainId)).toEqual([
      outgoing!.flowDomainIds[0], incoming!.flowDomainIds[0],
    ]);
  });

  it('retains a floating table in an empty outgoing region without charging incoming flow', () => {
    const model = sectionBoundaryModel(floatingTable(), 'continuous');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const [outgoing, incoming] = page.sectionRegions;
    const [floating, follower] = page.layers.body;

    expect(layout.pages).toHaveLength(1);
    expect(page.sectionRegions).toHaveLength(2);
    expect(outgoing!.blockStartPt).toBe(outgoing!.blockEndPt);
    expect(outgoing!.blockEndPt).toBe(incoming!.blockStartPt);
    expect(floating).toMatchObject({
      kind: 'table', ordinaryFlow: false, flowDomainId: outgoing!.flowDomainIds[0],
    });
    expect(follower).toMatchObject({
      kind: 'paragraph', ordinaryFlow: true, flowDomainId: incoming!.flowDomainIds[0],
      flowBounds: { yPt: incoming!.blockStartPt },
    });
    expect(page.flowDomains.find((domain) => domain.id === outgoing!.flowDomainIds[0]))
      .toMatchObject({ logicalBounds: { heightPt: 0 } });
  });

  it('keeps a non-continuous control on separate physical pages', () => {
    const model = sectionBoundaryModel(floatingTable(), 'nextPage');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => page.sectionRegions.length)).toEqual([1, 1]);
    expect(layout.pages[0]!.layers.body[0]).toMatchObject({
      kind: 'table', ordinaryFlow: false,
      flowDomainId: layout.pages[0]!.sectionRegions[0]!.flowDomainIds[0],
    });
    expect(layout.pages[1]!.layers.body[0]).toMatchObject({
      kind: 'paragraph', ordinaryFlow: true,
      flowDomainId: layout.pages[1]!.sectionRegions[0]!.flowDomainIds[0],
    });
  });

  it('advances a single-column nextColumn section to its own next-page geometry', () => {
    const outgoing = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5,
    };
    const incoming = {
      pageWidth: 300, pageHeight: 160,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 8, footerDistance: 8, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextColumn', columns: null,
    } as SectionProps;
    const endingSection = {
      type: 'sectionBreak',
      kind: 'nextPage',
      geom: outgoing,
      columns: null,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      titlePage: false,
    } as unknown as BodyElement;
    const model = {
      section: incoming,
      body: [
        ordinaryBodyParagraph('before'),
        endingSection,
        ordinaryBodyParagraph('after'),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => ({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    }))).toEqual([
      { widthPt: 200, heightPt: 100 },
      { widthPt: 300, heightPt: 160 },
    ]);
    expect(layout.pages.map((page) => page.sectionRegions.length)).toEqual([1, 1]);
  });

  it('retains page-owned wrap authority across a same-page nextColumn section cutover', () => {
    const columns = { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] };
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextColumn', columns,
    } as SectionProps;
    const endingSection = {
      type: 'sectionBreak',
      kind: 'nextPage',
      geom: { ...section, sectionStart: 'nextPage' },
      columns,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      titlePage: false,
    } as unknown as BodyElement;
    const model = {
      section,
      body: [ordinaryBodyParagraph('before'), endingSection, pageOwnedWrappingParagraph()],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]!.sectionRegions).toHaveLength(2);
    expect(layout.pages[0]!.layers.body.some((node) =>
      node.kind === 'paragraph' && node.drawings.length > 0)).toBe(true);
  });

  it.each([
    ['LTR left positive', 'left', 12, false, 22],
    ['LTR left negative', 'left', -12, false, -2],
    ['LTR center positive', 'center', 12, false, 92],
    ['LTR center negative', 'center', -12, false, 68],
    ['LTR right positive', 'right', 12, false, 162],
    ['LTR right negative', 'right', -12, false, 138],
    ['RTL left positive', 'left', 12, true, 138],
    ['RTL left negative', 'left', -12, true, 162],
    ['RTL center positive', 'center', 12, true, 68],
    ['RTL center negative', 'center', -12, true, 92],
    ['RTL right positive', 'right', 12, true, -2],
    ['RTL right negative', 'right', -12, true, 22],
  ] as const)(
    'projects parser-owned tblInd placement for %s',
    (_name, justification, indentPt, bidiVisual, expectedXPt) => {
      const section = {
        pageWidth: 200, pageHeight: 100,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5, titlePage: false,
        evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
      } as SectionProps;
      const model = {
        section, body: [ordinaryTable(justification, indentPt, bidiVisual)],
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        footnotes: [], endnotes: [], fontFamilyClasses: {},
      } as unknown as DocxDocumentModel;
      const services = createLayoutServices(model, { measureContext: measureContext() });

      const retained = layoutDocument(model, services, { currentDateMs: 0 })
        .pages[0]?.layers.body[0];

      expect(retained?.kind).toBe('table');
      if (retained?.kind !== 'table') throw new Error('expected retained table');
      expect(retained.flowBounds.xPt).toBe(expectedXPt);
      expect(retained.rows[0]?.flowBounds.xPt).toBe(expectedXPt);
    },
  );

  it('retains the note-reference id on the destination-page text placement', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [paragraph()],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [paragraph()] }],
      endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const placement = layout.pages[0]!.layers.body
      .filter((node) => node.kind === 'paragraph')
      .flatMap((node) => node.kind === 'paragraph' ? node.lines : [])
      .flatMap((line) => line.placements)
      .find((candidate) => candidate.kind === 'text');

    expect(placement).toMatchObject({ noteReference: { kind: 'footnote', id: '7' } });
  });

  it('removes only the explicitly empty footnote separator rule', () => {
    const section = {
      pageWidth: 200, pageHeight: 140,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const base = {
      section, body: [paragraph()],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [ordinaryParagraph('note')] }],
      endnotes: [], fontFamilyClasses: {},
    };
    const render = (separator?: string) => {
      const model = {
        ...base, __noteLayoutSettings: { footnoteSeparator: separator },
      } as unknown as DocxDocumentModel;
      const services = createLayoutServices(model, { measureContext: measureContext() });
      return layoutDocument(model, services, { currentDateMs: 0 })
        .pages.flatMap((page) => page.layers.notes)[0];
    };
    const standard = render();
    const empty = render('none');
    expect(standard?.kind).toBe('note');
    expect(empty?.kind).toBe('note');
    if (standard?.kind !== 'note' || empty?.kind !== 'note') {
      throw new Error('Expected retained footnote');
    }
    expect(standard?.separator).toHaveLength(1);
    expect(empty?.separator).toHaveLength(0);
    expect(empty?.flowBounds).toEqual(standard?.flowBounds);
  });

  it('uses the authored continuation rule width on a retained note fragment', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const render = (style?: 'short' | 'full' | 'none', ordinary?: 'short' | 'none', count = 8, longBody = false, leading = 0) => {
      const bodyParagraph = paragraph();
      if (longBody) bodyParagraph.runs = [...bodyParagraph.runs, ...ordinaryParagraph('body '.repeat(100)).runs];
      const model = {
        section, body: [...Array.from({ length: leading }, () => ordinaryParagraph('body')), bodyParagraph],
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        footnotes: [{ id: '7', content: Array.from({ length: count }, () => ordinaryParagraph('note')) }],
        endnotes: [], fontFamilyClasses: {},
        __noteLayoutSettings: { footnoteContinuationSeparator: style, footnoteSeparator: ordinary },
      } as unknown as DocxDocumentModel;
      const services = createLayoutServices(model, {
        measureContext: measureContext(), allowFootnoteContinuation: true,
      });
      return layoutDocument(model, services, { currentDateMs: 0 }).pages
        .flatMap((page) => page.layers.notes)
        .filter((note) => note.kind === 'note');
    };
    const relocated = render('full', undefined, 2, false, 6);
    expect(relocated[0]?.flowDomainId).toBe('notes:page:1');
    expect(relocated[0]?.story.blocks.every(block =>
      block.flowDomainId === 'notes:page:1:footnote:7')).toBe(true);
    const short = render('short', 'short');
    const full = render('full');
    expect(short.length).toBeGreaterThan(1);
    expect(full).toHaveLength(short.length);
    const width = (note: (typeof short)[number]) => {
      const edge = note.separator[0];
      return edge ? edge.to.xPt - edge.from.xPt : 0;
    };
    expect(width(short[0]!)).toBeCloseTo(width(full[0]!), 5);
    expect(width(short[1]!)).toBeCloseTo(60, 5);
    expect(width(full[1]!)).toBeCloseTo(180, 5);
    // A missing story retains the full-width continuation fallback; it must
    // not inherit the ordinary Short rule or become an explicit empty rule.
    const absent = render();
    expect(absent.map(width)).toEqual(full.map(width));
    const emptyFirst = render('full', 'none');
    expect(width(emptyFirst[0]!)).toBe(0);
    expect(width(emptyFirst[1]!)).toBeCloseTo(180, 5);
    const emptyContinuation = render('none');
    expect(width(emptyContinuation[0]!)).toBeCloseTo(60, 5);
    expect(width(emptyContinuation[1]!)).toBe(0);
    expect(emptyContinuation.map(note => note.advancePt)).toEqual(full.map(note => note.advancePt));
    // A note taller than a physical page must be acquired before partitioning.
    const long = render('full', undefined, 30);
    expect(long.length).toBeGreaterThan(2);
    expect(long.flatMap(note => note.story.blocks).filter(block => block.kind === 'paragraph')).toHaveLength(30);
    const continuedBody = render('full', undefined, 30, true);
    expect(continuedBody.length).toBeGreaterThan(2);
    expect(continuedBody.flatMap(note => note.story.blocks).filter(block => block.kind === 'paragraph')).toHaveLength(30);
  });

  it.each([{ prefix: 0, rejects: true }, { prefix: 35, rejects: false }])(
    'validates the source cut after a page-field reflow (prefix $prefix)', ({ prefix, rejects }) => {
    const note = ordinaryParagraph('x'.repeat(prefix));
    note.runs = [
      ...note.runs,
      { ...ordinaryParagraph('field').runs[0], type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '1' },
      ...ordinaryParagraph(' tail'.repeat(80)).runs,
    ] as DocParagraph['runs'];
    const model = {
      section: { pageWidth: 200, pageHeight: 100, marginTop: 10, marginBottom: 10,
        marginLeft: 10, marginRight: 10, headerDistance: 5, footerDistance: 5,
        titlePage: false, evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
        pageNumType: { start: 9, format: 'decimal' } },
      body: [paragraph()], headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [note] }], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: true,
    });
    // PAGE 9 -> 10 moves a word across the old line-six cut. Resuming at the
    // same shaped line would paint 81 copies of "tail" from a source of 80.
    if (rejects) {
      expect(() => layoutDocument(model, services, { currentDateMs: 0 }))
        .toThrow('Footnote continuation source cut changed during reflow');
    } else {
      const layout = layoutDocument(model, services, { currentDateMs: 0 });
      const text = layout.pages.flatMap(page => page.layers.notes).flatMap(note =>
        note.kind === 'note' ? note.story.blocks.flatMap(block => block.kind === 'paragraph'
          ? block.lines.flatMap(line => line.placements.flatMap(placement => placement.kind === 'text'
            ? [placement.text] : [])) : []) : []).join('');
      expect(text).toBe('x'.repeat(prefix) + '9' + ' tail'.repeat(80));
    }
  });

  it('rejects an over-budget note before shaping its source text', () => {
    const model = { section: {
      pageWidth: 200, pageHeight: 100, marginTop: 10, marginBottom: 10,
      marginLeft: 10, marginRight: 10, headerDistance: 5, footerDistance: 5,
      titlePage: false, evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    }, body: [paragraph()], headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [ordinaryParagraph('x'.repeat(1_000_001))] }],
      endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const context = measureContext();
    const measure = context.measureText;
    context.measureText = text => {
      if (text.length > 100) throw new Error('Over-budget source reached shaping');
      return measure(text);
    };
    const services = createLayoutServices(model, {
      measureContext: context, allowFootnoteContinuation: true,
    });
    expect(() => layoutDocument(model, services, { currentDateMs: 0 }))
      .toThrow('Footnote acquisition source budget exceeded');
  });

  it('retains a frame paragraph as an out-of-flow placed occurrence', () => {
    const framed = paragraph();
    framed.framePr = {
      dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
      hRule: 'auto', hSpace: 0, vSpace: 0, w: 40, h: 20, x: 15, y: 5,
    };
    const following = paragraph();
    (following.runs[0] as { noteRef?: unknown }).noteRef = undefined;
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [framed, following],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const body = layout.pages[0]!.layers.body;

    const frame = body[0]!;
    const follower = body[1]!;
    expect(frame.kind).toBe('paragraph');
    if (frame.kind !== 'paragraph') throw new Error('expected retained frame paragraph');
    expect(frame.ordinaryFlow).toBe(false);
    expect(frame.advancePt).toBe(0);
    expect(frame.flowBounds.heightPt).toBeCloseTo(
      frame.spacing.beforePt
        + fragmentLineAdvancesPt(frame)
        + frame.spacing.afterPt,
      6,
    );
    expect(frame.flowBounds).toMatchObject({ xPt: 25, yPt: 15 });
    expect(follower.flowBounds.yPt).toBe(10);
  });

  it('retains an effective positioned table without charging ordinary flow', () => {
    const following = paragraph();
    (following.runs[0] as { noteRef?: unknown }).noteRef = undefined;
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [floatingTable(), following],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const body = layout.pages[0]!.layers.body;

    const floating = body[0]!;
    const follower = body[1]!;
    expect(floating.kind).toBe('table');
    if (floating.kind !== 'table') throw new Error('expected retained floating table');
    expect(floating.ordinaryFlow).toBe(false);
    expect(floating.advancePt).toBeCloseTo(
      floating.rows.reduce((sum, row) => sum + row.advancePt, 0),
      6,
    );
    expect(floating.advancePt).toBeGreaterThan(0);
    expect(floating.flowBounds).toMatchObject({ xPt: 30, yPt: 20 });
    expect(follower.flowBounds.yPt).toBe(10);
  });

  it.each([
    { name: 'default top distance with an empty paragraph', topFromText: 0, emptyParagraph: true },
    { name: 'positive top distance without an empty paragraph', topFromText: 20, emptyParagraph: false },
  ])('keeps preceding text lines outside a page-positioned floating table: $name', ({
    topFromText, emptyParagraph,
  }) => {
    const section = {
      pageWidth: 200, pageHeight: 120,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
    table.colWidths = [180];
    table.rows[0]!.cells[0]!.widthPt = 180;
    table.tblpPr = {
      ...table.tblpPr!,
      horzAnchor: 'margin', vertAnchor: 'page', tblpX: 0, tblpY: 25,
      topFromText,
    };
    const tableIndex = emptyParagraph ? 2 : 1;
    const model = {
      section,
      body: [ordinaryBodyParagraph('before '.repeat(12)),
        ...(emptyParagraph ? [ordinaryBodyParagraph('')] : []), table],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
      currentDateMs: 0,
    });
    const body = layout.pages[0]!.layers.body;
    const heading = body.find((node) => node.kind === 'paragraph' && node.source.path[0] === 0);
    const placedTable = body.find((node) => node.kind === 'table'
      && node.source.path[0] === tableIndex);
    expect(heading?.kind).toBe('paragraph');
    expect(placedTable?.kind).toBe('table');
    if (heading?.kind !== 'paragraph' || placedTable?.kind !== 'table') return;
    expect(placedTable.flowBounds.yPt).toBe(25);
    expect(heading.lines.length).toBeGreaterThan(1);
    expect(heading.lines.every((line) => (
      line.bounds.yPt + line.bounds.heightPt <= placedTable.flowBounds.yPt
      || line.bounds.yPt >= placedTable.flowBounds.yPt + placedTable.flowBounds.heightPt
    ))).toBe(true);
    if (topFromText > 0) {
      expect(heading.lines.every((line) => line.bounds.yPt >= placedTable.flowBounds.yPt
        + placedTable.flowBounds.heightPt)).toBe(true);
    }
  });

  it('does not apply a later section page table to the preceding section', () => {
    const section = {
      pageWidth: 200, pageHeight: 120,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
    table.colWidths = [180];
    table.rows[0]!.cells[0]!.widthPt = 180;
    table.tblpPr = {
      ...table.tblpPr!,
      horzAnchor: 'margin', vertAnchor: 'page', tblpX: 0, tblpY: 25,
    };
    const breakMark = {
      ...sectionBreak(), kind: 'nextPage',
      geom: { ...section, sectionStart: 'nextPage' },
    } as unknown as BodyElement;
    const before = ordinaryBodyParagraph('before '.repeat(12));
    const model = (body: BodyElement[]) => ({
      section, body,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    }) as unknown as DocxDocumentModel;
    const control = model([before]);
    const withLaterTable = model([before, breakMark, table]);
    const controlLayout = layoutDocument(control,
      createLayoutServices(control, { measureContext: measureContext() }), { currentDateMs: 0 });
    const candidateLayout = layoutDocument(withLaterTable,
      createLayoutServices(withLaterTable, { measureContext: measureContext() }), { currentDateMs: 0 });
    const firstParagraph = (layout: ReturnType<typeof layoutDocument>) => layout.pages[0]!
      .layers.body.find((node) => node.kind === 'paragraph' && node.source.path[0] === 0);
    const controlParagraph = firstParagraph(controlLayout);
    const candidateParagraph = firstParagraph(candidateLayout);
    expect(controlParagraph?.kind).toBe('paragraph');
    expect(candidateParagraph?.kind).toBe('paragraph');
    if (controlParagraph?.kind !== 'paragraph' || candidateParagraph?.kind !== 'paragraph') return;
    expect(candidateParagraph.lines.map((line) => line.bounds.yPt))
      .toEqual(controlParagraph.lines.map((line) => line.bounds.yPt));
    expect(candidateLayout.pages[1]!.layers.body.some((node) => node.kind === 'table')).toBe(true);
  });

  it('keeps a continued floating table on its accepted page fragments', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
    table.colWidths = [180];
    table.rows = Array.from({ length: 8 }, (_, index) => ({
      ...table.rows[0]!,
      cells: [{
        ...table.rows[0]!.cells[0]!,
        widthPt: 180,
        content: [{ type: 'paragraph' as const, ...ordinaryParagraph(`row ${index} `.repeat(14)) }],
      }],
    }));
    table.tblpPr = {
      ...table.tblpPr!,
      horzAnchor: 'margin', vertAnchor: 'page', tblpX: 0, tblpY: 25,
    };
    const model = {
      section, body: [ordinaryBodyParagraph('before '.repeat(8)), table],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
      currentDateMs: 0,
    });
    const fragments = layout.pages.flatMap((page) => page.layers.body
      .filter((node) => node.kind === 'table' && node.source.path[0] === 1)
      .map((node) => ({ pageIndex: page.pageIndex, node })));
    expect(fragments.length).toBeGreaterThan(1);
    expect(fragments[0]!.pageIndex).toBe(1);
    expect(fragments[0]!.node.sectionFlowOwnership).toBe('page');
    expect(fragments.slice(1).every(({ node }) => node.sectionFlowOwnership !== 'page')).toBe(true);
    expect(new Set(fragments.map(({ pageIndex }) => pageIndex)).size).toBe(fragments.length);
  });

  it('rechecks page-owned table placement after a visible header changes the body band', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 30, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
    table.colWidths = [180];
    table.rows[0]!.cells[0]!.widthPt = 180;
    table.tblpPr = {
      ...table.tblpPr!,
      horzAnchor: 'margin', vertAnchor: 'page', tblpX: 0, tblpY: 65,
    };
    const model = {
      section,
      body: [ordinaryBodyParagraph('before '.repeat(20)), table],
      headers: { default: { body: [ordinaryParagraph('Header')] }, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
      currentDateMs: 0,
    });
    const headingPages = layout.pages.filter((page) => page.layers.body.some((node) =>
      node.kind === 'paragraph' && node.source.path[0] === 0));
    const tablePage = layout.pages.find((page) => page.layers.body.some((node) =>
      node.kind === 'table' && node.source.path[0] === 1));
    expect(headingPages.length).toBeGreaterThan(0);
    expect(tablePage).toBeDefined();
    expect(layout.pages[0]!.geometry.contentTopPt).toBeGreaterThan(section.marginTop);
    expect(layout.pages.every((page) => page.layers.body.length > 0)).toBe(true);
    expect(tablePage!.pageIndex).toBeGreaterThanOrEqual(headingPages.at(-1)!.pageIndex);
  });

  it('does not skip a page when the table source naturally follows multi-page text', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
    table.tblpPr = {
      ...table.tblpPr!,
      horzAnchor: 'margin', vertAnchor: 'page', tblpX: 0, tblpY: 25,
    };
    const model = {
      section,
      body: [ordinaryBodyParagraph('before '.repeat(150)), table],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
      currentDateMs: 0,
    });
    const tablePage = layout.pages.find((page) => page.layers.body.some((node) =>
      node.kind === 'table' && node.source.path[0] === 1));
    const lastTextPage = [...layout.pages].reverse().find((page) => page.layers.body.some((node) =>
      node.kind === 'paragraph' && node.source.path[0] === 0));
    expect(layout.pages.length).toBeGreaterThanOrEqual(3);
    expect(layout.pages.every((page) => page.layers.body.length > 0)).toBe(true);
    expect(tablePage).toBeDefined();
    expect(lastTextPage).toBeDefined();
    expect(tablePage!.pageIndex).toBeGreaterThanOrEqual(lastTextPage!.pageIndex);
  });
});

/** Measurement whose vertical metrics follow the requested CSS font size, so
 * a participant's own size is observable in the acquired line box. */
function sizedMeasureContext(): CanvasRenderingContext2D {
  const context = {
    font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
    measureText(text: string) {
      const px = Number(/([0-9.]+)px/.exec(context.font)?.[1] ?? 10);
      return {
        width: [...text].length * px / 2,
        actualBoundingBoxAscent: px * 0.8, actualBoundingBoxDescent: px * 0.2,
        fontBoundingBoxAscent: px * 0.8, fontBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  };
  return context as unknown as CanvasRenderingContext2D;
}

/** Native (MS-DOC 2.3.3) separator wire, shaped as the legacy producer emits it. */
const nativeAddress = (headerCp: number) => ({ headerCp, fc: 0x800 + headerCp * 2, prm: 0, paragraphStyle: 0 });
const nativeRun = (fontSize: number, color: string | null = null) => ({
  text: '', bold: false, italic: false, underline: false, strikethrough: false, fontSize, color,
  fontFamily: 'serif', isLink: false, background: null, vertAlign: null, hyperlink: null,
});
const nativeParagraph = (markPt: number, overrides: Record<string, unknown> = {}) => ({
  alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
  lineSpacing: null, numbering: null, tabStops: [], runs: [], defaultFontSize: markPt,
  defaultFontFamily: 'serif', widowControl: false,
  paragraphMarkFontFacts: { fontFamily: 'serif', fontSize: markPt }, ...overrides,
});
function nativeRule(mark: 'short' | 'full', cp: number, options: Readonly<{
  controlPt?: number; markPt?: number; color?: string; paragraph?: Record<string, unknown>;
}> = {}) {
  const markPt = options.markPt ?? 10;
  return {
    class: 'rule', contentStartCp: cp, contentEndCp: cp + 2, guardCp: cp + 2,
    rule: { mark, control: { run: nativeRun(options.controlPt ?? 10, options.color ?? null) }, source: nativeAddress(cp) },
    paragraph: {
      paragraph: nativeParagraph(markPt, options.paragraph), contentMark: { run: nativeRun(markPt) },
      source: nativeAddress(cp + 1),
    },
  };
}
const nativeNotice = (cp: number) => ({
  class: 'paragraphOnly', contentStartCp: cp, contentEndCp: cp + 1, guardCp: cp + 1,
  paragraph: { paragraph: nativeParagraph(10), contentMark: { run: nativeRun(10) }, source: nativeAddress(cp) },
});
const nativeEmpty = (cp: number) => ({ class: 'empty', contentStartCp: cp, contentEndCp: cp });

function nativeNotePages(
  stories: Readonly<{ separator: unknown; continuationSeparator: unknown; continuationNotice: unknown }>,
  options: Readonly<{ notes?: number; continuation?: boolean }> = {},
): NoteLayout[][] {
  const model = {
    section: {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    },
    body: [paragraph()],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [{ id: '7', content: Array.from({ length: options.notes ?? 1 }, () => ordinaryParagraph('note')) }],
    endnotes: [], fontFamilyClasses: {},
    __noteLayoutSettings: {
      footnoteSeparator: 'short', footnoteContinuationSeparator: 'full',
      nativeSeparators: { footnote: stories },
    },
  } as unknown as DocxDocumentModel;
  const services = createLayoutServices(model, {
    measureContext: sizedMeasureContext(), allowFootnoteContinuation: options.continuation === true,
  });
  return layoutDocument(model, services, { currentDateMs: 0 }).pages.map((page) =>
    page.layers.notes.filter((node): node is NoteLayout => node.kind === 'note'));
}

/** Ordinary text paragraph at an explicit size, optionally referencing a note. */
function sizedParagraph(text: string, fontSize: number, noteId?: string): DocParagraph {
  const result = ordinaryParagraph(text);
  result.defaultFontSize = fontSize;
  result.runs = [
    ...result.runs.map((run) => ({ ...run, fontSize })),
    ...(noteId ? [{ ...paragraph().runs[0]!, fontSize, noteRef: { kind: 'footnote' as const, id: noteId } }] : []),
  ] as DocParagraph['runs'];
  return result;
}

/** A long note continues from page 1 while a later body line on that page
 * references a short second note that still fits beneath it. */
function mixedNoteLayout(notice: unknown, options: Readonly<{ fillers?: number; longPt?: number }> = {}) {
  const model = {
    section: {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    },
    body: [
      sizedParagraph('a', 4, '7'),
      ...Array.from({ length: options.fillers ?? 0 }, () => sizedParagraph('b', 4)),
      sizedParagraph('c', 4, '8'),
    ],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [
      { id: '7', content: Array.from({ length: 30 }, (_, index) => sizedParagraph(`long-${index};`, options.longPt ?? 20)) },
      { id: '8', content: [sizedParagraph('short;', 4)] },
    ],
    endnotes: [], fontFamilyClasses: {},
    __noteLayoutSettings: {
      footnoteSeparator: 'short', footnoteContinuationSeparator: 'full',
      nativeSeparators: { footnote: {
        separator: nativeRule('short', 0), continuationSeparator: nativeRule('full', 3), continuationNotice: notice,
      } },
    },
  } as unknown as DocxDocumentModel;
  return { model, services: () => createLayoutServices(model, {
    measureContext: sizedMeasureContext(), allowFootnoteContinuation: true,
  }) };
}

describe('native note separator consumer', () => {
  const stories = (separator: unknown) => ({
    separator, continuationSeparator: nativeRule('full', 3), continuationNotice: nativeEmpty(6),
  });
  const leadingOf = (separator: unknown) => {
    const note = nativeNotePages(stories(separator))[0]![0]!;
    if (!note.leading) throw new Error('Expected a retained native separator occurrence');
    return { note, leading: note.leading };
  };

  it.each([
    // Fresh 180pt region: body 20 + native separator 10 + 13 lines of note 1
    // + the later note's 1-line minimum + ONE continuation notice 10 = 180.
    { pageHeight: 200, heads: [13, 1], bandPt: 160 },
    // Fresh 60pt region, exact fit: body 20 + separator 10 + one line of each
    // note + one notice = 60. Both notes then continue together on every
    // following page under the continuation separator and one notice.
    { pageHeight: 80, heads: [1, 1], bandPt: 40 },
  ])('reserves the single native notice once for a widow unit with two continuing references (page $pageHeight)', ({
    pageHeight, heads, bandPt,
  }) => {
    // A widow-controlled two-line paragraph references a 20-line note from
    // each line, so both lines are one admission unit. Exact 10pt lines.
    const exact = (text: string) => ({
      ...ordinaryParagraph(text), lineSpacing: { value: 10, rule: 'exact' as const, explicit: true },
    });
    const reference = (id: string) => ({ ...paragraph().runs[0]!, vertAlign: null, noteRef: { kind: 'footnote' as const, id } });
    const body = { ...exact('B01'), widowControl: true };
    body.runs = [
      reference('1'), ...body.runs, { type: 'break', breakType: 'line' }, reference('2'), ...exact('B02').runs,
    ] as DocParagraph['runs'];
    // Each note is one 20-line paragraph, so continuation resumes inside a
    // paragraph through its line cursor.
    const noteLines = (id: string) => {
      const note = exact(`N${id}-1;`);
      note.runs = Array.from({ length: 20 }, (_, index) => [
        ...(index > 0 ? [{ type: 'break', breakType: 'line' }] : []),
        ...exact(`N${id}-${index + 1};`).runs,
      ]).flat() as DocParagraph['runs'];
      return [note];
    };
    const lineCount = (note: NoteLayout) => note.story.blocks
      .reduce((sum, block) => sum + (block.kind === 'paragraph' ? block.lines.length : 0), 0);
    const notice = nativeNotice(6);
    const exactMark = { lineSpacing: { value: 10, rule: 'exact', explicit: true } };
    const model = {
      ...mixedNoteLayout(nativeEmpty(6)).model,
      section: { ...mixedNoteLayout(nativeEmpty(6)).model.section, pageHeight },
      body: [body],
      footnotes: [{ id: '1', content: noteLines('1') }, { id: '2', content: noteLines('2') }],
      __noteLayoutSettings: { nativeSeparators: { footnote: {
        separator: nativeRule('short', 0, { paragraph: exactMark }),
        continuationSeparator: nativeRule('full', 3, { paragraph: exactMark }),
        continuationNotice: { ...notice, paragraph: { ...notice.paragraph,
          paragraph: nativeParagraph(10, exactMark) } },
      } } },
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: sizedMeasureContext(), allowFootnoteContinuation: true,
    }), { currentDateMs: 0 });
    const pages = layout.pages.map((page) => page.layers.notes
      .filter((node): node is NoteLayout => node.kind === 'note'));
    expect(layout.pages[0]!.layers.body.flatMap((block) => block.kind === 'paragraph' ? block.lines : []))
      .toHaveLength(2);
    const first = pages[0]!;
    expect(first.map((note) => note.source.storyInstance)).toEqual(['1', '2']);
    expect(first.map(lineCount)).toEqual(heads);
    expect(first.map((note) => note.leading !== undefined)).toEqual([true, false]);
    expect(first.filter((note) => note.trailing)).toHaveLength(1);
    expect(first.reduce((sum, note) => sum + note.advancePt, 0)).toBe(bandPt);
    // Every source line of both notes is retained once across the pages, and
    // each page that still continues a note carries exactly one notice.
    for (const id of ['1', '2']) {
      const text = pages.flat().filter((note) => note.source.storyInstance === id)
        .flatMap((note) => note.story.blocks).flatMap((block) => block.kind === 'paragraph'
          ? block.lines.flatMap((line) => line.placements.flatMap((placement) =>
            placement.kind === 'text' ? [placement.text] : [])) : []).join('');
      expect(text).toBe(Array.from({ length: 20 }, (_, index) => `N${id}-${index + 1};`).join(''));
    }
    pages.forEach((notes, index) => {
      expect(notes.filter((note) => note.trailing)).toHaveLength(index + 1 < pages.length ? 1 : 0);
      expect(notes.filter((note) => note.leading)).toHaveLength(1);
    });
  });

  it('charges a native notice in the reference-line note plan', () => {
    // This is the native producer wire boundary, not a normal DOCX parser
    // fixture or an Office pagination oracle. Exact authored line heights
    // make separator, note, notice, and body ownership independently visible.
    const exact = (text: string, height = 10) => ({
      ...ordinaryParagraph(text), lineSpacing: { value: height, rule: 'exact' as const, explicit: true },
    });
    const body = exact('B01');
    body.runs = [
      { ...paragraph().runs[0]!, vertAlign: null }, ...body.runs,
      { type: 'break', breakType: 'line' }, ...exact('B02').runs,
    ] as DocParagraph['runs'];
    const layout = (remainingTwips: number, withNotice: boolean, continuation = true) => {
      const notice = nativeNotice(6);
      const measuredNotice = { ...notice, paragraph: { ...notice.paragraph,
        paragraph: nativeParagraph(10, { lineSpacing: { value: 10, rule: 'exact', explicit: true } }),
      } };
      const model = {
        ...mixedNoteLayout(nativeEmpty(6)).model,
        section: { ...mixedNoteLayout(nativeEmpty(6)).model.section, pageHeight: 160 + remainingTwips / 20 },
        body: [exact('P', 140), body],
        footnotes: [{ id: '7', content: Array.from({ length: 4 }, (_, index) => exact(`N${index + 1};`)) }],
        __noteLayoutSettings: { nativeSeparators: { footnote: {
          separator: nativeRule('short', 0, { paragraph: { lineSpacing: { value: 10, rule: 'exact', explicit: true } } }),
          continuationSeparator: nativeRule('full', 3, { paragraph: { lineSpacing: { value: 10, rule: 'exact', explicit: true } } }),
          continuationNotice: withNotice ? measuredNotice : nativeEmpty(6),
        } } },
      } as unknown as DocxDocumentModel;
      return layoutDocument(model, createLayoutServices(model, {
        measureContext: sizedMeasureContext(), allowFootnoteContinuation: continuation,
      }), { currentDateMs: 0 });
    };
    const notes = (value: ReturnType<typeof layout>) => value.pages.map(page => page.layers.notes
      .filter((node): node is NoteLayout => node.kind === 'note'));
    const bodyText = (value: ReturnType<typeof layout>) => value.pages.map(page => page.layers.body
      .flatMap(block => block.kind === 'paragraph' ? block.lines.flatMap(line => line.placements
        .flatMap(placement => placement.kind === 'text' ? [placement.text] : [])) : []).join(''));
    // 50pt remains after the 140pt line. The note plan is taken at the B01
    // reference line (capacity 40) and B02 is admitted only against it.
    const first = layout(1000, true);
    expect(bodyText(first)).toEqual(['P1B01', 'B02']);
    const note = notes(first)[0]![0]!;
    expect(note.advancePt).toBe(40); // separator10 + two note lines + notice10
    expect(note.story.blocks).toHaveLength(2);
    expect(note.trailing?.advancePt).toBe(10);
    expect(notes(first).at(-1)!.at(-1)!.trailing).toBeUndefined();
    const text = notes(first).flatMap(page => page.flatMap(note => note.story.blocks))
      .flatMap(block => block.kind === 'paragraph' ? block.lines.flatMap(line => line.placements
        .flatMap(placement => placement.kind === 'text' ? [placement.text] : [])) : []).join('');
    expect(text).toBe('N1;N2;N3;N4;');
    // One twip less: the notice-reserved partition keeps one real line.
    const lower = notes(layout(999, true))[0]![0]!;
    expect(lower.advancePt).toBe(30);
    expect(lower.story.blocks).toHaveLength(1);
    expect(lower.trailing?.advancePt).toBe(10);
    // Without a notice paragraph the same capacity holds three real lines.
    const withoutNotice = notes(layout(1000, false))[0]![0]!;
    expect(withoutNotice.story.blocks).toHaveLength(3);
    expect(withoutNotice.trailing).toBeUndefined();
    expect(layout(1000, true, false).pages[0]!.layers.notes).toEqual([]);
  });

  it('consumes control and content-mark CHPX and paragraph PAPX through the paragraph pipeline', () => {
    const { note, leading } = leadingOf(nativeRule('short', 0));
    // The native occurrence owns the rule; the modern scalar band is unused.
    expect(note.separator).toEqual([]);
    expect(note.advancePt).toBeCloseTo(leading.advancePt + note.story.advancePt);
    expect(note.story.flowBounds.yPt).toBeCloseTo(leading.flowBounds.yPt + leading.advancePt);
    expect(leading.source).toEqual({ story: 'footnote', storyInstance: 'reserved:separator', path: [] });
    // Both characters are zero-advance, inkless metric participants.
    expect(leading.paragraph?.lines.flatMap((line) => line.placements.map((placement) => placement.kind)))
      .toEqual(['anchor-host', 'anchor-host']);
    // Each participant's own size reaches the shared line box.
    expect(leadingOf(nativeRule('short', 0, { controlPt: 30 })).leading.advancePt)
      .toBeGreaterThan(leading.advancePt + 5);
    expect(leadingOf(nativeRule('short', 0, { markPt: 30 })).leading.advancePt)
      .toBeGreaterThan(leading.advancePt + 5);
    // Authored paragraph spacing is real flow, charged once.
    expect(leadingOf(nativeRule('short', 0, { paragraph: { spaceAfter: 12 } })).leading.advancePt)
      .toBeCloseTo(leading.advancePt + 12);
    // Indentation moves the paragraph's own placements; the rule keeps the
    // existing W/3 Short span from the main-story start.
    const indented = leadingOf(nativeRule('short', 0, { paragraph: { indentLeft: 20 } })).leading;
    expect(indented.paragraph?.lines[0]?.placements[0]?.bounds?.xPt)
      .toBeCloseTo((leading.paragraph?.lines[0]?.placements[0]?.bounds?.xPt ?? 0) + 20);
    expect(indented.rule?.segment).toEqual(leading.rule?.segment);
    expect(leading.rule).toMatchObject({
      mark: 'short',
      source: { storyInstance: 'reserved:separator', path: [0, 0] },
      segment: { color: '#000000', widthPt: 0.5 },
    });
    expect(leading.rule!.segment.to.xPt - leading.rule!.segment.from.xPt).toBeCloseTo(60);
    const line = leading.paragraph!.lines[0]!;
    expect(leading.rule!.segment.from.yPt).toBeCloseTo(line.bounds.yPt + line.bounds.heightPt / 2);
    // The control's own colour is the rule ink.
    expect(leadingOf(nativeRule('full', 0, { color: 'FF0000' })).leading.rule?.segment)
      .toMatchObject({ color: '#FF0000' });
  });

  it('owns one role-selected rule per page and a notice only on continuing fragments', () => {
    const render = (notice: unknown) => nativeNotePages({
      separator: nativeRule('short', 0),
      // An explicit Short continuation mark is kept independently of its role.
      continuationSeparator: nativeRule('short', 3),
      continuationNotice: notice,
    }, { notes: 8, continuation: true });
    const pages = render(nativeNotice(6));
    const notes = pages.flat();
    expect(notes.length).toBeGreaterThan(1);
    for (const page of pages) {
      expect(page.filter((note) => note.leading !== undefined).length).toBeLessThanOrEqual(1);
    }
    expect(notes[0]!.leading?.source.storyInstance).toBe('reserved:separator');
    expect(notes[1]!.leading?.source.storyInstance).toBe('reserved:continuation-separator');
    expect(notes[1]!.leading?.rule?.mark).toBe('short');
    // Source conservation across every fragment.
    expect(notes.flatMap((note) => note.story.blocks).filter((block) => block.kind === 'paragraph'))
      .toHaveLength(8);
    notes.forEach((note, index) => {
      const continues = index < notes.length - 1;
      expect(note.trailing !== undefined).toBe(continues);
      if (!note.trailing) return;
      expect(note.trailing.source.storyInstance).toBe('reserved:continuation-notice');
      expect(note.trailing.flowBounds.yPt + note.trailing.advancePt)
        .toBeCloseTo(note.flowBounds.yPt + note.advancePt);
      expect(note.advancePt).toBeCloseTo(
        (note.leading?.advancePt ?? 0) + note.story.advancePt + note.trailing.advancePt);
    });
    // The notice reserve can only shorten the admitted prefix.
    const admitted = (layout: NoteLayout[][]) => layout.flat()[0]!.story.blocks.length;
    expect(admitted(pages)).toBeLessThanOrEqual(admitted(render(nativeEmpty(6))));
    expect(render(nativeEmpty(6)).flat().every((note) => note.trailing === undefined)).toBe(true);
    // An intact note never carries a notice.
    const intact = nativeNotePages({ ...stories(nativeRule('short', 0)), continuationNotice: nativeNotice(6) },
      { continuation: true });
    expect(intact.flat()).toHaveLength(1);
    expect(intact[0]![0]!.trailing).toBeUndefined();
  });
  it('measures a visible paragraph-only notice through its own content-mark CHPX', () => {
    const notice = (mark: Record<string, unknown> | null, overrides: Record<string, unknown> = {}) => ({
      class: 'paragraphOnly', contentStartCp: 6, contentEndCp: 7, guardCp: 7,
      paragraph: {
        paragraph: nativeParagraph(10, overrides),
        contentMark: mark ? { run: { ...nativeRun(10), ...mark } } : {},
        source: nativeAddress(6),
      },
    });
    const trailingOf = (continuationNotice: unknown) => nativeNotePages({
      separator: nativeRule('short', 0), continuationSeparator: nativeRule('full', 3), continuationNotice,
    }, { notes: 8, continuation: true }).flat()[0]!.trailing;
    const plain = trailingOf(notice({}));
    if (!plain) throw new Error('Expected a reserved continuation notice');
    // Existing library line-box composition for this controlled fixture (not
    // an ECMA-376 §17.3.2.24 or Office formula): the raise adds to the mark's
    // own selected-face ascent instead of replacing it.
    expect(trailingOf(notice({ position: 6 }))?.advancePt).toBeCloseTo(plain.advancePt + 6);
    // §17.3.2.42 superscript reaches the mark through the existing
    // effective-size policy; no coefficient is introduced here.
    expect(trailingOf(notice({ vertAlign: 'super' }))?.advancePt).not.toBeCloseTo(plain.advancePt);
    // The retained participant carries its selected-face sides, not zeros.
    expect(plain.paragraph!.lines[0]!.placements[0]).toMatchObject({
      sourceMetrics: { ascentPt: 8, descentPt: 2 },
    });
    // The mark stays text-free and zero-advance.
    const placements = plain.paragraph!.lines.flatMap((line) => line.placements);
    expect(placements.filter((placement) => placement.kind === 'text')).toEqual([]);
    expect(placements.every((placement) => placement.bounds?.widthPt === 0)).toBe(true);
    // A vanished mark and a guard-only story own no notice occurrence.
    expect(trailingOf(notice(null, { markVanish: true }))).toBeUndefined();
    expect(trailingOf({ class: 'guardOnly', contentStartCp: 6, contentEndCp: 6, guardCp: 6 })).toBeUndefined();
  });

  it('composes a raised rule control or mark with the other participant unmasked', () => {
    const raised = (control: Record<string, unknown>, mark: Record<string, unknown>) => {
      const story = nativeRule('short', 0, { controlPt: 14, markPt: 9 }) as {
        rule: { control: { run: Record<string, unknown> } };
        paragraph: { contentMark: { run: Record<string, unknown> } };
      };
      Object.assign(story.rule.control.run, control);
      Object.assign(story.paragraph.contentMark.run, mark);
      return leadingOf(story).leading;
    };
    const plain = raised({}, {});
    // Existing library composition: the 14pt control owns the line box, so a
    // 6pt raise on it adds 6pt; the same raise on the 9pt mark adds only its
    // excess over the control's ascent.
    expect(raised({ position: 6 }, {}).advancePt).toBeCloseTo(plain.advancePt + 6);
    expect(raised({}, { position: 6 }).advancePt).toBeCloseTo(plain.advancePt + 2);
    // One rule ink, no text placements, zero inline extent.
    const leading = raised({ position: 6 }, {});
    expect(leading.rule?.segment.widthPt).toBe(0.5);
    const placements = leading.paragraph!.lines.flatMap((line) => line.placements);
    expect(placements.every((placement) => placement.kind === 'anchor-host'
      && placement.bounds?.widthPt === 0)).toBe(true);
  });

  it('acquires a small-caps rule control through its transformed probe at full metric size', () => {
    const story = nativeRule('short', 0) as { rule: { control: { run: Record<string, unknown> } } };
    story.rule.control.run.smallCaps = true;
    // Small caps transform the probe's display text; acquisition must still
    // succeed and keep the existing full-size metric policy.
    expect(leadingOf(story).leading.advancePt).toBeCloseTo(leadingOf(nativeRule('short', 0)).leading.advancePt);
  });

  it('keeps one band-final notice when a later short note joins a continuing page', async () => {
    const notes = (layout: ReturnType<typeof layoutDocument>) => layout.pages.map((page) =>
      page.layers.notes.filter((node): node is NoteLayout => node.kind === 'note'));
    const { model, services } = mixedNoteLayout(nativeNotice(6));
    const layout = layoutDocument(model, services(), { currentDateMs: 0 });
    const pages = notes(layout);
    const first = pages[0]!;
    // Both references stay on page 1, in reference order.
    expect(first.map((note) => note.source.storyInstance)).toEqual(['7', '8']);
    const [continuing, later] = first as [NoteLayout, NoteLayout];
    // §17.18.33: the one notice ends the band and is charged exactly once.
    expect(continuing.trailing).toBeUndefined();
    expect(continuing.advancePt).toBeCloseTo((continuing.leading?.advancePt ?? 0) + continuing.story.advancePt);
    expect(later.trailing?.role).toBe('continuationNotice');
    expect(later.advancePt).toBeCloseTo(later.story.advancePt + later.trailing!.advancePt);
    expect(later.trailing!.flowBounds.yPt + later.trailing!.advancePt)
      .toBeCloseTo(later.flowBounds.yPt + later.advancePt);
    expect(later.trailing!.paragraph?.flowDomainId).toBe(later.flowDomainId);
    expect(first.flatMap((note) => note.trailing ? [note] : [])).toHaveLength(1);
    // Every source row of the continued note is retained once, in order.
    const text = pages.flat().filter((note) => note.source.storyInstance === '7')
      .flatMap((note) => note.story.blocks).flatMap((block) => block.kind === 'paragraph'
        ? block.lines.flatMap((line) => line.placements.flatMap((placement) =>
          placement.kind === 'text' ? [placement.text] : [])) : []).join('');
    expect(text).toBe(Array.from({ length: 30 }, (_, index) => `long-${index};`).join(''));
    expect(pages.flat().filter((note) => note.source.storyInstance === '8')).toHaveLength(1);
    // Each later page that still continues carries its own single notice.
    pages.slice(1).forEach((page, index) => {
      expect(page.filter((note) => note.trailing !== undefined))
        .toHaveLength(index + 2 < pages.length ? 1 : 0);
    });
    // The sliced driver retains the same ownership.
    const slicedSource = layoutSourceStore(mixedNoteLayout(nativeNotice(6)).model);
    const sliced = await layoutDocumentInputAsync(slicedSource.bodyLayoutInput,
      createLayoutServices(slicedSource, {
        measureContext: sizedMeasureContext(), allowFootnoteContinuation: true,
      }), { currentDateMs: 0 }, { sliceMs: 0, yieldToHost: async () => undefined });
    expect(sliced).toEqual(layout);
    // Without a notice paragraph nothing is reserved or painted for it; both
    // notes are still retained once.
    const empty = mixedNoteLayout(nativeEmpty(6));
    const plain = notes(layoutDocument(empty.model, empty.services(), { currentDateMs: 0 })).flat();
    expect(plain.every((note) => note.trailing === undefined)).toBe(true);
    expect(plain.filter((note) => note.source.storyInstance === '8')).toHaveLength(1);
  });
});


describe('multiple pending footnote continuations', () => {
  it('preserves deferred cursors when an earlier continuation fills the page band', async () => {
    const base = mixedNoteLayout(nativeNotice(6));
    const model = {
      ...base.model,
      footnotes: base.model.footnotes!.map((note) => note.id === '8'
        ? { ...note, content: Array.from({ length: 30 }, (_, index) => sizedParagraph(`second-${index};`, 4)) }
        : note),
    } as DocxDocumentModel;
    const services = () => createLayoutServices(model, {
      measureContext: sizedMeasureContext(), allowFootnoteContinuation: true,
    });
    const layout = layoutDocument(model, services(), { currentDateMs: 0 });
    const pages = layout.pages.map((page) => page.layers.notes
      .filter((node): node is NoteLayout => node.kind === 'note'));
    // Both notes start with their references before their cursors are pending.
    expect(pages[0]!.map((note) => note.source.storyInstance)).toEqual(['7', '8']);
    for (const [id, prefix] of [['7', 'long'], ['8', 'second']]) {
      const text = pages.flat().filter((note) => note.source.storyInstance === id)
        .flatMap((note) => note.story.blocks).flatMap((block) => block.kind === 'paragraph'
          ? block.lines.flatMap((line) => line.placements.flatMap((placement) =>
            placement.kind === 'text' ? [placement.text] : [])) : []).join('');
      expect(text).toBe(Array.from({ length: 30 }, (_, index) => `${prefix}-${index};`).join(''));
    }
    // Every continuing page charges exactly one band-final notice. A deferred
    // note contributes no empty fragment or duplicate leading/notice reserve.
    pages.slice(0, -1).forEach((notes) => {
      expect(notes.filter((note) => note.trailing)).toHaveLength(1);
      expect(notes.at(-1)!.trailing).toBeDefined();
    });
    expect(pages.at(-1)!.every((note) => !note.trailing)).toBe(true);
    expect(pages.flat().every((note) => note.story.blocks.length > 0)).toBe(true);
    const source = layoutSourceStore(model);
    const sliced = await layoutDocumentInputAsync(source.bodyLayoutInput, createLayoutServices(source, {
      measureContext: sizedMeasureContext(), allowFootnoteContinuation: true,
    }), { currentDateMs: 0 }, { sliceMs: 0, yieldToHost: async () => undefined });
    expect(sliced).toEqual(layout);
  });
});
