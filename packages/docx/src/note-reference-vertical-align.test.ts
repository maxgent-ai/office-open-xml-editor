import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { BodyElement, DocNote, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';
import type { LineLayout, NoteLayout, ParagraphLayout } from './layout/types.js';

// ECMA-376 §17.3.2.42 (w:vertAlign) and §17.3.2.28 (direct w:rPr): a note
// reference mark (§17.11.14 footnoteReference, §17.11.7 endnoteReference) or
// placeholder (§17.11.13 footnoteRef, §17.11.6 endnoteRef) carries the run's
// effective vertical alignment; superscript comes from that formatting, never
// from the mark element. Synthetic glyph metrics (ascent 0.8em, descent
// 0.2em, advance 0.5em per character) verify the consumer contract, not Office
// glyph fidelity.

const FONT = 'Times New Roman';

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px', fontKerning: 'none',
    measureText: (text: string) => {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px * 0.5,
        fontBoundingBoxAscent: px * 0.8, fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8, actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function textRun(text: string, fontSize: number, extra: Record<string, unknown> = {}) {
  return {
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: FONT, fontFamilyEastAsia: '', isLink: false,
    background: null, hyperlink: null, ...extra,
  };
}

function para(runs: Array<Record<string, unknown>>): BodyElement {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs as DocParagraph['runs'], defaultFontSize: 5, defaultFontFamily: FONT,
    widowControl: false,
  } as unknown as BodyElement;
}

const note = (id: string): DocNote => ({ id, content: [para([textRun(`note ${id}`, 5)])] });

/** Each body paragraph pairs a small 5pt run with one 20pt subject run, so
 * the subject's effective formatting owns the line's ascent. */
function layoutSubjects(subjects: Array<Record<string, unknown>>) {
  const model = {
    section: {
      pageWidth: 400, pageHeight: 400, marginTop: 10, marginRight: 10, marginBottom: 10,
      marginLeft: 10, headerDistance: 4, footerDistance: 4, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage',
    } as SectionProps,
    body: subjects.map(subject => para([textRun('a', 5), subject])),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [FONT]: 'roman' },
    footnotes: ['plain', 'super', 'custom'].map(note),
  } as unknown as DocxDocumentModel;
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }),
    { currentDateMs: 0 });
  const lines = layout.pages[0]!.layers.body.map(block => (block as ParagraphLayout).lines[0]!);
  const notes = layout.pages[0]!.layers.notes
    .filter((node): node is NoteLayout => node.kind === 'note')
    .map(noteLayout => noteLayout.source.storyInstance);
  return { lines, notes };
}

const ascent = (line: LineLayout) => line.baselinePt - line.bounds.yPt;

describe('note reference marks keep their run\'s effective vertical alignment', () => {
  const plain = textRun('x', 20, { vertAlign: null });
  const visible = textRun('1', 20, { vertAlign: null, noteRef: { kind: 'footnote', id: 'plain' } });
  const superscript = textRun('2', 20, { vertAlign: 'super', noteRef: { kind: 'footnote', id: 'super' } });
  const custom = textRun('', 20, {
    vertAlign: null, noteRef: { kind: 'footnote', id: 'custom', customMarkFollows: true },
  });
  const { lines, notes } = layoutSubjects([plain, visible, superscript, custom]);
  const [plainLine, visibleLine, superLine, customLine] = lines as [LineLayout, LineLayout, LineLayout, LineLayout];
  const subject = (line: LineLayout) => line.placements.at(-1)!;

  it('prints an unstyled visible marker full size on the baseline', () => {
    const mark = subject(visibleLine);
    expect(mark.kind).toBe('text');
    if (mark.kind !== 'text') return;
    expect(mark.text).toBe('1');
    expect(mark.noteReference).toEqual({ kind: 'footnote', id: 'plain' });
    expect(mark.fontSizePt).toBe(20);
    expect(mark.verticalAlign).toBeUndefined();
    expect(mark.origin.yPt).toBe(visibleLine.baselinePt);
    expect(mark.paintOps[0]!.offset.yPt).toBe(0);
    expect(mark.advancePt).toBe(10);
    // Same line box as ordinary 20pt text with identical formatting.
    expect(ascent(visibleLine)).toBe(ascent(plainLine));
    expect(visibleLine.advancePt).toBe(plainLine.advancePt);
  });

  it('keeps explicitly superscript formatting on a marker', () => {
    const mark = subject(superLine);
    expect(mark.kind).toBe('text');
    if (mark.kind !== 'text') return;
    expect(mark.verticalAlign).toBe('super');
    expect(mark.fontSizePt).toBeLessThan(20);
    expect(mark.paintOps[0]!.offset.yPt).toBeLessThan(0);
  });

  it('gives a custom-mark zero-ink host the run\'s effective metrics and reference ownership', () => {
    const host = subject(customLine);
    expect(host.kind).toBe('anchor-host');
    if (host.kind !== 'anchor-host') return;
    expect(host.noteReference).toEqual({ kind: 'footnote', id: 'custom' });
    expect(host.bounds.widthPt).toBe(0);
    // A baseline 20pt host owns the same line box as visible 20pt text. Its
    // empty-probe baseline split is a separate, pre-existing metric contract.
    expect(customLine.bounds.heightPt).toBe(plainLine.bounds.heightPt);
    expect(customLine.advancePt).toBe(plainLine.advancePt);
  });

  it('keeps every referenced note on the reference page', () => {
    expect(notes).toEqual(['plain', 'super', 'custom']);
  });
});
