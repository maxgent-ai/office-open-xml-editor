import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { PaintNode } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxTextRun, FramePr, SectionProps } from './types';

// WORD_TEXT_FRAME_PRECEDING_TEXT_PUSH: consecutive text-anchored frames are
// placed from the content before them, push that content below their
// exclusions, and move with their anchor paragraph. The stub canvas makes a
// line exactly fontPx tall (0.8/0.2 em font box) and a glyph fontPx wide.

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// Content band 160 wide, from y=20 to pageHeight-20.
function section(pageHeight: number): SectionProps {
  return {
    pageWidth: 200, pageHeight,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  };
}

function textRun(text: string): DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocParagraph['runs'][number];
}

function para(text: string, extra: Partial<DocParagraph> = {}): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [textRun(text)], defaultFontSize: 20, defaultFontFamily: 'NotInMetrics',
    ...extra,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

/** A notBeside frame centred on the page, anchored to the text. */
function frame(w: number, vSpace: number, extra: Partial<FramePr> = {}): Partial<DocParagraph> {
  return {
    framePr: {
      dropCap: 'none', lines: 1, wrap: 'notBeside', hAnchor: 'page', vAnchor: 'text',
      hRule: 'auto', hSpace: 0, vSpace, w, xAlign: 'center', ...extra,
    },
  };
}

const textOf = (node: PaintNode): string => node.kind === 'paragraph'
  ? node.lines.flatMap((line) => line.placements)
      .filter((placement) => placement.kind === 'text')
      .map((placement) => placement.text).join('')
  : '';

function topsByText(body: readonly BodyElement[], pageHeight = 400) {
  const pages = layoutBodyModel(body, section(pageHeight), makeCtx()).pages;
  return pages.map((page) => Object.fromEntries(page.layers.body
    .filter((node) => node.kind === 'paragraph')
    .map((node) => [textOf(node), node.kind === 'paragraph' ? node.lines[0]!.bounds.yPt : 0])));
}

describe('text-anchored frames and the content before them', () => {
  it('pushes the preceding paragraph below each frame in turn and keeps each pusher in place', () => {
    // P (20..40); F1 is placed at P's bottom (40) and its 10pt vSpace meets P,
    // so P moves below F1's exclusion (70); F2 is then placed from P's new
    // bottom (90) and pushes P again (120). The anchor follows P.
    const [page] = topsByText([
      para('P'),
      para('F1', frame(100, 10)),
      para('F2', frame(80, 10)),
      para('A'),
    ]);
    expect(page).toEqual({ F1: 40, F2: 90, P: 120, A: 140 });
  });

  it('moves an earlier frame that did not push with the preceding paragraph', () => {
    // F1 has no vSpace and does not meet P; F2 does. F1 is not pinned: it is
    // re-placed at P's new bottom, overlapping F2's position untouched.
    const [page] = topsByText([
      para('P'),
      para('F1', frame(100, 0)),
      para('F2', frame(80, 10)),
      para('A'),
    ]);
    expect(page).toEqual({ F2: 40, P: 70, F1: 90, A: 110 });
  });

  it('leaves frames that meet nothing at one shared base, overlapping', () => {
    const [page] = topsByText([
      para('P'),
      para('F1', frame(100, 0)),
      para('F2', frame(80, 0)),
      para('A'),
    ]);
    expect(page).toEqual({ P: 20, F1: 40, F2: 40, A: 60 });
  });

  it('collapses the anchor spacing against the content before the frames', () => {
    // A narrow wrap-around frame leaves room beside it, so the anchor is
    // placed by its own space before after P, not against the frame
    // paragraph's space after.
    const [page] = topsByText([
      para('P'),
      para('F1', { ...frame(40, 0, { wrap: 'around', xAlign: 'right' }), spaceAfter: 30 }),
      para('A', { spaceBefore: 5 }),
    ]);
    expect(page).toEqual({ P: 20, F1: 40, A: 45 });
  });

  it('collapses the anchor spacing past every member of a multi-paragraph frame', () => {
    const member = { ...frame(40, 0, { wrap: 'around', xAlign: 'right' }) };
    const [page] = topsByText([
      para('P'),
      para('F1', member),
      para('F2', { ...member, spaceAfter: 30 }),
      para('A', { spaceBefore: 5 }),
    ]);
    expect(page).toEqual({ P: 20, F1: 40, F2: 60, A: 45 });
  });

  it('preserves preceding spacing when frame members have space before', () => {
    const member = { ...frame(40, 0, { wrap: 'around', xAlign: 'right' }), spaceBefore: 20 };
    const [page] = topsByText([
      para('P', { spaceAfter: 10 }),
      para('F1', member),
      para('F2', member),
      para('A', { spaceBefore: 20 }),
    ]);
    expect(page).toEqual({ P: 20, F1: 70, F2: 110, A: 60 });
  });

  it('moves the frames to the page their anchor paragraph starts on', () => {
    // Content band 20..120. Four lines fill 20..100; F1 at 100 fits, but its
    // anchor cannot start below the exclusion (110 + 10 vSpace), so the frame
    // run opens the next page with its anchor.
    const pages = topsByText([
      para('a'), para('b'), para('c'), para('d'),
      para('F1', frame(100, 10)),
      para('A'),
    ], 140);
    expect(pages).toEqual([
      { a: 20, b: 40, c: 60, d: 80 },
      { F1: 20, A: 50 },
    ]);
  });
});
