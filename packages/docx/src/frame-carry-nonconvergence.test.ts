import { describe, expect, it, vi } from 'vitest';

// WORD_TEXT_FRAME_PRECEDING_TEXT_PUSH: a frame-carry convergence failure is
// reported, never answered with a pass that lacks a required push.
vi.mock('./layout/convergence.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./layout/convergence.js')>();
  return {
    ...actual,
    // eslint-disable-next-line require-yield
    convergeExactStateSteps: function* () {
      throw new actual.ExactConvergenceError('limit', [], 1);
    },
  };
});

const { layoutBodyModel } = await import('./test-support/document-layout.test-support.js');

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => ({
      width: [...s].length * px(), fontBoundingBoxAscent: px() * 0.8, fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.8, actualBoundingBoxDescent: px() * 0.2,
    }),
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fillRect() {}, drawImage() {}, fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left', direction: 'ltr',
  } as unknown as CanvasRenderingContext2D;
}

const para = (text: string, extra: Record<string, unknown> = {}) => ({
  type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
  spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
  runs: [{ type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null }],
  defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', ...extra,
});

describe('text-frame carry convergence', () => {
  it('reports non-convergence instead of dropping a required push', () => {
    const framePr = {
      dropCap: 'none', lines: 1, wrap: 'notBeside', hAnchor: 'page', vAnchor: 'text',
      hRule: 'auto', hSpace: 0, vSpace: 10, w: 100, xAlign: 'center',
    };
    const body = [para('P'), para('F1', { framePr }), para('A')];
    const section = {
      pageWidth: 200, pageHeight: 400, marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    };
    expect(() => layoutBodyModel(body as never, section as never, makeCtx())).toThrow(/NON_CONVERGENCE/);
  });
});
