import { expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types.js';

function textRun(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily: 'Arial',
  };
}

it('starts and ends a visual line around a display equation in a PPTX paragraph', () => {
  const context = {
    font: '', measureText: (text: string) => ({ width: text.length * 10 }),
  } as unknown as CanvasRenderingContext2D;
  const paragraph = {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [],
    eaLnBrk: true,
    runs: [textRun('before'), { type: 'math', nodes: [], display: true }, textRun('after')],
  } as Paragraph;

  const lines = layoutParagraph(context, paragraph, 200, 20, '000000', 1, 0);
  expect(lines.map((line) => line.segments.map((segment) =>
    segment.math?.display ? '[equation]' : segment.text).join('')))
    .toEqual(['before', '[equation]', 'after']);
});
