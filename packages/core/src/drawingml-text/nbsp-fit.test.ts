import { expect, it } from 'vitest';
import { breakDrawingMlText } from './break.js';

it('keeps an NBSP word pair together across a table-cell fit boundary (T00/T01/T03)', () => {
  // Office-produced PowerPoint controls T00–T08 establish that NBSP binds its
  // adjacent words at the fit boundary; ordinary space does not (T03–T05).
  // In a separate Segoe UI table cell, usable width is 90.792 px: Office's
  // bound phrase is about 89.7 px, while the browser substitute is 91.787 px.
  // The previous renderer only matched Office there by incorrectly splitting
  // at NBSP. With the Office break rule, the substitute must move the bound
  // phrase to the next line; that remaining visual gap is a font-metric issue.
  // A wider initial word makes the NBSP group fit on the continuation line
  // at 24 units, but the entire first phrase only fits at 25 units.
  const advance = (text: string): number =>
    [...text].reduce((sum, ch) => sum + (ch === 'C' ? 6 : 1), 0);
  const wrap = (text: string, maxWidth: number): string[] =>
    breakDrawingMlText([{ type: 'text', text, style: 'same' }], {
      maxWidth, measureText: advance,
    }).map((line) => line.segments.map((segment) => segment.type === 'text' ? segment.text : '').join('').trimEnd());
  expect(wrap('Customer and\u00a0Partner Success', 24)).toEqual(['Customer', 'and\u00a0Partner Success']);
  expect(wrap('Customer and\u00a0Partner Success', 25)).toEqual(['Customer and\u00a0Partner', 'Success']);
  expect(wrap('Customer and Partner Success', 24)).toEqual(['Customer and', 'Partner Success']);
});
