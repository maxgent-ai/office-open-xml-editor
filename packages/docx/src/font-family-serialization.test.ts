import { expect, it } from 'vitest';
import { buildFont, quoteAll } from './line-breaker/font-routes.js';

it('preserves each named DOCX fallback family in its CSS string', () => {
  const families = ['Family "Quoted"', 'Family \\31'];
  const quoted = quoteAll(families).match(/"(?:[^"\\]|\\.)*"/gu) ?? [];
  expect(quoted.map((token) => JSON.parse(token) as string)).toEqual(families);
});

it('preserves a literal backslash in the primary DOCX Canvas family', () => {
  const font = buildFont(false, false, 20, 'Family \\31');
  const primary = font.match(/"(?:[^"\\]|\\.)*"/u)?.[0] as string;
  expect(JSON.parse(primary)).toBe('Family \\31');
});
