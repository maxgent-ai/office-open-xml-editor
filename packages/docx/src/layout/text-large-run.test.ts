import { expect, it } from 'vitest';
import { createFontResolver } from './font-service.js';
import { createTextLayoutService } from './text.js';

it('retains metrics for a long run with distinct scalar font slots', () => {
  // A valid multi-seam run must keep its ordinary slots rather than merge
  // faces merely to stay below the JavaScript call-argument limit.
  const text = 'Té'.repeat(65_536);
  const service = createTextLayoutService({
    fonts: createFontResolver([
      { requestedFamily: 'Ascii Face', resolvedFamily: 'Ascii Face', source: 'local' },
      { requestedFamily: 'Latin Face', resolvedFamily: 'Latin Face', source: 'local' },
    ]),
    measurer: {
      fingerprint: 'public-large-slot-run',
      measure: ({ text }) => ({
        advancePt: text.length * 2,
        ascentPt: text === 'é' ? 13 : 9,
        descentPt: text === 'é' ? 4 : 1,
      }),
    },
  });
  const shaped = service.shape({
    text, fontSizePt: 10,
    fonts: { ascii: 'Ascii Face', highAnsi: 'Latin Face' },
    clusterGeometry: false,
  });
  expect(shaped.advancePt).toBe(text.length * 2);
  expect([shaped.ascentPt, shaped.descentPt]).toEqual([13, 4]);
  expect(shaped.spans).toHaveLength(text.length);
  expect(shaped.spans.slice(0, 2).map(span => span.script)).toEqual(['ascii', 'highAnsi']);
  expect(shaped.spans.at(-1)?.end).toBe(text.length);
});
