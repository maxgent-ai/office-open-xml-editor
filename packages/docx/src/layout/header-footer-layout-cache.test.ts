import { expect, it } from 'vitest';
import { createHeaderFooterLayoutCache, MAX_HEADER_FOOTER_CACHE_ENTRIES,
  MAX_HEADER_FOOTER_CACHE_UNITS } from './header-footer-layout-cache.js';
import type { StoryLayout } from './types.js';

function story(): StoryLayout {
  const bounds = { xPt: 0, yPt: 0, widthPt: 100, heightPt: 0 };
  return { story: 'header', flowBounds: bounds, inkBounds: bounds, advancePt: 0, blocks: [], diagnostics: [] };
}
const request = (occurrence: string, placement = 'a') => ({ occurrence, placement, trial: true });

it('evicts least-recent placements and occurrences under independent entry and retained-size ceilings', () => {
  let acquisitions = 0;
  const acquire = () => { acquisitions++; return story(); };
  const cache = createHeaderFooterLayoutCache(() => undefined);
  const a = cache.layout(request('a'), acquire);
  cache.layout(request('a', 'b'), acquire);
  expect(cache.layout(request('a'), acquire)).toBe(a);
  cache.layout(request('a', 'c'), acquire);
  const before = acquisitions;
  cache.layout(request('a', 'b'), acquire);
  expect(acquisitions).toBe(before + 1);
  for (let i = 0; i < MAX_HEADER_FOOTER_CACHE_ENTRIES; i++) cache.layout(request(String(i)), acquire);
  const full = acquisitions;
  cache.layout(request('a'), acquire);
  expect(acquisitions).toBe(full + 1);

  // One over-sized key must not evict a useful entry or become retained itself.
  const recent = cache.layout(request('recent'), acquire);
  const oversized = request('large', 'x'.repeat(MAX_HEADER_FOOTER_CACHE_UNITS));
  cache.layout(oversized, acquire);
  const after = acquisitions;
  cache.layout(oversized, acquire);
  expect(acquisitions).toBe(after + 1);
  expect(cache.layout(request('recent'), acquire)).toBe(recent);

  // The size ceiling also covers acquired graphs, independently of entry count.
  const largeStory = () => ({ ...story(), diagnostics: [{ code: 'UNSUPPORTED_FEATURE',
    severity: 'warning', message: 'x'.repeat(MAX_HEADER_FOOTER_CACHE_UNITS / 2),
  }] }) as StoryLayout;
  cache.layout(request('heavy:1'), largeStory);
  cache.layout(request('heavy:2'), largeStory);
  const sizeBound = acquisitions;
  cache.layout(request('recent'), acquire);
  expect(acquisitions).toBe(sizeBound + 1);
});

it('charges before trial acquisition and retains no failed acquisition', () => {
  let charged = 0;
  let acquired = 0;
  const cache = createHeaderFooterLayoutCache(() => { if (++charged > 2) throw new Error('spent'); });
  expect(() => cache.layout(request('a'), () => { acquired++; throw new Error('layout failed'); })).toThrow('layout failed');
  cache.layout(request('a'), () => { acquired++; return story(); });
  expect(() => cache.layout(request('b'), () => { acquired++; return story(); })).toThrow('spent');
  expect(acquired).toBe(2);
});
