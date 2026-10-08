import { measureJoinedTextUnit } from './line-breaker/atomic-units.js';
import { it, expect } from 'vitest';
import { createPrefixWorkBudget } from './line-breaker/prefix-work-budget.js';
import { performExplicitTextSplit, performQueueEmergencyTail, performEmergencyTextSplit, performProspectiveSnapAdvance } from './line-breaker/pass-operations.js';
import { textBreakWindow } from './line-breaker/text-break-window.js';
it('bounds actual requested prefix work across dense ordinary hyphen queue tails', () => {
  for (const n of [100, 1600]) {
    let calls = 0, units = 0, lines = 0, longestRequest = 0;
    const state: any = { reservePrefixWork: createPrefixWorkBudget(), scale: 1, strNaturalAdvance: (_segment: any, text: string) => {
      calls++; units += text.length; longestRequest = Math.max(longestRequest, text.length); return text.length;
    }, prospectiveSnapAdvance: (_s: any, width: number) => width, breakerState: { queue: [] } };
    let segment: any = { src: { segIndex: 0, charOffset: 0 }, text: 'a-'.repeat(n) + 'a',
      explicitBreaks: textBreakWindow(Object.freeze(Array.from({ length: n }, (_, i) => 2 * (i + 1)))) };
    while (segment.text.length > 10) {
      const split = performExplicitTextSplit(state, segment, 10);
      expect(split).toBe(10);
      performQueueEmergencyTail(state, segment, split);
      segment = state.breakerState.queue.shift(); lines++;
    }
    expect(lines).toBe(n / 5);
    expect(longestRequest).toBeLessThanOrEqual(16);
    expect(calls).toBeLessThanOrEqual(lines * 6);
    expect(units).toBeLessThanOrEqual(lines * 52);
  }
});
it('preserves a later fitting opportunity after a signed-spacing overflow', () => {
  const state: any = { reservePrefixWork: createPrefixWorkBudget(), scale: 1, strNaturalAdvance: (_s: any, text: string) => ({ 2: 8, 4: 12, 6: 4, 8: 13 }[text.length]),
    prospectiveSnapAdvance: (_s: any, width: number) => width };
  expect(performExplicitTextSplit(state, { text: 'a-a-a-a-a', charSpacing: -1,
    explicitBreaks: textBreakWindow([2, 4, 6, 8]) } as any, 10)).toBe(6);
});

it('uses exact active Latin block rounding without exhaustive semantic or emergency fitting', () => {
  let calls = 0, units = 0;
  const state: any = { reservePrefixWork: createPrefixWorkBudget(), scale: 1,
    characterGrid: { type: 'snapToChars', characterPitchPt: 4 }, snapPitchPx: 4,
    breakerState: { snapBlock: { kind: 'latin', naturalWidthPx: 3, allocatedWidthPx: 4 } },
    strNaturalAdvance: (_s: any, text: string) => { calls++; units += text.length; return text.length; },
  };
  state.prospectiveSnapAdvance = (s: any, natural: number) => performProspectiveSnapAdvance(state, s, natural);
  const seg: any = { text: 'a-'.repeat(1600) + 'a', explicitBreaks: textBreakWindow(
    Array.from({ length: 1600 }, (_, i) => 2 * (i + 1))) };
  // ceil((3 + 8) / 4) * 4 - 4 = 8, whereas the next hyphen is12.
  expect(performExplicitTextSplit(state, seg, 10)).toBe(8);
  expect(calls).toBeLessThanOrEqual(6);
  expect(units).toBeLessThanOrEqual(52);
  calls = units = 0;
  // Character9 also fits the actual block; character10 requires12pt.
  expect(performEmergencyTextSplit(state, { text: 'a'.repeat(3201) } as any, 10)).toBe(9);
  expect(calls).toBeLessThanOrEqual(9);
  expect(units).toBeLessThanOrEqual(100);
});
it('shares a pre-shaping work quota across semantic and emergency searches and isolates later passes', () => {
  let requested = 0;
  const state: any = { reservePrefixWork: createPrefixWorkBudget(11), scale: 1,
    strNaturalAdvance: (_s: any, text: string) => { requested += text.length; return text.length; },
    prospectiveSnapAdvance: (_s: any, natural: number) => natural,
  };
  expect(performExplicitTextSplit(state, { text: 'a-a-a', charSpacing: -1,
    explicitBreaks: textBreakWindow([2, 4]) } as any, 10)).toBe(4);
  expect(() => performEmergencyTextSplit(state, { text: 'abc', charSpacing: -1 } as any, 1))
    .toThrow(RangeError);
  expect(requested).toBe(7); // The rejected two-unit request never entered shaping.
  state.reservePrefixWork = createPrefixWorkBudget(11);
  expect(performEmergencyTextSplit(state, { text: 'abc', charSpacing: -1 } as any, 1)).toBe(1);
  expect(requested).toBe(13);
});
it('fails a dense negative-pitch chain before unbounded queued-prefix shaping', () => {
  let requested = 0;
  const state: any = { reservePrefixWork: createPrefixWorkBudget(), scale: 1,
    strNaturalAdvance: (_s: any, text: string) => { requested += text.length; return text.length; },
    prospectiveSnapAdvance: (_s: any, natural: number) => natural,
    breakerState: { queue: [] },
  };
  let seg: any = { text: 'a-'.repeat(1600) + 'a', charSpacing: -1, src: { segIndex: 0, charOffset: 0 },
    explicitBreaks: textBreakWindow(Array.from({ length: 1600 }, (_, i) => 2 * (i + 1))) };
  expect(() => { while (seg.text.length > 10) {
    const split = performExplicitTextSplit(state, seg, 10);
    performQueueEmergencyTail(state, seg, split);
    seg = state.breakerState.queue.shift();
  } }).toThrow(RangeError);
  expect(requested).toBeLessThanOrEqual(16 * 1024 * 1024);
});

it('charges full hard-follower grapheme acquisition before choosing a short admitted prefix', () => {
  const measurement = { baseRtl: false, reservePrefixWork: createPrefixWorkBudget(60),
    segAdvance: (s: any) => s.text.length, strAdvance: (_s: any, text: string) => text.length };
  expect(() => measureJoinedTextUnit({ text: 'a-a', explicitBreaks: textBreakWindow([2]) } as any,
    [{ text: 'b'.repeat(100), joinPrev: true, hardJoinPrev: true } as any], measurement,
    3, 0, 0, false, 'whole-leader')).toThrow(RangeError);
});
