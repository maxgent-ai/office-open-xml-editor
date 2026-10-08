import { describe, expect, it, vi } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import * as breakQueue from './break-queue.js';
import { runLineBreakerPass } from './pass-driver.js';
import type { LayoutTextSeg } from './model.js';
import { SegmentQueue } from './segment-queue.js';

function token(text: string): LayoutTextSeg {
  return { text, fontSize: 10, fontFamily: 'Arial', bold: false, italic: false,
    underline: false, strikethrough: false, color: null, vertAlign: null, measuredWidth: 0 };
}

function paragraph(count: number) {
  return runLineBreakerPass({
    ctx: { measureText: (text: string) => ({ width: text.length * 5,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D,
    segs: Array.from({ length: count }, () => token('AAA ')),
    maxWidth: 200, firstIndent: 0, scale: 1, tabStops: [], fontFamilyClasses: {},
    tabOriginPx: 0, kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36, marginRightPx: 200,
    baseRtl: false, isJustified: false, stretchLastLine: false, widthPolicy: 'bounded',
    overflowPunct: true, passContext: { probeHeights: Array(count).fill(10) },
    wrapCtx: { floats: [], paraX: 0, startPageY: 0, columnXPt: 0, columnWidthPt: 200,
      pageH: count * 10, lineBoxH: () => 10,
      lineWindow: ({ topYPt, paragraphXPt }) => ({ topYPt,
        xOffsetPt: paragraphXPt === 0 ? 0 : 160 - paragraphXPt, maximumWidthPt: 40 }),
    },
  });
}

describe('gap snapshot work', () => {
  it('scales linearly with paragraph length rather than copying each remaining suffix', () => {
    const prepare = breakQueue.prepareBreakQueue;
    let visits = 0;
    // Count actual element reads at the prepared-queue boundary, including
    // native slice/spread copies and shift/unshift moves. The proxy leaves
    // production placement unchanged and introduces no timing assertion.
    const spy = vi.spyOn(breakQueue, 'prepareBreakQueue').mockImplementation((...args) =>
      new Proxy(prepare(...args), {
        get(target, key, receiver) {
          if (typeof key === 'string' && /^(0|[1-9]\d*)$/u.test(key)) visits += 1;
          return Reflect.get(target, key, receiver);
        },
      }));
    try {
      const work = (count: number) => {
        visits = 0;
        const lines = paragraph(count);
        expect(lines.flatMap(line => line.segments).map(segment => 'text' in segment ? segment.text : '').join(''))
          .toBe('AAA '.repeat(count));
        expect(lines.at(-1)?.consumedEnd).toEqual({ segIndex: count, charOffset: 0 });
        expect(lines.map(line => line.xOffset)).toEqual(
          Array.from({ length: count / 2 }, (_, index) => index % 2 === 0 ? 0 : 160));
        return visits;
      };
      const small = work(256);
      const large = work(1024);
      expect(small).toBeGreaterThan(0);
      expect(large / small).toBeLessThan(6);
    } finally {
      spy.mockRestore();
    }
  });

  it('restores consumed and prepended items, including a held item, on repeated rollback', () => {
    const original = [token('head'), token('middle'), token('last')];
    const queue = new SegmentQueue(original);
    const initial = queue.snapshot();
    queue.shift();
    const tail = token('split tail');
    queue.unshift(tail);
    const held = queue.shift();
    const fragment = queue.snapshot(held);
    queue.unshift(token('rejected replacement'));
    queue.shift();
    queue.shift();
    queue.unshift(token('retracted suffix'));
    queue.restore(fragment);
    expect([...queue]).toEqual([tail, original[1], original[2]]);
    while (queue.length) queue.shift();
    queue.restore(fragment);
    expect([...queue]).toEqual([tail, original[1], original[2]]);
    queue.restore(initial);
    expect([...queue]).toEqual(original);
  });
});
