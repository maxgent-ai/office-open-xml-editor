import { expect, it } from 'vitest';
import { breakDrawingMlText } from './break.js';
import { layoutStackedText } from './stacked.js';

// Inject counters into array operations, rather than timing or counting only
// measureText: the regression copied an unmeasured suffix on every CJK wrap.
// Count inserted/copied elements and flatMap visits during the synchronous
// production call; restore the native operations before any assertion/await.
function countWork(render: (visit: (count: number) => void) => void): number {
  let work = 0;
  const visit = (count: number): void => { work += count; };
  const push = Array.prototype.push;
  const slice = Array.prototype.slice;
  const flatMap = Array.prototype.flatMap;
  Array.prototype.push = function (this: unknown[], ...items: unknown[]): number {
    visit(items.length);
    return Reflect.apply(push, this, items);
  };
  Array.prototype.slice = function (this: unknown[], ...args: Parameters<typeof slice>): unknown[] {
    const result = Reflect.apply(slice, this, args);
    visit(result.length);
    return result;
  };
  Array.prototype.flatMap = function <U>(this: unknown[],
    callback: (value: unknown, index: number, array: unknown[]) => U | readonly U[], thisArg?: unknown,
  ): U[] {
    visit(this.length);
    return Reflect.apply(flatMap, this, [callback, thisArg]);
  };
  try { render(visit); } finally {
    Array.prototype.push = push;
    Array.prototype.slice = slice;
    Array.prototype.flatMap = flatMap;
  }
  return work;
}

it.each(['horizontal', 'wordArtVert', 'wordArtVertRtl'] as const)(
  '%s CJK wrapping scales linearly in visited/copied elements', (path) => {
    const render = (size: number): number => {
      const text = '日'.repeat(size);
      const runs = [{ type: 'text' as const, text, style: 'same' }];
      let painted = '';
      let lineCount = 0;
      const work = countWork((visit) => {
        if (path === 'horizontal') {
          const lines = breakDrawingMlText(runs, {
            maxWidth: 2,
            measureText(value) { visit(value.length); return value.length; },
          });
          lineCount = lines.length;
          painted = lines.map((line) => line.segments.map((seg) => seg.type === 'text' ? seg.text : '').join('')).join('');
        } else {
          const layout = layoutStackedText([{ runs, emptyThickness: 1 }], {
            direction: path, rect: { left: 0, top: 0, width: 10, height: 2 }, wrap: true, pxPerPt: 1,
            glyphs(value, style) {
              visit(value.length);
              return [...value].map((ch) => ({ text: ch, style, kind: 'upright' as const,
                advance: 1, thickness: 1, space: false }));
            },
          });
          lineCount = layout.columns.length;
          painted = layout.glyphs.map((glyph) => glyph.text).join('');
        }
      });
      expect(painted).toBe(text);
      expect(lineCount).toBe(size / 2);
      return work;
    };
    const small = render(2_000);
    const large = render(8_000);
    expect(small).toBeGreaterThan(0);
    expect(large / small).toBeLessThan(6);
  },
);

it('bounds suffix work when probing many authored tab cells', () => {
  const render = (size: number): number => {
    const text = `${'a\t'.repeat(size)}xyz`;
    let tabs = 0;
    let painted = '';
    const work = countWork((visit) => {
      const lines = breakDrawingMlText([{ type: 'text', text, style: 'same' }], {
        maxWidth: 2, tabStops: [{ pos: 1, algn: 'l' }],
        measureText(value) { visit(value.length); return value.length; },
      });
      for (const line of lines) for (const seg of line.segments) {
        if (seg.type === 'tab') tabs++;
        else if (seg.type === 'text') painted += seg.text;
      }
    });
    expect(tabs).toBe(size);
    expect(painted).toBe(`${'a'.repeat(size)}xyz`);
    return work;
  };
  expect(render(8_000) / render(2_000)).toBeLessThan(6);
});

it.each([
  { feature: 'CJK', tail: '日本語' },
  { feature: 'kinsoku punctuation', tail: '日「本、語' },
  { feature: 'tab', tail: '\txyz' },
  { feature: 'mixed scripts and clusters', tail: 'ก้ខ្មែរ日\u0301𠀀אב' },
])('bounds late $feature reads per line independently of consumed prefix length', ({ tail }) => {
  const maxWidth = 80;
  const render = (size: number, nonMonotoneMeasure: boolean) => {
    let atoms: unknown[] | undefined;
    const reads: number[] = [];
    let consumedReads = 0;
    let lineIndex = -1;
    const tailLine = size / maxWidth - 1;
    const consumed = size - maxWidth;
    const push = Array.prototype.push;
    // Capture the production atom array. The budget getter marks each line's
    // fitting phase, after one-time grapheme/script/opportunity preprocessing.
    // This also excludes the non-monotone model built on the first line.
    Array.prototype.push = function (this: unknown[], ...items: unknown[]): number {
      const item = items[0];
      if (typeof item === 'object' && item !== null && 'run' in item) atoms = this;
      return Reflect.apply(push, this, items);
    };
    let lines: ReturnType<typeof breakDrawingMlText<string>>;
    try {
      lines = breakDrawingMlText([{ type: 'text', text: 'a'.repeat(size) + tail, style: 'same' }], {
        get maxWidth() {
          if (lineIndex === -1 && atoms) {
            for (let i = 0; i < atoms.length; i++) {
              const atom = atoms[i];
              Object.defineProperty(atoms, i, {
                configurable: true, enumerable: true,
                get() {
                  if (lineIndex >= tailLine) {
                    reads[lineIndex - tailLine]++;
                    if (i < consumed) consumedReads++;
                  }
                  return atom;
                },
              });
            }
          }
          lineIndex++;
          if (lineIndex >= tailLine) reads.push(0);
          return maxWidth;
        },
        tabStops: [{ pos: 1, algn: 'l' }], nonMonotoneMeasure,
        measureText: (value) => value.length,
      });
    } finally {
      Array.prototype.push = push;
    }
    const painted = lines.map((line) => line.segments.map((seg) => seg.type === 'text' ? seg.text : '\t').join(''));
    expect(painted.slice(0, tailLine)).toEqual(Array<string>(tailLine).fill('a'.repeat(maxWidth)));
    expect(painted.slice(tailLine).join('')).toBe('a'.repeat(maxWidth) + tail);
    expect(consumedReads).toBe(0);
    expect(reads.length).toBeGreaterThan(0);
    for (const count of reads) {
      expect(count).toBeGreaterThan(0);
      expect(count).toBeLessThan(20 * (maxWidth + tail.length));
    }
    return { reads, lines: lines.slice(tailLine) };
  };
  for (const nonMonotoneMeasure of [false, true]) {
    expect(render(8_000, nonMonotoneMeasure)).toEqual(render(2_000, nonMonotoneMeasure));
  }
});

it('preserves code-point retraction across grapheme and supplementary-character offsets', () => {
  const lines = breakDrawingMlText([{ type: 'text', text: '日\u0301「本、𠀀「語、次', style: 'same' }], {
    maxWidth: 3, measureText: (text) => [...text].length,
  });
  expect(lines.map((line) => line.segments.map((seg) => seg.type === 'text' ? seg.text : '').join('')))
    .toEqual(['日\u0301', '「本、', '𠀀', '「語、', '次']);
});
