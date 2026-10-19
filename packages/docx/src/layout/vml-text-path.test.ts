import { expect, it } from 'vitest';
import { planVmlPathText } from './vml-text-path.js';
import type { ShapeRun } from '../types.js';

it('acquires nonparallel boundaries and bounds geometry and text replay work', () => {
  const shape: Pick<ShapeRun, 'subpaths'> = {
    subpaths: [
      [
        { cmd: 'moveTo', x: 0, y: 0 },
        { cmd: 'lineTo', x: 0.5, y: 0.5 },
      ],
      [
        { cmd: 'moveTo', x: 0, y: 1 },
        { cmd: 'lineTo', x: 1, y: 1 },
      ],
    ],
  };
  const command = { sourceBounds: { xPt: 0, yPt: -10, widthPt: 10, heightPt: 10 }, spans: [{ text: 'A\u0301' }] };
  const facts = { string: 'A\u0301', bold: false, italic: false, fitShape: true };
  const bounds = { xPt: 0, yPt: 0, widthPt: 100, heightPt: 100 };
  const slice = planVmlPathText(shape, bounds, facts, command)?.[128];
  expect(slice).toBeDefined();
  const u = 128.5 / 256;
  expect(slice!.xPt + slice!.c * -10).toBeCloseTo(u * 50);
  expect(slice!.yPt + slice!.d * -10).toBeCloseTo(u * 50);
  expect({ x: slice!.xPt, y: slice!.yPt }).toEqual({ x: u * 100, y: 100 });
  expect(() => planVmlPathText(shape, bounds, facts, command, 2)).toThrow(/work budget/);
  expect(() => planVmlPathText(shape, bounds, facts, {
    ...command, spans: [{ ...command.spans[0]!, text: 'A'.repeat(5000) }],
  })).toThrow(/work budget/);
});
