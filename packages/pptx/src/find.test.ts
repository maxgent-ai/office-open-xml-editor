import { expect, it } from 'vitest';
import { PptxFindController } from './find.js';
import type { PptxTextRunInfo } from './renderer';

function run(text: string, overrides: Partial<PptxTextRunInfo> = {}): PptxTextRunInfo {
  return {
    text, inShapeX: 0, inShapeY: 0, w: text.length, h: 10,
    fontSize: 10, font: '10px monospace', shapeX: 0, shapeY: 0,
    shapeW: 100, shapeH: 20, rotation: 0, ...overrides,
  };
}

it('maps shared search results and highlights to PPTX slides', async () => {
  const slides = [[run('Hel'), run('lo')], [run('hello')]];
  const find = new PptxFindController(() => slides.length, async (slide) => slides[slide]);
  expect((await find.find('hello')).map((match) => match.location)).toEqual([{ slide: 0 }, { slide: 1 }]);
  find.next();
  expect(find.activeSlide()).toBe(0);
  expect(find.slideHighlights(0)[0].slices).toHaveLength(2);
  expect(find.slideRuns(1)).toEqual(slides[1]);
});

it('matches within a table cell but rejects text crossing cell boundaries', async () => {
  const sameCell = [
    run('Hel', { elementIndex: 0, tableCell: { row: 0, column: 0 } }),
    run('lo', { elementIndex: 0, tableCell: { row: 0, column: 0 } }),
  ];
  const adjacentCells = [
    run('foo', { elementIndex: 0, tableCell: { row: 0, column: 0 } }),
    run('bar', { elementIndex: 0, tableCell: { row: 0, column: 1 } }),
  ];
  const find = new PptxFindController(() => 1, async () => sameCell);
  await expect(find.find('Hello')).resolves.toHaveLength(1);
  find.setSlideRuns(0, adjacentCells);
  await expect(find.find('foobar')).resolves.toEqual([]);
});
