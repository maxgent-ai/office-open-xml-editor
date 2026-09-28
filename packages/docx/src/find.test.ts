import { expect, it } from 'vitest';
import { DocxFindController } from './find.js';
import type { DocxTextRunInfo } from './renderer';

it('maps shared search results and highlight geometry to DOCX pages', async () => {
  const run = (text: string): DocxTextRunInfo => ({
    text, x: 0, y: 0, w: text.length, h: 10, fontSize: 10, font: '10px monospace',
  });
  const pages = [[run('Hel'), run('lo')], [run('hello')]];
  const find = new DocxFindController(() => pages.length, async (page) => pages[page]);
  expect((await find.find('hello')).map((match) => match.location)).toEqual([{ page: 0 }, { page: 1 }]);
  find.next();
  expect(find.activePage()).toBe(0);
  expect(find.pageHighlights(0)[0].slices).toHaveLength(2);
  expect(find.pageRuns(1)).toEqual(pages[1]);
});
